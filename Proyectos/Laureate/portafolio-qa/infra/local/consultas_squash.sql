-- Consultas de lectura para Squash TM (PostgreSQL).
-- En SQLTools, selecciona una consulta y ejecuta SQLTools: Run Query.
-- Las referencias de defectos apuntan al bugtracker externo. Squash conserva
-- aquí el ID remoto y el vínculo; el estado/severidad/título se consulta allí.

-- 1) Campañas por proyecto: fechas, estado e iteraciones asociadas.
WITH iteraciones AS (
    SELECT campaign_id, COUNT(DISTINCT iteration_id) AS cantidad_iteraciones
    FROM campaign_iteration
    GROUP BY campaign_id
), casos AS (
    SELECT campaign_id, COUNT(*) AS casos_asignados_directamente
    FROM campaign_test_plan_item
    GROUP BY campaign_id
)
SELECT
    p.name AS proyecto,
    cn.name AS campana,
    c.reference AS referencia,
    c.campaign_status AS estado,
    c.scheduled_start_date AS inicio_planificado,
    c.scheduled_end_date AS fin_planificado,
    c.actual_start_date AS inicio_real,
    c.actual_end_date AS fin_real,
    COALESCE(i.cantidad_iteraciones, 0) AS iteraciones,
    COALESCE(ca.casos_asignados_directamente, 0) AS casos_asignados_directamente
FROM campaign c
JOIN campaign_library_node cn ON cn.cln_id = c.cln_id
JOIN project p ON p.project_id = cn.project_id
LEFT JOIN iteraciones i ON i.campaign_id = c.cln_id
LEFT JOIN casos ca ON ca.campaign_id = c.cln_id
ORDER BY p.name, c.scheduled_start_date NULLS LAST, cn.name;

-- 2) Avance por iteración dentro de una campaña.
-- Porcentaje ejecutado = casos con resultado SUCCESS o FAILURE / total.
SELECT
    p.name AS proyecto,
    cn.name AS campana,
    it.name AS iteracion,
    it.iteration_status AS estado_iteracion,
    it.scheduled_start_date AS inicio_planificado,
    it.scheduled_end_date AS fin_planificado,
    COUNT(tpi.test_plan_item_id) AS casos_planificados,
    COUNT(*) FILTER (WHERE tpi.execution_status = 'SUCCESS') AS aprobados,
    COUNT(*) FILTER (WHERE tpi.execution_status = 'FAILURE') AS fallidos,
    COUNT(*) FILTER (WHERE tpi.execution_status = 'READY') AS pendientes,
    COUNT(*) FILTER (WHERE tpi.assignee_id IS NOT NULL) AS casos_con_asignado,
    ROUND(
        100.0 * COUNT(*) FILTER (WHERE tpi.execution_status IN ('SUCCESS', 'FAILURE'))
        / NULLIF(COUNT(tpi.test_plan_item_id), 0),
        1
    ) AS porcentaje_ejecutado
FROM campaign_iteration ci
JOIN campaign c ON c.cln_id = ci.campaign_id
JOIN campaign_library_node cn ON cn.cln_id = c.cln_id
JOIN project p ON p.project_id = cn.project_id
JOIN iteration it ON it.iteration_id = ci.iteration_id
LEFT JOIN test_plan tp ON tp.test_plan_id = it.test_plan_id
LEFT JOIN test_plan_item tpi ON tpi.test_plan_id = tp.test_plan_id
GROUP BY p.name, cn.name, it.name, it.iteration_status,
         it.scheduled_start_date, it.scheduled_end_date
ORDER BY p.name, cn.name, it.name;

-- 3) Casos asignados directamente a una campaña (fuera del detalle por iteración).
SELECT
    p.name AS proyecto,
    cn.name AS campana,
    tcn.name AS caso,
    tc.reference AS referencia_caso,
    u.login AS asignado_a,
    CONCAT_WS(' ', u.first_name, u.last_name) AS nombre_asignado,
    d.name AS conjunto_de_datos,
    cpi.test_plan_order AS orden
FROM campaign_test_plan_item cpi
JOIN campaign c ON c.cln_id = cpi.campaign_id
JOIN campaign_library_node cn ON cn.cln_id = c.cln_id
JOIN project p ON p.project_id = cn.project_id
JOIN test_case_library_node tcn ON tcn.cln_id = cpi.test_case_id
LEFT JOIN test_case tc ON tc.tcln_id = cpi.test_case_id
LEFT JOIN core_user u ON u.party_id = cpi.user_id
LEFT JOIN dataset d ON d.dataset_id = cpi.dataset_id
ORDER BY p.name, cn.name, cpi.test_plan_order;

-- 4) Casos de iteración y su responsable, incluyendo los que aún no tienen asignado.
SELECT
    p.name AS proyecto,
    cn.name AS campana,
    it.name AS iteracion,
    tcn.name AS caso,
    tc.reference AS referencia_caso,
    tpi.label AS etiqueta_plan,
    u.login AS asignado_a,
    CONCAT_WS(' ', u.first_name, u.last_name) AS nombre_asignado,
    d.name AS conjunto_de_datos,
    tpi.execution_status AS ultimo_estado,
    tpi.last_executed_on AS ultima_ejecucion
FROM test_plan_item tpi
JOIN test_plan tp ON tp.test_plan_id = tpi.test_plan_id
JOIN iteration it ON it.test_plan_id = tp.test_plan_id
JOIN campaign_iteration ci ON ci.iteration_id = it.iteration_id
JOIN campaign c ON c.cln_id = ci.campaign_id
JOIN campaign_library_node cn ON cn.cln_id = c.cln_id
JOIN project p ON p.project_id = cn.project_id
JOIN test_case_library_node tcn ON tcn.cln_id = tpi.tcln_id
LEFT JOIN test_case tc ON tc.tcln_id = tpi.tcln_id
LEFT JOIN core_user u ON u.party_id = tpi.assignee_id
LEFT JOIN dataset d ON d.dataset_id = tpi.dataset_id
ORDER BY p.name, cn.name, it.name, tpi.item_order;

-- 5) Historial de ejecuciones de casos incluidos en iteraciones de campañas.
SELECT
    p.name AS proyecto,
    cn.name AS campana,
    it.name AS iteracion,
    COALESCE(tcn.name, e.name) AS caso,
    e.reference AS referencia_caso,
    e.execution_status AS resultado,
    e.last_executed_on AS ejecutado_el,
    e.last_executed_by AS ejecutado_por_login,
    CONCAT_WS(' ', u.first_name, u.last_name) AS nombre_ejecutor,
    e.execution_id,
    e.issue_list_id
FROM execution e
LEFT JOIN test_plan_item tpi ON tpi.test_plan_item_id = e.test_plan_item_id
LEFT JOIN test_plan tp ON tp.test_plan_id = tpi.test_plan_id
LEFT JOIN iteration it ON it.test_plan_id = tp.test_plan_id
LEFT JOIN campaign_iteration ci ON ci.iteration_id = it.iteration_id
LEFT JOIN campaign c ON c.cln_id = ci.campaign_id
LEFT JOIN campaign_library_node cn ON cn.cln_id = c.cln_id
LEFT JOIN project p ON p.project_id = cn.project_id
LEFT JOIN test_case_library_node tcn ON tcn.cln_id = e.tcln_id
LEFT JOIN core_user u ON u.login = e.last_executed_by
ORDER BY e.last_executed_on DESC NULLS LAST, e.execution_id DESC;

-- 6) Defectos vinculados a ejecuciones o pasos de ejecución.
-- remote_issue_id es el ID del defecto en Bugzilla/bugtracker.
WITH ejecuciones AS (
    SELECT
        e.execution_id,
        NULL::bigint AS execution_step_id,
        'Ejecución'::text AS origen,
        e.execution_status AS resultado,
        e.issue_list_id,
        e.last_executed_on AS fecha,
        p.name AS proyecto,
        cn.name AS campana,
        it.name AS iteracion,
        COALESCE(tcn.name, e.name) AS caso,
        e.reference AS referencia_caso
    FROM execution e
    LEFT JOIN test_plan_item tpi ON tpi.test_plan_item_id = e.test_plan_item_id
    LEFT JOIN test_plan tp ON tp.test_plan_id = tpi.test_plan_id
    LEFT JOIN iteration it ON it.test_plan_id = tp.test_plan_id
    LEFT JOIN campaign_iteration ci ON ci.iteration_id = it.iteration_id
    LEFT JOIN campaign c ON c.cln_id = ci.campaign_id
    LEFT JOIN campaign_library_node cn ON cn.cln_id = c.cln_id
    LEFT JOIN project p ON p.project_id = cn.project_id
    LEFT JOIN test_case_library_node tcn ON tcn.cln_id = e.tcln_id

    UNION ALL

    SELECT
        e.execution_id,
        es.execution_step_id,
        'Paso de ejecución'::text AS origen,
        es.execution_status AS resultado,
        es.issue_list_id,
        es.last_executed_on AS fecha,
        p.name AS proyecto,
        cn.name AS campana,
        it.name AS iteracion,
        COALESCE(tcn.name, e.name) AS caso,
        e.reference AS referencia_caso
    FROM execution_step es
    JOIN execution_execution_steps ees ON ees.execution_step_id = es.execution_step_id
    JOIN execution e ON e.execution_id = ees.execution_id
    LEFT JOIN test_plan_item tpi ON tpi.test_plan_item_id = e.test_plan_item_id
    LEFT JOIN test_plan tp ON tp.test_plan_id = tpi.test_plan_id
    LEFT JOIN iteration it ON it.test_plan_id = tp.test_plan_id
    LEFT JOIN campaign_iteration ci ON ci.iteration_id = it.iteration_id
    LEFT JOIN campaign c ON c.cln_id = ci.campaign_id
    LEFT JOIN campaign_library_node cn ON cn.cln_id = c.cln_id
    LEFT JOIN project p ON p.project_id = cn.project_id
    LEFT JOIN test_case_library_node tcn ON tcn.cln_id = e.tcln_id
)
SELECT
    e.proyecto,
    e.campana,
    e.iteracion,
    e.caso,
    e.referencia_caso,
    e.origen,
    e.resultado,
    e.fecha,
    i.remote_issue_id AS id_defecto_externo,
    b.kind AS tipo_bugtracker,
    e.execution_id,
    e.execution_step_id
FROM ejecuciones e
JOIN issue i ON i.issue_list_id = e.issue_list_id
LEFT JOIN bugtracker b ON b.bugtracker_id = i.bugtracker_id
ORDER BY e.fecha DESC NULLS LAST, e.execution_id DESC, i.remote_issue_id;
