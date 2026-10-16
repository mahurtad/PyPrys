/* =====================================================================
   Limpieza de datos operativos - PortafolioQA
   CONSERVA : AspNet* (usuarios/roles), RecursosQA, Instituciones, Sistemas,
              SistemaInstituciones, Feriados, __EFMigrationsHistory
   BORRA    : Proyectos y todo lo que cuelga de ellos + Ausencias

   ANTES DE EJECUTAR
   1) Cerrar la aplicación.
   2) Backup:
      BACKUP DATABASE PortafolioQA TO DISK = N'C:\Temp\PortafolioQA_pre_limpieza.bak' WITH COPY_ONLY, CHECKSUM;
   3) Primera corrida con @Confirmar = 0 (hace ROLLBACK). Revisar conteos.
   4) Si AspNetUsers y RecursosQA no cambian, correr con @Confirmar = 1.
   ===================================================================== */
USE PortafolioQA;
SET XACT_ABORT ON;
SET NOCOUNT ON;

DECLARE @Confirmar bit = 0;   -- 0 = simulacro (ROLLBACK) | 1 = aplicar (COMMIT)

BEGIN TRAN;

-- Conteos previos
SELECT 'ANTES' AS momento, t, n FROM (
    SELECT 'AspNetUsers' t, COUNT(*) n FROM AspNetUsers
    UNION ALL SELECT 'AspNetUserRoles',   COUNT(*) FROM AspNetUserRoles
    UNION ALL SELECT 'RecursosQA',        COUNT(*) FROM RecursosQA
    UNION ALL SELECT 'Instituciones',     COUNT(*) FROM Instituciones
    UNION ALL SELECT 'Sistemas',          COUNT(*) FROM Sistemas
    UNION ALL SELECT 'Feriados',          COUNT(*) FROM Feriados
    UNION ALL SELECT 'Proyectos',         COUNT(*) FROM Proyectos
    UNION ALL SELECT 'Actividades',       COUNT(*) FROM Actividades
    UNION ALL SELECT 'Ausencias',         COUNT(*) FROM Ausencias
) x;

-- Borrado de hojas hacia la raíz
DELETE FROM SquashCasos;
DELETE FROM SquashIteraciones;
DELETE FROM SquashCampanas;
DELETE FROM AcuerdoActividades;
DELETE FROM Acuerdos;
DELETE FROM CierreParticipantes;
DELETE FROM CierreActividades;
DELETE FROM Cierres;
DELETE FROM ActividadAuditorias;
DELETE FROM Actividades;
DELETE FROM Hitos;
DELETE FROM Asignaciones;
DELETE FROM AsociacionesExternas;
DELETE FROM ProyectoSistemas;
DELETE FROM Proyectos;
DELETE FROM Ausencias;

-- Reiniciar identidades (TRUNCATE no aplica por las FKs). Próximo Id = 1.
DBCC CHECKIDENT ('SquashCasos',          RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('SquashIteraciones',    RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('SquashCampanas',       RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('Acuerdos',             RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('CierreParticipantes',  RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('CierreActividades',    RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('Cierres',              RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('ActividadAuditorias',  RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('Actividades',          RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('Hitos',                RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('Asignaciones',         RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('AsociacionesExternas', RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('Proyectos',            RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('Ausencias',            RESEED, 0) WITH NO_INFOMSGS;

-- Conteos posteriores: lo conservado debe ser idéntico a ANTES
SELECT 'DESPUES' AS momento, t, n FROM (
    SELECT 'AspNetUsers' t, COUNT(*) n FROM AspNetUsers
    UNION ALL SELECT 'AspNetUserRoles',   COUNT(*) FROM AspNetUserRoles
    UNION ALL SELECT 'RecursosQA',        COUNT(*) FROM RecursosQA
    UNION ALL SELECT 'Instituciones',     COUNT(*) FROM Instituciones
    UNION ALL SELECT 'Sistemas',          COUNT(*) FROM Sistemas
    UNION ALL SELECT 'Feriados',          COUNT(*) FROM Feriados
    UNION ALL SELECT 'Proyectos',         COUNT(*) FROM Proyectos
    UNION ALL SELECT 'Actividades',       COUNT(*) FROM Actividades
    UNION ALL SELECT 'Ausencias',         COUNT(*) FROM Ausencias
) x;

IF @Confirmar = 1
BEGIN
    COMMIT;
    PRINT 'Limpieza APLICADA (COMMIT).';
END
ELSE
BEGIN
    ROLLBACK;
    PRINT 'Simulacro: ROLLBACK. No se modificó nada.';
END
