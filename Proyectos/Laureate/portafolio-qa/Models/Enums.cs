namespace PortafolioQA.Models;

public enum TipoIniciativa
{
    Proyecto,
    RequerimientoNoBAU,
    BAU
}

/// <summary>
/// Los valores numéricos se persisten en BD: no reordenar. El valor 1 (antes "En ejecución") fue retirado;
/// el estado se deriva del avance (0 = Planificado, 1-99 = En certificación, 100 = Cerrado). Cancelado es manual.
/// </summary>
public enum EstadoProyecto
{
    Planificado = 0,
    EnCertificacion = 2,
    Cerrado = 3,
    Cancelado = 4
}

/// <summary>Los valores numéricos se persisten en BD: no reordenar.</summary>
public enum TipoAusencia
{
    Vacaciones = 0,
    Compensacion = 1,
    Cumpleanos = 2,
    LicenciaMedica = 3,
    LicenciaMaternidad = 4
}

public enum EquipoQA
{
    Change,
    RunBau
}

public enum HerramientaIntegracion
{
    SquashTM,
    Bugzilla
}

public enum EstadoSincronizacion
{
    PendienteConfigurar,
    Conectado,
    Desactualizado,
    Error
}
