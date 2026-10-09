namespace PortafolioQA.Models;

public static class EnumDisplay
{
    public static string Texto(this TipoIniciativa tipo) => tipo switch
    {
        TipoIniciativa.Proyecto => "Proyecto",
        TipoIniciativa.RequerimientoNoBAU => "Requerimiento no BAU",
        TipoIniciativa.BAU => "BAU",
        _ => tipo.ToString()
    };

    public static string Texto(this TipoAusencia tipo) => tipo switch
    {
        TipoAusencia.Vacaciones => "Vacaciones",
        TipoAusencia.Compensacion => "Compensación",
        TipoAusencia.Cumpleanos => "Cumpleaños",
        TipoAusencia.LicenciaMedica => "Licencia médica",
        TipoAusencia.LicenciaMaternidad => "Licencia maternidad",
        _ => tipo.ToString()
    };

    public static string Texto(this EstadoProyecto estado) => estado switch
    {
        EstadoProyecto.Planificado => "Planificado",
        EstadoProyecto.EnCertificacion => "En certificación",
        EstadoProyecto.Cerrado => "Cerrado",
        EstadoProyecto.Cancelado => "Cancelado",
        _ => estado.ToString()
    };

    public static string Texto(this EstadoActividad estado) => estado switch
    {
        EstadoActividad.Pendiente => "Pendiente",
        EstadoActividad.EnCurso => "En curso",
        EstadoActividad.Completada => "Completada",
        EstadoActividad.Cancelada => "Cancelada",
        _ => estado.ToString()
    };

    public static string BadgeClase(this EstadoActividad estado) => estado switch
    {
        EstadoActividad.Pendiente => "bg-secondary",
        EstadoActividad.EnCurso => "bg-primary",
        EstadoActividad.Completada => "bg-success",
        _ => "bg-dark"
    };

    public static string Texto(this EquipoQA equipo) => equipo switch
    {
        EquipoQA.Change => "Change",
        EquipoQA.RunBau => "RUN / BAU",
        _ => equipo.ToString()
    };

    public static string Texto(this HerramientaIntegracion herramienta) => herramienta switch
    {
        HerramientaIntegracion.SquashTM => "Squash TM",
        HerramientaIntegracion.Bugzilla => "Bugzilla",
        _ => herramienta.ToString()
    };

    public static string Texto(this EstadoSincronizacion estado) => estado switch
    {
        EstadoSincronizacion.PendienteConfigurar => "Pendiente de configurar",
        EstadoSincronizacion.Conectado => "Conectado",
        EstadoSincronizacion.Desactualizado => "Desactualizado",
        EstadoSincronizacion.Error => "Error de conexión",
        _ => estado.ToString()
    };

    public static string BadgeClase(this EstadoProyecto estado) => estado switch
    {
        EstadoProyecto.Planificado => "bg-secondary",
        EstadoProyecto.EnCertificacion => "bg-info text-dark",
        EstadoProyecto.Cerrado => "bg-success",
        EstadoProyecto.Cancelado => "bg-dark",
        _ => "bg-secondary"
    };

    public static string BadgeClase(this EstadoSincronizacion estado) => estado switch
    {
        EstadoSincronizacion.PendienteConfigurar => "bg-secondary",
        EstadoSincronizacion.Conectado => "bg-success",
        EstadoSincronizacion.Desactualizado => "bg-warning text-dark",
        EstadoSincronizacion.Error => "bg-danger",
        _ => "bg-secondary"
    };
}
