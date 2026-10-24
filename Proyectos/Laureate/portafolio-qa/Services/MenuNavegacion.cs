namespace PortafolioQA.Services;

/// <summary>
/// Entradas del menú lateral de Tabler (_MenuLateral). Solo presenta enlaces: la autorización la aplican las
/// convenciones de Razor Pages en Program.cs.
/// </summary>
public static class MenuNavegacion
{
    /// <param name="Carpeta">Ruta usada para resaltar la entrada activa.</param>
    /// <param name="Pagina">Página de destino.</param>
    public record Entrada(string Carpeta, string Pagina, string Texto, string Icono);

    public static readonly Entrada Portafolio = new("/Portafolio", "/Portafolio/Index", "Portafolio", "portafolio");
    public static readonly Entrada Dashboard = new("/Dashboard", "/Dashboard/Index", "Dashboard", "dashboard");

    public static readonly Entrada[] Planificacion =
    {
        new("/Planificacion/Gantt", "/Planificacion/Gantt", "Gantt", "gantt"),
        new("/Planificacion/Calendario", "/Planificacion/Calendario", "Calendario", "calendario"),
        new("/Planificacion/Capacidad", "/Planificacion/Capacidad", "Capacidad QA", "capacidad")
    };

    public static readonly Entrada[] Administracion =
    {
        new("/Admin/Proyectos", "/Admin/Proyectos/Index", "Proyectos", "proyectos"),
        new("/Admin/Asignaciones", "/Admin/Asignaciones/Index", "Asignaciones", "asignaciones"),
        new("/Admin/RecursosQA", "/Admin/RecursosQA/Index", "Recursos QA", "recursos"),
        new("/Admin/Instituciones", "/Admin/Instituciones/Index", "Instituciones", "instituciones"),
        new("/Admin/Sistemas", "/Admin/Sistemas/Index", "Sistemas / periféricos", "sistemas"),
        new("/Admin/Hitos", "/Admin/Hitos/Index", "Hitos", "hitos"),
        new("/Admin/Feriados", "/Admin/Feriados/Index", "Feriados", "feriados"),
        new("/Admin/Ausencias", "/Admin/Ausencias/Index", "Ausencias", "ausencias"),
        new("/Admin/Integraciones", "/Admin/Integraciones/Index", "Integraciones", "integraciones")
    };

    public static readonly Entrada[] Sistema =
    {
        new("/Admin/Usuarios", "/Admin/Usuarios/Index", "Usuarios", "usuarios"),
        new("/Admin/Bugzilla", "/Admin/Bugzilla/Index", "Integración Bugzilla", "integraciones"),
        new("/Setup", "/Setup/Index", "Conexión a la base de datos", "conexion")
    };

    public static bool EnRuta(string paginaActual, string carpeta) =>
        paginaActual.Equals(carpeta, StringComparison.OrdinalIgnoreCase)
        || paginaActual.StartsWith(carpeta + "/", StringComparison.OrdinalIgnoreCase);
}
