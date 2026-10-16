using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Text;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Models;
using PortafolioQA.Services.Conexion;

namespace PortafolioQA.Pages.Setup;

/// <summary>
/// Asistente de conexión a la base de datos. Dos formas de acceso:
///  - Aplicación pendiente (sin conexión o sin esquema): con el código de configuración del servidor.
///  - Aplicación activa: solo un Administrador con sesión, que además confirma su contraseña para guardar o migrar.
/// Nunca crea logins ni usuarios en SQL Server: para eso genera un script que ejecuta el administrador de la base.
/// </summary>
public class IndexModel : PageModel
{
    private readonly ConexionStore _store;
    private readonly ConexionProbador _probador;
    private readonly EstadoApp _estado;
    private readonly IHostApplicationLifetime _vida;
    private readonly UserManager<UsuarioApp> _usuarios;

    public IndexModel(ConexionStore store, ConexionProbador probador, EstadoApp estado,
        IHostApplicationLifetime vida, UserManager<UsuarioApp> usuarios)
    {
        _store = store;
        _probador = probador;
        _estado = estado;
        _vida = vida;
        _usuarios = usuarios;
    }

    public class EntradaModel
    {
        public string? Codigo { get; set; }

        [Required(ErrorMessage = "Indique el servidor SQL.")]
        public string Servidor { get; set; } = string.Empty;

        [Range(1, 65535, ErrorMessage = "El puerto debe estar entre 1 y 65535.")]
        public int? Puerto { get; set; }

        [Required(ErrorMessage = "Indique el nombre de la base de datos.")]
        public string Base { get; set; } = "PortafolioQA";

        public bool UsaSql { get; set; }
        public string? Usuario { get; set; }
        public string? Clave { get; set; }
        public bool ConfiarCertificado { get; set; } = true;

        /// <summary>Solo con la aplicación activa: contraseña del administrador que guarda o migra.</summary>
        public string? ClaveAdmin { get; set; }
    }

    [BindProperty]
    public EntradaModel Entrada { get; set; } = new();

    public bool ModoCodigo => _estado.Fase == FaseApp.Pendiente;
    public FaseApp Fase => _estado.Fase;
    public string? DetalleEstado => _estado.Detalle;
    public string OrigenActual => _store.Origen;
    public string? AvisoStore => _store.Aviso;
    public string ArchivoCodigo => _store.ArchivoCodigo;
    public string Equipo => Environment.MachineName;
    public bool TieneConexionActual => _store.Cadena is not null;

    public ResultadoPrueba? Prueba { get; private set; }
    public string? Mensaje { get; private set; }
    public string? Error { get; private set; }
    public bool Activado { get; private set; }
    public string? ScriptAcceso { get; private set; }
    public string CuentaWindows { get; private set; } = string.Empty;
    public bool Bloqueada { get; private set; }

    public bool ConexionSegura =>
        Request.IsHttps || IPAddress.IsLoopback(HttpContext.Connection.RemoteIpAddress ?? IPAddress.None);

    private string Quien => ModoCodigo
        ? $"(código de configuración) {HttpContext.Connection.RemoteIpAddress}"
        : $"{User.Identity?.Name} {HttpContext.Connection.RemoteIpAddress}";

    public IActionResult OnGet()
    {
        var acceso = Acceso();
        if (acceso is not null) return acceso;
        if (ModoCodigo && _store.AsegurarCodigo() is null) Error = _store.ErrorCodigo;

        var actual = ConexionProbador.Descomponer(_store.Cadena);
        if (actual is not null)
        {
            Entrada.Servidor = actual.Servidor;
            Entrada.Puerto = actual.Puerto;
            Entrada.Base = actual.Base;
            Entrada.UsaSql = actual.UsaSql;
            Entrada.Usuario = actual.Usuario;
            Entrada.ConfiarCertificado = actual.ConfiarCertificado;
        }
        PrepararAyuda();
        return Page();
    }

    public async Task<IActionResult> OnPostProbarAsync()
    {
        var bloqueo = await AutorizarAsync(confirmar: false);
        if (bloqueo is not null) return bloqueo;
        if (!ModelState.IsValid) { PrepararAyuda(); return Page(); }

        Prueba = await _probador.ProbarAsync(Datos());
        _store.Auditar(Quien, "PROBAR", $"{Entrada.Servidor}/{Entrada.Base} ({(Entrada.UsaSql ? "SQL" : "Windows")}) -> {(Prueba.ListaParaActivar ? "lista" : "con observaciones")}");
        PrepararAyuda();
        return Page();
    }

    public async Task<IActionResult> OnPostActivarAsync()
    {
        var bloqueo = await AutorizarAsync(confirmar: true);
        if (bloqueo is not null) return bloqueo;
        if (!ModelState.IsValid) { PrepararAyuda(); return Page(); }

        var datos = Datos();
        Prueba = await _probador.ProbarAsync(datos);
        if (!Prueba.ListaParaActivar)
        {
            Error = "No se activó: la conexión, los permisos y el esquema deben estar completos. Revise el resultado de la prueba.";
            PrepararAyuda();
            return Page();
        }

        _store.Guardar(ConexionProbador.Construir(datos));
        _store.EliminarCodigo();
        _store.Auditar(Quien, "ACTIVAR", $"{datos.Servidor}{(datos.Puerto is int p ? "," + p : "")}/{datos.Base} ({(datos.UsaSql ? "SQL: " + datos.Usuario : "Windows")})");
        Activado = true;
        ReiniciarPronto();
        PrepararAyuda();
        return Page();
    }

    public async Task<IActionResult> OnPostMigrarAsync()
    {
        var bloqueo = await AutorizarAsync(confirmar: true);
        if (bloqueo is not null) return bloqueo;
        if (!ModelState.IsValid) { PrepararAyuda(); return Page(); }

        var datos = Datos();
        Prueba = await _probador.ProbarAsync(datos);
        if (!Prueba.Conecta || !Prueba.BaseExiste || !Prueba.PermisosOk || Prueba.BaseAjena)
        {
            Error = "No se puede migrar: la conexión o los permisos no están completos.";
        }
        else if (!Prueba.PuedeAlterarEsquema)
        {
            Error = "La cuenta no tiene permiso para crear o modificar tablas. Descargue el script de esquema y pida al administrador de la base que lo ejecute (o otorgue db_ddladmin).";
        }
        else
        {
            try
            {
                await _probador.MigrarAsync(ConexionProbador.Construir(datos));
                _store.Auditar(Quien, "MIGRAR", $"{datos.Servidor}/{datos.Base}: {Prueba.Pendientes.Count} migraciones aplicadas");
                Prueba = await _probador.ProbarAsync(datos);
                Mensaje = "Migraciones aplicadas. Ya puede activar la aplicación.";
            }
            catch (Exception ex)
            {
                Error = "Falló la aplicación de migraciones: " + ex.Message;
            }
        }
        PrepararAyuda();
        return Page();
    }

    /// <summary>Con la conexión vigente (sin volver a escribir credenciales): para aplicar migraciones tras una actualización.</summary>
    public async Task<IActionResult> OnPostMigrarActualAsync()
    {
        var bloqueo = await AutorizarAsync(confirmar: true, validarFormulario: false);
        if (bloqueo is not null) return bloqueo;
        if (_store.Cadena is null) { Error = "No hay una conexión vigente."; PrepararAyuda(); return Page(); }

        try
        {
            await _probador.MigrarAsync(_store.Cadena);
            _store.Auditar(Quien, "MIGRAR", "conexión vigente");
            Activado = true;
            Mensaje = "Migraciones aplicadas. La aplicación se reinicia.";
            ReiniciarPronto();
        }
        catch (Exception ex)
        {
            Error = "No se pudieron aplicar las migraciones con la cuenta actual: " + ex.Message +
                    " Descargue el script de esquema y pida al administrador de la base que lo ejecute.";
        }
        PrepararAyuda();
        return Page();
    }

    public async Task<IActionResult> OnPostScriptEsquemaAsync()
    {
        var bloqueo = await AutorizarAsync(confirmar: false, validarFormulario: false);
        if (bloqueo is not null) return bloqueo;
        _store.Auditar(Quien, "DESCARGAR_SCRIPT_ESQUEMA", "-");
        return File(Encoding.UTF8.GetBytes(_probador.GenerarScriptEsquema()), "text/plain; charset=utf-8", "portafolio-qa-esquema.sql");
    }

    public async Task<IActionResult> OnPostScriptAccesoAsync()
    {
        var bloqueo = await AutorizarAsync(confirmar: false);
        if (bloqueo is not null) return bloqueo;
        if (!ModelState.IsValid) { PrepararAyuda(); return Page(); }
        PrepararAyuda();
        ScriptAcceso = ConexionProbador.GenerarScriptAcceso(Datos(), CuentaWindows);
        return Page();
    }

    private DatosConexion Datos() =>
        new(Entrada.Servidor, Entrada.Puerto, Entrada.Base, Entrada.UsaSql, Entrada.Usuario, Entrada.Clave, Entrada.ConfiarCertificado);

    private void PrepararAyuda()
    {
        CuentaWindows = ConexionProbador.EsServidorLocal(Entrada.Servidor)
            ? ConexionProbador.IdentidadLocal()
            : ConexionProbador.CuentaRemota();
        if (Prueba is { FaltaAcceso: true } && ScriptAcceso is null)
            ScriptAcceso = ConexionProbador.GenerarScriptAcceso(Datos(), CuentaWindows);
    }

    /// <summary>Control de acceso de la carga inicial (GET). Devuelve null si se permite.</summary>
    private IActionResult? Acceso()
    {
        if (!ConexionSegura)
        {
            Error = "El asistente solo se puede usar por HTTPS (o desde el propio servidor con localhost), porque transmite credenciales.";
            return Page();
        }
        if (ModoCodigo) return null;
        if (User.Identity?.IsAuthenticated != true) return Challenge();
        if (!User.IsInRole(Roles.Administrador)) return Forbid();
        return null;
    }

    /// <summary>Control de acceso de las acciones (POST). Devuelve null si se permite continuar.</summary>
    private async Task<IActionResult?> AutorizarAsync(bool confirmar, bool validarFormulario = true)
    {
        if (!ConexionSegura)
        {
            Error = "El asistente solo se puede usar por HTTPS (o desde el propio servidor con localhost).";
            PrepararAyuda();
            return Page();
        }

        if (ModoCodigo)
        {
            switch (_store.Verificar(Entrada.Codigo, out var minutos))
            {
                case ResultadoCodigo.Ok: break;
                case ResultadoCodigo.SinCodigo:
                    Error = _store.ErrorCodigo ?? "No se pudo generar el código de configuración.";
                    PrepararAyuda();
                    return Page();
                case ResultadoCodigo.Bloqueado:
                    Bloqueada = true;
                    Error = $"Demasiados intentos con un código incorrecto. Vuelva a intentar en {minutos} min.";
                    PrepararAyuda();
                    return Page();
                default:
                    Error = "Código de configuración incorrecto.";
                    PrepararAyuda();
                    return Page();
            }
        }
        else
        {
            if (User.Identity?.IsAuthenticated != true) return Challenge();
            if (!User.IsInRole(Roles.Administrador)) return Forbid();
            if (confirmar)
            {
                var usuario = await _usuarios.GetUserAsync(User);
                if (usuario is null || string.IsNullOrEmpty(Entrada.ClaveAdmin) || !await _usuarios.CheckPasswordAsync(usuario, Entrada.ClaveAdmin))
                {
                    Error = "Confirme su contraseña de administrador para continuar.";
                    PrepararAyuda();
                    return Page();
                }
            }
        }

        if (!validarFormulario) ModelState.Clear();
        return null;
    }

    private void ReiniciarPronto() =>
        _ = Task.Run(async () => { await Task.Delay(2000); _vida.StopApplication(); });
}
