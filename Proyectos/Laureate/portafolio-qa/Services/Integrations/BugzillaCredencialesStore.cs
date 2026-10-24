using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Extensions.Options;
using PortafolioQA.Services.Conexion;

namespace PortafolioQA.Services.Integrations;

/// <summary>Configuración vigente de Bugzilla para crear productos: la guardada desde la aplicación o, si no hay, la del servidor.</summary>
public interface IBugzillaServicioConfig
{
    BugzillaOptions Actual { get; }
}

/// <summary>Configuración inmutable (pruebas y usos que no necesitan recarga).</summary>
public class BugzillaServicioFija(BugzillaOptions opciones) : IBugzillaServicioConfig
{
    public BugzillaOptions Actual { get; } = opciones;
}

public enum OrigenBugzilla { SinConfigurar, Servidor, Aplicacion }

/// <summary>
/// Guarda cifrada (ASP.NET Data Protection) la conexión a Bugzilla: dirección, clave de la cuenta de servicio (grupo
/// editcomponents, crea productos), su login y clave de lectura (consulta de defectos), en la carpeta de datos del servidor junto
/// a la conexión a la base. Cada dato guardado desde Sistema → Integración Bugzilla tiene prioridad sobre la configuración del
/// servidor (appsettings, user-secrets o variables Bugzilla__*), que rige para lo que no se haya guardado: así las instalaciones
/// actuales siguen igual. Guardar o quitar surte efecto de inmediato, sin reiniciar la aplicación.
/// La clave nunca se devuelve a la interfaz ni se escribe en la auditoría.
/// </summary>
public class BugzillaCredencialesStore : IBugzillaServicioConfig
{
    private const string Proposito = "PortafolioQA.BugzillaServicio.v1";

    private sealed record Guardado(string? ApiKey, string? CuentaServicio, string GuardadoPor, DateTime GuardadoEn,
        string? BaseUrl = null, string? ApiKeyLectura = null, string? RutaRest = null);

    private readonly ConexionStore _conexion;
    private readonly IOptions<BugzillaOptions> _servidor;
    private readonly IDataProtector _protector;
    private readonly ILogger<BugzillaCredencialesStore> _logger;
    private readonly object _lock = new();
    private volatile Guardado? _guardado;
    private volatile BugzillaOptions _actual;

    public BugzillaCredencialesStore(ConexionStore conexion, IOptions<BugzillaOptions> servidor,
        IDataProtectionProvider proveedor, ILogger<BugzillaCredencialesStore> logger)
    {
        _conexion = conexion;
        _servidor = servidor;
        _protector = proveedor.CreateProtector(Proposito);
        _logger = logger;
        _actual = servidor.Value;
        Recargar();
    }

    public string Archivo => Path.Combine(_conexion.Ruta, "bugzilla-servicio.dat");

    public BugzillaOptions Actual => _actual;
    public string? Aviso { get; private set; }
    public string? GuardadoPor => _guardado?.GuardadoPor;
    public DateTime? GuardadoEn => _guardado?.GuardadoEn;

    /// <summary>Login de la cuenta de servicio en uso.</summary>
    public string CuentaServicio => _actual.CuentaServicio;

    public OrigenBugzilla Origen =>
        _guardado is not null ? OrigenBugzilla.Aplicacion
        : !string.IsNullOrWhiteSpace(_servidor.Value.AdminApiKey) ? OrigenBugzilla.Servidor
        : OrigenBugzilla.SinConfigurar;

    public void Recargar()
    {
        lock (_lock)
        {
            Aviso = null;
            _guardado = null;
            if (File.Exists(Archivo))
            {
                try
                {
                    _guardado = JsonSerializer.Deserialize<Guardado>(_protector.Unprotect(File.ReadAllText(Archivo, Encoding.UTF8).Trim()));
                }
                catch (Exception ex) when (ex is CryptographicException or FormatException or JsonException)
                {
                    Aviso = "Existe una clave de Bugzilla guardada pero no se pudo leer (cambiaron las claves de protección de datos). Vuelva a guardarla.";
                    _logger.LogWarning(ex, "No se pudo descifrar {Archivo}.", Archivo);
                }
            }
            Componer();
        }
    }

    private void Componer()
    {
        var s = _servidor.Value;
        var g = _guardado;
        _actual = new BugzillaOptions
        {
            BaseUrl = g?.BaseUrl ?? s.BaseUrl,
            RutaRest = g?.RutaRest ?? s.RutaRest,
            ApiKey = g?.ApiKeyLectura ?? s.ApiKey,
            AdminApiKey = g?.ApiKey ?? s.AdminApiKey,
            CuentaServicio = g?.CuentaServicio ?? s.CuentaServicio
        };
    }

    /// <summary>
    /// Guarda cifrado lo indicado (un dato en blanco conserva el valor guardado), conserva copia .bak y lo deja activo.
    /// La dirección se normaliza (sin barra final ni /rest).
    /// </summary>
    public void Guardar(string? baseUrl, string? apiKeyServicio, string? cuentaServicio, string? apiKeyLectura, string? rutaRest, string usuario)
    {
        static string? Dato(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        string detalle;
        lock (_lock)
        {
            var previo = _guardado;
            var url = Dato(baseUrl) is { } u ? NormalizarUrl(u) : previo?.BaseUrl;
            var datos = new Guardado(Dato(apiKeyServicio) ?? previo?.ApiKey, Dato(cuentaServicio) ?? previo?.CuentaServicio, usuario,
                DateTime.Now, url, Dato(apiKeyLectura) ?? previo?.ApiKeyLectura,
                Dato(baseUrl) is null ? previo?.RutaRest : Dato(rutaRest) ?? "rest");
            Directory.CreateDirectory(_conexion.Ruta);
            var temporal = Archivo + ".tmp";
            File.WriteAllText(temporal, _protector.Protect(JsonSerializer.Serialize(datos)), new UTF8Encoding(false));
            if (File.Exists(Archivo)) File.Copy(Archivo, Archivo + ".bak", overwrite: true);
            File.Move(temporal, Archivo, overwrite: true);
            _guardado = datos;
            Aviso = null;
            Componer();
            detalle = $"dirección {datos.BaseUrl ?? "(servidor)"} ({datos.RutaRest ?? "ruta del servidor"}); cuenta {datos.CuentaServicio ?? "(servidor)"}; " +
                      $"clave de servicio {(Dato(apiKeyServicio) is null ? "sin cambio" : "reemplazada")}; " +
                      $"clave de lectura {(Dato(apiKeyLectura) is null ? "sin cambio" : "reemplazada")}";
        }
        _conexion.Auditar(usuario, "Bugzilla: conexión guardada", detalle);
    }

    /// <summary>Quita el sufijo /rest o /rest.cgi y la barra final: la dirección que se guarda es la raíz de Bugzilla.</summary>
    public static string NormalizarUrl(string url)
    {
        var v = url.Trim().TrimEnd('/');
        if (v.EndsWith("/rest.cgi", StringComparison.OrdinalIgnoreCase)) v = v[..^9].TrimEnd('/');
        else if (v.EndsWith("/rest", StringComparison.OrdinalIgnoreCase)) v = v[..^5].TrimEnd('/');
        return v;
    }

    /// <summary>Quita todo lo guardado desde la pantalla: vuelve a regir la configuración del servidor (si la hay).</summary>
    public void Quitar(string usuario)
    {
        lock (_lock)
        {
            if (File.Exists(Archivo)) File.Copy(Archivo, Archivo + ".bak", overwrite: true);
            File.Delete(Archivo);
            _guardado = null;
            Aviso = null;
            Componer();
        }
        _conexion.Auditar(usuario, "Bugzilla: conexión guardada quitada", "vuelve la configuración del servidor");
    }
}
