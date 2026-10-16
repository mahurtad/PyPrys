using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.DataProtection;

namespace PortafolioQA.Services.Conexion;

public enum ResultadoCodigo { Ok, Incorrecto, Bloqueado, SinCodigo }

/// <summary>
/// Origen único de la cadena de conexión a SQL Server. Prioridad: archivo cifrado escrito por el asistente de
/// configuración (/Setup) y, si no existe, la configuración del servidor (variable de entorno, appsettings o secretos).
/// Así las instalaciones existentes siguen igual hasta que un administrador guarde una conexión nueva.
/// El archivo se cifra con ASP.NET Data Protection (las claves persisten en DataProtection:RutaClaves).
/// También administra el código de configuración que protege el asistente mientras la aplicación no está activada
/// y la auditoría de cambios (sin contraseñas).
/// </summary>
public class ConexionStore
{
    private const string Proposito = "PortafolioQA.Conexion.v1";
    private const int MaxIntentos = 5;
    private static readonly TimeSpan Bloqueo = TimeSpan.FromMinutes(15);
    private static readonly char[] Alfabeto = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789".ToCharArray();

    private readonly IConfiguration _config;
    private readonly IDataProtector _protector;
    private readonly ILogger<ConexionStore> _logger;
    private readonly object _lock = new();
    private int _fallidos;
    private DateTime _bloqueadoHasta = DateTime.MinValue;

    public ConexionStore(IConfiguration config, IDataProtectionProvider proveedor, ILogger<ConexionStore> logger)
    {
        _config = config;
        _protector = proveedor.CreateProtector(Proposito);
        _logger = logger;
        Recargar();
    }

    public string? Cadena { get; private set; }
    public string Origen { get; private set; } = "Sin configurar";
    public string? Aviso { get; private set; }

    /// <summary>Motivo por el que no se pudo crear o leer el código de configuración (carpeta de datos sin permiso de escritura).</summary>
    public string? ErrorCodigo { get; private set; }

    /// <summary>Carpeta de datos del servidor, fuera de la carpeta de la aplicación (sobrevive a las actualizaciones).</summary>
    public string Ruta => _config["Conexion:Ruta"]
        ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "PortafolioQA");

    public string ArchivoConexion => Path.Combine(Ruta, "conexion.dat");
    public string ArchivoCodigo => Path.Combine(Ruta, "codigo-configuracion.txt");
    public string ArchivoAuditoria => Path.Combine(Ruta, "auditoria-configuracion.log");

    public void Recargar()
    {
        Aviso = null;
        if (File.Exists(ArchivoConexion))
        {
            try
            {
                Cadena = _protector.Unprotect(File.ReadAllText(ArchivoConexion, Encoding.UTF8).Trim());
                Origen = "Asistente de configuración (archivo cifrado)";
                return;
            }
            catch (Exception ex) when (ex is CryptographicException or FormatException)
            {
                Aviso = "Existe una configuración guardada pero no se pudo leer (cambiaron las claves de protección de datos). Vuelva a configurar la conexión.";
                _logger.LogWarning(ex, "No se pudo descifrar {Archivo}.", ArchivoConexion);
            }
        }

        var cadena = _config.GetConnectionString("PortafolioQAContext");
        if (!string.IsNullOrWhiteSpace(cadena))
        {
            Cadena = cadena;
            Origen = "Configuración del servidor (variable de entorno, appsettings o secretos)";
        }
        else
        {
            Cadena = null;
            Origen = "Sin configurar";
        }
    }

    /// <summary>Guarda la cadena cifrada, conserva una copia .bak de la anterior y la deja activa.</summary>
    public void Guardar(string cadena)
    {
        Directory.CreateDirectory(Ruta);
        var temporal = ArchivoConexion + ".tmp";
        File.WriteAllText(temporal, _protector.Protect(cadena), new UTF8Encoding(false));
        if (File.Exists(ArchivoConexion)) File.Copy(ArchivoConexion, ArchivoConexion + ".bak", overwrite: true);
        File.Move(temporal, ArchivoConexion, overwrite: true);
        Recargar();
    }

    /// <summary>
    /// Devuelve el código de configuración, creándolo (y su archivo) si no existe. Devuelve null si la carpeta de datos
    /// no es accesible (la aplicación NO debe caerse por eso: el asistente muestra el motivo en <see cref="ErrorCodigo"/>).
    /// </summary>
    public string? AsegurarCodigo()
    {
        lock (_lock)
        {
            try
            {
                return AsegurarCodigoInterno();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                ErrorCodigo = $"La aplicación no puede escribir en la carpeta de datos '{Ruta}': {ex.Message} " +
                              "Dé permiso de modificación sobre esa carpeta a la cuenta del application pool (IIS APPPOOL\\<pool>).";
                _logger.LogError(ex, "No se pudo crear el código de configuración en {Ruta}.", Ruta);
                return null;
            }
        }
    }

    private string AsegurarCodigoInterno()
    {
        ErrorCodigo = null;
        if (File.Exists(ArchivoCodigo))
        {
            var existente = File.ReadAllText(ArchivoCodigo).Trim();
            if (existente.Length > 0) return existente;
        }

        var tramos = new string[3];
        for (var t = 0; t < tramos.Length; t++)
        {
            var c = new char[4];
            for (var i = 0; i < c.Length; i++) c[i] = Alfabeto[RandomNumberGenerator.GetInt32(Alfabeto.Length)];
            tramos[t] = new string(c);
        }
        var codigo = string.Join('-', tramos);
        Directory.CreateDirectory(Ruta);
        File.WriteAllText(ArchivoCodigo, codigo + Environment.NewLine);
        return codigo;
    }

    public void EliminarCodigo()
    {
        lock (_lock)
        {
            try { if (File.Exists(ArchivoCodigo)) File.Delete(ArchivoCodigo); }
            catch (IOException ex) { _logger.LogWarning(ex, "No se pudo borrar el código de configuración."); }
        }
    }

    /// <summary>Compara en tiempo constante; tras 5 fallos seguidos bloquea los intentos durante 15 minutos.</summary>
    public ResultadoCodigo Verificar(string? ingresado, out int minutosRestantes)
    {
        lock (_lock)
        {
            minutosRestantes = 0;
            var ahora = DateTime.UtcNow;
            if (ahora < _bloqueadoHasta)
            {
                minutosRestantes = (int)Math.Ceiling((_bloqueadoHasta - ahora).TotalMinutes);
                return ResultadoCodigo.Bloqueado;
            }

            var codigo = AsegurarCodigo();
            if (codigo is null) return ResultadoCodigo.SinCodigo;
            var esperado = Encoding.UTF8.GetBytes(codigo);
            var recibido = Encoding.UTF8.GetBytes((ingresado ?? string.Empty).Trim().ToUpperInvariant());
            if (recibido.Length == esperado.Length && CryptographicOperations.FixedTimeEquals(recibido, esperado))
            {
                _fallidos = 0;
                return ResultadoCodigo.Ok;
            }

            if (++_fallidos >= MaxIntentos)
            {
                _bloqueadoHasta = ahora + Bloqueo;
                _fallidos = 0;
                _logger.LogWarning("Asistente de configuración: {Max} códigos incorrectos seguidos; bloqueado {Min} min.", MaxIntentos, Bloqueo.TotalMinutes);
            }
            return ResultadoCodigo.Incorrecto;
        }
    }

    /// <summary>Registra quién cambió qué y cuándo. Nunca debe recibir contraseñas ni cadenas de conexión completas.</summary>
    public void Auditar(string quien, string accion, string detalle)
    {
        var linea = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {quien} | {accion} | {detalle}{Environment.NewLine}";
        _logger.LogInformation("Configuración: {Quien} {Accion} {Detalle}", quien, accion, detalle);
        try
        {
            Directory.CreateDirectory(Ruta);
            lock (_lock) File.AppendAllText(ArchivoAuditoria, linea, new UTF8Encoding(false));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _logger.LogWarning(ex, "No se pudo escribir la auditoría de configuración.");
        }
    }
}
