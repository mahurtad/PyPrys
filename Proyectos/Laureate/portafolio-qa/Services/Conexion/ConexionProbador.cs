using System.Net.NetworkInformation;
using System.Text;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Migrations;
using PortafolioQA.Data;

namespace PortafolioQA.Services.Conexion;

public record DatosConexion(string Servidor, int? Puerto, string Base, bool UsaSql, string? Usuario, string? Clave, bool ConfiarCertificado = true);

public class ResultadoPrueba
{
    public bool Conecta { get; set; }
    public string? Error { get; set; }
    public bool BaseExiste { get; set; }
    public string? Version { get; set; }
    public bool Lectura { get; set; }
    public bool Escritura { get; set; }
    public bool PuedeAlterarEsquema { get; set; }
    /// <summary>La base tiene tablas pero no el historial de migraciones: no es una base de Portafolio QA.</summary>
    public bool BaseAjena { get; set; }
    public List<string> Pendientes { get; set; } = new();
    public int Aplicadas { get; set; }
    /// <summary>El servidor rechazó el login o la base (el script para el DBA es la ayuda adecuada).</summary>
    public bool FaltaAcceso { get; set; }

    public bool PermisosOk => Lectura && Escritura;
    public bool EsquemaAlDia => Conecta && BaseExiste && Pendientes.Count == 0;
    public bool ListaParaActivar => Conecta && BaseExiste && PermisosOk && EsquemaAlDia && !BaseAjena;
}

/// <summary>Prueba la conexión y los permisos, aplica migraciones y genera los scripts (esquema y acceso para el DBA).</summary>
public class ConexionProbador
{
    private const string CadenaSinConexion = "Server=.;Database=PortafolioQA;Connect Timeout=1;TrustServerCertificate=True";

    public static string Construir(DatosConexion d, int? timeout = null)
    {
        var b = new SqlConnectionStringBuilder
        {
            DataSource = d.Puerto is int p ? $"{d.Servidor.Trim()},{p}" : d.Servidor.Trim(),
            InitialCatalog = d.Base.Trim(),
            TrustServerCertificate = d.ConfiarCertificado,
            MultipleActiveResultSets = true
        };
        if (d.UsaSql) { b.UserID = d.Usuario?.Trim() ?? string.Empty; b.Password = d.Clave ?? string.Empty; }
        else b.IntegratedSecurity = true;
        if (timeout is int t) b.ConnectTimeout = t;
        return b.ConnectionString;
    }

    /// <summary>Datos no secretos de una cadena existente, para precargar el formulario.</summary>
    public static DatosConexion? Descomponer(string? cadena)
    {
        if (string.IsNullOrWhiteSpace(cadena)) return null;
        try
        {
            var b = new SqlConnectionStringBuilder(cadena);
            var origen = b.DataSource;
            int? puerto = null;
            var coma = origen.LastIndexOf(',');
            if (coma > 0 && int.TryParse(origen[(coma + 1)..], out var p)) { puerto = p; origen = origen[..coma]; }
            return new DatosConexion(origen, puerto, b.InitialCatalog, !b.IntegratedSecurity, b.UserID, null, b.TrustServerCertificate);
        }
        catch (ArgumentException) { return null; }
    }

    public static bool EsServidorLocal(string? servidor)
    {
        var host = (servidor ?? string.Empty).Trim().ToLowerInvariant();
        if (host.StartsWith("tcp:")) host = host[4..];
        host = host.Split('\\')[0].Split(',')[0].Trim();
        if (host is "." or "(local)" or "localhost" or "127.0.0.1" or "::1" || host.StartsWith("(localdb)")) return true;
        var equipo = Environment.MachineName.ToLowerInvariant();
        return host == equipo || host.StartsWith(equipo + ".");
    }

    /// <summary>Cuenta de Windows con la que corre la aplicación (la que llega a SQL con autenticación Windows en el mismo equipo).</summary>
    public static string IdentidadLocal() =>
        OperatingSystem.IsWindows() ? System.Security.Principal.WindowsIdentity.GetCurrent().Name : Environment.UserName;

    /// <summary>Con SQL en otro equipo la aplicación se presenta como la cuenta de máquina del servidor web.</summary>
    public static string CuentaRemota()
    {
        var dominio = IPGlobalProperties.GetIPGlobalProperties().DomainName;
        var neta = string.IsNullOrWhiteSpace(dominio) ? "DOMINIO" : dominio.Split('.')[0].ToUpperInvariant();
        return $"{neta}\\{Environment.MachineName}$";
    }

    public async Task<ResultadoPrueba> ProbarAsync(DatosConexion d, CancellationToken ct = default)
    {
        var r = new ResultadoPrueba();
        var cadena = Construir(d, 8);
        try
        {
            await using var con = new SqlConnection(cadena);
            await con.OpenAsync(ct);
            r.Conecta = true;
            r.BaseExiste = true;
            r.Version = con.ServerVersion;

            await using var cmd = con.CreateCommand();
            cmd.CommandText = @"SELECT
                ISNULL(HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'SELECT'), 0),
                ISNULL(HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'INSERT'), 0),
                ISNULL(HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'UPDATE'), 0),
                ISNULL(HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'DELETE'), 0),
                ISNULL(HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'CREATE TABLE'), 0),
                ISNULL(HAS_PERMS_BY_NAME('dbo', 'SCHEMA', 'ALTER'), 0),
                (SELECT COUNT(*) FROM sys.tables),
                CASE WHEN OBJECT_ID(N'dbo.__EFMigrationsHistory') IS NULL THEN 0 ELSE 1 END";
            await using (var rd = await cmd.ExecuteReaderAsync(ct))
            {
                if (await rd.ReadAsync(ct))
                {
                    r.Lectura = rd.GetInt32(0) == 1;
                    r.Escritura = rd.GetInt32(1) == 1 && rd.GetInt32(2) == 1 && rd.GetInt32(3) == 1;
                    r.PuedeAlterarEsquema = rd.GetInt32(4) == 1 && rd.GetInt32(5) == 1;
                    r.BaseAjena = rd.GetInt32(7) == 0 && rd.GetInt32(6) > 0;
                }
            }
        }
        catch (SqlException ex)
        {
            r.Error = Describir(ex, d);
            r.FaltaAcceso = ex.Number is 18456 or 4060 or 916;
            if (ex.Number is 4060 or 916) r.Conecta = true; // el servidor respondió; falla la base
            return r;
        }
        catch (Exception ex)
        {
            r.Error = ex.Message;
            return r;
        }

        if (r.Lectura && !r.BaseAjena)
        {
            try
            {
                var (pendientes, aplicadas) = await EstadoEsquemaAsync(Construir(d, 8), ct);
                r.Pendientes = pendientes;
                r.Aplicadas = aplicadas;
            }
            catch (Exception ex)
            {
                r.Error = "No se pudo comprobar el esquema: " + ex.Message;
            }
        }
        return r;
    }

    private static string Describir(SqlException ex, DatosConexion d)
    {
        // Error de TLS: el servidor SQL presenta un certificado que este equipo no reconoce (típico en certificados internos o autofirmados).
        if (ex.Message.Contains("certificate chain", StringComparison.OrdinalIgnoreCase) || ex.Message.Contains("certificado", StringComparison.OrdinalIgnoreCase))
            return "El certificado del servidor SQL no es de confianza para este equipo. Marque \"Confiar en el certificado del servidor SQL\" y pruebe de nuevo " +
                   "(o instale la entidad certificadora en este servidor). Detalle: " + ex.Message;
        return DescribirNumero(ex, d);
    }

    private static string DescribirNumero(SqlException ex, DatosConexion d) => ex.Number switch
    {
        18456 => "El servidor rechazó el inicio de sesión: usuario o contraseña incorrectos, o la cuenta no tiene acceso al servidor.",
        4060 or 916 => $"La base '{d.Base}' no existe o la cuenta no tiene acceso a ella.",
        -1 or 2 or 53 or 258 or 10060 or 40 => "No se pudo contactar al servidor SQL. Verifique el nombre, el puerto y que el servicio acepte conexiones.",
        _ => ex.Message
    };

    public async Task<(List<string> Pendientes, int Aplicadas)> EstadoEsquemaAsync(string cadena, CancellationToken ct = default)
    {
        await using var ctx = CrearContexto(cadena);
        var pendientes = (await ctx.Database.GetPendingMigrationsAsync(ct)).ToList();
        var aplicadas = (await ctx.Database.GetAppliedMigrationsAsync(ct)).Count();
        return (pendientes, aplicadas);
    }

    public async Task MigrarAsync(string cadena, CancellationToken ct = default)
    {
        await using var ctx = CrearContexto(cadena);
        await ctx.Database.MigrateAsync(ct);
    }

    /// <summary>Script SQL idempotente con todo el esquema, para que lo ejecute el administrador de la base.</summary>
    public string GenerarScriptEsquema()
    {
        using var ctx = CrearContexto(CadenaSinConexion);
        return ctx.GetInfrastructure().GetRequiredService<IMigrator>().GenerateScript(options: MigrationsSqlGenerationOptions.Idempotent);
    }

    /// <summary>Script T-SQL para el DBA: crea la base si falta y otorga a la cuenta lectura y escritura. No se ejecuta solo.</summary>
    public static string GenerarScriptAcceso(DatosConexion d, string cuentaWindows)
    {
        static string Id(string s) => s.Replace("]", "]]");
        static string Lit(string s) => s.Replace("'", "''");
        var baseDatos = d.Base.Trim();
        var cuenta = d.UsaSql ? (d.Usuario ?? "usuario_portafolio").Trim() : cuentaWindows;
        var sb = new StringBuilder();
        sb.AppendLine("-- Script de acceso para Portafolio QA. Ejecútelo con una cuenta con permiso para crear logins y usuarios (p. ej. sysadmin).");
        sb.AppendLine("-- No contiene contraseñas. Revíselo antes de ejecutarlo.");
        sb.AppendLine("USE [master];");
        sb.AppendLine($"IF DB_ID(N'{Lit(baseDatos)}') IS NULL CREATE DATABASE [{Id(baseDatos)}];");
        sb.AppendLine("GO");
        if (d.UsaSql)
        {
            sb.AppendLine("-- La contraseña la define y custodia el administrador de la base; luego se escribe en el asistente.");
            sb.AppendLine($"IF SUSER_ID(N'{Lit(cuenta)}') IS NULL CREATE LOGIN [{Id(cuenta)}] WITH PASSWORD = N'<ESCRIBA_UNA_CONTRASEÑA_SEGURA>';");
        }
        else
        {
            sb.AppendLine($"IF SUSER_ID(N'{Lit(cuenta)}') IS NULL CREATE LOGIN [{Id(cuenta)}] FROM WINDOWS;");
        }
        sb.AppendLine("GO");
        sb.AppendLine($"USE [{Id(baseDatos)}];");
        sb.AppendLine($"IF USER_ID(N'{Lit(cuenta)}') IS NULL CREATE USER [{Id(cuenta)}] FOR LOGIN [{Id(cuenta)}];");
        sb.AppendLine($"ALTER ROLE db_datareader ADD MEMBER [{Id(cuenta)}];");
        sb.AppendLine($"ALTER ROLE db_datawriter ADD MEMBER [{Id(cuenta)}];");
        sb.AppendLine("-- Opcional: permite que el asistente de la aplicación aplique las migraciones por sí mismo.");
        sb.AppendLine("-- Si no se otorga, el esquema se aplica ejecutando el script de esquema que ofrece el asistente.");
        sb.AppendLine($"-- ALTER ROLE db_ddladmin ADD MEMBER [{Id(cuenta)}];");
        sb.AppendLine("GO");
        return sb.ToString();
    }

    private static PortafolioQAContext CrearContexto(string cadena) =>
        new(new DbContextOptionsBuilder<PortafolioQAContext>().UseSqlServer(cadena).Options);
}
