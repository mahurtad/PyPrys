using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Usuarios;

public class ResultadoImportacion
{
    public int Fila { get; set; }
    public string Correo { get; set; } = "";
    public string Estado { get; set; } = "";   // Creado | Omitido | Error
    public string Detalle { get; set; } = "";
    public string? ContrasenaGenerada { get; set; }
}

public class ImportarModel : PageModel
{
    private const int MaxFilas = 500;
    private const long MaxBytes = 1_000_000;

    private static readonly string[] Cabecera = { "Nombres", "Apellidos", "Correo", "Rol", "Contrasena", "Equipo", "Capacidad" };

    private readonly UserManager<UsuarioApp> _usuarios;
    private readonly PortafolioQAContext _context;

    public ImportarModel(UserManager<UsuarioApp> usuarios, PortafolioQAContext context)
    {
        _usuarios = usuarios;
        _context = context;
    }

    [BindProperty]
    public IFormFile? Archivo { get; set; }

    public List<ResultadoImportacion> Resultados { get; set; } = new();
    public string? Error { get; set; }
    public string[] RolesDisponibles => Roles.Todos;

    public void OnGet()
    {
    }

    public IActionResult OnGetPlantilla()
    {
        var sb = new StringBuilder();
        sb.AppendLine(string.Join(';', Cabecera));
        sb.AppendLine("Ana;Pérez Gómez;ana.perez@ejemplo.com;Editor;;Change;100");
        sb.AppendLine("Luis;Rojas Vega;luis.rojas@ejemplo.com;Lector;;;");
        var bytes = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        return File(bytes, "text/csv; charset=utf-8", "plantilla_usuarios.csv");
    }

    public async Task<IActionResult> OnPostAsync()
    {
        // Las contraseñas generadas se muestran una sola vez: la respuesta no debe quedar en caché.
        Response.Headers.CacheControl = "no-store";

        if (Archivo is null || Archivo.Length == 0)
        {
            Error = "Selecciona un archivo CSV.";
            return Page();
        }
        if (Archivo.Length > MaxBytes || !Archivo.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
        {
            Error = "El archivo debe ser un .csv de máximo 1 MB.";
            return Page();
        }

        string texto;
        using (var lector = new StreamReader(Archivo.OpenReadStream(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
        {
            texto = await lector.ReadToEndAsync();
        }

        var filas = CsvSimple.Leer(texto);
        if (filas.Count < 2)
        {
            Error = "El archivo no tiene filas de datos. Descarga la plantilla y respeta su formato.";
            return Page();
        }

        var indices = filas[0].Select((nombre, i) => (Nombre: Normalizar(nombre), i))
            .Where(x => x.Nombre.Length > 0).GroupBy(x => x.Nombre).ToDictionary(g => g.Key, g => g.First().i);
        var faltantes = new[] { "nombres", "apellidos", "correo", "rol" }.Where(c => !indices.ContainsKey(c)).ToList();
        if (faltantes.Count > 0)
        {
            Error = "Faltan columnas obligatorias: " + string.Join(", ", faltantes) + ". Descarga la plantilla y respeta su formato.";
            return Page();
        }
        if (filas.Count - 1 > MaxFilas)
        {
            Error = $"El archivo supera el máximo de {MaxFilas} filas por importación.";
            return Page();
        }

        string Valor(string[] fila, string col) => indices.TryGetValue(col, out var i) && i < fila.Length ? fila[i].Trim() : "";

        var correosVistos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var n = 1; n < filas.Count; n++)
        {
            var fila = filas[n];
            var numero = n + 1; // número de línea en el archivo (la 1 es la cabecera)
            var res = new ResultadoImportacion { Fila = numero, Correo = Valor(fila, "correo") };
            Resultados.Add(res);

            var nombres = Valor(fila, "nombres");
            var apellidos = Valor(fila, "apellidos");
            var rolTexto = Valor(fila, "rol");
            var rol = Roles.Todos.FirstOrDefault(r => r.Equals(rolTexto, StringComparison.OrdinalIgnoreCase));

            string? problema = null;
            if (nombres.Length == 0 || nombres.Length > 100) problema = "Nombres obligatorios (máximo 100 caracteres).";
            else if (apellidos.Length == 0 || apellidos.Length > 100) problema = "Apellidos obligatorios (máximo 100 caracteres).";
            else if (res.Correo.Length == 0 || res.Correo.Length > 150 || !new EmailAddressAttribute().IsValid(res.Correo)) problema = "Correo inválido.";
            else if (rol is null) problema = $"Rol no válido (usa: {string.Join(", ", Roles.Todos)}).";
            if (problema is not null) { res.Estado = "Error"; res.Detalle = problema; continue; }

            if (!correosVistos.Add(res.Correo)) { res.Estado = "Omitido"; res.Detalle = "Correo repetido en el archivo."; continue; }
            if (await _usuarios.FindByEmailAsync(res.Correo) is not null) { res.Estado = "Omitido"; res.Detalle = "Ya existe un usuario con ese correo."; continue; }

            // Perfil de recurso QA opcional: se crea solo si se indica Equipo.
            EquipoQA? equipo = null;
            decimal capacidad = 100;
            var equipoTexto = Valor(fila, "equipo");
            if (equipoTexto.Length > 0)
            {
                equipo = ParsearEquipo(equipoTexto);
                if (equipo is null) { res.Estado = "Error"; res.Detalle = "Equipo no válido (usa Change o RunBau)."; continue; }
                var capTexto = Valor(fila, "capacidad").TrimEnd('%').Trim();
                if (capTexto.Length > 0 && (!decimal.TryParse(capTexto.Replace(',', '.'), NumberStyles.Number, CultureInfo.InvariantCulture, out capacidad) || capacidad < 0 || capacidad > 100))
                {
                    res.Estado = "Error"; res.Detalle = "Capacidad debe ser un número entre 0 y 100."; continue;
                }
                if (await _context.RecursosQA.AnyAsync(r => r.Correo == res.Correo && r.UsuarioId != null))
                {
                    res.Estado = "Omitido"; res.Detalle = "Ya existe un recurso QA con ese correo."; continue;
                }
            }

            var contrasena = Valor(fila, "contrasena");
            var generada = contrasena.Length == 0;
            if (generada) contrasena = GenerarContrasena();

            var usuario = new UsuarioApp
            {
                UserName = res.Correo, Email = res.Correo, EmailConfirmed = true,
                Nombres = nombres, Apellidos = apellidos, Activo = true
            };
            var crear = await _usuarios.CreateAsync(usuario, contrasena);
            if (!crear.Succeeded)
            {
                res.Estado = "Error";
                res.Detalle = string.Join(" ", crear.Errors.Select(e => e.Description));
                continue;
            }
            await _usuarios.AddToRoleAsync(usuario, rol!);

            if (equipo.HasValue)
            {
                // Si ya existía un recurso sin usuario con ese correo, se vincula en lugar de duplicarlo.
                var huerfano = await _context.RecursosQA.FirstOrDefaultAsync(r => r.Correo == res.Correo && r.UsuarioId == null);
                if (huerfano is not null)
                {
                    huerfano.UsuarioId = usuario.Id;
                    huerfano.Nombres = nombres;
                    huerfano.Apellidos = apellidos;
                    huerfano.Equipo = equipo.Value;
                    huerfano.CapacidadDisponible = capacidad;
                }
                else
                {
                    _context.RecursosQA.Add(new RecursoQA
                    {
                        UsuarioId = usuario.Id, Nombres = nombres, Apellidos = apellidos, Correo = res.Correo,
                        Equipo = equipo.Value, CapacidadDisponible = capacidad, Activo = true
                    });
                }
                await _context.SaveChangesAsync();
            }

            res.Estado = "Creado";
            res.Detalle = $"Rol {rol}" + (equipo.HasValue ? $" · recurso QA {equipo.Value.Texto()} ({capacidad:0.##}%)" : "");
            if (generada) res.ContrasenaGenerada = contrasena;
        }

        return Page();
    }

    private static string Normalizar(string texto)
    {
        var sinAcentos = new string(texto.Normalize(NormalizationForm.FormD)
            .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray());
        return new string(sinAcentos.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();
    }

    private static EquipoQA? ParsearEquipo(string texto) => Normalizar(texto) switch
    {
        "change" => EquipoQA.Change,
        "runbau" or "run" or "bau" => EquipoQA.RunBau,
        _ => null
    };

    private static string GenerarContrasena()
    {
        const string mayus = "ABCDEFGHJKLMNPQRSTUVWXYZ", minus = "abcdefghijkmnopqrstuvwxyz", digitos = "23456789";
        var todos = mayus + minus + digitos;
        var chars = new List<char> { Elegir(mayus), Elegir(minus), Elegir(digitos) };
        while (chars.Count < 14) chars.Add(Elegir(todos));
        // Mezcla Fisher-Yates con RNG criptográfico.
        for (var i = chars.Count - 1; i > 0; i--)
        {
            var j = RandomNumberGenerator.GetInt32(i + 1);
            (chars[i], chars[j]) = (chars[j], chars[i]);
        }
        return new string(chars.ToArray());
    }

    private static char Elegir(string conjunto) => conjunto[RandomNumberGenerator.GetInt32(conjunto.Length)];
}
