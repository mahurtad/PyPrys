using System.Globalization;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Feriados;

public class ResultadoFeriado
{
    public int Fila { get; set; }
    public string Fecha { get; set; } = "";
    public string Nombre { get; set; } = "";
    public string Estado { get; set; } = "";   // Creado | Omitido | Error
    public string Detalle { get; set; } = "";
}

public class ImportarModel : PageModel
{
    private const int MaxFilas = 500;
    private const long MaxBytes = 1_000_000;

    private static readonly string[] Formatos = { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "d-M-yyyy", "yyyy/MM/dd" };

    private readonly PortafolioQAContext _context;

    public ImportarModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public IFormFile? Archivo { get; set; }

    public List<ResultadoFeriado> Resultados { get; set; } = new();
    public string? Error { get; set; }

    public void OnGet()
    {
    }

    public IActionResult OnGetPlantilla()
    {
        var sb = new StringBuilder();
        sb.AppendLine("Fecha;Nombre");
        sb.AppendLine("01/01/2027;Año Nuevo");
        sb.AppendLine("01/05/2027;Día del Trabajo");
        sb.AppendLine("28/07/2027;Fiestas Patrias");
        var bytes = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        return File(bytes, "text/csv; charset=utf-8", "plantilla_feriados.csv");
    }

    public async Task<IActionResult> OnPostAsync()
    {
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
        var faltantes = new[] { "fecha", "nombre" }.Where(c => !indices.ContainsKey(c)).ToList();
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

        var existentes = (await _context.Feriados.Select(f => f.Fecha).ToListAsync()).Select(f => f.Date).ToHashSet();
        var vistas = new HashSet<DateTime>();
        var nuevos = new List<Feriado>();

        for (var n = 1; n < filas.Count; n++)
        {
            var fila = filas[n];
            var res = new ResultadoFeriado { Fila = n + 1, Fecha = Valor(fila, "fecha"), Nombre = Valor(fila, "nombre") };
            Resultados.Add(res);

            if (!DateTime.TryParseExact(res.Fecha, Formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, out var fecha)
                || fecha.Year < 2000 || fecha.Year > 2100)
            {
                res.Estado = "Error"; res.Detalle = "Fecha inválida (usa dd/MM/yyyy o yyyy-MM-dd, años 2000-2100)."; continue;
            }
            if (res.Nombre.Length == 0 || res.Nombre.Length > 150)
            {
                res.Estado = "Error"; res.Detalle = "Nombre obligatorio (máximo 150 caracteres)."; continue;
            }
            if (!vistas.Add(fecha.Date)) { res.Estado = "Omitido"; res.Detalle = "Fecha repetida en el archivo."; continue; }
            if (existentes.Contains(fecha.Date)) { res.Estado = "Omitido"; res.Detalle = "Ya existe un feriado en esa fecha."; continue; }

            nuevos.Add(new Feriado { Fecha = fecha.Date, Nombre = res.Nombre });
            res.Estado = "Creado";
            res.Detalle = CultureInfo.GetCultureInfo("es-PE").DateTimeFormat.GetDayName(fecha.DayOfWeek)
                          + (fecha.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday ? " (fin de semana: ya no es laborable)" : "");
        }

        if (nuevos.Count > 0)
        {
            _context.Feriados.AddRange(nuevos);
            await _context.SaveChangesAsync();
        }

        return Page();
    }

    private static string Normalizar(string texto)
    {
        var sinAcentos = new string(texto.Normalize(NormalizationForm.FormD)
            .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray());
        return new string(sinAcentos.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();
    }
}
