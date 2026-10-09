using System.Globalization;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Ausencias;

public class ResultadoAusencia
{
    public int Fila { get; set; }
    public string Correo { get; set; } = "";
    public string Tipo { get; set; } = "";
    public string Periodo { get; set; } = "";
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

    public List<ResultadoAusencia> Resultados { get; set; } = new();
    public string? Error { get; set; }
    public string TiposDisponibles => string.Join(", ", Enum.GetValues<TipoAusencia>().Select(t => t.Texto()));

    public void OnGet()
    {
    }

    public IActionResult OnGetPlantilla()
    {
        var sb = new StringBuilder();
        sb.AppendLine("Correo;Tipo;FechaInicio;FechaFin;Motivo");
        sb.AppendLine("ana.perez@ejemplo.com;Vacaciones;02/02/2027;13/02/2027;");
        sb.AppendLine("ana.perez@ejemplo.com;Cumpleaños;17/06/2027;;");
        sb.AppendLine("luis.rojas@ejemplo.com;Licencia médica;10/03/2027;12/03/2027;Descanso médico");
        sb.AppendLine("luis.rojas@ejemplo.com;Compensación;21/05/2027;21/05/2027;Horas extra de abril");
        sb.AppendLine("rosa.vega@ejemplo.com;Licencia maternidad;01/08/2027;07/11/2027;");
        var bytes = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        return File(bytes, "text/csv; charset=utf-8", "plantilla_ausencias.csv");
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
        var faltantes = new[] { "correo", "tipo", "fechainicio" }.Where(c => !indices.ContainsKey(c)).ToList();
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

        var recursos = (await _context.RecursosQA.Where(r => r.Activo).ToListAsync())
            .GroupBy(r => r.Correo.ToLowerInvariant()).ToDictionary(g => g.Key, g => g.First());
        var existentes = await _context.Ausencias.AsNoTracking().ToListAsync();
        var nuevas = new List<Ausencia>();

        for (var n = 1; n < filas.Count; n++)
        {
            var fila = filas[n];
            var res = new ResultadoAusencia { Fila = n + 1, Correo = Valor(fila, "correo"), Tipo = Valor(fila, "tipo") };
            Resultados.Add(res);

            if (!recursos.TryGetValue(res.Correo.ToLowerInvariant(), out var recurso))
            {
                res.Estado = "Error"; res.Detalle = "No existe un recurso QA activo con ese correo."; continue;
            }
            var tipo = ParsearTipo(res.Tipo);
            if (tipo is null)
            {
                res.Estado = "Error"; res.Detalle = $"Tipo no válido (usa: {TiposDisponibles})."; continue;
            }
            if (!TryFecha(Valor(fila, "fechainicio"), out var inicio))
            {
                res.Estado = "Error"; res.Detalle = "FechaInicio inválida (usa dd/MM/yyyy o yyyy-MM-dd, años 2000-2100)."; continue;
            }
            var finTexto = Valor(fila, "fechafin");
            var fin = inicio;
            if (finTexto.Length > 0 && !TryFecha(finTexto, out fin))
            {
                res.Estado = "Error"; res.Detalle = "FechaFin inválida (usa dd/MM/yyyy o yyyy-MM-dd, años 2000-2100)."; continue;
            }
            if (fin < inicio)
            {
                res.Estado = "Error"; res.Detalle = "FechaFin debe ser igual o posterior a FechaInicio."; continue;
            }
            var motivo = Valor(fila, "motivo");
            if (motivo.Length > 150)
            {
                res.Estado = "Error"; res.Detalle = "Motivo máximo 150 caracteres."; continue;
            }

            res.Periodo = inicio == fin ? $"{inicio:dd/MM/yyyy}" : $"{inicio:dd/MM/yyyy} – {fin:dd/MM/yyyy}";

            var mismos = existentes.Concat(nuevas).Where(a => a.RecursoQAId == recurso.Id).ToList();
            if (mismos.Any(a => a.Tipo == tipo && a.FechaInicio.Date == inicio && a.FechaFin.Date == fin))
            {
                res.Estado = "Omitido"; res.Detalle = "Ya existe esa ausencia (mismo recurso, tipo y periodo)."; continue;
            }

            nuevas.Add(new Ausencia
            {
                RecursoQAId = recurso.Id, Tipo = tipo.Value, FechaInicio = inicio, FechaFin = fin,
                Motivo = motivo.Length == 0 ? null : motivo
            });
            res.Estado = "Creado";
            res.Detalle = $"{recurso.NombreCompleto} · {tipo.Value.Texto()}"
                          + (mismos.Any(a => a.FechaInicio.Date <= fin && inicio <= a.FechaFin.Date) ? " · se solapa con otra ausencia del recurso" : "");
        }

        if (nuevas.Count > 0)
        {
            _context.Ausencias.AddRange(nuevas);
            await _context.SaveChangesAsync();
        }

        return Page();
    }

    private static bool TryFecha(string texto, out DateTime fecha) =>
        DateTime.TryParseExact(texto, Formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha)
        && fecha.Year >= 2000 && fecha.Year <= 2100;

    private static TipoAusencia? ParsearTipo(string texto) => Normalizar(texto) switch
    {
        "vacaciones" => TipoAusencia.Vacaciones,
        "compensacion" => TipoAusencia.Compensacion,
        "cumpleanos" or "cumpleano" => TipoAusencia.Cumpleanos,
        "licenciamedica" => TipoAusencia.LicenciaMedica,
        "licenciamaternidad" => TipoAusencia.LicenciaMaternidad,
        _ => null
    };

    private static string Normalizar(string texto)
    {
        // La ñ se normaliza a n antes de quitar acentos (Cumpleaños = cumpleanos).
        var sinAcentos = new string(texto.Replace('ñ', 'n').Replace('Ñ', 'N').Normalize(NormalizationForm.FormD)
            .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray());
        return new string(sinAcentos.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();
    }
}
