using System.Globalization;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Sistemas;

public class ResultadoSistema
{
    public int Fila { get; set; }
    public string Nombre { get; set; } = "";
    public string Estado { get; set; } = "";   // Creado | Omitido | Error
    public string Detalle { get; set; } = "";
}

public class ImportarModel : PageModel
{
    private const int MaxFilas = 500;
    private const long MaxBytes = 1_000_000;

    private readonly PortafolioQAContext _context;

    public ImportarModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public IFormFile? Archivo { get; set; }

    public List<ResultadoSistema> Resultados { get; set; } = new();
    public string? Error { get; set; }
    public List<string> InstitucionesActivas { get; set; } = new();

    public async Task OnGetAsync() => await CargarInstitucionesAsync();

    public async Task<IActionResult> OnGetPlantillaAsync()
    {
        // Los ejemplos usan instituciones reales del catálogo para que la plantilla se pueda importar tal cual.
        var nombres = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).Select(i => i.Nombre).Take(2).ToListAsync();
        var a = nombres.ElementAtOrDefault(0) ?? "Institución A";
        var b = nombres.ElementAtOrDefault(1) ?? a;
        var sb = new StringBuilder();
        sb.AppendLine("Nombre;Instituciones;Activo");
        sb.AppendLine($"Sistema de ejemplo 1;{a}|{b};Sí".Replace($"|{a};", ";"));
        sb.AppendLine($"Sistema de ejemplo 2;{a};Sí");
        sb.AppendLine("Sistema en desuso;;No");
        var bytes = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        return File(bytes, "text/csv; charset=utf-8", "plantilla_sistemas.csv");
    }

    public async Task<IActionResult> OnPostAsync()
    {
        await CargarInstitucionesAsync();

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
        if (!indices.ContainsKey("nombre"))
        {
            Error = "Falta la columna obligatoria: Nombre. Descarga la plantilla y respeta su formato.";
            return Page();
        }
        if (filas.Count - 1 > MaxFilas)
        {
            Error = $"El archivo supera el máximo de {MaxFilas} filas por importación.";
            return Page();
        }

        string Valor(string[] fila, string col) => indices.TryGetValue(col, out var i) && i < fila.Length ? fila[i].Trim() : "";

        var instituciones = (await _context.Instituciones.Where(i => i.Activo).ToListAsync())
            .GroupBy(i => Normalizar(i.Nombre)).ToDictionary(g => g.Key, g => g.First());
        var existentes = (await _context.Sistemas.Select(s => s.Nombre).ToListAsync())
            .Select(Normalizar).ToHashSet();
        var vistos = new HashSet<string>();
        var nuevos = new List<Sistema>();

        for (var n = 1; n < filas.Count; n++)
        {
            var fila = filas[n];
            var res = new ResultadoSistema { Fila = n + 1, Nombre = Valor(fila, "nombre") };
            Resultados.Add(res);

            if (res.Nombre.Length == 0 || res.Nombre.Length > 150)
            {
                res.Estado = "Error"; res.Detalle = "Nombre obligatorio (máximo 150 caracteres)."; continue;
            }

            var activoTexto = Valor(fila, "activo");
            var activo = true;
            if (activoTexto.Length > 0)
            {
                switch (Normalizar(activoTexto))
                {
                    case "si" or "s" or "true" or "1" or "activo": activo = true; break;
                    case "no" or "n" or "false" or "0" or "inactivo": activo = false; break;
                    default:
                        res.Estado = "Error"; res.Detalle = "Activo debe ser Sí o No."; continue;
                }
            }

            var seleccion = new List<Institucion>();
            var desconocidas = new List<string>();
            foreach (var parte in Valor(fila, "instituciones").Split(new[] { '|', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (instituciones.TryGetValue(Normalizar(parte), out var inst))
                {
                    if (!seleccion.Contains(inst)) seleccion.Add(inst);
                }
                else desconocidas.Add(parte);
            }
            if (desconocidas.Count > 0)
            {
                res.Estado = "Error";
                res.Detalle = "Institución no encontrada o inactiva: " + string.Join(", ", desconocidas) + ".";
                continue;
            }

            var clave = Normalizar(res.Nombre);
            if (!vistos.Add(clave)) { res.Estado = "Omitido"; res.Detalle = "Nombre repetido en el archivo."; continue; }
            if (existentes.Contains(clave)) { res.Estado = "Omitido"; res.Detalle = "Ya existe un sistema con ese nombre."; continue; }

            nuevos.Add(new Sistema
            {
                Nombre = res.Nombre,
                Activo = activo,
                Instituciones = seleccion.Select(i => new SistemaInstitucion { InstitucionId = i.Id }).ToList()
            });
            res.Estado = "Creado";
            res.Detalle = (activo ? "Activo" : "Inactivo")
                          + (seleccion.Count > 0 ? " · " + string.Join(", ", seleccion.Select(i => i.Nombre)) : " · sin instituciones");
        }

        if (nuevos.Count > 0)
        {
            _context.Sistemas.AddRange(nuevos);
            await _context.SaveChangesAsync();
        }

        return Page();
    }

    private async Task CargarInstitucionesAsync() =>
        InstitucionesActivas = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).Select(i => i.Nombre).ToListAsync();

    private static string Normalizar(string texto)
    {
        var sinAcentos = new string(texto.Normalize(NormalizationForm.FormD)
            .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray());
        return new string(sinAcentos.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();
    }
}
