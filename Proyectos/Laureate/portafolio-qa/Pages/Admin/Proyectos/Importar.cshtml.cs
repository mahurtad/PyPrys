using System.Globalization;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Proyectos;

public class ResultadoProyecto
{
    public int Fila { get; set; }
    public string Nombre { get; set; } = "";
    public string Institucion { get; set; } = "";
    public string Estado { get; set; } = "";   // Creado | Omitido | Error
    public string Detalle { get; set; } = "";
}

/// <summary>
/// Importa proyectos NUEVOS: todos entran Planificados al 0 % (el estado se deriva del avance y el avance sale de las
/// actividades). Las tres actividades predeterminadas se crean al importar; el equipo y los hitos se cargan después en pantalla.
/// </summary>
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

    public List<ResultadoProyecto> Resultados { get; set; } = new();
    public string? Error { get; set; }
    public List<string> InstitucionesActivas { get; set; } = new();
    public List<string> SistemasActivos { get; set; } = new();

    public async Task OnGetAsync() => await CargarCatalogosAsync();

    public async Task<IActionResult> OnGetPlantillaAsync()
    {
        // Los ejemplos usan datos reales del catálogo para que la plantilla se pueda importar tal cual.
        var institucion = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).Select(i => i.Nombre).FirstOrDefaultAsync() ?? "Institución A";
        var sistemas = await _context.Sistemas.Where(s => s.Activo).OrderBy(s => s.Nombre).Select(s => s.Nombre).Take(2).ToListAsync();
        var hoy = DateTime.Today;

        var sb = new StringBuilder();
        sb.AppendLine("Nombre;Institucion;Tipo;FechaInicio;FechaFin;PM;Lider;Descripcion;Sistemas;Activo");
        sb.AppendLine($"Proyecto de ejemplo 1;{institucion};Proyecto;{hoy:dd/MM/yyyy};{hoy.AddMonths(2):dd/MM/yyyy};Nombre PM;Nombre líder;Descripción breve;{string.Join('|', sistemas)};Sí");
        sb.AppendLine($"Proyecto de ejemplo 2;{institucion};BAU;{hoy:dd/MM/yyyy};{hoy.AddMonths(1):dd/MM/yyyy};;;;;Sí");
        var bytes = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        return File(bytes, "text/csv; charset=utf-8", "plantilla_proyectos.csv");
    }

    public async Task<IActionResult> OnPostAsync()
    {
        await CargarCatalogosAsync();

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
        var faltantes = new[] { "nombre", "institucion", "fechainicio", "fechafin" }.Where(c => !indices.ContainsKey(c)).ToList();
        if (faltantes.Count > 0)
        {
            Error = "Faltan columnas obligatorias: " + string.Join(", ", faltantes) + ". Descarga la plantilla y respeta su formato.";
            return Page();
        }
        if (indices.ContainsKey("estado") || indices.ContainsKey("avance"))
        {
            Error = "El archivo no debe incluir las columnas Estado ni Avance: todos los proyectos entran Planificados al 0 %; "
                    + "el avance se calcula después con las actividades. Elimina esas columnas o usa la plantilla.";
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
        var sistemas = (await _context.Sistemas.Where(s => s.Activo).ToListAsync())
            .GroupBy(s => Normalizar(s.Nombre)).ToDictionary(g => g.Key, g => g.First());
        var existentes = (await _context.Proyectos.Select(p => new { p.InstitucionId, p.Nombre }).ToListAsync())
            .Select(p => (p.InstitucionId, Normalizar(p.Nombre))).ToHashSet();
        var vistos = new HashSet<(int, string)>();
        var nuevos = new List<Proyecto>();

        for (var n = 1; n < filas.Count; n++)
        {
            var fila = filas[n];
            var res = new ResultadoProyecto { Fila = n + 1, Nombre = Valor(fila, "nombre"), Institucion = Valor(fila, "institucion") };
            Resultados.Add(res);

            if (res.Nombre.Length == 0 || res.Nombre.Length > 200)
            {
                res.Estado = "Error"; res.Detalle = "Nombre obligatorio (máximo 200 caracteres)."; continue;
            }
            if (!instituciones.TryGetValue(Normalizar(res.Institucion), out var institucion))
            {
                res.Estado = "Error"; res.Detalle = "Institución no encontrada o inactiva."; continue;
            }

            var tipoTexto = Valor(fila, "tipo");
            var tipo = TipoIniciativa.Proyecto;
            if (tipoTexto.Length > 0)
            {
                var parsed = ParsearTipo(tipoTexto);
                if (parsed is null)
                {
                    res.Estado = "Error"; res.Detalle = "Tipo no válido (usa: Proyecto, Requerimiento no BAU, BAU)."; continue;
                }
                tipo = parsed.Value;
            }

            if (!TryFecha(Valor(fila, "fechainicio"), out var inicio))
            {
                res.Estado = "Error"; res.Detalle = "FechaInicio inválida (usa dd/MM/yyyy o yyyy-MM-dd, años 2000-2100)."; continue;
            }
            if (!TryFecha(Valor(fila, "fechafin"), out var fin))
            {
                res.Estado = "Error"; res.Detalle = "FechaFin inválida (usa dd/MM/yyyy o yyyy-MM-dd, años 2000-2100)."; continue;
            }
            if (fin < inicio)
            {
                res.Estado = "Error"; res.Detalle = "FechaFin debe ser igual o posterior a FechaInicio."; continue;
            }

            var pm = Valor(fila, "pm");
            var lider = Valor(fila, "lider");
            var descripcion = Valor(fila, "descripcion");
            if (pm.Length > 150 || lider.Length > 150) { res.Estado = "Error"; res.Detalle = "PM y Lider: máximo 150 caracteres."; continue; }
            if (descripcion.Length > 500) { res.Estado = "Error"; res.Detalle = "Descripcion: máximo 500 caracteres."; continue; }

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

            var seleccion = new List<Sistema>();
            var desconocidos = new List<string>();
            foreach (var parte in Valor(fila, "sistemas").Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (sistemas.TryGetValue(Normalizar(parte), out var sis))
                {
                    if (!seleccion.Contains(sis)) seleccion.Add(sis);
                }
                else desconocidos.Add(parte);
            }
            if (desconocidos.Count > 0)
            {
                res.Estado = "Error"; res.Detalle = "Sistema no encontrado o inactivo: " + string.Join(", ", desconocidos) + "."; continue;
            }

            var clave = (institucion.Id, Normalizar(res.Nombre));
            if (!vistos.Add(clave)) { res.Estado = "Omitido"; res.Detalle = "Proyecto repetido en el archivo (mismo nombre e institución)."; continue; }
            if (existentes.Contains(clave)) { res.Estado = "Omitido"; res.Detalle = "Ya existe un proyecto con ese nombre en la institución."; continue; }

            nuevos.Add(new Proyecto
            {
                Nombre = res.Nombre,
                Descripcion = descripcion.Length == 0 ? null : descripcion,
                InstitucionId = institucion.Id,
                Tipo = tipo,
                Estado = EstadoProyecto.Planificado,
                Avance = 0,
                PM = pm.Length == 0 ? null : pm,
                Lider = lider.Length == 0 ? null : lider,
                FechaInicio = inicio,
                FechaFin = fin,
                Activo = activo,
                Sistemas = seleccion.Select(s => new ProyectoSistema { SistemaId = s.Id }).ToList()
            });
            res.Estado = "Creado";
            res.Detalle = $"{tipo.Texto()} · {inicio:dd/MM/yyyy} – {fin:dd/MM/yyyy}"
                          + (seleccion.Count > 0 ? " · " + string.Join(", ", seleccion.Select(s => s.Nombre)) : "")
                          + (activo ? "" : " · inactivo");
        }

        if (nuevos.Count > 0)
        {
            foreach (var proyecto in nuevos) proyecto.Actividades = ActividadesPredeterminadas.Crear(proyecto);
            _context.Proyectos.AddRange(nuevos);
            await _context.SaveChangesAsync();
        }

        return Page();
    }

    private async Task CargarCatalogosAsync()
    {
        InstitucionesActivas = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).Select(i => i.Nombre).ToListAsync();
        SistemasActivos = await _context.Sistemas.Where(s => s.Activo).OrderBy(s => s.Nombre).Select(s => s.Nombre).ToListAsync();
    }

    private static bool TryFecha(string texto, out DateTime fecha) =>
        DateTime.TryParseExact(texto, Formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha)
        && fecha.Year >= 2000 && fecha.Year <= 2100;

    private static TipoIniciativa? ParsearTipo(string texto) => Normalizar(texto) switch
    {
        "proyecto" => TipoIniciativa.Proyecto,
        "requerimientonobau" or "requerimiento" or "nobau" => TipoIniciativa.RequerimientoNoBAU,
        "bau" => TipoIniciativa.BAU,
        _ => null
    };

    private static string Normalizar(string texto)
    {
        var sinAcentos = new string(texto.Normalize(NormalizationForm.FormD)
            .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray());
        return new string(sinAcentos.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();
    }
}
