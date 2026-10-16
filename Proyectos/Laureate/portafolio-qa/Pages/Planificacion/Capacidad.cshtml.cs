using System.Globalization;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Planificacion;

public class CapacidadModel : PageModel
{
    private readonly CapacidadService _capacidad;

    public CapacidadModel(CapacidadService capacidad)
    {
        _capacidad = capacidad;
    }

    public DateTime Desde { get; set; }
    public DateTime Hasta { get; set; }
    public EquipoQA? Equipo { get; set; }
    /// <summary>Atajo de periodo elegido: mes, proximo o trimestre (vacío = fechas manuales).</summary>
    public string? Periodo { get; set; }
    /// <summary>Filtro de situación: todos, sobre, espacio o completo.</summary>
    public string Situacion { get; set; } = "todos";

    /// <summary>Todos los recursos del equipo/periodo; los totales y las tarjetas se calculan sobre este conjunto.</summary>
    public List<ResumenRecurso> Recursos { get; set; } = new();
    /// <summary>Recursos que se listan, tras aplicar el filtro de situación.</summary>
    public List<ResumenRecurso> Visibles { get; set; } = new();

    public decimal HorasPorDia => _capacidad.HorasPorDia;
    public string? Error { get; set; }

    public int Sobreasignados => Recursos.Count(r => Clasificar(r) == "sobre");
    public int ConEspacio => Recursos.Count(r => Clasificar(r) is "espacio" or "libre");
    public int Completos => Recursos.Count(r => Clasificar(r) == "completo");
    public decimal HorasAsignadas => Recursos.Sum(r => r.HorasAsignadas);
    public decimal HorasCapacidad => Recursos.Sum(r => r.HorasCapacidad);
    public decimal HorasLibres => Recursos.Sum(r => Math.Max(0, r.HorasCapacidad - r.HorasAsignadas));
    public decimal? OcupacionEquipo => HorasCapacidad == 0 ? null : Math.Round(HorasAsignadas * 100 / HorasCapacidad, 0);
    public int DiasLaborables => Recursos.Count == 0 ? 0 : Recursos.Max(r => r.DiasLaborables);

    public string EtiquetaPeriodo
    {
        get
        {
            var cultura = CultureInfo.GetCultureInfo("es-PE");
            if (Desde.Day == 1 && Hasta == Desde.AddMonths(1).AddDays(-1))
            {
                var texto = Desde.ToString("MMMM yyyy", cultura);
                return char.ToUpper(texto[0], cultura) + texto[1..];
            }
            return $"{Desde:dd/MM/yyyy} – {Hasta:dd/MM/yyyy}";
        }
    }

    /// <summary>Clasifica la situación de un recurso en el periodo.</summary>
    public static string Clasificar(ResumenRecurso r)
    {
        if (r.Sobreasignado) return "sobre";
        if (r.DiasLaborables == 0) return "sindias";
        if (r.CargaMaxima == 0) return "libre";
        return r.DiasConDisponibilidad > 0 ? "espacio" : "completo";
    }

    public static decimal? Ocupacion(ResumenRecurso r) =>
        r.HorasCapacidad == 0 ? null : Math.Round(r.HorasAsignadas * 100 / r.HorasCapacidad, 0);

    public async Task OnGetAsync(DateTime? desde, DateTime? hasta, EquipoQA? equipo, string? periodo, string? situacion)
    {
        var hoy = DateTime.Today;
        var inicioMes = new DateTime(hoy.Year, hoy.Month, 1);

        Periodo = periodo;
        switch (periodo)
        {
            case "mes": desde = inicioMes; hasta = inicioMes.AddMonths(1).AddDays(-1); break;
            case "proximo": desde = inicioMes.AddMonths(1); hasta = inicioMes.AddMonths(2).AddDays(-1); break;
            case "trimestre": desde = inicioMes; hasta = inicioMes.AddMonths(3).AddDays(-1); break;
            default: Periodo = null; break;
        }

        Desde = (desde ?? inicioMes).Date;
        Hasta = (hasta ?? Desde.AddMonths(1).AddDays(-1)).Date;
        Equipo = equipo;
        Situacion = situacion is "sobre" or "espacio" or "completo" ? situacion : "todos";

        if (Hasta < Desde)
        {
            Error = "La fecha final no puede ser anterior a la inicial; se restableció el periodo por defecto.";
            Desde = inicioMes;
            Hasta = Desde.AddMonths(1).AddDays(-1);
            Periodo = null;
        }
        if ((Hasta - Desde).Days + 1 > CalculadoraCapacidad.MaximoDiasPeriodo)
        {
            Error = $"El periodo máximo es de {CalculadoraCapacidad.MaximoDiasPeriodo} días; se recortó el fin.";
            Hasta = Desde.AddDays(CalculadoraCapacidad.MaximoDiasPeriodo - 1);
            Periodo = null;
        }

        Recursos = (await _capacidad.CalcularAsync(Desde, Hasta, equipo: Equipo))
            .OrderByDescending(r => r.Sobreasignado)
            .ThenByDescending(r => r.CargaPromedio)
            .ToList();

        Visibles = Situacion switch
        {
            "sobre" => Recursos.Where(r => Clasificar(r) == "sobre").ToList(),
            "espacio" => Recursos.Where(r => Clasificar(r) is "espacio" or "libre").ToList(),
            "completo" => Recursos.Where(r => Clasificar(r) == "completo").ToList(),
            _ => Recursos
        };
    }
}
