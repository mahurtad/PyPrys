using PortafolioQA.Models;

namespace PortafolioQA.Services;

public static class ActividadesPredeterminadas
{
    public static readonly IReadOnlyList<(string Nombre, decimal Peso)> Definiciones =
        Array.AsReadOnly(new[] {
            ("Diseño y redacción de Casos de Prueba", 30m),
            ("Ejecución de Casos de prueba", 50m),
            ("Elaborar informe", 20m)
        });

    // Días inclusivos. Con menos de tres días las actividades comparten un día.
    public static List<Actividad> Crear(Proyecto proyecto)
    {
        var inicio = proyecto.FechaInicio.Date;
        var dias = (proyecto.FechaFin.Date - inicio).Days + 1;
        if (dias < 1) throw new ArgumentException("El fin debe ser posterior o igual al inicio.");
        var resultado = new List<Actividad>();
        var anterior = 0;
        decimal acumulado = 0;
        for (var i = 0; i < Definiciones.Count; i++)
        {
            var (nombre, peso) = Definiciones[i];
            acumulado += peso;
            var limite = dias >= 3
                ? Math.Clamp((int)Math.Round(dias * acumulado / 100, MidpointRounding.AwayFromZero), anterior + 1, dias - (2 - i))
                : Math.Clamp((int)Math.Ceiling(dias * acumulado / 100), 1, dias);
            resultado.Add(new Actividad {
                ProyectoId = proyecto.Id, Nombre = nombre, Peso = peso,
                EsPredeterminada = true, Orden = i + 1,
                FechaInicio = inicio.AddDays(Math.Min(anterior, dias - 1)),
                FechaFin = inicio.AddDays(limite - 1)
            });
            anterior = limite;
        }
        return resultado;
    }
}
