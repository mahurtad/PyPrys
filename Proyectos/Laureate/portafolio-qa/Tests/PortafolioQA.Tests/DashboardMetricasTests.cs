using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Tests;

public class DashboardMetricasTests
{
    private static SquashCaso Caso(int item, int? tc, string estado, DateTime? ultima = null) =>
        new() { SquashItemId = item, SquashTestCaseId = tc, EstadoSquash = estado, Porcentaje = SquashEstados.Porcentaje(estado), UltimaEjecucion = ultima };

    [Fact]
    public void Certificar_cuenta_cada_caso_una_vez_y_separa_porcentajes()
    {
        var casos = new[]
        {
            Caso(1, 10, "FAILURE", new DateTime(2026, 1, 1)), Caso(2, 10, "SUCCESS", new DateTime(2026, 2, 1)), // mismo caso de prueba: gana el más reciente
            Caso(3, 11, "SUCCESS"), Caso(4, 12, "FAILURE"), Caso(5, 13, "BLOCKED"), Caso(6, 14, "READY"), Caso(7, 15, "WARNING"),
            Caso(8, 16, "UNTESTABLE")
        };
        var c = DashboardMetricas.Certificar(casos);
        Assert.Equal(6, c.Total);
        Assert.Equal(1, c.Excluidos);
        Assert.Equal(4, c.Ejecutados);       // 10, 11, 12 y 15 (advertencia)
        Assert.Equal(2, c.Aprobados);
        Assert.Equal(1, c.Fallidos);
        Assert.Equal(1, c.Bloqueados);
        Assert.Equal(1, c.ConAdvertencia);
        Assert.Equal(2, c.Pendientes);
        Assert.Equal(66.7m, c.PorcentajeEjecutado);
        Assert.Equal(50m, c.TasaAprobacion);
    }

    [Fact]
    public void Estados_de_pruebas_separan_en_curso_y_sin_ejecutar_y_suman_el_total()
    {
        var c = DashboardMetricas.Certificar(new[]
        {
            Caso(1, 1, "SUCCESS"), Caso(2, 2, "SETTLED"), Caso(3, 3, "FAILURE"), Caso(4, 4, "ERROR"), Caso(5, 5, "WARNING"),
            Caso(6, 6, "BLOCKED"), Caso(7, 7, "RUNNING"), Caso(8, 8, "RUNNING"), Caso(9, 9, "READY"), Caso(10, 10, "NOT_RUN"), Caso(11, 11, "SKIPPED")
        });
        Assert.Equal(10, c.Total);
        Assert.Equal((2, 2, 1, 1, 2, 2), (c.Aprobados, c.Fallidos, c.ConAdvertencia, c.Bloqueados, c.EnCurso, c.SinEjecutar));
        Assert.Equal(c.Total, c.Aprobados + c.Fallidos + c.ConAdvertencia + c.Bloqueados + c.EnCurso + c.SinEjecutar);
        Assert.Equal(1, c.Excluidos);
    }

    [Fact]
    public void Sin_casos_o_sin_ejecuciones_los_porcentajes_son_nulos_no_cero()
    {
        var vacio = DashboardMetricas.Certificar(Array.Empty<SquashCaso>());
        Assert.Equal(0, vacio.Total);
        Assert.Null(vacio.PorcentajeEjecutado);
        Assert.Null(vacio.TasaAprobacion);
        var sinEjecutar = DashboardMetricas.Certificar(new[] { Caso(1, 1, "READY"), Caso(2, 2, "BLOCKED") });
        Assert.Equal(0m, sinEjecutar.PorcentajeEjecutado);
        Assert.Null(sinEjecutar.TasaAprobacion);
    }

    [Fact]
    public void Atencion_solo_incluye_hechos_con_regla_explicita()
    {
        var hoy = new DateTime(2026, 10, 16);
        var cert = new ResumenCertificacion(10, 5, 4, 1, 2, 0, 0);
        var def = new Pages.Portafolio.DetalleModel.ResumenDefectos(5, 3, 1, 2, 0, 0, 2);
        var hechos = DashboardMetricas.Atencion(EstadoProyecto.EnCertificacion, new DateTime(2026, 10, 10), hoy, cert, def);
        Assert.Equal(new[] { "vencido", "pruebas-bloqueadas", "defectos-bloqueantes", "pendiente-retest" }, hechos.Select(h => h.Codigo));
        // Cerrado o cancelado no está "vencido"; las fuentes sin datos no generan hechos.
        Assert.Empty(DashboardMetricas.Atencion(EstadoProyecto.Cerrado, new DateTime(2026, 10, 10), hoy, null, null));
        Assert.Empty(DashboardMetricas.Atencion(EstadoProyecto.Cancelado, new DateTime(2026, 10, 10), hoy, null, null));
        Assert.Empty(DashboardMetricas.Atencion(EstadoProyecto.Planificado, hoy, hoy, null, null)); // vence hoy: aún no vencido
    }
}
