using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Tests;

/// <summary>Base temporal propia en .\SQLEXPRESS (se crea y se elimina por prueba); Squash y Bugzilla son simulados: no se toca ninguna base real.</summary>
public class DashboardServicioTests : IAsyncLifetime
{
    private readonly string _cadena = $"Server=.\\SQLEXPRESS;Database=PortafolioQA_test_{Guid.NewGuid():N};Trusted_Connection=True;TrustServerCertificate=True;MultipleActiveResultSets=true";
    private readonly DefectosFalsos _bz = new();
    private readonly DateTime _hoy = new(2026, 10, 16, 12, 0, 0);
    private int _instA, _instB, _sinVinculo, _conSquash, _soloBugzilla, _bugCaido, _bugVacio, _campana1, _it1, _it2, _modX, _modY;

    private PortafolioQAContext Ctx() => new(new DbContextOptionsBuilder<PortafolioQAContext>().UseSqlServer(_cadena).Options);

    private DashboardService Servicio(PortafolioQAContext db, bool squashConf = true, bool bugzillaConf = true, DashboardDefectosCache? cache = null) =>
        new(db, _bz, cache ?? new DashboardDefectosCache(new MemoryCache(new MemoryCacheOptions())),
            Options.Create(new SquashOptions { BaseUrl = squashConf ? "http://localhost:1" : null, Token = squashConf ? "t" : null }),
            new BugzillaServicioFija(new BugzillaOptions { BaseUrl = bugzillaConf ? "http://localhost:2" : null, ApiKey = bugzillaConf ? "k" : null }),
            new RelojFijo(_hoy));

    private static SquashCaso K(int item, int tc, string estado, DateTime? ultima = null) =>
        new() { SquashItemId = item, SquashTestCaseId = tc, Nombre = "c" + item, EstadoSquash = estado, Porcentaje = SquashEstados.Porcentaje(estado), UltimaEjecucion = ultima, Vigente = true };

    public async Task InitializeAsync()
    {
        using var db = Ctx();
        await db.Database.EnsureCreatedAsync();
        var a = new Institucion { Nombre = "Institución A" }; var b = new Institucion { Nombre = "Institución B" };
        db.Instituciones.AddRange(a, b);
        Proyecto P(string n, Institucion i, int? sq, EstadoProyecto e = EstadoProyecto.EnCertificacion, int diasFin = 30) => new()
        { Nombre = n, Institucion = i, SquashProjectId = sq, Estado = e, Avance = 40, FechaInicio = _hoy.AddDays(-60), FechaFin = _hoy.AddDays(diasFin) };

        var pSin = P("Sin vínculos", a, null, diasFin: -5);
        var pSq = P("Con Squash y Bugzilla", a, 1);
        var pBz = P("Solo Bugzilla manual", b, null, EstadoProyecto.Cerrado, -10);
        var pCaido = P("Bugzilla caído", b, 2);
        var pVacio = P("Bugzilla sin defectos", b, null);
        pSq.SquashSincronizadoEn = _hoy.AddHours(-3);
        pCaido.SquashSincronizadoEn = _hoy.AddHours(-1);   // sincronizado pero sin casos → "Sin datos", no cero
        pSq.VinculoBugzilla = new VinculoBugzilla { SquashProjectId = 1, ProductoNombre = "ProdSq", Estado = EstadoBugzilla.Configurado };
        pBz.Integraciones.Add(new AsociacionExterna { Herramienta = HerramientaIntegracion.Bugzilla, IdentificadorExterno = "ProdSq" }); // mismo producto que otro proyecto
        pCaido.Integraciones.Add(new AsociacionExterna { Herramienta = HerramientaIntegracion.Bugzilla, IdentificadorExterno = "ProdCaido" });
        pVacio.Integraciones.Add(new AsociacionExterna { Herramienta = HerramientaIntegracion.Bugzilla, IdentificadorExterno = "ProdVacio" });
        var r = new RecursoQA { Nombres = "Ana", Apellidos = "Quispe", Correo = "ana@x.pe" };
        db.RecursosQA.Add(r);
        pSq.Asignaciones.Add(new Asignacion { RecursoQA = r, Rol = "QA", Porcentaje = 50, FechaInicio = _hoy.AddDays(-10), FechaFin = _hoy.AddDays(10) });

        // Squash: 1 campaña y 2 iteraciones. El caso de prueba tc=1 está en ambas; módulos X (it1) e Y (it2).
        var camp = new SquashCampana { Proyecto = pSq, SquashId = 1, Nombre = "Campaña 1", Vigente = true };
        var it1 = new SquashIteracion { SquashId = 1, Nombre = "Iteración 1", Vigente = true };
        var it2 = new SquashIteracion { SquashId = 2, Nombre = "Iteración 2", Vigente = true };
        camp.Iteraciones.AddRange(new[] { it1, it2 });
        it1.Casos.AddRange(new[] { K(1, 1, "FAILURE", new DateTime(2026, 9, 1)), K(2, 2, "SUCCESS"), K(3, 3, "BLOCKED") });
        it2.Casos.AddRange(new[] { K(4, 1, "SUCCESS", new DateTime(2026, 10, 1)), K(5, 4, "READY"), K(6, 5, "UNTESTABLE") });
        var mx = new SquashModulo { SquashId = 10, Nombre = "Módulo X", Vigente = true, Casos = { it1.Casos[0], it1.Casos[1] } };
        var my = new SquashModulo { SquashId = 11, Nombre = "Módulo Y", Vigente = true, Casos = { it2.Casos[0] } };
        it1.Modulos.Add(mx); it2.Modulos.Add(my);
        // Dato no vigente: no debe contar.
        it2.Casos.Add(new SquashCaso { SquashItemId = 7, SquashTestCaseId = 99, Nombre = "viejo", EstadoSquash = "FAILURE", Porcentaje = 50, Vigente = false });
        db.SquashCampanas.Add(camp);
        db.Proyectos.AddRange(pSin, pSq, pBz, pCaido, pVacio);
        await db.SaveChangesAsync();
        (_instA, _instB) = (a.Id, b.Id);
        (_sinVinculo, _conSquash, _soloBugzilla, _bugCaido, _bugVacio) = (pSin.Id, pSq.Id, pBz.Id, pCaido.Id, pVacio.Id);
        (_campana1, _it1, _it2, _modX, _modY) = (camp.Id, it1.Id, it2.Id, mx.Id, my.Id);

        _bz.PorProducto["ProdSq"] = new ResultadoDefectos
        {
            Disponible = true, UltimaActualizacion = DateTime.UtcNow,
            Defectos =
            {
                DefectosFalsos.Bug(1, "CONFIRMED", "blocker"), DefectosFalsos.Bug(2, "IN_PROGRESS", "normal"), DefectosFalsos.Bug(3, "UNCONFIRMED", "enhancement"),
                DefectosFalsos.Bug(4, "RESOLVED", "major", "FIXED"), DefectosFalsos.Bug(5, "VERIFIED", "blocker", "FIXED"), DefectosFalsos.Bug(6, "CONFIRMED", "")
            }
        };
        _bz.PorProducto["ProdVacio"] = new ResultadoDefectos { Disponible = true, UltimaActualizacion = DateTime.UtcNow };
        // "ProdCaido" no está en el diccionario: el falso responde con error.
    }

    public async Task DisposeAsync() { using var db = Ctx(); await db.Database.EnsureDeletedAsync(); }

    private static readonly FiltrosDashboard SinFiltros = new(null, null, null, null, null);

    [Fact]
    public async Task Distingue_cero_real_sin_datos_sin_vinculo_y_error_y_conserva_el_resto()
    {
        using var db = Ctx();
        var r = await Servicio(db).CargarAsync(SinFiltros, false, default);
        ProyectoDashboard F(int id) => r.Filas.Single(f => f.Id == id);

        Assert.Equal(EstadoFuente.SinVinculo, F(_sinVinculo).Squash.Estado);
        Assert.Equal(EstadoFuente.SinVinculo, F(_sinVinculo).Bugzilla.Estado);
        Assert.Null(F(_sinVinculo).Certificacion);
        Assert.Null(F(_sinVinculo).Defectos);

        Assert.Equal(EstadoFuente.Ok, F(_conSquash).Squash.Estado);
        Assert.Equal(EstadoFuente.Ok, F(_conSquash).Bugzilla.Estado);

        // Bugzilla caído: esa fuente queda en Error y sin cifras; el resto del proyecto sigue disponible.
        var caido = F(_bugCaido);
        Assert.Equal(EstadoFuente.Error, caido.Bugzilla.Estado);
        Assert.Null(caido.Defectos);
        Assert.Contains(caido.Incompleto, x => x.StartsWith("Bugzilla"));
        Assert.Equal(EstadoFuente.SinDatos, caido.Squash.Estado);   // vinculado y sincronizado, sin casos
        Assert.Equal(40, caido.Avance);

        // Cero real: Bugzilla respondió y no hay defectos.
        var vacio = F(_bugVacio);
        Assert.Equal(EstadoFuente.Ok, vacio.Bugzilla.Estado);
        Assert.Equal(0, vacio.Defectos!.Abiertos);
        Assert.Equal(0, vacio.Defectos.Total);
        Assert.Equal(EstadoFuente.Error, r.FuenteBugzilla.Estado);  // al menos un proyecto con error
    }

    [Fact]
    public async Task Calcula_certificacion_y_defectos_del_proyecto_sin_duplicar()
    {
        using var db = Ctx();
        var r = await Servicio(db).CargarAsync(SinFiltros, false, default);
        var f = r.Filas.Single(x => x.Id == _conSquash);
        var c = f.Certificacion!;
        // Únicos vigentes que cuentan: tc1 (gana SUCCESS, es más reciente), tc2 SUCCESS, tc3 BLOCKED, tc4 READY → 4; tc5 UNTESTABLE excluido; el no vigente no cuenta.
        Assert.Equal(4, c.Total);
        Assert.Equal(1, c.Excluidos);
        Assert.Equal(2, c.Ejecutados);
        Assert.Equal(2, c.Aprobados);
        Assert.Equal(0, c.Fallidos);
        Assert.Equal(1, c.Bloqueados);
        Assert.Equal(50m, c.PorcentajeEjecutado);
        Assert.Equal(100m, c.TasaAprobacion);
        // Abiertos = 1, 2, 3 y 6 (el 4 es RESOLVED y el 5 VERIFIED). Retest = bug 4.
        var d = f.Defectos!;
        Assert.Equal(6, d.Total);
        Assert.Equal(4, d.Abiertos);
        Assert.Equal(1, d.Bloqueantes);
        Assert.Equal(1, d.NoBloqueantes);
        Assert.Equal(1, d.Mejoras);
        Assert.Equal(1, d.SinClasificar);
        Assert.Equal(1, d.PendienteRetest);
        Assert.Equal(new[] { "1" }, f.DefectosBloqueantes.Select(x => x.Id));
        Assert.Equal(new[] { "4" }, f.DefectosRetest.Select(x => x.Id));
        Assert.Equal(new[] { "pruebas-bloqueadas", "defectos-bloqueantes", "pendiente-retest" }, f.Atencion.Select(h => h.Codigo));
        Assert.Equal(_hoy.AddHours(-3), f.Squash.ActualizadoEn);
        Assert.Single(f.Equipo);
        Assert.True(f.Equipo[0].VigenteHoy);
    }

    [Fact]
    public async Task Producto_inexistente_en_bugzilla_es_error_y_no_cero_real()
    {
        using (var alta = Ctx())
        {
            var inst = await alta.Instituciones.FirstAsync();
            var p = new Proyecto { Nombre = "Producto mal escrito", Institucion = inst, FechaInicio = _hoy, FechaFin = _hoy.AddDays(5) };
            p.Integraciones.Add(new AsociacionExterna { Herramienta = HerramientaIntegracion.Bugzilla, IdentificadorExterno = "ProdFantasma" });
            alta.Proyectos.Add(p);
            await alta.SaveChangesAsync();
        }
        _bz.PorProducto["ProdFantasma"] = new ResultadoDefectos { Disponible = true, UltimaActualizacion = DateTime.UtcNow }; // Bugzilla real responde lista vacía
        _bz.Inexistentes.Add("ProdFantasma");
        using var db = Ctx();
        var r = await Servicio(db).CargarAsync(SinFiltros, false, default);
        var f = r.Filas.Single(x => x.Nombre == "Producto mal escrito");
        Assert.Equal(EstadoFuente.Error, f.Bugzilla.Estado);
        Assert.Null(f.Defectos);
        Assert.Equal(EstadoFuente.Ok, r.Filas.Single(x => x.Id == _bugVacio).Bugzilla.Estado); // el producto que sí existe y no tiene defectos sigue siendo cero real
    }

    [Fact]
    public async Task Vencido_solo_si_el_proyecto_sigue_abierto()
    {
        using var db = Ctx();
        var r = await Servicio(db).CargarAsync(SinFiltros, false, default);
        Assert.Equal(5, r.Filas.Single(f => f.Id == _sinVinculo).DiasVencido);
        Assert.Contains(r.Filas.Single(f => f.Id == _sinVinculo).Atencion, h => h.Codigo == "vencido");
        Assert.Null(r.Filas.Single(f => f.Id == _soloBugzilla).DiasVencido);   // cerrado
    }

    [Fact]
    public async Task Filtra_por_institucion_y_el_proyecto_manda_sobre_la_institucion()
    {
        using var db = Ctx();
        var s = Servicio(db);
        var a = await s.CargarAsync(new(_instA, null, null, null, null), false, default);
        Assert.All(a.Filas, f => Assert.Equal(_instA, f.InstitucionId));
        Assert.Equal(2, a.Filas.Count);
        var b = await s.CargarAsync(new(_instA, _bugVacio, null, null, null), false, default);   // el proyecto es de B
        Assert.Equal(_instB, b.Filtros.InstitucionId);
        Assert.All(b.Filas, f => Assert.Equal(_instB, f.InstitucionId));
        Assert.Equal(_bugVacio, b.Detalle!.Proyecto.Id);
    }

    [Fact]
    public async Task Resumen_e_instituciones_usan_el_alcance_correcto_y_son_coherentes()
    {
        using var db = Ctx();
        var s = Servicio(db);
        var todo = await s.CargarAsync(SinFiltros, false, default);
        // Alcance completo: 5 proyectos; 4 en certificación; 1 vencido (abierto con fecha pasada; el cerrado no cuenta).
        Assert.Equal(5, todo.Resumen.Total);
        Assert.Equal(todo.Filas.Count, todo.Resumen.Total);
        Assert.Equal(4, todo.Resumen.EnCertificacion);
        Assert.Equal(1, todo.Resumen.Vencidos);
        // Bugzilla respondió en 3 proyectos (el caído y los sin vínculo no se evalúan); los dos que comparten ProdSq tienen 1 bloqueante cada uno.
        Assert.Equal(3, todo.Resumen.ConBugzillaConsultado);
        Assert.Equal(2, todo.Resumen.ConBloqueantes);
        Assert.Equal(todo.Resumen.Total, todo.Resumen.PorEstado.Values.Sum());
        Assert.Equal(todo.Resumen.Total, todo.Resumen.PorTipo.Values.Sum());

        // Las tarjetas cuentan siempre el portafolio completo y suman el total; el filtro solo cambia el resumen.
        var a = await s.CargarAsync(new(_instA, null, null, null, null), false, default);
        Assert.Equal(2, a.Resumen.Total);
        Assert.Equal(todo.TarjetasInstituciones.Select(t => (t.Id, t.Proyectos)), a.TarjetasInstituciones.Select(t => (t.Id, t.Proyectos)));
        Assert.Equal(todo.Resumen.Total, todo.TarjetasInstituciones.Sum(t => t.Proyectos));
        Assert.Equal(2, todo.TarjetasInstituciones.Single(t => t.Id == _instA).Proyectos);
        var tb = todo.TarjetasInstituciones.Single(t => t.Id == _instB);
        Assert.Equal((3, 1, 2), (tb.Proyectos, tb.Estado(EstadoProyecto.Cerrado), tb.Estado(EstadoProyecto.EnCertificacion)));
    }

    [Fact]
    public async Task Estados_de_pruebas_suman_los_casos_totales_y_defectos_abiertos_no_incluyen_retest()
    {
        using var db = Ctx();
        var r = await Servicio(db).CargarAsync(SinFiltros, false, default);
        var f = r.Filas.Single(x => x.Id == _conSquash);
        var c = f.Certificacion!;
        Assert.Equal(c.Total, c.Aprobados + c.Fallidos + c.ConAdvertencia + c.Bloqueados + c.EnCurso + c.SinEjecutar);
        Assert.Equal((4, 2, 1, 1), (c.Total, c.Aprobados, c.Bloqueados, c.SinEjecutar));
        var d = f.Defectos!;
        Assert.Equal(4, d.Abiertos);                 // 1, 2, 3 y 6; el RESOLVED (4) y el VERIFIED (5) no son abiertos
        Assert.Equal(1, d.PendienteRetest);          // el 4 (RESOLVED + FIXED) se cuenta aparte, no dentro de los abiertos
        Assert.Equal(d.Abiertos - d.Bloqueantes, d.NoBloqueantes + d.Mejoras + d.SinClasificar);
    }

    [Fact]
    public async Task Filtros_de_campana_iteracion_modulo_afectan_solo_a_certificacion_y_usan_la_relacion_real()
    {
        using var db = Ctx();
        var s = Servicio(db);
        var todo = await s.CargarAsync(new(null, _conSquash, null, null, null), false, default);
        var defectosTodo = todo.Detalle!.Proyecto.Defectos;

        var it1 = await s.CargarAsync(new(null, _conSquash, null, _it1, null), false, default);
        Assert.Equal(_campana1, it1.Detalle!.CampanaId);                     // la campaña se deduce de la jerarquía
        var c1 = it1.Detalle.Proyecto.Certificacion!;
        Assert.Equal(3, c1.Total);                                            // tc1 (FAILURE en esta iteración), tc2, tc3
        Assert.Equal(1, c1.Fallidos);
        Assert.Equal(1, c1.Bloqueados);
        Assert.Equal(defectosTodo, it1.Detalle.Proyecto.Defectos);            // Bugzilla no relaciona campaña/iteración: los defectos no cambian

        var modX = await s.CargarAsync(new(null, _conSquash, null, null, _modX), false, default);
        Assert.Equal(_it1, modX.Detalle!.IteracionId);
        Assert.Equal(2, modX.Detalle.Proyecto.Certificacion!.Total);          // solo los ítems que la suite X contiene
        Assert.Single(modX.Detalle.Campanas.Single().Iteraciones.Single().Modulos);

        // Iteración y módulo incoherentes: manda el más específico (módulo) y se avisa.
        var incoherente = await s.CargarAsync(new(null, _conSquash, null, _it2, _modX), false, default);
        Assert.Equal(_it1, incoherente.Detalle!.IteracionId);
        Assert.Contains(incoherente.Avisos, a => a.Contains("coherentes"));

        // Ids ajenos o inexistentes se ignoran con aviso: nunca se atribuye por nombre.
        var ajeno = await s.CargarAsync(new(null, _conSquash, 99999, 99999, 99999), false, default);
        Assert.Null(ajeno.Detalle!.CampanaId);
        Assert.Equal(4, ajeno.Detalle.Proyecto.Certificacion!.Total);
        Assert.True(ajeno.Avisos.Count >= 1);

        // Sin proyecto elegido, los filtros finos no aplican.
        var sinProyecto = await s.CargarAsync(new(null, null, _campana1, null, null), false, default);
        Assert.Null(sinProyecto.Detalle);
        Assert.Contains(sinProyecto.Avisos, a => a.Contains("requieren elegir un proyecto"));
    }

    [Fact]
    public async Task Desglose_calcula_cada_nivel_por_separado_e_informa_los_casos_sin_modulo()
    {
        using var db = Ctx();
        var r = await Servicio(db).CargarAsync(new(null, _conSquash, null, null, null), false, default);
        var camp = r.Detalle!.Campanas.Single();
        Assert.Equal(4, camp.Certificacion.Total);                            // campaña: tc1 una sola vez
        var i1 = camp.Iteraciones.Single(i => i.Id == _it1);
        var i2 = camp.Iteraciones.Single(i => i.Id == _it2);
        Assert.Equal(3, i1.Certificacion.Total);
        Assert.Equal(2, i2.Certificacion.Total);                              // tc1 y tc4; tc5 está excluido
        Assert.Equal(1, i1.CasosSinModulo);                                   // tc3
        Assert.Equal(1, i2.CasosSinModulo);                                   // tc4
        Assert.Equal(2, i1.Modulos.Single().Certificacion.Total);
        Assert.Equal(1, i2.Modulos.Single().Certificacion.Total);
    }

    [Fact]
    public async Task Bugzilla_se_consulta_una_vez_por_producto_con_cache_y_cupo()
    {
        using var db = Ctx();
        var cache = new DashboardDefectosCache(new MemoryCache(new MemoryCacheOptions()));
        var s = Servicio(db, cache: cache);
        await s.CargarAsync(SinFiltros, false, default);
        await s.CargarAsync(SinFiltros, false, default);
        Assert.Equal(1, _bz.Llamadas["ProdSq"]);          // dos proyectos usan ProdSq y la página se pidió dos veces: una sola consulta
        Assert.Equal(1, _bz.Llamadas["ProdCaido"]);
        await s.CargarAsync(SinFiltros, true, default);   // forzar ignora la caché
        Assert.Equal(2, _bz.Llamadas["ProdSq"]);
        Assert.InRange(_bz.MaximoSimultaneo, 1, 4);
    }

    [Fact]
    public async Task Integraciones_no_configuradas_se_distinguen_y_no_consultan()
    {
        using var db = Ctx();
        var r = await Servicio(db, squashConf: false, bugzillaConf: false).CargarAsync(SinFiltros, false, default);
        var f = r.Filas.Single(x => x.Id == _conSquash);
        Assert.Equal(EstadoFuente.NoConfigurada, f.Bugzilla.Estado);
        Assert.Null(f.Defectos);
        Assert.Equal(EstadoFuente.NoConfigurada, f.Squash.Estado);            // hay datos sincronizados: se muestran y se avisa
        Assert.NotNull(f.Certificacion);
        Assert.Contains(f.Incompleto, x => x.StartsWith("Squash TM"));
        Assert.Empty(_bz.Llamadas);
    }

    [Fact]
    public async Task No_escribe_en_la_base()
    {
        using var db = Ctx();
        static string Huella(PortafolioQAContext c) =>
            string.Join("|", c.Proyectos.AsNoTracking().OrderBy(p => p.Id).Select(p => p.Id + ":" + p.Avance + ":" + p.SquashSincronizadoEn + ":" + p.Estado).ToList())
            + "#" + c.SquashCasos.Count() + "#" + c.SquashCampanas.Count() + "#" + c.VinculosBugzilla.Count();
        var antes = Huella(db);
        await Servicio(db).CargarAsync(new(null, _conSquash, null, _it1, null), false, default);
        Assert.False(db.ChangeTracker.HasChanges());
        Assert.Empty(db.ChangeTracker.Entries());
        using var db2 = Ctx();
        Assert.Equal(antes, Huella(db2));
    }
}
