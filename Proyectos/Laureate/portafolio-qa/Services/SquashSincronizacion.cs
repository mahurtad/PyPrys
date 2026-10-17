using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Services;

/// <summary>
/// Copia a la base la última lectura de Squash TM del proyecto vinculado (campañas, iteraciones y casos con su porcentaje)
/// y recalcula el avance del proyecto con el mismo flujo de siempre (CierreProyecto.RecalcularAvanceAsync).
/// </summary>
public static class SquashSincronizacion
{
    public static async Task<string> SincronizarAsync(PortafolioQAContext context, ISquashApi squash, Proyecto proyecto, string usuario, CancellationToken ct = default)
    {
        if (proyecto.SquashProjectId is not int squashProjectId) return "El proyecto no está vinculado a un proyecto de Squash TM.";
        // Un proyecto cerrado es historial: reabrirlo por un cambio en Squash movería la capacidad de los recursos.
        if (proyecto.Estado is EstadoProyecto.Cerrado or EstadoProyecto.Cancelado)
            return "El proyecto está Cerrado o Cancelado: no se sincroniza para no alterar el historial ni la capacidad.";

        var lectura = await squash.ObtenerEjecucionAsync(squashProjectId, ct);
        // Sin lectura no se toca nada: un resultado parcial bajaría el avance sin motivo.
        if (!lectura.Ok) return "No se actualizó el avance de Squash TM. " + lectura.Error;

        var existentes = await context.SquashCampanas
            .Include(c => c.Iteraciones).ThenInclude(i => i.Casos)
            .Include(c => c.Iteraciones).ThenInclude(i => i.Modulos).ThenInclude(m => m.Casos)
            .Where(c => c.ProyectoId == proyecto.Id).ToListAsync(ct);
        foreach (var c in existentes)
        {
            c.Vigente = false;
            foreach (var i in c.Iteraciones) { i.Vigente = false; foreach (var k in i.Casos) k.Vigente = false; foreach (var m in i.Modulos) m.Vigente = false; }
        }

        var recursos = (await context.RecursosQA.Where(r => r.SquashLogin != null).ToListAsync(ct))
            .ToDictionary(r => r.SquashLogin!.Trim().ToLowerInvariant(), r => r.Id);
        int totalIteraciones = 0, totalCasos = 0;

        foreach (var cd in lectura.Datos!)
        {
            var campana = existentes.FirstOrDefault(c => c.SquashId == cd.Id);
            if (campana is null)
            {
                campana = new SquashCampana { ProyectoId = proyecto.Id, SquashId = cd.Id };
                context.SquashCampanas.Add(campana);
                existentes.Add(campana);
            }
            campana.Nombre = Recortar(cd.Nombre, 200);
            campana.Estado = cd.Estado;
            campana.Vigente = true;

            foreach (var idto in cd.Iteraciones)
            {
                var iteracion = campana.Iteraciones.FirstOrDefault(i => i.SquashId == idto.Id);
                if (iteracion is null) { iteracion = new SquashIteracion { SquashId = idto.Id }; campana.Iteraciones.Add(iteracion); }
                iteracion.Nombre = Recortar(idto.Nombre, 200);
                iteracion.Estado = idto.Estado;
                iteracion.Vigente = true;
                totalIteraciones++;

                foreach (var kd in idto.Casos)
                {
                    var caso = iteracion.Casos.FirstOrDefault(k => k.SquashItemId == kd.ItemId);
                    if (caso is null) { caso = new SquashCaso { SquashItemId = kd.ItemId }; iteracion.Casos.Add(caso); }
                    caso.SquashTestCaseId = kd.TestCaseId;
                    caso.Nombre = Recortar(kd.Nombre, 300);
                    caso.EstadoSquash = Recortar(kd.Estado, 30);
                    caso.Porcentaje = SquashEstados.Porcentaje(kd.Estado);
                    caso.AsignadoLogin = kd.AsignadoLogin is null ? null : Recortar(kd.AsignadoLogin, 100);
                    caso.RecursoQAId = kd.AsignadoLogin is not null && recursos.TryGetValue(kd.AsignadoLogin.Trim().ToLowerInvariant(), out var rid) ? rid : null;
                    caso.UltimaEjecucion = kd.UltimaEjecucion;
                    caso.Vigente = true;
                    totalCasos++;
                }
                // Módulos: la pertenencia se reconstruye en cada lectura con lo que informa Squash.
                foreach (var m in iteracion.Modulos) m.Casos.Clear();
                foreach (var md in idto.Modulos ?? new())
                {
                    var modulo = iteracion.Modulos.FirstOrDefault(m => m.SquashId == md.Id);
                    if (modulo is null) { modulo = new SquashModulo { SquashId = md.Id }; iteracion.Modulos.Add(modulo); }
                    modulo.Nombre = Recortar(md.Nombre, 200);
                    modulo.Vigente = true;
                    foreach (var itemId in md.ItemIds.Distinct())
                        if (iteracion.Casos.FirstOrDefault(k => k.SquashItemId == itemId) is { } caso) modulo.Casos.Add(caso);
                }
                var cuentan = iteracion.Casos.Where(k => k.Vigente && k.Porcentaje is not null).ToList();
                iteracion.Porcentaje = cuentan.Count == 0 ? 0 : Math.Round(cuentan.Average(k => k.Porcentaje!.Value), 2);
            }
        }

        // El nombre del proyecto es el de Squash TM. Si no se puede leer la lista, se conserva el guardado.
        var proyectosSquash = await squash.ListarProyectosAsync(ct);
        var nombreSquash = proyectosSquash.Datos?.FirstOrDefault(p => p.Id == squashProjectId)?.Nombre;
        if (!string.IsNullOrWhiteSpace(nombreSquash) && proyecto.Nombre != Recortar(nombreSquash, 200))
        {
            CierreProyecto.Auditar(context, proyecto, null, "Squash TM", usuario, "Nombre", $"«{proyecto.Nombre}» → «{Recortar(nombreSquash, 200)}»");
            proyecto.Nombre = Recortar(nombreSquash, 200);
        }

        var avanceAnterior = proyecto.Avance;
        proyecto.SquashSincronizadoEn = DateTime.Now;
        await context.SaveChangesAsync(ct);
        var aviso = await CierreProyecto.RecalcularAvanceAsync(context, proyecto, usuario);

        var detalle = $"{lectura.Datos!.Count} campaña(s), {totalIteraciones} iteración(es), {totalCasos} caso(s)";
        if (proyecto.Avance != avanceAnterior)
        {
            CierreProyecto.Auditar(context, proyecto, null, "Squash TM", usuario, "Sincronización", $"{avanceAnterior:0.##} % → {proyecto.Avance:0.##} % ({detalle})");
            await context.SaveChangesAsync(ct);
        }
        return $"Squash TM sincronizado: {detalle}. Avance del proyecto: {proyecto.Avance:0.##} %." + aviso;
    }

    /// <summary>Un proyecto de Squash solo puede estar vinculado a un proyecto del portafolio (1 a 1) y debe existir.</summary>
    public static async Task<string?> ValidarVinculoAsync(PortafolioQAContext context, int? squashProjectId, int proyectoId, List<SquashProyectoDto>? proyectosSquash)
    {
        if (squashProjectId is null) return null;
        if (proyectosSquash is not null && proyectosSquash.All(p => p.Id != squashProjectId))
            return "El proyecto seleccionado no existe en Squash TM.";
        if (await context.Proyectos.AnyAsync(p => p.SquashProjectId == squashProjectId && p.Id != proyectoId))
            return "Ese proyecto de Squash TM ya está vinculado a otro proyecto del portafolio.";
        return null;
    }

    private static string Recortar(string texto, int max) => texto.Length <= max ? texto : texto[..max];
}
