using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Pages.Admin.Proyectos;
using PortafolioQA.Services;

var connection = args.FirstOrDefault() ?? "Server=.\\SQLEXPRESS;Database=PortafolioQA_Verificacion11;Trusted_Connection=True;TrustServerCertificate=True";
var cs = new SqlConnectionStringBuilder(connection);
if (!cs.InitialCatalog.StartsWith("PortafolioQA_Verificacion", StringComparison.Ordinal))
    throw new InvalidOperationException("La verificación solo admite una BD PortafolioQA_Verificacion*. Restaura una copia previa a Etapa11.");
await using var db = new PortafolioQAContext(new DbContextOptionsBuilder<PortafolioQAContext>().UseSqlServer(connection).Options);
if ((await db.Database.GetAppliedMigrationsAsync()).Any(m => m.EndsWith("Etapa11ActividadesPredeterminadas")))
    throw new InvalidOperationException("Esta copia ya fue migrada; usa una copia nueva del respaldo previo.");

// Captura datos existentes sin usar las columnas nuevas antes de migrar.
var antes = new Dictionary<int, (int Estado, decimal Avance, int Cantidad)>();
await db.Database.OpenConnectionAsync();
await using (var cmd = db.Database.GetDbConnection().CreateCommand())
{
    cmd.CommandText = "SELECT p.Id,p.Estado,p.Avance,COUNT(a.Id) FROM Proyectos p LEFT JOIN Actividades a ON a.ProyectoId=p.Id GROUP BY p.Id,p.Estado,p.Avance";
    await using var reader = await cmd.ExecuteReaderAsync();
    while (await reader.ReadAsync()) antes[reader.GetInt32(0)] = (reader.GetInt32(1), reader.GetDecimal(2), reader.GetInt32(3));
}
Console.WriteLine("Conteo previo (copia real): " + string.Join("; ", antes.Values.GroupBy(p => p.Estado).Select(g => $"estado {g.Key}: {g.Count()}")));
// Cobertura de estados ausentes en la copia: solo se insertan fixtures en la BD de prueba.
await db.Database.ExecuteSqlRawAsync("""
    INSERT INTO Proyectos (Nombre,InstitucionId,Tipo,Estado,FechaInicio,FechaFin,Avance,Activo)
    SELECT 'Verificacion legado '+CONVERT(varchar,v.Estado),MIN(i.Id),0,v.Estado,'2026-10-01','2026-10-10',v.Avance,1
    FROM Instituciones i CROSS JOIN (VALUES (0,0),(2,47),(3,100),(4,35)) v(Estado,Avance) GROUP BY v.Estado,v.Avance;
    """);
await db.Database.MigrateAsync();
foreach (var proyecto in await db.Proyectos.ToListAsync()) await CierreProyecto.RecalcularAvanceAsync(db, proyecto, "Verificación backfill");
await CierreProyecto.ReparaCoherenciaAsync(db);

void Check(bool ok, string detalle) { if (!ok) throw new Exception(detalle); }
void Pass(int n, string texto) => Console.WriteLine($"PASS {n}: {texto}");
T Setup<T>(T page) where T : PageModel
{
    var http = new DefaultHttpContext();
    http.User = new ClaimsPrincipal(new ClaimsIdentity(new[] { new Claim(ClaimTypes.Name, "Verificación"), new Claim(ClaimTypes.Role, Roles.Administrador) }, "test"));
    page.PageContext = new PageContext { HttpContext = http };
    page.TempData = new TempDataDictionary(http, new MemoryTempData());
    return page;
}
var institucion = await db.Instituciones.FirstAsync(i => i.Activo);
var create = Setup(new CreateModel(db, new CapacidadService(db, new ConfigurationBuilder().Build())));
create.Proyecto = new Proyecto { Nombre = "Verificacion nuevo", InstitucionId = institucion.Id, FechaInicio = new DateTime(2026,10,1), FechaFin = new DateTime(2026,10,10), Estado = EstadoProyecto.Cerrado, Avance = 99 };
await create.OnPostAsync();
var p = create.Proyecto;
var acts = await db.Actividades.Where(a => a.ProyectoId == p.Id).OrderBy(a => a.Orden).ToListAsync();
Check(acts.Count==3 && acts.All(a => a.EsPredeterminada && a.RecursoQAId==null) && acts.Select(a=>a.Peso).SequenceEqual(new[]{30m,50m,20m}) && p.Estado==EstadoProyecto.Planificado && p.Avance==0, "Alta incorrecta");
Check((acts[0].FechaFin-acts[0].FechaInicio).Days==2 && (acts[1].FechaFin-acts[1].FechaInicio).Days==4 && (acts[2].FechaFin-acts[2].FechaInicio).Days==1, "Proporciones de fechas incorrectas");
Check(acts[0].FechaInicio==p.FechaInicio && acts[2].FechaFin==p.FechaFin && acts[1].FechaInicio==acts[0].FechaFin.AddDays(1) && acts[2].FechaInicio==acts[1].FechaFin.AddDays(1), "Fechas no secuenciales");
for(var dias=1;dias<=100;dias++)
{
    var corto = new Proyecto { FechaInicio=p.FechaInicio, FechaFin=p.FechaInicio.AddDays(dias-1) };
    Check(ActividadesPredeterminadas.Crear(corto).All(a=>a.FechaInicio>=corto.FechaInicio && a.FechaFin<=corto.FechaFin && a.FechaInicio<=a.FechaFin), "Fechas fuera del rango");
}
Pass(1,"Alta con 3 actividades 30/50/20; fechas inclusivas y rangos de 1–100 días; se ignora avance/estado recibido");
var edit = Setup(new EditModel(db));
await edit.OnPostAvanceActividadAsync(p.Id,acts[0].Id,100);
Check(p.Avance==30 && p.Estado==EstadoProyecto.EnCertificacion,"Diseño no da 30 %");
await edit.OnPostAvanceActividadAsync(p.Id,acts[1].Id,50);
Check(p.Avance==55 && acts[1].Estado==EstadoActividad.EnCurso,"Ejecución no da 55 %");
await edit.OnPostCancelarActividadAsync(p.Id,acts[2].Id);
Check(p.Avance==68.75m && acts[2].Estado==EstadoActividad.Cancelada,"Cancelada no se excluye del denominador");
await edit.OnPostCancelarActividadAsync(p.Id,acts[2].Id);
Check(p.Avance==55 && acts[2].Estado==EstadoActividad.Pendiente,"Reactivar actividad no deriva estado");
Pass(2,"Avance ponderado 30 % y 55 %; canceladas excluidas; estado derivado");
foreach(var a in acts) { a.Avance=0; a.Estado=EstadoActividad.Pendiente; }
acts[0].Estado=EstadoActividad.EnCurso;
await db.SaveChangesAsync(); await CierreProyecto.RecalcularAvanceAsync(db,p);
Check(p.Avance==0 && p.Estado==EstadoProyecto.EnCertificacion,"En curso al 0 % no mantiene certificación");
Pass(3,"Actividad heredada En curso al 0 % implica En certificación");
var recurso = await db.RecursosQA.FirstAsync();
var asig = new Asignacion { ProyectoId=p.Id,RecursoQAId=recurso.Id,Rol="QA",Porcentaje=10,FechaInicio=DateTime.Today.AddDays(-5),FechaFin=DateTime.Today.AddDays(10) };
db.Asignaciones.Add(asig); await db.SaveChangesAsync();
foreach(var a in acts) await edit.OnPostAvanceActividadAsync(p.Id,a.Id,100);
Check(p.Estado==EstadoProyecto.Cerrado && p.Avance==100 && asig.FechaFin==DateTime.Today.AddDays(-1) && asig.FechaFinOriginal!=null,"Cierre no recorta asignación");
var cierre=await db.Cierres.Include(c=>c.Actividades).SingleAsync(c=>c.ProyectoId==p.Id && c.Vigente);
Check(cierre.Actividades.Count==3 && cierre.Actividades.All(a=>a.Avance==100),"Foto incorrecta");
Pass(4,"100 % cierra, crea foto y recorta asignaciones");
await edit.OnPostAgregarActividadAsync(p.Id,"Extra",p.FechaInicio,p.FechaFin,0,null);
Check(p.Estado==EstadoProyecto.EnCertificacion && p.Avance==90.91m && !cierre.Vigente && asig.FechaFinOriginal==null && asig.FechaFin==DateTime.Today.AddDays(10),"Reapertura incorrecta");
Pass(5,"Extra de peso 10 reabre y restaura asignación/cierre histórico");
await edit.OnPostEliminarActividadAsync(p.Id,acts[0].Id);
Check(await db.Actividades.AnyAsync(a=>a.Id==acts[0].Id),"Borró predeterminada");
await edit.OnPostEditarActividadAsync(p.Id,acts[0].Id,"Otro nombre",acts[0].FechaInicio,acts[0].FechaFin);
Check(acts[0].Nombre==ActividadesPredeterminadas.Definiciones[0].Nombre,"Renombró predeterminada");
acts[0].Nombre="Cambio directo";
try { await db.SaveChangesAsync(); throw new Exception("Falta protección del modelo"); }
catch(InvalidOperationException) { acts[0].Nombre=ActividadesPredeterminadas.Definiciones[0].Nombre; db.Entry(acts[0]).State=EntityState.Unchanged; }
db.Actividades.Remove(acts[0]);
try { await db.SaveChangesAsync(); throw new Exception("Falta protección de borrado"); }
catch(InvalidOperationException) { db.Entry(acts[0]).State=EntityState.Unchanged; }
await edit.OnPostEditarActividadAsync(p.Id,acts[0].Id,acts[0].Nombre,acts[0].FechaInicio,acts[0].FechaFin);
Check(acts[0].Peso==30,"La edición no debe alterar el peso interno");
var inicioAntes=acts[0].FechaInicio;
edit.ModelState.AddModelError("inicio","Fecha inválida");
await edit.OnPostEditarActividadAsync(p.Id,acts[0].Id,acts[0].Nombre,acts[0].FechaInicio.AddDays(1),acts[0].FechaFin);
Check(acts[0].FechaInicio==inicioAntes,"Entrada inválida altera las fechas");
edit.ModelState.Clear();
Pass(6,"Borrado/renombrado rechazados en handlers y persistencia; peso interno intacto y entrada inválida rechazada");
var avance=p.Avance;
await edit.OnPostCancelarProyectoAsync(p.Id);
Check(p.Estado==EstadoProyecto.Cancelado && p.Avance==avance,"Cancelación cambia avance");
await edit.OnPostCancelarProyectoAsync(p.Id);
Check(p.Estado==EstadoProyecto.EnCertificacion && p.Avance==avance,"Reactivación cambia avance");
Pass(7,"Cancelar/reactivar conserva avance");
var csv=Encoding.UTF8.GetBytes($"Nombre;Institucion;FechaInicio;FechaFin\nVerificacion CSV A;{institucion.Nombre};2026-10-01;2026-10-10\nVerificacion CSV B;{institucion.Nombre};2026-10-01;2026-10-02\n");
using var stream=new MemoryStream(csv);
var import=Setup(new ImportarModel(db));
import.Archivo=new FormFile(stream,0,csv.Length,"Archivo","verificacion.csv");
await import.OnPostAsync();
var importados=await db.Proyectos.Include(x=>x.Actividades).Where(x=>x.Nombre.StartsWith("Verificacion CSV")).ToListAsync();
Check(import.Resultados.Count(r=>r.Estado=="Creado")==2 && importados.Count==2 && importados.All(x=>x.Actividades.Count==3 && x.Avance==0 && x.Estado==EstadoProyecto.Planificado),"CSV no crea predeterminadas");
Pass(8,"Importación CSV real crea tres actividades en cada proyecto");
foreach(var original in antes)
{
    var actual=await db.Proyectos.Include(x=>x.Actividades).SingleAsync(x=>x.Id==original.Key);
    Check(actual.Actividades.Count(a=>a.EsPredeterminada)==3,"Backfill incompleto");
    if(original.Value.Cantidad==0 && original.Value.Estado==3)
        Check(actual.Estado==EstadoProyecto.Cerrado && actual.Avance==100 && actual.Actividades.All(a=>a.Avance==100),"Cerrado sin actividades no se preservó");
    if(original.Value.Cantidad>0)
        Check(actual.Actividades.Count>=original.Value.Cantidad,"Se perdieron actividades previas");
}
var fixtures=await db.Proyectos.Include(x=>x.Actividades).Where(x=>x.Nombre.StartsWith("Verificacion legado")).ToListAsync();
var legado=fixtures.Single(x=>x.Nombre.EndsWith("2"));
Check(legado.Avance==47 && legado.AvanceLegado==47 && legado.AvancePendienteConfirmar && legado.Estado==EstadoProyecto.EnCertificacion,"Avance heredado no preservado");
await edit.OnPostAvanceActividadAsync(legado.Id,legado.Actividades[0].Id,100);
Check(legado.Avance==47 && legado.AvancePendienteConfirmar,"Una edición confirma prematuramente");
await edit.OnPostConfirmarAvanceAsync(legado.Id,new List<ConfirmacionActividad> { new() { Id=legado.Actividades[0].Id,Avance=100 } });
Check(legado.AvancePendienteConfirmar,"Confirmación parcial aceptada");
await edit.OnPostConfirmarAvanceAsync(legado.Id,legado.Actividades.Select(a=>new ConfirmacionActividad { Id=a.Id,Avance=a.Orden==1?100:0 }).ToList());
Check(!legado.AvancePendienteConfirmar && legado.Avance==30 && legado.AvanceLegado==47,"Confirmación PM incorrecta");
Check(fixtures.Single(x=>x.Nombre.EndsWith("0")).Estado==EstadoProyecto.Planificado,"Backfill Planificado incorrecto");
var cancelado=fixtures.Single(x=>x.Nombre.EndsWith("4"));
Check(cancelado.Estado==EstadoProyecto.Cancelado && cancelado.Avance==35 && cancelado.Actividades.All(a=>a.Estado==EstadoActividad.Cancelada),"Backfill Cancelado incorrecto");
Console.WriteLine("Conteo posterior (solo proyectos originales): " + string.Join("; ",(await db.Proyectos.Where(x=>antes.Keys.Contains(x.Id)).ToListAsync()).GroupBy(x=>x.Estado).Select(g=>$"{g.Key}: {g.Count()}")));
Pass(9,"Backfill de copia real; fixtures de los 4 estados; legado preservado hasta confirmación completa");
Console.WriteLine("9/9 escenarios aprobados. Base original intacta.");

sealed class MemoryTempData : ITempDataProvider
{
    public IDictionary<string, object> LoadTempData(HttpContext context) => new Dictionary<string,object>();
    public void SaveTempData(HttpContext context, IDictionary<string,object> values) { }
}
