using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Conexion;
using PortafolioQA.Services.Integrations;

var builder = WebApplication.CreateBuilder(args);

// La conexión a SQL Server se resuelve en tiempo de ejecución (archivo cifrado del asistente /Setup o configuración del
// servidor). Sin conexión configurada la aplicación arranca en modo "pendiente": solo el asistente está disponible.
builder.Services.AddSingleton<ConexionStore>();
builder.Services.AddSingleton<ConexionProbador>();
builder.Services.AddSingleton<EstadoApp>();
builder.Services.AddDbContext<PortafolioQAContext>((sp, options) =>
    options.UseSqlServer(sp.GetRequiredService<ConexionStore>().Cadena
        ?? "Server=.;Database=PortafolioQA_pendiente;Connect Timeout=1;TrustServerCertificate=True"));

// Autenticación con cuentas locales (ASP.NET Core Identity) y roles Administrador / Editor / Lector.
builder.Services.AddIdentity<UsuarioApp, IdentityRole>(options =>
    {
        options.Password.RequiredLength = 10;
        options.Password.RequireDigit = true;
        options.Password.RequireLowercase = true;
        options.Password.RequireUppercase = true;
        options.Password.RequireNonAlphanumeric = false;
        options.User.RequireUniqueEmail = true;
        options.Lockout.AllowedForNewUsers = true;
        options.Lockout.MaxFailedAccessAttempts = 5;
        options.Lockout.DefaultLockoutTimeSpan = TimeSpan.FromMinutes(15);
    })
    .AddEntityFrameworkStores<PortafolioQAContext>()
    .AddDefaultTokenProviders();

builder.Services.ConfigureApplicationCookie(options =>
{
    options.LoginPath = "/Account/Login";
    options.AccessDeniedPath = "/Account/AccesoDenegado";
    options.Cookie.Name = "PortafolioQA.Auth";
    options.Cookie.HttpOnly = true;
    options.Cookie.SameSite = SameSiteMode.Lax;
    // En desarrollo se usa http://localhost; en cualquier otro entorno la cookie solo viaja por HTTPS.
    options.Cookie.SecurePolicy = builder.Environment.IsDevelopment()
        ? CookieSecurePolicy.SameAsRequest
        : CookieSecurePolicy.Always;
    options.ExpireTimeSpan = TimeSpan.FromHours(8);
    options.SlidingExpiration = true;
});

// Revalida la cuenta cada minuto: un usuario desactivado o con la contraseña cambiada pierde la sesión rápido.
builder.Services.Configure<SecurityStampValidatorOptions>(o => o.ValidationInterval = TimeSpan.FromMinutes(1));

builder.Services.AddAuthorization(options =>
{
    options.AddPolicy(Politicas.PuedeEditar, p => p.RequireRole(Roles.Administrador, Roles.Editor));
    options.AddPolicy(Politicas.SoloAdmin, p => p.RequireRole(Roles.Administrador));
});

// Las claves de protección de datos (cookies, antiforgery) deben persistir fuera del proceso en IIS;
// de lo contrario cada reciclaje del application pool cierra todas las sesiones.
var rutaClaves = builder.Configuration["DataProtection:RutaClaves"];
if (!string.IsNullOrWhiteSpace(rutaClaves))
{
    builder.Services.AddDataProtection()
        .SetApplicationName("PortafolioQA")
        .PersistKeysToFileSystem(new DirectoryInfo(rutaClaves));
}

builder.Services.AddRazorPages(options =>
{
    // Todo requiere sesión, salvo el acceso. Lector: solo consulta (Portafolio y Planificación).
    options.Conventions.AuthorizeFolder("/");
    options.Conventions.AllowAnonymousToPage("/Account/Login");
    options.Conventions.AllowAnonymousToPage("/Account/AccesoDenegado");
    options.Conventions.AllowAnonymousToPage("/Error");
    // El asistente de conexión controla su propio acceso: código de configuración (aplicación pendiente) o Administrador.
    options.Conventions.AllowAnonymousToFolder("/Setup");
    options.Conventions.AuthorizeFolder("/Admin", Politicas.PuedeEditar);
    options.Conventions.AuthorizeFolder("/Admin/Usuarios", Politicas.SoloAdmin);
});

builder.Services.AddScoped<CapacidadService>();

// Integraciones externas: en la etapa 1 no hay conexión real configurada.
// Las etapas 4 y 5 registrarán implementaciones reales sin cambiar quien las consume.
// Squash TM: el token va en user-secrets (Squash:Token). Sin token se mantiene la implementación "no configurada".
builder.Services.Configure<SquashOptions>(builder.Configuration.GetSection("Squash"));
builder.Services.AddHttpClient<SquashTmClient>((sp, http) =>
{
    var o = sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<SquashOptions>>().Value;
    if (!o.Configurado) return;
    http.BaseAddress = new Uri(o.BaseUrl!.TrimEnd('/') + "/api/rest/latest/");
    http.Timeout = TimeSpan.FromSeconds(20);
    http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", o.Token);
    http.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
});
builder.Services.AddScoped<IIntegracionCertificacion>(sp =>
    sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<SquashOptions>>().Value.Configurado
        ? sp.GetRequiredService<SquashTmClient>()
        : new IntegracionCertificacionNoConfigurada());
builder.Services.AddScoped<ISquashApi>(sp =>
    sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<SquashOptions>>().Value.Configurado
        ? sp.GetRequiredService<SquashTmClient>()
        : new SquashNoConfigurado());
builder.Services.Configure<BugzillaOptions>(builder.Configuration.GetSection("Bugzilla"));
builder.Services.AddHttpClient<BugzillaClient>((sp, http) =>
{
    var o = sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<BugzillaOptions>>().Value;
    if (!o.Configurado) return;
    http.BaseAddress = new Uri(o.BaseUrl!.TrimEnd('/') + "/rest/");
    http.Timeout = TimeSpan.FromSeconds(20);
    http.DefaultRequestHeaders.Add("X-BUGZILLA-API-KEY", o.ApiKey);
    http.DefaultRequestHeaders.Accept.ParseAdd("application/json");
});
// Cuenta de servicio con editcomponents (Bugzilla:AdminApiKey): crea el producto y el componente de cada iniciativa por la API REST.
builder.Services.AddHttpClient<BugzillaProductosClient>((sp, http) =>
{
    var o = sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<BugzillaOptions>>().Value;
    if (!o.ProductosConfigurado) return;
    http.BaseAddress = new Uri(o.BaseUrl!.TrimEnd('/') + "/rest/");
    http.Timeout = TimeSpan.FromSeconds(20);
    http.DefaultRequestHeaders.Accept.ParseAdd("application/json");
})
.AddHttpMessageHandler(sp => new BugzillaApiKeyHandler(sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<BugzillaOptions>>().Value.AdminApiKey ?? string.Empty));
builder.Logging.AddFilter("System.Net.Http.HttpClient.BugzillaProductosClient", LogLevel.Warning);
builder.Services.AddScoped<IBugzillaProductos>(sp =>
    sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<BugzillaOptions>>().Value.ProductosConfigurado
        ? sp.GetRequiredService<BugzillaProductosClient>()
        : new BugzillaProductosNoConfigurado());
builder.Services.AddScoped<BugzillaProvision>();
builder.Services.AddScoped<IIntegracionDefectos>(sp =>
    sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<BugzillaOptions>>().Value.Configurado
        ? sp.GetRequiredService<BugzillaClient>()
        : new IntegracionDefectosNoConfigurada());

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();

app.Use(async (context, next) =>
{
    var h = context.Response.Headers;
    h["X-Content-Type-Options"] = "nosniff";
    h["X-Frame-Options"] = "DENY";
    h["Referrer-Policy"] = "same-origin";
    await next();
});

app.UseStaticFiles();

app.UseRouting();

// Mientras la aplicación no esté activada (sin conexión, sin esquema o con migraciones pendientes) todo lleva al asistente.
app.Use(async (context, next) =>
{
    var estado = context.RequestServices.GetRequiredService<EstadoApp>();
    if (estado.Fase != FaseApp.Activa)
    {
        var ruta = context.Request.Path;
        var libre = ruta.StartsWithSegments("/Setup") || ruta.StartsWithSegments("/healthz") || ruta.StartsWithSegments("/Error")
            || (estado.Fase == FaseApp.EsquemaDesactualizado && ruta.StartsWithSegments("/Account"));
        if (!libre)
        {
            context.Response.Redirect("/Setup");
            return;
        }
    }
    await next();
});

app.UseAuthentication();
app.UseAuthorization();

// Verificación de disponibilidad para IIS / monitoreo: anónima y sin detalles internos.
// ok = activa y con base alcanzable; pendiente / esquema-pendiente = falta configurar o migrar (no es una falla).
app.MapGet("/healthz", async (PortafolioQAContext db, EstadoApp estado) =>
{
    if (estado.Fase == FaseApp.Pendiente) return Results.Text("pendiente", statusCode: 503);
    if (estado.Fase == FaseApp.EsquemaDesactualizado) return Results.Text("esquema-pendiente", statusCode: 503);
    return await db.Database.CanConnectAsync() ? Results.Text("ok") : Results.StatusCode(503);
}).AllowAnonymous();

app.MapRazorPages();

var estadoApp = app.Services.GetRequiredService<EstadoApp>();
await estadoApp.EvaluarAsync(app.Services.GetRequiredService<ConexionStore>(), app.Services.GetRequiredService<ConexionProbador>(), app.Logger);

if (estadoApp.Fase == FaseApp.Activa)
{
    await InicializadorSeguridad.EjecutarAsync(app.Services);

    // Verifica las relaciones proyecto / asignaciones / cierre y corrige las incoherencias heredadas.
    try
    {
        using var scope = app.Services.CreateScope();
        var db = scope.ServiceProvider.GetRequiredService<PortafolioQAContext>();
        foreach (var proyecto in await db.Proyectos.ToListAsync())
            await CierreProyecto.RecalcularAvanceAsync(db, proyecto);
        var cambios = await CierreProyecto.ReparaCoherenciaAsync(db);
        foreach (var c in cambios) app.Logger.LogWarning("Coherencia: {Cambio}", c);
    }
    catch (Exception ex)
    {
        app.Logger.LogError(ex, "No se pudo verificar la coherencia de proyectos y asignaciones (¿migraciones pendientes?).");
    }
}
else
{
    var store = app.Services.GetRequiredService<ConexionStore>();
    app.Logger.LogWarning("Aplicación sin activar ({Fase}): {Detalle} Use /Setup{Codigo}.", estadoApp.Fase, estadoApp.Detalle,
        estadoApp.Fase == FaseApp.Pendiente ? $" con el código de {store.ArchivoCodigo}" : " como Administrador");
    if (estadoApp.Fase == FaseApp.Pendiente && store.AsegurarCodigo() is null) app.Logger.LogError("{Motivo}", store.ErrorCodigo);
}

app.Run();
