using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddDbContext<PortafolioQAContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("PortafolioQAContext")));

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
    options.Conventions.AuthorizeFolder("/Admin", Politicas.PuedeEditar);
    options.Conventions.AuthorizeFolder("/Admin/Usuarios", Politicas.SoloAdmin);
});

builder.Services.AddScoped<CapacidadService>();

// Integraciones externas: en la etapa 1 no hay conexión real configurada.
// Las etapas 4 y 5 registrarán implementaciones reales sin cambiar quien las consume.
builder.Services.AddScoped<IIntegracionCertificacion, IntegracionCertificacionNoConfigurada>();
builder.Services.AddScoped<IIntegracionDefectos, IntegracionDefectosNoConfigurada>();

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

app.UseAuthentication();
app.UseAuthorization();

// Verificación de disponibilidad para IIS / monitoreo: anónima y sin detalles internos.
app.MapGet("/healthz", async (PortafolioQAContext db) =>
    await db.Database.CanConnectAsync() ? Results.Text("ok") : Results.StatusCode(503))
    .AllowAnonymous();

app.MapRazorPages();

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

app.Run();
