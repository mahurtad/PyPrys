using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Pages.Admin.Bugzilla;

/// <summary>
/// Sistema → Integración Bugzilla (solo Administrador). Configura la dirección de Bugzilla, la clave de la cuenta de servicio
/// (grupo editcomponents: crea productos) y la clave de lectura (consulta de defectos). Todo se prueba contra Bugzilla antes de
/// guardar, se guarda cifrado y se pide la contraseña del administrador. Las claves nunca se vuelven a mostrar.
/// </summary>
public class IndexModel(BugzillaCredencialesStore store, BugzillaProbador probador, UserManager<UsuarioApp> usuarios,
    PortafolioQAContext db) : PageModel
{
    public class EntradaModel
    {
        public string? BaseUrl { get; set; }
        public string? CuentaServicio { get; set; }
        public string? ApiKey { get; set; }
        public string? ApiKeyLectura { get; set; }
        public string? ClaveAdmin { get; set; }
        public bool ConfirmaCambioUrl { get; set; }
    }

    [BindProperty]
    public EntradaModel Entrada { get; set; } = new();

    public string? BaseUrlActual => store.Actual.BaseUrl;
    public OrigenBugzilla Origen => store.Origen;
    public string CuentaActual => store.CuentaServicio;
    public bool ClaveServicioConfigurada => !string.IsNullOrWhiteSpace(store.Actual.AdminApiKey);
    public bool ClaveLecturaConfigurada => !string.IsNullOrWhiteSpace(store.Actual.ApiKey);
    public string? GuardadoPor => store.GuardadoPor;
    public DateTime? GuardadoEn => store.GuardadoEn;
    public string? Aviso => store.Aviso;

    public ResultadoPruebaBugzilla? PruebaServicio { get; private set; }
    public ResultadoPruebaBugzilla? PruebaLectura { get; private set; }
    /// <summary>Proyectos con producto ya creado en la dirección actual: al cambiarla quedarían en Error para revisión.</summary>
    public int VinculosAfectados { get; private set; }
    public string? Error { get; private set; }

    public void OnGet()
    {
        Entrada.BaseUrl = store.Actual.BaseUrl;
        Entrada.CuentaServicio = store.CuentaServicio;
    }

    public async Task<IActionResult> OnPostProbarAsync()
    {
        var c = Resolver();
        if (c.Error is not null) Error = c.Error;
        else await ProbarAsync(c);
        LimpiarSecretos();
        return Page();
    }

    public async Task<IActionResult> OnPostGuardarAsync()
    {
        if (!await ConfirmarAdminAsync()) return Page();

        var c = Resolver();
        if (c.Error is not null) { Error = c.Error; LimpiarSecretos(); return Page(); }

        if (!await ProbarAsync(c))
        {
            Error = "No se guardó: corrija lo marcado abajo.";
            LimpiarSecretos();
            return Page();
        }

        if (c.CambiaUrl)
        {
            var ambiente = BugzillaProvision.NormalizarAmbiente(c.Url);
            VinculosAfectados = await db.VinculosBugzilla.CountAsync(v =>
                (v.ProductoId != null || v.ComponenteId != null) && v.BugzillaAmbiente != ambiente, HttpContext.RequestAborted);
            if (VinculosAfectados > 0 && !Entrada.ConfirmaCambioUrl)
            {
                Error = $"No se guardó: {VinculosAfectados} proyecto(s) ya tienen producto creado en la dirección actual. Con otra dirección " +
                        "quedarán en Error para su revisión (no se borra ni se reasigna nada). Marque la confirmación para continuar.";
                LimpiarSecretos();
                return Page();
            }
        }

        var rutaRest = PruebaServicio?.RutaRest ?? PruebaLectura?.RutaRest;
        store.Guardar(c.Url, Entrada.ApiKey, c.Cuenta, Entrada.ApiKeyLectura, rutaRest, User.Identity?.Name ?? "desconocido");
        TempData["Mensaje"] = "Conexión guardada y activa. Los proyectos que quedaron Pendientes se reintentan desde su detalle.";
        return RedirectToPage();
    }

    public async Task<IActionResult> OnPostQuitarAsync()
    {
        if (!await ConfirmarAdminAsync()) return Page();
        store.Quitar(User.Identity?.Name ?? "desconocido");
        TempData["Mensaje"] = "Se quitó lo guardado desde la aplicación. " +
            (store.Origen == OrigenBugzilla.Servidor ? "Rige la configuración del servidor." : "Bugzilla queda sin configurar: los productos nuevos quedarán Pendientes.");
        return RedirectToPage();
    }

    private sealed record Candidatos(string? Url, string? Cuenta, string? ClaveServicio, string? ClaveLectura, bool CambiaUrl, string? Error);

    /// <summary>
    /// Valores con los que se prueba y guarda: lo escrito o, si está en blanco, lo vigente. Con otra dirección las claves vigentes
    /// no sirven (son de otra instalación de Bugzilla), así que hay que pegarlas de nuevo.
    /// </summary>
    private Candidatos Resolver()
    {
        var actual = store.Actual;
        static string? Dato(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        var url = Dato(Entrada.BaseUrl) is { } u ? BugzillaCredencialesStore.NormalizarUrl(u) : actual.BaseUrl;
        var cambia = !string.Equals(BugzillaProvision.NormalizarAmbiente(url), BugzillaProvision.NormalizarAmbiente(actual.BaseUrl), StringComparison.Ordinal);
        var servicio = Dato(Entrada.ApiKey) ?? (cambia ? null : actual.AdminApiKey);
        var lectura = Dato(Entrada.ApiKeyLectura) ?? (cambia ? null : actual.ApiKey);
        var cuenta = Dato(Entrada.CuentaServicio) ?? actual.CuentaServicio;

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https"))
            return new(url, cuenta, servicio, lectura, cambia, "Indique la dirección de Bugzilla, por ejemplo https://bugzilla.su-dominio.");
        if (cambia && (Dato(Entrada.ApiKey) is null || (ClaveLecturaConfigurada && Dato(Entrada.ApiKeyLectura) is null)))
            return new(url, cuenta, servicio, lectura, cambia,
                "Con otra dirección hay que pegar de nuevo las claves (son de otra instalación de Bugzilla): la de servicio" +
                (ClaveLecturaConfigurada ? " y la de lectura." : "."));
        return new(url, cuenta, servicio, lectura, cambia, null);
    }

    /// <summary>Prueba lo que haya (servicio y/o lectura). Devuelve false si algo no es aceptable.</summary>
    private async Task<bool> ProbarAsync(Candidatos c)
    {
        var ok = true;
        if (c.ClaveServicio is not null)
        {
            PruebaServicio = await probador.ProbarAsync(c.Url, c.ClaveServicio, c.Cuenta, HttpContext.RequestAborted);
            ok &= PruebaServicio.Aceptable;
        }
        if (c.ClaveLectura is not null)
        {
            PruebaLectura = await probador.ProbarLecturaAsync(c.Url, c.ClaveLectura, HttpContext.RequestAborted);
            ok &= PruebaLectura.Aceptable;
        }
        if (PruebaServicio is null && PruebaLectura is null)
        {
            Error = "No hay ninguna clave que probar: pegue la clave de servicio y/o la de lectura.";
            return false;
        }
        return ok;
    }

    private async Task<bool> ConfirmarAdminAsync()
    {
        var usuario = await usuarios.GetUserAsync(User);
        if (usuario is null || string.IsNullOrEmpty(Entrada.ClaveAdmin) || !await usuarios.CheckPasswordAsync(usuario, Entrada.ClaveAdmin))
        {
            Error = "Confirme con la contraseña de su usuario de Portafolio (la que usa para iniciar sesión en esta aplicación). No es la de ninguna cuenta de Bugzilla.";
            LimpiarSecretos();
            return false;
        }
        return true;
    }

    private void LimpiarSecretos()
    {
        Entrada.ApiKey = null;
        Entrada.ApiKeyLectura = null;
        Entrada.ClaveAdmin = null;
        ModelState.Clear();
    }
}
