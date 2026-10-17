(function () {
    var CLAVE = 'portafolioqa.menu';
    var raiz = document.documentElement;
    var boton = document.getElementById('btn-plegar-menu');
    var panel = document.getElementById('menu-navegacion');
    var movil = window.matchMedia('(max-width: 991.98px)');

    // Escritorio: plegar/desplegar el menú lateral y recordar la preferencia (si no hay almacenamiento, sigue funcionando).
    function actualizarBoton() {
        if (!boton) return;
        var plegado = raiz.classList.contains('menu-plegado');
        var texto = plegado ? 'Mostrar el menú lateral' : 'Ocultar el menú lateral';
        boton.setAttribute('aria-expanded', plegado ? 'false' : 'true');
        boton.setAttribute('aria-label', texto);
        boton.setAttribute('title', texto);
    }

    if (boton) {
        boton.addEventListener('click', function () {
            var plegado = raiz.classList.toggle('menu-plegado');
            try { localStorage.setItem(CLAVE, plegado ? 'plegado' : 'visible'); } catch (e) { /* sin almacenamiento */ }
            actualizarBoton();
        });
        actualizarBoton();
    }

    // Móvil: Escape cierra el menú desplegado y devuelve el foco al botón que lo abre.
    document.addEventListener('keydown', function (e) {
        if (e.key !== 'Escape' || !panel || !movil.matches || !panel.classList.contains('show')) return;
        var activo = document.activeElement;
        var dentro = !activo || activo === document.body || panel.contains(activo) || activo.matches('[data-bs-target="#menu-navegacion"]');
        if (!dentro) return;
        // tabler.min.js expone los componentes de Bootstrap en window.tabler
        (window.tabler || window.bootstrap).Collapse.getOrCreateInstance(panel, { toggle: false }).hide();
        var toggler = document.querySelector('[data-bs-target="#menu-navegacion"]');
        if (toggler) toggler.focus();
    });
})();
