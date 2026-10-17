// Filtros del Dashboard: al cambiar un filtro se vacían los que dependen de él (para no combinar elementos de otro
// proyecto o campaña) y se envía el formulario. Sin JavaScript queda el botón "Aplicar filtros".
(function () {
    var form = document.getElementById('filtros-dashboard');
    if (!form) return;
    form.querySelectorAll('select').forEach(function (sel) {
        sel.addEventListener('change', function () {
            (sel.getAttribute('data-limpia') || '').split(',').forEach(function (id) {
                var dep = id && document.getElementById(id);
                if (dep) dep.value = '';
            });
            // Los vacíos no se envían (URL limpia); form.submit() no dispara el evento submit, por eso se hace aquí.
            form.querySelectorAll('select').forEach(function (s) { if (!s.value) s.disabled = true; });
            form.submit();
        });
    });
    // Los selects deshabilitados (sin proyecto elegido) no se envían; los vacíos tampoco ensucian la URL.
    form.addEventListener('submit', function () {
        form.querySelectorAll('select').forEach(function (s) { if (!s.value) s.disabled = true; });
    });
})();
