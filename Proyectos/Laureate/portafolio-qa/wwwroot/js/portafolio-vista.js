(function () {
    const btnTarjetas = document.getElementById('btn-vista-tarjetas');
    const btnTabla = document.getElementById('btn-vista-tabla');
    const vistaTarjetas = document.getElementById('vista-tarjetas');
    const vistaTabla = document.getElementById('vista-tabla');

    if (!btnTarjetas || !btnTabla) {
        return;
    }

    function mostrarTarjetas() {
        vistaTarjetas.classList.remove('d-none');
        vistaTabla.classList.add('d-none');
        btnTarjetas.classList.add('active');
        btnTabla.classList.remove('active');
        localStorage.setItem('portafolio-vista', 'tarjetas');
    }

    function mostrarTabla() {
        vistaTabla.classList.remove('d-none');
        vistaTarjetas.classList.add('d-none');
        btnTabla.classList.add('active');
        btnTarjetas.classList.remove('active');
        localStorage.setItem('portafolio-vista', 'tabla');
    }

    btnTarjetas.addEventListener('click', mostrarTarjetas);
    btnTabla.addEventListener('click', mostrarTabla);

    try {
        if (localStorage.getItem('portafolio-vista') === 'tabla') {
            mostrarTabla();
        }
    } catch (e) {
        // almacenamiento no disponible; se mantiene la vista de tarjetas por defecto
    }
})();
