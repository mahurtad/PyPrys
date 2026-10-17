// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

// Barras de progreso con el porcentaje dentro: si el avance es 0 la barra no tiene ancho, así que el texto se muestra a su lado.
document.querySelectorAll('.progress-texto .progress-bar').forEach(function (barra) {
    if (parseFloat(barra.style.width) === 0 && barra.textContent.trim()) {
        var etiqueta = document.createElement('span');
        etiqueta.className = 'progress-cero';
        etiqueta.textContent = barra.textContent.trim();
        barra.textContent = '';
        barra.parentNode.appendChild(etiqueta);
    }
});
