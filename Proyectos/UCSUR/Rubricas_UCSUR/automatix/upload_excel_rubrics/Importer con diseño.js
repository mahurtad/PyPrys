// ==UserScript==
// @name        Rubric Importer from Excel (Enhanced UI)
// @namespace   https://github.com/jamesjonesmath/canvancement
// @description Create rubrics from Excel with progress tracking and preview
// @include     https://*.instructure.com/courses/*/rubrics
// @include     https://*.instructure.com/accounts/*/rubrics
// @include     https://cientificavirtual.cientifica.edu.pe/courses/*/rubrics
// @version     3.2
// @grant       GM_addStyle
// @require     https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function() {
    'use strict';

    // Configuración principal
    const config = {
        debugMode: false,
        maxFileSizeMB: 5,
        allowedExtensions: ['xlsx', 'xls']
    };

    // Estado de la aplicación
    const state = {
        rubricsToImport: [],
        currentImportIndex: 0,
        isImporting: false
    };

    // Elementos UI principales
    const ui = {
        dialog: null,
        fileInput: null,
        previewContainer: null,
        progressContainer: null,
        progressBar: null,
        progressText: null,
        importBtn: null,
        cancelBtn: null,
        errorDisplay: null,
        browseBtn: null,
        fileNameDisplay: null
    };

    // Inicialización
    function init() {
        if (isRubricsPage()) {
            addImportButton();
            loadStyles();
        }
    }

    // Verificar si estamos en la página correcta
    function isRubricsPage() {
        return /^\/(course|account)s\/\d+\/rubrics$/.test(window.location.pathname);
    }

    // Añadir botón de importación (mejorado para visibilidad)
    function addImportButton() {
        const buttonId = 'jj_rubric_excel_enhanced';
        const existingBtn = document.getElementById(buttonId);

        if (!existingBtn) {
            const parent = document.querySelector('aside#right-side');
            if (parent) {
                // Crear contenedor si no existe
                let container = parent.querySelector('.rubric-importer-container');
                if (!container) {
                    container = document.createElement('div');
                    container.className = 'rubric-importer-container';
                    parent.insertBefore(container, parent.firstChild);
                }

                // Crear botón
                const btn = document.createElement('a');
                btn.className = 'btn button-sidebar-wide';
                btn.id = buttonId;
                btn.innerHTML = '<i class="icon-upload"></i> Import Rubrics from Excel';
                btn.addEventListener('click', showMainDialog);

                // Añadir al contenedor
                container.appendChild(btn);

                // Forzar visibilidad
                btn.style.display = 'block';
                btn.style.visibility = 'visible';
                btn.style.opacity = '1';
            }
        }
    }

    // Cargar estilos CSS (mejorados)
    function loadStyles() {
        const css = `
            /* Contenedor del botón */
            .rubric-importer-container {
                margin-bottom: 15px;
            }

            /* Estilos para el botón */
            #jj_rubric_excel_enhanced {
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                position: static !important;
            }

            /* Diálogo principal */
            .rubric-importer-dialog {
                font-family: 'Lato', 'Helvetica Neue', Arial, sans-serif;
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: white;
                z-index: 9999;
                box-shadow: 0 2px 10px rgba(0,0,0,0.2);
                border-radius: 5px;
                width: 700px;
                max-width: 95%;
                max-height: 90vh;
                display: flex;
                flex-direction: column;
                padding: 25px;
            }

            /* Overlay */
            .dialog-overlay {
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: rgba(0,0,0,0.5);
                z-index: 9998;
            }

            /* Contenedor de archivo */
            .file-input-container {
                margin-bottom: 20px;
                padding: 20px;
                border: 2px dashed #d9d9d9;
                border-radius: 6px;
                text-align: center;
                transition: all 0.3s;
                background: #f8f9fa;
            }

            .file-input-container:hover {
                border-color: #008EE2;
                background: #f0f7fc;
            }

            /* Vista previa */
            .preview-container {
                max-height: 350px;
                overflow-y: auto;
                margin: 20px 0;
                border: 1px solid #e1e1e1;
                padding: 20px;
                display: none;
                background: #f9f9f9;
                border-radius: 6px;
            }

            /* Barra de progreso */
            .progress-container {
                display: none;
                margin: 20px 0;
            }

            .progress-bar {
                height: 24px;
                background: #f0f0f0;
                border-radius: 6px;
                overflow: hidden;
                margin-bottom: 8px;
                box-shadow: inset 0 1px 3px rgba(0,0,0,0.1);
            }

            .progress-fill {
                height: 100%;
                background: linear-gradient(to right, #4CAF50, #8BC34A);
                width: 0%;
                transition: width 0.4s ease;
            }

            /* Mensajes de error */
            .error-message {
                color: #D32F2F;
                margin: 15px 0;
                padding: 12px;
                background: #FFEBEE;
                border-radius: 6px;
                display: none;
                border-left: 4px solid #D32F2F;
            }

            /* Botones */
            .dialog-buttons {
                display: flex;
                justify-content: flex-end;
                margin-top: 20px;
                gap: 12px;
            }

            .dialog-btn {
                padding: 10px 20px;
                border-radius: 6px;
                cursor: pointer;
                font-weight: bold;
                transition: all 0.2s;
                font-size: 14px;
            }

            .primary-btn {
                background: #008EE2;
                color: white;
                border: none;
            }

            .primary-btn:hover {
                background: #0077C6;
            }

            .primary-btn:disabled {
                background: #B0BEC5;
                cursor: not-allowed;
                opacity: 0.7;
            }

            .secondary-btn {
                background: white;
                border: 1px solid #B0BEC5;
                color: #546E7A;
            }

            .secondary-btn:hover {
                background: #f5f5f5;
            }

            /* Preview items */
            .rubric-preview {
                margin-bottom: 25px;
                border-left: 4px solid #008EE2;
                padding-left: 15px;
                background: white;
                border-radius: 6px;
                padding: 15px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }

            .rubric-title {
                font-size: 16px;
                font-weight: bold;
                color: #263238;
                margin-bottom: 10px;
                display: flex;
                justify-content: space-between;
            }

            .rubric-points {
                color: #4CAF50;
                font-weight: normal;
            }

            .criteria-item {
                margin: 12px 0;
                padding: 12px;
                background: #f8f9fa;
                border-radius: 4px;
            }

            .criteria-name {
                font-weight: 600;
                color: #37474F;
                margin-bottom: 8px;
            }

            .rating-item {
                margin: 8px 0;
                padding: 8px 12px;
                border-left: 3px solid #B0BEC5;
                background: white;
                border-radius: 4px;
            }

            .rating-points {
                font-weight: bold;
                color: #008EE2;
            }

            .rating-desc {
                font-size: 13px;
                color: #546E7A;
                margin-top: 4px;
                line-height: 1.4;
            }

            /* Texto del nombre del archivo */
            .file-name-display {
                margin-top: 12px;
                font-size: 14px;
                color: #546E7A;
                word-break: break-all;
            }

            /* Títulos */
            .dialog-title {
                margin-top: 0;
                margin-bottom: 20px;
                color: #263238;
                font-size: 20px;
                font-weight: 600;
            }

            .preview-title {
                margin-top: 0;
                margin-bottom: 15px;
                color: #37474F;
                font-size: 16px;
                border-bottom: 1px solid #e1e1e1;
                padding-bottom: 8px;
            }

            /* Animaciones */
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }

            .spinner {
                animation: spin 1s linear infinite;
                display: inline-block;
                margin-right: 8px;
            }
        `;

        GM_addStyle(css);
    }

    // Mostrar diálogo principal (mejorado)
    function showMainDialog() {
        // Crear overlay
        const overlay = document.createElement('div');
        overlay.className = 'dialog-overlay';
        overlay.addEventListener('click', closeDialog);

        // Crear diálogo
        ui.dialog = document.createElement('div');
        ui.dialog.className = 'rubric-importer-dialog';
        ui.dialog.innerHTML = `
            <h2 class="dialog-title">Import Rubrics from Excel</h2>

            <div class="file-input-container">
                <label for="excelFileInput" style="display: block; margin-bottom: 12px; font-weight: bold; font-size: 15px;">
                    <i class="icon-upload" style="margin-right: 8px;"></i> Select Excel File
                </label>
                <input type="file" id="excelFileInput" accept=".xlsx,.xls" style="display: none;">
                <button id="browseBtn" class="dialog-btn secondary-btn" style="margin: 0 auto; font-size: 14px;">
                    Browse Files...
                </button>
                <div id="fileNameDisplay" class="file-name-display">
                    No file selected
                </div>
            </div>

            <div id="previewContainer" class="preview-container">
                <h3 class="preview-title">Rubrics Preview</h3>
                <div id="rubricsPreview"></div>
            </div>

            <div id="progressContainer" class="progress-container">
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill"></div>
                </div>
                <div id="progressText" style="text-align: center; font-size: 14px; color: #546E7A;">
                    Ready to import
                </div>
            </div>

            <div id="errorMessage" class="error-message"></div>

            <div class="dialog-buttons">
                <button id="cancelBtn" class="dialog-btn secondary-btn">Cancel</button>
                <button id="importBtn" class="dialog-btn primary-btn" disabled>Import Rubrics</button>
            </div>
        `;

        // Añadir al documento
        document.body.appendChild(overlay);
        document.body.appendChild(ui.dialog);

        // Inicializar elementos UI
        ui.fileInput = document.getElementById('excelFileInput');
        ui.previewContainer = document.getElementById('previewContainer');
        ui.progressContainer = document.getElementById('progressContainer');
        ui.progressBar = document.getElementById('progressFill');
        ui.progressText = document.getElementById('progressText');
        ui.importBtn = document.getElementById('importBtn');
        ui.cancelBtn = document.getElementById('cancelBtn');
        ui.errorDisplay = document.getElementById('errorMessage');
        ui.browseBtn = document.getElementById('browseBtn');
        ui.fileNameDisplay = document.getElementById('fileNameDisplay');
        ui.rubricsPreview = document.getElementById('rubricsPreview');

        // Configurar eventos
        ui.browseBtn.addEventListener('click', () => ui.fileInput.click());
        ui.fileInput.addEventListener('change', handleFileSelect);
        ui.cancelBtn.addEventListener('click', closeDialog);
        ui.importBtn.addEventListener('click', startImportProcess);
    }

    // Manejar selección de archivo
    function handleFileSelect(event) {
        resetUI();

        const file = event.target.files[0];
        if (!file) return;

        // Validar archivo
        const fileError = validateFile(file);
        if (fileError) {
            showError(fileError);
            return;
        }

        // Actualizar UI
        ui.fileNameDisplay.textContent = `Selected: ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
        ui.previewContainer.style.display = 'block';
        ui.rubricsPreview.innerHTML = `
            <div style="text-align: center; padding: 30px;">
                <div style="display: inline-block;">
                    <i class="icon-spinner spinner" style="font-size: 20px;"></i>
                </div>
                <div style="margin-top: 10px;">Processing file...</div>
            </div>
        `;

        // Leer archivo
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                processExcelFile(e.target.result);
            } catch (error) {
                showError(`Error processing file: ${error.message}`);
                logError(error);
            }
        };

        reader.onerror = function() {
            showError('Error reading file. Please try again.');
        };

        reader.readAsArrayBuffer(file);
    }

    // Validar archivo seleccionado
    function validateFile(file) {
        if (!file) return 'No file selected';

        // Validar tamaño
        if (file.size > config.maxFileSizeMB * 1024 * 1024) {
            return `File is too large (max ${config.maxFileSizeMB}MB allowed)`;
        }

        // Validar extensión
        const fileExt = file.name.split('.').pop().toLowerCase();
        if (!config.allowedExtensions.includes(fileExt)) {
            return `Invalid file type. Please use .xlsx or .xls files`;
        }

        return null;
    }

    // Procesar archivo Excel
    function processExcelFile(data) {
        const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
        state.rubricsToImport = [];

        // Procesar cada hoja
        workbook.SheetNames.forEach(sheetName => {
            try {
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                // Filtrar filas vacías y saltar cabecera
                const rows = jsonData.slice(1).filter(row =>
                    row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '')
                );

                if (rows.length > 0) {
                    const criteria = parseRubricData(rows);
                    if (criteria.length > 0) {
                        state.rubricsToImport.push({
                            title: sheetName,
                            criteria: criteria,
                            pointsPossible: calculateTotalPoints(criteria)
                        });
                    }
                }
            } catch (error) {
                logError(`Error processing sheet "${sheetName}": ${error.message}`);
            }
        });

        // Mostrar vista previa
        showRubricsPreview();
    }

    // Parsear datos de rúbrica
    function parseRubricData(rows) {
        const criteria = [];

        rows.forEach(row => {
            try {
                if (row.length >= 4) { // Mínimo: nombre + 1 rating (puntos, título, descripción)
                    const criterion = {
                        description: String(row[0]).trim(),
                        ratings: []
                    };

                    // Procesar ratings (grupos de 3 columnas)
                    for (let i = 1; i < row.length; i += 3) {
                        if (i + 2 < row.length) {
                            const points = parseFloat(row[i]);
                            if (!isNaN(points)) {
                                criterion.ratings.push({
                                    points: points,
                                    description: String(row[i+1] || '').trim(),
                                    long_description: String(row[i+2] || '').trim()
                                });
                            }
                        }
                    }

                    // Ordenar ratings por puntos (mayor a menor)
                    if (criterion.ratings.length > 0) {
                        criterion.ratings.sort((a, b) => b.points - a.points);
                        criteria.push(criterion);
                    }
                }
            } catch (error) {
                logError(`Error parsing row: ${error.message}`);
            }
        });

        return criteria;
    }

    // Calcular puntos totales
    function calculateTotalPoints(criteria) {
        return criteria.reduce((total, criterion) => {
            return total + (criterion.ratings.length > 0 ? Number(criterion.ratings[0].points) : 0);
        }, 0);
    }

    // Mostrar vista previa (mejorada)
    function showRubricsPreview() {
        if (state.rubricsToImport.length === 0) {
            ui.rubricsPreview.innerHTML = `
                <div style="color: #FF5722; text-align: center; padding: 30px;">
                    No valid rubrics found in the selected file.
                </div>
            `;
            ui.importBtn.disabled = true;
            return;
        }

        let html = '';

        state.rubricsToImport.forEach((rubric, index) => {
            html += `
                <div class="rubric-preview">
                    <div class="rubric-title">
                        <span>${index + 1}. ${rubric.title}</span>
                        <span class="rubric-points">${rubric.pointsPossible} points</span>
                    </div>

                    <div style="margin-left: 10px;">
                        ${rubric.criteria.map((criterion, critIndex) => `
                            <div class="criteria-item">
                                <div class="criteria-name">
                                    ${critIndex + 1}. ${criterion.description}
                                </div>

                                <div style="margin-left: 10px;">
                                    ${criterion.ratings.map(rating => `
                                        <div class="rating-item">
                                            <div>
                                                <span class="rating-points">${rating.points} points:</span>
                                                ${rating.description}
                                            </div>
                                            ${rating.long_description ? `
                                                <div class="rating-desc">
                                                    ${rating.long_description}
                                                </div>
                                            ` : ''}
                                        </div>
                                    `).join('')}
                                </div>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
        });

        ui.rubricsPreview.innerHTML = html;
        ui.importBtn.disabled = false;
    }

    // Iniciar proceso de importación
    function startImportProcess() {
        if (state.isImporting || state.rubricsToImport.length === 0) return;

        state.isImporting = true;
        state.currentImportIndex = 0;

        // Configurar UI para importación
        ui.importBtn.disabled = true;
        ui.cancelBtn.disabled = true;
        ui.fileInput.disabled = true;
        ui.browseBtn.disabled = true;
        ui.progressContainer.style.display = 'block';
        ui.progressBar.style.width = '0%';
        ui.progressText.textContent = 'Starting import...';

        // Obtener información de asociación
        const match = window.location.pathname.match(/^\/(course|account)s\/(\d+)\/rubrics$/);
        if (!match) {
            showError('Invalid page for rubric creation');
            resetImportState();
            return;
        }

        const association = {
            type: match[1].charAt(0).toUpperCase() + match[1].slice(1),
            id: match[2]
        };

        // Iniciar importación secuencial
        importNextRubric(association);
    }

    // Importar rúbricas una por una
    function importNextRubric(association) {
        if (state.currentImportIndex >= state.rubricsToImport.length) {
            // Importación completada
            ui.progressText.innerHTML = `
                <div style="color: #4CAF50; font-weight: bold;">
                    <i class="icon-check" style="margin-right: 5px;"></i>
                    Successfully imported ${state.rubricsToImport.length} rubrics!
                </div>
            `;
            ui.importBtn.textContent = 'Done';
            ui.cancelBtn.textContent = 'Close';
            ui.cancelBtn.disabled = false;
            state.isImporting = false;

            // Recargar después de 2 segundos
            setTimeout(() => {
                closeDialog();
                window.location.reload();
            }, 2000);
            return;
        }

        const rubric = state.rubricsToImport[state.currentImportIndex];
        const progressPercent = Math.round((state.currentImportIndex / state.rubricsToImport.length) * 100);

        // Actualizar UI de progreso
        ui.progressBar.style.width = `${progressPercent}%`;
        ui.progressText.innerHTML = `
            Importing <strong>${state.currentImportIndex + 1} of ${state.rubricsToImport.length}</strong>:
            ${rubric.title}
        `;

        // Preparar datos para Canvas
        const rubricData = {
            'rubric': {
                'title': rubric.title,
                'points_possible': rubric.pointsPossible,
                'free_form_criterion_comments': 0,
                'criteria': rubric.criteria
            },
            'rubric_association': {
                'id': '',
                'use_for_grading': 0,
                'hide_score_total': 0,
                'association_type': association.type,
                'association_id': association.id,
                'purpose': 'bookmark'
            },
            'title': rubric.title,
            'points_possible': rubric.pointsPossible,
            'rubric_id': 'new',
            'rubric_association_id': '',
            'skip_updating_points_possible': 0,
            'authenticity_token': getCsrfToken()
        };

        // Enviar a Canvas
        sendRubricToCanvas(rubricData,
            () => {
                // Éxito - procesar siguiente
                state.currentImportIndex++;
                setTimeout(() => importNextRubric(association), 300);
            },
            (error) => {
                // Error - registrar pero continuar
                showError(`Error importing "${rubric.title}": ${error}`);
                logError(`Import failed for "${rubric.title}": ${error}`);
                state.currentImportIndex++;
                setTimeout(() => importNextRubric(association), 300);
            }
        );
    }

    // Enviar rúbrica a Canvas
    function sendRubricToCanvas(data, successCallback, errorCallback) {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', window.location.pathname, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded; charset=UTF-8');
        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');

        xhr.onload = function() {
            if (xhr.status >= 200 && xhr.status < 300) {
                successCallback();
            } else {
                let errorMsg = xhr.statusText;
                try {
                    const response = JSON.parse(xhr.responseText);
                    if (response.errors) {
                        errorMsg = response.errors.join(', ');
                    }
                } catch (e) {
                    logError('Error parsing response:', e);
                }
                errorCallback(errorMsg);
            }
        };

        xhr.onerror = function() {
            errorCallback('Network error occurred');
        };

        // Convertir datos a formato x-www-form-urlencoded
        const formData = Object.keys(data).map(key => {
            const value = typeof data[key] === 'object' ? JSON.stringify(data[key]) : data[key];
            return `${encodeURIComponent(key)}=${encodeURIComponent(value)}`;
        }).join('&');

        xhr.send(formData);
    }

    // Obtener token CSRF
    function getCsrfToken() {
        const csrfRegex = /^_csrf_token=(.*)$/;
        const cookies = document.cookie.split(';');

        for (let i = 0; i < cookies.length; i++) {
            const cookie = cookies[i].trim();
            const match = csrfRegex.exec(cookie);
            if (match) {
                return decodeURIComponent(match[1]);
            }
        }

        return '';
    }

    // Mostrar mensaje de error
    function showError(message) {
        ui.errorDisplay.textContent = message;
        ui.errorDisplay.style.display = 'block';

        // Auto-ocultar después de 5 segundos
        setTimeout(() => {
            ui.errorDisplay.style.display = 'none';
        }, 5000);
    }

    // Reiniciar estado de UI
    function resetUI() {
        ui.previewContainer.style.display = 'none';
        ui.progressContainer.style.display = 'none';
        ui.errorDisplay.style.display = 'none';
        ui.importBtn.disabled = true;
        state.rubricsToImport = [];
    }

    // Reiniciar estado de importación
    function resetImportState() {
        state.isImporting = false;
        state.currentImportIndex = 0;
        ui.importBtn.disabled = false;
        ui.cancelBtn.disabled = false;
        ui.fileInput.disabled = false;
        ui.browseBtn.disabled = false;
    }

    // Cerrar diálogo
    function closeDialog() {
        if (ui.dialog && ui.dialog.parentNode) {
            document.body.removeChild(ui.dialog);
        }

        const overlay = document.querySelector('.dialog-overlay');
        if (overlay) {
            document.body.removeChild(overlay);
        }

        resetUI();
        resetImportState();
    }

    // Registrar errores (solo en debug)
    function logError(...args) {
        if (config.debugMode) {
            console.error('[Rubric Importer]', ...args);
        }
    }

    // Inicializar cuando el DOM esté listo
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        setTimeout(init, 100);
    } else {
        document.addEventListener('DOMContentLoaded', init);
    }
})();