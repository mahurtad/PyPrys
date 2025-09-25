// ==UserScript==
// @name        Rubric Importer from Excel por Curso (Enhanced UI)
// @namespace   https://github.com/jamesjonesmath/canvancement
// @description Create rubrics by importing from an Excel file (one rubric per sheet) with improved UI
// @include     https://*.instructure.com/courses/*/rubrics
// @include     https://*.instructure.com/accounts/*/rubrics
// @include     https://cientificavirtual.cientifica.edu.pe/courses/*/rubrics
// @version     1.1
// @grant       GM_addStyle
// @require     https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function() {
  'use strict';
  var assocRegex = new RegExp('^/(course|account)s/([0-9]+)/rubrics$');
  var errors = [];
  var outcomes = [];
  var pendingOutcomes = 0;
  var criteria = [];
  var rubricTitle;
  var rubricAssociation;

  // Elementos UI principales
  const ui = {
      dialog: null,
      fileInput: null,
      progressContainer: null,
      progressBar: null,
      progressText: null,
      importBtn: null,
      cancelBtn: null,
      errorDisplay: null,
      browseBtn: null,
      fileNameDisplay: null
  };

  if (assocRegex.test(window.location.pathname)) {
      add_button();
      loadStyles();
  }

  // Cargar estilos CSS
  function loadStyles() {
      const css = `
          .rubric-importer-container {
              margin-bottom: 15px;
          }
          #jj_rubric_excel {
              display: block !important;
              visibility: visible !important;
              opacity: 1 !important;
              position: static !important;
          }
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
              width: 500px;
              max-width: 95%;
              max-height: 90vh;
              display: flex;
              flex-direction: column;
              padding: 25px;
          }
          .dialog-overlay {
              position: fixed;
              top: 0;
              left: 0;
              right: 0;
              bottom: 0;
              background: rgba(0,0,0,0.5);
              z-index: 9998;
          }
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
          .error-message {
              color: #D32F2F;
              margin: 15px 0;
              padding: 12px;
              background: #FFEBEE;
              border-radius: 6px;
              display: none;
              border-left: 4px solid #D32F2F;
              white-space: pre-wrap;
              max-height: 200px;
              overflow-y: auto;
          }
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
          .file-name-display {
              margin-top: 12px;
              font-size: 14px;
              color: #546E7A;
              word-break: break-all;
          }
          .dialog-title {
              margin-top: 0;
              margin-bottom: 20px;
              color: #263238;
              font-size: 20px;
              font-weight: 600;
          }
      `;
      GM_addStyle(css);
  }

  function checkPointsRow(cols) {
      if (typeof cols === 'undefined' || cols.length < 3 || isPoints(cols[0])) {
          return false;
      }
      var points = [];
      var isValid = true;
      var i = 1;
      var firstCol;
      var block = false;
      while (isValid && i < cols.length) {
          var item = cols[i++];
          if (isPoints(item)) {
              points.push(item);
              if (typeof firstCol === 'undefined') {
                  firstCol = i - 1;
              }
          } else if (i > 2) {
              isValid = false;
          }
      }
      if (isValid && points.length > 1) {
          var order = checkMonotonic(points, true);
          if (order) {
              block = {
                  'order' : order,
                  'points' : points,
                  'start' : firstCol,
                  'end' : firstCol + points.length - 1,
                  'columns' : cols.length
              };
              if (typeof checkPointsRow.initial === 'undefined') {
                  checkPointsRow.initial = block.start;
              }
          }
      }
      return block;
  }

  function checkMonotonic(A, isStrict) {
      if (typeof A === 'undefined') {
          return 0;
      }
      if (typeof isStrict === 'undefined') {
          isStrict = true;
      }
      var order;
      var isMonotonic = true;
      var a, b;
      a = Number(A[0]);
      var k = 1;
      while (isMonotonic && k < A.length) {
          b = Number(A[k]);
          k++;
          var diff = b - a;
          a = b;
          if (Math.abs(diff) < 0.000001) {
              if (isStrict) {
                  isMonotonic = false;
              }
          } else {
              if (typeof order === 'undefined') {
                  order = diff < 0 ? -1 : 1;
              } else {
                  if (order * (diff < 0 ? -1 : 1) < 0) {
                      isMonotonic = false;
                  }
              }
          }
      }
      if (typeof order === 'undefined') {
          order = 0;
      }
      return isMonotonic ? order : 0;
  }

  function addRubric(title, criteria, association) {
      var pointsPossible = 0;
      for (var i = 0; i < criteria.length; i++) {
          if (typeof criteria[i].ignore_for_scoring === 'undefined' || !getBoolean(criteria[i].ignore_for_scoring)) {
              pointsPossible += Number(criteria[i].ratings[0].points);
          }
      }
      var F = {
          'rubric' : {
              'title' : title,
              'points_possible' : pointsPossible,
              'free_form_criterion_comments' : 0,
              'criteria' : criteria,
          },
          'rubric_association' : {
              'id' : '',
              'use_for_grading' : 0,
              'hide_score_total' : 0,
              'association_type' : association.type,
              'association_id' : association.id,
              'purpose' : 'bookmark'
          },
          'title' : title,
          'points_possible' : pointsPossible,
          'rubric_id' : 'new',
          'rubric_association_id' : '',
          'skip_updating_points_possible' : 0,
      };
      return F;
  }

  function addCriterion(item) {
      if (typeof item === 'undefined') {
          return false;
      }
      var name = item.name.replace(/\s*\\n\s*/g, ' ').replace(/\s+/g, ' ');
      var longDesc = typeof item.longDesc === 'undefined' ? '' : item.longDesc.replace(/\\n/g, '\n');
      var ratings = addRatings(item.ratings, item.points);
      var criterion = {
          'description' : name,
          'long_description' : longDesc
      };
      if (ratings !== false) {
          criterion.ratings = ratings;
          criterion.points = ratings[0].points;
      }
      if (item.outcome !== false) {
          criterion.learning_outcome_id = item.outcome.id;
          if (typeof item.outcome.ignore !== 'undefined' && item.outcome.ignore) {
              criterion.ignore_for_scoring = 1;
          }
      }
      return criterion;
  }

  function addRatings(descriptions, points) {
      if (descriptions.length === 0 || descriptions.length !== points.length) {
          return false;
      }
      var mono = checkMonotonic(points, false);
      if (mono === 0) {
          return false;
      }
      var ratings = [];
      var n = points.length;
      var j;
      for (var i = 0; i < n; i++) {
          j = mono < 0 ? i : n - 1 - i;
          ratings.push({
              'description' : descriptions[j].replace(/\\n/g, ' ').replace(/\s+/g, ' '),
              'points' : points[j],
          });
      }
      return ratings;
  }

  function getCsrfToken() {
      var csrfRegex = new RegExp('^_csrf_token=(.*)$');
      var csrf;
      var cookies = document.cookie.split(';');
      for (var i = 0; i < cookies.length; i++) {
          var cookie = cookies[i].trim();
          var match = csrfRegex.exec(cookie);
          if (match) {
              csrf = decodeURIComponent(match[1]);
              break;
          }
      }
      return csrf;
  }

  function add_button() {
      var parent = document.querySelector('aside#right-side');
      if (parent) {
          var el = parent.querySelector('#jj_rubric_excel');
          if (!el) {
              el = document.createElement('a');
              el.classList.add('btn', 'button-sidebar-wide');
              el.id = 'jj_rubric_excel';
              var icon = document.createElement('i');
              icon.classList.add('icon-upload');
              el.appendChild(icon);
              var txt = document.createTextNode(' Importar Rubrica');
              el.appendChild(txt);
              el.addEventListener('click', showMainDialog);
              parent.appendChild(el);
          }
      }
  }

  // Mostrar diálogo principal con nuevo diseño
  function showMainDialog() {
      // Crear overlay
      const overlay = document.createElement('div');
      overlay.className = 'dialog-overlay';
      overlay.addEventListener('click', closeDialog);

      // Crear diálogo
      ui.dialog = document.createElement('div');
      ui.dialog.className = 'rubric-importer-dialog';
      ui.dialog.innerHTML = `
          <h2 class="dialog-title">Importar Rubrica</h2>
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

      document.body.appendChild(overlay);
      document.body.appendChild(ui.dialog);

      // Inicializar elementos UI
      ui.fileInput = document.getElementById('excelFileInput');
      ui.progressContainer = document.getElementById('progressContainer');
      ui.progressBar = document.getElementById('progressFill');
      ui.progressText = document.getElementById('progressText');
      ui.importBtn = document.getElementById('importBtn');
      ui.cancelBtn = document.getElementById('cancelBtn');
      ui.errorDisplay = document.getElementById('errorMessage');
      ui.browseBtn = document.getElementById('browseBtn');
      ui.fileNameDisplay = document.getElementById('fileNameDisplay');

      // Configurar eventos
      ui.browseBtn.addEventListener('click', () => ui.fileInput.click());
      ui.fileInput.addEventListener('change', handleFileSelect);
      ui.cancelBtn.addEventListener('click', closeDialog);
      ui.importBtn.addEventListener('click', processExcelFile);
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
      ui.importBtn.disabled = false;
  }

  // Validar archivo seleccionado
  function validateFile(file) {
      if (!file) return 'No file selected';
      if (file.size > 5 * 1024 * 1024) {
          return 'File is too large (max 5MB allowed)';
      }
      const fileExt = file.name.split('.').pop().toLowerCase();
      if (!['xlsx', 'xls'].includes(fileExt)) {
          return 'Invalid file type. Please use .xlsx or .xls files';
      }
      return null;
  }

  // Mostrar mensaje de error
  function showError(message) {
      ui.errorDisplay.textContent = message;
      ui.errorDisplay.style.display = 'block';
  }

  // Reiniciar UI
  function resetUI() {
      ui.progressContainer.style.display = 'none';
      ui.errorDisplay.style.display = 'none';
      ui.importBtn.disabled = true;
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
  }

  // Procesar archivo Excel (función original modificada para nuevo UI)
  function processExcelFile() {
      errors = [];
      if (!ui.fileInput.files || ui.fileInput.files.length === 0) {
          showError('Please select an Excel file.');
          return;
      }

      const file = ui.fileInput.files[0];
      const reader = new FileReader();

      // Configurar UI de progreso
      ui.importBtn.disabled = true;
      ui.cancelBtn.disabled = true;
      ui.fileInput.disabled = true;
      ui.browseBtn.disabled = true;
      ui.progressContainer.style.display = 'block';
      ui.progressBar.style.width = '0%';
      ui.progressText.textContent = 'Processing file...';

      reader.onload = function(e) {
          try {
              var data = new Uint8Array(e.target.result);
              var workbook = XLSX.read(data, { type: 'array' });

              var assocMatch = assocRegex.exec(window.location.pathname);
              if (!assocMatch) {
                  showError('Unable to determine where to place these rubrics.');
                  return;
              }

              var associationType = assocMatch[1].charAt(0).toUpperCase() + assocMatch[1].slice(1);
              var association = {
                  'type' : associationType,
                  'id' : assocMatch[2],
              };
              rubricAssociation = association;

              // Contar hojas válidas
              var validSheets = 0;
              workbook.SheetNames.forEach(function() { validSheets++; });

              var processedSheets = 0;

              // Procesar cada hoja
              workbook.SheetNames.forEach(function(sheetName) {
                  try {
                      var worksheet = workbook.Sheets[sheetName];
                      var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                      // Actualizar progreso
                      processedSheets++;
                      var progress = Math.round((processedSheets / validSheets) * 100);
                      ui.progressBar.style.width = `${progress}%`;
                      ui.progressText.textContent = `Processing sheet ${processedSheets} of ${validSheets}: ${sheetName}`;

                      // Saltar fila de cabecera
                      jsonData = jsonData.slice(1);

                      // Convertir a texto delimitado por tabulaciones para analizar
                      var tabText = jsonData.map(function(row) {
                          return row.join('\t');
                      }).join('\n');

                      // Analizar los datos de la rúbrica
                      var sheetCriteria = parseRubric(tabText);
                      if (sheetCriteria && sheetCriteria.length > 0) {
                          prepareRubric(sheetName, sheetCriteria, association);
                      }
                  } catch (e) {
                      errors.push('Error processing sheet "' + sheetName + '": ' + e.message);
                  }
              });

              if (errors.length === 0) {
                  ui.progressText.innerHTML = `
                      <div style="color: #4CAF50; font-weight: bold;">
                          <i class="icon-check" style="margin-right: 5px;"></i>
                          Successfully processed ${validSheets} sheets!
                      </div>
                  `;
                  setTimeout(() => {
                      closeDialog();
                      window.location.reload();
                  }, 2000);
              } else {
                  updateMsgs();
              }
          } catch (e) {
              showError('Error processing Excel file: ' + e.message);
          }
      };

      reader.onerror = function() {
          showError('Error reading the file.');
      };

      reader.readAsArrayBuffer(file);
  }

  // Resto de funciones originales (sin cambios)
  function parseRubric(txt) {
      var criteria = [];
      var lines = txt.split(/\r?\n/);
      for (var i = 0; i < lines.length; i++) {
          var line = lines[i].trim();
          if (line === '') continue;

          var parts = line.split(/\t/);
          if (parts.length < 4) {
              errors.push('Invalid format in line ' + (i + 1));
              continue;
          }

          // Extraer nombre del criterio (primera columna)
          var criterionName = parts[0];

          var criterion = {
              'description': criterionName,
              'ratings': []
          };

          // Procesar cada valoración (grupos de 3 columnas: puntos, título, descripción)
          for (var j = 1; j < parts.length; j += 3) {
              if (j + 2 >= parts.length) {
                  errors.push('Incomplete rating in line ' + (i + 1));
                  break;
              }

              var points = parseFloat(parts[j]);
              if (isNaN(points)) {
                  errors.push('Invalid points value in line ' + (i + 1));
                  continue;
              }

              var ratingTitle = parts[j + 1];
              var ratingDescription = parts[j + 2];

              criterion.ratings.push({
                  'description': ratingTitle,
                  'points': points,
                  'long_description': ratingDescription
              });
          }

          criteria.push(criterion);
      }

      return criteria;
  }

  function prepareRubric(title, criteria, association) {
      if (typeof criteria === 'object' && criteria.length > 0) {
          var formData = addRubric(title, criteria, association);
          if (typeof formData !== 'undefined') {
              saveRubric(formData);
          }
      }
  }

  function saveRubric(formData) {
      formData.authenticity_token = getCsrfToken();
      var url = window.location.pathname;
      $.ajax({
          'cache' : false,
          'url' : url,
          'type' : 'POST',
          'data' : formData,
      }).done(function() {
          // Éxito ya manejado en processExcelFile
      }).fail(function(jqXHR, textStatus, errorThrown) {
          errors.push('Error saving rubric: ' + textStatus);
          updateMsgs();
      });
  }

  function updateMsgs() {
      if (errors.length === 0) return;

      let errorHtml = '<ul>';
      for (var i = 0; i < errors.length; i++) {
          errorHtml += `<li>${errors[i]}</li>`;
      }
      errorHtml += '</ul>';

      ui.errorDisplay.innerHTML = errorHtml;
      ui.errorDisplay.style.display = 'block';
  }

  function isInteger(t) {
      if (typeof t === 'undefined') {
          return false;
      }
      return /^[0-9]+$/.test(t.trim());
  }

  function isBoolean(t, req) {
      if (typeof t === 'undefined') {
          return false;
      }
      if (typeof req !== 'undefined' && req && t === '') {
          return false;
      }
      return /^$|^[01]$/.test(t.trim());
  }

  function getBoolean(t) {
      return (t === false || t === null || t === '' || t === 0 || t === '0') ? false : true;
  }

  function isPoints(t) {
      if (typeof t === 'undefined') {
          return false;
      }
      return /^([0-9]+|[0-9]+[.][0-9]+|0?[.][0-9]+)$/.test(t);
  }
})();