// ==UserScript==
// @name        Rubric Importer from Excel por Curso
// @namespace   https://github.com/jamesjonesmath/canvancement
// @description Create rubrics by importing from an Excel file (one rubric per sheet)
// @include     https://*.instructure.com/courses/*/rubrics
// @include     https://*.instructure.com/accounts/*/rubrics
// @include     https://cientificavirtual.cientifica.edu.pe/courses/*/rubrics
// @version     1.0
// @grant       GM_xmlhttpRequest
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

  if (assocRegex.test(window.location.pathname)) {
    add_button();
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
        icon.classList.add('icon-import');
        el.appendChild(icon);
        var txt = document.createTextNode(' Import Rubrics from Excel');
        el.appendChild(txt);
        el.addEventListener('click', openExcelDialog);
        parent.appendChild(el);
      }
    }
  }

  function createExcelDialog() {
    var el = document.querySelector('#jj_rubric_excel_dialog');
    if (!el) {
      el = document.createElement('div');
      el.id = 'jj_rubric_excel_dialog';
      el.classList.add('ic-Form-control');

      var label = document.createElement('label');
      label.htmlFor = 'jj_rubric_excel_file';
      label.textContent = 'Excel File: ';
      label.classList.add('ic-Label');
      el.appendChild(label);

      var input = document.createElement('input');
      input.id = 'jj_rubric_excel_file';
      input.classList.add('ic-Input');
      input.type = 'file';
      input.accept = '.xlsx,.xls';
      el.appendChild(input);

      var msg = document.createElement('div');
      msg.id = 'jj_rubric_excel_msg';
      msg.classList.add('ic-flash-warning');
      msg.style.display = 'none';
      el.appendChild(msg);

      var parent = document.querySelector('body');
      parent.appendChild(el);
    }
  }

  function openExcelDialog() {
    try {
      createExcelDialog();
      $('#jj_rubric_excel_dialog').dialog({
        'title' : 'Import Rubrics from Excel',
        'autoOpen' : false,
        'buttons' : [ {
          'text' : 'Import',
          'click' : processExcelFile
        }, {
          'text' : 'Cancel',
          'click' : function() {
            $(this).dialog('close');
            var el = document.getElementById('jj_rubric_excel_file');
            if (el) {
              el.value = '';
            }
            errors = [];
            updateMsgs();
          }
        } ],
        'modal' : true,
        'height' : 'auto',
        'width' : '50%'
      });
      if (!$('#jj_rubric_excel_dialog').dialog('isOpen')) {
        $('#jj_rubric_excel_dialog').dialog('open');
      }
    } catch (e) {
      console.log(e);
    }
  }

  function processExcelFile() {
    errors = [];
    var fileInput = document.getElementById('jj_rubric_excel_file');
    if (!fileInput.files || fileInput.files.length === 0) {
      errors.push('Please select an Excel file.');
      updateMsgs();
      return;
    }

    var file = fileInput.files[0];
    var reader = new FileReader();

    reader.onload = function(e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });

        var assocMatch = assocRegex.exec(window.location.pathname);
        if (!assocMatch) {
          errors.push('Unable to determine where to place these rubrics.');
          updateMsgs();
          return;
        }

        var associationType = assocMatch[1].charAt(0).toUpperCase() + assocMatch[1].slice(1);
        var association = {
          'type' : associationType,
          'id' : assocMatch[2],
        };
        rubricAssociation = association;

        // Process each sheet
        workbook.SheetNames.forEach(function(sheetName) {
          var worksheet = workbook.Sheets[sheetName];
          var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          // Skip header row (first row)
          jsonData = jsonData.slice(1);

          // Convert to tab-delimited text for parsing
          var tabText = jsonData.map(function(row) {
            return row.join('\t');
          }).join('\n');

          // Parse the rubric data
          var sheetCriteria = parseRubric(tabText);
          if (sheetCriteria && sheetCriteria.length > 0) {
            prepareRubric(sheetName, sheetCriteria, association);
          }
        });

        if (errors.length === 0) {
          $('#jj_rubric_excel_dialog').dialog('close');
          window.location.reload(true);
        } else {
          updateMsgs();
        }
      } catch (e) {
        errors.push('Error processing Excel file: ' + e.message);
        updateMsgs();
      }
    };

    reader.onerror = function() {
      errors.push('Error reading the file.');
      updateMsgs();
    };

    reader.readAsArrayBuffer(file);
  }

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

      // Extract criterion name (first column)
      var criterionName = parts[0];

      var criterion = {
        'description': criterionName,
        'ratings': []
      };

      // Process each rating (groups of 3 columns: points, title, description)
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
      updateMsgs();
    }).fail(function() {
      errors.push('All the information was supplied correctly, but there was an error saving rubric to Canvas.');
      updateMsgs();
    });
  }

  function updateMsgs() {
    var msg = document.getElementById('jj_rubric_excel_msg');
    if (!msg) {
      return;
    }
    if (msg.hasChildNodes()) {
      while (msg.firstChild) {
        msg.removeChild(msg.firstChild);
      }
    }
    if (typeof errors === 'undefined' || errors.length === 0) {
      msg.style.display = 'none';
    } else {
      var ul = document.createElement('ul');
      var li;
      for (var i = 0; i < errors.length; i++) {
        li = document.createElement('li');
        li.textContent = errors[i];
        ul.appendChild(li);
      }
      msg.appendChild(ul);
      msg.style.display = 'inline-block';
    }
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