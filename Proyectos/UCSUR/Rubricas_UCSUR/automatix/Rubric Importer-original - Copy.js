// ==UserScript==
// @name        Rubric Importer
// @namespace   https://github.com/jamesjonesmath/canvancement
// @description Create a rubric by copying from a spreadsheet and pasting into Canvas
// @include     https://*.instructure.com/courses/*/rubrics
// @include     https://*.instructure.com/accounts/*/rubrics
// @include     https://cientificavirtual.cientifica.edu.pe/courses/*/rubrics
// @version     5
// @grant       none
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
    var pointsPossible = criteria.reduce((sum, crit) => sum + crit.points, 0);
    var F = {
      'rubric': {
        'title': title,
        'points_possible': pointsPossible,
        'free_form_criterion_comments': false,
        'criteria': criteria
      },
      'rubric_association': {
        'id': '',
        'use_for_grading': false,
        'hide_score_total': false,
        'association_type': association.type,
        'association_id': association.id,
        'purpose': 'bookmark'
      },
      'title': title,
      'points_possible': pointsPossible,
      'rubric_id': 'new',
      'rubric_association_id': '',
      'skip_updating_points_possible': false
    };
    return F;
  }

  function isPoints(t) {
    return /^([0-9]+|[0-9]+[.][0-9]+|0?[.][0-9]+)$/.test(t);
  }

  function addCriterion(item) {
    if (typeof item === 'undefined') {
      return false;
    }
    var name = item.name.trim();
    var longDesc = typeof item.longDesc === 'undefined' ? '' : item.longDesc.trim();
    var ratings = addRatings(item.ratings, item.points, item.long_descriptions);

    var criterion = {
      'description': name,
      'long_description': longDesc,
      'ratings': ratings
    };

    if (ratings !== false) {
      criterion.points = ratings[0].points;
    }

    return criterion;
  }

  function addRatings(descriptions, points, long_descriptions) {
    if (descriptions.length === 0 || descriptions.length !== points.length || descriptions.length !== long_descriptions.length) {
      return false;
    }

    var ratings = [];
    for (var i = 0; i < descriptions.length; i++) {
      ratings.push({
        'description': descriptions[i].trim(),
        'points': Number(points[i]),
        'long_description': long_descriptions[i].trim()
      });
    }

    return ratings;
  }

  function parseRubricText(txt) {
    var lines = txt.trim().split(/\r?\n/);
    var criteria = [];

    for (var i = 0; i < lines.length; i++) {
      var parts = lines[i].split(/\t/);
      if (parts.length < 10) continue;  // Validación de estructura mínima

      var name = parts[0].trim();
      var ratings = [];
      var points = [];
      var long_descriptions = [];

      for (var j = 1; j < parts.length; j += 3) {
        if (j + 2 >= parts.length) break;
        points.push(parts[j].trim());
        ratings.push(parts[j + 1].trim());
        long_descriptions.push(parts[j + 2].trim());
      }

      criteria.push({
        'name': name,
        'ratings': ratings,
        'points': points,
        'long_descriptions': long_descriptions
      });
    }

    return criteria;
  }

  function getCsrfToken() {
    var csrfMatch = document.cookie.match(/_csrf_token=([^;]+)/);
    return csrfMatch ? decodeURIComponent(csrfMatch[1]) : '';
  }

  function add_button() {
    var parent = document.querySelector('aside#right-side');
    if (parent) {
      var el = parent.querySelector('#jj_rubric');
      if (!el) {
        el = document.createElement('a');
        el.classList.add('btn', 'button-sidebar-wide');
        el.id = 'jj_rubric';
        var icon = document.createElement('i');
        icon.classList.add('icon-import');
        el.appendChild(icon);
        var txt = document.createTextNode(' Import Rubric');
        el.appendChild(txt);
        el.addEventListener('click', openDialog);
        parent.appendChild(el);
      }
    }
  }

  function createDialog() {
    var el = document.querySelector('#jj_rubric_dialog');
    if (!el) {
      el = document.createElement('div');
      el.id = 'jj_rubric_dialog';
      el.classList.add('ic-Form-control');
      var label = document.createElement('label');
      label.htmlFor = 'jj_rubric_title';
      label.textContent = 'Rubric Title: ';
      label.classList.add('ic-Label');
      el.appendChild(label);
      var input = document.createElement('input');
      input.id = 'jj_rubric_title';
      input.classList.add('ic-Input');
      input.type = 'text';
      input.placeholder = 'Enter name of rubric';
      el.appendChild(input);
      label = document.createElement('label');
      label.htmlFor = 'jj_rubric_text';
      label.textContent = 'Rubric Contents';
      label.classList.add('ic-Label');
      el.appendChild(label);
      var textarea = document.createElement('textarea');
      textarea.id = 'jj_rubric_text';
      textarea.classList.add('ic-Input');
      textarea.placeholder = 'Paste a tab-delimited rubric into this textbox.';
      el.appendChild(textarea);
      var msg = document.createElement('div');
      msg.id = 'jj_rubric_msg';
      msg.classList.add('ic-flash-warning');
      msg.style.display = 'none';
      el.appendChild(msg);
      var parent = document.querySelector('body');
      parent.appendChild(el);
    }
  }

  function openDialog() {
    try {
      createDialog();
      $('#jj_rubric_dialog').dialog({
        'title' : 'Import Rubric',
        'autoOpen' : false,
        'buttons' : [ {
          'text' : 'Create',
          'click' : processDialog
        }, {
          'text' : 'Cancel',
          'click' : function() {
            $(this).dialog('close');
            var el = document.getElementById('jj_rubric_text');
            if (el) {
              el.value = '';
            }
            errors = [];
            updateMsgs();
          }
        } ],
        'modal' : true,
        'height' : 'auto',
        'width' : '80%'
      });
      if (!$('#jj_rubric_dialog').dialog('isOpen')) {
        $('#jj_rubric_dialog').dialog('open');
      }
    } catch (e) {
      console.log(e);
    }
  }

  function processDialog() {
    var el = document.getElementById('jj_rubric_text');
    if (!el.value.trim()) {
      alert('Debe ingresar una rúbrica en formato tabulado.');
      return;
    }

    var criteria = parseRubricText(el.value);
    if (criteria.length === 0) {
      alert('El formato del texto ingresado no es válido.');
      return;
    }

    var rubricTitle = document.getElementById('jj_rubric_title').value.trim();
    if (!rubricTitle) {
      alert('Debe ingresar un título para la rúbrica.');
      return;
    }

    var association = {
      'type': 'Course',
      'id': 'new'
    };

    var formData = {
      'rubric': {
        'title': rubricTitle,
        'points_possible': 4,
        'free_form_criterion_comments': 0,
        'criteria': criteria
      },
      'rubric_association': {
        'id': '',
        'use_for_grading': 0,
        'hide_score_total': 0,
        'association_type': association.type,
        'association_id': association.id,
        'purpose': 'bookmark'
      },
      'title': rubricTitle,
      'points_possible': 4,
      'rubric_id': 'new',
      'rubric_association_id': '',
      'skip_updating_points_possible': 0
    };

    saveRubric(formData);
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
      'cache': false,
      'url': url,
      'type': 'POST',
      'data': formData
    }).done(function() {
      alert('Rúbrica importada con éxito.');
      window.location.reload(true);
    }).fail(function() {
      alert('Error al guardar la rúbrica.');
    });
  }

  function updateMsgs() {
    var msg = document.getElementById('jj_rubric_msg');
    if (!msg) {
      return;
    }
    if (msg.hasChildNodes()) {
      msg.removeChild(msg.childNodes[0]);
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

  function blockRubric(txt) {
    var criteria = [];
    var isValid = true;
    var isMethod = false;
    try {
      var dequoted = dequote(txt);
      var lines = dequoted.split(/\r?\n/);
      var i = 0;
      while (isValid && i < lines.length) {
        var words = lines[i].replace(/[\s\uFEFF\xA0]+$/, '').split(/\t/);
        i++;
        words.map(function(s) {
          return s.trim();
        });
        if (words.length === 1 && words[0] === '') {
          continue; // Ignore blank lines
        }
        var name = '';
        var descriptions = [];
        var points = [];
        var titles = [];

        if (words.length >= 4) {
          name = words[0];
          var j = 1;
          while (j + 2 < words.length) {
            var point = words[j];
            var title = words[j + 1];
            var desc = words[j + 2];
            j += 3;

            if (!isPoints(point)) {
              errors.push(`Expected a number at line ${i}, column ${j - 2}`);
              isValid = false;
              break;
            }

            points.push(Number(point)); // Asegurar que sean numéricos
            titles.push(title);
            descriptions.push(desc);
          }
        } else {
          isValid = false;
          errors.push(`Invalid format in line ${i}`);
        }

        if (isValid) {
          var ratings = [];
          for (let k = 0; k < points.length; k++) {
            ratings.push({
              'description': descriptions[k],
              'points': points[k]
            });
          }

          var item = {
            'description': name,
            'long_description': descriptions[0],
            'ratings': ratings,
            'points': points[0] // Usar el puntaje más alto para Canvas
          };
          criteria.push(item);
        }
      }
    } catch (e) {
      console.log(e);
    }
    return isValid && criteria.length > 0 ? criteria : isMethod;
  }

  function isPoints(t) {
    return /^([0-9]+|[0-9]+[.][0-9]+|0?[.][0-9]+)$/.test(t);
  }

  function flexRubric(txt) {
    var criteria = [];
    var isValid = true;
    try {
      var dequoted = dequote(txt);
      var lines = dequoted.split(/\r?\n/);
      var i = 0;
      while (isValid && i < lines.length) {
        var words = lines[i].trim().split(/\t/);
        i++;
        if (words.length === 1 && words[0] === '') {
          // Ignore blank lines
          continue;
        }
        words.map(function(s) {
          return s.trim();
        });
        var outcome = false;
        var name = '';
        var longDesc = '';
        var descriptions = [];
        var points = [];
        var validLine = true;
        var k = 0;
        if (isInteger(words[k])) {
          if (words.length > k + 1 && isBoolean(words[k + 1], true)) {
            outcome = {
              'id' : words[k++],
              'ignore' : getBoolean(words[k++]) ? false : true
            };
          } else {
            outcome = {
              'id' : words[k++]
            };
          }
          while (words.length > k && words[k] === '') {
            k++;
          }
        }
        if (words.length > k) {

          if (outcome && words.length > k + 1 && !isPoints(words[k]) && isPoints(words[k + 1])) {
          } else {
            name = words[k++];
          }
        }
        if (!outcome && words.length > k) {
          if (isInteger(words[k])) {
            // What should be the long description or a rating is an integer
            if (words.length > k + 1 && isBoolean(words[k + 1], true)) {
              outcome = {
                'id' : words[k++],
                'ignore' : getBoolean(words[k++]) ? false : true
              };
            } else {
              outcome = {
                'id' : words[k++]
              };
            }
          }
        }
        if (isValid && words.length > k) {
          if (!isPoints(words[k])) {
            if (words.length == k) {
              longDesc = words[k++];
            } else if (words.length > k + 1 && !isPoints(words[k + 1])) {
              longDesc = words[k++];
            }
          }
        }
        while (isValid && validLine && words.length > k) {
          var description;
          var point;
          if (words.length > k + 1) {
            description = words[k++];
            point = words[k++];
            if (description === '' && point === '') {
              // Skip completely blank pairs
              continue;
            }
            if (!outcome && isInteger(description) && isBoolean(point)) {
              // Have an integer in the rating descriptions
              // Valid if this is the last rating, otherwise invalid
              if (words.length == k) {
                outcome = {
                  'id' : description,
                  'ignore' : getBoolean(point) ? false : true
                };
              }
            } else {
              if (description === '') {
                validLine = false;
                if (i > 1) {
                  errors.push('Empty rating description in line ' + i + ', column ' + (k - 1));
                  isValid = false;
                }
              }
              if (!isPoints(point)) {
                validLine = false;
                if (i > 1) {
                  errors.push('Second item in pair is not a number in line ' + i + ', column ' + k);
                  isValid = false;
                }
              }
              if (isValid && validLine) {
                descriptions.push(description);
                points.push(point);
              }
            }
          } else {
            // Odd number of entries on line, check for outcome
            // or ignore points for outcome
            description = words[k++];
            if (!outcome) {
              if (isInteger(description)) {
                outcome = {
                  'id' : description
                };
              } else {
                errors.push('Unbalanced rating/points at the end of line ' + i);
                isValid = false;
              }
            } else {
              if (isBoolean(description)) {
                var thisIgnore = getBoolean(description) ? false : true;
                if (typeof outcome.ignore === 'undefined') {
                  outcome.ignore = thisIgnore;
                } else if (outcome.ignore !== thisIgnore) {
                  errors.push('Conflicting directives for using outcome for scoring in line ' + i);
                  isValid = false;
                }
              }
            }
          }
        }
        if (isValid && validLine) {
          var item = {
            'name' : name,
            'longDesc' : longDesc,
            'outcome' : outcome,
            'ratings' : descriptions,
            'points' : points,
          };
          var criterion = addCriterion(item);
          if (criterion !== false) {
            criteria.push(criterion);
          } else {
            errors.push('Unable to create criterion for line ' + i);
            isValid = false;
          }
        }
      }
    } catch (e) {
      console.log(e);
    }
    return isValid ? criteria : false;
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

  function dequote(txt) {
    var regex = new RegExp('([^"]*\\\\n)?["](.*)["](\\\\n[^*"])?$');
    var s = txt.replace(/\r?\n/g, '\\n');
    s = s.replace(/([^\t]+)/g, quoted);
    return s;

    function quoted(s) {
      s = s.trim();
      if (regex.test(s)) {
        var match = regex.exec(s);
        if (match) {
          var prefix = typeof match[1] !== 'undefined' ? match[1].replace(/\\n/g, '\n') : '';
          var postfix = typeof match[3] !== 'undefined' ? match[3].replace(/\\n/g, '\n') : '';
          s = prefix + match[2].trim().replace(/""/g, '"') + postfix;
        }
      } else {
        s = s.replace(/\\n/g, '\n');
      }
      return s;
    }
  }

  function checkOutcomes() {
    if (criteria === false) {
      return;
    }
    var i;
    var id;
    outcomes = [];
    pendingOutcomes = 0;
    var outcomeIds = [];
    // Go through list of outcomes check for duplicates
    for (i = 0; i < criteria.length; i++) {
      if (typeof criteria[i].learning_outcome_id !== 'undefined' && isInteger(criteria[i].learning_outcome_id)) {
        id = criteria[i].learning_outcome_id;
        if (outcomeIds.indexOf(id) === -1) {
          outcomes.push({
            'id' : id,
            'loaded' : false
          });
          outcomeIds.push(id);
        }
      }
    }
    if (outcomes.length > 0) {
      pendingOutcomes = outcomes.length;
      for (i = 0; i < outcomes.length; i++) {
        id = outcomes[i].id;
        var url = '/api/v1/outcomes/' + id;
        $.ajax({
          'url' : url,
          'dataType' : 'json',
          'timeout' : 3000
        }).done(replaceCriteria).fail(checkPendingOutcomes);
      }
    } else {
      checkPendingOutcomes();
    }
  }

  function checkPendingOutcomes() {
    pendingOutcomes--;
    if (pendingOutcomes <= 0) {
      // we're all finished with the fetching of the rubric
      var outstanding = [];
      if (outcomes.length > 0) {
        for (var i = 0; i < outcomes.length; i++) {
          if (outcomes[i].loaded === false) {
            outstanding.push(outcomes[i].id);
          }
        }
      }
      if (outstanding.length > 0) {
        errors.push('Unable to locate outcomes: ' + outstanding.join(', '));
        updateMsgs();
      } else {
        prepareRubric(rubricTitle, criteria, rubricAssociation);
      }
    }
  }

  function replaceCriteria(data) {
    if (criteria === false || typeof data === 'undefined') {
      return;
    }
    for (var i = 0; i < criteria.length; i++) {
      if (criteria[i].learning_outcome_id == data.id) {
        if (criteria[i].description === '') {
          criteria[i].description = data.title;
        }
        if (criteria[i].long_description === '' && typeof data.description !== 'undefined') {
          criteria[i].long_description = data.description;
        }
        if (typeof data.mastery_points !== 'undefined') {
          criteria[i].mastery_points = data.mastery_points;
        }
        if (typeof data.ratings !== 'undefined') {
          criteria[i].ratings = data.ratings;
          criteria[i].points = data.ratings[0].points;
        }
      }
    }
    for (var j = 0; j < outcomes.length; j++) {
      if (outcomes[j].id == data.id) {
        outcomes[j].loaded = true;
      }
    }
    checkPendingOutcomes();
  }

})();
