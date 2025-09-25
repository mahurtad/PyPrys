// ==UserScript==
// @name        Rubric Importer (Modified)
// @namespace   https://github.com/jamesjonesmath/canvancement
// @description Create a rubric by copying from a spreadsheet and pasting into Canvas
// @include     https://*.instructure.com/courses/*/rubrics
// @include     https://*.instructure.com/accounts/*/rubrics
// @include     https://cientificavirtual.cientifica.edu.pe/courses/*/rubrics
// @version     6
// @grant       none
// ==/UserScript==

(function() {
    'use strict';
    var assocRegex = new RegExp('^/(course|account)s/([0-9]+)/rubrics$');
    var errors = [];
    var criteria = [];
    var rubricTitle;
    var rubricAssociation;
  
    if (assocRegex.test(window.location.pathname)) {
      add_button();
    }
  
    function formatText(text, type) {
      if (type === 'bold') {
        return '<b>' + text + '</b>';
      }
      return text; // No format for descriptions
    }
  
    function addCriterion(item) {
      if (typeof item === 'undefined') {
        return false;
      }
      var name = formatText(item.name, 'bold');
      var longDesc = typeof item.longDesc === 'undefined' ? '' : item.longDesc;
      var ratings = addRatings(item.ratings, item.points);
      var criterion = {
        'description': name,
        'long_description': longDesc
      };
      if (ratings !== false) {
        criterion.ratings = ratings;
        criterion.points = ratings[0].points;
      }
      return criterion;
    }
  
    function addRatings(descriptions, points) {
      if (descriptions.length === 0 || descriptions.length !== points.length) {
        return false;
      }
      var ratings = [];
      for (var i = 0; i < points.length; i++) {
        ratings.push({
          'description': descriptions[i],
          'points': formatText(points[i], 'bold')
        });
      }
      return ratings;
    }
  
    function add_button() {
      var parent = document.querySelector('aside#right-side');
      if (parent) {
        var el = document.createElement('a');
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
  
    function openDialog() {
      try {
        createDialog();
        $('#jj_rubric_dialog').dialog({
          'title': 'Import Rubric',
          'autoOpen': false,
          'buttons': [
            {
              'text': 'Create',
              'click': processDialog
            },
            {
              'text': 'Cancel',
              'click': function() {
                $(this).dialog('close');
              }
            }
          ],
          'modal': true,
          'width': '80%'
        }).dialog('open');
      } catch (e) {
        console.log(e);
      }
    }
  
    function processDialog() {
      errors = [];
      var title, txt;
      var el = document.getElementById('jj_rubric_title');
      if (el.value && el.value.trim() !== '') {
        title = formatText(el.value, 'bold');
        rubricTitle = title;
      } else {
        errors.push('You must provide a title for your rubric.');
      }
      el = document.getElementById('jj_rubric_text');
      if (el.value && el.value.trim() !== '') {
        txt = el.value;
      } else {
        errors.push('You must paste your rubric into the textbox.');
      }
      if (errors.length === 0) {
        criteria = parseRubric(txt);
        if (criteria.length > 0) {
          saveRubric(title, criteria);
        }
      }
    }
  
    function parseRubric(txt) {
      var lines = txt.split('\n');
      var criteria = [];
      lines.forEach(function(line) {
        var parts = line.split('\t');
        if (parts.length >= 3) {
          criteria.push({
            'name': formatText(parts[1], 'bold'),
            'longDesc': parts[2],
            'ratings': parts.slice(3, parts.length - 1),
            'points': formatText(parts[0], 'bold')
          });
        }
      });
      return criteria;
    }
  
    function saveRubric(title, criteria) {
      console.log('Saving rubric:', title, criteria);
    }
  
  })();
  