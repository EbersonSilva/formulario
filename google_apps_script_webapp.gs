// Apps Script: WebApp para receber submissões JSON e gravar no Google Sheets
// Como usar: crie uma Google Sheet, abra Extensões → Apps Script, cole este código e publique como WebApp (Anyone, even anonymous) para aceitar POSTs.

function doPost(e) {
  try {
    var ssId = PropertiesService.getScriptProperties().getProperty('TARGET_SHEET_ID');
    if (!ssId) {
      return jsonResponse({ error: 'TARGET_SHEET_ID não configurado nas Script Properties' });
    }

    var sheet = SpreadsheetApp.openById(ssId).getSheetByName('Responses');
    if (!sheet) {
      // cria sheet e cabeçalho
      sheet = SpreadsheetApp.openById(ssId).insertSheet('Responses');
      sheet.appendRow(['timestamp','name','score','percentage','answers']);
    }

    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }

    var timestamp = data.timestamp || (new Date()).toISOString();
    var name = data.name || '';
    var score = data.score || '';
    var percentage = data.percentage || '';
    var answers = '';
    if (data.answers) {
      try { answers = JSON.stringify(data.answers); } catch (err) { answers = String(data.answers); }
    }

    sheet.appendRow([timestamp, name, score, percentage, answers]);

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ error: String(err) });
  }
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// Número de perguntas esperadas (ajuste se mudar o quiz)
const NUM_QUESTIONS = 10;

// Letras das respostas corretas, na ordem (a,b,c,d -> 'a','b'...)
const CORRECT_ANSWERS = ['b','c','b','c','c','b','c','b','b','b'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('KPI')
    .addItem('Criar/Atualizar Dashboard', 'createDashboard')
    .addItem('Parsear Responses', 'parseResponses')
    .addToUi();
}

function parseResponses() {
  var ssId = PropertiesService.getScriptProperties().getProperty('TARGET_SHEET_ID');
  if (!ssId) throw new Error('TARGET_SHEET_ID não configurado nas Script Properties');

  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName('Responses');
  if (!sheet) throw new Error('Aba Responses não encontrada');

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues(); // includes header
  if (values.length <= 1) return; // nothing to parse

  // Prepare headers: ensure we have Q1..Qn headers after 'answers'
  var header = values[0];
  var baseHeader = ['timestamp','name','score','percentage','answers'];
  // If header doesn't match, try to reset first 5 columns
  for (var i=0;i<5;i++) header[i] = baseHeader[i];

  for (var q=1;q<=NUM_QUESTIONS;q++) {
    header[5 + (q-1)] = 'Q' + q;
    header[5 + NUM_QUESTIONS + (q-1)] = 'Q' + q + '_ok';
  }
  // write header back
  sheet.getRange(1,1,1, header.length).setValues([header]);

  // Iterate rows and fill Q columns and Q_ok columns
  var out = [];
  for (var r=1;r<values.length;r++) {
    var row = values[r];
    var answersCell = row[4];
    var parsed = [];
    if (answersCell) {
      try {
        parsed = JSON.parse(answersCell);
      } catch (err) {
        // fallback: remove brackets/quotes and split
        var s = String(answersCell).replace(/[\[\]"]+/g,'');
        parsed = s === '' ? [] : s.split(',').map(function(x){return x.trim();});
      }
    }

    // prepare output row slice for the columns after col E
    var outRow = [];
    for (var q=0;q<NUM_QUESTIONS;q++) {
      outRow.push(parsed[q] || '');
    }
    for (var q=0;q<NUM_QUESTIONS;q++) {
      var ok = (parsed[q] && parsed[q].toString().toLowerCase() === CORRECT_ANSWERS[q]) ? 1 : 0;
      outRow.push(ok);
    }
    out.push(outRow);
  }

  // write parsed data starting at row 2, column 6 (F)
  if (out.length>0) {
    sheet.getRange(2,6,out.length, out[0].length).setValues(out);
  }
}

function createDashboard() {
  var ssId = PropertiesService.getScriptProperties().getProperty('TARGET_SHEET_ID');
  if (!ssId) throw new Error('TARGET_SHEET_ID não configurado nas Script Properties');

  var ss = SpreadsheetApp.openById(ssId);
  var responses = ss.getSheetByName('Responses');
  if (!responses) throw new Error('Aba Responses não encontrada');

  // Primeiro parse
  parseResponses();

  var lastRow = responses.getLastRow();
  var total = Math.max(0, lastRow - 1);

  var percentages = [];
  if (total>0) {
    percentages = responses.getRange(2,4,total,1).getValues().map(function(r){return Number(r[0]) || 0;});
  }

  var avg = 0;
  if (percentages.length>0) avg = percentages.reduce(function(a,b){return a+b;},0)/percentages.length;
  var passCount = percentages.filter(function(p){return p>=80;}).length;
  var passRate = total>0 ? (passCount/total)*100 : 0;

  // per-question percent using Q_ok columns starting at col (6 + NUM_QUESTIONS)
  var perQuestion = [];
  if (total>0) {
    var okRange = responses.getRange(2, 6 + NUM_QUESTIONS, total, NUM_QUESTIONS).getValues();
    for (var q=0;q<NUM_QUESTIONS;q++) {
      var sum = 0;
      for (var r=0;r<okRange.length;r++) sum += Number(okRange[r][q])||0;
      perQuestion.push(total>0 ? (sum/total)*100 : 0);
    }
  } else {
    for (var q=0;q<NUM_QUESTIONS;q++) perQuestion.push(0);
  }

  // distribution buckets
  var buckets = {"0-20":0,"21-40":0,"41-60":0,"61-80":0,"81-100":0};
  percentages.forEach(function(p){
    if (p<=20) buckets['0-20']++;
    else if (p<=40) buckets['21-40']++;
    else if (p<=60) buckets['41-60']++;
    else if (p<=80) buckets['61-80']++;
    else buckets['81-100']++;
  });

  // Create or clear Dashboard
  var dash = ss.getSheetByName('Dashboard');
  if (!dash) dash = ss.insertSheet('Dashboard');
  else dash.clear();

  // Write KPIs
  dash.getRange(1,1).setValue('Total participantes');
  dash.getRange(1,2).setValue(total);
  dash.getRange(2,1).setValue('Média (%)');
  dash.getRange(2,2).setValue(Math.round(avg*10)/10);
  dash.getRange(3,1).setValue('Taxa >=80%');
  dash.getRange(3,2).setValue(Math.round(passRate*10)/10);

  // Per-question table
  dash.getRange(5,1).setValue('Pergunta');
  dash.getRange(5,2).setValue('% Acerto');
  for (var q=0;q<NUM_QUESTIONS;q++) {
    dash.getRange(6+q,1).setValue('Q' + (q+1));
    dash.getRange(6+q,2).setValue(Math.round(perQuestion[q]*10)/10);
  }

  // Distribution
  var distStart = 6 + NUM_QUESTIONS + 2;
  dash.getRange(distStart,1).setValue('Faixa');
  dash.getRange(distStart,2).setValue('Contagem');
  var dkeys = Object.keys(buckets);
  for (var i=0;i<dkeys.length;i++) {
    dash.getRange(distStart+1+i,1).setValue(dkeys[i]);
    dash.getRange(distStart+1+i,2).setValue(buckets[dkeys[i]]);
  }

  // Gráfico de barras: % acerto por pergunta
  var chartBuilder = dash.newChart();
  chartBuilder.addRange(dash.getRange(5,1,NUM_QUESTIONS+1,2));
  chartBuilder.setChartType(Charts.ChartType.BAR);
  chartBuilder.setPosition(1,4,0,0);
  dash.insertChart(chartBuilder.build());

  // Gráfico de pizza: distribuição
  var chart2 = dash.newChart();
  chart2.addRange(dash.getRange(distStart,1, dkeys.length+1,2));
  chart2.setChartType(Charts.ChartType.PIE);
  chart2.setPosition(12,4,0,0);
  dash.insertChart(chart2.build());
}
