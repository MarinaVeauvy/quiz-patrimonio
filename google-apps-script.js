var SHEET_NAME = 'Leads';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'add';

    if (action === 'add') {
      return addLead(data);
    }

    if (action === 'mark_sent') {
      return markSent(data);
    }

    return jsonResponse({ error: 'Action invalida' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) || 'pending';

    if (action === 'pending') {
      return getPending();
    }

    if (action === 'all') {
      return getAll();
    }

    return jsonResponse({ error: 'Action invalida' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function addLead(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ error: 'Sheet nao encontrada' });

  var now = new Date();
  var msg2Date = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000);
  var msg3Date = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  sheet.appendRow([
    Utilities.formatDate(now, 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm'),
    data.nome || '',
    data.email || '',
    data.whatsapp || '',
    data.perfil || '',
    'SIM',
    Utilities.formatDate(msg2Date, 'America/Sao_Paulo', 'yyyy-MM-dd'),
    'NAO',
    Utilities.formatDate(msg3Date, 'America/Sao_Paulo', 'yyyy-MM-dd'),
    'NAO',
    'NAO',
    'NAO',
    ''
  ]);

  return jsonResponse({ success: true });
}

function getPending() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ error: 'Sheet nao encontrada' });

  var today = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  var data = sheet.getDataRange().getValues();
  var pending = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var nome = row[1];
    var telefone = String(row[3]);
    var perfil = row[4];
    var msg2Date = String(row[6]).substring(0, 10);
    var msg2Sent = String(row[7]).toUpperCase();
    var msg3Date = String(row[8]).substring(0, 10);
    var msg3Sent = String(row[9]).toUpperCase();

    if (msg2Date <= today && msg2Sent === 'NAO') {
      pending.push({
        row: i + 1,
        nome: nome,
        telefone: telefone,
        perfil: perfil,
        msg: 'msg2',
        col: 'H'
      });
    }

    if (msg3Date <= today && msg3Sent === 'NAO') {
      pending.push({
        row: i + 1,
        nome: nome,
        telefone: telefone,
        perfil: perfil,
        msg: 'msg3',
        col: 'J'
      });
    }
  }

  return jsonResponse({ success: true, today: today, pending: pending, count: pending.length });
}

function getAll() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ error: 'Sheet nao encontrada' });

  var data = sheet.getDataRange().getValues();
  var leads = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    leads.push({
      data: row[0],
      nome: row[1],
      email: row[2],
      telefone: String(row[3]),
      perfil: row[4],
      msg1: row[5],
      msg2_data: row[6],
      msg2_enviada: row[7],
      msg3_data: row[8],
      msg3_enviada: row[9],
      respondeu: row[10],
      agendou: row[11],
      obs: row[12]
    });
  }

  return jsonResponse({ success: true, leads: leads, total: leads.length });
}

function markSent(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ error: 'Sheet nao encontrada' });

  var row = data.row;
  var col = data.col;

  if (!row || !col) return jsonResponse({ error: 'row e col obrigatorios' });

  var colMap = { 'H': 8, 'J': 10 };
  var colNum = colMap[col.toUpperCase()];

  if (!colNum) return jsonResponse({ error: 'col invalida' });

  sheet.getRange(row, colNum).setValue('SIM');

  return jsonResponse({ success: true, row: row, col: col });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
