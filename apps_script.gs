// Backend de estado compartilhado do funil.html, rodando como Web App do Google Apps Script
// vinculado a uma planilha do Google.
//
// Setup (passo-a-passo no README do funil ou na conversa que gerou este arquivo):
//   1. Criar uma planilha do Google em branco (sheets.new). Renomear pra "Funil State" (ou o que preferir).
//   2. Menu Extensões > Apps Script. Apagar o código padrão e colar TUDO deste arquivo.
//   3. Salvar (ícone de disquete).
//   4. Implantar > Nova implantação > Tipo: "Aplicativo da Web".
//        - Descrição: "Funil shared state"
//        - Executar como: Eu (talentos@allugator.com)
//        - Quem tem acesso: Qualquer pessoa em allugator.com
//   5. Implantar. Autorizar os escopos quando pedir.
//   6. Copiar a URL do tipo https://script.google.com/macros/s/AKfy.../exec
//      e colar no funil.html na constante SHEETS_API_URL.

const SHEET_TAB = 'state';

function _sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_TAB);
  if (!sh) {
    sh = ss.insertSheet(SHEET_TAB);
    sh.appendRow(['section', 'key', 'value', 'updated_at', 'updated_by']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _readAll_() {
  const sh = _sheet_();
  const rows = sh.getDataRange().getValues();
  const out = { status: {}, moves: {}, completed: {} };
  for (let i = 1; i < rows.length; i++) {
    const [section, key, value] = rows[i];
    if (!section || !key) continue;
    if (out[section]) out[section][String(key)] = String(value);
  }
  return out;
}

function _json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return _json_(_readAll_());
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return _json_({ ok: false, error: 'bad_json' }); }

  const user = (Session.getActiveUser().getEmail() || '').toString();
  const now = new Date().toISOString();

  if (body.action === 'bulk') return _handleBulk_(body.state, user, now);

  const { section, key, value } = body;
  if (!['status', 'moves', 'completed'].includes(section) || !key) {
    return _json_({ ok: false, error: 'invalid_payload' });
  }

  const sh = _sheet_();
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === section && String(data[i][1]) === String(key)) {
      rowIdx = i + 1; // Sheet rows são 1-indexed
      break;
    }
  }

  if (value == null || value === '') {
    if (rowIdx > 0) sh.deleteRow(rowIdx);
  } else if (rowIdx > 0) {
    sh.getRange(rowIdx, 3, 1, 3).setValues([[String(value), now, user]]);
  } else {
    sh.appendRow([section, String(key), String(value), now, user]);
  }

  return _json_({ ok: true });
}

function _handleBulk_(state, user, now) {
  if (!state || typeof state !== 'object') return _json_({ ok: false, error: 'no_state' });
  const sh = _sheet_();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, Math.max(5, sh.getLastColumn())).clearContent();

  const rows = [];
  for (const section of ['status', 'moves', 'completed']) {
    const sec = state[section] || {};
    for (const k of Object.keys(sec)) {
      const v = sec[k];
      if (v == null || v === '') continue;
      rows.push([section, String(k), String(v), now, user]);
    }
  }
  if (rows.length > 0) sh.getRange(2, 1, rows.length, 5).setValues(rows);
  return _json_({ ok: true, written: rows.length });
}
