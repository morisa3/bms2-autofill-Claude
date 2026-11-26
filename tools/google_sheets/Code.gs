/**
 * 公路橋梁填報助手 v4.0 — Google 試算表直寫端點（Google Apps Script）
 *
 * 使用方式：
 * 1. 開啟《115年北雲林自檢橋梁總表(請每日更新)》→ 上方選單「擴充功能」→「Apps Script」
 * 2. 將本檔全部程式碼貼入 Code.gs，把下方 SECRET_TOKEN 改成你們團隊自訂的密碼字串
 * 3. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    - 執行身分：「我」
 *    - 誰可以存取：「所有人」
 * 4. 複製部署後的「網頁應用程式網址」與你設定的 SECRET_TOKEN，貼進填報助手「⑥ 同步到橋梁總表 → 雲端設定」即可。
 */

const SECRET_TOKEN = '請改成你的自訂金鑰';
const GUIDE_SHEET = '說明';
const GENERATED_TAG = '填報助手產生';
const COLOR_CHECK = '#f8cbad';    // 與系統資料不同／本期淨高異動：橘底
const COLOR_DONE = '#b7e1cd';     // 系統填寫完成：橋名綠底
const COLOR_SERIAL = '#92d050';   // 前期已建檔：序號綠底

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const req = JSON.parse(raw);
    if (SECRET_TOKEN && req.token !== SECRET_TOKEN) {
      return jsonOut({ ok: false, error: '金鑰驗證失敗，請確認填報助手與 Apps Script 的金鑰一致' });
    }
    const action = String(req.action || '');
    if (action === 'ping') return jsonOut(handlePing());
    if (action === 'list_bridges') return jsonOut(handleListBridges(req.sheet || ''));
    if (action === 'sync_bridge') return jsonOut(handleSyncBridge(req));
    if (action === 'export_all') return jsonOut(handleExportAll());
    return jsonOut({ ok: false, error: '不支援的操作：' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  return jsonOut({ ok: true, message: '公路橋梁填報助手 Apps Script 運作中（請使用 POST 呼叫）' });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 與 bridgefill/core/bridge_list.py 的 key() 完全一致的表頭正規化 */
function normKey(text) {
  let s = String(text == null ? '' : text).normalize('NFKC');
  s = s.replace(/\s+/g, '');
  s = s.replace(/\(?\d{2,3}年\)?/g, '');
  s = s.replace(/須修正/g, '').replace(/檢測新增/g, '新增').replace(/≦/g, '<=').replace(/≧/g, '>=');
  s = s.replace(/\(列出過去維修紀錄\)|\(橋長\*淨寬\)|\(V\)|\(是[\/／、]?否\)|\(ch\d+(?:\.\d+)?\)/gi, '');
  s = s.replace(/[\/、]/g, '');
  return s.replace(/梁下橋下淨高/g, '梁下淨高');
}

/** 計算與 Python hashlib.sha1(...).hexdigest()[:8] 一致的內容識別碼 */
function contentStamp(value) {
  const text = String(value == null ? '' : value).normalize('NFKC').trim();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, text, Utilities.Charset.UTF_8);
  return digest.slice(0, 4).map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function generatedState(val, note) {
  const m = String(note || '').match(/填報助手產生[\s\S]*?識別碼\s*([0-9a-f]{8})/);
  if (!m) return null;
  return contentStamp(val) === m[1] ? 'same' : 'edited';
}

function sameValue(a, b) {
  const sa = String(a == null ? '' : a).normalize('NFKC').trim();
  const sb = String(b == null ? '' : b).normalize('NFKC').trim();
  if (sa !== '' && sb !== '' && !isNaN(Number(sa)) && !isNaN(Number(sb))) {
    return Number(sa) === Number(sb);
  }
  return sa.replace(/\s+/g, '') === sb.replace(/\s+/g, '');
}

function toValue(v) {
  const s = String(v == null ? '' : v);
  if (/^-?\d+(?:\.\d+)?$/.test(s)) return Number(s);
  return s;
}

function sheetHeadersMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  for (let i = 0; i < row.length; i++) {
    const k = normKey(row[i]);
    if (k && !map[k]) map[k] = i + 1;
  }
  return map;
}

function handlePing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ss.getSheets().map(s => s.getName()).filter(n => n !== GUIDE_SHEET);
  return { ok: true, title: ss.getName(), sheets: names };
}

function handleListBridges(onlySheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName() !== GUIDE_SHEET && (!onlySheet || s.getName() === onlySheet));
  const bridges = [];
  for (const sheet of sheets) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    const heads = sheetHeadersMap(sheet);
    const idCol = heads[normKey('二代系統ID')] || 0;
    const nameCol = heads[normKey('橋梁名稱')] || 0;
    if (!idCol && !nameCol) continue;
    const serialCol = heads[normKey('序號')] || 0;
    const codeCol = heads[normKey('橋梁編號')] || 0;
    const doneCol = heads[normKey('系統填寫完成')] || 0;
    const pathCol = heads[normKey('下橋路徑(ABCD)')] || 0;
    const auxCol = heads[normKey('輔助設備(A/B)')] || 0;
    const vals = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const id = idCol ? String(r[idCol - 1] || '').trim() : '';
      const name = nameCol ? String(r[nameCol - 1] || '').trim() : '';
      if (!id && !name) continue;
      bridges.push({
        sheet: sheet.getName(),
        row: i + 2,
        serial: serialCol ? String(r[serialCol - 1] || '').trim() : '',
        id: id,
        code: codeCol ? String(r[codeCol - 1] || '').trim() : '',
        name: name,
        done: doneCol ? String(r[doneCol - 1] || '').trim() : '',
        access_path: pathCol ? String(r[pathCol - 1] || '').trim() : '',
        access_aux: auxCol ? String(r[auxCol - 1] || '').trim() : '',
      });
    }
  }
  return { ok: true, sheets: sheets.map(s => s.getName()), bridges: bridges };
}

function findBridgeRow(ss, invId, code, name, onlySheet) {
  const sheets = ss.getSheets().filter(s => s.getName() !== GUIDE_SHEET && (!onlySheet || s.getName() === onlySheet));
  const tries = [
    [String(invId || '').trim(), normKey('二代系統ID')],
    [String(code || '').trim(), normKey('橋梁編號')],
    [String(name || '').trim(), normKey('橋梁名稱')],
  ];
  for (const [wanted, colKey] of tries) {
    if (!wanted) continue;
    for (const sheet of sheets) {
      const heads = sheetHeadersMap(sheet);
      const col = heads[colKey];
      const lastRow = sheet.getLastRow();
      if (!col || lastRow < 2) continue;
      const colVals = sheet.getRange(2, col, lastRow - 1, 1).getDisplayValues();
      for (let i = 0; i < colVals.length; i++) {
        if (colVals[i][0] !== '' && sameValue(colVals[i][0], wanted)) {
          return { sheet: sheet, row: i + 2, heads: heads };
        }
      }
    }
  }
  return null;
}

function computeRailingToBed(heightVal, railing, slab, girder) {
  if (!heightVal || railing === '' || slab === '' || girder === '') return '';
  const r = Number(railing), s = Number(slab), g = Number(girder);
  if (isNaN(r) || isNaN(s) || isNaN(g)) return '';
  const extra = r + s + g;
  const text = String(heightVal).normalize('NFKC').trim();
  const piers = [...text.matchAll(/\bP\s*0*(\d+)\s*[=:]\s*(\d+(?:\.\d+)?)/gi)];
  if (piers.length > 0) {
    return piers.map(m => 'P' + String(Number(m[1])).padStart(2, '0') + '=' + (Number(m[2]) + extra).toFixed(2)).join('\n');
  }
  const m = text.match(/\bH\s*[=:]\s*(\d+(?:\.\d+)?)/i) || text.match(/^(\d+(?:\.\d+)?)\s*m?$/i);
  if (m) return Number((Number(m[1]) + extra).toFixed(2));
  return '';
}

function handleSyncBridge(req) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(25000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const found = findBridgeRow(ss, req.inventory_id, req.code, req.name, req.sheet || '');
    if (!found) {
      return { ok: false, error: `總表中找不到這座橋（二代系統ID ${req.inventory_id}、${req.name}）` };
    }
    const { sheet, row, heads } = found;
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const colKeys = req.column_keys || {};
    const colLabels = req.column_labels || {};
    const ticks = new Set(req.ticks || []);
    const preserveManual = req.preserve_manual !== false;
    const locked = new Set(req.locked || ['prior1', 'prior2', 'prior3', 'height_prior', 'repairs']);
    const report = [`${sheet.getName()} 第 ${row} 列：${req.name || req.inventory_id}`];

    // 1. 系統資料核對（空白填入；不同標橘底加附註，不覆寫）
    for (const [name, val] of Object.entries(req.system || {})) {
      const col = heads[colKeys[name]];
      if (!col || val === '' || val == null) continue;
      const cell = sheet.getRange(row, col);
      const cur = cell.getDisplayValue().trim();
      if (cur === '') {
        cell.setValue(toValue(val));
        report.push(`${colLabels[name] || name}：填入系統值 ${val}`);
      } else if (!sameValue(cur, val)) {
        cell.setBackground(COLOR_CHECK);
        cell.setNote(`系統值：${val}（${today} 填報助手核對）`);
        report.push(`${colLabels[name] || name}：清單 ${cur} 與系統 ${val} 不同（已標橘底，未修改）`);
      }
    }

    // 2. 程式整理欄位（只寫入程式欄位，人工修改或既有無記號文字保留不覆寫）
    for (const [name, val] of Object.entries(req.values || {})) {
      const col = heads[colKeys[name]];
      if (!col) continue;
      const cell = sheet.getRange(row, col);
      const cur = cell.getDisplayValue().trim();
      const note = cell.getNote();
      const state = generatedState(cur, note);

      if (!locked.has(name)) {
        // 本期資料不鎖定：每次以最新覆寫，移除舊的程式附註與「淨高與前期不同」橘底（2026-11-29 使用者）
        if (ticks.has(name) && !val && cur !== '') continue;
        if (name === 'height_now' && !val) continue;
        if (cur !== String(val)) cell.setValue(toValue(val));
        if (note && (note.indexOf(GENERATED_TAG) >= 0 || note.indexOf('淨高與前期不同') >= 0)) {
          cell.setNote('');
          if (String(cell.getBackground()).toLowerCase() === String(COLOR_CHECK).toLowerCase()) cell.setBackground(null);
        }
        if (name === 'done' && val) {
          const nameCol = heads[normKey('橋梁名稱')];
          if (nameCol) sheet.getRange(row, nameCol).setBackground(COLOR_DONE);
        }
        continue;
      }
      if (ticks.has(name) && !val && state === null && cur !== '') continue;
      if (state === 'edited') {
        report.push(`${colLabels[name] || name}：已人工修正，保留不覆寫`);
        continue;
      }
      if (preserveManual && state === null && cur !== '' && cur !== '無' && name !== 'done' && !ticks.has(name)) {
        report.push(`${colLabels[name] || name}：已有人工內容，保留不覆寫`);
        continue;
      }
      if (name === 'height_prior' && cur !== '' && state === null) continue;

      const extra = '';
      if (cur !== String(val)) {
        cell.setValue(toValue(val));
      }
      const stamp = contentStamp(val);
      cell.setNote((extra ? extra + '\n' : '') + `${GENERATED_TAG} ${today}｜識別碼 ${stamp}（修改這格後，程式同步時不會覆寫）`);
      if (name === 'done' && val) {
        const nameCol = heads[normKey('橋梁名稱')];
        if (nameCol) sheet.getRange(row, nameCol).setBackground(COLOR_DONE);
      }
    }

    // 3. 多跨超過 8 墩或非公式時自動更新「欄杆至河床高度(m)」
    const rtbCol = heads[normKey('欄杆至河床高度(m)')];
    if (rtbCol) {
      const rtbCell = sheet.getRange(row, rtbCol);
      const formula = rtbCell.getFormula();
      const hNowCol = heads[colKeys['height_now']];
      const hPriCol = heads[colKeys['height_prior']];
      const rCol = heads[colKeys['railing']];
      const sCol = heads[normKey('橋面板厚(m)')];
      const gCol = heads[normKey('大梁高度(m)')];
      if (rCol && sCol && gCol && (hNowCol || hPriCol)) {
        const hNow = hNowCol ? sheet.getRange(row, hNowCol).getDisplayValue().trim() : '';
        const hPri = hPriCol ? sheet.getRange(row, hPriCol).getDisplayValue().trim() : '';
        const activeH = hNow || hPri;
        const piers = [...String(activeH).matchAll(/\bP\s*0*(\d+)\s*[=:]/gi)];
        if (piers.length > 8 || (!formula && rtbCell.getDisplayValue().trim() === '')) {
          const calc = computeRailingToBed(
            activeH,
            sheet.getRange(row, rCol).getDisplayValue().trim(),
            sheet.getRange(row, sCol).getDisplayValue().trim(),
            sheet.getRange(row, gCol).getDisplayValue().trim()
          );
          if (calc !== '') rtbCell.setValue(calc);
        }
      }
    }

    // 4. 序號標綠（非只核對基本資料時）
    if (!req.only_basic) {
      const sCol = heads[normKey('序號')];
      if (sCol) sheet.getRange(row, sCol).setBackground(COLOR_SERIAL);
    }

    return { ok: true, sheet: sheet.getName(), row: row, report: report };
  } finally {
    lock.releaseLock();
  }
}

function handleExportAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];
  for (const sheet of ss.getSheets()) {
    if (sheet.getName() === GUIDE_SHEET) continue;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) continue;
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    if (lastRow < 2) {
      out.push({ name: sheet.getName(), headers: headers, rows: [], fills: [], notes: [] });
      continue;
    }
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    out.push({
      name: sheet.getName(),
      headers: headers,
      rows: range.getDisplayValues(),
      fills: range.getBackgrounds(),
      notes: range.getNotes(),
    });
  }
  return { ok: true, title: ss.getName(), sheets: out };
}
