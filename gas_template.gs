// ============================================================
// 実地棚卸支援ツール - GASテンプレート
// ============================================================
// 【使い方】
// 1. 新しいGoogleスプレッドシートを作成する
// 2. メニュー「拡張機能」→「Apps Script」を開く
// 3. デフォルトで入っているコードを全部消し、このファイルの中身を
//    全部貼り付けて保存する（フロッピーアイコン、またはCtrl+S）
// 4. 右上の「デプロイ」→「新しいデプロイ」をクリック
//    - 種類の選択：歯車アイコン→「ウェブアプリ」
//    - 説明：任意（例：棚卸2026）
//    - 次のユーザーとして実行：自分
//    - アクセスできるユーザー：全員
//    →「デプロイ」をクリック
// 5. 初回のみ権限の承認を求められるので許可する
//    （「このアプリは Google で確認されていません」と出た場合は
//     「詳細」→「安全ではないページに移動」で進めて問題ありません。
//     これは自分自身が作成したスクリプトのため安全です）
// 6. 表示された「ウェブアプリのURL」をコピーし、
//    棚卸ツール（管理者モード）の設定画面に貼り付ける
//
// シートやその中の見出し・数式は、ツール側から取込データを送った時点で
// 自動的に作成・更新されます。あらかじめシートを手作業で用意する必要は
// ありません。コードを更新した場合は、既存のデプロイに対して
// 「デプロイを管理」→「新バージョン」で反映させてください
// （コードの保存だけでは、公開中のURLには反映されません）。
// ============================================================

const SHEET_GUIDE = '使い方';
const SHEET_MASTER = 'マスタ';
const SHEET_ANSWERS = '回答';
const SHEET_SUMMARY = '集計';
const SHEET_FINAL = '最終リスト';
const SHEET_ORIGINAL = '元データ';
const SHEET_FINAL_ORIGINAL = '最終（元レイアウト）';

// ---------------- ヘッダー定義（軸の数に応じて可変） ----------------

function masterHeader_(axisLabels) {
  return ['識別キー', '表示名'].concat(axisLabels).concat(['帳簿数量', '元データ(JSON)']);
}
function answersHeader_(axisLabels) {
  return ['タイムスタンプ', '担当者名', '識別キー'].concat(axisLabels).concat(['数量', '場所メモ']);
}
function summaryHeader_(axisLabels) {
  return ['識別キー'].concat(axisLabels).concat(['表示名', '帳簿数量', '実地数（合算）', '入力件数', '表示用実地数', '差異', '内訳（担当者:数量(メモ)）', '管理者上書き', '上書き理由', '最終実地数']);
}
function finalHeader_(axisLabels) {
  return ['識別キー', '表示名'].concat(axisLabels).concat(['最終数量']);
}

// ---------------- エントリーポイント ----------------

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'ping') {
      return jsonOut({ ok: true });
    }
    if (action === 'getMaster') {
      return jsonOut({ ok: true, items: getMasterData(), config: getConfig() });
    }
    if (action === 'getConfig') {
      return jsonOut({ ok: true, config: getConfig() });
    }
    return jsonOut({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'setup') {
      return jsonOut(setupMaster(body));
    }
    if (action === 'submitCounts') {
      return jsonOut(submitCounts(body));
    }
    return jsonOut({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------- 初期化 ----------------

function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const axisLabels = getAxisLabels_();
  writeUsageGuideSheet_();
  ensureSheet_(ss, SHEET_MASTER, masterHeader_(axisLabels));
  ensureSheet_(ss, SHEET_ANSWERS, answersHeader_(axisLabels));
  ensureSheet_(ss, SHEET_SUMMARY, summaryHeader_(axisLabels));
  ensureSheet_(ss, SHEET_FINAL, finalHeader_(axisLabels));
  // 元データ・最終（元レイアウト）は元ファイルのヘッダーが分かってから作成する（setupMaster内）
}

// 「使い方」シート：管理者がスプレッドシート上で何を確認・操作すればよいかを、
// 時系列・自動/手動の別が分かる形でまとめる。内容は固定なので、毎回上書きして最新版に揃える。
function writeUsageGuideSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_GUIDE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_GUIDE);
    sh.setFrozenRows(1);
  }

  const header = ['No.', '区分', '担当', '作業場所', '内容'];
  const COLS = header.length;

  // セクション見出し行は、作成者メッセージで謳っている「まるごとサポート」の
  // 5パート（棚卸リスト作成／配付／カウント／回収／カウント結果入力）との対応を示す。
  // カウント・回収・結果入力は、送信と同時に自動完了するため1セクションに統合している。
  const items = [
    { section: '棚卸リスト作成' },
    { row: ['1', '手動', '管理者', '実地棚卸支援ツール', '在庫リストを取込み、列（識別キー・表示名・軸・帳簿数量）を指定する。プレビュー画面で件数・帳簿数量の合計が元データと一致するか確認してから送信する'] },
    { row: ['1-1', '自動', 'システム', 'マスタ／元データ', 'シートが作成され、取り込んだデータが保存される'] },
    { row: ['1-2', '自動', 'システム', '集計／最終リスト／最終（元レイアウト）', 'シートが作成され、差異・最終結果が計算される'] },
    { section: '棚卸リスト配付' },
    { row: ['2', '手動', '管理者', '実地棚卸支援ツール', '発行されたQRコード・リンクをカウント担当者に共有する（印刷・PDF保存推奨）'] },
    { section: 'カウント／棚卸リスト回収／カウント結果入力（送信と同時に自動完了）' },
    { row: ['3', '手動', 'カウント担当者', '実地棚卸支援ツール', 'QR・リンクを開き、名前を入力→担当（ロケーション等）を選択→実地でカウントした数量を入力して送信する'] },
    { row: ['3-1', '自動', 'システム', '回答', '入力内容が追記される'] },
    { row: ['3-2', '自動', 'システム', '集計', '再計算され、帳簿数量と実地数が異なる行が赤く強調表示される'] },
    { section: '棚卸差異の確認' },
    { row: ['4', '手動', '管理者', '集計', '差異のある品目は「内訳」列（誰が・いつ・どんなメモで入力したか）を見て、重複入力なのか、本当に複数箇所にあった（分散カウント）のかを判断する'] },
    { row: ['5', '手動', '管理者', '集計', '是正が必要な場合、該当行の「管理者上書き」に確定数量を、「上書き理由」に理由を入力する（現場での再確認が必要な場合もある）'] },
    { row: ['6', '手動', '管理者', 'スプレッドシートのメニュー', '「棚卸ツール→🔄 集計を再計算」を実行し、上書き内容を反映させる（上書き入力だけでは自動反映されないため必須）'] },
    { row: ['6-1', '自動', 'システム', '集計／最終リスト／最終（元レイアウト）', '上書き内容を反映して再計算される'] },
    { section: '棚卸結果の出力・ダウンロード' },
    { row: ['7', '手動', '管理者', '最終（元レイアウト）', 'CSV等で出力・ダウンロードし、在庫管理システムへ反映する'] }
  ];

  sh.clear();

  sh.getRange(1, 1, 1, COLS).setValues([header]).setFontWeight('bold');
  sh.getRange(1, 1, 1, COLS).setBackground('#e2e5ea');

  let r = 2;
  items.forEach(function (item) {
    if (item.section) {
      // セルの結合はせず、行全体を色付けしてA列にラベルを入れるだけにする
      sh.getRange(r, 1, 1, COLS).setBackground('#2563eb');
      sh.getRange(r, 1).setValue(item.section).setFontWeight('bold').setFontColor('#ffffff');
    } else {
      sh.getRange(r, 1, 1, COLS).setValues([item.row]);
      if (item.row[1] === '手動') {
        sh.getRange(r, 1, 1, COLS).setBackground('#fff7e6');
      }
    }
    r++;
  });
  const totalRows = r - 1;
  sh.getRange(2, 1, totalRows - 1, 1).setNumberFormat('@'); // No.列は"1-1"等の混在を統一して文字列表示にする

  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 60);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(5, 480);
  sh.getRange(2, 4, totalRows - 1, 1).setWrap(true);
  sh.getRange(2, 5, totalRows - 1, 1).setWrap(true);

  trimSheet_(sh, totalRows, COLS);
}

function ensureSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.setFrozenRows(1);
  }
  // 既存シートでも、コード側の見出し定義が変わった場合に追従できるよう毎回上書きする
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  return sh;
}

// 見出し+データで使っている行・列数を超える分を削除し、見た目を整える
function trimSheet_(sh, usedRows, usedCols) {
  if (sh.getMaxRows() > usedRows) sh.deleteRows(usedRows + 1, sh.getMaxRows() - usedRows);
  if (sh.getMaxColumns() > usedCols) sh.deleteColumns(usedCols + 1, sh.getMaxColumns() - usedCols);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('棚卸ツール')
    .addItem('🔄 集計を再計算', 'recomputeSummary')
    .addToUi();
}

// ---------------- 設定の保存・読み出し ----------------

function getAxisLabels_() {
  const raw = PropertiesService.getScriptProperties().getProperty('axisLabels');
  return raw ? JSON.parse(raw) : [];
}
function getOriginalHeaders_() {
  const raw = PropertiesService.getScriptProperties().getProperty('originalHeaders');
  return raw ? JSON.parse(raw) : [];
}
function getAxisColIndices_() {
  const raw = PropertiesService.getScriptProperties().getProperty('axisColIndices');
  return raw ? JSON.parse(raw) : [];
}
function getBookQtyColIndex_() {
  const raw = PropertiesService.getScriptProperties().getProperty('bookQtyColIndex');
  return (raw !== null && raw !== undefined && raw !== '') ? Number(raw) : -1;
}

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    keyLabel: props.getProperty('keyLabel') || '識別キー',
    displayLabel: props.getProperty('displayLabel') || '表示名',
    axisLabels: getAxisLabels_()
  };
}

// ---------------- マスタ取込 ----------------

// body = {
//   action:'setup',
//   config:{keyLabel, displayLabel, axisLabels:[...]},
//   items:[{key, display, axes:[...], bookQty, originalRow:[...]}],
//   originalHeaders:[...], originalRows:[[...], ...],
//   axisColIndices:[...], bookQtyColIndex: number
// }
function setupMaster(body) {
  const cfg = body.config || {};
  const axisLabels = cfg.axisLabels || [];

  const props = PropertiesService.getScriptProperties();
  props.setProperty('keyLabel', cfg.keyLabel || '識別キー');
  props.setProperty('displayLabel', cfg.displayLabel || '表示名');
  props.setProperty('axisLabels', JSON.stringify(axisLabels));
  props.setProperty('originalHeaders', JSON.stringify(body.originalHeaders || []));
  props.setProperty('axisColIndices', JSON.stringify(body.axisColIndices || []));
  props.setProperty('bookQtyColIndex', String(body.bookQtyColIndex != null ? body.bookQtyColIndex : -1));

  // シートの並び順を「元データ→マスタ→回答→集計→最終リスト→最終（元レイアウト）」にするため、
  // 元データを先に作成しておく
  writeOriginalDataSheet_(body.originalHeaders || [], body.originalRows || []);

  ensureSheetsExist();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MASTER);
  const header = masterHeader_(axisLabels);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, header.length).clearContent();
  }

  const items = body.items || [];
  const rows = items.map(function (it) {
    const axes = it.axes || [];
    return [it.key, it.display].concat(axes).concat([Number(it.bookQty) || 0, JSON.stringify(it.originalRow || [])]);
  });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  trimSheet_(sh, rows.length + 1, header.length);

  recomputeSummary();
  return { ok: true, count: rows.length };
}

function getMasterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MASTER);
  if (!sh) return [];
  const axisLabels = getAxisLabels_();
  const n = axisLabels.length;
  const headerLen = masterHeader_(axisLabels).length;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, headerLen).getValues();
  return values
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      return { key: r[0], display: r[1], axes: r.slice(2, 2 + n), bookQty: r[2 + n] };
    });
}

function writeOriginalDataSheet_(originalHeaders, originalRows) {
  if (!originalHeaders || originalHeaders.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureSheet_(ss, SHEET_ORIGINAL, originalHeaders);
  const clearRows = sh.getMaxRows() - 1;
  if (clearRows > 0) sh.getRange(2, 1, clearRows, originalHeaders.length).clearContent();
  if (originalRows.length > 0) sh.getRange(2, 1, originalRows.length, originalHeaders.length).setValues(originalRows);
  trimSheet_(sh, originalRows.length + 1, originalHeaders.length);
}

// ---------------- カウント結果の受付 ----------------

// body = { action:'submitCounts', name:'担当者名', items:[{key, axes:[...], qty, memo}, ...] }
function submitCounts(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSheetsExist();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_ANSWERS);
    const axisLabels = getAxisLabels_();
    const header = answersHeader_(axisLabels);
    const now = new Date();
    const name = body.name || '';
    const items = body.items || [];
    const rows = items.map(function (it) {
      const axes = it.axes || [];
      return [now, name, it.key].concat(axes).concat([Number(it.qty) || 0, it.memo || '']);
    });
    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
    }
    trimSheet_(sh, Math.max(sh.getLastRow(), 1), header.length);
    // 集計の再計算もロックを保持したまま行う（複数端末の同時送信による競合を避けるため）
    try { recomputeSummary(); } catch (e) { /* 集計失敗しても回答は保存済みなので握りつぶす */ }
    return { ok: true, count: rows.length };
  } finally {
    lock.releaseLock();
  }
}

// ---------------- 集計の再計算 ----------------
// マスタと回答、両方に登場する(識別キー,軸1..軸N)の組み合わせを全て拾い、
// 帳簿数量・実地数（合算）・入力件数・差異を算出してシートへ書き込む。
// 「管理者上書き」「上書き理由」は既存の値を保持する。
// 併せて「最終リスト」「最終（元レイアウト）」も更新する。

function recomputeSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetsExist();
  const axisLabels = getAxisLabels_();
  const n = axisLabels.length;

  const masterSh = ss.getSheetByName(SHEET_MASTER);
  const ansSh = ss.getSheetByName(SHEET_ANSWERS);
  const sumSh = ss.getSheetByName(SHEET_SUMMARY);

  const masterHeaderLen = masterHeader_(axisLabels).length;
  const masterLast = masterSh.getLastRow();
  const masterRows = masterLast > 1 ? masterSh.getRange(2, 1, masterLast - 1, masterHeaderLen).getValues() : [];
  trimSheet_(masterSh, Math.max(masterLast, 1), masterHeaderLen);

  const ansHeaderLen = answersHeader_(axisLabels).length;
  const ansLast = ansSh.getLastRow();
  const ansRows = ansLast > 1 ? ansSh.getRange(2, 1, ansLast - 1, ansHeaderLen).getValues() : [];
  trimSheet_(ansSh, Math.max(ansLast, 1), ansHeaderLen);

  const originalHeadersForTrim = getOriginalHeaders_();
  if (originalHeadersForTrim.length > 0) {
    const origSh = ss.getSheetByName(SHEET_ORIGINAL);
    if (origSh) trimSheet_(origSh, Math.max(origSh.getLastRow(), 1), originalHeadersForTrim.length);
  }

  const sumHeaderLen = summaryHeader_(axisLabels).length;
  const sumLast = sumSh.getLastRow();
  const oldSumRows = sumLast > 1 ? sumSh.getRange(2, 1, sumLast - 1, sumHeaderLen).getValues() : [];

  // 集計シートの列位置（軸の数nに応じて可変）
  const idx = {
    idKey: 0,
    axisStart: 1,
    display: n + 1,
    bookQty: n + 2,
    actualSum: n + 3,
    inputCount: n + 4,
    displayActual: n + 5,
    diff: n + 6,
    breakdown: n + 7,
    override: n + 8,
    reason: n + 9,
    finalQty: n + 10
  };

  // 既存の管理者上書き・理由を保持
  const overrideMap = {};
  oldSumRows.forEach(function (r) {
    const axes = r.slice(idx.axisStart, idx.axisStart + n);
    const k = combineKey_(r[idx.idKey], axes);
    overrideMap[k] = { override: r[idx.override], reason: r[idx.reason] };
  });

  // マスタ側キー: [識別キー, 表示名, 軸(n), 帳簿数量, 元データ(JSON)]
  const bookMap = {};
  masterRows.forEach(function (r) {
    if (r[0] === '' || r[0] === null) return;
    const axes = r.slice(2, 2 + n);
    const bookQty = Number(r[2 + n]) || 0;
    const key = combineKey_(r[0], axes);
    bookMap[key] = { idKey: r[0], display: r[1], axes: axes, bookQty: bookQty };
  });

  // 回答側の合算（担当者名・メモも内訳表示用に保持）: [タイムスタンプ, 担当者名, 識別キー, 軸(n), 数量, メモ]
  const ansAgg = {};
  ansRows.forEach(function (r) {
    const name = r[1], idKey = r[2];
    const axes = r.slice(3, 3 + n);
    const qty = Number(r[3 + n]) || 0;
    const memo = r[3 + n + 1] || '';
    if (idKey === '' || idKey === null) return;
    const key = combineKey_(idKey, axes);
    if (!ansAgg[key]) ansAgg[key] = { sum: 0, count: 0, idKey: idKey, axes: axes, details: [] };
    ansAgg[key].sum += qty;
    ansAgg[key].count += 1;
    ansAgg[key].details.push({ name: name, qty: qty, memo: memo });
  });

  const allKeys = {};
  Object.keys(bookMap).forEach(function (k) { allKeys[k] = true; });
  Object.keys(ansAgg).forEach(function (k) { allKeys[k] = true; });

  const outRows = [];
  Object.keys(allKeys).forEach(function (key) {
    const book = bookMap[key];
    const ans = ansAgg[key];
    const idKey = book ? book.idKey : ans.idKey;
    const display = book ? book.display : '';
    const axes = book ? book.axes : ans.axes;
    const bookQty = book ? book.bookQty : 0;
    const actualSum = ans ? ans.sum : 0;
    const inputCount = ans ? ans.count : 0;
    const displayActual = inputCount === 0 ? bookQty : actualSum;
    const diff = displayActual - bookQty;

    const prev = overrideMap[key] || {};
    const override = (prev.override === undefined || prev.override === null) ? '' : prev.override;
    const reason = prev.reason || '';
    const finalQty = (override !== '') ? Number(override) : displayActual;
    const breakdown = buildBreakdown_(ans);

    const row = [idKey].concat(axes).concat([display, bookQty, actualSum, inputCount, displayActual, diff, breakdown, override, reason, finalQty]);
    outRows.push(row);
  });

  outRows.sort(function (a, b) { return Math.abs(b[idx.diff]) - Math.abs(a[idx.diff]); });

  const clearRows = sumSh.getMaxRows() - 1;
  if (clearRows > 0) {
    sumSh.getRange(2, 1, clearRows, sumHeaderLen).clearContent();
  }
  if (outRows.length > 0) {
    sumSh.getRange(2, 1, outRows.length, sumHeaderLen).setValues(outRows);
  }
  trimSheet_(sumSh, outRows.length + 1, sumHeaderLen);

  applyConditionalFormatting_(sumSh, idx.diff, sumHeaderLen);
  updateFinalList_(outRows, axisLabels, idx);
  updateOriginalLayoutFinal_(outRows, axisLabels, idx);
  reorderSheets_();
}

// シートのタブ順を「元データ→マスタ→回答→集計→最終リスト→最終（元レイアウト）」に揃える。
// 既存のスプレッドシート（過去のバージョンで別順に作られたもの）でも、実行するたびに揃う。
function reorderSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const order = [SHEET_GUIDE, SHEET_ORIGINAL, SHEET_MASTER, SHEET_ANSWERS, SHEET_SUMMARY, SHEET_FINAL, SHEET_FINAL_ORIGINAL];
  order.forEach(function (name, i) {
    const sh = ss.getSheetByName(name);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    }
  });
}

function combineKey_(idKey, axes) {
  function norm(v) { return (v === null || v === undefined) ? '' : String(v); }
  return norm(idKey) + '|' + (axes || []).map(norm).join('|');
}

function buildBreakdown_(ans) {
  if (!ans || !ans.details || ans.details.length === 0) return '';
  return ans.details.map(function (d) {
    let label = (d.name || '(名前なし)') + ':' + d.qty;
    if (d.memo) label += '(' + d.memo + ')';
    return label;
  }).join(' / ');
}

function colLetter_(colNum1based) {
  let s = '', n = colNum1based;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function applyConditionalFormatting_(sh, diffIdx, totalCols) {
  sh.clearConditionalFormatRules();
  const lastColLetter = colLetter_(totalCols);
  const diffColLetter = colLetter_(diffIdx + 1);
  const range = sh.getRange('A2:' + lastColLetter + '2000');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + diffColLetter + '2<>0')
    .setBackground('#fce8e6')
    .setRanges([range])
    .build();
  sh.setConditionalFormatRules([rule]);
}

function updateFinalList_(outRows, axisLabels, idx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_FINAL);
  const header = finalHeader_(axisLabels);
  const clearRows = sh.getMaxRows() - 1;
  if (clearRows > 0) {
    sh.getRange(2, 1, clearRows, header.length).clearContent();
  }
  const rows = outRows.map(function (r) {
    const idKey = r[idx.idKey];
    const axes = r.slice(idx.axisStart, idx.axisStart + axisLabels.length);
    const display = r[idx.display];
    const finalQty = r[idx.finalQty];
    return [idKey, display].concat(axes).concat([finalQty]);
  });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  trimSheet_(sh, rows.length + 1, header.length);
}

// 元ファイルと同じレイアウトのまま、帳簿数量の列だけ最終実地数に置き換えたシートを作る。
// 元の取込時に無かった（品番,軸）の組み合わせ（ロケーション不一致等で新規発見されたもの）は、
// 同じ識別キーを持つ既存行（マスタに保存した元データJSON）を複製し、軸列と数量列だけ書き換えて追加する。
function updateOriginalLayoutFinal_(outRows, axisLabels, idx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalHeaders = getOriginalHeaders_();
  if (!originalHeaders || originalHeaders.length === 0) return;
  const sh = ensureSheet_(ss, SHEET_FINAL_ORIGINAL, originalHeaders);
  const bookQtyColIndex = getBookQtyColIndex_();
  const axisColIndices = getAxisColIndices_();

  const masterSh = ss.getSheetByName(SHEET_MASTER);
  const masterHeaderLen = masterHeader_(axisLabels).length;
  const masterLast = masterSh.getLastRow();
  const masterRows = masterLast > 1 ? masterSh.getRange(2, 1, masterLast - 1, masterHeaderLen).getValues() : [];
  const jsonColIdx = masterHeaderLen - 1;

  // 識別キー -> 元データ(JSON) のテンプレート（新規発見された組み合わせの複製元に使う）
  const templateByKey = {};
  masterRows.forEach(function (r) {
    const idKey = r[0];
    if (idKey === '' || idKey === null) return;
    if (!templateByKey[idKey]) {
      try { templateByKey[idKey] = JSON.parse(r[jsonColIdx]); } catch (e) { /* ignore */ }
    }
  });

  const outSheetRows = [];
  outRows.forEach(function (row) {
    const idKey = row[idx.idKey];
    const axes = row.slice(idx.axisStart, idx.axisStart + axisLabels.length);
    const finalQty = row[idx.finalQty];
    const template = templateByKey[idKey];
    if (!template) return; // 元データが無い場合（元データ未設定等）はスキップ
    const clone = template.slice();
    axisColIndices.forEach(function (colIdx, i) {
      if (colIdx > -1 && i < axes.length) clone[colIdx] = axes[i];
    });
    if (bookQtyColIndex > -1) clone[bookQtyColIndex] = finalQty;
    outSheetRows.push(clone);
  });

  const clearRows = sh.getMaxRows() - 1;
  if (clearRows > 0) sh.getRange(2, 1, clearRows, originalHeaders.length).clearContent();
  if (outSheetRows.length > 0) sh.getRange(2, 1, outSheetRows.length, originalHeaders.length).setValues(outSheetRows);
  trimSheet_(sh, outSheetRows.length + 1, originalHeaders.length);
}
