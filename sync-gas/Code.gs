/**
 * 暗記チェックブック 同期用スクリプト（Google Apps Script）
 *
 * スプレッドシートに紐づけて「ウェブアプリ」としてデプロイし，
 * 発行されたURL（.../exec）と下の SECRET をアプリの設定欄に入力します。
 *
 * 記録は「記録」シートに次の形式で保存されます（人が見て確認できます）。
 *   A列: 問題（Q1〜）   B列: 通算の解答回数   C列: 直近5回の正誤（古い→新しい，○=正解 ✕=誤答）
 */

/* ▼ ここを自分だけの合言葉に書き換えてください（アプリに入力するものと同じ文字列） */
var SECRET = 'CHANGE-ME';

var SHEET_NAME = '記録';
var KEEP = 5;            /* 保存する正誤の数（アプリ側と揃える） */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (String(p.k || '') !== SECRET) {
      out = { ok: false, err: 'key' };                 /* 合言葉が違う */
    } else if (p.a === 'ping') {
      out = { ok: true, mode: 'ping', g: getGen_() };  /* 接続テスト */
    } else if (p.a === 'reset') {
      out = resetAll_();                               /* 全端末の記録を消去 */
    } else {
      out = sync_(String(p.d || '1'), p.g, p.p);
    }
  } catch (err) {
    out = { ok: false, err: String((err && err.message) || err) };
  }
  var body = JSON.stringify(out);
  if (p.cb) {
    return ContentService.createTextOutput(p.cb + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 世代番号 ----------
   リセットするたびに1つ増やす。端末が持つ世代より新しければ，
   その端末の送信内容は採用せず，空の記録を返して消去を伝える。 */
function getGen_() {
  var v = PropertiesService.getScriptProperties().getProperty('gen');
  return Number(v || 0);
}
function setGen_(g) {
  PropertiesService.getScriptProperties().setProperty('gen', String(g));
  var sh = getSheet_();
  sh.getRange('E3').setValue('世代');
  sh.getRange('E4').setValue(g);
}

/* ---------- 最後に解いた問題（続きから用） ---------- */
function getPos_() {
  var v = PropertiesService.getScriptProperties().getProperty('pos');
  return (v === null || v === '' || v == null) ? null : Number(v);
}
function setPos_(i) {
  var props = PropertiesService.getScriptProperties();
  if (i === null) { props.deleteProperty('pos'); }
  else { props.setProperty('pos', String(i)); }
  var sh = getSheet_();
  sh.getRange('E5').setValue('最後に解いた問題');
  sh.getRange('E6').setValue(i === null ? '' : 'Q' + (i + 1));
}

/* ---------- 同期 ---------- */
function sync_(payload, clientGen, clientPos) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);                                 /* 同時書き込みを防ぐ */
  try {
    var gen = getGen_();
    var cg = (clientGen === '' || clientGen == null) ? gen : Number(clientGen);

    /* 端末の世代が古い＝別の端末でリセットされた後。送信内容は捨てて現状を返す。 */
    if (cg < gen) {
      return { ok: true, d: currentPayload_(), g: gen, p: getPos_(), reset: true };
    }

    /* 位置は「その端末で進んだときだけ」送られてくる。後に届いた方を採用する。 */
    if (clientPos !== '' && clientPos != null) {
      var pi = Number(clientPos);
      if (pi >= 0 && pi < 1000) setPos_(pi);
    }

    if (payload.charAt(0) !== '1') throw new Error('format');
    var inc = decodeB64_(payload.slice(1));
    var cur = readSheet_();

    /* 問題ごとに「解答回数が多い方」を採用 */
    var len = Math.max(cur.length, inc.length);
    var out = [];
    for (var j = 0; j < len; j += 2) {
      var cn = cur[j] || 0, cb = cur[j + 1] || 0;
      var nn = inc[j] || 0, nb = inc[j + 1] || 0;
      if ((nn || nb) && nn >= cn) { out[j] = nn; out[j + 1] = nb; }
      else                        { out[j] = cn; out[j + 1] = cb; }
    }
    writeSheet_(out);
    return { ok: true, d: '1' + encodeB64_(out), g: gen, p: getPos_() };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 全消去 ---------- */
function resetAll_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
    }
    var g = getGen_() + 1;
    setGen_(g);
    setPos_(null);
    sh.getRange('E1').setValue('最終更新');
    sh.getRange('E2').setValue(new Date());
    return { ok: true, d: '1', g: g, p: null, reset: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- シート入出力 ---------- */
function readSheet_() {
  var sh = getSheet_();
  var out = [];
  var last = sh.getLastRow();
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < vals.length; i++) {
      out[i * 2]     = Number(vals[i][1]) || 0;
      out[i * 2 + 1] = encodeHist_(vals[i][2]);
    }
  }
  return out;
}
function writeSheet_(bytes) {
  var sh = getSheet_();
  var rows = [];
  for (var k = 0; k * 2 < bytes.length; k++) {
    rows.push(['Q' + (k + 1), bytes[k * 2] || 0, decodeHist_(bytes[k * 2 + 1] || 0)]);
  }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  sh.getRange('E1').setValue('最終更新');
  sh.getRange('E2').setValue(new Date());
}
function currentPayload_() {
  return '1' + encodeB64_(readSheet_());
}
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['問題', '解答回数', '直近の正誤']]);
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['問題', '解答回数', '直近の正誤']]);
  }
  return sh;
}

/* ---------- 変換 ---------- */
/* 「○✕○」→ 1バイト（上位3bit=件数, 下位5bit=正誤） */
function encodeHist_(s) {
  s = String(s == null ? '' : s);
  var arr = [];
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    if (c === '○') arr.push(1);
    else if (c === '✕' || c === '×' || c === 'x') arr.push(0);
  }
  if (arr.length > KEEP) arr = arr.slice(arr.length - KEEP);
  var bits = 0;
  for (var k = 0; k < arr.length; k++) if (arr[k]) bits |= (1 << k);
  return (arr.length << 5) | (bits & 31);
}
function decodeHist_(b) {
  var len = (b >> 5) & 7, bits = b & 31, s = '';
  for (var k = 0; k < len; k++) s += ((bits >> k) & 1) ? '○' : '✕';
  return s;
}

/* base64url（符号なしバイト配列と相互変換） */
function decodeB64_(s) {
  s = String(s || '').replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  if (!s) return [];
  var raw = Utilities.base64Decode(s);
  var out = [];
  for (var i = 0; i < raw.length; i++) out.push(raw[i] < 0 ? raw[i] + 256 : raw[i]);
  return out;
}
function encodeB64_(b) {
  var signed = [];
  for (var i = 0; i < b.length; i++) {
    var v = b[i] || 0;
    signed.push(v > 127 ? v - 256 : v);
  }
  return Utilities.base64Encode(signed)
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
