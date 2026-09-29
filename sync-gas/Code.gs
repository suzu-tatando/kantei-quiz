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
      out = { ok: true, mode: 'ping' };                /* 接続テスト */
    } else {
      out = { ok: true, d: mergeRecords_(String(p.d || '1')) };
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

/* 送られてきた記録とシート上の記録を突き合わせ，統合結果を返す */
function mergeRecords_(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);                                 /* 同時書き込みを防ぐ */
  try {
    if (payload.charAt(0) !== '1') throw new Error('format');
    var inc = decodeB64_(payload.slice(1));
    var sh = getSheet_();

    /* 現在の内容を読む */
    var cur = [];
    var last = sh.getLastRow();
    if (last > 1) {
      var vals = sh.getRange(2, 1, last - 1, 3).getValues();
      for (var i = 0; i < vals.length; i++) {
        cur[i * 2]     = Number(vals[i][1]) || 0;
        cur[i * 2 + 1] = encodeHist_(vals[i][2]);
      }
    }

    /* 問題ごとに「解答回数が多い方」を採用 */
    var len = Math.max(cur.length, inc.length);
    var out = [];
    for (var j = 0; j < len; j += 2) {
      var cn = cur[j] || 0, cb = cur[j + 1] || 0;
      var nn = inc[j] || 0, nb = inc[j + 1] || 0;
      if ((nn || nb) && nn >= cn) { out[j] = nn; out[j + 1] = nb; }
      else                        { out[j] = cn; out[j + 1] = cb; }
    }

    /* シートへ書き戻す */
    var rows = [];
    for (var k = 0; k * 2 < out.length; k++) {
      rows.push(['Q' + (k + 1), out[k * 2] || 0, decodeHist_(out[k * 2 + 1] || 0)]);
    }
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
    sh.getRange('E1').setValue('最終更新');
    sh.getRange('E2').setValue(new Date());

    return '1' + encodeB64_(out);
  } finally {
    lock.releaseLock();
  }
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
