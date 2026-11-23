/* kijun-map.js を生成する。
   kousei-data.js の各骨子項目が持つ根拠条文（c[].s）を、
   checkbook-data.js（基準・留意事項 全200問）の該当Qに対応づける。

   使い方:  node tools/gen-kijun-map.js
   出力:    kijun-map.js（リポジトリ直下）＋ 標準出力に照合レポート
*/
'use strict';
const fs = require('fs');
const path = require('path');
const ROOT = path.join(__dirname, '..');
global.window = {};
require(path.join(ROOT, 'kousei-data.js'));
require(path.join(ROOT, 'checkbook-data.js'));
const KOUSEI = window.KOUSEI_DATA;
const CHECK = window.CHECKBOOK_DATA;

/* ---------- 文字列ユーティリティ ---------- */
const strip = s => String(s || '').replace(/<[^>]+>/g, '').replace(/\s+/g, '');
const NORM = s => strip(s)
  .replace(/[，、]/g, '')
  .replace(/[（(][^）)]*[）)]/g, '')
  .replace(/[ＡＤＣＦ]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));

function bigrams(s) {
  const t = NORM(s), m = new Map();
  for (let i = 0; i < t.length - 1; i++) {
    const g = t.slice(i, i + 2);
    m.set(g, (m.get(g) || 0) + 1);
  }
  return m;
}
/* a=項目側, b=基準Q側。
   dice      … 全体としての似かた
   contain   … 項目の内容がどれだけ Q に含まれているか（長い Q が不利にならない）
   長い基準条文の中の一部を骨子項目が指すことが多いので contain を重めに見る。 */
function sim(a, b) {
  let inter = 0, na = 0, nb = 0;
  for (const v of a.values()) na += v;
  for (const [g, v] of b) { nb += v; if (a.has(g)) inter += Math.min(v, a.get(g)); }
  if (!na || !nb) return 0;
  const dice = (2 * inter) / (na + nb);
  const contain = inter / na;
  return 0.45 * dice + 0.55 * contain;
}

/* ---------- 章の解決 ---------- */
/* checkbook の ch をキーに引けるよう索引化 */
const byCh = new Map();
CHECK.forEach(q => {
  const k = q.part + '/' + q.ch;
  if (!byCh.has(k)) byCh.set(k, []);
  byCh.get(k).push(q);
});
const allCh = [...byCh.keys()];
const chStartsWith = pre => allCh.filter(k => k.startsWith(pre));

/* 総第7章 の節 → checkbook の 第7章-N */
const S7 = [
  [/差額配分法|利回り法|スライド法|継続賃料を求める鑑定評価の手法/, '総論/第7章-7'],
  [/積算法|賃貸事例比較法|収益分析法|新規賃料を求める鑑定評価の手法/, '総論/第7章-6'],
  [/賃料を求める場合の一般的留意事項|実質賃料|支払賃料|賃料の算定の期間|継続賃料を求める場合|第2節 賃料を求める鑑定評価の手法/, '総論/第7章-5'],
  [/収益還元法/, '総論/第7章-4'],
  [/取引事例比較法/, '総論/第7章-3'],
  [/原価法/, '総論/第7章-2'],
  [/試算価格を求める場合の一般的留意事項|事情補正|時点修正|事例の収集|一般的要因と鑑定評価/, '総論/第7章-1'],
];
/* 各論第1章 の節 → checkbook の 各論第1章-N（順序が重要） */
const K1 = [
  [/借地権付建物/, '各論/各論第1章-9'],
  [/区分所有建物/, '各論/各論第1章-10'],
  [/区分地上権/, '各論/各論第1章-4'],
  [/宅地見込地/, '各論/各論第1章-6'],
  [/自用の建物及びその敷地/, '各論/各論第1章-7'],
  [/貸家及びその敷地/, '各論/各論第1章-8'],
  [/借地権及び底地|底地/, '各論/各論第1章-3'],
  [/農地|林地/, '各論/各論第1章-5'],
  [/更地/, '各論/各論第1章-1'],
  [/建付地/, '各論/各論第1章-2'],
  [/借家権|第3節 建物/, '各論/各論第1章-11'],
];
const K2 = [
  [/建物及びその敷地/, '各論/各論第2章-2'],
  [/宅地/, '各論/各論第2章-1'],
];

function chaptersFor(s) {
  const out = new Set();
  const re = /(総|各)(?:論)?第(\d+)章/g;
  let m;
  while ((m = re.exec(s))) {
    const part = m[1] === '総' ? '総論' : '各論';
    const n = Number(m[2]);
    if (part === '総論') {
      if (n === 7) {
        const hit = S7.find(([r]) => r.test(s));
        if (hit) chStartsWith(hit[1]).forEach(k => out.add(k));
        else chStartsWith('総論/第7章').forEach(k => out.add(k));
      } else {
        chStartsWith('総論/第' + n + '章').forEach(k => out.add(k));
      }
    } else {
      if (n === 1) {
        const hit = K1.find(([r]) => r.test(s));
        if (hit) chStartsWith(hit[1]).forEach(k => out.add(k));
        else chStartsWith('各論/各論第1章').forEach(k => out.add(k));
      } else if (n === 2) {
        const hit = K2.find(([r]) => r.test(s));
        if (hit) chStartsWith(hit[1]).forEach(k => out.add(k));
        else chStartsWith('各論/各論第2章').forEach(k => out.add(k));
      } else {
        chStartsWith('各論/各論第' + n + '章').forEach(k => out.add(k));
      }
    }
  }
  return [...out];
}

/* 引用文字列の末尾の見出し語（Ⅰ 対象確定条件 / 1.正常価格 など） */
function tailTerms(s) {
  const t = [];
  const re = /(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|\d+\.)\s*([^\sⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)/g;
  let m; while ((m = re.exec(s))) t.push(m[1]);
  const sec = /第\d+節\s*([^\sⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)/.exec(s);
  if (sec) t.push(sec[1]);
  return t;
}

/* ---------- 本処理 ---------- */
const qText = new Map();
CHECK.forEach(q => qText.set(q.id, bigrams(q.title + strip(q.body))));
const qTitle = new Map(CHECK.map(q => [q.id, q.title]));

const MAP = {};
const report = { high: 0, mid: 0, chOnly: 0, none: 0 };
const samples = [];
const unresolved = [];

KOUSEI.forEach(d => {
  (d.subs || []).forEach((sb, si) => {
    (sb.items || []).forEach((it, ii) => {
      const cites = (it.c || []).filter(x => x.k === '基' || x.k === '留');
      if (!cites.length) { report.none++; return; }

      const chs = new Set();
      const terms = [];
      let vague = true;            /* 節・項の指定がない（章だけの）引用か */
      cites.forEach(c => {
        chaptersFor(c.s).forEach(k => chs.add(k));
        const t = tailTerms(c.s);
        terms.push(...t);
        if (t.length || /第\d+節/.test(c.s)) vague = false;
      });
      if (!chs.size) { report.none++; return; }

      const cand = [];
      chs.forEach(k => (byCh.get(k) || []).forEach(q => cand.push(q)));
      if (!cand.length) { report.none++; unresolved.push(d.n + ' ' + it.t + ' ← ' + cites.map(c=>c.s).join(' / ')); return; }

      const itemBg = bigrams(it.t + ' ' + (it.p || []).join(' '));
      const scored = cand.map(q => {
        let sc = sim(itemBg, qText.get(q.id));
        /* 引用の見出し語が Q のタイトル／secに出てくれば加点 */
        terms.forEach(t => {
          if (t.length >= 2 && (NORM(q.title).includes(NORM(t)) || NORM(q.sec).includes(NORM(t)))) sc += 0.12;
        });
        /* 項目名が Q タイトルに含まれる／その逆 */
        const a = NORM(it.t), b = NORM(q.title);
        if (a.length >= 3 && (b.includes(a) || a.includes(b))) sc += 0.10;
        return { id: q.id, sc, ch: q.part + '/' + q.ch };
      }).sort((x, y) => y.sc - x.sc);

      const best = scored[0], second = scored[1];
      const gap = best.sc - (second ? second.sc : 0);
      /* 章だけの引用で候補章が広いときは、よほど一致しない限り「候補」どまりにする */
      const hiBar  = vague ? 0.46 : 0.34;
      const hiGap  = vague ? 0.10 : 0.04;
      let conf, qs;
      if (best.sc >= hiBar && gap >= hiGap) { conf = 'high'; qs = [best.id]; }
      else if (best.sc >= 0.24) { conf = 'mid'; qs = scored.filter(x => x.sc >= best.sc - 0.06).slice(0, 3).map(x => x.id); }
      else { conf = 'ch'; qs = []; }

      report[conf === 'high' ? 'high' : conf === 'mid' ? 'mid' : 'chOnly']++;
      MAP[d.n + '-' + si + '_' + ii] = {
        q: qs.map(id => ({ id, t: qTitle.get(id) })),
        ch: [...chs].map(k => k.split('/')[1]),
        c: conf
      };
      if (samples.length < 40 && (d.n <= 12 || d.n === 50)) {
        samples.push(`問${d.n} ${sb.name} ${it.i}. ${it.t}\n    → ${conf}${vague ? '(章のみ引用)' : ''} ${qs.map(x => x + '「' + qTitle.get(x) + '」').join(' / ') || '(章: ' + [...chs].join(',') + ')'}  [best=${best.sc.toFixed(3)} gap=${gap.toFixed(3)}]`);
      }
    });
  });
});

fs.writeFileSync(path.join(ROOT, 'kijun-map.js'),
  '/* 自動生成ファイル — kousei-data.js の根拠条文と checkbook-data.js を照合して作成。手で編集しない。\n' +
  '   生成元: tools/gen-kijun-map.js　骨子項目' + Object.keys(MAP).length + '件 → 基準Q200件 */\n' +
  'window.KIJUN_MAP = ' + JSON.stringify(MAP) + ';\n', 'utf8');

console.log('=== 照合結果 ===');
console.log('高確度(Q直リンク):', report.high);
console.log('中確度(候補複数)  :', report.mid);
console.log('章のみ            :', report.chOnly);
console.log('根拠なし(補)      :', report.none);
if (unresolved.length) {
  console.log('\n!!! 章を解決できなかった項目:', unresolved.length, '件');
  unresolved.slice(0, 15).forEach(x => console.log('   ', x));
}
console.log('\n=== サンプル（問題1〜12, 50） ===');
console.log(samples.join('\n'));
