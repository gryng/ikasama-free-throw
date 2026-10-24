/**
 * イカサマフリースロー オンラインランキング（Google Apps Script）
 *
 * 使い方
 *   1. 新しいGoogleスプレッドシートを作り、「拡張機能 → Apps Script」を開く
 *   2. このファイルの中身をまるごと貼り付けて保存
 *   3. 「デプロイ → 新しいデプロイ」で種類「ウェブアプリ」を選び、
 *      実行ユーザー「自分」、アクセスできるユーザー「全員」でデプロイ
 *   4. 発行された「ウェブアプリのURL（…/exec）」をゲームの RANK_URL に設定
 *
 * 記録
 *   scores シート … 1プレイ1行で追記（日時, プレイヤーID, 名前, スコア, 成功, 最大コンボ, 実力, 称号, バージョン）
 *   names  シート … プレイヤーIDごとの今の名前（名前を変えたら上書き）
 * ランキングはプレイヤーIDごとの最高スコアで並べる。プレイヤーIDはレスポンスに含めない。
 */
const SCORES = 'scores';
const NAMES = 'names';
const TOP_N = 50;
const DEFAULT_NAME = 'ななしのイカサマ師';

/** ランキングを返す。?pid=… を付けると、そのプレイヤーの順位（me）と、上位の中の自分（me: true）も返す */
function doGet(e) {
  const q = (e && e.parameter) || {};
  const limit = Math.min(50, Math.max(1, Math.floor(Number(q.limit)) || TOP_N));
  return json_(buildRanking_(cleanPid_(q.pid), null, limit));
}

/**
 * POST（本文はJSON。Content-Type: text/plain で送る）
 *   { action: 'score',  pid, name, score, makes, maxCombo, honest, title, version } … スコアを登録
 *   { action: 'rename', pid, name }                                             … 名前を変更
 */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  const pid = cleanPid_(d && d.pid);
  if (!pid) return json_({ ok: false, error: 'no_pid' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (d.action === 'rename') {
      setName_(pid, cleanName_(d.name));
      return json_(buildRanking_(pid, null, TOP_N));
    }
    const entry = validate_(d);
    if (!entry) return json_({ ok: false, error: 'invalid' });
    sheet_(SCORES).appendRow([new Date(), pid, entry.name, entry.score, entry.makes, entry.maxCombo, entry.honest, entry.title, entry.version]);
    setName_(pid, entry.name);
    return json_(buildRanking_(pid, entry.score, TOP_N));
  } finally {
    lock.releaseLock();
  }
}

function cleanPid_(v) { const s = String(v || '').replace(/[^0-9A-Za-z_-]/g, '').slice(0, 64); return s.length >= 6 ? s : ''; }
function cleanName_(v) { return String(v || '').replace(/[\u0000-\u001f\u007f]/g, '').replace(/\s+/g, ' ').trim().slice(0, 10) || DEFAULT_NAME; }

/** 送られてきた値の形と、ゲームのルール上ありえる範囲かを確認する */
function validate_(d) {
  const score = Math.floor(Number(d.score)), makes = Math.floor(Number(d.makes));
  const maxCombo = Math.floor(Number(d.maxCombo)), honest = Math.floor(Number(d.honest));
  if (![score, makes, maxCombo, honest].every(n => Number.isFinite(n) && n >= 0)) return null;
  if (score <= 0 || makes > 3000 || maxCombo > makes || honest > makes) return null;
  if (score < makes || score > maxScore_(makes)) return null;
  return {
    name: cleanName_(d.name), score, makes, maxCombo, honest,
    title: String(d.title || '').replace(/[\u0000-\u001f\u007f]/g, '').slice(0, 20),
    version: String(d.version || '').slice(0, 16),
  };
}

/**
 * その成功数で理論上とれる最高点（全部ゴールを動かさずに決めて、ノーミスでコンボを続けた場合）。
 * ゲーム側の配点（1本 = 動かさず2点 / 動かして1点 × コンボ倍率。3本ごとに+1、最大×4）と合わせること。
 */
function maxScore_(makes) {
  let max = 0;
  for (let i = 1; i <= makes; i++) max += 2 * Math.min(4, 1 + Math.floor((i - 1) / 3));
  return max;
}

function setName_(pid, name) {
  const sh = sheet_(NAMES), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === pid) { sh.getRange(i + 1, 2, 1, 2).setValues([[name, new Date()]]); return; }
  }
  sh.appendRow([pid, name, new Date()]);
}

function buildRanking_(pid, entryScore, limit) {
  const names = {};
  sheet_(NAMES).getDataRange().getValues().slice(1).forEach(r => { names[String(r[0])] = String(r[1]); });
  const best = {};
  sheet_(SCORES).getDataRange().getValues().slice(1).forEach(r => {
    const id = String(r[1]), score = Number(r[3]) || 0;
    if (!best[id] || score > best[id].score) best[id] = { id, score, title: String(r[7] || ''), name: String(r[2]) };
  });
  const list = Object.keys(best).map(k => best[k]).sort((a, b) => b.score - a.score);
  const res = {
    ok: true,
    players: list.length,
    top: list.slice(0, limit).map(p => {
      const o = { name: names[p.id] || p.name || DEFAULT_NAME, score: p.score, title: p.title };
      if (pid && p.id === pid) o.me = true;
      return o;
    }),
    me: null,
  };
  if (pid) {
    const i = list.findIndex(p => p.id === pid);
    if (i >= 0) res.me = { rank: i + 1, score: list[i].score };
  }
  // 今回のスコアが何位にあたるか（自分の過去の記録は数えない）
  if (entryScore != null) res.rank = list.filter(p => p.id !== pid && p.score > entryScore).length + 1;
  return res;
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === SCORES) sh.appendRow(['日時', 'プレイヤーID', '名前', 'スコア', '成功', '最大コンボ', '実力', '称号', 'バージョン']);
    else sh.appendRow(['プレイヤーID', '名前', '更新日時']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
