/**
 * ============================================================
 *  古典クエスト ― サーバー側（Google Apps Script）
 *  埼玉県公立高校入試 大問4（古文）対策
 *
 *  漢字クエスト／語句クエストと同じ構成です。
 *   ・集計用スプレッドシート … 名簿・解答履歴・集計・疑義報告・設定
 *   ・問題マスター           … 本文と設問のプール
 *
 *  ★ 集計用シートは「古典_」で始まる名前にしてあります。
 *    そのため、語句クエストなどの既存の集計スプレッドシートを
 *    そのまま SPREADSHEET_ID に指定して相乗りさせることもできます
 *    （その場合「名簿」シートは既存のものがそのまま使われます）。
 * ============================================================
 */

const SPREADSHEET_ID = "ここに集計用スプレッドシートのIDを貼る";
const MASTER_ID      = "ここに問題マスターのIDを貼る";

/** index.html 側の CLIENT_VERSION と必ず同じ値にすること */
const CLIENT_VERSION = 1;

const APP_NAME  = "古典クエスト";
const PREFIX    = "古典_";                 // 集計側シート名の接頭辞
const SH_ROSTER = "名簿";                  // 既存アプリと共用できるよう接頭辞なし
const SH_RESULT = PREFIX + "結果";
const SH_ANSWER = PREFIX + "解答履歴";
const SH_TOTAL  = PREFIX + "集計";
const SH_ISSUE  = PREFIX + "疑義報告";
const SH_CONFIG = PREFIX + "設定";

const TEACHER_NO   = "999";                // 教員アカウント（各種制限をバイパス）
const CACHE_KEY    = "kobun_pool_v1";
const CACHE_SEC    = 21600;                // 問題プールのキャッシュ 6時間
const TARGET_SEC   = 300;                  // 1ステージの目安時間（5分）
const PLACEHOLDER  = "ここに";

/** 称号（合計スコア基準・下が上位） */
const RANKS = [
  { min: 22000, name: "古典マスター" },
  { min: 14000, name: "古典の師範"   },
  { min:  8000, name: "助動詞の達人" },
  { min:  4000, name: "仮名遣い名人" },
  { min:  1500, name: "説話の読み手" },
  { min:     0, name: "古文の旅人"   }
];

const Q_TYPES = ["仮名遣い", "主語・語意", "内容理解", "要旨・話し合い"];

/* ============================================================
   ウェブアプリ
   ============================================================ */
function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ============================================================
   スプレッドシートのユーティリティ
   ============================================================ */
function book_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf(PLACEHOLDER) === 0) {
    throw new Error("Code.gs の SPREADSHEET_ID が未設定です。集計用スプレッドシートのIDを貼り付けてください。");
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function master_() {
  if (!MASTER_ID || MASTER_ID.indexOf(PLACEHOLDER) === 0) {
    throw new Error("Code.gs の MASTER_ID が未設定です。問題マスターのIDを貼り付けてください。");
  }
  return SpreadsheetApp.openById(MASTER_ID);
}

/** シートが無ければヘッダ付きで作る */
function sheet_(name, headers) {
  const ss = book_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight("bold").setBackground("#16223d").setFontColor("#ffffff");
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

/** 見出し行を {列名: 列番号(1始まり)} に */
function headerMap_(sh) {
  const last = sh.getLastColumn();
  if (last < 1) return {};
  const head = sh.getRange(1, 1, 1, last).getValues()[0];
  const map = {};
  head.forEach(function (h, i) { if (h !== "") map[String(h).trim()] = i + 1; });
  return map;
}

function uuid_() { return Utilities.getUuid(); }
function now_()  { return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"); }
function today_(){ return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd"); }

/* ============================================================
   シート定義
   ============================================================ */
const H_RESULT = ["UUID","保存日時","出席番号","氏名","ステージID","ステージ名","出典",
                  "得点","満点","正答数","問題数","正答率(%)","所要時間(秒)",
                  "基礎点","タイム点","満点ボーナス","倍率","獲得スコア","自己ベスト更新",
                  "誤答問題ID","正答問題ID","フォーカス離脱回数","不正疑い","重複排除キー","クライアント版"];

const H_ANSWER = ["UUID","結果UUID","保存日時","出席番号","氏名","ステージID","問題ID","問番号",
                  "設問タイプ","生徒の解答","正解","正誤","配点","所要時間(秒)"];

const H_TOTAL  = ["順位","出席番号","氏名","合計スコア","挑戦回数","クリアステージ数","満点ステージ数",
                  "平均正答率(%)","最高得点","最速タイム(秒)","称号","最終挑戦日時"];

const H_ISSUE  = ["UUID","報告日時","出席番号","氏名","結果UUID","問題ID","設問文",
                  "選択肢ア","選択肢イ","選択肢ウ","選択肢エ","正解","生徒の回答",
                  "疑義種別","自由記述","ステータス","承認日時","加算済み","メモ"];

const H_ROSTER = ["出席番号","コード","氏名"];

/* ============================================================
   初回セットアップ
   ============================================================ */
function setupSheets() {
  const ui = (typeof SpreadsheetApp.getUi === "function") ? tryUi_() : null;
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf(PLACEHOLDER) === 0) {
    const msg = "Code.gs の SPREADSHEET_ID がプレースホルダのままです。先に集計用スプレッドシートのIDを設定してください。";
    if (ui) ui.alert(msg); else Logger.log(msg);
    return;
  }
  sheet_(SH_RESULT, H_RESULT);
  sheet_(SH_ANSWER, H_ANSWER);
  sheet_(SH_TOTAL,  H_TOTAL);
  sheet_(SH_ISSUE,  H_ISSUE);

  // 名簿（既存アプリのものがあればそのまま使う）
  const ss = book_();
  if (!ss.getSheetByName(SH_ROSTER)) {
    const sh = sheet_(SH_ROSTER, H_ROSTER);
    sh.getRange(2, 1, 3, 3).setValues([
      [TEACHER_NO, "9999",    "せんせい"],
      ["101",      "SAMPLE1", "サンプル太郎"],
      ["102",      "SAMPLE2", "サンプル花子"]
    ]);
  }

  // 設定
  let cfg = ss.getSheetByName(SH_CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(SH_CONFIG);
    cfg.getRange(1, 1, 2, 2).setValues([
      ["授業中ボーナス", false],
      ["時間制限解除",   false]
    ]);
    cfg.getRange(1, 1, 2, 1).setFontWeight("bold");
    cfg.setColumnWidth(1, 160);
  }
  const msg = "セットアップが完了しました。\n\n・" + SH_ROSTER + " に生徒を入力してください（出席番号 / 個人コード / 氏名）。\n・出席番号 " + TEACHER_NO + " は教員アカウントです。消さないでください。";
  if (ui) ui.alert(APP_NAME, msg, ui.ButtonSet.OK); else Logger.log(msg);
}

function tryUi_() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } }

function onOpen() {
  const ui = tryUi_();
  if (!ui) return;
  ui.createMenu("⚙ " + APP_NAME + "管理")
    .addItem("初回セットアップ（必要なシートを生成）", "setupSheets")
    .addSeparator()
    .addItem("問題プールのキャッシュをクリア", "clearPoolCache")
    .addItem("集計を再計算", "rebuildTotals")
    .addSeparator()
    .addItem("疑義報告サマリを表示", "showIssueSummary")
    .addToUi();
}

function clearPoolCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  const ui = tryUi_();
  if (ui) ui.alert(APP_NAME, "問題プールのキャッシュをクリアしました。次回出題から最新の問題マスターが読み込まれます。", ui.ButtonSet.OK);
}

/* ============================================================
   設定
   ============================================================ */
function readConfig_() {
  const ss = book_();
  const sh = ss.getSheetByName(SH_CONFIG);
  const cfg = { classBonus: false, timeFree: false };
  if (!sh) return cfg;
  const v = sh.getRange(1, 1, 2, 2).getValues();
  v.forEach(function (row) {
    const key = String(row[0]).trim();
    const on  = (row[1] === true || String(row[1]).toUpperCase() === "TRUE");
    if (key === "授業中ボーナス") cfg.classBonus = on;
    if (key === "時間制限解除")   cfg.timeFree   = on;
  });
  return cfg;
}

/** 深夜0:00〜5:00は挑戦禁止（睡眠時間の保護）。教員・解除設定時はバイパス */
function timeBlocked_(isTeacher) {
  if (isTeacher) return "";
  const cfg = readConfig_();
  if (cfg.timeFree || cfg.classBonus) return "";
  const h = Number(Utilities.formatDate(new Date(), "Asia/Tokyo", "H"));
  if (h >= 0 && h < 5) {
    return "深夜0時から5時までは挑戦できません。しっかり睡眠をとりましょう。";
  }
  return "";
}

/* ============================================================
   ログイン
   ============================================================ */
function login(number, code) {
  try {
    const no = String(number || "").trim();
    const cd = String(code || "").trim();
    if (!no || !cd) return { ok: false, message: "出席番号と個人コードを入力してください。" };

    const sh = book_().getSheetByName(SH_ROSTER);
    if (!sh) return { ok: false, message: "名簿シートがありません。先生にお知らせください（初回セットアップ未実行）。" };

    const last = sh.getLastRow();
    if (last < 2) return { ok: false, message: "名簿が空です。先生にお知らせください。" };

    const rows = sh.getRange(2, 1, last - 1, 3).getValues();
    let me = null;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === no && String(rows[i][1]).trim() === cd) {
        me = { no: no, name: String(rows[i][2]).trim() };
        break;
      }
    }
    if (!me) return { ok: false, message: "出席番号か個人コードが違います。" };

    me.teacher = (no === TEACHER_NO);
    const cfg = readConfig_();
    return {
      ok: true,
      student: me,
      settings: { classBonus: cfg.classBonus, timeFree: cfg.timeFree },
      blocked: timeBlocked_(me.teacher),
      stages: getStages_(),
      progress: getProgress_(me.no),
      version: CLIENT_VERSION
    };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/* ============================================================
   問題プール（問題マスターから読み込み・キャッシュ付き）
   ============================================================ */
function getStages_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* 壊れていたら読み直す */ }
  }
  const pool = loadPool_();
  try { cache.put(CACHE_KEY, JSON.stringify(pool), CACHE_SEC); } catch (e) { /* 6MB超は素通り */ }
  return pool;
}

function loadPool_() {
  const ms = master_();

  const shT = ms.getSheetByName("本文");
  const shQ = ms.getSheetByName("設問");
  if (!shT || !shQ) throw new Error("問題マスターに「本文」「設問」シートが見つかりません。");

  const tMap = headerMap_(shT);
  const qMap = headerMap_(shQ);

  const tLast = shT.getLastRow();
  const qLast = shQ.getLastRow();
  if (tLast < 2 || qLast < 2) throw new Error("問題マスターに問題が入力されていません。");

  const tRows = shT.getRange(2, 1, tLast - 1, shT.getLastColumn()).getValues();
  const qRows = shQ.getRange(2, 1, qLast - 1, shQ.getLastColumn()).getValues();

  function pick(row, map, key) {
    const c = map[key];
    return (c ? String(row[c - 1] == null ? "" : row[c - 1]) : "");
  }

  // 本文
  const stages = [];
  const byId = {};
  tRows.forEach(function (row) {
    const id = pick(row, tMap, "本文ID").trim();
    if (!id) return;
    const pub = pick(row, tMap, "公開").trim().toUpperCase();
    if (pub === "FALSE") return;
    const ord = Number(pick(row, tMap, "表示順")) || 9999;
    const st = {
      id: id,
      src: pick(row, tMap, "出典"),
      title: pick(row, tMap, "タイトル"),
      lead: pick(row, tMap, "リード文"),
      honbun: pick(row, tMap, "本文HTML"),
      chu: pick(row, tMap, "注"),
      yaku: pick(row, tMap, "現代語訳"),
      ord: ord,
      qs: []
    };
    stages.push(st);
    byId[id] = st;
  });

  // 設問
  const MARK = ["ア", "イ", "ウ", "エ"];
  qRows.forEach(function (row) {
    const qid = pick(row, qMap, "問題ID").trim();
    const sid = pick(row, qMap, "本文ID").trim();
    if (!qid || !sid || !byId[sid]) return;

    const opts = [
      pick(row, qMap, "選択肢ア"), pick(row, qMap, "選択肢イ"),
      pick(row, qMap, "選択肢ウ"), pick(row, qMap, "選択肢エ")
    ];
    if (opts.some(function (o) { return String(o).trim() === ""; })) return; // 空欄があれば出題しない

    const ansMark = pick(row, qMap, "正解").trim();
    const a = MARK.indexOf(ansMark);
    if (a < 0) return;

    byId[sid].qs.push({
      id: qid,
      n: pick(row, qMap, "問番号") || "問",
      type: pick(row, qMap, "設問タイプ") || "内容理解",
      q: pick(row, qMap, "設問文"),
      talk: pick(row, qMap, "話し合い文"),
      opts: opts,
      a: a,
      exp: pick(row, qMap, "解説"),
      pt: Number(pick(row, qMap, "配点")) || 3
    });
  });

  const usable = stages.filter(function (s) { return s.qs.length > 0; });
  usable.sort(function (x, y) { return x.ord - y.ord || (x.id < y.id ? -1 : 1); });
  return usable;
}

/* ============================================================
   進捗（自己ベスト・設問タイプ別の正答率・まちがえた問題）
   ============================================================ */
function getProgress_(no) {
  const prog = { best: {}, stats: {}, wrong: [], totalScore: 0, tries: 0, rankName: RANKS[RANKS.length - 1].name };
  Q_TYPES.forEach(function (t) { prog.stats[t] = { c: 0, t: 0 }; });

  const ss = book_();

  // 結果シートから自己ベストと合計スコア
  const shR = ss.getSheetByName(SH_RESULT);
  if (shR && shR.getLastRow() >= 2) {
    const m = headerMap_(shR);
    const rows = shR.getRange(2, 1, shR.getLastRow() - 1, shR.getLastColumn()).getValues();
    rows.forEach(function (r) {
      if (String(r[m["出席番号"] - 1]).trim() !== String(no)) return;
      const sid = String(r[m["ステージID"] - 1]).trim();
      const sc  = Number(r[m["得点"] - 1]) || 0;
      if (prog.best[sid] == null || sc > prog.best[sid]) prog.best[sid] = sc;
      prog.totalScore += Number(r[m["獲得スコア"] - 1]) || 0;
      prog.tries++;
    });
  }

  // 解答履歴シートから設問タイプ別の正答率と、いま間違えたままの問題
  const shA = ss.getSheetByName(SH_ANSWER);
  if (shA && shA.getLastRow() >= 2) {
    const m = headerMap_(shA);
    const rows = shA.getRange(2, 1, shA.getLastRow() - 1, shA.getLastColumn()).getValues();
    const latest = {};   // 問題ID → 直近の正誤
    rows.forEach(function (r) {
      if (String(r[m["出席番号"] - 1]).trim() !== String(no)) return;
      const type = String(r[m["設問タイプ"] - 1]).trim();
      const ok   = String(r[m["正誤"] - 1]).trim() === "○";
      if (prog.stats[type]) { prog.stats[type].t++; if (ok) prog.stats[type].c++; }
      latest[String(r[m["問題ID"] - 1]).trim()] = ok;   // 後の行ほど新しい
    });
    Object.keys(latest).forEach(function (qid) { if (!latest[qid]) prog.wrong.push(qid); });
  }

  prog.rankName = rankName_(prog.totalScore);
  return prog;
}

function rankName_(score) {
  for (let i = 0; i < RANKS.length; i++) if (score >= RANKS[i].min) return RANKS[i].name;
  return RANKS[RANKS.length - 1].name;
}

/* ============================================================
   採点結果の保存（解答履歴の記録はここ）
   ============================================================ */
function submitResult(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    return { ok: false, message: "混み合っています。少し待ってからもう一度お試しください。" };
  }
  try {
    const p = payload || {};
    const no = String(p.no || "").trim();
    if (!no) return { ok: false, message: "ログインし直してください。" };

    // 名簿で本人確認（クライアントからの氏名は信用しない）
    const who = findStudent_(no);
    if (!who) return { ok: false, message: "名簿に見つかりません。先生にお知らせください。" };

    const isTeacher = (no === TEACHER_NO);
    const blocked = timeBlocked_(isTeacher);
    if (blocked) return { ok: false, message: blocked };

    // 問題プールで答え合わせ（得点はサーバー側で計算し直す）
    const stages = getStages_();
    let stage = null;
    stages.forEach(function (s) { if (s.id === String(p.stageId)) stage = s; });
    if (!stage) return { ok: false, message: "ステージが見つかりません。画面を再読み込みしてください。" };

    const picks = p.picks || [];
    const secs  = Math.max(0, Math.round(Number(p.seconds) || 0));
    const blur  = Math.max(0, Math.round(Number(p.blur) || 0));

    let score = 0, fullScore = 0, correct = 0;
    const wrongIds = [], rightIds = [], answerRows = [], detail = [];
    const resultUuid = uuid_();
    const stamp = now_();

    stage.qs.forEach(function (q, i) {
      const mine = (picks[i] == null) ? -1 : Number(picks[i]);
      const ok = (mine === q.a);
      fullScore += q.pt;
      if (ok) { score += q.pt; correct++; rightIds.push(q.id); } else { wrongIds.push(q.id); }
      detail.push({ id: q.id, mine: mine, a: q.a, ok: ok });
      answerRows.push([
        uuid_(), resultUuid, stamp, no, who.name, stage.id, q.id, q.n,
        q.type, markOf_(mine), markOf_(q.a), ok ? "○" : "×", q.pt, secs
      ]);
    });

    const rate = fullScore ? Math.round(score / fullScore * 100) : 0;

    // スコア計算
    const cfg = readConfig_();
    const basePt  = correct * 100;
    const timePt  = (secs > 0 && secs < TARGET_SEC) ? Math.round((TARGET_SEC - secs) * 0.5) : 0;
    const fullPt  = (score === fullScore) ? 200 : 0;
    const mult    = cfg.classBonus ? 3 : 1;
    // 極端に速い＝まともに読んでいない、とみなしタイム点を無効化
    const suspect = (secs > 0 && secs < 20) ? "要確認" : "";
    const gained  = (basePt + (suspect ? 0 : timePt) + fullPt) * mult;

    // 重複排除（同じ提出の二重書き込みを防ぐ）
    const dedupe = [no, stage.id, String(p.startedAt || ""), String(score)].join("|");
    const shR = sheet_(SH_RESULT, H_RESULT);
    if (dedupe && alreadySaved_(shR, dedupe)) {
      return { ok: true, duplicated: true, score: score, fullScore: fullScore, detail: detail,
               progress: getProgress_(no) };
    }

    const prevBest = bestOf_(no, stage.id);
    const isBest = (prevBest == null || score > prevBest);

    shR.appendRow([
      resultUuid, stamp, no, who.name, stage.id, stage.title, stage.src,
      score, fullScore, correct, stage.qs.length, rate, secs,
      basePt, (suspect ? 0 : timePt), fullPt, mult, gained, isBest ? "○" : "",
      wrongIds.join(","), rightIds.join(","), blur, suspect, dedupe, CLIENT_VERSION
    ]);

    if (answerRows.length) {
      const shA = sheet_(SH_ANSWER, H_ANSWER);
      shA.getRange(shA.getLastRow() + 1, 1, answerRows.length, H_ANSWER.length).setValues(answerRows);
    }

    updateTotalsFor_(no);

    const prog = getProgress_(no);
    return {
      ok: true,
      resultUuid: resultUuid,
      score: score, fullScore: fullScore, correct: correct, rate: rate,
      basePt: basePt, timePt: (suspect ? 0 : timePt), fullPt: fullPt, mult: mult, gained: gained,
      isBest: isBest, prevBest: prevBest,
      detail: detail,
      progress: prog,
      rankName: prog.rankName,
      classBonus: cfg.classBonus
    };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function markOf_(i) { return ["ア", "イ", "ウ", "エ"][i] || "―"; }

function findStudent_(no) {
  const sh = book_().getSheetByName(SH_ROSTER);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(no)) return { no: String(no), name: String(rows[i][2]).trim() };
  }
  return null;
}

function alreadySaved_(sh, dedupe) {
  if (sh.getLastRow() < 2) return false;
  const m = headerMap_(sh);
  const col = m["重複排除キー"];
  if (!col) return false;
  const vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === dedupe) return true;
  return false;
}

function bestOf_(no, stageId) {
  const sh = book_().getSheetByName(SH_RESULT);
  if (!sh || sh.getLastRow() < 2) return null;
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  let best = null;
  rows.forEach(function (r) {
    if (String(r[m["出席番号"] - 1]).trim() !== String(no)) return;
    if (String(r[m["ステージID"] - 1]).trim() !== String(stageId)) return;
    const sc = Number(r[m["得点"] - 1]) || 0;
    if (best == null || sc > best) best = sc;
  });
  return best;
}

/* ============================================================
   集計・ランキング
   ============================================================ */
function collectTotals_() {
  const sh = book_().getSheetByName(SH_RESULT);
  const agg = {};
  if (!sh || sh.getLastRow() < 2) return agg;
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  rows.forEach(function (r) {
    const no = String(r[m["出席番号"] - 1]).trim();
    if (!no) return;
    if (!agg[no]) {
      agg[no] = { no: no, name: String(r[m["氏名"] - 1]).trim(), score: 0, tries: 0,
                  best: {}, rateSum: 0, maxScore: 0, fastest: null, last: "" };
    }
    const a = agg[no];
    a.name    = String(r[m["氏名"] - 1]).trim() || a.name;
    a.score  += Number(r[m["獲得スコア"] - 1]) || 0;
    a.tries  += 1;
    a.rateSum += Number(r[m["正答率(%)"] - 1]) || 0;
    const sc = Number(r[m["得点"] - 1]) || 0;
    const sid = String(r[m["ステージID"] - 1]).trim();
    if (a.best[sid] == null || sc > a.best[sid]) a.best[sid] = sc;
    if (sc > a.maxScore) a.maxScore = sc;
    const secs = Number(r[m["所要時間(秒)"] - 1]) || 0;
    if (secs > 0 && (a.fastest == null || secs < a.fastest)) a.fastest = secs;
    a.last = String(r[m["保存日時"] - 1]);
  });
  return agg;
}

function rebuildTotals() {
  const agg = collectTotals_();
  const sh = sheet_(SH_TOTAL, H_TOTAL);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, H_TOTAL.length).clearContent();

  const list = Object.keys(agg).map(function (k) { return agg[k]; });
  list.sort(function (x, y) { return y.score - x.score; });

  const rows = list.map(function (a, i) {
    const cleared = Object.keys(a.best).length;
    let fullStages = 0;
    Object.keys(a.best).forEach(function (sid) { if (a.best[sid] >= 12) fullStages++; });
    return [i + 1, a.no, a.name, a.score, a.tries, cleared, fullStages,
            a.tries ? Math.round(a.rateSum / a.tries) : 0,
            a.maxScore, a.fastest == null ? "" : a.fastest, rankName_(a.score), a.last];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, H_TOTAL.length).setValues(rows);

  const ui = tryUi_();
  if (ui) ui.alert(APP_NAME, "集計を再計算しました（" + rows.length + "名）。", ui.ButtonSet.OK);
  return rows.length;
}

/** 1人分だけ更新（保存のたびに全件書き直さない） */
function updateTotalsFor_(no) {
  const agg = collectTotals_();
  const sh = sheet_(SH_TOTAL, H_TOTAL);
  const list = Object.keys(agg).map(function (k) { return agg[k]; });
  list.sort(function (x, y) { return y.score - x.score; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, H_TOTAL.length).clearContent();
  const rows = list.map(function (a, i) {
    const cleared = Object.keys(a.best).length;
    let fullStages = 0;
    Object.keys(a.best).forEach(function (sid) { if (a.best[sid] >= 12) fullStages++; });
    return [i + 1, a.no, a.name, a.score, a.tries, cleared, fullStages,
            a.tries ? Math.round(a.rateSum / a.tries) : 0,
            a.maxScore, a.fastest == null ? "" : a.fastest, rankName_(a.score), a.last];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, H_TOTAL.length).setValues(rows);
}

/** 画面に出すランキング（上位20名＋自分） */
function getRanking(no) {
  try {
    const agg = collectTotals_();
    const list = Object.keys(agg).map(function (k) { return agg[k]; });
    list.sort(function (x, y) { return y.score - x.score; });

    const top = [], mine = { rank: null, score: 0 };
    list.forEach(function (a, i) {
      if (String(a.no) === String(no)) { mine.rank = i + 1; mine.score = a.score; }
      if (i < 20) top.push({ rank: i + 1, no: a.no, name: a.name, score: a.score, rankName: rankName_(a.score) });
    });
    return { ok: true, top: top, mine: mine, total: list.length };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/** 自分の解答履歴（新しい順） */
function getMyHistory(no, limit) {
  try {
    const sh = book_().getSheetByName(SH_RESULT);
    const out = [];
    if (!sh || sh.getLastRow() < 2) return { ok: true, rows: out };
    const m = headerMap_(sh);
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    rows.forEach(function (r) {
      if (String(r[m["出席番号"] - 1]).trim() !== String(no)) return;
      out.push({
        at: String(r[m["保存日時"] - 1]),
        stage: String(r[m["ステージ名"] - 1]),
        score: Number(r[m["得点"] - 1]) || 0,
        full: Number(r[m["満点"] - 1]) || 0,
        secs: Number(r[m["所要時間(秒)"] - 1]) || 0,
        gained: Number(r[m["獲得スコア"] - 1]) || 0
      });
    });
    out.reverse();
    return { ok: true, rows: out.slice(0, Number(limit) || 30) };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/* ============================================================
   疑義報告
   ============================================================ */
function reportIssue(payload) {
  try {
    const p = payload || {};
    const no = String(p.no || "").trim();
    const who = findStudent_(no);
    if (!who) return { ok: false, message: "ログインし直してください。" };

    const stages = getStages_();
    let q = null;
    stages.forEach(function (s) {
      s.qs.forEach(function (x) { if (x.id === String(p.qid)) q = x; });
    });
    if (!q) return { ok: false, message: "問題が見つかりませんでした。" };

    const sh = sheet_(SH_ISSUE, H_ISSUE);
    sh.appendRow([
      uuid_(), now_(), no, who.name, String(p.resultUuid || ""), q.id,
      stripTags_(q.q), q.opts[0], q.opts[1], q.opts[2], q.opts[3],
      markOf_(q.a), markOf_(Number(p.mine)),
      String(p.kind || "その他"), String(p.memo || ""), "未確認", "", "", ""
    ]);
    return { ok: true, message: "報告を受け付けました。ありがとうございます。先生が確認します。" };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function stripTags_(s) { return String(s || "").replace(/<[^>]*>/g, ""); }

function showIssueSummary() {
  const ui = tryUi_();
  if (!ui) return;
  const sh = book_().getSheetByName(SH_ISSUE);
  if (!sh || sh.getLastRow() < 2) { ui.alert(APP_NAME, "疑義報告はまだありません。", ui.ButtonSet.OK); return; }
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const byQ = {};
  let open = 0;
  rows.forEach(function (r) {
    const qid = String(r[m["問題ID"] - 1]);
    const st  = String(r[m["ステータス"] - 1]).trim();
    byQ[qid] = (byQ[qid] || 0) + 1;
    if (st === "未確認" || st === "") open++;
  });
  const lines = Object.keys(byQ).sort(function (a, b) { return byQ[b] - byQ[a]; })
    .slice(0, 20).map(function (k) { return k + " … " + byQ[k] + "件"; });
  ui.alert(APP_NAME + " 疑義報告",
    "報告 " + rows.length + "件（未確認 " + open + "件）\n\n報告の多い問題:\n" + lines.join("\n"),
    ui.ButtonSet.OK);
}

/* ============================================================
   接続確認（教員用・エディタから実行）
   ============================================================ */
function testConnection() {
  const out = [];
  try { out.push("集計簿: " + book_().getName()); } catch (e) { out.push("集計簿: NG " + e.message); }
  try {
    const pool = loadPool_();
    let n = 0; pool.forEach(function (s) { n += s.qs.length; });
    out.push("問題マスター: " + master_().getName() + " / " + pool.length + "ステージ・" + n + "問");
  } catch (e) { out.push("問題マスター: NG " + e.message); }
  Logger.log(out.join("\n"));
  return out.join("\n");
}
