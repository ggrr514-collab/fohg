/**
 * ============================================================
 *  古典クエスト ― サーバー側（Google Apps Script）
 *  埼玉県公立高校入試 大問4（古文）対策
 *
 *  ログイン方式
 *   ・Google アカウントのメールアドレスで自動ログイン
 *   ・「名簿」シート（email / 氏名 / 組 / 番号）を参照する
 *   ・名簿に無い場合だけ、初回に 18s から始まる番号と氏名を入力 →
 *     名簿へ追記し、以後は自動ログイン
 *
 *  ★ 集計側のシートは「古典_」で始まる名前にしてあります。
 *    文章読解クエストなど既存の集計スプレッドシートを
 *    そのまま SPREADSHEET_ID に指定して相乗りできます
 *    （その場合「名簿」は既存のものがそのまま使われます）。
 * ============================================================
 */

const SPREADSHEET_ID = "1Yuf_jzWZYfQKaYuZV6LN6-RDAODoOLeX-uyn9nYecwU";
const MASTER_ID      = "1OaMsGfk_-s-BMa_osO3d8hwK04ikP0G34J7TWbs2lJo";

/** index.html 側の CLIENT_VERSION と必ず同じ値にすること */
const CLIENT_VERSION = 14;

const APP_NAME = "古典クエスト";

/** 学校のメールドメインと、生徒アカウントの接頭辞 */
const MAIL_DOMAIN     = "saitama-city.ed.jp";
const STUDENT_PREFIX  = "18s";          // 例) 18s10659@saitama-city.ed.jp
/** 教員アカウントのローカル部（t15434@… など）。
 *  例外的に教員扱いしたいアドレスは TEACHER_EMAILS に追加する。 */
const TEACHER_LOCAL   = /^t\d+$/i;
const TEACHER_EMAILS  = [];

/** Googleアカウントが取得できないときに、番号の手入力でのログインを許すか。
 *  true にしても、名簿にすでにある番号でしかログインできない（新規登録は不可）。
 *  ※ 本来は「アクセスできるユーザー：（学校ドメイン）内の全員」でデプロイすれば
 *    メールが取得でき、この設定は不要です。 */
const ALLOW_MANUAL_LOGIN = false;

const PREFIX    = "古典_";
const SH_ROSTER = "名簿";               // 既存アプリと共用（接頭辞なし）
const SH_RESULT = PREFIX + "結果";
const SH_ANSWER = PREFIX + "解答履歴";
const SH_TOTAL  = PREFIX + "集計";
const SH_ISSUE  = PREFIX + "疑義報告";
const SH_CONFIG = PREFIX + "設定";
const SH_QUIZ   = PREFIX + "小テスト";
const SH_ROUND  = PREFIX + "小テスト回";

const CACHE_KEY   = "kobun_pool_v2";
const CACHE_SEC   = 21600;              // 問題プールのキャッシュ 6時間
const TARGET_SEC  = 300;                // 1ステージの目安時間（5分・画面表示用）
const QUIZ_SEC    = 300;                // 小テストの制限時間（5分）。先生が「開始」を押した時刻から数える
const QUIZ_GRACE  = 45;                 // 締切ぎりぎりに押した解答を受け取るための余裕（秒）
const QUIZ_COUNT  = 2;                  // 1回の小テストで使う大問の数（クラス全員に同じものを出す）

/* --- 連打・速すぎる解答を成績から外すための目安 ---
 * 「速すぎる」かつ「正答率が高い」回だけを成績に入れません。
 * 速くて正答率が低い回はそのまま成績になるので、
 * わざと連打して悪い点を無効にする、ということはできません。 */
const FAIR_MIN_SEC_PER_Q = 5;    // 1問あたりの平均がこれ未満なら「速すぎる」（秒）
const FAIR_MIN_GAP_SEC   = 1.5;  // 解答と解答の間隔の中央値がこれ未満なら「連打」（秒）
const FAIR_SUSPECT_RATE  = 75;   // 上にあてはまり、かつ成績がこの％以上なら成績に入れない
const PLACEHOLDER = "ここに";

/** 称号（累計得点が満点の何％かで決まる・上が上位） */
const RANKS = [
  { ratio: 0.90, name: "古典マスター" },
  { ratio: 0.70, name: "古典の師範"   },
  { ratio: 0.50, name: "助動詞の達人" },
  { ratio: 0.30, name: "仮名遣い名人" },
  { ratio: 0.10, name: "説話の読み手" },
  { ratio: 0.00, name: "古文の旅人"   }
];

const Q_TYPES = ["仮名遣い", "主語・指示語", "内容理解", "要旨・話し合い"];

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
/**
 * 見出し名で1行追加する。既存シートに無い見出しは右端に足すので、
 * あとから列が増えても、前からある列がずれない。
 */
function appendByHeader_(sh, headers, obj) {
  let m = headerMap_(sh);
  const missing = headers.filter(function (h) { return !m[h]; });
  if (missing.length) {
    const from = sh.getLastColumn() + 1;
    const style = sh.getRange(1, 1);
    sh.getRange(1, from, 1, missing.length).setValues([missing])
      .setFontWeight("bold")
      .setBackground(style.getBackground())
      .setFontColor(style.getFontColor());
    m = headerMap_(sh);
  }
  const width = sh.getLastColumn();
  const row = new Array(width).fill("");
  Object.keys(obj).forEach(function (k) {
    const c = m[k];
    if (c) row[c - 1] = obj[k];
  });
  sh.appendRow(row);
}

function headerMap_(sh) {
  const last = sh.getLastColumn();
  if (last < 1) return {};
  const head = sh.getRange(1, 1, 1, last).getValues()[0];
  const map = {};
  head.forEach(function (h, i) { if (h !== "") map[String(h).trim()] = i + 1; });
  return map;
}
function uuid_()  { return Utilities.getUuid(); }
function now_()   { return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"); }
function tryUi_() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } }

/* ============================================================
   シート定義
   ============================================================ */
/** 名簿シートを新しく作るときの見出し（既存シートがあればそちらの見出しに合わせる） */
const H_ROSTER = ["出席番号", "名前", "アドレス"];

/** 名簿の見出しのゆれを吸収する。先に完全一致、無ければ部分一致で探す。 */
const ROSTER_ALIASES = {
  email: ["email", "Email", "EMAIL", "mail", "Mail", "メール", "メールアドレス", "アドレス", "Gmail"],
  name:  ["氏名", "名前", "生徒名", "なまえ", "Name"],
  klass: ["組", "クラス", "学級", "Class"],
  no:    ["出席番号", "番号", "出席No", "No", "no"]
};

const H_RESULT = ["UUID","保存日時","email","氏名","組","番号","ステージID","ステージ名","出典",
                  "得点","満点","正答数","問題数","正答率(%)","所要時間(秒)","自己ベスト更新",
                  "誤答問題ID","正答問題ID","フォーカス離脱回数","不正疑い","重複排除キー","クライアント版"];

const H_ANSWER = ["UUID","結果UUID","保存日時","email","氏名","組","番号","ステージID","問題ID","問番号",
                  "設問タイプ","生徒の解答","正解","正誤","配点","所要時間(秒)"];

const H_TOTAL  = ["順位","email","氏名","組","番号","累計得点","満点","挑戦回数","クリアステージ数","満点ステージ数",
                  "平均正答率(%)","最高得点","最速タイム(秒)","称号",
                  "小テスト・解いた大問数","小テスト最高成績(%)","小テスト平均成績(%)","成績に入れなかった大問数","最終挑戦日時"];

const H_QUIZ   = ["UUID","保存日時","email","氏名","組","番号","回ID","何問目",
                  "成績(%)","成績に反映","除外理由","ステージID","ステージ名",
                  "問題数","正答数","所要時間(秒)","1問あたり(秒)","解答間隔の中央値(秒)","時間切れ",
                  "誤答問題ID","正答問題ID","フォーカス離脱回数","クライアント版",
                  "キー入力回数","不正の疑い"];

/** 小テストの「回」。1回＝クラス全員に同じ大問を出す1コマぶん */
const H_ROUND  = ["回ID","開始日時","終了予定","終了日時","状態","ステージID1","ステージ名1",
                  "ステージID2","ステージ名2","開始した先生","開始ミリ秒"];
const ROUND_OPEN = "実施中";
const ROUND_DONE = "終了";

const H_ISSUE  = ["UUID","報告日時","email","氏名","結果UUID","問題ID","設問文",
                  "選択肢ア","選択肢イ","選択肢ウ","選択肢エ","正解","生徒の回答",
                  "疑義種別","自由記述","ステータス","承認日時","メモ"];

/* ============================================================
   メールアドレスの扱い
   ============================================================ */
/** 「18s10659」「18S10659」「18s10659@saitama-city.ed.jp」→ 正規化したメール */
function normalizeEmail_(input) {
  let s = String(input || "").trim().toLowerCase().replace(/\s/g, "");
  if (!s) return "";
  if (s.indexOf("@") >= 0) return s;
  return s + "@" + MAIL_DOMAIN;
}
function localPart_(email) {
  const s = String(email || "");
  const i = s.indexOf("@");
  return (i < 0 ? s : s.slice(0, i)).toLowerCase();
}
function isStudentEmail_(email) {
  const local = localPart_(email);
  return local.indexOf(STUDENT_PREFIX) === 0 && /\d/.test(local.slice(STUDENT_PREFIX.length));
}
function isTeacher_(email) {
  const e = String(email || "").toLowerCase();
  if (!e) return false;
  if (TEACHER_EMAILS.indexOf(e) >= 0) return true;
  return TEACHER_LOCAL.test(localPart_(e));
}
/** このアプリで受け付けてよいアドレスか（生徒 or 教員） */
function isKnownFormat_(email) {
  return isStudentEmail_(email) || isTeacher_(email);
}
/** 実際に「誰として扱うか」を決める。クライアントの自己申告は原則使わない。 */
function resolveEmail_(claimed) {
  const signed = activeEmail_();
  if (signed) return signed;
  if (!ALLOW_MANUAL_LOGIN) return "";
  const e = String(claimed || "").trim().toLowerCase();
  return findByEmail_(e) ? e : "";       // 名簿にある番号のみ。新規は作らせない
}
/** ログイン中の Google アカウント。取得できないときは空文字 */
function activeEmail_() {
  let e = "";
  try { e = Session.getActiveUser().getEmail() || ""; } catch (err) { e = ""; }
  if (!e) { try { e = Session.getEffectiveUser().getEmail() || ""; } catch (err) { e = ""; } }
  return String(e).trim().toLowerCase();
}

/* ============================================================
   名簿
   ============================================================ */
function rosterSheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(SH_ROSTER);
  if (!sh) sh = sheet_(SH_ROSTER, H_ROSTER);
  return sh;
}

/**
 * 名簿の列位置を見出しから判定する。見つからない列は 0（＝無い）。
 * 例）「出席番号 / 名前 / アドレス」でも「email / 氏名 / 組 / 番号」でも動く。
 */
function rosterCols_(sh) {
  const m = headerMap_(sh);
  const keys = Object.keys(m);

  function exact(list) {
    for (let i = 0; i < list.length; i++) if (m[list[i]]) return m[list[i]];
    return 0;
  }
  function loose(re, exclude) {
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (exclude && exclude.indexOf(m[k]) >= 0) continue;
      if (re.test(k)) return m[k];
    }
    return 0;
  }

  const cols = {
    email: exact(ROSTER_ALIASES.email),
    name:  exact(ROSTER_ALIASES.name),
    klass: exact(ROSTER_ALIASES.klass),
    no:    exact(ROSTER_ALIASES.no)
  };
  const used = [cols.email, cols.name, cols.klass, cols.no].filter(function (c) { return c; });
  if (!cols.email) cols.email = loose(/mail|メール|アドレス|@/i, used);
  if (!cols.name)  cols.name  = loose(/氏名|名前|name/i, used);

  // 見出しが1つも当たらない古い形式（email / 氏名 / 組 / 番号 の並び）への保険
  if (!cols.email && !cols.name && sh.getLastColumn() >= 2) { cols.email = 1; cols.name = 2; }
  return cols;
}

/** 名簿を email で引く。無ければ null */
/* 名簿は毎回読むと重いので、しばらく覚えておく（登録があれば捨てる） */
const ROSTER_CACHE_KEY = "kobun_roster_v1";

function rosterMap_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(ROSTER_CACHE_KEY);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  const map = {};
  const sh = rosterSheet_();
  if (sh.getLastRow() >= 2) {
    const c = rosterCols_(sh);
    if (c.email) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      function cell(row, col) { return col ? String(row[col - 1] == null ? "" : row[col - 1]).trim() : ""; }
      rows.forEach(function (row, i) {
        const e = cell(row, c.email).toLowerCase();
        if (!e) return;
        if (!map[e]) map[e] = { email: e, name: cell(row, c.name), klass: cell(row, c.klass),
                                no: cell(row, c.no), row: i + 2 };
      });
    }
  }
  try { cache.put(ROSTER_CACHE_KEY, JSON.stringify(map), CACHE_SEC); } catch (e) {}
  return map;
}
function clearRosterCache_() {
  try { CacheService.getScriptCache().remove(ROSTER_CACHE_KEY); } catch (e) {}
}

function findByEmail_(email) {
  const e = String(email || "").trim().toLowerCase();
  if (!e) return null;
  const hit = rosterMap_()[e];
  if (!hit) return null;
  return { email: e, name: hit.name, klass: hit.klass, no: hit.no, row: hit.row,
           teacher: isTeacher_(e) };
}

/** 氏名の表記ゆれを吸収して比べるためのキー（空白を除く） */
function nameKey_(s) { return String(s || "").replace(/[\s\u3000]/g, ""); }

/**
 * 名簿へ登録する。
 *  1) すでに同じメールの行があれば、空いている項目だけ補う
 *  2) 氏名が一致していてメールだけ空の行があれば、その行にメールを入れる
 *     （名簿にいるがアカウント未登録の生徒を、行を増やさずに拾う）
 *  3) どちらでもなければ最終行に追加する
 * 実在する列にしか書き込まないので、他の列を壊さない。
 */
function upsertRoster_(email, name, klass, no) {
  const sh = rosterSheet_();
  const c = rosterCols_(sh);
  if (!c.email) {
    throw new Error("名簿シートに「アドレス」（またはemail）の列が見つかりません。見出し行をご確認ください。");
  }

  function put(row, col, value) {
    if (col && value !== "" && value != null) sh.getRange(row, col).setValue(value);
  }

  // 1) すでにいる
  const found = findByEmail_(email);
  if (found) {
    if (!found.name  && name)  put(found.row, c.name,  name);
    if (!found.klass && klass) put(found.row, c.klass, klass);
    if (!found.no    && no)    put(found.row, c.no,    no);
    return findByEmail_(email);
  }

  // 2) 氏名が同じでメールだけ空の行を探す
  if (c.name && sh.getLastRow() >= 2) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const key = nameKey_(name);
    let hit = 0, count = 0;
    for (let i = 0; i < rows.length; i++) {
      const mail = String(rows[i][c.email - 1] || "").trim();
      const nm   = nameKey_(rows[i][c.name - 1]);
      if (!mail && key && nm === key) { hit = i + 2; count++; }
    }
    if (count === 1) {                       // 候補がちょうど1人のときだけ
      put(hit, c.email, email);
      if (klass) put(hit, c.klass, klass);
      if (no)    put(hit, c.no,    no);
      clearRosterCache_();
      return findByEmail_(email);
    }
  }

  // 3) 追加
  const width = Math.max(sh.getLastColumn(), c.email, c.name, c.klass, c.no, 1);
  const row = [];
  for (let i = 0; i < width; i++) row.push("");
  row[c.email - 1] = email;
  if (c.name)  row[c.name  - 1] = name;
  if (c.klass) row[c.klass - 1] = klass;
  if (c.no)    row[c.no    - 1] = no;
  sh.getRange(sh.getLastRow() + 1, 1, 1, width).setValues([row]);
  clearRosterCache_();
  return findByEmail_(email);
}

/* ============================================================
   初回セットアップ／メニュー
   ============================================================ */
function setupSheets() {
  const ui = tryUi_();
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf(PLACEHOLDER) === 0) {
    const msg = "Code.gs の SPREADSHEET_ID がプレースホルダのままです。先に集計用スプレッドシートのIDを設定してください。";
    if (ui) ui.alert(msg); else Logger.log(msg);
    return;
  }
  sheet_(SH_RESULT, H_RESULT);
  sheet_(SH_ANSWER, H_ANSWER);
  sheet_(SH_TOTAL,  H_TOTAL);
  sheet_(SH_ISSUE,  H_ISSUE);
  sheet_(SH_QUIZ,   H_QUIZ);
  sheet_(SH_ROUND,  H_ROUND);

  const ss = book_();
  if (!ss.getSheetByName(SH_ROSTER)) sheet_(SH_ROSTER, H_ROSTER);

  let cfg = ss.getSheetByName(SH_CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(SH_CONFIG);
    cfg.getRange(1, 1, 3, 2).setValues([["授業中モード", false], ["時間制限解除", false], ["小テストモード", false]]);
    cfg.getRange(1, 1, 3, 1).setFontWeight("bold");
    cfg.setColumnWidth(1, 160);
  }
  // 既存の設定シートに「小テストモード」が無ければ足す
  const cfg2 = ss.getSheetByName(SH_CONFIG);
  if (cfg2 && cfg2.getLastRow() >= 1) {
    const rows = cfg2.getRange(1, 1, cfg2.getLastRow(), 1).getValues();
    let has = false;
    rows.forEach(function (r) { if (String(r[0]).trim() === "小テストモード") has = true; });
    if (!has) {
      const at = cfg2.getLastRow() + 1;
      cfg2.getRange(at, 1).setValue("小テストモード").setFontWeight("bold");
      cfg2.getRange(at, 2).setValue(false);
    }
  }

  const msg = "セットアップが完了しました。\n\n" +
    "・ログインは Google アカウントのメールで自動判定します。\n" +
    "・「" + SH_ROSTER + "」シート（email / 氏名 / 組 / 番号）に生徒を入れておくと、そのまま自動ログインできます。\n" +
    "・名簿に無い生徒は、初回だけ " + STUDENT_PREFIX + " から始まる番号と氏名を入力すると名簿へ追記されます。";
  if (ui) ui.alert(APP_NAME, msg, ui.ButtonSet.OK); else Logger.log(msg);
}

function onOpen() {
  const ui = tryUi_();
  if (!ui) return;
  ui.createMenu("⚙ " + APP_NAME + "管理")
    .addItem("初回セットアップ（必要なシートを生成）", "setupSheets")
    .addSeparator()
    .addItem("小テスト用の問題を割り当てる（用途列を整える）", "assignQuizUse")
    .addItem("読み込みをやり直す（キャッシュをクリア）", "clearPoolCache")
    .addItem("集計を今すぐ作り直す", "rebuildTotals")
    .addSeparator()
    .addItem("小テストモードを開始", "quizModeOn")
    .addItem("小テストモードを終了", "quizModeOff")
    .addSeparator()
    .addItem("疑義報告サマリを表示", "showIssueSummary")
    .addSeparator()
    .addItem("成績に入れなかった小テストを表示", "showExcludedQuiz")
    .addToUi();
}

function quizModeOn()  { toggleQuizFromMenu_(true); }
function quizModeOff() { toggleQuizFromMenu_(false); }
function toggleQuizFromMenu_(on) {
  const ui = tryUi_();
  const r = switchQuizMode_(on, "");
  const msg = r.ok ? quizModeMessage_(r) : ("できませんでした。\n" + r.message);
  if (ui) ui.alert(APP_NAME, msg, ui.ButtonSet.OK); else Logger.log(msg);
}

/** 小テストモードの開始／終了に共通の処理 */
function switchQuizMode_(on, byEmail) {
  try {
    let round = null, reused = false, closed = null;
    if (on) {
      const r = startRound_(byEmail);
      if (!r.ok) return r;
      round = r.round; reused = r.reused;
    } else {
      closed = endRound_();
      clearMemo_();
      clearProgressCache_();       // 練習に出る本数が変わるので、覚えていた進捗を捨てる
      rebuildTotalsIfDirty_();     // 生徒を待たせずにためておいた集計を、ここで作り直す
    }
    writeConfigFlag_("小テストモード", !!on);
    return { ok: true, quizMode: !!on, round: round, reused: reused, closed: closed,
             remain: quizStages_().length };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function quizModeMessage_(r) {
  if (!r.quizMode) {
    return "小テストモードを終了しました。生徒は通常の練習に戻ります。\n\n" +
      (r.closed
        ? "この回で使った大問（" + r.closed.titles.join("／") + "）は、\n" +
          "これから練習でも解けるようになります。\n\n"
        : "") +
      "小テストに使える大問は、あと " + r.remain + " 本です。";
  }
  const end = r.round && r.round.endsAt
    ? Utilities.formatDate(new Date(r.round.endsAt), "Asia/Tokyo", "H:mm:ss")
    : "";
  return (r.reused ? "小テストモードを続けています。\n\n" : "小テストモードを開始しました。\n\n") +
    (end ? "終了時刻　" + end + "（ここから5分。クラス全員が同じ時刻に終わります）\n\n" : "") +
    "この回の大問（クラス全員が同じ問題に取り組みます）\n" +
    r.round.titles.map(function (t, i) { return "　" + (i + 1) + "つ目　" + t; }).join("\n") + "\n\n" +
    "・生徒は画面を再読み込みすると小テストに変わります。\n" +
    "・1人1回だけです。途中でやめても受け直せません。\n" +
    "・終わったら必ず「小テストモードを終了」を実行してください。\n" +
    "　終了すると、この2本は次の回には出ず、練習で解けるようになります。\n\n" +
    "次の回に使える大問は、残り " + Math.max(0, r.remain) + " 本です。";
}

/** 設定シートの1行を書き換える（無ければ足す） */
function writeConfigFlag_(key, value) {
  const ss = book_();
  let sh = ss.getSheetByName(SH_CONFIG);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_CONFIG); }
  const last = Math.max(sh.getLastRow(), 1);
  const rows = sh.getRange(1, 1, last, 1).getValues();
  let at = 0;
  rows.forEach(function (r, i) { if (String(r[0]).trim() === key) at = i + 1; });
  if (!at) { at = last + 1; sh.getRange(at, 1).setValue(key).setFontWeight("bold"); }
  sh.getRange(at, 2).setValue(value);
}

/** 用途列が無い・空の問題マスターに、練習／小テストの既定を書き込む */
const DEFAULT_USE = {
  "S01": "練習",   "S02": "小テスト", "S03": "練習",   "S04": "小テスト", "S05": "練習",
  "S06": "練習",   "S07": "小テスト", "S08": "練習",   "S09": "小テスト", "S10": "小テスト"
};

/**
 * 問題マスターの「用途」列を整える。
 * 列が無ければ作り、空欄は既定（S01…の並びで練習5本・小テスト5本）で埋める。
 * すでに入っている値は書き換えない。
 */
function assignQuizUse() {
  const ui = tryUi_();
  const r = assignQuizUse_();
  const msg = r.ok
    ? "「用途」列を整えました。\n\n" +
      "・練習　　" + r.practice + "本\n" +
      "・小テスト " + r.quiz + "本\n" +
      (r.filled ? "・今回うめた行　" + r.filled + "\n" : "・うめる行はありませんでした\n") +
      (r.unknown.length ? "・見覚えのない本文IDは交互に振りました：" + r.unknown.join("、") + "\n" : "") +
      "\n問題プールのキャッシュも消したので、次の出題から反映されます。"
    : "うまくいきませんでした。\n" + r.message;
  if (ui) ui.alert(APP_NAME, msg, ui.ButtonSet.OK); else Logger.log(msg);
  return r;
}

function assignQuizUse_() {
  try {
    const sh = master_().getSheetByName("本文");
    if (!sh) return { ok: false, message: "問題マスターに「本文」シートが見つかりません。" };
    const last = sh.getLastRow();
    if (last < 2) return { ok: false, message: "問題マスターの「本文」シートに行がありません。" };

    let m = headerMap_(sh);
    if (!m["本文ID"]) return { ok: false, message: "「本文」シートに「本文ID」の列が見つかりません。" };

    // 用途列が無ければ、右端に作る
    let col = m["用途"];
    if (!col) {
      col = sh.getLastColumn() + 1;
      const head = sh.getRange(1, col);
      head.setValue("用途").setFontWeight("bold");
      const style = sh.getRange(1, m["本文ID"]);
      head.setBackground(style.getBackground()).setFontColor(style.getFontColor());
    }

    const ids  = sh.getRange(2, m["本文ID"], last - 1, 1).getValues();
    const uses = sh.getRange(2, col, last - 1, 1).getValues();
    const unknown = [];
    let filled = 0, practice = 0, quiz = 0, odd = 0;

    for (let i = 0; i < ids.length; i++) {
      const id = String(ids[i][0]).trim();
      if (!id) continue;
      let v = String(uses[i][0]).trim();
      if (v !== "練習" && v !== "小テスト") {
        if (DEFAULT_USE[id]) {
          v = DEFAULT_USE[id];
        } else {
          // 見覚えのない本文IDは、練習と小テストが半々になるように交互に振る
          v = (odd++ % 2) ? "小テスト" : "練習";
          unknown.push(id);
        }
        uses[i][0] = v;
        filled++;
      }
      if (v === "小テスト") quiz++; else practice++;
    }
    sh.getRange(2, col, last - 1, 1).setValues(uses);

    CacheService.getScriptCache().remove(CACHE_KEY);
    clearProgressCache_();
    return { ok: true, filled: filled, practice: practice, quiz: quiz, unknown: unknown };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function clearPoolCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  clearRosterCache_();
  clearProgressCache_();
  const ui = tryUi_();
  if (ui) ui.alert(APP_NAME, "キャッシュをクリアしました。\n\n問題マスター・名簿・生徒の進捗を、次のアクセスから読み直します。\nシートを手で書き換えたあとは、これを実行してください。", ui.ButtonSet.OK);
}

/* ============================================================
   設定
   ============================================================ */
function readConfig_() {
  const cfg = { classBonus: false, timeFree: false, quizMode: false };
  const sh = book_().getSheetByName(SH_CONFIG);
  if (!sh || sh.getLastRow() < 1) return cfg;
  const v = sh.getRange(1, 1, Math.min(sh.getLastRow(), 10), 2).getValues();
  v.forEach(function (row) {
    const key = String(row[0]).trim();
    const on  = (row[1] === true || String(row[1]).toUpperCase() === "TRUE");
    if (key === "授業中モード" || key === "授業中ボーナス") cfg.classBonus = on;
    if (key === "時間制限解除")   cfg.timeFree = on;
    if (key === "小テストモード") cfg.quizMode = on;
  });
  return cfg;
}

/** 小テストモードの切り替え（教員のみ）。開始すると、その回の大問2つが決まる */
function setQuizMode(on) {
  try {
    const email = resolveEmail_(null);
    if (!email || !isTeacher_(email)) return { ok: false, message: "先生のアカウントでのみ切り替えられます。" };
    const r = switchQuizMode_(!!on, email);
    if (!r.ok) return r;
    return { ok: true, quizMode: r.quizMode, reused: r.reused, remain: r.remain,
             round: r.round ? { id: r.round.id, titles: r.round.titles, count: r.round.ids.length,
                                endsAt: r.round.endsAt || 0 } : null,
             serverNow: Date.now(),
             closed: r.closed ? { titles: r.closed.titles } : null };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/** 深夜0:00〜5:00は挑戦禁止。教員・解除設定時はバイパス */
function timeBlocked_(isTeacher) {
  if (isTeacher) return "";
  const cfg = readConfig_();
  if (cfg.timeFree || cfg.classBonus) return "";
  const h = Number(Utilities.formatDate(new Date(), "Asia/Tokyo", "H"));
  if (h >= 0 && h < 5) return "深夜0時から5時までは挑戦できません。しっかり睡眠をとりましょう。";
  return "";
}

/* ============================================================
   セッション（自動ログイン）
   ============================================================ */
function getSession() {
  clearMemo_();
  try {
    const email = activeEmail_();

    // Google アカウントが取れない＝学校アカウントでサインインしていない
    if (!email) {
      return {
        ok: true, state: "noaccount", version: CLIENT_VERSION,
        domain: MAIL_DOMAIN, prefix: STUDENT_PREFIX, allowManual: ALLOW_MANUAL_LOGIN,
        message: "学校の Google アカウント（" + STUDENT_PREFIX + "…@" + MAIL_DOMAIN + "）でログインしてから、もう一度開いてください。"
      };
    }

    const me = findByEmail_(email);

    // 名簿に無い → 初回登録へ
    if (!me) {
      const c = rosterCols_(rosterSheet_());
      if (!c.email) {
        return { ok: false, message: "名簿シートに「アドレス」（またはemail）の列が見つかりません。先生にお知らせください。" };
      }
      return {
        ok: true, state: "register", version: CLIENT_VERSION,
        email: email, suggestNo: localPart_(email),
        domain: MAIL_DOMAIN, prefix: STUDENT_PREFIX,
        fields: { klass: !!c.klass, no: !!c.no }     // 名簿にある欄だけ入力させる
      };
    }
    return sessionFor_(me);
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/** 初回登録：18s から始まる番号と氏名を受け取って名簿に追記 */
function registerStudent(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, message: "混み合っています。少し待ってからもう一度お試しください。" };
  }
  try {
    const p = payload || {};
    const name = String(p.name || "").trim();
    if (!name) return { ok: false, message: "氏名を入力してください。" };

    const signed = activeEmail_();
    const typed  = normalizeEmail_(p.number);
    if (!typed) return { ok: false, message: STUDENT_PREFIX + " から始まる番号を入力してください。" };

    if (!isKnownFormat_(typed)) {
      return { ok: false, message: "番号は " + STUDENT_PREFIX + " から始まる形（例：" + STUDENT_PREFIX + "10659）で入力してください。" };
    }

    // Googleアカウントが取れないときは、名簿を書き換えさせない
    if (!signed) {
      if (!ALLOW_MANUAL_LOGIN) {
        return { ok: false, message: "学校の Google アカウントでログインしてから開いてください。この画面からは新しく登録できません。" };
      }
      const known = findByEmail_(typed);
      if (!known) {
        return { ok: false, message: "その番号は名簿にありません。学校の Google アカウントでログインしてから開いてください。" };
      }
      return sessionFor_(known);
    }

    // サインイン中のアカウントと違う番号は登録させない（なりすまし防止）
    if (typed !== signed) {
      return {
        ok: false,
        message: "入力された番号（" + typed + "）が、いまログインしているアカウント（" + signed + "）と一致しません。" +
                 "自分の番号を入力するか、自分の学校アカウントでログインし直してください。"
      };
    }

    const me = upsertRoster_(signed, name, String(p.klass || "").trim(), String(p.no || "").trim());
    if (!me) return { ok: false, message: "名簿に登録できませんでした。先生にお知らせください。" };
    return sessionFor_(me);
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function sessionFor_(me) {
  const cfg = readConfig_();
  return {
    ok: true, state: "ready", version: CLIENT_VERSION,
    student: { email: me.email, name: me.name, klass: me.klass, no: me.no, teacher: me.teacher },
    settings: { classBonus: cfg.classBonus, timeFree: cfg.timeFree, quizMode: cfg.quizMode },
    blocked: timeBlocked_(me.teacher),
    stages: practiceStages_(),          // 小テスト用の本文は画面に渡さない
    progress: progressOf_(me.email),
    quiz: quizStatus_(me.email)
  };
}

/* ============================================================
   問題プール（問題マスターから読み込み・キャッシュ付き）
   ============================================================ */
function getStages_() {
  const cache = CacheService.getScriptCache();
  let pool = null;
  const hit = cache.get(CACHE_KEY);
  if (hit) { try { pool = JSON.parse(hit); } catch (e) { pool = null; } }
  if (!pool) {
    pool = loadPool_();
    try { cache.put(CACHE_KEY, JSON.stringify(pool), CACHE_SEC); } catch (e) {}
  }
  return withScenes_(pool);
}

/**
 * 図解データを、キャッシュの外で毎回つけ直す。
 * こうしておくと、Scenes.gs を書き換えたときに
 * キャッシュを消さなくてもすぐ反映される。
 */
function withScenes_(pool) {
  const has = (typeof scenesOf_ === "function");     // Scenes.gs を貼っていなくても動くように
  pool.forEach(function (s) {
    if (!s.scenes && has) s.scenes = scenesOf_(s.id);   // 問題マスターの「図解」列があればそちらが残る
    if (!s.scenes) s.scenes = null;
  });
  return pool;
}

function loadPool_() {
  const ms = master_();
  const shT = ms.getSheetByName("本文");
  const shQ = ms.getSheetByName("設問");
  if (!shT || !shQ) throw new Error("問題マスターに「本文」「設問」シートが見つかりません。");

  const tMap = headerMap_(shT), qMap = headerMap_(shQ);
  const tLast = shT.getLastRow(), qLast = shQ.getLastRow();
  if (tLast < 2 || qLast < 2) throw new Error("問題マスターに問題が入力されていません。");

  const tRows = shT.getRange(2, 1, tLast - 1, shT.getLastColumn()).getValues();
  const qRows = shQ.getRange(2, 1, qLast - 1, shQ.getLastColumn()).getValues();

  function pick(row, map, key) {
    const c = map[key];
    return (c ? String(row[c - 1] == null ? "" : row[c - 1]) : "");
  }

  const stages = [], byId = {};
  tRows.forEach(function (row) {
    const id = pick(row, tMap, "本文ID").trim();
    if (!id) return;
    if (pick(row, tMap, "公開").trim().toUpperCase() === "FALSE") return;
    const st = {
      id: id, src: pick(row, tMap, "出典"), title: pick(row, tMap, "タイトル"),
      lead: pick(row, tMap, "リード文"), intro: pick(row, tMap, "前書き"), honbun: pick(row, tMap, "本文HTML"),
      chu: pick(row, tMap, "注"), yaku: pick(row, tMap, "現代語訳"),
      use: (pick(row, tMap, "用途").trim() === "小テスト") ? "小テスト" : "練習",
      ord: Number(pick(row, tMap, "表示順")) || 9999, qs: []
    };
    // 登場人物とできごとの図解。問題マスターの「図解」列があればそちらを優先する
    st.scenes = parseScenes_(pick(row, tMap, "図解")) || null;
    stages.push(st); byId[id] = st;
  });

  const MARK = ["ア", "イ", "ウ", "エ"];
  qRows.forEach(function (row) {
    const qid = pick(row, qMap, "問題ID").trim();
    const sid = pick(row, qMap, "本文ID").trim();
    if (!qid || !sid || !byId[sid]) return;
    const opts = [pick(row, qMap, "選択肢ア"), pick(row, qMap, "選択肢イ"),
                  pick(row, qMap, "選択肢ウ"), pick(row, qMap, "選択肢エ")];
    if (opts.some(function (o) { return String(o).trim() === ""; })) return;
    const a = MARK.indexOf(pick(row, qMap, "正解").trim());
    if (a < 0) return;
    byId[sid].qs.push({
      id: qid, n: pick(row, qMap, "問番号") || "問",
      type: pick(row, qMap, "設問タイプ") || "内容理解",
      q: pick(row, qMap, "設問文"), blank: pick(row, qMap, "空欄文"), talk: pick(row, qMap, "話し合い文"),
      opts: opts, a: a, exp: pick(row, qMap, "解説"),
      pt: Number(pick(row, qMap, "配点")) || 3
    });
  });

  const usable = stages.filter(function (s) { return s.qs.length > 0; });
  usable.sort(function (x, y) { return x.ord - y.ord || (x.id < y.id ? -1 : 1); });
  return usable;
}

/** 問題マスターの「図解」列（JSON）を読む。空や壊れていれば null */
function parseScenes_(text) {
  const t = String(text || "").trim();
  if (!t) return null;
  try {
    const v = JSON.parse(t);
    return (v && v.people && v.events) ? v : null;
  } catch (e) { return null; }
}

/**
 * 練習で出すステージ。
 * 用途が「練習」のものに加えて、小テストで一度使い終わった大問も練習に回る。
 */
function practiceStages_() {
  const opened = finishedQuizIds_();
  return getStages_().filter(function (s) {
    return s.use !== "小テスト" || opened[s.id];
  });
}
/** 小テストにまだ使っていない大問（次の回に出せるもの） */
function quizStages_() {
  const used = usedQuizIds_();
  return getStages_().filter(function (s) {
    return s.use === "小テスト" && !used[s.id];
  });
}
/** 回で指定された大問を、順番どおりに取り出す */
function roundStages_(round) {
  const by = {};
  getStages_().forEach(function (s) { by[s.id] = s; });
  return (round.ids || []).map(function (id) { return by[id]; })
                          .filter(function (s) { return !!s; });
}
/** 採点前に画面へ渡す用。正解・解説・図解を落とす */
function hideAnswers_(stages) {
  return stages.map(function (s) {
    return {
      id: s.id, src: s.src, title: s.title, lead: s.lead, intro: s.intro,
      honbun: s.honbun, chu: s.chu, use: s.use,
      qs: s.qs.map(function (q) {
        return { id: q.id, n: q.n, type: q.type, q: q.q, blank: q.blank, talk: q.talk,
                 opts: q.opts, pt: q.pt };
      })
    };
  });
}

/* ============================================================
   小テストの「回」
   1回＝その授業の1コマ。クラス全員が同じ大問2つに取り組む。
   使い終わった大問は、次の回には出ず、練習に回る。
   ============================================================ */
/* --- 1回の呼び出しのあいだだけ覚えておく（同じシートを何度も読まないため） --- */
let _memoRounds = null;      // 小テスト回シートの中身
let _memoQuiz   = null;      // 小テストシートの中身
let _memoSheet  = {};        // その他のシートの中身（シート名ごと）
function clearMemo_() { _memoRounds = null; _memoQuiz = null; _memoSheet = {}; }

/** シートを丸ごと1回だけ読む。書き込んだあとは clearMemo_() で捨てること */
function sheetValues_(name) {
  if (_memoSheet[name]) return _memoSheet[name];
  const sh = book_().getSheetByName(name);
  let out = { map: {}, rows: [] };
  if (sh && sh.getLastRow() >= 2 && sh.getLastColumn() >= 1) {
    out = { map: headerMap_(sh),
            rows: sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() };
  }
  _memoSheet[name] = out;
  return out;
}

/** 集計シートは、生徒を待たせないように、あとでまとめて作り直す */
function markTotalsDirty_() {
  try { PropertiesService.getScriptProperties().setProperty("totalsDirty", "1"); } catch (e) {}
}
function rebuildTotalsIfDirty_() {
  let dirty = false;
  try { dirty = (PropertiesService.getScriptProperties().getProperty("totalsDirty") === "1"); } catch (e) {}
  if (!dirty) return 0;
  const n = updateTotals_();
  try { PropertiesService.getScriptProperties().deleteProperty("totalsDirty"); } catch (e) {}
  return n;
}

function roundSheet_() { return sheet_(SH_ROUND, H_ROUND); }

/** 回シートの全行を読む（古い順）。同じ実行の中では1回だけ読む */
function roundRows_() {
  if (_memoRounds) return _memoRounds;
  _memoRounds = roundRowsFresh_();
  return _memoRounds;
}
function roundRowsFresh_() {
  const sh = book_().getSheetByName(SH_ROUND);
  if (!sh || sh.getLastRow() < 2) return [];
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const out = [];
  rows.forEach(function (r, i) {
    const id = String(r[m["回ID"] - 1] || "").trim();
    if (!id) return;
    const ids = [], titles = [];
    [["ステージID1", "ステージ名1"], ["ステージID2", "ステージ名2"]].forEach(function (pair) {
      const c = m[pair[0]];
      const sid = c ? String(r[c - 1] || "").trim() : "";
      if (sid) { ids.push(sid); titles.push(m[pair[1]] ? String(r[m[pair[1]] - 1] || "") : ""); }
    });
    let startMs = m["開始ミリ秒"] ? Number(r[m["開始ミリ秒"] - 1]) : 0;
    if (!startMs) {
      // 古い記録には開始ミリ秒がないので、開始日時の文字から読み取る
      const t = Date.parse(String(r[m["開始日時"] - 1] || "").replace(/\//g, "-").replace(" ", "T") + "+09:00");
      startMs = isFinite(t) ? t : 0;
    }
    out.push({ row: i + 2, id: id, state: String(r[m["状態"] - 1] || "").trim(),
               startedAt: String(r[m["開始日時"] - 1] || ""),
               startMs: startMs, endsAt: startMs ? startMs + QUIZ_SEC * 1000 : 0,
               ids: ids, titles: titles });
  });
  return out;
}

/** いま実施中の回（無ければ null） */
function activeRound_() {
  const rows = roundRows_();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i].state === ROUND_OPEN) return rows[i];
  }
  return null;
}

/** これまでの回で使った大問ID（実施中の回も含む＝次の回には出さない） */
function usedQuizIds_() {
  const d = {};
  roundRows_().forEach(function (r) { r.ids.forEach(function (id) { d[id] = true; }); });
  return d;
}

/** 終わった回で使った大問ID（＝練習に出してよいもの） */
function finishedQuizIds_() {
  const d = {};
  roundRows_().forEach(function (r) {
    if (r.state === ROUND_DONE) r.ids.forEach(function (id) { d[id] = true; });
  });
  return d;
}

/** 新しい回を始める。未使用の大問から QUIZ_COUNT 本を選んで固定する */
function startRound_(byEmail) {
  const open = activeRound_();
  if (open) return { ok: true, round: open, reused: true };

  const pool = quizStages_().slice();
  if (!pool.length) {
    return { ok: false, message: "小テストに使える大問が残っていません。\n" +
      "問題マスターの「用途」列に「小テスト」の大問を足してください。" };
  }
  for (let i = pool.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = pool[i]; pool[i] = pool[j]; pool[j] = t;
  }
  const use = pool.slice(0, QUIZ_COUNT);
  const sh  = roundSheet_();
  const id  = uuid_();
  const startMs = Date.now();
  const endsAt  = startMs + QUIZ_SEC * 1000;
  appendByHeader_(sh, H_ROUND, {
    "回ID": id, "開始日時": now_(),
    "終了予定": Utilities.formatDate(new Date(endsAt), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"),
    "終了日時": "", "状態": ROUND_OPEN,
    "ステージID1": use[0] ? use[0].id : "", "ステージ名1": use[0] ? use[0].title : "",
    "ステージID2": use[1] ? use[1].id : "", "ステージ名2": use[1] ? use[1].title : "",
    "開始した先生": byEmail || "", "開始ミリ秒": startMs
  });
  clearMemo_();
  return { ok: true, reused: false,
    round: { id: id, state: ROUND_OPEN, startedAt: now_(), startMs: startMs, endsAt: endsAt,
             ids: use.map(function (s) { return s.id; }),
             titles: use.map(function (s) { return s.title; }) } };
}

/** 実施中の回を終わらせる。使った大問はここで練習に回る */
function endRound_() {
  const sh = book_().getSheetByName(SH_ROUND);
  if (!sh || sh.getLastRow() < 2) return null;
  const m = headerMap_(sh);
  const rows = roundRows_();
  let closed = null;
  rows.forEach(function (r) {
    if (r.state !== ROUND_OPEN) return;
    sh.getRange(r.row, m["状態"]).setValue(ROUND_DONE);
    sh.getRange(r.row, m["終了日時"]).setValue(now_());
    closed = r;
  });
  return closed;
}

/** 小テストシートを、その生徒のぶんだけ整理して返す（同じ実行の中では1回だけ読む） */
function quizLog_() {
  if (_memoQuiz) return _memoQuiz;
  const out = [];
  const sh = book_().getSheetByName(SH_QUIZ);
  if (sh && sh.getLastRow() >= 2 && sh.getLastColumn() >= 1) {
    const m = headerMap_(sh);
    if (m["email"]) {
      // いるのは前のほうの列だけなので、そこまでで読み止める
      let need = 1;
      ["保存日時", "email", "回ID", "何問目", "成績(%)", "成績に反映", "ステージID"].forEach(function (h) {
        if (m[h] && m[h] > need) need = m[h];
      });
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, need).getValues();
      rows.forEach(function (r) {
        const email = String(r[m["email"] - 1] || "").trim().toLowerCase();
        if (!email) return;
        out.push({
          email: email,
          roundId: m["回ID"] ? String(r[m["回ID"] - 1] || "").trim() : "",
          stageId: m["ステージID"] ? String(r[m["ステージID"] - 1] || "").trim() : "",
          index: m["何問目"] ? (Number(r[m["何問目"] - 1]) || 0) : 0,
          score: m["成績(%)"] ? (Number(r[m["成績(%)"] - 1]) || 0) : 0,
          at: m["保存日時"] ? String(r[m["保存日時"] - 1] || "") : "",
          counted: quizCounts_(r, m)
        });
      });
    }
  }
  _memoQuiz = out;
  return out;
}

/** この回が終わる時刻（ミリ秒）。全員同じ */
function roundEndsAt_(round) {
  return (round && round.endsAt) ? round.endsAt : 0;
}
/** 締切を過ぎたか。grace は、ぎりぎりの解答を受け取るための余裕（秒） */
function roundOver_(round, graceSec) {
  const end = roundEndsAt_(round);
  if (!end) return false;
  return Date.now() > end + (graceSec || 0) * 1000;
}

/** その生徒が、この回の小テストをもう受けたか */
function roundTakenBy_(email, roundId) {
  const key = String(email || "").toLowerCase();
  const mine = quizLog_().filter(function (r) {
    return r.email === key && r.roundId === String(roundId);
  });
  if (!mine.length) return null;
  const counted = mine.filter(function (x) { return x.counted; });
  const avg = counted.length
      ? Math.round(counted.reduce(function (a, x) { return a + x.score; }, 0) / counted.length)
      : null;
  return { parts: mine.length, avg: avg, ids: mine.map(function (x) { return x.stageId; }) };
}

/* ============================================================
   小テスト
   ============================================================ */
/** 解答の速さを測る。stamps は { 問題ID: 開始からのミリ秒 } */
function paceOf_(stamps, secs, total) {
  const ms = [];
  if (stamps) {
    Object.keys(stamps).forEach(function (k) {
      const v = Number(stamps[k]);
      if (isFinite(v) && v >= 0) ms.push(v);
    });
  }
  // 最後に押した時刻と、画面が数えた所要時間の、長いほうを使う
  let spanSec = 0;
  ms.forEach(function (v) { if (v / 1000 > spanSec) spanSec = v / 1000; });
  const effSec = Math.max(Number(secs) || 0, spanSec);
  const perQ = total ? (effSec / total) : 0;

  let medGap = null;
  if (ms.length >= 2) {
    ms.sort(function (a, b) { return a - b; });
    const gaps = [];
    let prev = 0;
    ms.forEach(function (v) { gaps.push((v - prev) / 1000); prev = v; });
    gaps.sort(function (a, b) { return a - b; });
    const mid = Math.floor(gaps.length / 2);
    medGap = (gaps.length % 2) ? gaps[mid] : (gaps[mid - 1] + gaps[mid]) / 2;
    medGap = Math.round(medGap * 10) / 10;
  }
  return { perQ: Math.round(perQ * 10) / 10, medGap: medGap, answered: ms.length,
           effSec: Math.round(effSec * 10) / 10 };
}

/**
 * 成績に入れてよい大問かどうかを判定する。
 * ・ほかの画面に移った／キーボードを使った → 不正あつかい（成績に入れない）
 * ・速すぎる、かつ正答率が高い → 成績に入れない
 */
function fairnessOf_(pace, score, answered, total, blur, keys) {
  const cheat = [];
  if (blur > 0) cheat.push("ほかの画面に移りました（" + blur + "回）");
  if (keys > 0) cheat.push("キーボードを使いました（" + keys + "回）");
  if (cheat.length) {
    return { include: false, cheated: true, reason: cheat.join("・") };
  }

  const reasons = [];
  // 未解答が多い回は、そもそも速くて当然なので対象にしない
  const fastPerQ = (answered >= 2) && (pace.perQ < FAIR_MIN_SEC_PER_Q);
  const burst    = (pace.medGap != null) && (pace.medGap < FAIR_MIN_GAP_SEC);
  if (fastPerQ) reasons.push("1問あたり" + pace.perQ + "秒");
  if (burst)    reasons.push("解答間隔の中央値" + pace.medGap + "秒");

  if ((fastPerQ || burst) && score >= FAIR_SUSPECT_RATE) {
    return { include: false, cheated: false, reason: "速すぎます（" + reasons.join("・") + "）" };
  }
  return { include: true, cheated: false, reason: "" };
}

/** 記録の1行を成績に入れてよいか。先生が「×」に書き換えた行は外れる */
function quizCounts_(row, m) {
  const c = m["成績に反映"];
  if (!c) return true;                       // 旧い記録は従来どおり数える
  const v = String(row[c - 1]).trim();
  return v !== "×" && v !== "x" && v.toUpperCase() !== "FALSE";
}

/** 小テストの状況（この回のこと＋これまでの成績） */
function quizStatus_(email) {
  const round = activeRound_();
  const taken = round ? roundTakenBy_(email, round.id) : null;

  let best = null, sum = 0, tries = 0, last = null, skipped = 0;
  const key = String(email || "").toLowerCase();
  quizLog_().forEach(function (r) {
    if (r.email !== key) return;
    if (!r.counted) { skipped++; return; }
    tries++;
    if (best == null || r.score > best) best = r.score;
    sum += r.score; last = r.score;
  });

  const allQuiz = getStages_().filter(function (s) { return s.use === "小テスト"; }).length;
  return {
    round: round ? { id: round.id, parts: round.ids.length,
                     titles: round.titles, startedAt: round.startedAt,
                     endsAt: roundEndsAt_(round), over: roundOver_(round, 0) } : null,
    serverNow: Date.now(),
    taken: taken ? { parts: taken.parts, avg: taken.avg } : null,
    remain: quizStages_().length,      // 次の回に使える大問
    total: allQuiz,                    // 小テスト用として用意してある大問
    tries: tries, skipped: skipped, best: best,
    avg: tries ? Math.round(sum / tries) : null, last: last,
    limitSec: QUIZ_SEC, maxParts: QUIZ_COUNT
  };
}

/**
 * 小テストを始める。いま実施中の回で決まっている1つ目を返す。
 * クラス全員が同じ大問に取り組み、1人1回だけ受けられる。
 */
function startQuiz() {
  clearMemo_();
  try {
    const email = resolveEmail_(null);
    const me = email ? findByEmail_(email) : null;
    if (!me) return { ok: false, message: "ログイン状態が確認できません。ページを再読み込みしてください。" };

    const round = activeRound_();
    if (!round) {
      return { ok: false, message: "いまは小テストの時間ではありません。" };
    }
    if (roundOver_(round, 0)) {
      return { ok: false, over: true,
        message: "この時間の小テストは、もう終わりました。先生の指示を待ちましょう。" };
    }
    const taken = roundTakenBy_(me.email, round.id);
    if (taken) {
      return { ok: false, taken: true,
        message: "この時間の小テストは、すでに受けています。" +
          (taken.avg != null ? "（成績 " + taken.avg + "％）" : "") };
    }
    const stages = roundStages_(round);
    if (!stages.length) {
      return { ok: false, message: "この回の問題が読み込めませんでした。先生にお知らせください。" };
    }
    return {
      ok: true,
      sessionId: round.id,
      index: 1,
      parts: stages.length,
      stage: hideAnswers_([stages[0]])[0],
      limitSec: QUIZ_SEC,
      endsAt: roundEndsAt_(round),      // クラス全員で同じ終了時刻
      serverNow: Date.now(),            // 端末の時計がずれていても合わせられるように
      status: quizStatus_(me.email)
    };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/** 「2つ目に挑戦する」を選んだときに、この回の2つ目を渡す */
function nextQuizStage(exclude) {
  clearMemo_();
  try {
    const email = resolveEmail_(null);
    const me = email ? findByEmail_(email) : null;
    if (!me) return { ok: false, message: "ログイン状態が確認できません。ページを再読み込みしてください。" };

    const round = activeRound_();
    if (!round) return { ok: false, message: "いまは小テストの時間ではありません。" };
    if (roundOver_(round, 0)) {
      return { ok: true, over: true, endsAt: roundEndsAt_(round), serverNow: Date.now() };
    }

    const done = {};
    (exclude || []).forEach(function (id) { done[String(id)] = true; });
    const next = roundStages_(round).filter(function (s) { return !done[s.id]; })[0];
    if (!next) return { ok: true, noMore: true };
    return { ok: true, stage: hideAnswers_([next])[0],
             endsAt: roundEndsAt_(round), serverNow: Date.now() };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/**
 * 大問1つ分を採点して記録する。
 * 記録した時点で、その大問の成績は確定する。
 */
function submitQuizPart(payload) {
  // クラス全員がいっせいに送るので、待たせないことを第一にする。
  // 集計シートの作り直しは後回しにし、二重記録はスクリプトプロパティで防ぐ。
  try {
    clearMemo_();
    const p = payload || {};
    const email = resolveEmail_(p.email);
    const me = email ? findByEmail_(email) : null;
    if (!me) return { ok: false, message: "ログイン状態が確認できません。ページを再読み込みしてください。" };

    const sid = String(p.stageId || "");
    const round = activeRound_();
    if (!round) return { ok: false, message: "いまは小テストの時間ではありません。" };
    if (round.ids.indexOf(sid) < 0) {
      return { ok: false, message: "この回の問題ではありません。ページを再読み込みしてください。" };
    }
    let stage = null;
    roundStages_(round).forEach(function (s) { if (s.id === sid) stage = s; });
    if (!stage) return { ok: false, message: "問題が見つかりません。もう一度やり直してください。" };

    // 同じ回の同じ大問は、1人1回だけ
    const already = roundTakenBy_(me.email, round.id);
    if (already && already.ids.indexOf(sid) >= 0) {
      return { ok: false, taken: true, message: "この大問は、すでに提出しています。" };
    }

    const picks  = p.picks || {};
    const secs   = Math.max(0, Math.round(Number(p.seconds) || 0));
    const blur   = Math.max(0, Math.round(Number(p.blur) || 0));
    const keys   = Math.max(0, Math.round(Number(p.keys) || 0));
    // 押した時点では間に合っていても、保存が遅れることがあるので少し余裕をみる
    const late   = roundOver_(round, QUIZ_GRACE);
    const timeUp = !!p.timeUp || late;
    const index  = Math.max(1, Math.round(Number(p.index) || 1));
    const sessionId = round.id;          // 回ID＝その授業の小テスト1回分

    let total = 0, correct = 0;
    const detail = [], wrongIds = [], rightIds = [];
    stage.qs.forEach(function (q) {
      total++;
      const raw = picks[q.id];
      const mine = (raw == null || raw === "") ? -1 : Number(raw);   // 未回答は不正解
      const ok = (mine === q.a);
      if (ok) { correct++; rightIds.push(q.id); } else { wrongIds.push(q.id); }
      detail.push({ stageId: sid, id: q.id, n: q.n, type: q.type,
                    mine: mine, a: q.a, ok: ok, exp: q.exp, opts: q.opts });
    });
    if (!total) return { ok: false, message: "問題が見つかりませんでした。" };

    const score = Math.round(correct / total * 100);

    // 連打・速すぎる解答は成績に入れない
    const answered = Object.keys(picks).filter(function (k) {
      return picks[k] != null && picks[k] !== "";
    }).length;
    const pace = paceOf_(p.stamps, secs, total);
    const fair = fairnessOf_(pace, score, answered, total, blur, keys);

    const sh = sheet_(SH_QUIZ, H_QUIZ);
    const dedupe = [me.email, sessionId, sid, String(index)].join("|");
    if (!alreadyQuizSaved_(sh, dedupe)) {
      appendByHeader_(sh, H_QUIZ, {
        "UUID": uuid_(), "保存日時": now_(), "email": me.email, "氏名": me.name,
        "組": me.klass, "番号": me.no, "回ID": sessionId, "何問目": index,
        "成績(%)": score, "成績に反映": fair.include ? "○" : "×", "除外理由": fair.reason,
        "ステージID": sid, "ステージ名": stage.title,
        "問題数": total, "正答数": correct, "所要時間(秒)": secs,
        "1問あたり(秒)": pace.perQ, "解答間隔の中央値(秒)": pace.medGap == null ? "" : pace.medGap,
        "時間切れ": timeUp ? "○" : "", "誤答問題ID": wrongIds.join(","), "正答問題ID": rightIds.join(","),
        "フォーカス離脱回数": blur, "キー入力回数": keys,
        "不正の疑い": fair.cheated ? "○" : "", "クライアント版": CLIENT_VERSION
      });
      markTotalsDirty_();      // 集計シートは、生徒を待たせないようにあとで作り直す
      // いま書いた1行を手元の控えにも足す（シートを読み直さずに状況を返すため）
      quizLog_().push({ email: me.email.toLowerCase(), roundId: sessionId, stageId: sid,
                        index: index, score: score, at: now_(), counted: fair.include });
    }

    const res = {
      ok: true, index: index, stageId: sid, title: stage.title, src: stage.src,
      total: total, correct: correct, score: score,
      scenes: stage.scenes || null,
      counted: fair.include, cheated: !!fair.cheated, reason: fair.reason, pace: pace,
      blur: blur, keys: keys,
      timeUp: timeUp, seconds: secs, detail: detail,
      status: quizStatus_(me.email)
    };

    return res;
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function alreadyQuizSaved_(sh, dedupe) {
  return markOnce_("qdup_", dedupe);
}

/* ============================================================
   進捗
   ============================================================ */
/* 進捗（自己ベスト・タイプ別正誤・まちがえた問題）は、
   結果シートと解答履歴シートを全部読まないと出せない。
   毎回読むと人数ぶん重くなるので、生徒ごとに覚えておき、
   解くたびに「足し算」で更新する。 */
function progressEpoch_() {
  try {
    const v = PropertiesService.getScriptProperties().getProperty("progEpoch");
    return v || "1";
  } catch (e) { return "1"; }
}
function clearProgressCache_() {
  // 生徒ごとのキーを消して回るかわりに、世代番号を1つ進める
  try {
    const props = PropertiesService.getScriptProperties();
    const v = Number(props.getProperty("progEpoch") || "1") + 1;
    props.setProperty("progEpoch", String(v));
  } catch (e) {}
}
function progKey_(email) {
  return "prog" + progressEpoch_() + "_" +
         Utilities.base64EncodeWebSafe(String(email || "").toLowerCase()).slice(0, 80);
}

/**
 * 覚えてあるものがあれば使う。無ければシートから作って覚える。
 * どのみちシートを全部読むので、そのとき「全員ぶん」まとめて作って覚えておく。
 * こうすると、クラスで最初の1人だけが読み込み、あとの人は読まずに済む。
 */
function progressCoreCached_(email) {
  const cache = CacheService.getScriptCache();
  const key = progKey_(email);
  const hit = cache.get(key);
  let core = null;
  if (hit) { try { core = JSON.parse(hit); } catch (e) { core = null; } }
  if (!core) {
    const all = progressAllCores_();
    core = all[String(email || "").toLowerCase()] || emptyCore_();
    const put = {};
    Object.keys(all).forEach(function (e) { put[progKey_(e)] = JSON.stringify(all[e]); });
    put[key] = JSON.stringify(core);          // 名簿にいない人のぶんも覚えておく
    try { cache.putAll(put, CACHE_SEC); } catch (e) {
      try { cache.put(key, JSON.stringify(core), CACHE_SEC); } catch (e2) {}
    }
  }
  if (!core.best)  core.best  = {};
  if (!core.stats) core.stats = {};
  if (!core.wrong) core.wrong = [];
  if (!core.tries) core.tries = 0;
  return core;
}
function progressOf_(email) { return finishProgress_(progressCoreCached_(email)); }
function saveProgressCore_(email, core) {
  try { CacheService.getScriptCache().put(progKey_(email), JSON.stringify(core), CACHE_SEC); } catch (e) {}
}

/** 満点・累計・称号は、そのつど計算し直す（練習の本数が増えると変わるため） */
function finishProgress_(core) {
  const prog = { best: core.best || {}, stats: core.stats || {}, wrong: core.wrong || [],
                 tries: core.tries || 0, points: 0, maxPoints: 0, rankName: "" };
  Q_TYPES.forEach(function (t) { if (!prog.stats[t]) prog.stats[t] = { c: 0, t: 0 }; });
  Object.keys(prog.best).forEach(function (sid) { prog.points += prog.best[sid]; });
  prog.maxPoints = maxPoints_();
  prog.rankName = rankName_(prog.points, prog.maxPoints);
  return prog;
}

function emptyCore_() {
  const stats = {};
  Q_TYPES.forEach(function (t) { stats[t] = { c: 0, t: 0 }; });
  return { best: {}, stats: stats, wrong: [], tries: 0 };
}

/** 結果シートと解答履歴シートを1回ずつ読んで、全員ぶんの進捗を作る */
function progressAllCores_() {
  const all = {};
  function core(e) {
    if (!all[e]) all[e] = emptyCore_();
    return all[e];
  }

  const vR = sheetValues_(SH_RESULT);
  if (vR.map["email"]) {
    const m = vR.map;
    vR.rows.forEach(function (r) {
      const e = String(r[m["email"] - 1] || "").trim().toLowerCase();
      if (!e) return;
      const c = core(e);
      const sid = String(r[m["ステージID"] - 1] || "").trim();
      const sc = Number(r[m["得点"] - 1]) || 0;
      if (c.best[sid] == null || sc > c.best[sid]) c.best[sid] = sc;
      c.tries++;
    });
  }

  const vA = sheetValues_(SH_ANSWER);
  if (vA.map["email"]) {
    const m = vA.map;
    const latest = {};                       // email → 問題ID → 正誤（最後の解答）
    vA.rows.forEach(function (r) {
      const e = String(r[m["email"] - 1] || "").trim().toLowerCase();
      if (!e) return;
      const c = core(e);
      const type = String(r[m["設問タイプ"] - 1] || "").trim();
      const ok = String(r[m["正誤"] - 1] || "").trim() === "○";
      if (c.stats[type]) { c.stats[type].t++; if (ok) c.stats[type].c++; }
      if (!latest[e]) latest[e] = {};
      latest[e][String(r[m["問題ID"] - 1] || "").trim()] = ok;
    });
    Object.keys(latest).forEach(function (e) {
      const c = core(e);
      Object.keys(latest[e]).forEach(function (qid) { if (!latest[e][qid]) c.wrong.push(qid); });
    });
  }
  return all;
}

function getProgress_(email) {
  return finishProgress_(progressCore_(email));
}

function progressCore_(email) {
  const key = String(email || "").toLowerCase();
  const prog = { best: {}, stats: {}, wrong: [], points: 0, maxPoints: 0, tries: 0, rankName: RANKS[RANKS.length - 1].name };
  Q_TYPES.forEach(function (t) { prog.stats[t] = { c: 0, t: 0 }; });

  const vR = sheetValues_(SH_RESULT);
  if (vR.map["email"]) {
    const m = vR.map;
    vR.rows.forEach(function (r) {
      if (String(r[m["email"] - 1]).trim().toLowerCase() !== key) return;
      const sid = String(r[m["ステージID"] - 1]).trim();
      const sc  = Number(r[m["得点"] - 1]) || 0;
      if (prog.best[sid] == null || sc > prog.best[sid]) prog.best[sid] = sc;
      prog.tries++;
    });
  }

  const vA = sheetValues_(SH_ANSWER);
  if (vA.map["email"]) {
    const m = vA.map;
    const latest = {};
    vA.rows.forEach(function (r) {
      if (String(r[m["email"] - 1]).trim().toLowerCase() !== key) return;
      const type = String(r[m["設問タイプ"] - 1]).trim();
      const ok   = String(r[m["正誤"] - 1]).trim() === "○";
      if (prog.stats[type]) { prog.stats[type].t++; if (ok) prog.stats[type].c++; }
      latest[String(r[m["問題ID"] - 1]).trim()] = ok;
    });
    Object.keys(latest).forEach(function (qid) { if (!latest[qid]) prog.wrong.push(qid); });
  }

  return { best: prog.best, stats: prog.stats, wrong: prog.wrong, tries: prog.tries };
}

function rankName_(points, maxPoints) {
  const max = Number(maxPoints) || 0;
  const r = max > 0 ? (Number(points) || 0) / max : 0;
  for (let i = 0; i < RANKS.length; i++) if (r >= RANKS[i].ratio) return RANKS[i].name;
  return RANKS[RANKS.length - 1].name;
}
/** 全ステージを満点で取ったときの点数 */
function maxPoints_() {
  let n = 0;
  practiceStages_().forEach(function (s) { s.qs.forEach(function (q) { n += (q.pt || 3); }); });
  return n;
}

/* ============================================================
   採点結果の保存
   ============================================================ */
function submitResult(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {
    return { ok: false, message: "混み合っています。少し待ってからもう一度お試しください。" };
  }
  try {
    clearMemo_();
    const p = payload || {};

    // 誰が解いたかは、クライアントの申告ではなくログイン中のアカウントで決める
    const email = resolveEmail_(p.email);
    if (!email) return { ok: false, message: "ログイン状態が確認できません。ページを再読み込みしてください。" };

    const me = findByEmail_(email);
    if (!me) return { ok: false, message: "名簿に見つかりません。ページを再読み込みして登録し直してください。" };

    const blocked = timeBlocked_(me.teacher);
    if (blocked) return { ok: false, message: blocked };

    const cfgNow = readConfig_();
    if (cfgNow.quizMode && !me.teacher) {
      return { ok: false, message: "いまは小テストの時間です。画面を再読み込みしてください。" };
    }

    let stage = null;
    practiceStages_().forEach(function (s) { if (s.id === String(p.stageId)) stage = s; });
    if (!stage) return { ok: false, message: "ステージが見つかりません。画面を再読み込みしてください。" };

    const picks = p.picks || [];
    const secs  = Math.max(0, Math.round(Number(p.seconds) || 0));
    const blur  = Math.max(0, Math.round(Number(p.blur) || 0));

    let score = 0, fullScore = 0, correct = 0;
    const wrongIds = [], rightIds = [], answerRows = [], detail = [];
    const resultUuid = uuid_(), stamp = now_();

    stage.qs.forEach(function (q, i) {
      const mine = (picks[i] == null) ? -1 : Number(picks[i]);
      const ok = (mine === q.a);
      fullScore += q.pt;
      if (ok) { score += q.pt; correct++; rightIds.push(q.id); } else { wrongIds.push(q.id); }
      detail.push({ id: q.id, mine: mine, a: q.a, ok: ok });
      answerRows.push([
        uuid_(), resultUuid, stamp, me.email, me.name, me.klass, me.no, stage.id, q.id, q.n,
        q.type, markOf_(mine), markOf_(q.a), ok ? "○" : "×", q.pt, secs
      ]);
    });

    const rate = fullScore ? Math.round(score / fullScore * 100) : 0;

    // 極端に速い提出＝まともに読んでいない、として印を付ける（得点はそのまま記録）
    const suspect = (secs > 0 && secs < 20) ? "要確認" : "";

    const dedupe = [me.email, stage.id, String(p.startedAt || ""), String(score)].join("|");
    const shR = sheet_(SH_RESULT, H_RESULT);
    if (dedupe && alreadySaved_(shR, dedupe)) {
      return { ok: true, duplicated: true, score: score, fullScore: fullScore, correct: correct, rate: rate,
               isBest: false, prevBest: bestOf_(me.email, stage.id), detail: detail,
               progress: progressOf_(me.email), resultUuid: resultUuid };
    }

    const prevBest = bestOf_(me.email, stage.id);
    const isBest = (prevBest == null || score > prevBest);

    shR.appendRow([
      resultUuid, stamp, me.email, me.name, me.klass, me.no, stage.id, stage.title, stage.src,
      score, fullScore, correct, stage.qs.length, rate, secs, isBest ? "○" : "",
      wrongIds.join(","), rightIds.join(","), blur, suspect, dedupe, CLIENT_VERSION
    ]);

    if (answerRows.length) {
      const shA = sheet_(SH_ANSWER, H_ANSWER);
      shA.getRange(shA.getLastRow() + 1, 1, answerRows.length, H_ANSWER.length).setValues(answerRows);
    }

    markTotalsDirty_();        // 集計シートは、生徒を待たせないようにあとで作り直す

    // 進捗は、シートを読み直さずに「いま解いたぶん」を足して更新する
    const core = progressCoreCached_(me.email);
    core.tries += 1;
    if (core.best[stage.id] == null || score > core.best[stage.id]) core.best[stage.id] = score;
    stage.qs.forEach(function (q, i) {
      const ok = detail[i].ok;
      if (!core.stats[q.type]) core.stats[q.type] = { c: 0, t: 0 };
      core.stats[q.type].t += 1;
      if (ok) core.stats[q.type].c += 1;
      const at = core.wrong.indexOf(q.id);
      if (ok && at >= 0) core.wrong.splice(at, 1);      // 直した問題は復習から外す
      if (!ok && at < 0) core.wrong.push(q.id);
    });
    saveProgressCore_(me.email, core);
    const prog = finishProgress_(core);
    return {
      ok: true, resultUuid: resultUuid,
      score: score, fullScore: fullScore, correct: correct, rate: rate,
      isBest: isBest, prevBest: prevBest, detail: detail,
      progress: prog, rankName: prog.rankName
    };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function markOf_(i) { return ["ア", "イ", "ウ", "エ"][i] || "―"; }

/* 二重送信の見張りは、貯め込まないようにキャッシュで行う。
   スクリプトプロパティは全体で500KBまでなので、増え続ける用途には向かない。 */
function markOnce_(prefix, dedupe) {
  const key = prefix + Utilities.base64EncodeWebSafe(String(dedupe)).slice(0, 100);
  const cache = CacheService.getScriptCache();
  try {
    if (cache.get(key)) return true;
    cache.put(key, "1", CACHE_SEC);
  } catch (e) {}
  return false;
}

function alreadySaved_(sh, dedupe) {
  return markOnce_("dup_", dedupe);
}

function bestOf_(email, stageId) {
  const core = progressCoreCached_(email);
  const v = core.best[String(stageId)];
  return (v == null) ? null : v;
}

/* ============================================================
   集計・ランキング
   ============================================================ */
function collectTotals_() {
  const agg = {};
  const v = sheetValues_(SH_RESULT);
  const m = v.map;
  if (!m["email"]) return agg;
  v.rows.forEach(function (r) {
    const email = String(r[m["email"] - 1]).trim().toLowerCase();
    if (!email) return;
    if (!agg[email]) {
      agg[email] = { email: email, name: "", klass: "", no: "", tries: 0,
                     best: {}, rateSum: 0, maxScore: 0, fastest: null, last: "" };
    }
    const a = agg[email];
    a.name  = String(r[m["氏名"] - 1] || "").trim() || a.name;
    a.klass = String(r[m["組"] - 1] || "").trim() || a.klass;
    a.no    = String(r[m["番号"] - 1] || "").trim() || a.no;
    a.tries += 1;
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

/** 累計得点＝ステージごとの自己ベストの合計 */
function pointsOf_(a) {
  let n = 0;
  Object.keys(a.best).forEach(function (sid) { n += a.best[sid]; });
  return n;
}
function fullStagesOf_(a, per) {
  let n = 0;
  Object.keys(a.best).forEach(function (sid) { if (a.best[sid] >= per) n++; });
  return n;
}
function rankedList_() {
  const agg = collectTotals_();
  const list = Object.keys(agg).map(function (k) { agg[k].points = pointsOf_(agg[k]); return agg[k]; });
  list.sort(function (x, y) { return y.points - x.points; });
  return list;
}
/** 小テストの成績を email ごとに集める */
function collectQuiz_() {
  const agg = {};
  quizLog_().forEach(function (r) {
    if (!agg[r.email]) agg[r.email] = { tries: 0, best: null, sum: 0, skipped: 0, last: "" };
    const a = agg[r.email];
    a.last = r.at;
    if (!r.counted) { a.skipped++; return; }
    a.tries++;
    if (a.best == null || r.score > a.best) a.best = r.score;
    a.sum += r.score;
  });
  return agg;
}

function totalRows_() {
  const max = maxPoints_();
  const stages = practiceStages_();
  const per = stages.length ? Math.round(max / stages.length) : 12;
  const quiz = collectQuiz_();
  const list = rankedList_();
  const seen = {};
  list.forEach(function (a) { seen[a.email] = true; });
  // 練習はしていないが小テストだけ受けた生徒も行に出す
  Object.keys(quiz).forEach(function (email) {
    if (seen[email]) return;
    const who = findByEmail_(email) || { name: "", klass: "", no: "" };
    list.push({ email: email, name: who.name, klass: who.klass, no: who.no,
                points: 0, tries: 0, best: {}, rateSum: 0, maxScore: 0, fastest: null, last: quiz[email].last });
  });
  return list.map(function (a, i) {
    const q = quiz[a.email] || { tries: 0, best: null, sum: 0, skipped: 0, last: "" };
    return [i + 1, a.email, a.name, a.klass, a.no, a.points, max, a.tries,
            Object.keys(a.best).length, fullStagesOf_(a, per),
            a.tries ? Math.round(a.rateSum / a.tries) : 0,
            a.maxScore, a.fastest == null ? "" : a.fastest, rankName_(a.points, max),
            q.tries, q.best == null ? "" : q.best,
            q.tries ? Math.round(q.sum / q.tries) : "",
            q.skipped || "",
            a.last || q.last];
  });
}

function updateTotals_() {
  const sh = sheet_(SH_TOTAL, H_TOTAL);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, H_TOTAL.length).clearContent();
  const rows = totalRows_();
  if (rows.length) sh.getRange(2, 1, rows.length, H_TOTAL.length).setValues(rows);
  return rows.length;
}

function rebuildTotals() {
  clearMemo_();
  clearProgressCache_();
  const n = updateTotals_();
  try { PropertiesService.getScriptProperties().deleteProperty("totalsDirty"); } catch (e) {}
  const ui = tryUi_();
  if (ui) ui.alert(APP_NAME, "集計を再計算しました（" + n + "名）。", ui.ButtonSet.OK);
  return n;
}

/** 画面に出すランキング（上位20名＋自分） */
function getRanking(claimed) {
  clearMemo_();
  try {
    const email = resolveEmail_(claimed);
    const max = maxPoints_();
    const list = rankedList_();

    const top = [], mine = { rank: null, points: 0 };
    list.forEach(function (a, i) {
      if (a.email === email) { mine.rank = i + 1; mine.points = a.points; }
      if (i < 20) top.push({
        rank: i + 1, me: (a.email === email),
        name: a.name || localPart_(a.email), klass: a.klass,
        points: a.points, rankName: rankName_(a.points, max)
      });
    });
    return { ok: true, top: top, mine: mine, total: list.length, maxPoints: max };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/** 自分の解答履歴（新しい順） */
function getMyHistory(limit, claimed) {
  clearMemo_();
  try {
    const email = resolveEmail_(claimed);
    const sh = book_().getSheetByName(SH_RESULT);
    const out = [];
    if (!email || !sh || sh.getLastRow() < 2) return { ok: true, rows: out };
    const m = headerMap_(sh);
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    rows.forEach(function (r) {
      if (String(r[m["email"] - 1]).trim().toLowerCase() !== email) return;
      out.push({
        at: String(r[m["保存日時"] - 1]), stage: String(r[m["ステージ名"] - 1]),
        score: Number(r[m["得点"] - 1]) || 0, full: Number(r[m["満点"] - 1]) || 0,
        secs: Number(r[m["所要時間(秒)"] - 1]) || 0
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
    const email = resolveEmail_(p.email);
    const me = email ? findByEmail_(email) : null;
    if (!me) return { ok: false, message: "ログイン状態が確認できません。ページを再読み込みしてください。" };

    let q = null;
    getStages_().forEach(function (s) {
      s.qs.forEach(function (x) { if (x.id === String(p.qid)) q = x; });
    });   // 小テストの問題も報告できるよう、ここは全体から探す
    if (!q) return { ok: false, message: "問題が見つかりませんでした。" };

    sheet_(SH_ISSUE, H_ISSUE).appendRow([
      uuid_(), now_(), me.email, me.name, String(p.resultUuid || ""), q.id,
      stripTags_(q.q), q.opts[0], q.opts[1], q.opts[2], q.opts[3],
      markOf_(q.a), markOf_(Number(p.mine)),
      String(p.kind || "その他"), String(p.memo || ""), "未確認", "", ""
    ]);
    return { ok: true, message: "報告を受け付けました。ありがとうございます。先生が確認します。" };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function stripTags_(s) { return String(s || "").replace(/<[^>]*>/g, ""); }

/** 速すぎるなどで成績に入れなかった小テストの一覧 */
function showExcludedQuiz() {
  const ui = tryUi_();
  if (!ui) return;
  const sh = book_().getSheetByName(SH_QUIZ);
  if (!sh || sh.getLastRow() < 2) { ui.alert(APP_NAME, "小テストの記録がまだありません。", ui.ButtonSet.OK); return; }
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const out = [];
  rows.forEach(function (r, i) {
    if (quizCounts_(r, m)) return;
    const cheat = m["不正の疑い"] && String(r[m["不正の疑い"] - 1]).trim() === "○";
    out.push((i + 2) + "行目　" + (cheat ? "【不正】" : "【速すぎ】") +
             String(r[m["氏名"] - 1] || r[m["email"] - 1]) +
             "　" + String(r[m["ステージ名"] - 1] || "") +
             "　成績" + r[m["成績(%)"] - 1] + "％　" + String(r[m["除外理由"] - 1] || ""));
  });
  ui.alert(APP_NAME + "　成績に入れなかった小テスト",
    out.length ? (out.length + "件\n\n" + out.slice(0, 30).join("\n") +
      "\n\n※【不正】＝解答中にほかの画面に移った、またはキーボードを使った回です。\n" +
      "　　成績には入りませんが、記録は残ります。0点として扱うかは先生がお決めください。\n" +
      "※ 問題ないと判断したら、その行の「成績に反映」を ○ に書き換えてください。次の集計から数えられます。")
      : "ありません。",
    ui.ButtonSet.OK);
}

function showIssueSummary() {
  const ui = tryUi_();
  if (!ui) return;
  const sh = book_().getSheetByName(SH_ISSUE);
  if (!sh || sh.getLastRow() < 2) { ui.alert(APP_NAME, "疑義報告はまだありません。", ui.ButtonSet.OK); return; }
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const byQ = {}; let open = 0;
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
   接続確認（エディタから実行）
   ============================================================ */
function testConnection() {
  const out = [];
  out.push("ログイン中のアカウント: " + (activeEmail_() || "（取得できません。デプロイ設定を確認してください）"));
  try {
    const ss = book_();
    out.push("集計簿: " + ss.getName());
    const r = ss.getSheetByName(SH_ROSTER);
    if (r) {
      const c = rosterCols_(r);
      out.push("名簿: " + (r.getLastRow() - 1) + "行");
      out.push("  列の判定 … アドレス:" + (c.email ? "第" + c.email + "列" : "見つかりません") +
               " / 氏名:" + (c.name ? "第" + c.name + "列" : "なし") +
               " / 組:" + (c.klass ? "第" + c.klass + "列" : "なし") +
               " / 出席番号:" + (c.no ? "第" + c.no + "列" : "なし"));
      if (!c.email) out.push("  ※ アドレスの列が判定できません。見出しを「アドレス」か「email」にしてください。");
    } else {
      out.push("名簿: なし（初回セットアップ未実行）");
    }
  } catch (e) { out.push("集計簿: NG " + e.message); }
  try {
    const pool = loadPool_();
    let n = 0; pool.forEach(function (s) { n += s.qs.length; });
    let np = 0, nq = 0;
    pool.forEach(function (s) { if (s.use === "小テスト") nq++; else np++; });
    out.push("問題マスター: " + master_().getName() + " / " + pool.length + "ステージ・" + n + "問");
    out.push("  用途の内訳 … 練習 " + np + "ステージ / 小テスト " + nq + "ステージ");
    if (typeof scenesOf_ !== "function") {
      out.push("  図解: Scenes.gs が見つかりません（採点後の「登場人物とできごと」が出ません）");
    } else {
      let ns = 0;
      withScenes_(pool).forEach(function (s) { if (s.scenes) ns++; });
      out.push("  図解: " + ns + "／" + pool.length + "ステージに用意されています" +
               (ns < pool.length ? "（足りない分は Scenes.gs に本文IDを追加してください）" : ""));
    }
    out.push("  小テストモード: " + (readConfig_().quizMode ? "ON" : "OFF"));
  } catch (e) { out.push("問題マスター: NG " + e.message); }
  Logger.log(out.join("\n"));
  return out.join("\n");
}
