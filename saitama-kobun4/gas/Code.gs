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
const CLIENT_VERSION = 8;

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

const CACHE_KEY   = "kobun_pool_v1";
const CACHE_SEC   = 21600;              // 問題プールのキャッシュ 6時間
const TARGET_SEC  = 300;                // 1ステージの目安時間（5分・画面表示用）
const QUIZ_SEC    = 300;                // 小テストの制限時間（5分）
const QUIZ_COUNT  = 2;                  // 小テストで出す大問の数（5分でこの数に挑戦する）

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
                  "誤答問題ID","正答問題ID","フォーカス離脱回数","クライアント版"];

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
function findByEmail_(email) {
  const e = String(email || "").trim().toLowerCase();
  if (!e) return null;
  const sh = rosterSheet_();
  if (sh.getLastRow() < 2) return null;
  const c = rosterCols_(sh);
  if (!c.email) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  function cell(row, col) { return col ? String(row[col - 1] == null ? "" : row[col - 1]).trim() : ""; }
  for (let i = 0; i < rows.length; i++) {
    if (cell(rows[i], c.email).toLowerCase() === e) {
      return {
        email: e,
        name:  cell(rows[i], c.name),
        klass: cell(rows[i], c.klass),
        no:    cell(rows[i], c.no),
        row:   i + 2,
        teacher: isTeacher_(e)
      };
    }
  }
  return null;
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
    .addItem("問題プールのキャッシュをクリア", "clearPoolCache")
    .addItem("集計を再計算", "rebuildTotals")
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
  const ss = book_();
  let sh = ss.getSheetByName(SH_CONFIG);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_CONFIG); }
  const last = Math.max(sh.getLastRow(), 1);
  const rows = sh.getRange(1, 1, last, 1).getValues();
  let at = 0;
  rows.forEach(function (r, i) { if (String(r[0]).trim() === "小テストモード") at = i + 1; });
  if (!at) { at = last + 1; sh.getRange(at, 1).setValue("小テストモード").setFontWeight("bold"); }
  sh.getRange(at, 2).setValue(!!on);
  const ui = tryUi_();
  if (ui) ui.alert(APP_NAME, on
    ? "小テストモードを開始しました。\n生徒の画面は小テストに切り替わります（再読み込みが必要です）。\n終わったら必ず「小テストモードを終了」を実行してください。"
    : "小テストモードを終了しました。生徒は通常の練習に戻ります。", ui.ButtonSet.OK);
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

/** 小テストモードの切り替え（教員のみ） */
function setQuizMode(on) {
  try {
    const email = resolveEmail_(null);
    if (!email || !isTeacher_(email)) return { ok: false, message: "先生のアカウントでのみ切り替えられます。" };
    const ss = book_();
    let sh = ss.getSheetByName(SH_CONFIG);
    if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_CONFIG); }
    const last = Math.max(sh.getLastRow(), 1);
    const rows = sh.getRange(1, 1, last, 2).getValues();
    let at = 0;
    rows.forEach(function (r, i) { if (String(r[0]).trim() === "小テストモード") at = i + 1; });
    if (!at) { at = last + 1; sh.getRange(at, 1).setValue("小テストモード").setFontWeight("bold"); }
    sh.getRange(at, 2).setValue(!!on);
    return { ok: true, quizMode: !!on };
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
    progress: getProgress_(me.email),
    quiz: quizStatus_(me.email)
  };
}

/* ============================================================
   問題プール（問題マスターから読み込み・キャッシュ付き）
   ============================================================ */
function getStages_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const pool = loadPool_();
  try { cache.put(CACHE_KEY, JSON.stringify(pool), CACHE_SEC); } catch (e) {}
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

/** 練習で出すステージ（問題マスターの用途が「練習」のもの） */
function practiceStages_() {
  return getStages_().filter(function (s) { return s.use !== "小テスト"; });
}
/** 小テストで出すステージ */
function quizStages_() {
  return getStages_().filter(function (s) { return s.use === "小テスト"; });
}
/** 採点前に画面へ渡す用。正解と解説を落とす */
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
   小テスト
   ============================================================ */
/** その生徒が小テストですでに出題されたステージID */
function quizServedIds_(email) {
  const key = String(email || "").toLowerCase();
  const done = {};
  const sh = book_().getSheetByName(SH_QUIZ);
  if (!sh || sh.getLastRow() < 2) return done;
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  rows.forEach(function (r) {
    if (String(r[m["email"] - 1]).trim().toLowerCase() !== key) return;
    const id = String(r[m["ステージID"] - 1]).trim();
    if (id) done[id] = true;
  });
  return done;
}

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

/** 成績に入れてよい回かどうかを判定する */
function fairnessOf_(pace, score, answered, total) {
  const reasons = [];
  // 未解答が多い回は、そもそも速くて当然なので対象にしない
  const fastPerQ = (answered >= 2) && (pace.perQ < FAIR_MIN_SEC_PER_Q);
  const burst    = (pace.medGap != null) && (pace.medGap < FAIR_MIN_GAP_SEC);
  if (fastPerQ) reasons.push("1問あたり" + pace.perQ + "秒");
  if (burst)    reasons.push("解答間隔の中央値" + pace.medGap + "秒");

  if ((fastPerQ || burst) && score >= FAIR_SUSPECT_RATE) {
    return { include: false, reason: "速すぎます（" + reasons.join("・") + "）" };
  }
  return { include: true, reason: "" };
}

/** 記録の1行を成績に入れてよいか。先生が「×」に書き換えた行は外れる */
function quizCounts_(row, m) {
  const c = m["成績に反映"];
  if (!c) return true;                       // 旧い記録は従来どおり数える
  const v = String(row[c - 1]).trim();
  return v !== "×" && v !== "x" && v.toUpperCase() !== "FALSE";
}

/** 小テストの状況（残りの問題数・最高正答率） */
function quizStatus_(email) {
  const done = quizServedIds_(email);
  const pool = quizStages_();
  let remain = 0;
  pool.forEach(function (s) { if (!done[s.id]) remain++; });

  let best = null, sum = 0, tries = 0, last = null, skipped = 0;
  const sh = book_().getSheetByName(SH_QUIZ);
  if (sh && sh.getLastRow() >= 2) {
    const m = headerMap_(sh);
    const key = String(email || "").toLowerCase();
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    rows.forEach(function (r) {
      if (String(r[m["email"] - 1]).trim().toLowerCase() !== key) return;
      if (!quizCounts_(r, m)) { skipped++; return; }
      tries++;
      const sc = Number(r[m["成績(%)"] - 1]) || 0;
      if (best == null || sc > best) best = sc;
      sum += sc; last = sc;
    });
  }
  return { remain: remain, total: pool.length, tries: tries, skipped: skipped,
           best: best, avg: tries ? Math.round(sum / tries) : null, last: last,
           limitSec: QUIZ_SEC, maxParts: QUIZ_COUNT };
}

/** まだ出していない大問から1本ランダムに選ぶ（正解は含めない） */
function pickQuizStage_(email, exclude) {
  const done = quizServedIds_(email);
  (exclude || []).forEach(function (id) { done[id] = true; });
  const pool = quizStages_().filter(function (s) { return !done[s.id]; });
  if (!pool.length) return null;
  return pool[Math.floor(Math.random() * pool.length)];
}

/**
 * 小テストを始める。まず1問目（大問1つ）だけを返す。
 * 2問目は「2つ目に挑戦する」を選んだときに初めて渡すので、
 * 先に読まれてしまうことがない。
 */
function startQuiz() {
  try {
    const email = resolveEmail_(null);
    const me = email ? findByEmail_(email) : null;
    if (!me) return { ok: false, message: "ログイン状態が確認できません。ページを再読み込みしてください。" };

    const cfg = readConfig_();
    if (!cfg.quizMode && !me.teacher) {
      return { ok: false, message: "いまは小テストの時間ではありません。" };
    }

    const st = pickQuizStage_(me.email, []);
    if (!st) {
      return { ok: false, remain: 0,
        message: "小テストで出せる問題が、もうありません。先生にお知らせください。" };
    }
    return {
      ok: true,
      sessionId: uuid_(),
      index: 1,
      stage: hideAnswers_([st])[0],
      limitSec: QUIZ_SEC,
      startedAt: new Date().toISOString(),
      status: quizStatus_(me.email)
    };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/** 「2つ目に挑戦する」を選んだときに、次の大問を渡す */
function nextQuizStage(exclude) {
  try {
    const email = resolveEmail_(null);
    const me = email ? findByEmail_(email) : null;
    if (!me) return { ok: false, message: "ログイン状態が確認できません。ページを再読み込みしてください。" };
    const cfg = readConfig_();
    if (!cfg.quizMode && !me.teacher) return { ok: false, message: "いまは小テストの時間ではありません。" };
    const nx = pickQuizStage_(me.email, exclude || []);
    if (!nx) return { ok: true, noMore: true };
    return { ok: true, stage: hideAnswers_([nx])[0] };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/**
 * 大問1つ分を採点して記録する。
 * 記録した時点で、その大問の成績は確定する。
 */
function submitQuizPart(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, message: "混み合っています。少し待ってからもう一度お試しください。" };
  }
  try {
    const p = payload || {};
    const email = resolveEmail_(p.email);
    const me = email ? findByEmail_(email) : null;
    if (!me) return { ok: false, message: "ログイン状態が確認できません。ページを再読み込みしてください。" };

    const sid = String(p.stageId || "");
    let stage = null;
    quizStages_().forEach(function (s) { if (s.id === sid) stage = s; });
    if (!stage) return { ok: false, message: "問題が見つかりません。もう一度やり直してください。" };

    const picks  = p.picks || {};
    const secs   = Math.max(0, Math.round(Number(p.seconds) || 0));
    const blur   = Math.max(0, Math.round(Number(p.blur) || 0));
    const timeUp = !!p.timeUp;
    const index  = Math.max(1, Math.round(Number(p.index) || 1));
    const sessionId = String(p.sessionId || "");

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
    const fair = fairnessOf_(pace, score, answered, total);

    const sh = sheet_(SH_QUIZ, H_QUIZ);
    const dedupe = [me.email, sessionId, sid, String(index)].join("|");
    if (!alreadyQuizSaved_(sh, dedupe)) {
      sh.appendRow([
        uuid_(), now_(), me.email, me.name, me.klass, me.no, sessionId, index,
        score, fair.include ? "○" : "×", fair.reason, sid, stage.title,
        total, correct, secs, pace.perQ, pace.medGap == null ? "" : pace.medGap,
        timeUp ? "○" : "", wrongIds.join(","), rightIds.join(","), blur, CLIENT_VERSION
      ]);
      updateTotals_();
    }

    const res = {
      ok: true, index: index, stageId: sid, title: stage.title, src: stage.src,
      total: total, correct: correct, score: score,
      counted: fair.include, reason: fair.reason, pace: pace,
      timeUp: timeUp, seconds: secs, detail: detail,
      status: quizStatus_(me.email)
    };

    return res;
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function alreadyQuizSaved_(sh, dedupe) {
  const key = "quizdedupe_" + Utilities.base64EncodeWebSafe(dedupe).slice(0, 80);
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(key) === "1") return true;
  props.setProperty(key, "1");
  return false;
}

/* ============================================================
   進捗
   ============================================================ */
function getProgress_(email) {
  const key = String(email || "").toLowerCase();
  const prog = { best: {}, stats: {}, wrong: [], points: 0, maxPoints: 0, tries: 0, rankName: RANKS[RANKS.length - 1].name };
  Q_TYPES.forEach(function (t) { prog.stats[t] = { c: 0, t: 0 }; });

  const ss = book_();

  const shR = ss.getSheetByName(SH_RESULT);
  if (shR && shR.getLastRow() >= 2) {
    const m = headerMap_(shR);
    const rows = shR.getRange(2, 1, shR.getLastRow() - 1, shR.getLastColumn()).getValues();
    rows.forEach(function (r) {
      if (String(r[m["email"] - 1]).trim().toLowerCase() !== key) return;
      const sid = String(r[m["ステージID"] - 1]).trim();
      const sc  = Number(r[m["得点"] - 1]) || 0;
      if (prog.best[sid] == null || sc > prog.best[sid]) prog.best[sid] = sc;
      prog.tries++;
    });
  }

  const shA = ss.getSheetByName(SH_ANSWER);
  if (shA && shA.getLastRow() >= 2) {
    const m = headerMap_(shA);
    const rows = shA.getRange(2, 1, shA.getLastRow() - 1, shA.getLastColumn()).getValues();
    const latest = {};
    rows.forEach(function (r) {
      if (String(r[m["email"] - 1]).trim().toLowerCase() !== key) return;
      const type = String(r[m["設問タイプ"] - 1]).trim();
      const ok   = String(r[m["正誤"] - 1]).trim() === "○";
      if (prog.stats[type]) { prog.stats[type].t++; if (ok) prog.stats[type].c++; }
      latest[String(r[m["問題ID"] - 1]).trim()] = ok;
    });
    Object.keys(latest).forEach(function (qid) { if (!latest[qid]) prog.wrong.push(qid); });
  }

  // 累計得点は「ステージごとの自己ベストの合計」。解き直しても水増しされない。
  Object.keys(prog.best).forEach(function (sid) { prog.points += prog.best[sid]; });
  prog.maxPoints = maxPoints_();
  prog.rankName = rankName_(prog.points, prog.maxPoints);
  return prog;
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
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, message: "混み合っています。少し待ってからもう一度お試しください。" };
  }
  try {
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
               progress: getProgress_(me.email), resultUuid: resultUuid };
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

    updateTotals_();

    const prog = getProgress_(me.email);
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

function alreadySaved_(sh, dedupe) {
  if (sh.getLastRow() < 2) return false;
  const m = headerMap_(sh);
  const col = m["重複排除キー"];
  if (!col) return false;
  const vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === dedupe) return true;
  return false;
}

function bestOf_(email, stageId) {
  const key = String(email).toLowerCase();
  const sh = book_().getSheetByName(SH_RESULT);
  if (!sh || sh.getLastRow() < 2) return null;
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  let best = null;
  rows.forEach(function (r) {
    if (String(r[m["email"] - 1]).trim().toLowerCase() !== key) return;
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
  const sh = book_().getSheetByName(SH_QUIZ);
  if (!sh || sh.getLastRow() < 2) return agg;
  const m = headerMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  rows.forEach(function (r) {
    const email = String(r[m["email"] - 1]).trim().toLowerCase();
    if (!email) return;
    if (!agg[email]) agg[email] = { tries: 0, best: null, sum: 0, skipped: 0, last: "" };
    const a = agg[email];
    a.last = String(r[m["保存日時"] - 1]);
    if (!quizCounts_(r, m)) { a.skipped++; return; }
    a.tries++;
    const sc = Number(r[m["成績(%)"] - 1]) || 0;
    if (a.best == null || sc > a.best) a.best = sc;
    a.sum += sc;
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
  const n = updateTotals_();
  const ui = tryUi_();
  if (ui) ui.alert(APP_NAME, "集計を再計算しました（" + n + "名）。", ui.ButtonSet.OK);
  return n;
}

/** 画面に出すランキング（上位20名＋自分） */
function getRanking(claimed) {
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
    out.push((i + 2) + "行目　" + String(r[m["氏名"] - 1] || r[m["email"] - 1]) +
             "　" + String(r[m["ステージ名"] - 1] || "") +
             "　成績" + r[m["成績(%)"] - 1] + "％　" + String(r[m["除外理由"] - 1] || ""));
  });
  ui.alert(APP_NAME + "　成績に入れなかった小テスト",
    out.length ? (out.length + "件\n\n" + out.slice(0, 30).join("\n") +
      "\n\n※ 問題ないと判断したら、その行の「成績に反映」を ○ に書き換えてください。次の集計から数えられます。")
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
    out.push("  小テストモード: " + (readConfig_().quizMode ? "ON" : "OFF"));
  } catch (e) { out.push("問題マスター: NG " + e.message); }
  Logger.log(out.join("\n"));
  return out.join("\n");
}
