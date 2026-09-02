/**
 * クラスまど - サーバー側ロジック
 *
 * このスクリプトをスプレッドシートに紐づけて実行すること（拡張機能 > Apps Script）。
 * 詳しいセットアップ手順は README.md を参照。
 */

// 学校行事を外部のスプレッドシートから参照する場合は、そのIDをここに設定する。
// URLの https://docs.google.com/spreadsheets/d/【ここ】/edit の部分を貼り付ける。
// 次の2通りの形式に対応している：
//  (1) 年間行事予定ファイル（「４月」〜「３月」の月別シートがある形式）
//      → 「行事を取り込む」（syncEventsFromMaster）で、このスプレッドシートの
//        「行事取込」シートに日付・曜日・週・行事・日課が自動転記される。
//  (2) 「予定」シート1枚に日付とタイトルを並べた行事マスター形式（従来どおり）
// 空欄のままなら、このスプレッドシート内の「予定」シートだけを使う。
var EVENT_SS_ID = '';
// (2)の形式のとき、行事が入っているシート名
var EVENT_SHEET_NAME = '予定';

// 後期の開始日（この日以降は「時間割」シートの後期の教科列を表示する）。
// 「設定」シートに キー「後期開始日」・値「2026-11-09」の行を作れば、コードを触らずに変更できる。
var SECOND_TERM_START_DEFAULT = '2026-11-09';

var SHEET = {
  CLASS_LIST: 'クラス設定',
  USERS: '利用者情報',
  ASSIGNMENTS: '担当割当',
  SCHEDULE: '時間割',
  SCHEDULE_CHANGES: '時間割変更',
  DAILY_NOTES: '日別連絡',
  EVENTS: '予定',
  EVENT_IMPORT: '行事取込',
  TASKS: '課題',
  SUBMISSIONS: '提出状況',
  POSTS: '連絡',
  SETTINGS: '設定'
};

var SUBJECTS = ["国語", "数学", "英語", "理科", "社会", "音楽", "美術", "保健", "技術", "家庭"];

var CATEGORY_NAME = {
  subject: '教科',
  duty: '係',
  committee: '委員会',
  homeroom: '学級・担任'
};

/* ============================================================
 * エントリーポイント
 * ============================================================ */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('クラスまど')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================================================
 * スプレッドシート読み書きヘルパー
 * ============================================================ */

// このアプリ自身のデータ（担当割当・時間割・予定・課題・提出状況・連絡）が入っているスプレッドシート
function getSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('スプレッドシートに接続できません。スクリプトをスプレッドシートに紐づけるか、'
      + 'スクリプトプロパティに SPREADSHEET_ID を設定してください。');
  }
  return SpreadsheetApp.openById(id);
}

function getSheetFrom_(ss, name, sourceLabel) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('シート「' + name + '」が' + sourceLabel + 'に見つかりません。シート名を確認してください。');
  return sheet;
}

function getSheet_(name) {
  return getSheetFrom_(getSS_(), name, 'このスプレッドシート');
}

// 見出し行が何行目にあるかを探す。
// テンプレートの各シートは1行目が「※〜」の注意書きで、見出しは3行目にあるため、
// 上から順に見て「空でなく、※で始まらない最初の行」を見出し行とみなす。
function findHeaderRow_(sheet) {
  var lastRow = Math.min(sheet.getLastRow(), 10);
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return 0;
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  for (var i = 0; i < values.length; i++) {
    var first = values[i][0];
    if (first !== '' && first !== null && String(first).charAt(0) !== '※') return i + 1;
  }
  return 0;
}

// シートをヘッダー行をキーにしたオブジェクトの配列として読み込む
function readSheetFrom_(ss, name, sourceLabel) {
  var sheet = getSheetFrom_(ss, name, sourceLabel);
  var headerRow = findHeaderRow_(sheet);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (!headerRow || lastRow <= headerRow) return [];
  var values = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, lastCol).getValues();
  var headers = values.shift();
  return values
    .filter(function (row) { return row.some(function (c) { return c !== '' && c !== null; }); })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

// このアプリ自身のスプレッドシートからシートを読み込む
function readSheet_(name) {
  return readSheetFrom_(getSS_(), name, 'このスプレッドシート');
}

// ヘッダー行に合わせて1行追加する（このアプリ自身のスプレッドシートのみ）
function appendRow_(name, obj) {
  var sheet = getSheet_(name);
  var headerRow = findHeaderRow_(sheet);
  if (!headerRow) throw new Error('シート「' + name + '」の見出し行が見つかりません。');
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var row = headers.map(function (h) { return (obj[h] !== undefined && obj[h] !== null) ? obj[h] : ''; });
  sheet.appendRow(row);
}

// 「設定」シート（キー・値の2列）から設定値を読む。シートや行がなければ既定値を返す。
function getSetting_(key, defaultValue) {
  try {
    var row = readSheet_(SHEET.SETTINGS).filter(function (r) { return r['キー'] === key; })[0];
    return row && row['値'] !== '' && row['値'] != null ? String(row['値']) : defaultValue;
  } catch (e) {
    return defaultValue;
  }
}

// 「設定」シートに設定値を書き込む（既存の行があれば更新、なければ追加）
function setSetting_(key, value) {
  var sheet = getSheet_(SHEET.SETTINGS);
  var headerRow = findHeaderRow_(sheet);
  if (!headerRow) throw new Error('「設定」シートの見出し行が見つかりません。');
  var lastRow = sheet.getLastRow();
  if (lastRow > headerRow) {
    var values = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, sheet.getLastColumn()).getValues();
    var headers = values[0];
    var keyIdx = headers.indexOf('キー');
    var valIdx = headers.indexOf('値');
    for (var i = 1; i < values.length; i++) {
      if (values[i][keyIdx] === key) {
        sheet.getRange(headerRow + i, valIdx + 1).setValue(value);
        return;
      }
    }
  }
  appendRow_(SHEET.SETTINGS, { 'キー': key, '値': value });
}

function toDateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v ? String(v) : '';
}

function toDateTimeLabel_(v) {
  var d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v || '');
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'M/d HH:mm');
}

/* ============================================================
 * ユーザー識別・権限
 * ============================================================ */

// メールアドレスのローカル部分が「t」で始まる場合は教員とみなす
// （例：t13133@saitama-city.ed.jp）。この学校のアカウント命名規則に基づく。
function isTeacherEmail_(email) {
  var local = String(email).split('@')[0];
  return /^t/i.test(local);
}

// 「クラス設定」シートに登録されているクラス名の一覧を返す
function getClassList_() {
  return readSheet_(SHEET.CLASS_LIST)
    .map(function (r) { return String(r['クラス'] || '').trim(); })
    .filter(function (name) { return name; });
}

// 「利用者情報」シートから、このメールアドレスの登録内容を探す
// （生徒が初回アクセス時に自分で名前・クラスを登録したものが積み上がっていく）
function findUserInfo_(email) {
  return readSheet_(SHEET.USERS).filter(function (u) { return u['メールアドレス'] === email; })[0] || null;
}

// ログイン中のユーザーを判定する。
// ・メールアドレスが t で始まる → 教員（全クラス閲覧・投稿可）
// ・それ以外 → 「利用者情報」シートに自己登録済みのクラスを使う
//   （まだ登録がなければ needsRegistration:true を返し、クライアント側で登録フォームを出す）
function getContext_() {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('ログイン情報を取得できませんでした。学校のGoogleアカウントでログインしているか確認してください。');
  }

  var isTeacher = isTeacherEmail_(email);
  var info = findUserInfo_(email);

  if (!info) {
    return { email: email, role: isTeacher ? 'teacher' : 'student', needsRegistration: true };
  }

  if (isTeacher) {
    return { email: email, name: info['氏名'] || email, role: 'teacher', myClass: null, needsRegistration: false };
  }
  return { email: email, name: info['氏名'] || email, role: 'student', myClass: info['クラス'], needsRegistration: false };
}

function assertCanView_(ctx, cls) {
  if (ctx.role === 'teacher') return;
  if (ctx.myClass !== cls) throw new Error('このクラスのページは閲覧できません。');
}

// 指定クラス・指定用途（kind: 'tasks' | 'announcement'）で
// このユーザーがどの立場（教科名・係名など）として投稿できるかを返す
function getPostPermissions_(ctx, cls, kind) {
  var assignments = readSheet_(SHEET.ASSIGNMENTS).filter(function (a) { return a['クラス'] === cls; });
  var perms = [];

  if (ctx.role === 'teacher') {
    perms.push({ category: 'homeroom', label: '学級・担任' });
    SUBJECTS.forEach(function (s) { perms.push({ category: 'subject', label: s }); });
    assignments
      .filter(function (a) { return a['カテゴリ'] === 'duty' || a['カテゴリ'] === 'committee'; })
      .forEach(function (a) { perms.push({ category: a['カテゴリ'], label: a['表示名'] }); });
  } else {
    assignments
      .filter(function (a) { return a['担当者メール'] === ctx.email; })
      .forEach(function (a) { perms.push({ category: a['カテゴリ'], label: a['表示名'] }); });
  }

  // 課題・提出物の投稿は教員のみ（教科係・係・委員会の生徒は連絡と持ち物の入力まで）
  if (kind === 'tasks') {
    if (ctx.role !== 'teacher') return [];
    perms = perms.filter(function (p) { return p.category === 'subject' || p.category === 'homeroom'; });
  }

  var seen = {};
  return perms.filter(function (p) {
    var key = p.category + '|' + p.label;
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function getAllClasses_() {
  return getClassList_().sort();
}

/* ============================================================
 * API（クライアントの google.script.run から呼ばれる関数）
 * ============================================================ */

function api_getContext() {
  var ctx = getContext_();
  if (ctx.needsRegistration) {
    return { needsRegistration: true, email: ctx.email, role: ctx.role, classes: getAllClasses_() };
  }
  var classes = ctx.role === 'teacher' ? getAllClasses_() : [ctx.myClass];

  // 行事予定ファイルを取り込み済みなら、そのA/B週から今週の週区分を自動判定する
  var autoWeek = null;
  try { autoWeek = getWeekFromImport_(); } catch (e) { autoWeek = null; }

  return {
    needsRegistration: false, email: ctx.email, name: ctx.name, role: ctx.role,
    myClass: ctx.myClass, classes: classes,
    currentWeek: autoWeek || getSetting_('今週の週区分', 'A'),
    currentWeekAuto: !!autoWeek,
    eventSyncTime: getSetting_('行事取込日時', ''),
    secondTermStart: getSecondTermStart_()
  };
}

// 今週がA週かB週かを設定する（教員のみ）
function api_setCurrentWeek(week) {
  var ctx = getContext_();
  if (ctx.role !== 'teacher') throw new Error('週区分の変更は教員のみ行えます。');
  if (week !== 'A' && week !== 'B') throw new Error('週区分はAまたはBを指定してください。');
  setSetting_('今週の週区分', week);
  return { ok: true, currentWeek: week };
}

// 生徒・教員が初回アクセス時に自分で名前（生徒はクラスも）を登録する
function api_register(payload) {
  var email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('ログイン情報を取得できませんでした。');

  var name = payload && payload.name ? String(payload.name).trim() : '';
  if (!name) throw new Error('お名前を入力してください。');

  var isTeacher = isTeacherEmail_(email);
  var cls = '';
  if (!isTeacher) {
    cls = payload && payload.cls ? String(payload.cls).trim() : '';
    if (!cls) throw new Error('クラスを選択してください。');
    if (getAllClasses_().indexOf(cls) < 0) throw new Error('選択されたクラスが見つかりません。');
  }

  var sheet = getSheet_(SHEET.USERS);
  var headerRow = findHeaderRow_(sheet);
  var lastRow = sheet.getLastRow();
  if (headerRow && lastRow > headerRow) {
    var values = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, sheet.getLastColumn()).getValues();
    var headers = values[0];
    var emailIdx = headers.indexOf('メールアドレス');
    var nameIdx = headers.indexOf('氏名');
    var clsIdx = headers.indexOf('クラス');
    var timeIdx = headers.indexOf('登録日時');
    for (var i = 1; i < values.length; i++) {
      if (values[i][emailIdx] === email) {
        var rowNum = headerRow + i;
        sheet.getRange(rowNum, nameIdx + 1).setValue(name);
        sheet.getRange(rowNum, clsIdx + 1).setValue(cls);
        sheet.getRange(rowNum, timeIdx + 1).setValue(new Date());
        return { ok: true };
      }
    }
  }
  appendRow_(SHEET.USERS, { 'メールアドレス': email, '氏名': name, 'クラス': cls, '登録日時': new Date() });
  return { ok: true };
}

// 後期の開始日を yyyy-MM-dd で返す（「設定」シートの「後期開始日」が優先）
function getSecondTermStart_() {
  try {
    var row = readSheet_(SHEET.SETTINGS).filter(function (r) { return r['キー'] === '後期開始日'; })[0];
    if (row && row['値'] != null && row['値'] !== '') {
      var s = toDateStr_(row['値']);  // セルが日付型でも文字列でも yyyy-MM-dd に揃える
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
      var m = String(row['値']).match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);  // 2026/11/9 形式も許容
      if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    }
  } catch (e) {}
  return SECOND_TERM_START_DEFAULT;
}

function api_getSchedule(cls) {
  var ctx = getContext_();
  assertCanView_(ctx, cls);
  var rows = readSheet_(SHEET.SCHEDULE);

  // 「後期」を含む列（例：後期教科・後期の教科）があれば、後期の教科として読む
  var secondKey = null;
  if (rows.length > 0) {
    Object.keys(rows[0]).some(function (k) {
      if (String(k).indexOf('後期') >= 0) { secondKey = k; return true; }
      return false;
    });
  }

  return rows
    .filter(function (r) { return r['クラス'] === cls; })
    .map(function (r) {
      return {
        day: r['曜日'], period: Number(r['時限']), subject: r['教科'],
        subjectSecond: secondKey && r[secondKey] ? String(r[secondKey]).trim() : '',  // 空欄=前期と同じ
        belongings: r['持ち物'] || '',
        week: r['週'] ? String(r['週']).trim().toUpperCase() : ''   // 空欄=A週・B週共通
      };
    })
    .sort(function (a, b) { return a.period - b.period; });
}

// 時間割の1コマを設定する（教員のみ）。
// week: '' = A週・B週共通 / 'A' / 'B'。subject を空にするとそのコマを削除する。
// belongings（毎回の持ち物）は指定されたときだけ書き換える。
function api_setScheduleCell(payload) {
  var ctx = getContext_();
  if (ctx.role !== 'teacher') throw new Error('時間割の編集は教員のみ行えます。');

  var cls = payload && payload.cls;
  var week = payload && payload.week ? String(payload.week).trim().toUpperCase() : '';
  var day = payload && payload.day;
  var period = payload && Number(payload.period);
  var subject = payload && payload.subject != null ? String(payload.subject).trim() : '';
  var belongings = payload && payload.belongings != null ? String(payload.belongings).trim() : null;
  if (!cls || !day || !period) throw new Error('対象のコマが指定されていません。');
  if (week !== '' && week !== 'A' && week !== 'B') throw new Error('週の指定が正しくありません。');

  var sheet = getSheet_(SHEET.SCHEDULE);
  var headerRow = findHeaderRow_(sheet);
  if (!headerRow) throw new Error('時間割シートの見出し行が見つかりません。');
  var lastRow = sheet.getLastRow();

  var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var clsIdx = headers.indexOf('クラス');
  var weekIdx = headers.indexOf('週');   // 週列がない古いシートでも動くようにする
  var dayIdx = headers.indexOf('曜日');
  var periodIdx = headers.indexOf('時限');
  var subjectIdx = headers.indexOf('教科');

  if (lastRow > headerRow) {
    var values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    for (var i = 0; i < values.length; i++) {
      var rowWeek = weekIdx >= 0 && values[i][weekIdx] ? String(values[i][weekIdx]).trim().toUpperCase() : '';
      if (values[i][clsIdx] === cls && rowWeek === week
          && values[i][dayIdx] === day && Number(values[i][periodIdx]) === period) {
        if (subject === '') {
          sheet.deleteRow(headerRow + 1 + i);
        } else {
          sheet.getRange(headerRow + 1 + i, subjectIdx + 1).setValue(subject);
          var belongingsIdx = headers.indexOf('持ち物');
          if (belongings !== null && belongingsIdx >= 0) {
            sheet.getRange(headerRow + 1 + i, belongingsIdx + 1).setValue(belongings);
          }
        }
        return { ok: true };
      }
    }
  }

  if (subject === '') return { ok: true };  // 削除対象がなければ何もしない
  var newRow = { 'クラス': cls, '曜日': day, '時限': period, '教科': subject, '持ち物': belongings !== null ? belongings : '' };
  if (weekIdx >= 0) newRow['週'] = week;
  appendRow_(SHEET.SCHEDULE, newRow);
  return { ok: true };
}

// 行事マスター形式（「予定」シート）の1行を行事オブジェクトに変換する。
// 列名の揺れ（タイトル／行事名／行事／内容、日付／月日）はある程度吸収する。
function mapEventRow_(r) {
  var date = r['日付'] !== undefined ? r['日付'] : r['月日'];
  var title = r['タイトル'] !== undefined ? r['タイトル']
    : (r['行事名'] !== undefined ? r['行事名']
    : (r['行事'] !== undefined ? r['行事'] : r['内容']));
  return { date: toDateStr_(date), cls: r['クラス'] || '全校', title: title || '' };
}

// 学校行事を読み込む。次の3つを合わせて返す：
//  ① 年間行事予定ファイルから取り込んだ「行事取込」シート（全校行事）
//  ② このスプレッドシートの「予定」シート（手入力の予定・クラス別の予定）
//  ③ EVENT_SS_ID 先の「予定」シート（従来の行事マスター形式との互換）
function readEvents_() {
  var events = [];

  readImportedEvents_().forEach(function (e) {
    if (e.title) events.push({ date: e.date, cls: '全校', title: e.title, you: e.you, week: e.week });
  });

  readSheet_(SHEET.EVENTS).forEach(function (r) {
    var e = mapEventRow_(r);
    if (e.date && e.title) events.push(e);
  });

  if (EVENT_SS_ID) {
    try {
      var master = SpreadsheetApp.openById(EVENT_SS_ID);
      // 年間行事予定ファイル（月別シート形式）を指している場合は「予定」シートがないので読み飛ばす
      if (master.getSheetByName(EVENT_SHEET_NAME)) {
        readSheetFrom_(master, EVENT_SHEET_NAME, '行事マスター').forEach(function (r) {
          var e = mapEventRow_(r);
          if (e.date && e.title) events.push(e);
        });
      }
    } catch (err) {
      // 取込済みの行事があれば、マスターに一時的にアクセスできなくても表示は続ける
      if (events.length === 0) {
        throw new Error('行事のスプレッドシートを開けませんでした。Code.gs の EVENT_SS_ID と共有設定を確認してください。');
      }
    }
  }
  return events;
}

function api_getEvents(cls) {
  var ctx = getContext_();
  assertCanView_(ctx, cls);
  // 年間分を取り込むと過去の行事も大量に含まれるため、今日以降だけ返す
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return readEvents_()
    .filter(function (r) { return (r.cls === cls || r.cls === '全校') && r.date >= today; })
    .sort(function (a, b) { return a.date.localeCompare(b.date); });
}

/* ============================================================
 * 年間行事予定ファイルからの行事取込
 * （西原中ダッシュボードの syncEventImportFromR8 と同じ仕組み）
 *
 * 年間行事予定ファイルは「４月」〜「３月」の月別シートを持ち、各シートに
 *   日 | 曜 | 週 | 行事計画等 | … | 日直 | 清掃 | 日課 | …
 * という見出し行がある想定。ここから 日付・曜日・週・行事・日課 を抜き出して
 * このスプレッドシートの「行事取込」シートに書き込む。
 * ============================================================ */

function toHalfWidthDigits_(s) {
  return String(s).replace(/[０-９]/g, function (c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
}

// 「４月」「10月」のような月別シート名なら月の数値を返す（違えば0）
function monthFromSheetName_(name) {
  var m = toHalfWidthDigits_(String(name).trim()).match(/^(\d{1,2})月$/);
  if (!m) return 0;
  var n = Number(m[1]);
  return (n >= 1 && n <= 12) ? n : 0;
}

var YOUBI_CHARS_ = ['日', '月', '火', '水', '木', '金', '土'];

// 月別シート1枚から行事を抽出する
function extractEventsFromMonthSheet_(sheet, month) {
  var values = sheet.getDataRange().getValues();

  // 見出し行（「日」「曜」「行事…」が並ぶ行）を探し、各列の位置を覚える
  var headerIdx = -1, cols = null;
  for (var i = 0; i < Math.min(values.length, 10); i++) {
    var row = values[i].map(function (v) { return String(v == null ? '' : v).trim(); });
    var dayCol = row.indexOf('日');
    var youCol = row.indexOf('曜');
    var eventCol = -1;
    for (var j = 0; j < row.length; j++) {
      if (row[j].indexOf('行事') === 0) { eventCol = j; break; }
    }
    if (dayCol >= 0 && youCol >= 0 && eventCol >= 0) {
      headerIdx = i;
      cols = { day: dayCol, you: youCol, week: row.indexOf('週'), event: eventCol, nikka: row.indexOf('日課') };
      break;
    }
  }
  if (headerIdx < 0) return [];

  // 学年ごとの校時列を探す。見出し行に「１年」「２年」…があり、
  // その次の行に 1,2,3… と時限番号が並ぶ（「時数」列で止まる）。
  // セルの値は「A水3」（=A週水曜3限の授業）・「授」（=通常どおり）・
  // 「学活」「訓練」などの特別活動名・空欄（=その時限なし）のいずれか。
  var grades = [];
  var numRow = values[headerIdx + 1] || [];
  for (var j = 0; j < values[headerIdx].length; j++) {
    var h = toHalfWidthDigits_(String(values[headerIdx][j] == null ? '' : values[headerIdx][j]).trim());
    var gm = h.match(/^(\d)年$/);
    if (!gm) continue;
    var periodCols = [];
    for (var k = j; k < numRow.length; k++) {
      var pn = Number(numRow[k]);
      if (!pn || pn !== Math.floor(pn) || pn !== periodCols.length + 1) break;
      periodCols.push(k);
    }
    if (periodCols.length > 0) grades.push({ grade: Number(gm[1]), periodCols: periodCols });
  }

  // 見出しより上の行から西暦年を探す（1〜3月のシートは年が翌年になっている）
  var year = 0;
  for (var i = 0; i < headerIdx && !year; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var n = Number(values[i][j]);
      if (n >= 2000 && n <= 2100 && n === Math.floor(n)) { year = n; break; }
    }
  }
  if (!year) {
    // 年が見つからなければ、今日を基準にした年度から推定する（4月〜3月区切り）
    var today = new Date();
    var fiscalYear = today.getFullYear() - ((today.getMonth() + 1) < 4 ? 1 : 0);
    year = month >= 4 ? fiscalYear : fiscalYear + 1;
  }

  var tz = Session.getScriptTimeZone();
  var events = [];
  for (var r = headerIdx + 1; r < values.length; r++) {
    var d = Number(values[r][cols.day]);
    if (!d || d !== Math.floor(d) || d < 1 || d > 31) continue;
    var you = String(values[r][cols.you] == null ? '' : values[r][cols.you]).trim();
    if (YOUBI_CHARS_.indexOf(you) < 0) continue;  // 曜日のない行は集計行などとみなして飛ばす
    var date = new Date(year, month - 1, d);
    if (date.getMonth() !== month - 1) continue;  // 存在しない日付（2/30など）は飛ばす
    var week = cols.week >= 0 ? String(values[r][cols.week] || '').trim().toUpperCase() : '';
    if (week !== 'A' && week !== 'B') week = '';
    var plans = {};
    grades.forEach(function (g) {
      plans[g.grade] = g.periodCols.map(function (c) {
        return String(values[r][c] == null ? '' : values[r][c]).trim();
      });
    });
    events.push({
      date: Utilities.formatDate(date, tz, 'yyyy-MM-dd'),
      you: you,
      week: week,
      title: String(values[r][cols.event] == null ? '' : values[r][cols.event]).trim(),
      nikka: cols.nikka >= 0 ? String(values[r][cols.nikka] == null ? '' : values[r][cols.nikka]).trim() : '',
      plans: plans
    });
  }
  return events;
}

// 年間行事予定ファイル（EVENT_SS_ID）の月別シートを読み、「行事取込」シートに書き込む。
// エディタから手動実行してもよいし、教員は画面の「行事を取り込む」ボタンからも実行できる。
function syncEventsFromMaster() {
  if (!EVENT_SS_ID) {
    throw new Error('年間行事予定ファイルのIDが設定されていません。Code.gs の EVENT_SS_ID に設定してください。');
  }
  var master;
  try {
    master = SpreadsheetApp.openById(EVENT_SS_ID);
  } catch (e) {
    throw new Error('年間行事予定ファイルを開けませんでした。Code.gs の EVENT_SS_ID と共有設定を確認してください。');
  }

  var all = [];
  master.getSheets().forEach(function (sheet) {
    var month = monthFromSheetName_(sheet.getName());
    if (!month) return;
    all = all.concat(extractEventsFromMonthSheet_(sheet, month));
  });
  if (all.length === 0) {
    throw new Error('月別シート（「４月」〜「３月」）から行事を読み取れませんでした。年間行事予定ファイルの形式を確認してください。');
  }

  var seen = {};
  all = all.filter(function (e) {
    if (seen[e.date]) return false;
    seen[e.date] = true;
    return true;
  });
  all.sort(function (a, b) { return a.date.localeCompare(b.date); });

  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET.EVENT_IMPORT);
  if (!sheet) sheet = ss.insertSheet(SHEET.EVENT_IMPORT);
  sheet.clearContents();

  // 学年ごとの校時列（１年校時・２年校時…）。6時限分を「|」区切りの1セルにまとめる
  var gradeSet = {};
  all.forEach(function (e) {
    Object.keys(e.plans || {}).forEach(function (g) { gradeSet[g] = true; });
  });
  var gradeList = Object.keys(gradeSet).map(Number).sort(function (a, b) { return a - b; });

  var headers = ['日付', '曜日', '週', '行事', '日課'].concat(
    gradeList.map(function (g) { return g + '年校時'; })
  );
  var noteRow = ['※ このシートは年間行事予定ファイルから自動で取り込まれます。手動で書き換えても、次回の取込で上書きされます。校時列は各時限を「|」区切りで保存しています。'];
  while (noteRow.length < headers.length) noteRow.push('');

  var rows = [noteRow, headers];
  all.forEach(function (e) {
    var row = [e.date, e.you, e.week, e.title, e.nikka];
    gradeList.forEach(function (g) {
      row.push(e.plans && e.plans[g] ? e.plans[g].join('|') : '');
    });
    rows.push(row);
  });
  sheet.getRange(1, 1, rows.length, 1).setNumberFormat('@');  // 日付を文字列のまま保つ
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

  var syncedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  setSetting_('行事取込日時', syncedAt);

  var count = all.filter(function (e) { return e.title; }).length;
  return { ok: true, days: all.length, count: count, syncedAt: syncedAt };
}

// 「行事取込」シートを読み込む（シートがまだなければ空配列）
function readImportedEvents_() {
  var ss = getSS_();
  if (!ss.getSheetByName(SHEET.EVENT_IMPORT)) return [];
  return readSheet_(SHEET.EVENT_IMPORT).map(function (r) {
    // 「１年校時」のような列を、学年→時限ごとの校時の配列に戻す
    var plans = {};
    Object.keys(r).forEach(function (k) {
      var m = toHalfWidthDigits_(String(k)).match(/^(\d)年校時$/);
      if (m && r[k] != null && String(r[k]) !== '') {
        plans[Number(m[1])] = String(r[k]).split('|').map(function (t) { return t.trim(); });
      }
    });
    return {
      date: toDateStr_(r['日付']),
      you: String(r['曜日'] == null ? '' : r['曜日']).trim(),
      week: r['週'] ? String(r['週']).trim().toUpperCase() : '',
      title: String(r['行事'] == null ? '' : r['行事']).trim(),
      nikka: String(r['日課'] == null ? '' : r['日課']).trim(),
      plans: plans
    };
  }).filter(function (e) { return e.date; });
}

// 取り込んだ行事予定から「日付→A/B週」の対応表を作る
function buildWeekByDate_() {
  var byDate = {};
  readImportedEvents_().forEach(function (e) { if (e.week) byDate[e.date] = e.week; });
  return byDate;
}

// 指定した日付（yyyy-MM-dd）のA/B週を、取り込んだ行事予定から判定する。
// その日の行に週があればそれを、なければ同じ週の月〜金で最初に見つかった週を返す。
// 判定できなければ null（「設定」シートの値を使う）。
function getWeekForDate_(dateStr, byDate) {
  if (!byDate) byDate = buildWeekByDate_();
  if (byDate[dateStr]) return byDate[dateStr];

  var parts = String(dateStr).split('-');
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  if (isNaN(d.getTime())) return null;

  var tz = Session.getScriptTimeZone();
  var monday = new Date(d);
  monday.setDate(d.getDate() - ((d.getDay() + 6) % 7));  // その週の月曜
  for (var i = 0; i < 5; i++) {
    var x = new Date(monday);
    x.setDate(monday.getDate() + i);
    var s = Utilities.formatDate(x, tz, 'yyyy-MM-dd');
    if (byDate[s]) return byDate[s];
  }
  return null;
}

// 今週の週区分を自動判定する（今日の日付で判定）
function getWeekFromImport_() {
  return getWeekForDate_(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), null);
}

// 行事の取込（教員のみ・画面のボタンから呼ばれる）
function api_syncEvents() {
  var ctx = getContext_();
  if (ctx.role !== 'teacher') throw new Error('行事の取込は教員のみ行えます。');
  return syncEventsFromMaster();
}

// 毎日早朝に自動で取り込みたい場合は、エディタからこの関数を一度実行してトリガーを作る
function setupEventSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncEventsFromMaster') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncEventsFromMaster').timeBased().everyDays(1).atHour(5).create();
}

/* ---------- 係・担当の管理 ---------- */

// このクラスの担当一覧（教科係・係・委員会）。生徒も自分のクラスの分は閲覧できる。
function api_getAssignments(cls) {
  var ctx = getContext_();
  assertCanView_(ctx, cls);
  return readSheet_(SHEET.ASSIGNMENTS)
    .filter(function (a) { return a['クラス'] === cls; })
    .map(function (a) {
      return {
        category: a['カテゴリ'], label: a['表示名'],
        email: a['担当者メール'], name: a['担当者氏名（任意）'] || ''
      };
    });
}

// このクラスで初回登録を済ませた生徒の一覧（担当を選ぶプルダウン用・教員のみ）
function api_getClassStudents(cls) {
  var ctx = getContext_();
  if (ctx.role !== 'teacher') throw new Error('教員のみ利用できます。');
  return readSheet_(SHEET.USERS)
    .filter(function (u) { return u['クラス'] === cls && u['メールアドレス'] && !isTeacherEmail_(u['メールアドレス']); })
    .map(function (u) { return { email: u['メールアドレス'], name: u['氏名'] || u['メールアドレス'] }; });
}

// 担当を追加する（教員のみ）。「担当割当」シートに1行追加される。
function api_addAssignment(payload) {
  var ctx = getContext_();
  if (ctx.role !== 'teacher') throw new Error('担当の設定は教員のみ行えます。');

  var cls = payload && payload.cls;
  var category = payload && payload.category;
  var label = payload && payload.label ? String(payload.label).trim() : '';
  var email = payload && payload.email ? String(payload.email).trim() : '';

  if (!cls) throw new Error('クラスが指定されていません。');
  if (['subject', 'duty', 'committee'].indexOf(category) < 0) throw new Error('カテゴリの指定が正しくありません。');
  if (!label) throw new Error('係・委員会の名前を入力してください。');
  if (category === 'subject' && SUBJECTS.indexOf(label) < 0) {
    throw new Error('教科係の教科名は ' + SUBJECTS.join('・') + ' のいずれかにしてください。');
  }
  if (!email) throw new Error('担当する生徒を選択してください。');

  var dup = readSheet_(SHEET.ASSIGNMENTS).some(function (a) {
    return a['クラス'] === cls && a['カテゴリ'] === category
      && a['表示名'] === label && a['担当者メール'] === email;
  });
  if (dup) throw new Error('同じ担当がすでに登録されています。');

  var student = readSheet_(SHEET.USERS).filter(function (u) { return u['メールアドレス'] === email; })[0];
  var name = student ? (student['氏名'] || '') : '';

  appendRow_(SHEET.ASSIGNMENTS, {
    'クラス': cls, 'カテゴリ': category, '表示名': label,
    '担当者メール': email, '担当者氏名（任意）': name
  });
  return { ok: true };
}

// 担当を削除する（教員のみ）。「担当割当」シートから該当行を削除する。
function api_deleteAssignment(payload) {
  var ctx = getContext_();
  if (ctx.role !== 'teacher') throw new Error('担当の設定は教員のみ行えます。');

  var cls = payload && payload.cls;
  var category = payload && payload.category;
  var label = payload && payload.label;
  var email = payload && payload.email;
  if (!cls || !category || !label || !email) throw new Error('削除する担当が指定されていません。');

  var sheet = getSheet_(SHEET.ASSIGNMENTS);
  var headerRow = findHeaderRow_(sheet);
  var lastRow = sheet.getLastRow();
  if (!headerRow || lastRow <= headerRow) throw new Error('担当が登録されていません。');

  var values = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, sheet.getLastColumn()).getValues();
  var headers = values[0];
  var clsIdx = headers.indexOf('クラス');
  var catIdx = headers.indexOf('カテゴリ');
  var labelIdx = headers.indexOf('表示名');
  var emailIdx = headers.indexOf('担当者メール');

  for (var i = 1; i < values.length; i++) {
    if (values[i][clsIdx] === cls && values[i][catIdx] === category
        && values[i][labelIdx] === label && values[i][emailIdx] === email) {
      sheet.deleteRow(headerRow + i);
      return { ok: true };
    }
  }
  throw new Error('該当する担当が見つかりませんでした。すでに削除されている可能性があります。');
}

function api_getTasks(cls) {
  var ctx = getContext_();
  assertCanView_(ctx, cls);

  var doneSet = {};
  if (ctx.role === 'student') {
    readSheet_(SHEET.SUBMISSIONS)
      .filter(function (s) { return s['生徒メール'] === ctx.email; })
      .forEach(function (s) {
        if (s['完了フラグ'] === true || s['完了フラグ'] === 'TRUE') doneSet[s['課題ID']] = true;
      });
  }

  return readSheet_(SHEET.TASKS)
    .filter(function (t) { return t['クラス'] === cls; })
    .map(function (t) {
      return {
        id: String(t['ID']), cls: t['クラス'], category: t['カテゴリ'], label: t['表示名'],
        title: t['タイトル'], due: toDateStr_(t['締切日']), detail: t['詳細'] || '',
        postedBy: t['投稿者表示名'] || '', done: !!doneSet[t['ID']]
      };
    })
    .sort(function (a, b) { return a.due.localeCompare(b.due); });
}

function api_getAnnouncements(cls) {
  var ctx = getContext_();
  assertCanView_(ctx, cls);
  return readSheet_(SHEET.POSTS)
    .filter(function (p) { return p['クラス'] === cls; })
    .map(function (p) {
      return {
        id: String(p['ID']), cls: p['クラス'], category: p['カテゴリ'], label: p['表示名'],
        title: p['タイトル'], body: p['本文'] || '', postedBy: p['投稿者表示名'] || '',
        time: toDateTimeLabel_(p['投稿日時']), _sort: p['投稿日時'] instanceof Date ? p['投稿日時'].getTime() : 0
      };
    })
    .sort(function (a, b) { return b._sort - a._sort; })
    .map(function (p) { delete p._sort; return p; });
}

function api_getPostPermissions(cls, kind) {
  var ctx = getContext_();
  if (ctx.role !== 'teacher' && ctx.myClass !== cls) return [];
  return getPostPermissions_(ctx, cls, kind);
}

function api_postTask(payload) {
  var ctx = getContext_();
  var cls = payload && payload.cls;
  if (!cls) throw new Error('クラスが指定されていません。');
  if (ctx.role !== 'teacher' && ctx.myClass !== cls) throw new Error('このクラスには投稿できません。');

  var perms = getPostPermissions_(ctx, cls, 'tasks');
  var allowed = perms.some(function (p) { return p.category === payload.category && p.label === payload.label; });
  if (!allowed) throw new Error('その立場では課題・提出物を投稿できません。');

  var title = payload.title ? String(payload.title).trim() : '';
  if (!title) throw new Error('タイトルを入力してください。');
  if (!payload.due) throw new Error('締切日を入力してください。');

  var id = Utilities.getUuid();
  appendRow_(SHEET.TASKS, {
    'ID': id, 'クラス': cls, 'カテゴリ': payload.category, '表示名': payload.label,
    'タイトル': title, '締切日': payload.due,
    '詳細': payload.detail ? String(payload.detail).trim() : '',
    '投稿者メール': ctx.email, '投稿者表示名': ctx.name, '投稿日時': new Date()
  });
  return { ok: true, id: id };
}

function api_postAnnouncement(payload) {
  var ctx = getContext_();
  var cls = payload && payload.cls;
  if (!cls) throw new Error('クラスが指定されていません。');
  if (ctx.role !== 'teacher' && ctx.myClass !== cls) throw new Error('このクラスには投稿できません。');

  var perms = getPostPermissions_(ctx, cls, 'announcement');
  var allowed = perms.some(function (p) { return p.category === payload.category && p.label === payload.label; });
  if (!allowed) throw new Error('その立場では連絡を投稿できません。');

  var title = payload.title ? String(payload.title).trim() : '';
  if (!title) throw new Error('タイトルを入力してください。');

  var id = Utilities.getUuid();
  appendRow_(SHEET.POSTS, {
    'ID': id, 'クラス': cls, 'カテゴリ': payload.category, '表示名': payload.label,
    'タイトル': title, '本文': payload.body ? String(payload.body).trim() : '',
    '投稿者メール': ctx.email, '投稿者表示名': ctx.name, '投稿日時': new Date()
  });
  return { ok: true, id: id };
}

/* ---------- 今日・明日の教科連絡（日別連絡） ---------- */

// 「日別連絡」シートがなければ作る（注意書き＋見出し行つき）
function getOrCreateDailyNotesSheet_() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET.DAILY_NOTES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.DAILY_NOTES);
    sheet.getRange(1, 1, 2, 7).setValues([
      ['※ 今日・明日の教科連絡や持ち物は、このシートに日付ごとに保存されます（画面から入力します）。', '', '', '', '', '', ''],
      ['日付', 'クラス', '時限', '教科', '連絡', '入力者メール', '更新日時']
    ]);
  }
  return sheet;
}

// クラス名から学年を読み取る（例：1年1組・１年２組 → 1。読み取れなければ 0）
function gradeFromClassName_(cls) {
  var m = toHalfWidthDigits_(String(cls)).match(/(\d)\s*年/);
  return m ? Number(m[1]) : 0;
}

// 「今日」と「次の登校日（通常は明日。金曜なら月曜）」の日付・曜日・週区分・
// その日の校時（行事予定ファイル由来）と、そのクラスの日別連絡・時間割変更を返す
function api_getDayInfo(cls) {
  var ctx = getContext_();
  assertCanView_(ctx, cls);
  var tz = Session.getScriptTimeZone();

  var byDate = {}, rowByDate = {};
  try {
    readImportedEvents_().forEach(function (e) {
      if (e.week) byDate[e.date] = e.week;
      rowByDate[e.date] = e;
    });
  } catch (e) {}
  var fallbackWeek = getSetting_('今週の週区分', 'A');
  var grade = gradeFromClassName_(cls);

  function planFor(dateStr) {
    var row = rowByDate[dateStr];
    return (grade && row && row.plans && row.plans[grade]) ? row.plans[grade] : null;
  }

  // 行事予定の校時に授業が入っている日か（土曜の公開授業なども登校日として扱う）
  function hasSchoolPlan(dateStr) {
    var plan = planFor(dateStr);
    return !!(plan && plan.some(function (t) { return String(t).trim() !== ''; }));
  }

  function dayEntry(d, label) {
    var dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    var dow = d.getDay();
    var isWeekday = dow >= 1 && dow <= 5;
    return {
      date: dateStr,
      label: label,
      dayKey: (isWeekday || hasSchoolPlan(dateStr)) ? YOUBI_CHARS_[dow] : '',
      week: getWeekForDate_(dateStr, byDate) || fallbackWeek,
      // この日の校時（時限ごとの「A水3」「授」「学活」など）。行事予定にない日は null
      plan: planFor(dateStr)
    };
  }

  var today = new Date();
  var next = new Date(today);
  next.setDate(today.getDate() + 1);
  var isLiterallyTomorrow = true;
  // 土日は飛ばして次の登校日へ（ただし校時に授業が入っている土曜などは登校日扱い）
  while ((next.getDay() === 0 || next.getDay() === 6)
      && !hasSchoolPlan(Utilities.formatDate(next, tz, 'yyyy-MM-dd'))) {
    next.setDate(next.getDate() + 1);
    isLiterallyTomorrow = false;
  }
  var days = [dayEntry(today, '今日'), dayEntry(next, isLiterallyTomorrow ? '明日' : '次の登校日')];

  var todayStr = days[0].date;
  var notes = [];
  if (getSS_().getSheetByName(SHEET.DAILY_NOTES)) {
    readSheet_(SHEET.DAILY_NOTES).forEach(function (r) {
      var date = toDateStr_(r['日付']);
      if (r['クラス'] !== cls || !date || date < todayStr) return;
      notes.push({
        date: date, period: Number(r['時限']),
        subject: r['教科'] || '', text: r['連絡'] || ''
      });
    });
  }

  // 教員が入力した日付ごとの時間割変更（今日以降の分）
  var overrides = [];
  if (getSS_().getSheetByName(SHEET.SCHEDULE_CHANGES)) {
    readSheet_(SHEET.SCHEDULE_CHANGES).forEach(function (r) {
      var date = toDateStr_(r['日付']);
      if (r['クラス'] !== cls || !date || date < todayStr) return;
      overrides.push({ date: date, period: Number(r['時限']), subject: String(r['教科'] || '').trim() });
    });
  }

  return { days: days, notes: notes, overrides: overrides };
}

// 「時間割変更」シートがなければ作る
function getOrCreateScheduleChangesSheet_() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET.SCHEDULE_CHANGES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.SCHEDULE_CHANGES);
    sheet.getRange(1, 1, 2, 6).setValues([
      ['※ その日だけの時間割変更は、このシートに日付ごとに保存されます（教員が画面から入力します）。', '', '', '', '', ''],
      ['日付', 'クラス', '時限', '教科', '入力者メール', '更新日時']
    ]);
  }
  return sheet;
}

// その日だけの時間割変更を入力する（教員のみ・教科を空欄にすると元に戻す）。
// 固定の「時間割」シートは変更せず、「時間割変更」シートに日付つきで保存する。
function api_setScheduleOverride(payload) {
  var ctx = getContext_();
  if (ctx.role !== 'teacher') throw new Error('時間割変更は教員のみ行えます。');

  var cls = payload && payload.cls;
  var date = payload && payload.date ? String(payload.date).trim() : '';
  var period = payload && Number(payload.period);
  var subject = payload && payload.subject != null ? String(payload.subject).trim() : '';
  if (!cls || !date || !period) throw new Error('対象の時限が指定されていません。');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日付の形式が正しくありません。');

  var sheet = getOrCreateScheduleChangesSheet_();
  var headerRow = findHeaderRow_(sheet);
  if (!headerRow) throw new Error('「時間割変更」シートの見出し行が見つかりません。');
  var lastRow = sheet.getLastRow();
  if (lastRow > headerRow) {
    var values = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, sheet.getLastColumn()).getValues();
    var headers = values[0];
    var dateIdx = headers.indexOf('日付');
    var clsIdx = headers.indexOf('クラス');
    var periodIdx = headers.indexOf('時限');
    var subjectIdx = headers.indexOf('教科');
    var emailIdx = headers.indexOf('入力者メール');
    var timeIdx = headers.indexOf('更新日時');
    for (var i = 1; i < values.length; i++) {
      if (toDateStr_(values[i][dateIdx]) === date && values[i][clsIdx] === cls
          && Number(values[i][periodIdx]) === period) {
        var rowNum = headerRow + i;
        if (subject === '') {
          sheet.deleteRow(rowNum);
        } else {
          sheet.getRange(rowNum, subjectIdx + 1).setValue(subject);
          sheet.getRange(rowNum, emailIdx + 1).setValue(ctx.email);
          sheet.getRange(rowNum, timeIdx + 1).setValue(new Date());
        }
        return { ok: true };
      }
    }
  }
  if (subject === '') return { ok: true };  // 削除対象がなければ何もしない
  appendRow_(SHEET.SCHEDULE_CHANGES, {
    '日付': date, 'クラス': cls, '時限': period, '教科': subject,
    '入力者メール': ctx.email, '更新日時': new Date()
  });
  return { ok: true };
}

// 日付を指定して教科連絡・持ち物を入力する（空欄で削除）。
// 教員は全クラス・全時限、教科係の生徒は自分のクラスの担当教科の時限のみ入力できる。
function api_setDailyNote(payload) {
  var ctx = getContext_();

  var cls = payload && payload.cls;
  var date = payload && payload.date ? String(payload.date).trim() : '';
  var period = payload && Number(payload.period);
  var subject = payload && payload.subject != null ? String(payload.subject).trim() : '';
  var text = payload && payload.text != null ? String(payload.text).trim() : '';
  if (!cls || !date || !period) throw new Error('対象の時限が指定されていません。');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日付の形式が正しくありません。');

  if (ctx.role !== 'teacher') {
    if (ctx.myClass !== cls) throw new Error('このクラスの連絡は入力できません。');
    var isRep = readSheet_(SHEET.ASSIGNMENTS).some(function (a) {
      return a['クラス'] === cls && a['カテゴリ'] === 'subject'
        && a['表示名'] === subject && a['担当者メール'] === ctx.email;
    });
    if (!isRep) throw new Error('「' + subject + '」の教科係に登録されている生徒のみ入力できます。');
  }

  var sheet = getOrCreateDailyNotesSheet_();
  var headerRow = findHeaderRow_(sheet);
  if (!headerRow) throw new Error('「日別連絡」シートの見出し行が見つかりません。');
  var lastRow = sheet.getLastRow();
  if (lastRow > headerRow) {
    var values = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, sheet.getLastColumn()).getValues();
    var headers = values[0];
    var dateIdx = headers.indexOf('日付');
    var clsIdx = headers.indexOf('クラス');
    var periodIdx = headers.indexOf('時限');
    var subjectIdx = headers.indexOf('教科');
    var textIdx = headers.indexOf('連絡');
    var emailIdx = headers.indexOf('入力者メール');
    var timeIdx = headers.indexOf('更新日時');
    for (var i = 1; i < values.length; i++) {
      if (toDateStr_(values[i][dateIdx]) === date && values[i][clsIdx] === cls
          && Number(values[i][periodIdx]) === period) {
        var rowNum = headerRow + i;
        if (text === '') {
          sheet.deleteRow(rowNum);
        } else {
          sheet.getRange(rowNum, subjectIdx + 1).setValue(subject);
          sheet.getRange(rowNum, textIdx + 1).setValue(text);
          sheet.getRange(rowNum, emailIdx + 1).setValue(ctx.email);
          sheet.getRange(rowNum, timeIdx + 1).setValue(new Date());
        }
        return { ok: true };
      }
    }
  }
  if (text === '') return { ok: true };  // 削除対象がなければ何もしない
  appendRow_(SHEET.DAILY_NOTES, {
    '日付': date, 'クラス': cls, '時限': period, '教科': subject,
    '連絡': text, '入力者メール': ctx.email, '更新日時': new Date()
  });
  return { ok: true };
}

function api_setTaskDone(taskId, done) {
  var ctx = getContext_();
  if (ctx.role !== 'student') throw new Error('生徒のみチェックできます。');

  var sheet = getSheet_(SHEET.SUBMISSIONS);
  var headerRow = findHeaderRow_(sheet);
  var lastRow = sheet.getLastRow();
  if (headerRow && lastRow > headerRow) {
    var values = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, sheet.getLastColumn()).getValues();
    var headers = values[0];
    var idIdx = headers.indexOf('課題ID');
    var emailIdx = headers.indexOf('生徒メール');
    var doneIdx = headers.indexOf('完了フラグ');
    var timeIdx = headers.indexOf('更新日時');
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idIdx]) === String(taskId) && values[i][emailIdx] === ctx.email) {
        var rowNum = headerRow + i;
        sheet.getRange(rowNum, doneIdx + 1).setValue(!!done);
        sheet.getRange(rowNum, timeIdx + 1).setValue(new Date());
        return { ok: true };
      }
    }
  }
  appendRow_(SHEET.SUBMISSIONS, {
    '課題ID': taskId, '生徒メール': ctx.email, '完了フラグ': !!done, '更新日時': new Date()
  });
  return { ok: true };
}
