/*************************************************************
 * 読みの観点システム（系統表対応版）
 * 西原中学校 国語科
 *
 * ■ この版で追加したこと（系統表のねらいを反映）
 *  1) 系統表：学年 × 単元 × 観点 のマトリクス。
 *     「先生のねらい（◎）」＝単元マスタの「推奨観点」を、
 *     生徒の記録（深さ1〜3／☆）と同じ表に重ねて表示する。
 *  2) 観点ごとの具体例：HTML側（index_*.html）に「例」と「書き出しの型」を明示。
 *  3) 文字数の指定：単元マスタの「文字数目安」「振り返り文字数」で単元ごとに指定。
 *     記述は観点ごとに1列ずつ、シートに保存する。
 *  4) 総括的な振り返り：単元ごとの「単元のふりかえり」を保存し、
 *     系統表の下に「これまでの歩み」としてまとめて返す。
 *
 * ■ 照合の考え方（C案：両対応）
 *  - 教員：ログインのメールが「教員名簿」にあれば教員と判定。
 *  - 生徒：メールが「生徒名簿」にあればメールで自動判定。
 *          メールが空でも、生徒が初回にクラス＋番号を選ぶと名簿の氏名に結びつく。
 *  - 記録は「ログインのメール＋単元ID」で1行に集約（上書き）。
 *  - 教員ダッシュボードは、クラス＋番号で生徒の履歴を引く（名簿メールが空でも可）。
 *
 * ■ 既存シートは自動で列を追加します（並び替え・作り直しは不要）。
 *************************************************************/

// ===== 設定：ここだけ書き換える =====
const CONFIG = {
  RESULT_SS_ID: '',            // 記録先SSのID（このスクリプトを紐付けたSSを使うなら空 ''）
  ROSTER_SS_ID: '1L8RbP6aGUVPXPLfir6uebfjd-EXIXgTodVqVV66EdNw',            // 名簿マスターのSS ID（必須）

  TEACHER_SHEET: '教員名簿',
  STUDENT_SHEET: '生徒名簿',
  TCOL: { name: '教員名', email: 'メールアドレス', role: '役割' },
  SCOL: { klass: 'クラス', number: '番号', name: '氏名', email: 'メールアドレス', grade: '学年' },

  // 名簿に載っていない教員を足したいときだけ使う（任意）
  TEACHER_EMAILS: []
};

// ===== 観点マスタ（HTMLと必ず一致）=====
// 系統表と同じ「層 → 観点」の並び。ここを直すとシートの列も自動で増えます。
const LAYERS = {
  bungaku: [
    ['登場人物', '設定', '構成', '視点・語り'],
    ['表現の工夫', '題名', '主題'],
    ['批評・評価', '自分に生かす']
  ],
  setsumei: [
    ['話題と問い', '文章の構成', '要点と要旨'],
    ['主張と根拠', '事例のえらび方', '図表・写真'],
    ['筆者の工夫', '批判的な読み', '自分の考え']
  ]
};
const OBSERVATIONS = {
  bungaku:  [].concat.apply([], LAYERS.bungaku),
  setsumei: [].concat.apply([], LAYERS.setsumei)
};

const REC_SHEET  = { bungaku: '文学_記録', setsumei: '説明文_記録' };
const UNIT_SHEET = '単元マスタ';
const LABEL_GENRE = { '文学': 'bungaku', '説明文': 'setsumei' };
const PAGE_TITLE = { bungaku: '文学的文章　読みの観点', setsumei: '説明的文章　読みの観点', teacher: '教員用　読みの観点ダッシュボード' };

// 記述欄の設定
const MEMO_SUFFIX     = 'の記述';   // 記録シートの列名：「登場人物の記述」など
const REFLECT_COL     = '単元のふりかえり';
const DEFAULT_CHARS   = 100;        // 観点1つあたりの目安文字数
const DEFAULT_RCHARS  = 200;        // 単元のふりかえりの目安文字数
const MAX_TEXT        = 2000;       // 保存する記述の上限（安全のため）

/* ============================================================
 *  画面配信
 * ==========================================================*/
function doGet(e) {
  const p = (e && e.parameter && e.parameter.p) || 'menu';
  let file, title;
  if (p === 'bungaku')       { file = 'index_bungaku';  title = PAGE_TITLE.bungaku; }
  else if (p === 'setsumei') { file = 'index_setsumei'; title = PAGE_TITLE.setsumei; }
  else if (p === 'teacher')  { file = 'teacher';        title = PAGE_TITLE.teacher; }
  else                       { file = 'menu';           title = '読みの観点｜西原中学校 国語科'; }

  const appUrl = getAppUrl();
  let content = HtmlService.createHtmlOutputFromFile(file).getContent();
  content = content.replace('__APP_URL__', function () { return appUrl; });

  return HtmlService.createHtmlOutput(content)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getAppUrl() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/* ============================================================
 *  共通の小道具
 * ==========================================================*/
function safeEmail() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}
function rosterSS() { return SpreadsheetApp.openById(CONFIG.ROSTER_SS_ID); }
function colIndex(head, name) { return head.map(String).indexOf(name); }
function isValidEmail(s) { return /^[^@\s]+@[^@\s]+/.test(String(s || '').trim()); }
function normNum(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isNaN(n) ? String(v).trim() : String(Math.round(n));
}
function normKlass(v) { return String(v || '').trim(); }
function toHankaku(s) {
  return String(s == null ? '' : s).replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
}
// 「1」「1年」「１年２組」→「1年」。数字が無ければそのまま返す。
function normGrade(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  const m = toHankaku(s).match(/\d+/);
  return m ? m[0] + '年' : s;
}
function gradeNum(g) {
  const m = toHankaku(g).match(/\d+/);
  return m ? Number(m[0]) : 99;
}
function clipText(s) {
  const t = String(s == null ? '' : s);
  return t.length > MAX_TEXT ? t.slice(0, MAX_TEXT) : t;
}
// 見出し行に足りない列を右端に足す（既存データは動かさない）
function ensureHeaders(sh, header) {
  const lastCol = sh.getLastColumn();
  let head = (lastCol > 0 && sh.getLastRow() > 0) ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  while (head.length && String(head[head.length - 1]).trim() === '') head.pop();
  if (!head.length) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return header.slice();
  }
  const add = header.filter(function (h) { return head.indexOf(h) < 0; });
  if (add.length) {
    sh.getRange(1, head.length + 1, 1, add.length).setValues([add]).setFontWeight('bold');
    head = head.concat(add);
  }
  return head;
}

/* ============================================================
 *  名簿マスター読み取り
 * ==========================================================*/
// 教員名簿 → [{email, name, role}]
function getTeacherRows() {
  if (!CONFIG.ROSTER_SS_ID) return [];
  try {
    const sh = rosterSS().getSheetByName(CONFIG.TEACHER_SHEET);
    if (!sh) return [];
    const v = sh.getDataRange().getValues();
    if (v.length < 2) return [];
    const head = v[0];
    const iE = colIndex(head, CONFIG.TCOL.email), iN = colIndex(head, CONFIG.TCOL.name), iR = colIndex(head, CONFIG.TCOL.role);
    const out = [];
    for (let r = 1; r < v.length; r++) {
      const em = iE >= 0 ? String(v[r][iE] || '').trim() : '';
      if (!em) continue;
      out.push({ email: em, name: iN >= 0 ? String(v[r][iN] || '') : '', role: iR >= 0 ? String(v[r][iR] || '') : '' });
    }
    return out;
  } catch (err) { return []; }
}

// 生徒名簿 → [{email, klass, number, name, grade}]
// 「学年」列が無ければクラス名（例：１年２組）から学年を取り出す。
function getStudentRows() {
  if (!CONFIG.ROSTER_SS_ID) return [];
  try {
    const sh = rosterSS().getSheetByName(CONFIG.STUDENT_SHEET);
    if (!sh) return [];
    const v = sh.getDataRange().getValues();
    if (v.length < 2) return [];
    const head = v[0];
    const iK = colIndex(head, CONFIG.SCOL.klass), iN = colIndex(head, CONFIG.SCOL.number),
          iM = colIndex(head, CONFIG.SCOL.name), iE = colIndex(head, CONFIG.SCOL.email),
          iG = colIndex(head, CONFIG.SCOL.grade);
    const out = [];
    for (let r = 1; r < v.length; r++) {
      const nm = iM >= 0 ? String(v[r][iM] || '').trim() : '';
      const kl = iK >= 0 ? normKlass(v[r][iK]) : '';
      if (!nm && !kl) continue;
      out.push({
        klass: kl,
        number: iN >= 0 ? normNum(v[r][iN]) : '',
        name: nm,
        email: iE >= 0 ? String(v[r][iE] || '').trim() : '',
        grade: normGrade(iG >= 0 ? v[r][iG] : kl)
      });
    }
    return out;
  } catch (err) { return []; }
}

function isTeacher(email) {
  if (!email) return false;
  const t = email.toLowerCase().trim();
  if ((CONFIG.TEACHER_EMAILS || []).some(x => String(x).toLowerCase().trim() === t)) return true;
  return getTeacherRows().some(r => r.email.toLowerCase().trim() === t);
}

/* ============================================================
 *  ユーザー情報（自動判定）
 * ==========================================================*/
function getUserInfo() {
  const email = safeEmail();
  const info = { email: email, klass: '', number: '', name: '', grade: '', role: '', resolved: false, isTeacher: false };

  // 教員判定（教員名簿）
  const trow = getTeacherRows().find(r => r.email.toLowerCase().trim() === email.toLowerCase().trim());
  if (trow || (CONFIG.TEACHER_EMAILS || []).some(x => String(x).toLowerCase().trim() === email.toLowerCase().trim())) {
    info.isTeacher = true;
    if (trow) { info.name = trow.name; info.role = trow.role; }
    info.resolved = true;   // 教員は氏名入力不要
    return info;
  }

  // 生徒：メールが名簿にあれば自動照合
  if (isValidEmail(email)) {
    const s = getStudentRows().find(x => isValidEmail(x.email) && x.email.toLowerCase().trim() === email.toLowerCase().trim());
    if (s) { info.klass = s.klass; info.number = s.number; info.name = s.name; info.grade = s.grade; info.resolved = true; return info; }
  }
  // メールで引けない → フロントでクラス＋番号を選んでもらう（resolved=false）
  return info;
}

// 生徒：クラス＋番号 → 名簿の氏名を返す
function resolveStudent(klass, number) {
  const kl = normKlass(klass), nu = normNum(number);
  const s = getStudentRows().find(x => x.klass === kl && x.number === nu);
  return s ? { ok: true, klass: kl, number: nu, name: s.name, grade: s.grade }
           : { ok: false, klass: kl, number: nu, name: '', grade: normGrade(kl) };
}

// 生徒名簿のクラス一覧（初回入力のプルダウン用）
function getRosterClasses() {
  const seen = {}, out = [];
  getStudentRows().forEach(s => { if (s.klass && !seen[s.klass]) { seen[s.klass] = true; out.push(s.klass); } });
  return out;
}

/* ============================================================
 *  単元マスタ（＝系統表のもと）
 * ==========================================================*/
const UNIT_HEADER = ['単元ID', '学年', 'ジャンル', '単元名', '教材名', '順序', '有効', '推奨観点', '文字数目安', '振り返り文字数'];

function getResultSS() {
  return CONFIG.RESULT_SS_ID ? SpreadsheetApp.openById(CONFIG.RESULT_SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function getUnitSheet() {
  const ss = getResultSS();
  let sh = ss.getSheetByName(UNIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(UNIT_SHEET);
    sh.getRange(1, 1, 1, UNIT_HEADER.length).setValues([UNIT_HEADER])
      .setFontWeight('bold').setBackground('#9a7b3f').setFontColor('#fff');
    sh.setFrozenRows(1);
    // 系統表のイメージをそのまま初期値に（各単元でねらう観点は2〜3個）
    const seed = [
      ['B01', '1年', '文学',   '（例）ベンチ',               'ベンチ',               1, false, '登場人物、視点・語り、題名',         DEFAULT_CHARS, DEFAULT_RCHARS],
      ['B02', '1年', '文学',   '（例）オツベルと象',         'オツベルと象',         2, false, '設定、表現の工夫、主題',             DEFAULT_CHARS, DEFAULT_RCHARS],
      ['B03', '1年', '文学',   '（例）少年の日の思い出',     '少年の日の思い出',     3, true,  '構成、批評・評価、自分に生かす',     DEFAULT_CHARS, DEFAULT_RCHARS],
      ['S01', '1年', '説明文', '（例）ダイコンは大きな根？', 'ダイコンは大きな根？', 1, true,  '話題と問い、文章の構成、要点と要旨', DEFAULT_CHARS, DEFAULT_RCHARS],
      ['S02', '1年', '説明文', '（例）モアイは語る',         'モアイは語る',         2, false, '文章の構成、主張と根拠、自分の考え', DEFAULT_CHARS, DEFAULT_RCHARS]
    ];
    sh.getRange(2, 1, seed.length, UNIT_HEADER.length).setValues(seed);
    sh.getRange('F:F').setNumberFormat('0');
    sh.setColumnWidth(4, 200); sh.setColumnWidth(8, 260);
  } else {
    ensureHeaders(sh, UNIT_HEADER);
  }
  return sh;
}

function splitObs(s) {
  return String(s || '').split(/[,、\s／\/]+/).filter(Boolean);
}

function getUnits(genre) {
  const g = (genre === 'setsumei') ? 'setsumei' : 'bungaku';
  const sh = getUnitSheet();
  const v = sh.getDataRange().getValues();
  const out = [];
  if (v.length < 2) return out;
  const head = v[0].map(String);
  const c = {
    id: head.indexOf('単元ID'), grade: head.indexOf('学年'), genre: head.indexOf('ジャンル'),
    name: head.indexOf('単元名'), material: head.indexOf('教材名'), order: head.indexOf('順序'),
    active: head.indexOf('有効'), rec: head.indexOf('推奨観点'),
    chars: head.indexOf('文字数目安'), rchars: head.indexOf('振り返り文字数')
  };
  for (let r = 1; r < v.length; r++) {
    if (LABEL_GENRE[String(v[r][c.genre] || '').trim()] !== g) continue;
    const id = String(v[r][c.id] || '').trim();
    if (!id) continue;
    const chars  = c.chars  >= 0 ? Number(v[r][c.chars]  || 0) : 0;
    const rchars = c.rchars >= 0 ? Number(v[r][c.rchars] || 0) : 0;
    out.push({
      id: id,
      grade: c.grade >= 0 ? normGrade(v[r][c.grade]) : '',
      name: String(v[r][c.name] || '').trim() || id,
      material: c.material >= 0 ? String(v[r][c.material] || '').trim() : '',
      order: c.order >= 0 ? Number(v[r][c.order] || 0) : 0,
      active: c.active >= 0 ? (v[r][c.active] === true || String(v[r][c.active]).toUpperCase() === 'TRUE' || v[r][c.active] === 1) : false,
      recommend: c.rec >= 0 ? splitObs(v[r][c.rec]) : [],
      chars: chars > 0 ? Math.round(chars) : DEFAULT_CHARS,
      reflectChars: rchars > 0 ? Math.round(rchars) : DEFAULT_RCHARS
    });
  }
  // 系統表の並び：学年 → 順序 → 単元ID
  out.sort(function (a, b) {
    return (gradeNum(a.grade) - gradeNum(b.grade)) || (a.order - b.order) || a.id.localeCompare(b.id);
  });
  return out;
}

// 系統表（先生のねらいだけ／生徒データなし）。教員画面などから使える。
function getSyllabus(genre) {
  const g = (genre === 'setsumei') ? 'setsumei' : 'bungaku';
  return { genre: g, layers: LAYERS[g], observations: OBSERVATIONS[g], units: getUnits(g) };
}

/* ============================================================
 *  記録：保存 / 読み込み（単元別）
 * ==========================================================*/
function recHeader(genre) {
  const obs = OBSERVATIONS[genre];
  return ['更新日時', 'メール', '学年', 'クラス', '番号', '氏名', '単元ID', '単元名', '選んだ観点']
    .concat(obs)                                                // 深さ 1〜3
    .concat(obs.map(function (o) { return o + MEMO_SUFFIX; }))  // 観点ごとの記述
    .concat([REFLECT_COL]);                                     // 総括の振り返り
}
function getRecSheet(genre) {
  const ss = getResultSS();
  const name = REC_SHEET[genre];
  const header = recHeader(genre);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#20303f').setFontColor('#fff');
    sh.setFrozenRows(1);
  } else {
    ensureHeaders(sh, header);   // 旧シートに「◯◯の記述」「学年」などを自動追加
  }
  return sh;
}

// payload: {genre, unitId, unitName, picked:[], levels:{}, memos:{}, reflection:'', klass, number, name, grade}
function saveRecord(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) { return { ok: false, message: '混み合っています。もう一度試してね。' }; }
  try {
    const genre = (payload && payload.genre === 'setsumei') ? 'setsumei' : 'bungaku';
    const obs = OBSERVATIONS[genre];
    const unitId = String(payload.unitId || '').trim();
    if (!unitId) return { ok: false, message: '単元が選ばれていません。' };

    const sh = getRecSheet(genre);
    const info = getUserInfo();
    const email = info.email;
    const klass  = normKlass(info.klass  || payload.klass  || '');
    const number = normNum(info.number || payload.number || '');
    const name   = info.name   || payload.name   || '';
    const grade  = normGrade(info.grade || payload.grade || klass);

    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const cE = head.indexOf('メール'), cK = head.indexOf('クラス'), cN = head.indexOf('番号'), cU = head.indexOf('単元ID');

    // 先に既存行をさがす（メールがあれば「メール＋単元」、無ければ「クラス＋番号＋単元」）
    const data = sh.getDataRange().getValues();
    const keyE = email.toLowerCase().trim();
    let target = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][cU] || '').trim() !== unitId) continue;
      const rowE = String(data[r][cE] || '').toLowerCase().trim();
      const sameEmail = keyE && rowE === keyE;
      const sameCN = !keyE && klass && String(data[r][cK] || '').trim() === klass && normNum(data[r][cN]) === number;
      if (sameEmail || sameCN) { target = r + 1; break; }
    }

    // 既存の値を土台にする（送られてこなかった項目は消さない）
    const row = new Array(head.length).fill('');
    if (target > 0) {
      for (let i = 0; i < head.length; i++) row[i] = (data[target - 1][i] === undefined) ? '' : data[target - 1][i];
    }

    const set = function (col, val) { const i = head.indexOf(col); if (i >= 0) row[i] = val; };
    set('更新日時', new Date());
    set('メール', email);
    set('学年', grade);
    set('クラス', klass);
    set('番号', number);
    set('氏名', name);
    set('単元ID', unitId);
    set('単元名', payload.unitName || '');

    if (Array.isArray(payload.picked)) set('選んだ観点', payload.picked.join('、'));

    if (payload.levels && typeof payload.levels === 'object') {
      const levels = payload.levels;
      obs.forEach(function (o) {
        const v = levels[o];
        set(o, (v === 1 || v === 2 || v === 3) ? v : '');   // 選び直し・取り消しも反映
      });
    }
    if (payload.memos && typeof payload.memos === 'object') {
      const memos = payload.memos;
      obs.forEach(function (o) { set(o + MEMO_SUFFIX, clipText(memos[o] || '')); });
    }
    if (typeof payload.reflection === 'string') set(REFLECT_COL, clipText(payload.reflection));

    if (target > 0) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: String(err) };
  } finally {
    lock.releaseLock();
  }
}

// filter = {email} または {klass, number}（null なら全件）
function readRows(genre, filter) {
  const g = (genre === 'setsumei') ? 'setsumei' : 'bungaku';
  const sh = getRecSheet(g);
  const v = sh.getDataRange().getValues();
  const obs = OBSERVATIONS[g];
  const out = [];
  if (v.length < 2) return out;
  const head = v[0].map(String);
  const cE = head.indexOf('メール'), cK = head.indexOf('クラス'), cN = head.indexOf('番号'),
        cU = head.indexOf('単元ID'), cName = head.indexOf('単元名'), cP = head.indexOf('選んだ観点'),
        cT = head.indexOf('更新日時'), cR = head.indexOf(REFLECT_COL);
  const fEmail = filter && filter.email ? String(filter.email).toLowerCase().trim() : '';
  const fKlass = filter && filter.klass ? normKlass(filter.klass) : '';
  const fNum   = filter && (filter.number !== undefined) ? normNum(filter.number) : '';
  for (let r = 1; r < v.length; r++) {
    if (fEmail) { if (String(v[r][cE] || '').toLowerCase().trim() !== fEmail) continue; }
    else if (fKlass) { if (normKlass(v[r][cK]) !== fKlass || normNum(v[r][cN]) !== fNum) continue; }
    const levels = {}, memos = {};
    obs.forEach(function (o) {
      const ci = head.indexOf(o); const val = ci >= 0 ? v[r][ci] : '';
      if (val === 1 || val === 2 || val === 3) levels[o] = val;
      else if (typeof val === 'string' && /^[123]$/.test(val.trim())) levels[o] = Number(val.trim());
      const mi = head.indexOf(o + MEMO_SUFFIX);
      const mv = mi >= 0 ? String(v[r][mi] || '') : '';
      if (mv) memos[o] = mv;
    });
    out.push({
      unitId: String(v[r][cU] || '').trim(), unitName: String(v[r][cName] || '').trim(),
      picked: String(v[r][cP] || '').split(/[,、\s]+/).filter(Boolean), levels: levels, memos: memos,
      reflection: cR >= 0 ? String(v[r][cR] || '') : '',
      updated: cT >= 0 ? v[r][cT] : ''
    });
  }
  return out;
}

/* ============================================================
 *  系統表つきの履歴
 *   - units / counts は従来どおり（教員画面の互換のため）
 *   - syllabus / records / best / summary を追加
 * ==========================================================*/
function buildHistory(genre, filter) {
  const g = (genre === 'setsumei') ? 'setsumei' : 'bungaku';
  const units = getUnits(g);
  const rows = readRows(g, filter);
  const byUnit = {}; rows.forEach(function (r) { byUnit[r.unitId] = r; });

  // 従来の list（記録のある単元だけ）
  const seen = {}, list = [];
  units.forEach(function (u) {
    seen[u.id] = true;
    const rec = byUnit[u.id];
    if (rec) list.push({ id: u.id, name: u.name || rec.unitName, grade: u.grade, picked: rec.picked, levels: rec.levels,
                         memos: rec.memos, reflection: rec.reflection, updated: rec.updated });
  });
  rows.forEach(function (r) {
    if (!seen[r.unitId] && r.unitId) list.push({ id: r.unitId, name: r.unitName || r.unitId, grade: '', picked: r.picked,
                                                 levels: r.levels, memos: r.memos, reflection: r.reflection, updated: r.updated });
  });

  const obs = OBSERVATIONS[g];
  const counts = {}, best = {};
  obs.forEach(function (o) { counts[o] = 0; best[o] = 0; });
  list.forEach(function (row) {
    obs.forEach(function (o) {
      if (row.levels[o]) { counts[o]++; if (row.levels[o] > best[o]) best[o] = row.levels[o]; }
    });
  });

  // 系統表の列：単元マスタの全単元（＝これから学ぶ単元も並ぶ）
  const syllabus = units.map(function (u) {
    return { id: u.id, name: u.name, material: u.material, grade: u.grade, order: u.order, active: u.active, recommend: u.recommend };
  });
  list.forEach(function (row) {
    if (!syllabus.some(function (u) { return u.id === row.id; })) {
      syllabus.push({ id: row.id, name: row.name, material: '', grade: row.grade || '', order: 999, active: false, recommend: [] });
    }
  });

  const records = {};
  list.forEach(function (row) {
    records[row.id] = { picked: row.picked, levels: row.levels, reflection: row.reflection, updated: row.updated };
  });

  const unused = obs.filter(function (o) { return !counts[o]; });
  const summary = {
    total: obs.length,
    done: obs.length - unused.length,
    unused: unused,
    deep: { 1: 0, 2: 0, 3: 0 },
    unitsDone: list.length,
    unitsAll: syllabus.length,
    reflections: list.filter(function (r) { return String(r.reflection || '').trim(); }).length
  };
  obs.forEach(function (o) { if (best[o]) summary.deep[best[o]]++; });

  return { genre: g, observations: obs, layers: LAYERS[g], units: list, counts: counts, best: best,
           syllabus: syllabus, records: records, summary: summary };
}

// 単一単元の記録（生徒の続き読み込み）
function getRecord(genre, unitId, klass, number) {
  const empty = { picked: [], levels: {}, memos: {}, reflection: '' };
  const email = safeEmail();
  let filter = null;
  if (isValidEmail(email)) filter = { email: email };
  else if (klass) filter = { klass: klass, number: number };
  if (!filter) return empty;
  const rows = readRows(genre, filter);
  const hit = rows.find(function (r) { return r.unitId === String(unitId).trim(); });
  return hit ? { picked: hit.picked, levels: hit.levels, memos: hit.memos, reflection: hit.reflection } : empty;
}

// 生徒：自分の履歴（ログインのメール、無ければクラス＋番号で引く）
function getMyHistory(genre, klass, number) {
  const email = safeEmail();
  const filter = isValidEmail(email) ? { email: email } : { klass: klass, number: number };
  return buildHistory(genre, filter);
}

/* ============================================================
 *  教員用
 * ==========================================================*/
function getTeacherContext() {
  const email = safeEmail();
  const t = getTeacherRows().find(r => r.email.toLowerCase().trim() === email.toLowerCase().trim());
  return { email: email, isTeacher: isTeacher(email), name: t ? t.name : '' };
}

// 教員：生徒一覧（生徒名簿）
function getStudents() {
  if (!isTeacher(safeEmail())) return { ok: false, message: '教員のみ利用できます。', students: [] };
  return { ok: true, students: getStudentRows() };  // [{klass, number, name, email, grade}]
}

// 教員：指定生徒の履歴（クラス＋番号で引く。名簿メールが空でも可）
function getStudentHistory(klass, number, genre) {
  if (!isTeacher(safeEmail())) return { ok: false, message: '教員のみ利用できます。' };
  const filter = { klass: klass, number: number };
  const g = (genre === 'setsumei') ? 'setsumei' : (genre === 'bungaku' ? 'bungaku' : null);
  const res = { ok: true };
  if (g) res.data = buildHistory(g, filter);
  else { res.bungaku = buildHistory('bungaku', filter); res.setsumei = buildHistory('setsumei', filter); }
  return res;
}

// 教員：クラス全体の系統表（だれが、どの観点を、どこまで使えたか）
// 返り値 students[] = {klass, number, name, best:{観点:最高の深さ}, done:経験した観点数}
function getClassSyllabus(klass, genre) {
  if (!isTeacher(safeEmail())) return { ok: false, message: '教員のみ利用できます。' };
  const g = (genre === 'setsumei') ? 'setsumei' : 'bungaku';
  const obs = OBSERVATIONS[g];
  const roster = getStudentRows().filter(function (s) { return !klass || s.klass === normKlass(klass); });
  const list = roster.map(function (s) { return { klass: s.klass, number: s.number, name: s.name, best: {}, done: 0 }; });

  const sh = getRecSheet(g);
  const v = sh.getDataRange().getValues();
  if (v.length > 1) {
    const head = v[0].map(String);
    const cK = head.indexOf('クラス'), cN = head.indexOf('番号'), cE = head.indexOf('メール');
    const byKey = {};
    roster.forEach(function (s, i) {
      byKey[s.klass + '/' + s.number] = i;
      if (s.email) byKey[String(s.email).toLowerCase().trim()] = i;
    });
    for (let r = 1; r < v.length; r++) {
      let idx = byKey[normKlass(v[r][cK]) + '/' + normNum(v[r][cN])];
      if (idx === undefined) idx = byKey[String(v[r][cE] || '').toLowerCase().trim()];
      if (idx === undefined) continue;
      obs.forEach(function (o) {
        const ci = head.indexOf(o); if (ci < 0) return;
        const val = Number(v[r][ci]);
        if (val >= 1 && val <= 3 && val > (list[idx].best[o] || 0)) list[idx].best[o] = val;
      });
    }
    list.forEach(function (s) { s.done = obs.filter(function (o) { return s.best[o]; }).length; });
  }
  return { ok: true, genre: g, observations: obs, layers: LAYERS[g], units: getUnits(g), students: list };
}

/* ============================================================
 *  メニュー（任意）
 * ==========================================================*/
function onOpen() {
  SpreadsheetApp.getUi().createMenu('読みの観点')
    .addItem('単元マスタを開く（作成）', 'openUnitSheet')
    .addItem('記録シートの列を最新にする', 'upgradeSheets')
    .addToUi();
}
function openUnitSheet() { getResultSS().setActiveSheet(getUnitSheet()); }
function upgradeSheets() {
  getUnitSheet(); getRecSheet('bungaku'); getRecSheet('setsumei');
  SpreadsheetApp.getUi().alert('単元マスタと記録シートの列を最新にしました。');
}
