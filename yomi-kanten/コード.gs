/*************************************************************
 * 読みの観点システム（閲覧版・系統表対応）
 * 西原中学校 国語科
 *
 * ■ この版の考え方
 *  - 生徒の画面は「見るだけ」。ログインも入力も保存もしない。
 *  - 教員が単元ごとに「全員がおさえる観点（◎）」と「ひとこと」を決め、
 *    それが生徒の画面と系統表にそのまま出る。
 *  - データは「単元マスタ」シート1枚だけ。生徒の記録は一切とらない。
 *
 * ■ ファイル構成
 *   コード.gs            … これ
 *   index_bungaku.html   … 生徒用（文学的文章）※閲覧のみ
 *   index_setsumei.html  … 生徒用（説明的文章）※閲覧のみ
 *   teacher.html         … 教員用（おすすめ観点の設定・系統表の確認）
 *   menu.html            … 入口（任意）
 *
 * ■ 公開設定のめやす
 *   デプロイ →「次のユーザーとして実行：自分」
 *             「アクセスできるユーザー：組織内の全員（または全員）」
 *   ※教員画面は、ログインのメールが「教員名簿」にある人だけ編集できます。
 *     メールが取得できない公開設定のときは、教員画面は使えないので
 *     スプレッドシートの「単元マスタ」を直接編集してください。
 *************************************************************/

// ===== 設定：ここだけ書き換える =====
const CONFIG = {
  UNIT_SS_ID: '',          // 単元マスタを置くSSのID（このスクリプトを紐付けたSSなら空 ''）
  ROSTER_SS_ID: '1L8RbP6aGUVPXPLfir6uebfjd-EXIXgTodVqVV66EdNw',   // 教員名簿のSS ID（教員画面を使うとき）

  TEACHER_SHEET: '教員名簿',
  TCOL: { name: '教員名', email: 'メールアドレス', role: '役割' },

  // 名簿に載っていない教員を足したいときだけ（任意）
  TEACHER_EMAILS: []
};

// ===== 観点マスタ（HTMLと必ず一致）=====
// 系統表と同じ「層 → 観点」の並び。
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

const UNIT_SHEET  = '単元マスタ';
const UNIT_HEADER = ['単元ID', '学年', 'ジャンル', '単元名', '教材名', '順序', '有効', 'おさえる観点', 'ひとこと'];
const OLD_REC_COL = '推奨観点';   // 旧版の列名。あれば読み込みだけ引き継ぐ。
const LABEL_GENRE = { '文学': 'bungaku', '説明文': 'setsumei' };
const GENRE_LABEL = { bungaku: '文学', setsumei: '説明文' };
const PAGE_TITLE  = { bungaku: '文学的文章　読みの観点', setsumei: '説明的文章　読みの観点', teacher: '教員用　単元とおさえる観点の設定' };
const MAX_NOTE    = 400;   // 「ひとこと」の上限

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
 *  小道具
 * ==========================================================*/
function safeEmail() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}
function colIndex(head, name) { return head.map(String).indexOf(name); }
function toHankaku(s) {
  return String(s == null ? '' : s).replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
}
// 「1」「1年」「１年２組」→「1年」
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
function splitObs(s) {
  return String(s || '').split(/[,、\s／\/]+/).filter(Boolean);
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
 *  教員判定（教員名簿）
 * ==========================================================*/
function getTeacherRows() {
  if (!CONFIG.ROSTER_SS_ID) return [];
  try {
    const sh = SpreadsheetApp.openById(CONFIG.ROSTER_SS_ID).getSheetByName(CONFIG.TEACHER_SHEET);
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
function isTeacher(email) {
  if (!email) return false;
  const t = email.toLowerCase().trim();
  if ((CONFIG.TEACHER_EMAILS || []).some(function (x) { return String(x).toLowerCase().trim() === t; })) return true;
  return getTeacherRows().some(function (r) { return r.email.toLowerCase().trim() === t; });
}
function getTeacherContext() {
  const email = safeEmail();
  const t = getTeacherRows().find(function (r) { return r.email.toLowerCase().trim() === email.toLowerCase().trim(); });
  return { email: email, isTeacher: isTeacher(email), name: t ? t.name : '' };
}
function requireTeacher() {
  const email = safeEmail();
  if (!isTeacher(email)) {
    throw new Error(email ? '教員のみ編集できます。' : 'ログイン情報が取れません。スプレッドシートの「単元マスタ」を直接編集してください。');
  }
  return email;
}

/* ============================================================
 *  単元マスタ（＝系統表のもと）
 * ==========================================================*/
function getUnitSS() {
  return CONFIG.UNIT_SS_ID ? SpreadsheetApp.openById(CONFIG.UNIT_SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function getUnitSheet() {
  const ss = getUnitSS();
  let sh = ss.getSheetByName(UNIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(UNIT_SHEET);
    sh.getRange(1, 1, 1, UNIT_HEADER.length).setValues([UNIT_HEADER])
      .setFontWeight('bold').setBackground('#9a7b3f').setFontColor('#fff');
    sh.setFrozenRows(1);
    // 系統表のイメージをそのまま初期値に（各単元で全員がおさえる観点は2〜3個）
    const seed = [
      ['B01', '1年', '文学',   '（例）ベンチ',               'ベンチ',               1, false, '登場人物、視点・語り、題名',         ''],
      ['B02', '1年', '文学',   '（例）オツベルと象',         'オツベルと象',         2, false, '設定、表現の工夫、主題',             ''],
      ['B03', '1年', '文学',   '（例）少年の日の思い出',     '少年の日の思い出',     3, true,  '構成、批評・評価、自分に生かす',     '今回は「構成」を手がかりに、山場から人物の変化を読みます。'],
      ['S01', '1年', '説明文', '（例）ダイコンは大きな根？', 'ダイコンは大きな根？', 1, true,  '話題と問い、文章の構成、要点と要旨', '問いの文をさがすところから始めます。'],
      ['S02', '2年', '説明文', '（例）モアイは語る',         'モアイは語る',         2, false, '文章の構成、主張と根拠、自分の考え', '']
    ];
    sh.getRange(2, 1, seed.length, UNIT_HEADER.length).setValues(seed);
    sh.getRange('F:F').setNumberFormat('0');
    sh.setColumnWidth(4, 200); sh.setColumnWidth(8, 260); sh.setColumnWidth(9, 280);
  } else {
    ensureHeaders(sh, UNIT_HEADER);
  }
  return sh;
}
function unitCols(head) {
  return {
    id: head.indexOf('単元ID'), grade: head.indexOf('学年'), genre: head.indexOf('ジャンル'),
    name: head.indexOf('単元名'), material: head.indexOf('教材名'), order: head.indexOf('順序'),
    active: head.indexOf('有効'), rec: head.indexOf('おさえる観点'), recOld: head.indexOf(OLD_REC_COL),
    note: head.indexOf('ひとこと')
  };
}
function rowToUnit(row, c) {
  const id = String(row[c.id] || '').trim();
  if (!id) return null;
  return {
    id: id,
    grade: c.grade >= 0 ? normGrade(row[c.grade]) : '',
    genre: LABEL_GENRE[String(row[c.genre] || '').trim()] || '',
    name: String(row[c.name] || '').trim() || id,
    material: c.material >= 0 ? String(row[c.material] || '').trim() : '',
    order: c.order >= 0 ? Number(row[c.order] || 0) : 0,
    active: c.active >= 0 ? (row[c.active] === true || String(row[c.active]).toUpperCase() === 'TRUE' || row[c.active] === 1) : false,
    must: mustOf(row, c),
    note: c.note >= 0 ? String(row[c.note] || '').trim() : ''
  };
}
// 「おさえる観点」。旧「推奨観点」列しか無いシートからも読めるようにする。
function mustOf(row, c) {
  const cur = c.rec >= 0 ? splitObs(row[c.rec]) : [];
  if (cur.length) return cur;
  return c.recOld >= 0 ? splitObs(row[c.recOld]) : [];
}

function sortUnits(list) {
  // 系統表の並び：学年 → 順序 → 単元ID
  return list.sort(function (a, b) {
    return (gradeNum(a.grade) - gradeNum(b.grade)) || (a.order - b.order) || a.id.localeCompare(b.id);
  });
}

// 生徒・教員ともに使う：ジャンルごとの単元一覧
function getUnits(genre) {
  const g = (genre === 'setsumei') ? 'setsumei' : 'bungaku';
  const v = getUnitSheet().getDataRange().getValues();
  if (v.length < 2) return [];
  const c = unitCols(v[0].map(String));
  const out = [];
  for (let r = 1; r < v.length; r++) {
    const u = rowToUnit(v[r], c);
    if (u && u.genre === g) out.push(u);
  }
  return sortUnits(out);
}

// 生徒ページが最初に1回だけ呼ぶ（観点の並び＋全単元）
function getSyllabus(genre) {
  const g = (genre === 'setsumei') ? 'setsumei' : 'bungaku';
  return { genre: g, layers: LAYERS[g], observations: OBSERVATIONS[g], units: getUnits(g) };
}

/* ============================================================
 *  教員用：単元の追加・編集
 * ==========================================================*/
// payload: {id, genre, grade, name, material, order, active, must:[], note}
// id が空なら新規追加（IDは自動採番）。active=true にすると同じジャンルの他の単元は自動でfalse。
function saveUnit(payload) {
  requireTeacher();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) { return { ok: false, message: '混み合っています。もう一度試してください。' }; }
  try {
    const g = (payload && payload.genre === 'setsumei') ? 'setsumei' : 'bungaku';
    const name = String((payload && payload.name) || '').trim();
    if (!name) return { ok: false, message: '単元名を入力してください。' };

    const sh = getUnitSheet();
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const c = unitCols(head);
    const data = sh.getDataRange().getValues();

    // 観点は、このジャンルにあるものだけ受け取る（順番は観点マスタの並びにそろえる）
    const valid = OBSERVATIONS[g];
    const picked = (Array.isArray(payload.must) ? payload.must : []).map(function (s) { return String(s).trim(); });
    const must = valid.filter(function (o) { return picked.indexOf(o) >= 0; });

    let id = String((payload && payload.id) || '').trim();
    let target = -1;
    if (id) {
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][c.id] || '').trim() === id) { target = r + 1; break; }
      }
      if (target < 0) return { ok: false, message: 'その単元が見つかりませんでした。' };
    } else {
      id = nextUnitId(data, c, g);
    }

    const row = new Array(head.length).fill('');
    if (target > 0) for (let i = 0; i < head.length; i++) row[i] = (data[target - 1][i] === undefined) ? '' : data[target - 1][i];

    const set = function (col, val) { const i = head.indexOf(col); if (i >= 0) row[i] = val; };
    set('単元ID', id);
    set('学年', normGrade(payload.grade));
    set('ジャンル', GENRE_LABEL[g]);
    set('単元名', name);
    set('教材名', String(payload.material || '').trim());
    set('順序', Number(payload.order || 0) || nextOrder(data, c, g, normGrade(payload.grade), id));
    set('有効', payload.active === true);
    set('おさえる観点', must.join('、'));
    set('ひとこと', String(payload.note || '').trim().slice(0, MAX_NOTE));

    if (target > 0) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else { sh.appendRow(row); target = sh.getLastRow(); }

    // 「いま学習中」は同じジャンルで1つだけ
    if (payload.active === true && c.active >= 0) {
      const v2 = sh.getDataRange().getValues();
      for (let r = 1; r < v2.length; r++) {
        if (r + 1 === target) continue;
        const u = rowToUnit(v2[r], c);
        if (u && u.genre === g && u.active) sh.getRange(r + 1, c.active + 1).setValue(false);
      }
    }
    return { ok: true, id: id, units: getUnits(g) };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

function nextUnitId(data, c, g) {
  const prefix = (g === 'setsumei') ? 'S' : 'B';
  let max = 0;
  for (let r = 1; r < data.length; r++) {
    const id = String(data[r][c.id] || '').trim();
    const m = id.match(new RegExp('^' + prefix + '(\\d+)$'));
    if (m) max = Math.max(max, Number(m[1]));
  }
  const n = max + 1;
  return prefix + (n < 10 ? '0' + n : String(n));
}
function nextOrder(data, c, g, grade, selfId) {
  let max = 0;
  for (let r = 1; r < data.length; r++) {
    const u = rowToUnit(data[r], c);
    if (!u || u.genre !== g || u.id === selfId) continue;
    if (grade && u.grade !== grade) continue;
    max = Math.max(max, u.order || 0);
  }
  return max + 1;
}

// 単元を1つ消す（教員のみ）
function deleteUnit(unitId, genre) {
  requireTeacher();
  const id = String(unitId || '').trim();
  if (!id) return { ok: false, message: '単元IDがありません。' };
  const sh = getUnitSheet();
  const v = sh.getDataRange().getValues();
  const c = unitCols(v[0].map(String));
  for (let r = 1; r < v.length; r++) {
    if (String(v[r][c.id] || '').trim() === id) {
      sh.deleteRow(r + 1);
      return { ok: true, units: getUnits(genre) };
    }
  }
  return { ok: false, message: 'その単元が見つかりませんでした。' };
}

// 教員画面の初期読み込み（権限＋両ジャンルの単元＋観点の並び）
function getTeacherData() {
  const ctx = getTeacherContext();
  return {
    ctx: ctx,
    bungaku:  { layers: LAYERS.bungaku,  observations: OBSERVATIONS.bungaku,  units: getUnits('bungaku') },
    setsumei: { layers: LAYERS.setsumei, observations: OBSERVATIONS.setsumei, units: getUnits('setsumei') }
  };
}

/* ============================================================
 *  メニュー（任意）
 * ==========================================================*/
function onOpen() {
  SpreadsheetApp.getUi().createMenu('読みの観点')
    .addItem('単元マスタを開く（作成）', 'openUnitSheet')
    .addToUi();
}
function openUnitSheet() { getUnitSS().setActiveSheet(getUnitSheet()); }
