/**
 * ディベートフローシート統合Web App
 * フェーズ3.0
 *   - 単一画面UI: 教員も生徒とまったく同じ画面を見る。教員アカウントにだけ
 *     画面下部に操作バーが出る(教員ビュー/生徒ビューの切替は廃止)
 *   - フェーズ同期の高速化:
 *       * pollState() 1往復でフェーズ判定と全コンテキスト取得を同時に行う
 *       * CacheService による共有キャッシュでシート読み取りを大幅削減
 *       * 状態書き込みはバッチ化(1回のsetValuesにまとめる)
 *       * 生徒側の「画面を更新」ボタンを廃止し、自動で画面が切り替わる
 *
 * 名簿シート構成:
 *   A列: クラス / B列: 番号 / C列: 氏名 / D列: メール / E列: 役割
 *
 * 班マスタシート構成(名簿と同じ並び + F・G):
 *   A列: クラス / B列: 番号 / C列: 氏名 / D列: メールアドレス / E列: 役割
 *   F列: 班番号 / G列: ディベート役割("立論", "質疑", "第1反駁", "第2反駁" など)
 */

// ======================================================
// 定数
// ======================================================

const SHEET_NAMES = {
  MEIBO: '名簿',
  GROUPS: '班マスタ',
  TOPICS: '論題マスタ',
  DEBATERS: '発表者マスタ',
  VOTES: '投票',
  SUMMARIES: 'まとめ',
  STATE: '状態管理',
};

const MEIBO_COL = { KUMI: 1, BAN: 2, NAME: 3, EMAIL: 4, ROLE: 5 };

const STATE_KEYS = {
  CURRENT_TOPIC_ID: '現在の論題ID',
  PHASE: '現在のフェーズ',
  GD_REVEAL_STEP: 'GD公開段階',
  TODAYS_AFFIRMATIVE_GROUP: '本日の肯定班番号',
  TODAYS_NEGATIVE_GROUP: '本日の否定班番号',
  LOTTERY_RESULT: 'くじ結果',
  FORCE_SUBMIT_DEADLINE: '強制提出期限',
  USED_GROUPS: '本日の使用済み班',
};

// 全クラス共通の状態(旧形式からの移行用スコープ名)
const COMMON_SCOPE = '_共通_';

const PHASES = {
  IDLE: '待機中',
  LOTTERY_READY: 'くじ準備中',
  LOTTERY_DONE: '論題決定',
  VOTING: '投票受付中',
  VOTE_CLOSED_PRIVATE: '締切(集計中・結果非公開)',
  REVEAL_WINNER: '勝敗発表中',
  REVEAL_GD: 'グッドディベーター発表中',
  SUMMARY: 'まとめ受付中',
  SUMMARY_FORCE_COUNTDOWN: 'まとめ強制提出カウント中',
  CLOSED: '終了',
};

// ======================================================
// キャッシュ基盤
//   _memo   : 1回の実行内のメモ化(同じシートを何度も読まない)
//   Cache   : 全ユーザー共有の短期キャッシュ(同時アクセス時の負荷を下げる)
// 書き込み側で必ず drop / write-through するので、アプリ操作は即時反映される。
// (スプレッドシートを手で編集した場合だけ TTL 分だけ反映が遅れる)
// ======================================================

const CACHE_PREFIX = 'dbt30:';
const _memo = {};

function cacheRead_(key) {
  if (Object.prototype.hasOwnProperty.call(_memo, key)) return _memo[key];
  try {
    const raw = CacheService.getScriptCache().get(CACHE_PREFIX + key);
    if (raw) {
      const wrapped = JSON.parse(raw);
      _memo[key] = wrapped;
      return wrapped;
    }
  } catch (err) { /* キャッシュが使えなくても動作は継続する */ }
  return null;
}

function cacheWrite_(key, wrapped, ttlSec) {
  _memo[key] = wrapped;
  try {
    CacheService.getScriptCache().put(CACHE_PREFIX + key, JSON.stringify(wrapped), ttlSec || 30);
  } catch (err) { /* 値が大きすぎる等。無視して継続 */ }
}

function cacheDrop_(key) {
  delete _memo[key];
  try { CacheService.getScriptCache().remove(CACHE_PREFIX + key); } catch (err) { /* noop */ }
}

/** キャッシュ付き取得。producer は キャッシュミス時だけ実行される */
function cached_(key, ttlSec, producer) {
  const hit = cacheRead_(key);
  if (hit) return hit.v;
  const value = producer();
  cacheWrite_(key, { v: value }, ttlSec);
  return value;
}

// ======================================================
// Web App エントリーポイント
// ======================================================

function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setTitle('ディベート振り返り')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<p>エラーが発生しました: ' + err.message + '</p>');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ======================================================
// シート読み取り(すべてキャッシュ経由)
// ======================================================

/** 班マスタ全行(A〜G) */
function readGroups_() {
  return cached_('groups', 60, function() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.GROUPS);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, 7).getValues().map(function(row) {
      return {
        kumi: String(row[0]).trim(),
        ban: String(row[1]).trim(),
        name: String(row[2]).trim(),
        email: String(row[3]).trim().toLowerCase(),
        meiboRole: String(row[4]).trim(),
        groupNo: String(row[5]).trim(),
        role: String(row[6]).trim() || '発表者',
      };
    });
  });
}

/** 班マスタからユニークなクラス一覧 */
function getAllClasses_() {
  return cached_('classes', 60, function() {
    const seen = {};
    const list = [];
    for (const g of readGroups_()) {
      if (g.kumi && !seen[g.kumi]) { seen[g.kumi] = true; list.push(g.kumi); }
    }
    list.sort();
    return list;
  });
}

function getGroupMembers_(kumi, groupNo) {
  const k = String(kumi).trim();
  const g = String(groupNo).trim();
  return readGroups_().filter(function(row) {
    return row.kumi === k && row.groupNo === g;
  }).map(function(row) {
    return { email: row.email, name: row.name, role: row.role };
  });
}

function getTopics_() {
  return cached_('topics', 60, function() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TOPICS);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const topics = [];
    for (const row of data) {
      if (!row[0]) continue;
      topics.push({
        id: String(row[0]).trim(),
        title: String(row[1]).trim(),
        date: String(row[2]).trim(),
        affirmative: String(row[3]).trim(),
        negative: String(row[4]).trim(),
        targetClass: String(row[5]).trim(),
        memo: String(row[6]).trim(),
      });
    }
    return topics;
  });
}

/** 論題IDで引ける連想配列(getTopicById_ がシートを何度も読まないように) */
function getTopicMap_() {
  if (_memo.__topicMap) return _memo.__topicMap;
  const map = {};
  for (const t of getTopics_()) map[t.id] = t;
  _memo.__topicMap = map;
  return map;
}

function getTopicById_(topicId) {
  if (!topicId) return null;
  return getTopicMap_()[String(topicId).trim()] || null;
}

function getDebatersByTopic_(topicId) {
  if (!topicId) return [];
  return cached_('deb:' + topicId, 60, function() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DEBATERS);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const debaters = [];
    for (const row of data) {
      if (String(row[0]).trim() !== topicId) continue;
      debaters.push({
        email: String(row[1]).trim().toLowerCase(),
        name: String(row[2]).trim(),
        kumi: String(row[3]).trim(),
        ban: String(row[4]).trim(),
        team: String(row[5]).trim(),
        role: String(row[6]).trim(),
        order: Number(row[7]) || 999,
      });
    }
    debaters.sort(function(a, b) { return a.order - b.order; });
    return debaters;
  });
}

/** 名簿(メール → ユーザー情報)。認証は全リクエストで走るのでキャッシュ必須 */
function readMeiboIndex_() {
  return cached_('meibo', 120, function() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MEIBO);
    if (!sheet) return null;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const index = {};
    for (const row of data) {
      const email = String(row[MEIBO_COL.EMAIL - 1]).trim().toLowerCase();
      if (!email) continue;
      index[email] = {
        email: String(row[MEIBO_COL.EMAIL - 1]).trim(),
        name: String(row[MEIBO_COL.NAME - 1]).trim(),
        kumi: String(row[MEIBO_COL.KUMI - 1]).trim(),
        ban: String(row[MEIBO_COL.BAN - 1]).trim(),
        role: String(row[MEIBO_COL.ROLE - 1]).trim() || '生徒',
      };
    }
    return index;
  });
}

function getTeacherEmails_() {
  return cached_('teachers', 120, function() {
    const index = readMeiboIndex_() || {};
    const result = {};
    for (const email in index) {
      if (index[email].role === '教員') result[email] = true;
    }
    return result;
  });
}

/** 投票シート全行(1実行内メモのみ。投票中は頻繁に変わるため共有キャッシュはしない) */
function readVotes_() {
  if (_memo.__votes) return _memo.__votes;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.VOTES);
  let rows = [];
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) rows = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  }
  _memo.__votes = rows;
  return rows;
}

/** まとめシート全行(1実行内メモ。過去まとめ・フィードバック・自分の提出をまとめて1回の読み取りで賄う) */
function readSummaries_() {
  if (_memo.__summaries) return _memo.__summaries;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SUMMARIES);
  const out = { rows: [], numCols: 16 };
  if (sheet) {
    const lastRow = sheet.getLastRow();
    const numCols = Math.max(16, sheet.getLastColumn());
    out.numCols = numCols;
    if (lastRow >= 2) out.rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  }
  _memo.__summaries = out;
  return out;
}

function dropSummaryMemo_() { delete _memo.__summaries; }
function dropVoteMemo_() { delete _memo.__votes; }

// ======================================================
// 状態管理(クラス別 / キャッシュ+バッチ書き込み)
// ======================================================

/** 状態管理シート全体を { クラス: { キー: 値 } } の形で返す */
function readStateMap_() {
  return cached_('stateMap', 25, function() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STATE);
    const map = {};
    if (!sheet) return map;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (const row of data) {
      const scope = String(row[0]).trim();
      const key = String(row[1]).trim();
      if (!scope || !key) continue;
      if (!map[scope]) map[scope] = {};
      map[scope][key] = String(row[2]).trim();
    }
    return map;
  });
}

function defaultState_(kumi) {
  return {
    kumi: kumi,
    topicId: '', phase: PHASES.IDLE, gdStep: 0,
    affirmativeGroup: '', negativeGroup: '', lotteryResult: '',
    forceSubmitDeadline: 0, usedGroups: [],
  };
}

/**
 * 指定クラスの状態を取得する
 * @param {string} kumi クラス名(例: "3年1組")
 */
function getState_(kumi) {
  const targetKumi = String(kumi || '').trim();
  const state = defaultState_(targetKumi);
  if (!targetKumi) return state;
  const entries = readStateMap_()[targetKumi];
  if (!entries) return state;

  const val = function(key) { return entries[key] !== undefined ? entries[key] : ''; };
  state.topicId = val(STATE_KEYS.CURRENT_TOPIC_ID);
  state.phase = val(STATE_KEYS.PHASE) || PHASES.IDLE;
  state.gdStep = Number(val(STATE_KEYS.GD_REVEAL_STEP)) || 0;
  state.affirmativeGroup = val(STATE_KEYS.TODAYS_AFFIRMATIVE_GROUP);
  state.negativeGroup = val(STATE_KEYS.TODAYS_NEGATIVE_GROUP);
  state.lotteryResult = val(STATE_KEYS.LOTTERY_RESULT);
  state.forceSubmitDeadline = Number(val(STATE_KEYS.FORCE_SUBMIT_DEADLINE)) || 0;
  const used = val(STATE_KEYS.USED_GROUPS);
  state.usedGroups = used
    ? used.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; })
    : [];
  return state;
}

/**
 * 複数の状態キーをまとめて書き込む(シート読み書きは1往復ずつ)
 * @param {string} kumi
 * @param {Object} pairs { キー: 値, ... }
 */
function setStates_(kumi, pairs) {
  const targetKumi = String(kumi || '').trim();
  if (!targetKumi) throw new Error('クラスが指定されていません(setStates_)');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STATE);
  if (!sheet) throw new Error('状態管理シートがありません');

  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];

  const remaining = {};
  for (const k in pairs) remaining[k] = String(pairs[k] === null || pairs[k] === undefined ? '' : pairs[k]);

  let dirty = false;
  for (let i = 0; i < data.length; i++) {
    const scope = String(data[i][0]).trim();
    const key = String(data[i][1]).trim();
    if (scope !== targetKumi) continue;
    if (!Object.prototype.hasOwnProperty.call(remaining, key)) continue;
    if (String(data[i][2]) !== remaining[key]) { data[i][2] = remaining[key]; dirty = true; }
    delete remaining[key];
  }
  if (dirty && data.length > 0) sheet.getRange(2, 1, data.length, 3).setValues(data);

  const appends = [];
  for (const key in remaining) appends.push([targetKumi, key, remaining[key]]);
  if (appends.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 3).setValues(appends);
  }

  // キャッシュを書き戻す(次の読み取りでシートを読み直さなくて済む)
  const map = readStateMap_();
  if (!map[targetKumi]) map[targetKumi] = {};
  for (const key in pairs) {
    map[targetKumi][key] = String(pairs[key] === null || pairs[key] === undefined ? '' : pairs[key]);
  }
  cacheWrite_('stateMap', { v: map }, 25);
}

function setState_(kumi, key, value) {
  const pairs = {};
  pairs[key] = value;
  setStates_(kumi, pairs);
}

/**
 * 全クラスの状態を一覧で返す
 */
function getAllStates_() {
  return getAllClasses_().map(function(kumi) { return getState_(kumi); });
}

/** 教員パネル用の全クラス概要(シート読み取りは状態1回+論題1回のみ) */
function buildClassSummaries_() {
  return getAllClasses_().map(function(kumi) {
    const s = getState_(kumi);
    const topic = getTopicById_(s.topicId);
    return {
      kumi: kumi,
      phase: s.phase,
      topicId: s.topicId,
      topicTitle: topic ? topic.title : '',
      gdStep: s.gdStep,
    };
  });
}

/** クライアントが「変化したか」を1つの文字列で判定するためのリビジョン */
function buildRev_(state) {
  return [
    state.topicId, state.phase, state.gdStep,
    state.affirmativeGroup, state.negativeGroup, state.lotteryResult,
    state.forceSubmitDeadline,
  ].join('|');
}

// ======================================================
// マスタ初期化
// ======================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const meibo = ss.getSheetByName(SHEET_NAMES.MEIBO);
  if (!meibo) {
    ui.alert('エラー', '「' + SHEET_NAMES.MEIBO + '」シートが存在しません。', ui.ButtonSet.OK);
    return;
  }

  const logs = [];

  const e1 = meibo.getRange(1, MEIBO_COL.ROLE).getValue();
  if (!e1) {
    meibo.getRange(1, MEIBO_COL.ROLE).setValue('役割');
    meibo.getRange(1, MEIBO_COL.ROLE).setFontWeight('bold');
    logs.push('名簿E列に「役割」ヘッダを追加');
  } else if (String(e1).trim() !== '役割') {
    logs.push('警告: 名簿E1の値が「役割」ではありません(' + e1 + ')');
  } else {
    logs.push('名簿: E列確認OK');
  }

  logs.push(createSheetIfNotExists_(ss, SHEET_NAMES.GROUPS, [
    'クラス', '番号', '氏名', 'メールアドレス', '役割', '班番号', 'ディベート役割'
  ]));

  logs.push(createSheetIfNotExists_(ss, SHEET_NAMES.TOPICS, [
    '論題ID', '論題文', '実施日', '肯定班', '否定班', '対象クラス', '備考'
  ]));

  logs.push(createSheetIfNotExists_(ss, SHEET_NAMES.DEBATERS, [
    '論題ID', '生徒メール', '氏名', '組', '番', '班', '役割', '表示順'
  ]));

  logs.push(createSheetIfNotExists_(ss, SHEET_NAMES.VOTES, [
    'タイムスタンプ', '論題ID', '投票者メール', '投票者氏名', '組', '番',
    '勝敗選択', 'GD1', 'GD2', 'GD3', 'GD4', 'GD5'
  ]));

  logs.push(createSheetIfNotExists_(ss, SHEET_NAMES.SUMMARIES, [
    'タイムスタンプ', '論題ID', '生徒メール', '氏名', '組', '番',
    '立場', '役立った発表者メール', 'どの意見', 'どう広がり深まったか',
    '最終更新', '教員コメント', '確認済み', '評価',
    '違反ロック', '違反詳細'
  ]));

  ensureSummaryViolationColumns_(logs);

  logs.push(createStateSheetIfNotExists_(ss));
  migrateStateSheetIfNeeded_(logs);
  flushAllCaches_();
  initializeAllClassStates_(logs);
  flushAllCaches_();

  ui.alert('セットアップ完了', logs.join('\n'), ui.ButtonSet.OK);
}

/** 手作業でシートを直したあとにキャッシュを捨てる(メニューからも呼べる) */
function flushAllCaches_() {
  const keys = ['groups', 'classes', 'topics', 'meibo', 'teachers', 'stateMap'];
  for (const k of keys) cacheDrop_(k);
  delete _memo.__topicMap;
  dropSummaryMemo_();
  dropVoteMemo_();
}

function flushAllCaches() {
  flushAllCaches_();
  SpreadsheetApp.getUi().alert('キャッシュを消去しました。次の読み込みでシートの最新内容が反映されます。');
}

function createStateSheetIfNotExists_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.STATE);
  if (sheet) return '「' + SHEET_NAMES.STATE + '」: 既存(変更なし)';
  sheet = ss.insertSheet(SHEET_NAMES.STATE);
  sheet.getRange(1, 1, 1, 3).setValues([['クラス', 'キー', '値']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e0e0e0');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getMaxRows(), 3).setNumberFormat('@');
  return '「' + SHEET_NAMES.STATE + '」: 新規作成(3列構成)';
}

/** 旧形式(2列: キー/値)を新形式(3列: クラス/キー/値)に移行 */
function migrateStateSheetIfNeeded_(logs) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.STATE);
  if (!sheet) return;

  if (sheet.getLastColumn() >= 3) {
    const a1 = String(sheet.getRange(1, 1).getValue()).trim();
    if (a1 === 'クラス') return; // 既に新形式
  }

  const lastRow = sheet.getLastRow();
  let oldData = [];
  if (lastRow >= 2) oldData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([['クラス', 'キー', '値']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e0e0e0');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getMaxRows(), 3).setNumberFormat('@');

  if (oldData.length > 0) {
    const newData = oldData
      .filter(function(row) { return row[0]; })
      .map(function(row) { return [COMMON_SCOPE, row[0], row[1]]; });
    if (newData.length > 0) sheet.getRange(2, 1, newData.length, 3).setValues(newData);
  }
  if (logs) logs.push('状態管理シート: 2列→3列に移行(クラス別状態管理)');
}

/** 班マスタのクラスごとに必要な状態キーを初期化 */
function initializeAllClassStates_(logs) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.STATE);
  if (!sheet) return;

  const classes = getAllClasses_();
  if (classes.length === 0) {
    if (logs) logs.push('状態管理: 班マスタにクラスがないため、状態キーは初期化されませんでした');
    return;
  }

  const required = [
    { key: STATE_KEYS.CURRENT_TOPIC_ID, def: '' },
    { key: STATE_KEYS.PHASE, def: PHASES.IDLE },
    { key: STATE_KEYS.GD_REVEAL_STEP, def: '0' },
    { key: STATE_KEYS.TODAYS_AFFIRMATIVE_GROUP, def: '' },
    { key: STATE_KEYS.TODAYS_NEGATIVE_GROUP, def: '' },
    { key: STATE_KEYS.LOTTERY_RESULT, def: '' },
    { key: STATE_KEYS.FORCE_SUBMIT_DEADLINE, def: '' },
    { key: STATE_KEYS.USED_GROUPS, def: '' },
  ];

  const lastRow = sheet.getLastRow();
  const existing = {};
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (const row of data) {
      const scope = String(row[0]).trim();
      const key = String(row[1]).trim();
      if (!scope || !key) continue;
      if (!existing[scope]) existing[scope] = {};
      existing[scope][key] = true;
    }
  }

  const newRows = [];
  for (const kumi of classes) {
    for (const r of required) {
      if (!(existing[kumi] && existing[kumi][r.key])) newRows.push([kumi, r.key, r.def]);
    }
  }
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
  }
  if (logs) {
    logs.push('状態管理: ' + classes.length + 'クラス × ' + required.length + 'キー(' + newRows.length + '件追加)');
  }
}

/** まとめシートに違反ロック列・違反詳細列が無ければ追加 */
function ensureSummaryViolationColumns_(logs) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUMMARIES);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol >= 16) return;

  if (lastCol < 15) {
    sheet.getRange(1, 15).setValue('違反ロック');
    sheet.getRange(1, 15).setFontWeight('bold').setBackground('#e0e0e0');
    if (logs) logs.push('まとめシート: O列「違反ロック」を追加');
  }
  if (lastCol < 16) {
    sheet.getRange(1, 16).setValue('違反詳細');
    sheet.getRange(1, 16).setFontWeight('bold').setBackground('#e0e0e0');
    if (logs) logs.push('まとめシート: P列「違反詳細」を追加');
  }
}

function createSheetIfNotExists_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return '「' + name + '」: 既存(変更なし)';
  sheet = ss.insertSheet(name);
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold');
  range.setBackground('#e0e0e0');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setNumberFormat('@');
  return '「' + name + '」: 新規作成';
}

function fillDefaultRoles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const meibo = ss.getSheetByName(SHEET_NAMES.MEIBO);
  if (!meibo) { ui.alert('名簿が見つかりません'); return; }
  const lastRow = meibo.getLastRow();
  if (lastRow < 2) { ui.alert('名簿が空です'); return; }

  const roleRange = meibo.getRange(2, MEIBO_COL.ROLE, lastRow - 1, 1);
  const roles = roleRange.getValues();
  let filled = 0;
  for (let i = 0; i < roles.length; i++) {
    if (!roles[i][0] || String(roles[i][0]).trim() === '') {
      roles[i][0] = '生徒';
      filled++;
    }
  }
  if (filled > 0) roleRange.setValues(roles);
  flushAllCaches_();
  ui.alert(filled + '件の行に「生徒」を設定しました。');
}

// ======================================================
// メニュー
// ======================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ディベートアプリ')
    .addItem('初期セットアップ', 'setupSheets')
    .addItem('名簿の空欄役割に「生徒」を一括入力', 'fillDefaultRoles')
    .addSeparator()
    .addItem('ディベート進行はWeb画面で操作します', 'showWebOperationGuide')
    .addItem('全クラスの現在の状態', 'showAllStates')
    .addItem('◇ クラスのフェーズをリセット(待機中へ)', 'resetClassToIdle')
    .addItem('◇ まとめの違反ロックを解除', 'unlockStudentSummary')
    .addSeparator()
    .addItem('シートを手で直した後のキャッシュ消去', 'flushAllCaches')
    .addToUi();
}

function showWebOperationGuide() {
  SpreadsheetApp.getUi().alert(
    'Web画面で操作します',
    'ディベートの進行操作は、Web画面を教員アカウントで開いたときに画面下部に出る「操作バー」から行います。\n\n' +
    '教員も生徒とまったく同じ画面を見ます。教員アカウントの場合だけ、その画面の下に操作ボタンが表示されます。\n\n' +
    'クラスの切り替えも操作バーのプルダウンから行えます。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showAllStates() {
  const ui = SpreadsheetApp.getUi();
  const states = getAllStates_();
  if (states.length === 0) { ui.alert('班マスタにクラスが登録されていません。'); return; }
  const lines = states.map(function(s) {
    const topic = getTopicById_(s.topicId);
    return '【' + s.kumi + '】 ' + s.phase + ' / 論題: ' + (s.topicId || '-') + ' "' + (topic ? topic.title : '-') + '"';
  });
  ui.alert('全クラスの現在の状態', lines.join('\n\n'), ui.ButtonSet.OK);
}

function resetClassToIdle() {
  const ui = SpreadsheetApp.getUi();
  const classes = getAllClasses_();
  if (classes.length === 0) { ui.alert('班マスタにクラスがありません'); return; }
  const choices = classes.map(function(c, i) { return (i + 1) + '. ' + c; }).join('\n');
  const response = ui.prompt('リセット対象のクラスを選択', choices + '\n\n番号を入力:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const idx = Number(response.getResponseText()) - 1;
  if (isNaN(idx) || idx < 0 || idx >= classes.length) { ui.alert('正しい番号を入力'); return; }
  const kumi = classes[idx];
  const pairs = {};
  pairs[STATE_KEYS.PHASE] = PHASES.IDLE;
  pairs[STATE_KEYS.GD_REVEAL_STEP] = '0';
  setStates_(kumi, pairs);
  ui.alert(kumi + ' のフェーズを「待機中」に戻しました。投票・まとめデータは残っています。');
}

// ======================================================
// 班マスタから発表者自動登録
// ======================================================

/**
 * 論題マスタの指定論題の肯定班・否定班から発表者マスタを自動生成する
 */
function registerDebatersFromGroups() {
  const ui = SpreadsheetApp.getUi();

  const topics = getTopics_();
  if (topics.length === 0) { ui.alert('論題マスタに論題が登録されていません。'); return; }

  const choices = topics.map(function(t, i) {
    return (i + 1) + '. [' + t.id + '] ' + t.title +
      ' (対象クラス: ' + (t.targetClass || '-') + ', 肯定班: ' + (t.affirmative || '-') + ', 否定班: ' + (t.negative || '-') + ')';
  }).join('\n');

  const response = ui.prompt(
    '発表者を自動登録する論題を選択',
    choices + '\n\n番号(1〜' + topics.length + ')を入力:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const idx = Number(response.getResponseText()) - 1;
  if (isNaN(idx) || idx < 0 || idx >= topics.length) { ui.alert('正しい番号を入力してください'); return; }

  const topic = topics[idx];
  if (!topic.targetClass) {
    ui.alert('この論題には「対象クラス」が設定されていません。論題マスタを修正してください。');
    return;
  }
  if (!topic.affirmative || !topic.negative) {
    ui.alert('この論題には「肯定班」「否定班」のどちらかが未設定です。論題マスタを修正してください。');
    return;
  }

  const affirmativeMembers = getGroupMembers_(topic.targetClass, topic.affirmative);
  const negativeMembers = getGroupMembers_(topic.targetClass, topic.negative);

  if (affirmativeMembers.length === 0 && negativeMembers.length === 0) {
    ui.alert('班マスタに該当する生徒が見つかりません。\nクラス: ' + topic.targetClass +
      '\n肯定班: ' + topic.affirmative + '\n否定班: ' + topic.negative);
    return;
  }

  const preview =
    '◆肯定側(' + topic.affirmative + '班): ' + affirmativeMembers.length + '名\n' +
    affirmativeMembers.map(function(m) { return '  ' + m.name + ' (' + m.role + ')'; }).join('\n') +
    '\n\n◆否定側(' + topic.negative + '班): ' + negativeMembers.length + '名\n' +
    negativeMembers.map(function(m) { return '  ' + m.name + ' (' + m.role + ')'; }).join('\n') +
    '\n\n※既に登録済みの同論題の発表者情報は一旦削除して、上記の内容で入れ替えます。';

  if (ui.alert('発表者登録の確認', preview, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const count = writeDebaters_(topic.id, topic.targetClass, affirmativeMembers, negativeMembers);
  ui.alert('完了', count + '件の発表者を登録しました。', ui.ButtonSet.OK);
}

/**
 * 発表者マスタを書き換える(同論題の既存行は削除して入れ替え)
 * @return {number} 書き込んだ行数
 */
function writeDebaters_(topicId, kumi, affirmativeMembers, negativeMembers) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DEBATERS);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const keep = data.filter(function(row) { return String(row[0]).trim() !== topicId; });
    sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
    if (keep.length > 0) sheet.getRange(2, 1, keep.length, 8).setValues(keep);
  }

  const newRows = [];
  let order = 1;
  for (const m of affirmativeMembers) newRows.push([topicId, m.email, m.name, kumi, '', '肯定側', m.role, order++]);
  for (const m of negativeMembers) newRows.push([topicId, m.email, m.name, kumi, '', '否定側', m.role, order++]);

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
  }
  cacheDrop_('deb:' + topicId);
  return newRows.length;
}

// ======================================================
// 認証
// ======================================================

function authenticate() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return { ok: false, user: null, message: 'メールアドレスが取得できません。' };

    const index = readMeiboIndex_();
    if (index === null) return { ok: false, user: null, message: '名簿シートが見つかりません。' };

    const hit = index[email.toLowerCase()];
    if (!hit) return { ok: false, user: null, message: 'あなた(' + email + ')は名簿に登録されていません。' };

    return {
      ok: true,
      user: { email: email, name: hit.name, kumi: hit.kumi, ban: hit.ban, role: hit.role },
      message: 'OK',
    };
  } catch (err) {
    return { ok: false, user: null, message: 'エラー: ' + err.message };
  }
}

/** 教員操作の共通ガード。成功なら { ok:true, auth } を返す */
function requireTeacher_(kumi) {
  const auth = authenticate();
  if (!auth.ok) return { ok: false, message: auth.message };
  if (auth.user.role !== '教員') return { ok: false, message: '教員のみ操作できます' };
  if (kumi !== undefined) {
    if (!kumi) return { ok: false, message: 'クラスが指定されていません' };
    if (getAllClasses_().indexOf(String(kumi).trim()) < 0) {
      return { ok: false, message: 'クラス「' + kumi + '」は班マスタに存在しません' };
    }
  }
  return { ok: true, auth: auth };
}

/**
 * 教員の操作対象クラスを決める
 * 教員: 指定クラス > 自分のクラス(班マスタにある場合) > 先頭クラス
 * 生徒: 必ず自分のクラス
 */
function resolveTargetKumi_(auth, kumiOverride) {
  if (!auth.ok || !auth.user) return '';
  const allClasses = getAllClasses_();
  if (auth.user.role !== '教員') return auth.user.kumi || '';

  const requested = String(kumiOverride || '').trim();
  if (requested && allClasses.indexOf(requested) >= 0) return requested;
  if (auth.user.kumi && allClasses.indexOf(auth.user.kumi) >= 0) return auth.user.kumi;
  return allClasses.length > 0 ? allClasses[0] : '';
}

// ======================================================
// 教員操作: フェーズ制御
// ======================================================

/**
 * Web画面の操作バーから呼ばれる唯一の入口。
 * 成功時は最新コンテキストも同梱するので、操作 → 画面更新が1往復で済む。
 */
function teacherAction(kumi, action, arg) {
  try {
    const guard = requireTeacher_(kumi);
    if (!guard.ok) return guard;

    let res;
    switch (action) {
      case 'startLottery': res = opStartLottery_(kumi); break;
      case 'resetLotteryHistory': res = opResetLotteryHistory_(kumi); break;
      case 'openVoting': res = opOpenVoting_(kumi); break;
      case 'closeVoting': res = opCloseVoting_(kumi); break;
      case 'revealWinner': res = opRevealWinner_(kumi); break;
      case 'revealGd': res = opRevealGd_(kumi); break;
      case 'advanceGd': res = opAdvanceGd_(kumi); break;
      case 'openSummary': res = opOpenSummary_(kumi); break;
      case 'startForceSubmit': res = opStartForceSubmit_(kumi); break;
      case 'finalizeForceSubmit': res = opFinalizeForceSubmit_(kumi); break;
      case 'closeSummary': res = opCloseSummary_(kumi); break;
      case 'resetToIdle': res = opResetToIdle_(kumi); break;
      case 'setPhase': res = opSetPhase_(kumi, arg); break;
      default: return { ok: false, message: '不明な操作: ' + action };
    }

    if (res && res.ok) {
      // 操作後の画面を同じレスポンスで返す(往復を減らす)
      res.context = buildAppContext_(guard.auth, kumi, getState_(kumi));
    }
    return res;
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

function opStartLottery_(kumi) {
  const pairs = {};
  pairs[STATE_KEYS.PHASE] = PHASES.LOTTERY_READY;
  pairs[STATE_KEYS.LOTTERY_RESULT] = '';
  pairs[STATE_KEYS.TODAYS_AFFIRMATIVE_GROUP] = '';
  pairs[STATE_KEYS.TODAYS_NEGATIVE_GROUP] = '';
  pairs[STATE_KEYS.CURRENT_TOPIC_ID] = '';
  pairs[STATE_KEYS.GD_REVEAL_STEP] = '0';
  pairs[STATE_KEYS.FORCE_SUBMIT_DEADLINE] = '';
  setStates_(kumi, pairs);
  return { ok: true, message: kumi + ' をくじ準備中に設定しました' };
}

function opResetLotteryHistory_(kumi) {
  const pairs = {};
  pairs[STATE_KEYS.USED_GROUPS] = '';
  pairs[STATE_KEYS.PHASE] = PHASES.LOTTERY_READY;
  pairs[STATE_KEYS.LOTTERY_RESULT] = '';
  pairs[STATE_KEYS.TODAYS_AFFIRMATIVE_GROUP] = '';
  pairs[STATE_KEYS.TODAYS_NEGATIVE_GROUP] = '';
  pairs[STATE_KEYS.CURRENT_TOPIC_ID] = '';
  pairs[STATE_KEYS.GD_REVEAL_STEP] = '0';
  setStates_(kumi, pairs);
  return { ok: true, message: kumi + ' のくじ履歴をリセットしました' };
}

function opOpenVoting_(kumi) {
  const state = getState_(kumi);
  if (!state.topicId) return { ok: false, message: '先にくじを引いて論題を確定してください' };
  if (getDebatersByTopic_(state.topicId).length === 0) {
    return { ok: false, message: '発表者が登録されていません' };
  }
  const pairs = {};
  pairs[STATE_KEYS.PHASE] = PHASES.VOTING;
  pairs[STATE_KEYS.GD_REVEAL_STEP] = '0';
  setStates_(kumi, pairs);
  return { ok: true, message: kumi + ' の投票受付を開始しました' };
}

function opCloseVoting_(kumi) {
  const state = getState_(kumi);
  if (state.phase !== PHASES.VOTING) {
    return { ok: false, message: '現在のフェーズは「' + state.phase + '」です。投票受付中ではありません。' };
  }
  setState_(kumi, STATE_KEYS.PHASE, PHASES.VOTE_CLOSED_PRIVATE);
  cacheDrop_('tally:' + state.topicId);
  return { ok: true, message: '投票を締め切りました', tally: tallyVotes_(state.topicId) };
}

function opRevealWinner_(kumi) {
  const state = getState_(kumi);
  if (state.phase !== PHASES.VOTE_CLOSED_PRIVATE) {
    return { ok: false, message: '現在のフェーズは「' + state.phase + '」です。投票締切中ではありません。' };
  }
  setState_(kumi, STATE_KEYS.PHASE, PHASES.REVEAL_WINNER);
  return { ok: true, message: '勝敗発表に切り替えました' };
}

function opRevealGd_(kumi) {
  const state = getState_(kumi);
  if (state.phase !== PHASES.REVEAL_WINNER) {
    return { ok: false, message: '現在のフェーズは「' + state.phase + '」です。勝敗発表中ではありません。' };
  }
  const pairs = {};
  pairs[STATE_KEYS.PHASE] = PHASES.REVEAL_GD;
  pairs[STATE_KEYS.GD_REVEAL_STEP] = '0';
  setStates_(kumi, pairs);
  return { ok: true, message: 'GD発表モードに切り替えました' };
}

const GD_MAX_STEP = 2;

function opAdvanceGd_(kumi) {
  const state = getState_(kumi);
  if (state.phase !== PHASES.REVEAL_GD) {
    return { ok: false, message: '現在のフェーズは「' + state.phase + '」です。GD発表モードではありません。' };
  }
  const currentStep = Number(state.gdStep) || 0;
  if (currentStep >= GD_MAX_STEP) {
    return { ok: false, message: 'すべてのグッドディベーターを発表済みです。' };
  }
  const nextStep = currentStep + 1;
  setState_(kumi, STATE_KEYS.GD_REVEAL_STEP, String(nextStep));
  return { ok: true, step: nextStep, total: GD_MAX_STEP, message: '第' + nextStep + '段階を発表しました' };
}

function opOpenSummary_(kumi) {
  const state = getState_(kumi);
  if (!state.topicId) return { ok: false, message: '論題が設定されていません' };
  const pairs = {};
  pairs[STATE_KEYS.PHASE] = PHASES.SUMMARY;
  pairs[STATE_KEYS.FORCE_SUBMIT_DEADLINE] = '';
  setStates_(kumi, pairs);
  return { ok: true, message: kumi + ' のまとめ受付を開始しました' };
}

const FORCE_SUBMIT_COUNTDOWN_SEC = 20;

function opStartForceSubmit_(kumi) {
  const state = getState_(kumi);
  if (state.phase !== PHASES.SUMMARY) {
    return { ok: false, message: 'まとめ受付中ではありません(現在: ' + state.phase + ')' };
  }
  const deadline = Date.now() + FORCE_SUBMIT_COUNTDOWN_SEC * 1000;
  const pairs = {};
  pairs[STATE_KEYS.FORCE_SUBMIT_DEADLINE] = String(deadline);
  pairs[STATE_KEYS.PHASE] = PHASES.SUMMARY_FORCE_COUNTDOWN;
  setStates_(kumi, pairs);
  return {
    ok: true, message: 'カウントダウン開始',
    deadline: deadline, countdownSec: FORCE_SUBMIT_COUNTDOWN_SEC,
  };
}

function opFinalizeForceSubmit_(kumi) {
  const state = getState_(kumi);
  if (state.phase !== PHASES.SUMMARY_FORCE_COUNTDOWN && state.phase !== PHASES.SUMMARY) {
    return { ok: false, message: 'カウントダウン中ではありません(現在: ' + state.phase + ')' };
  }
  const lockedCount = lockAllSubmissionsForTopic_(state.topicId);
  const pairs = {};
  pairs[STATE_KEYS.PHASE] = PHASES.CLOSED;
  pairs[STATE_KEYS.FORCE_SUBMIT_DEADLINE] = '';
  setStates_(kumi, pairs);
  return { ok: true, message: 'まとめ受付を終了しました(全' + lockedCount + '件を確定)' };
}

function opCloseSummary_(kumi) {
  const state = getState_(kumi);
  if (state.phase !== PHASES.SUMMARY && state.phase !== PHASES.SUMMARY_FORCE_COUNTDOWN) {
    return { ok: false, message: '現在のフェーズは「' + state.phase + '」です。まとめ受付中ではありません。' };
  }
  const lockedCount = state.topicId ? lockAllSubmissionsForTopic_(state.topicId) : 0;
  const pairs = {};
  pairs[STATE_KEYS.PHASE] = PHASES.CLOSED;
  pairs[STATE_KEYS.FORCE_SUBMIT_DEADLINE] = '';
  setStates_(kumi, pairs);
  return { ok: true, message: kumi + ' のまとめ受付を締め切りました(' + lockedCount + '件確定)' };
}

function opResetToIdle_(kumi) {
  const pairs = {};
  pairs[STATE_KEYS.PHASE] = PHASES.IDLE;
  pairs[STATE_KEYS.GD_REVEAL_STEP] = '0';
  pairs[STATE_KEYS.FORCE_SUBMIT_DEADLINE] = '';
  setStates_(kumi, pairs);
  return { ok: true, message: kumi + ' を待機中に戻しました' };
}

/** 任意のフェーズへ直接ジャンプ */
function opSetPhase_(kumi, phaseKey) {
  if (!PHASES[phaseKey]) return { ok: false, message: '不明なフェーズキー: ' + phaseKey };
  const newPhase = PHASES[phaseKey];
  const state = getState_(kumi);

  const needsTopic = ['VOTING', 'VOTE_CLOSED_PRIVATE', 'REVEAL_WINNER', 'REVEAL_GD', 'SUMMARY'];
  if (needsTopic.indexOf(phaseKey) >= 0 && !state.topicId) {
    return { ok: false, message: '論題が未設定のため「' + newPhase + '」に進めません。先にくじを引いて論題を確定してください。' };
  }

  const pairs = {};
  pairs[STATE_KEYS.PHASE] = newPhase;
  if (['IDLE', 'LOTTERY_READY', 'VOTING', 'REVEAL_GD'].indexOf(phaseKey) >= 0) {
    pairs[STATE_KEYS.GD_REVEAL_STEP] = '0';
  }
  if (phaseKey !== 'SUMMARY' && phaseKey !== 'SUMMARY_FORCE_COUNTDOWN') {
    pairs[STATE_KEYS.FORCE_SUBMIT_DEADLINE] = '';
  }
  if (phaseKey === 'CLOSED' && state.topicId) lockAllSubmissionsForTopic_(state.topicId);

  setStates_(kumi, pairs);
  return { ok: true, message: kumi + ' を「' + newPhase + '」に変更しました' };
}

// ---- 旧API(互換のため残す。内部は teacherAction と同じ実装を呼ぶ) ----
function startLotteryForClass(kumi) { return teacherAction(kumi, 'startLottery'); }
function resetLotteryHistoryForClass(kumi) { return teacherAction(kumi, 'resetLotteryHistory'); }
function openVotingForClass(kumi) { return teacherAction(kumi, 'openVoting'); }
function closeVotingForClass(kumi) { return teacherAction(kumi, 'closeVoting'); }
function revealWinnerForClass(kumi) { return teacherAction(kumi, 'revealWinner'); }
function revealGoodDebaterForClass(kumi) { return teacherAction(kumi, 'revealGd'); }
function advanceGdReveal(kumi) { return teacherAction(kumi, 'advanceGd'); }
function openSummaryForClass(kumi) { return teacherAction(kumi, 'openSummary'); }
function closeSummaryForClass(kumi) { return teacherAction(kumi, 'closeSummary'); }
function resetToIdleForClass(kumi) { return teacherAction(kumi, 'resetToIdle'); }
function startForceSummarySubmit(kumi) { return teacherAction(kumi, 'startForceSubmit'); }
function finalizeForceSummarySubmit(kumi) { return teacherAction(kumi, 'finalizeForceSubmit'); }
function setClassPhase(kumi, phaseKey) { return teacherAction(kumi, 'setPhase', phaseKey); }

function getAllClassSummariesForTeacher() {
  const guard = requireTeacher_();
  if (!guard.ok) return guard;
  return { ok: true, summaries: buildClassSummaries_() };
}

// ======================================================
// くじ関連
// ======================================================

/**
 * 対象クラスの論題マスタから「班番号 → 論題」の逆引きテーブルを作る
 */
function buildGroupToTopicMap_(kumi) {
  const map = {};
  for (const t of getTopics_()) {
    if (String(t.targetClass).trim() !== String(kumi).trim()) continue;
    const a = String(t.affirmative).trim();
    const b = String(t.negative).trim();
    if (a) map[a] = { topicId: t.id, title: t.title, pairGroup: b };
    if (b) map[b] = { topicId: t.id, title: t.title, pairGroup: a };
  }
  return map;
}

/** くじで利用可能な班番号(論題マスタにあって、まだ使用されていない)を返す */
function getAvailableLotteryNumbers(kumi) {
  try {
    const guard = requireTeacher_(kumi);
    if (!guard.ok) return guard;

    const state = getState_(kumi);
    const map = buildGroupToTopicMap_(kumi);
    const used = state.usedGroups || [];
    const available = Object.keys(map)
      .map(function(s) { return Number(s); })
      .filter(function(n) { return !isNaN(n) && used.indexOf(String(n)) < 0; })
      .sort(function(a, b) { return a - b; });

    return {
      ok: true,
      available: available,
      used: used.map(function(s) { return Number(s); }).filter(function(n) { return !isNaN(n); }),
      allDone: available.length === 0,
      todaysClass: kumi,
    };
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

/**
 * くじ結果を受け取って、論題・肯定・否定を確定する
 * @param {number} winningNumber 班番号
 */
function confirmLottery(kumi, winningNumber) {
  try {
    const guard = requireTeacher_(kumi);
    if (!guard.ok) return guard;

    const num = Number(winningNumber);
    if (!(num >= 1 && num <= 99)) return { ok: false, message: 'くじの数字が不正です' };

    const state = getState_(kumi);
    if (state.usedGroups.indexOf(String(num)) >= 0) {
      return { ok: false, message: '班' + num + 'は既に使用済みです。別の班を選んでください。' };
    }

    const hit = buildGroupToTopicMap_(kumi)[String(num)];
    if (!hit) {
      return { ok: false, message: 'クラス「' + kumi + '」の班' + num + 'に対応する論題が論題マスタに登録されていません。' };
    }
    if (!hit.pairGroup) {
      return { ok: false, message: '論題「' + hit.title + '」の相手班が不明です。論題マスタを確認してください。' };
    }

    const newUsed = state.usedGroups.slice();
    if (newUsed.indexOf(String(num)) < 0) newUsed.push(String(num));
    if (newUsed.indexOf(String(hit.pairGroup)) < 0) newUsed.push(String(hit.pairGroup));

    const pairs = {};
    pairs[STATE_KEYS.CURRENT_TOPIC_ID] = hit.topicId;
    pairs[STATE_KEYS.TODAYS_AFFIRMATIVE_GROUP] = String(num);
    pairs[STATE_KEYS.TODAYS_NEGATIVE_GROUP] = hit.pairGroup;
    pairs[STATE_KEYS.LOTTERY_RESULT] = String(num);
    pairs[STATE_KEYS.PHASE] = PHASES.LOTTERY_DONE;
    pairs[STATE_KEYS.GD_REVEAL_STEP] = '0';
    pairs[STATE_KEYS.USED_GROUPS] = newUsed.join(',');
    setStates_(kumi, pairs);

    // 発表者マスタを自動登録(肯定班と否定班から)
    const affirmativeMembers = getGroupMembers_(kumi, String(num));
    const negativeMembers = getGroupMembers_(kumi, hit.pairGroup);
    let warning = '';
    if (affirmativeMembers.length === 0) warning += '肯定班(' + num + ')のメンバーが班マスタに見つかりません。 ';
    if (negativeMembers.length === 0) warning += '否定班(' + hit.pairGroup + ')のメンバーが班マスタに見つかりません。 ';
    const count = writeDebaters_(hit.topicId, kumi, affirmativeMembers, negativeMembers);

    const res = {
      ok: true,
      topicId: hit.topicId,
      topicTitle: hit.title,
      affirmativeGroup: String(num),
      negativeGroup: hit.pairGroup,
      debatersRegistered: count,
      warning: warning,
      message: '論題が決定しました',
    };
    res.context = buildAppContext_(guard.auth, kumi, getState_(kumi));
    return res;
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

// ======================================================
// 投票集計
// ======================================================

/**
 * 教員の投票は勝敗について3票換算。GDは教員票を集計対象外とする。
 */
function tallyVotes_(topicId) {
  const TEACHER_VOTE_WEIGHT = 3;
  const empty = {
    affirmativeCount: 0, negativeCount: 0, totalVoters: 0,
    studentAffirmative: 0, studentNegative: 0,
    teacherAffirmative: 0, teacherNegative: 0,
    teacherVoterCount: 0, teacherWeight: TEACHER_VOTE_WEIGHT,
    topFive: [], allRanked: [],
  };
  if (!topicId) return empty;

  return cached_('tally:' + topicId, 20, function() {
    const teacherEmails = getTeacherEmails_();
    const rows = readVotes_();
    let studentAff = 0, studentNeg = 0, teacherAff = 0, teacherNeg = 0, teacherVoters = 0, total = 0;
    const gdVotes = {};

    for (const row of rows) {
      if (String(row[1]).trim() !== topicId) continue;
      total++;

      const isTeacher = teacherEmails[String(row[2]).trim().toLowerCase()] === true;
      const winner = String(row[6]).trim();
      if (winner === '肯定側') { if (isTeacher) teacherAff++; else studentAff++; }
      else if (winner === '否定側') { if (isTeacher) teacherNeg++; else studentNeg++; }
      if (isTeacher) { teacherVoters++; continue; } // GDは生徒票のみ集計

      for (let i = 7; i <= 11; i++) {
        const email = String(row[i]).trim().toLowerCase();
        if (email) gdVotes[email] = (gdVotes[email] || 0) + 1;
      }
    }

    const debaterByEmail = {};
    for (const d of getDebatersByTopic_(topicId)) debaterByEmail[d.email] = d;

    const ranked = Object.keys(gdVotes).map(function(email) {
      const d = debaterByEmail[email];
      return { email: email, name: d ? d.name : email, team: d ? d.team : '', votes: gdVotes[email] };
    });
    ranked.sort(function(a, b) { return b.votes - a.votes; });

    return {
      affirmativeCount: studentAff + teacherAff * TEACHER_VOTE_WEIGHT,
      negativeCount: studentNeg + teacherNeg * TEACHER_VOTE_WEIGHT,
      totalVoters: total,
      studentAffirmative: studentAff,
      studentNegative: studentNeg,
      teacherAffirmative: teacherAff,
      teacherNegative: teacherNeg,
      teacherVoterCount: teacherVoters,
      teacherWeight: TEACHER_VOTE_WEIGHT,
      topFive: ranked.slice(0, 5),
      allRanked: ranked,
    };
  });
}

// ======================================================
// フロント向けAPI(コンテキスト取得 / ポーリング)
// ======================================================

/**
 * 画面表示に必要な情報を1つにまとめて返す
 * @param {string} kumiOverride 教員のみ有効。生徒は常に自分のクラス
 */
function getAppContext(kumiOverride) {
  const auth = authenticate();
  const targetKumi = resolveTargetKumi_(auth, kumiOverride);
  return buildAppContext_(auth, targetKumi, getState_(targetKumi));
}

function buildAppContext_(auth, targetKumi, state) {
  const allClasses = getAllClasses_();
  const now = Date.now();

  const ctx = {
    auth: auth,
    state: state,
    phases: PHASES,
    topic: null,
    debaters: [],
    hasVoted: false,
    results: null,
    gdStep: state.gdStep,
    mySummary: null,
    targetKumi: targetKumi,
    allClasses: allClasses,
    allClassSummaries: null,
    isStudentClassMismatch: false,
    myLatestPastSummary: null,
    myFeedbacksByTopic: null,
    rev: buildRev_(state),
    serverNow: now,
    // 端末の時計がずれていてもカウントダウンが狂わないよう、残りミリ秒で渡す
    forceRemainingMs: state.forceSubmitDeadline ? Math.max(0, state.forceSubmitDeadline - now) : 0,
  };

  if (!auth.ok || !auth.user) return ctx;

  const isTeacher = auth.user.role === '教員';
  if (isTeacher) ctx.allClassSummaries = buildClassSummaries_();

  if (!isTeacher && (!targetKumi || allClasses.indexOf(targetKumi) < 0)) {
    ctx.isStudentClassMismatch = true;
    return ctx;
  }

  // 待機系フェーズでは、過去のまとめと自分へのフィードバックを振り返れるようにする
  // (教員は生徒と同じ画面を見るので、条件は生徒と揃える)
  const isWaitingPhase = state.phase === PHASES.IDLE || state.phase === PHASES.LOTTERY_READY;
  if (isWaitingPhase) {
    ctx.myLatestPastSummary = getMyLatestPastSummary_(auth.user.email, state.topicId);
    ctx.myFeedbacksByTopic = getAllFeedbacksForMe_(auth.user.email);
  }

  if (!state.topicId) return ctx;

  ctx.topic = getTopicById_(state.topicId);
  ctx.debaters = getDebatersByTopic_(state.topicId);
  ctx.hasVoted = hasUserVoted_(state.topicId, auth.user.email);

  const summaryPhases = [PHASES.SUMMARY, PHASES.SUMMARY_FORCE_COUNTDOWN, PHASES.CLOSED];
  if (summaryPhases.indexOf(state.phase) >= 0) {
    const found = findSummaryRow_(state.topicId, auth.user.email);
    if (found.rowNum) {
      const row = found.data;
      ctx.mySummary = {
        position: String(row[6] || '').trim(),
        debaterEmail: String(row[7] || '').trim().toLowerCase(),
        whichOpinion: String(row[8] || ''),
        howExpanded: String(row[9] || ''),
        lastUpdated: String(row[10] || ''),
        teacherComment: String(row[11] || ''),
        confirmed: isConfirmed_(row),
        forceSubmitted: String(row[12] || '').trim() === '強制提出',
        evaluation: String(row[13] || ''),
        violationLocked: isViolationLocked_(row),
        violationDetail: String(row[15] || ''),
      };
    }
  }

  const resultPhases = [
    PHASES.REVEAL_WINNER, PHASES.REVEAL_GD, PHASES.SUMMARY,
    PHASES.SUMMARY_FORCE_COUNTDOWN, PHASES.CLOSED,
  ];
  if (resultPhases.indexOf(state.phase) >= 0) ctx.results = tallyVotes_(state.topicId);

  return ctx;
}

/**
 * 3秒ごとのポーリング用。クライアントが持っている rev と一致すれば軽い応答だけ、
 * 変化していればフルコンテキストも同じ応答に載せて返す。
 * → フェーズが変わったとき、追加の往復なしで即座に画面を切り替えられる。
 */
function pollState(kumiOverride, clientRev) {
  try {
    const auth = authenticate();
    if (!auth.ok) return { ok: false, message: auth.message };

    const targetKumi = resolveTargetKumi_(auth, kumiOverride);
    const state = getState_(targetKumi);
    const now = Date.now();
    const rev = buildRev_(state);

    const out = {
      ok: true,
      kumi: targetKumi,
      rev: rev,
      changed: String(clientRev || '') !== rev,
      topicId: state.topicId,
      phase: state.phase,
      gdStep: state.gdStep,
      lotteryResult: state.lotteryResult,
      serverNow: now,
      forceRemainingMs: state.forceSubmitDeadline ? Math.max(0, state.forceSubmitDeadline - now) : 0,
    };

    if (auth.user.role === '教員') out.allClassSummaries = buildClassSummaries_();
    if (out.changed) out.context = buildAppContext_(auth, targetKumi, state);
    return out;
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/** 旧API(互換) */
function getPhaseSnapshot(kumiOverride) {
  return pollState(kumiOverride, null);
}

// ======================================================
// 投票
// ======================================================

function hasUserVoted_(topicId, email) {
  if (!topicId) return false;
  const emailLower = String(email).trim().toLowerCase();
  for (const row of readVotes_()) {
    if (String(row[1]).trim() !== topicId) continue;
    if (String(row[2]).trim().toLowerCase() === emailLower) return true;
  }
  return false;
}

/**
 * 投票を受け付ける
 * payload: { winner, gdEmails[5], kumi(教員が担当クラスを指定する場合) }
 * 教員も投票できる(勝敗は3票換算、GDは集計対象外)
 */
function submitVote(payload) {
  try {
    const auth = authenticate();
    if (!auth.ok) return { ok: false, message: auth.message };

    const isTeacher = auth.user.role === '教員';
    const scopeKumi = isTeacher
      ? resolveTargetKumi_(auth, payload && payload.kumi)
      : String(auth.user.kumi || '').trim();
    if (!scopeKumi) return { ok: false, message: 'あなたのクラスが名簿に登録されていません' };

    const state = getState_(scopeKumi);
    if (state.phase !== PHASES.VOTING) {
      return { ok: false, message: '現在は投票受付中ではありません(' + state.phase + ')' };
    }
    if (!state.topicId) return { ok: false, message: '論題が設定されていません' };
    if (hasUserVoted_(state.topicId, auth.user.email)) {
      return { ok: false, message: 'すでに投票済みです。投票は1回のみです。' };
    }

    if (!payload || !payload.winner || !Array.isArray(payload.gdEmails)) {
      return { ok: false, message: '送信データが不正です' };
    }
    if (payload.winner !== '肯定側' && payload.winner !== '否定側') {
      return { ok: false, message: '勝敗選択が不正です' };
    }
    if (payload.gdEmails.length !== 5) {
      return { ok: false, message: 'GDは5名選択してください(現在 ' + payload.gdEmails.length + ' 名)' };
    }

    const uniq = {};
    for (const e of payload.gdEmails) {
      const k = String(e).trim().toLowerCase();
      if (!k) return { ok: false, message: 'GD選択に空が含まれています' };
      if (uniq[k]) return { ok: false, message: '同一人物を複数回選択できません' };
      uniq[k] = true;
    }
    const validEmails = {};
    for (const d of getDebatersByTopic_(state.topicId)) validEmails[d.email] = true;
    for (const e of payload.gdEmails) {
      if (!validEmails[String(e).trim().toLowerCase()]) {
        return { ok: false, message: '選択肢に発表者以外が含まれています' };
      }
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.VOTES);
    sheet.appendRow([
      new Date(), state.topicId,
      auth.user.email, auth.user.name, auth.user.kumi, auth.user.ban,
      payload.winner,
      payload.gdEmails[0], payload.gdEmails[1], payload.gdEmails[2],
      payload.gdEmails[3], payload.gdEmails[4],
    ]);
    dropVoteMemo_();
    cacheDrop_('tally:' + state.topicId);

    return { ok: true, message: '投票を受け付けました' };
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

// ======================================================
// まとめ (Summary) CRUD
// ======================================================

/** まとめシートから指定ユーザー・論題の行番号(1-indexed)と既存データを返す */
function findSummaryRow_(topicId, email) {
  const store = readSummaries_();
  const emailLower = String(email).trim().toLowerCase();
  for (let i = 0; i < store.rows.length; i++) {
    const row = store.rows[i];
    if (String(row[1]).trim() !== topicId) continue;
    if (String(row[2]).trim().toLowerCase() !== emailLower) continue;
    return { rowNum: i + 2, data: row };
  }
  return { rowNum: null, data: null };
}

/** 行データから違反ロック状態を取得 (O列 / index 14) */
function isViolationLocked_(rowData) {
  if (!rowData) return false;
  const v = String(rowData[14] === undefined || rowData[14] === null ? '' : rowData[14]).trim().toUpperCase();
  return v === 'TRUE' || v === '✓' || v === 'LOCKED';
}

/** M列(確認済み)から提出が確定しているかを判定 */
function isConfirmed_(rowData) {
  if (!rowData) return false;
  const raw = String(rowData[12] === undefined || rowData[12] === null ? '' : rowData[12]).trim();
  const v = raw.toUpperCase();
  return v === 'TRUE' || raw === '✓' || raw === '強制提出';
}

/** 自分のまとめを取得(生徒画面の表示用) */
function getMySummary() {
  try {
    const auth = authenticate();
    if (!auth.ok) return { ok: false, message: auth.message };
    const myKumi = String(auth.user.kumi || '').trim();
    if (!myKumi) return { ok: false, message: 'あなたのクラスが名簿に登録されていません' };
    const state = getState_(myKumi);
    if (!state.topicId) return { ok: false, message: '論題が設定されていません' };

    const found = findSummaryRow_(state.topicId, auth.user.email);
    if (!found.rowNum) return { ok: true, exists: false, summary: null };

    const row = found.data;
    return {
      ok: true,
      exists: true,
      summary: {
        timestamp: String(row[0] || ''),
        position: String(row[6] || '').trim(),
        debaterEmail: String(row[7] || '').trim().toLowerCase(),
        whichOpinion: String(row[8] || ''),
        howExpanded: String(row[9] || ''),
        lastUpdated: String(row[10] || ''),
        teacherComment: String(row[11] || ''),
        confirmed: isConfirmed_(row),
        forceSubmitted: String(row[12] || '').trim() === '強制提出',
        evaluation: String(row[13] || ''),
        violationLocked: isViolationLocked_(row),
        violationDetail: String(row[15] || ''),
      },
    };
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

/** 自分の最新の過去まとめ(現在の論題以外) */
function getMyLatestPastSummary_(email, currentTopicId) {
  const store = readSummaries_();
  const emailLower = String(email).trim().toLowerCase();
  const candidates = [];

  for (const row of store.rows) {
    if (String(row[2]).trim().toLowerCase() !== emailLower) continue;
    if (currentTopicId && String(row[1]).trim() === currentTopicId) continue;
    const hasContent = String(row[6] || '').trim() || String(row[8] || '').trim() || String(row[9] || '').trim();
    if (!hasContent) continue;
    candidates.push(row);
  }
  if (candidates.length === 0) return null;

  candidates.sort(function(a, b) {
    return String(b[10] || b[0] || '').localeCompare(String(a[10] || a[0] || ''));
  });
  const row = candidates[0];
  const topicId = String(row[1]).trim();
  const topic = getTopicById_(topicId);

  return {
    topicId: topicId,
    topicTitle: topic ? topic.title : '(論題不明)',
    position: String(row[6] || '').trim(),
    whichOpinion: String(row[8] || ''),
    howExpanded: String(row[9] || ''),
    lastUpdated: String(row[10] || ''),
    teacherComment: String(row[11] || ''),
    evaluation: String(row[13] || ''),
  };
}

/**
 * 自分が「役立った発表者」として選ばれた全フィードバックを論題ごとに取得
 * 匿名化: 書いた生徒の氏名・メール・組番は返さない
 */
function getAllFeedbacksForMe_(email) {
  const store = readSummaries_();
  const emailLower = String(email).trim().toLowerCase();
  const byTopic = {};

  for (const row of store.rows) {
    if (String(row[7] || '').trim().toLowerCase() !== emailLower) continue;
    if (isViolationLocked_(row)) continue;
    const whichOpinion = String(row[8] || '').trim();
    const howExpanded = String(row[9] || '').trim();
    if (!whichOpinion && !howExpanded) continue;

    const topicId = String(row[1]).trim();
    if (!byTopic[topicId]) byTopic[topicId] = [];
    byTopic[topicId].push({
      position: String(row[6] || '').trim(),
      whichOpinion: whichOpinion,
      howExpanded: howExpanded,
      timestamp: String(row[0] || ''),
    });
  }

  const result = [];
  for (const topicId in byTopic) {
    const topic = getTopicById_(topicId);
    result.push({
      topicId: topicId,
      topicTitle: topic ? topic.title : '(論題不明)',
      feedbacks: byTopic[topicId],
    });
  }
  result.sort(function(a, b) { return b.topicId.localeCompare(a.topicId); });
  return result;
}

/** 自分が選ばれたフィードバック(指定論題) */
function getFeedbackForMe(topicId) {
  try {
    const auth = authenticate();
    if (!auth.ok) return { ok: false, message: auth.message };

    let targetTopicId = topicId;
    if (!targetTopicId) targetTopicId = getState_(String(auth.user.kumi || '').trim()).topicId;
    if (!targetTopicId) return { ok: false, message: '論題が指定されていません' };

    const all = getAllFeedbacksForMe_(auth.user.email);
    for (const t of all) {
      if (t.topicId === targetTopicId) return { ok: true, feedbacks: t.feedbacks, count: t.feedbacks.length };
    }
    return { ok: true, feedbacks: [], count: 0 };
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

/**
 * まとめ提出/更新
 * payload: { position, debaterEmail, whichOpinion, howExpanded, violations:[{type,at}], forceSubmit }
 */
function submitSummary(payload) {
  try {
    const auth = authenticate();
    if (!auth.ok) return { ok: false, message: auth.message };

    const myKumi = String(auth.user.kumi || '').trim();
    if (!myKumi) return { ok: false, message: 'あなたのクラスが名簿に登録されていません' };

    const state = getState_(myKumi);
    if (state.phase !== PHASES.SUMMARY && state.phase !== PHASES.SUMMARY_FORCE_COUNTDOWN) {
      return { ok: false, message: '現在はまとめ受付中ではありません(' + state.phase + ')' };
    }
    if (!state.topicId) return { ok: false, message: '論題が設定されていません' };
    if (!payload) return { ok: false, message: '送信データが不正です' };

    const isForceMode = payload.forceSubmit === true;
    const validPositions = ['肯定', 'どちらかといえば肯定', 'どちらかといえば否定', '否定'];
    const whichOpinion = String(payload.whichOpinion || '').trim();
    const howExpanded = String(payload.howExpanded || '').trim();
    const debaterEmailLower = String(payload.debaterEmail || '').trim().toLowerCase();

    if (!isForceMode) {
      if (validPositions.indexOf(payload.position) < 0) return { ok: false, message: '立場の選択が不正です' };
      if (!payload.debaterEmail) return { ok: false, message: '役に立った発表者を選択してください' };
      if (whichOpinion.length === 0) return { ok: false, message: '「どの意見か」を記入してください' };
      if (howExpanded.length === 0) return { ok: false, message: '「どう広がり深まったか」を記入してください' };

      let debaterValid = false;
      for (const d of getDebatersByTopic_(state.topicId)) {
        if (d.email === debaterEmailLower) { debaterValid = true; break; }
      }
      if (!debaterValid) return { ok: false, message: '発表者が無効です' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SUMMARIES);
    const existing = findSummaryRow_(state.topicId, auth.user.email);

    if (!isForceMode && existing.rowNum && isViolationLocked_(existing.data)) {
      return {
        ok: false,
        message: '違反(タブ切替・ウィンドウ切替)が検出されたため、この提出はロックされています。「やり直す」ボタンで再開できます(入力内容は復元されます)。',
        violationLocked: true,
      };
    }

    const clientViolations = Array.isArray(payload.violations) ? payload.violations : [];
    const hasViolation = !isForceMode && clientViolations.length > 0;
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

    if (hasViolation) {
      // 違反ありの場合: 入力内容は保存せず、違反記録だけを残して拒否する
      const violationText = clientViolations.map(function(v) {
        return '[' + (v.at || '?') + '] ' + (v.type || '?');
      }).join(' / ');

      if (existing.rowNum) {
        sheet.getRange(existing.rowNum, 15).setValue('LOCKED');
        const currentDetail = String(existing.data[15] || '');
        sheet.getRange(existing.rowNum, 16)
          .setValue(currentDetail ? (currentDetail + ' ; ' + violationText) : violationText);
      } else {
        sheet.appendRow([
          nowStr, state.topicId, auth.user.email, auth.user.name, auth.user.kumi, auth.user.ban,
          '', '', '', '',
          nowStr, '', '', '',
          'LOCKED', violationText,
        ]);
      }
      dropSummaryMemo_();
      return {
        ok: false,
        message: '違反(タブ切替・ウィンドウ切替)が検出されたため、送信できません。「やり直す」ボタンで再開できます(入力内容は復元されます)。',
        violationLocked: true,
        violationDetail: violationText,
      };
    }

    const forceSubmitNote = isForceMode ? ('[強制提出:' + nowStr + ']') : '';

    if (existing.rowNum) {
      if (isConfirmed_(existing.data)) {
        return { ok: false, message: '既に提出が確定しているため、修正できません。' };
      }

      // 強制提出時の保護: 新しい値が空欄なら既存の値を保持する
      const keepIfEmpty = function(newVal, oldVal) {
        return (isForceMode && !String(newVal).trim()) ? oldVal : newVal;
      };
      const values = [[
        keepIfEmpty(String(payload.position || '').trim(), String(existing.data[6] || '').trim()),
        keepIfEmpty(debaterEmailLower, String(existing.data[7] || '').trim().toLowerCase()),
        keepIfEmpty(whichOpinion, String(existing.data[8] || '')),
        keepIfEmpty(howExpanded, String(existing.data[9] || '')),
        nowStr,
      ]];
      sheet.getRange(existing.rowNum, 7, 1, 5).setValues(values);

      if (isForceMode) {
        sheet.getRange(existing.rowNum, 13).setValue('強制提出');
        const prev = String(existing.data[15] || '');
        sheet.getRange(existing.rowNum, 16).setValue(prev ? (prev + ' ' + forceSubmitNote) : forceSubmitNote);
      }
      dropSummaryMemo_();
      return {
        ok: true,
        message: isForceMode ? '強制提出で保存しました' : '提出内容を更新しました',
        updated: true, forced: isForceMode,
      };
    }

    sheet.appendRow([
      nowStr, state.topicId, auth.user.email, auth.user.name, auth.user.kumi, auth.user.ban,
      payload.position || '', debaterEmailLower, whichOpinion, howExpanded,
      nowStr, '',
      isForceMode ? '強制提出' : '',
      '',
      '', forceSubmitNote,
    ]);
    dropSummaryMemo_();
    return {
      ok: true,
      message: isForceMode ? '強制提出で保存しました' : 'まとめを提出しました',
      updated: false, forced: isForceMode,
    };
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

/**
 * 指定論題の全提出行を確定させる(M列が未確定の行のみ「強制提出」を書き込む)
 * @return {number} 新たにロックした行数
 */
function lockAllSubmissionsForTopic_(topicId) {
  if (!topicId) return 0;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SUMMARIES);
  if (!sheet) return 0;
  const store = readSummaries_();
  if (store.rows.length === 0) return 0;

  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  const closeNote = '[全体締切:' + nowStr + ']';

  // M列(13)とP列(16)を1回ずつの setValues でまとめて書く
  const confirmedCol = [];
  const detailCol = [];
  let lockedCount = 0;

  for (const row of store.rows) {
    const isTarget = String(row[1]).trim() === topicId && !isConfirmed_(row);
    if (isTarget) {
      confirmedCol.push(['強制提出']);
      const prev = String(row[15] || '');
      detailCol.push([prev ? (prev + ' ' + closeNote) : closeNote]);
      lockedCount++;
    } else {
      // 対象外の行は元の値をそのまま書き戻す(チェックボックス等の型を壊さない)
      confirmedCol.push([row[12] === undefined || row[12] === null ? '' : row[12]]);
      detailCol.push([row[15] === undefined || row[15] === null ? '' : row[15]]);
    }
  }

  if (lockedCount > 0) {
    sheet.getRange(2, 13, confirmedCol.length, 1).setValues(confirmedCol);
    sheet.getRange(2, 16, detailCol.length, 1).setValues(detailCol);
    dropSummaryMemo_();
  }
  return lockedCount;
}

/** 教員による違反ロックの解除 */
function clearViolationLock(studentEmail, kumi, topicId) {
  try {
    const guard = requireTeacher_();
    if (!guard.ok) return guard;

    let targetTopicId = topicId;
    if (!targetTopicId && kumi) targetTopicId = getState_(kumi).topicId;
    if (!targetTopicId) return { ok: false, message: '論題が指定されていません' };
    if (!studentEmail) return { ok: false, message: '生徒のメールアドレスが指定されていません' };

    const found = findSummaryRow_(targetTopicId, studentEmail);
    if (!found.rowNum) return { ok: false, message: '該当の提出が見つかりません' };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SUMMARIES);
    sheet.getRange(found.rowNum, 15).setValue('');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    const currentDetail = String(found.data[15] || '');
    const releaseNote = '[解除:' + nowStr + ']';
    sheet.getRange(found.rowNum, 16).setValue(currentDetail ? (currentDetail + ' ' + releaseNote) : releaseNote);
    dropSummaryMemo_();

    return { ok: true, message: '違反ロックを解除しました' };
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

/** 生徒自身が自分のロックを解除する(やり直しボタン) */
function resetMyViolationLock() {
  try {
    const auth = authenticate();
    if (!auth.ok) return { ok: false, message: auth.message };

    const myKumi = String(auth.user.kumi || '').trim();
    if (!myKumi) return { ok: false, message: 'あなたのクラスが名簿に登録されていません' };
    const state = getState_(myKumi);
    if (!state.topicId) return { ok: false, message: '論題が設定されていません' };

    const found = findSummaryRow_(state.topicId, auth.user.email);
    if (!found.rowNum) return { ok: true, message: 'ロックはありません' };
    if (!isViolationLocked_(found.data)) return { ok: true, message: 'ロックされていません' };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SUMMARIES);
    sheet.getRange(found.rowNum, 15).setValue('');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    const currentDetail = String(found.data[15] || '');
    const selfResetNote = '[本人リセット:' + nowStr + ']';
    sheet.getRange(found.rowNum, 16).setValue(currentDetail ? (currentDetail + ' ' + selfResetNote) : selfResetNote);
    dropSummaryMemo_();

    return { ok: true, message: 'ロックを解除しました' };
  } catch (err) {
    return { ok: false, message: 'エラー: ' + err.message };
  }
}

/**
 * 違反ロックの解除(教員メニュー用)
 */
function unlockStudentSummary() {
  const ui = SpreadsheetApp.getUi();
  const classes = getAllClasses_();
  if (classes.length === 0) { ui.alert('班マスタにクラスがありません'); return; }
  const choices = classes.map(function(c, i) { return (i + 1) + '. ' + c; }).join('\n');
  const r = ui.prompt('クラスを選択', choices + '\n\n番号を入力:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const idx = Number(r.getResponseText()) - 1;
  if (isNaN(idx) || idx < 0 || idx >= classes.length) { ui.alert('正しい番号を入力'); return; }

  const kumi = classes[idx];
  const state = getState_(kumi);
  if (!state.topicId) { ui.alert(kumi + ' に現在の論題が設定されていません。'); return; }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SUMMARIES);
  if (!sheet) { ui.alert('まとめシートが見つかりません'); return; }

  const store = readSummaries_();
  const locked = [];
  for (let i = 0; i < store.rows.length; i++) {
    const row = store.rows[i];
    if (String(row[1]).trim() !== state.topicId) continue;
    if (!isViolationLocked_(row)) continue;
    locked.push({
      rowNum: i + 2,
      name: String(row[3] || ''),
      kumi: String(row[4] || ''),
      ban: String(row[5] || ''),
      email: String(row[2] || ''),
      detail: String(row[15] || ''),
    });
  }

  if (locked.length === 0) { ui.alert('現在の論題でロックされている提出はありません。'); return; }

  const lockedChoices = locked.map(function(l, i) {
    return (i + 1) + '. ' + l.name + ' (' + l.kumi + ' ' + l.ban + '番): ' + l.detail;
  }).join('\n');

  const response = ui.prompt(
    '違反ロックの解除',
    lockedChoices + '\n\n解除する生徒の番号を入力(1〜' + locked.length + ')、もしくは「all」で全員解除:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = String(response.getResponseText()).trim().toLowerCase();
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  const releaseNote = '[解除:' + nowStr + ']';

  const release = function(target) {
    sheet.getRange(target.rowNum, 15).setValue('');
    sheet.getRange(target.rowNum, 16).setValue(target.detail ? (target.detail + ' ' + releaseNote) : releaseNote);
  };

  if (input === 'all') {
    for (const l of locked) release(l);
    dropSummaryMemo_();
    ui.alert(locked.length + '件のロックを解除しました。');
    return;
  }

  const targetIdx = Number(input) - 1;
  if (isNaN(targetIdx) || targetIdx < 0 || targetIdx >= locked.length) {
    ui.alert('正しい番号または「all」を入力してください');
    return;
  }
  release(locked[targetIdx]);
  dropSummaryMemo_();
  ui.alert(locked[targetIdx].name + ' さんのロックを解除しました。');
}
