/**
 * ============================================================
 * アレルギー対応情報セクション（出張欄の下に新設）
 * ============================================================
 * 「アレルギー対応入力用ファイル」(ID: 1lP50kkohHS0XQCDq78etutBja0_CBdkl8DZqsCtEqWU)
 * の「アレルギー」シートから、その日の給食アレルギー対応情報を取り込んで
 * ダッシュボードの出張欄の下に表示します。
 *
 * 【マスター側「アレルギー」シートの列構成（1行目は見出し、データは2行目から）】
 *   A=月  B=日  C=№（使わない）  D=アレルギー対応内容（複数件は1セル内で改行区切り）
 * 年の情報が無いため、月日だけで当日のデータと照合します
 * （行事・出張と違って年度をまたいで使い続けるファイルではない前提）。
 *
 * 【導入手順】
 * 1. このファイルの中身をまるごとコピーして、Apps Scriptエディタで
 *    新しいファイル（例: AllergySync.gs）を作り、貼り付けて保存する。
 * 2. 「設定」シートに新しい行を追加する：
 *      設定項目: アレルギー対応マスターID
 *      値      : 1lP50kkohHS0XQCDq78etutBja0_CBdkl8DZqsCtEqWU
 * 3. Code.gs の getNippo(dateStr) 関数の中、
 *      try { result.deadlines = getDeadlineList_(ss); } catch (e) { ... }
 *    という行のすぐ後ろに、次の1行を追加する：
 *      try { result.allergy = getAllergyList_(ss, date); } catch (e) { console.log('allergy error: ' + e.message); }
 *    （resultオブジェクトの初期化部分にも、他の項目に倣って allergy: [] を
 *      追加しておくと、エラー時の見た目が安定します。必須ではありません）
 * 4. Apps Scriptエディタの関数選択で syncAllergyFromMaster を選び、
 *    「実行」を1回押す（初回はアクセス許可の承認が必要）。
 *    → 「アレルギー取込」シートが自動作成され、データが書き込まれれば成功。
 * 5. gas/allergy-section-patch.html の手順にしたがって、index.htmlに
 *    表示セクション・スタイル・反映ボタンを追加する。
 *
 * 【任意】1時間おきの自動同期も併用したい場合は、setupAllergySyncTrigger を
 * 1回だけ実行してください。ボタンだけで十分な場合は実行しなくて構いません。
 * ============================================================
 */

const ALLERGY_MASTER_SHEET_NAME = 'アレルギー';
const ALLERGY_IMPORT_SHEET_NAME = 'アレルギー取込';
const ALLERGY_SRC_COL = { MONTH: 1, DAY: 2, CONTENT: 4 };

/**
 * 「アレルギー取込」シートを取得（無ければ作成）
 */
function getOrCreateAllergyImportSheet_(ss) {
  let sh = ss.getSheetByName(ALLERGY_IMPORT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ALLERGY_IMPORT_SHEET_NAME);
    sh.getRange(1, 1).setValue(
      '※ このシートは「アレルギー対応入力用ファイル」から Apps Script (syncAllergyFromMaster) で自動取込されます。\n' +
      '手動で書き換えても、次回の同期で上書きされます。\n' +
      '【列構成】 A=月  B=日  C=内容'
    );
    sh.getRange(2, 1, 1, 3).setValues([['月', '日', '内容']]);
    sh.getRange(2, 1, 1, 3).setFontWeight('bold').setBackground('#fee2e2');
    sh.setFrozenRows(2);
  }
  return sh;
}

/**
 * 「アレルギー取込」シートを マスターファイル の内容で洗い替える
 * 手動実行 / 時間主導トリガーの両方から呼ばれる想定
 */
function syncAllergyFromMaster() {
  console.log('アレルギー同期: 開始');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings_(ss);
  const masterId = settings['アレルギー対応マスターID'];
  if (!masterId) {
    throw new Error('設定シートに「アレルギー対応マスターID」が設定されていません');
  }
  console.log('アレルギー同期: 設定取得完了。マスターID=' + masterId);

  let masterSs;
  try {
    masterSs = SpreadsheetApp.openById(String(masterId).trim());
  } catch (e) {
    throw new Error('アレルギー対応マスターを開けませんでした（IDまたは共有設定を確認してください）: ' + (e.message || e));
  }
  console.log('アレルギー同期: マスターを開けました');

  const sh = masterSs.getSheetByName(ALLERGY_MASTER_SHEET_NAME);
  if (!sh) {
    throw new Error('マスター側に「' + ALLERGY_MASTER_SHEET_NAME + '」シートが見つかりません');
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    throw new Error('アレルギー対応マスターにデータがありません');
  }
  console.log('アレルギー同期: マスター取得完了。データ行数=' + (lastRow - 1));

  // 1行目は見出しなので2行目から読み込む
  const data = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  const rows = [];
  data.forEach(row => {
    const month = row[ALLERGY_SRC_COL.MONTH - 1];
    const day = row[ALLERGY_SRC_COL.DAY - 1];
    const content = row[ALLERGY_SRC_COL.CONTENT - 1];
    if (!month || !day || !content) return;  // 空行はスキップ

    rows.push([
      allergyCleanValue_(month),
      allergyCleanValue_(day),
      allergyCleanValue_(content)
    ]);
  });
  console.log('アレルギー同期: データ整形完了。件数=' + rows.length);

  if (!rows.length) {
    throw new Error('アレルギー対応マスターから1件もデータを取得できませんでした。シート名・列構成を確認してください。');
  }

  const importSh = getOrCreateAllergyImportSheet_(ss);

  const clearLastRow = Math.max(importSh.getLastRow(), rows.length + 3);
  if (clearLastRow >= 3) {
    try {
      importSh.getRange(3, 1, clearLastRow - 2, 3).clearContent();
    } catch (e) {
      throw new Error('「アレルギー取込」シートの3行目以降をクリアできませんでした: ' + (e.message || e));
    }
  }
  console.log('アレルギー同期: 「アレルギー取込」クリア完了。書き込み開始');
  importSh.getRange(3, 1, rows.length, 3).setValues(rows);

  console.log('アレルギー同期完了: ' + rows.length + '件のデータを「アレルギー取込」に書き込みました');
  return { ok: true, count: rows.length };
}

/**
 * マスターのセル値を整形する（null/undefined/エラー値を空文字に）
 */
function allergyCleanValue_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string' && v.indexOf('#') === 0) return '';  // #N/A, #REF! など
  if (typeof v === 'number') return v;
  return String(v).trim();
}

/**
 * ダッシュボードの「🔄 アレルギー対応を反映」ボタン（管理者用）から呼ばれる。
 * 使い方: gas/allergy-section-patch.html を参照し、index.html にボタンを追加する。
 */
function manualSyncAllergyFromMaster() {
  const user = getCurrentUser();
  if (!user.ok) return { ok: false, error: user.reason };
  if (!user.isAdmin) return { ok: false, error: '管理者のみ実行できます' };
  try {
    const result = syncAllergyFromMaster();
    return { ok: true, count: result.count };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

/**
 * （任意）1時間おきにアレルギー対応同期を自動実行するトリガーを設定する
 */
function setupAllergySyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'syncAllergyFromMaster') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncAllergyFromMaster')
    .timeBased()
    .everyHours(1)
    .create();
  console.log('アレルギー対応の自動同期トリガーを設定しました（1時間おき）');
  return { ok: true, message: '1時間おきにアレルギー対応マスターから自動同期されます' };
}

/**
 * 指定日のアレルギー対応情報を取得する（getNippoから呼ばれる）
 * 年の情報が無いため、月・日だけで一致判定する。
 * 1つのセルに複数件（改行区切り）入っている場合は、行ごとに分割して返す。
 */
function getAllergyList_(ss, date) {
  const list = [];
  const sh = ss.getSheetByName(ALLERGY_IMPORT_SHEET_NAME);
  if (!sh || sh.getLastRow() < 3) return list;

  const targetMonth = date.getMonth() + 1;
  const targetDay = date.getDate();

  const data = sh.getRange(3, 1, sh.getLastRow() - 2, 3).getValues();
  data.forEach(row => {
    const month = parseInt(row[0], 10);
    const day = parseInt(row[1], 10);
    const content = row[2];
    if (!month || !day || !content) return;
    if (month !== targetMonth || day !== targetDay) return;

    String(content).split('\n').forEach(line => {
      const trimmed = line.trim();
      if (trimmed) list.push(trimmed);
    });
  });

  return list;
}

/**
 * アレルギー対応マスターのURL取得（出張マスターと同様、閲覧用リンクを開くため）
 */
function getAllergyMasterUrl() {
  const user = getCurrentUser();
  if (!user.ok) return { ok: false, error: user.reason };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings_(ss);
  const masterId = settings['アレルギー対応マスターID'];
  if (!masterId) return { ok: true, allergyMasterUrl: '' };
  return {
    ok: true,
    allergyMasterUrl: 'https://docs.google.com/spreadsheets/d/' + String(masterId).trim() + '/edit'
  };
}
