/**
 * ============================================================
 * 出張データを「出張マスター」から直接取得する（管理者用ボタン）
 * ============================================================
 * これまで「出張取込」シートは、
 *   =QUERY(IMPORTRANGE("出張マスターID", "出張!A2:I"), "select Col1,Col2,Col3,Col6,Col7,Col8,Col9", 0)
 * という数式で出張マスター（出張流し込みファイル）から取り込んでいました。
 * IMPORTRANGE数式は、出張マスター側の変更が自動で（ただし多少のタイムラグを
 * 伴って）反映される仕組みでしたが、「今すぐ確実に反映したい」という要望に
 * 応えるため、行事予定と同じ「管理者用の反映ボタン」を追加します。
 *
 * 【出張マスター側「出張」シートの列構成（1行目は見出し、データは2行目から）】
 *   A=年  B=月  C=日  D=人数  E=番号  F=氏名  G=用務  H=時間  I=場所
 * これを「出張取込」シートの7列フォーマット
 *   A=年度  B=月  C=日  D=氏名  E=用務  F=時間  G=場所
 * にマッピングします（D=人数, E=番号 は使わずスキップ）。
 * ※ A列「年」が令和年数（例:8）で入っている場合は西暦（2026）に自動変換します
 *   （shuchoNormalizeYear_）。西暦がそのまま入っている場合は変換しません。
 *
 * 【導入手順】
 * 1. このファイルの中身をまるごとコピーして、Apps Scriptエディタで
 *    新しいファイル（例: ShuchoSync.gs）を作り、貼り付けて保存する。
 * 2. 「出張取込」シートのA3セルに残っている、古いQUERY(IMPORTRANGE(...))の
 *    数式を削除する（Delete）。空にしておけばOK
 *    （このスクリプトが動くたびに、A3以降を書き換えるようになります）。
 * 3. Apps Scriptエディタの関数選択で syncShuchoFromMaster を選び、
 *    「実行」を1回押す（初回はアクセス許可の承認が必要）。
 *    → 「出張取込」シートにデータが書き込まれれば成功。
 * 4. gas/shucho-sync-button.html の手順にしたがって、ダッシュボードに
 *    「🔄 出張を反映」ボタンを追加する。
 *
 * 【任意】1時間おきの自動同期も併用したい場合は、setupShuchoSyncTrigger を
 * 1回だけ実行してください（行事予定と同じ仕組みです）。ボタンだけで
 * 十分な場合は実行しなくて構いません。
 * ============================================================
 */

// 出張マスター側「出張」シートの列番号（1-indexed）。1行目は見出し、データは2行目から
const SHUCHO_SOURCE_SHEET_NAME = '出張';
const SHUCHO_SRC_COL = {
  YEAR: 1, MONTH: 2, DAY: 3, NAME: 6, BUSYO: 7, TIME: 8, PLACE: 9
};

/**
 * 「出張取込」シートを 出張マスター の内容で洗い替える
 * 手動実行 / 時間主導トリガーの両方から呼ばれる想定
 */
function syncShuchoFromMaster() {
  console.log('出張同期: 開始');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings_(ss);
  const masterId = settings['出張マスターID'];
  if (!masterId) {
    throw new Error('設定シートに「出張マスターID」が設定されていません');
  }
  console.log('出張同期: 設定取得完了。マスターID=' + masterId);

  let masterSs;
  try {
    masterSs = SpreadsheetApp.openById(String(masterId).trim());
  } catch (e) {
    throw new Error('出張マスターを開けませんでした（IDまたは共有設定を確認してください）: ' + (e.message || e));
  }
  console.log('出張同期: マスターを開けました');

  const sh = masterSs.getSheetByName(SHUCHO_SOURCE_SHEET_NAME);
  if (!sh) {
    throw new Error('出張マスター側に「' + SHUCHO_SOURCE_SHEET_NAME + '」シートが見つかりません');
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    throw new Error('出張マスターにデータがありません');
  }
  console.log('出張同期: マスター取得完了。データ行数=' + (lastRow - 1));

  // 1行目は見出しなので2行目から読み込む
  const data = sh.getRange(2, 1, lastRow - 1, 9).getValues();
  const rows = [];
  data.forEach(row => {
    const year = row[SHUCHO_SRC_COL.YEAR - 1];
    const month = row[SHUCHO_SRC_COL.MONTH - 1];
    const day = row[SHUCHO_SRC_COL.DAY - 1];
    if (!year || !month || !day) return;  // 空行・不正な行はスキップ

    rows.push([
      shuchoNormalizeYear_(year),
      shuchoCleanValue_(month),
      shuchoCleanValue_(day),
      shuchoCleanValue_(row[SHUCHO_SRC_COL.NAME - 1]),
      shuchoCleanValue_(row[SHUCHO_SRC_COL.BUSYO - 1]),
      shuchoCleanValue_(row[SHUCHO_SRC_COL.TIME - 1]),
      shuchoCleanValue_(row[SHUCHO_SRC_COL.PLACE - 1])
    ]);
  });
  console.log('出張同期: データ整形完了。件数=' + rows.length);

  if (!rows.length) {
    throw new Error('出張マスターから1件もデータを取得できませんでした。シート名・列構成を確認してください。');
  }

  const importSh = ss.getSheetByName(SHEET_NAMES.SHUCHO_IMPORT);
  if (!importSh) {
    throw new Error('「' + SHEET_NAMES.SHUCHO_IMPORT + '」シートが見つかりません');
  }

  // 説明文（1行目）を更新し、データ範囲（3行目以降）を洗い替える
  importSh.getRange(1, 1).setValue(
    '※ このシートは「出張マスター」から Apps Script (syncShuchoFromMaster) で自動取込されます。\n' +
    '手動で書き換えても、次回の同期で上書きされます。\n' +
    '【列構成】 A=年度  B=月  C=日  D=氏名  E=用務  F=時間  G=場所'
  );

  const clearLastRow = Math.max(importSh.getLastRow(), rows.length + 2);
  if (clearLastRow >= 3) {
    try {
      importSh.getRange(3, 1, clearLastRow - 2, 7).clearContent();
    } catch (e) {
      throw new Error('「出張取込」シートの3行目以降をクリアできませんでした。A3セルに古いQUERY/IMPORTRANGE数式が残っていないか確認してください: ' + (e.message || e));
    }
  }
  console.log('出張同期: 「出張取込」クリア完了。書き込み開始');
  importSh.getRange(3, 1, rows.length, 7).setValues(rows);

  console.log('出張同期完了: ' + rows.length + '件のデータを「出張取込」に書き込みました');
  return { ok: true, count: rows.length };
}

/**
 * 出張マスターのセル値を整形する（null/undefined/エラー値を空文字に）
 */
function shuchoCleanValue_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string' && v.indexOf('#') === 0) return '';  // #N/A, #REF! など
  if (typeof v === 'number') return v;
  return String(v).trim();
}

/**
 * 出張マスターの「年」列は、令和の年数（例: 8）で入っていることがある。
 * ダッシュボード側は西暦（例: 2026）で日付を照合するため、
 * 100未満の小さい数値は「令和のX年」とみなして西暦に変換する。
 * すでに西暦（1900以上）で入っている場合はそのまま使う。
 */
function shuchoNormalizeYear_(rawYear) {
  const n = Number(rawYear);
  if (!n || isNaN(n)) return '';
  return (n < 100) ? (n + 2018) : n;
}

/**
 * ダッシュボードの「🔄 出張を反映」ボタン（管理者用）から呼ばれる。
 * 使い方: gas/shucho-sync-button.html を参照し、index.html にボタンを追加する。
 */
function manualSyncShuchoFromMaster() {
  const user = getCurrentUser();
  if (!user.ok) return { ok: false, error: user.reason };
  if (!user.isAdmin) return { ok: false, error: '管理者のみ実行できます' };
  try {
    const result = syncShuchoFromMaster();
    return { ok: true, count: result.count };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

/**
 * （任意）1時間おきに出張同期を自動実行するトリガーを設定する
 */
function setupShuchoSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'syncShuchoFromMaster') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncShuchoFromMaster')
    .timeBased()
    .everyHours(1)
    .create();
  console.log('出張の自動同期トリガーを設定しました（1時間おき）');
  return { ok: true, message: '1時間おきに出張マスターから自動同期されます' };
}
