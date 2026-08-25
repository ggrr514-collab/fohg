/**
 * ============================================================
 * 行事取込データを「R8_行事予定Ver 2」から直接取得する
 * ============================================================
 * これまで「行事取込」シートは、
 *   =IMPORTRANGE(設定!B4, "行事取込!A3:AF")
 * という数式で「行事マスター(改)」という中間スプレッドシートから
 * 取り込んでいました。
 *
 * このコードは、その中間スプレッドシートを廃止し、正式な
 * 「R8_行事予定Ver 2」(月ごとのシート「４月」〜「３月」)から
 * 直接データを抽出して「行事取込」シートに書き込みます。
 *
 * 【重要】「R8_行事予定Ver 2」側の月別シート(４月〜３月)の列構成:
 *   C=日 D=曜 E=週 F=行事計画等 H=日直 I=清掃 J=日課
 *   K-M=給食(1-3年) N-T=1年①〜⑥+時数 U-AA=2年①〜⑥+時数
 *   AB-AH=3年①〜⑥+時数 AI=部活
 * これを「行事取込」シートの32列フォーマット(A=日付...AF=部活)に
 * そのままマッピングします。
 *
 * 【導入手順】
 * 1. このファイルの中身をまるごとコピーして、Apps Scriptエディタで
 *    新しいファイル（例: R8EventSync.gs）を作り、貼り付けて保存する。
 * 2. 「設定」シートの「行事マスターSS_ID」の値(B4セル)を
 *    R8_行事予定Ver 2 のスプレッドシートID
 *      1VB7PDDjjHrcMpSoUglT0q4D-Bo2FhlgEEZd1Iqr3H4k
 *    に書き換える。
 * 3. 「行事取込」シートの A3セルに残っている
 *      =IMPORTRANGE(設定!B4, "行事取込!A3:AF")
 *    という数式を削除する（Delete）。空にしておけばOK。
 * 4. Apps Scriptエディタの関数選択で syncEventImportFromR8 を選び、
 *    「実行」を1回押す（初回はアクセス許可の承認が必要）。
 *    → 「行事取込」シートに R8_行事予定Ver 2 のデータが書き込まれれば成功。
 * 5. 続けて setupR8EventSyncTrigger を1回だけ実行する。
 *    → 1時間おきに自動で再同期されるようになる（手動で毎回実行しなくてよい）。
 *
 * 【R8_行事予定Ver 2 が来年度以降差し替わる場合】
 * 新しい年度のファイルができたら、「設定」シートのIDを新しいものに
 * 書き換えて syncEventImportFromR8 を再実行するだけでよい。
 * ============================================================
 */

// R8_行事予定Ver 2 側、月別シートの名前（「1月」だけ半角数字なので注意）
const R8_MONTH_SHEET_NAMES = [
  '４月', '５月', '６月', '７月', '８月', '９月',
  '１０月', '１１月', '１２月', '1月', '２月', '３月'
];

// R8_行事予定Ver 2 側、月別シートの列番号（1-indexed）
const R8_COL = {
  DAY: 3, WEEK: 5, EVENT: 6, NICCHOKU: 8, SEISO: 9, NIKKA: 10,
  LUNCH1: 11, LUNCH2: 12, LUNCH3: 13,
  G1_PERIODS: [14, 15, 16, 17, 18, 19], G1_JISU: 20,
  G2_PERIODS: [21, 22, 23, 24, 25, 26], G2_JISU: 27,
  G3_PERIODS: [28, 29, 30, 31, 32, 33], G3_JISU: 34,
  BUKATSU: 35
};
const R8_DATA_START_ROW = 5;      // 月別シートのデータ開始行（日=1のある行）
const R8_YEAR_CELL = 'C1';        // 月別シートの年（西暦）が入っているセル

/**
 * 「行事取込」シートを R8_行事予定Ver 2 の内容で洗い替える
 * 手動実行 / 時間主導トリガーの両方から呼ばれる想定
 */
function syncEventImportFromR8() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings_(ss);
  const masterId = settings['行事マスターSS_ID'] || settings['行事マスターID'];
  if (!masterId) {
    throw new Error('設定シートに「行事マスターSS_ID」（R8_行事予定Ver 2 のID）が設定されていません');
  }

  let masterSs;
  try {
    masterSs = SpreadsheetApp.openById(String(masterId).trim());
  } catch (e) {
    throw new Error('R8_行事予定Ver 2 を開けませんでした（IDまたは共有設定を確認してください）: ' + (e.message || e));
  }

  const rows = [];
  R8_MONTH_SHEET_NAMES.forEach(sheetName => {
    try {
      const monthRows = extractR8MonthRows_(masterSs, sheetName);
      rows.push(...monthRows);
    } catch (e) {
      console.log('R8同期: 「' + sheetName + '」シートの読込でエラー: ' + (e.message || e));
    }
  });

  if (!rows.length) {
    throw new Error('R8_行事予定Ver 2 から1件もデータを取得できませんでした。シート名・列構成を確認してください。');
  }

  // 日付順に整列（月別シートを順番に読んでいるので基本的には既に順序通り）
  rows.sort((a, b) => a[0].getTime() - b[0].getTime());

  const importSh = ss.getSheetByName(SHEET_NAMES.EVENT_IMPORT);
  if (!importSh) {
    throw new Error('「' + SHEET_NAMES.EVENT_IMPORT + '」シートが見つかりません');
  }

  // 説明文（1行目）を更新し、データ範囲（3行目以降）を洗い替える
  importSh.getRange(1, 1).setValue(
    '※ このシートは「R8_行事予定Ver 2」(月別シート)から Apps Script (syncEventImportFromR8) で自動取込されます。\n' +
    '手動で書き換えても、次回の同期で上書きされます。\n' +
    '【列構成】 A=日付  B=曜日  C=週  D=行事  E=日直  F=清掃  G=日課  H-J=給食  K-AE=学年校時  AF=部活'
  );

  const lastRow = Math.max(importSh.getLastRow(), rows.length + 2);
  if (lastRow >= 3) {
    importSh.getRange(3, 1, lastRow - 2, 32).clearContent();
  }
  importSh.getRange(3, 1, rows.length, 32).setValues(rows);

  console.log('R8同期完了: ' + rows.length + '件のデータを「行事取込」に書き込みました');
  return { ok: true, count: rows.length };
}

/**
 * R8_行事予定Ver 2 の月別シート1枚から、日ごとの32列データ配列を作る
 */
function extractR8MonthRows_(masterSs, sheetName) {
  const sh = masterSs.getSheetByName(sheetName);
  if (!sh) {
    console.log('R8同期: シート「' + sheetName + '」が見つかりません（スキップ）');
    return [];
  }

  const year = parseInt(sh.getRange(R8_YEAR_CELL).getValue(), 10);
  if (!year || isNaN(year)) {
    console.log('R8同期: 「' + sheetName + '」の年（' + R8_YEAR_CELL + '）を取得できません（スキップ）');
    return [];
  }
  const month = parseInt(String(sheetName).replace(/[^0-9０-９]/g, '').replace(/[０-９]/g, ch =>
    String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30)), 10);
  if (!month) return [];

  // データ範囲をまとめて読み込む（1日〜31日分、35列）
  const maxDays = 31;
  const data = sh.getRange(R8_DATA_START_ROW, 1, maxDays, 35).getValues();

  const out = [];
  data.forEach(row => {
    const day = row[R8_COL.DAY - 1];
    if (!day || typeof day !== 'number') return;  // 日付が無い行（月末以降の余白行）はスキップ

    const date = new Date(year, month - 1, day);
    if (isNaN(date.getTime()) || date.getMonth() !== month - 1) return;  // 存在しない日付はスキップ（例: 2/30）

    const weekdayJp = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
    const get = (col) => r8CleanValue_(row[col - 1]);
    const getPeriods = (cols) => cols.map(get);

    out.push([
      date, weekdayJp, get(R8_COL.WEEK), get(R8_COL.EVENT),
      get(R8_COL.NICCHOKU), get(R8_COL.SEISO), get(R8_COL.NIKKA),
      get(R8_COL.LUNCH1), get(R8_COL.LUNCH2), get(R8_COL.LUNCH3),
      ...getPeriods(R8_COL.G1_PERIODS), get(R8_COL.G1_JISU),
      ...getPeriods(R8_COL.G2_PERIODS), get(R8_COL.G2_JISU),
      ...getPeriods(R8_COL.G3_PERIODS), get(R8_COL.G3_JISU),
      get(R8_COL.BUKATSU)
    ]);
  });

  return out;
}

/**
 * R8シートのセル値を整形する
 * - #N/A 等のエラー文字列は空文字に（土日の日直欄などで発生）
 * - null/undefined は空文字に
 * - それ以外は文字列化してトリム
 */
function r8CleanValue_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string' && v.indexOf('#') === 0) return '';  // #N/A, #REF! など
  if (typeof v === 'number') return v;
  return String(v).trim();
}

/**
 * 1時間おきにR8同期を自動実行するトリガーを設定する（初回1回だけ実行）
 */
function setupR8EventSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'syncEventImportFromR8') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncEventImportFromR8')
    .timeBased()
    .everyHours(1)
    .create();
  console.log('R8行事予定の自動同期トリガーを設定しました（1時間おき）');
  return { ok: true, message: '1時間おきに「R8_行事予定Ver 2」から自動同期されます' };
}
