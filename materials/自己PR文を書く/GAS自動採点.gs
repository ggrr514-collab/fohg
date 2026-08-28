/**
 * 「自分の歩みを言葉にする ―自己PR文を書こう―」AI自動採点スクリプト
 *
 * 教員用評価シート（本ファイルをコンテナバインドしたスプレッドシート）に設置する。
 * 生徒は Google Classroom 経由で配布された「生徒用ワークシート」（Googleスプレッドシート化した
 * 生徒用ワークシート.xlsx）に記入し、教員が指定フォルダにそれらを集約している前提。
 *
 * 運用フロー：
 *   第1時・第2時 → 生徒が「発問・回答」シートに記入
 *   第2時終了後 → 教員が【Step1：回答を収集】（クラス・番号順にソートして「回答入力」シートへ集約）
 *              → 教員が【Step2：一括自動採点】（F列の自己PR文をAIが評価し、G・H列に記入）
 *              → 教員がH列のコメントを確認し、I列に最終評価（教師評価）を入力
 *              → 教員が【Step3：AI評価を生徒に返却】（クラス+番号で突合し、生徒シートC28へ書き込み）
 *   その後      → 生徒がAI評価を見て、J列（自己評価）とK・L列（評価差考察・振り返り）を記入
 *              → 教員が再度【Step1：回答を収集】して最終記録を残す
 */

const CONFIG = {
  STUDENT: {
    INFO_SHEET: '単元ガイド',
    CLASS_CELL: 'A2',
    NUMBER_CELL: 'B2',
    NAME_CELL: 'D2',
    ANSWER_SHEET: '発問・回答',
    Q1_CELL: 'A7',
    Q2_CELL: 'A11',
    ESSAY_CELL: 'A23',
    AI_EVAL_CELL: 'C28',
    SELF_EVAL_CELL: 'C29',
    DIFF_CELL: 'A33',
    REFLECT_CELL: 'A38',
  },
  COL: {
    CLASS: 1, NUMBER: 2, NAME: 3,
    Q1: 4, Q2: 5, ESSAY: 6,
    GRADE: 7, COMMENT: 8, TEACHER: 9,
    SELF_EVAL: 10, DIFF: 11, REFLECT: 12,
  },
  TOTAL_COLS: 12,
  CFG_SHEET: 'AI設定',
  DATA_SHEET: '回答入力',
  DATA_START: 4,
  DATA_END: 43,
};

// ============================================================
// メニュー
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 AI採点')
    .addItem('📂 Step1：回答を収集', 'collectOnly')
    .addItem('📊 Step2：一括自動採点', 'gradeOnly')
    .addItem('📤 Step3：AI評価を生徒に返却', 'returnOnly')
    .addSeparator()
    .addItem('🔄 採点結果クリア', 'clearGradingResults')
    .addItem('❓ セットアップガイド', 'showSetupGuide')
    .addToUi();
}

function showSetupGuide() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'セットアップガイド',
    '① 「拡張機能」→「Apps Script」→「プロジェクトの設定」→「スクリプト プロパティ」で\n' +
    '   GEMINI_API_KEY を設定してください。\n' +
    '② 「AI設定」シートのB11に、生徒ワークシートを集めたGoogle DriveフォルダのIDを入力してください。\n' +
    '③ Step1→Step2→（教師評価の確認・入力）→Step3 の順に実行してください。\n' +
    '④ 初回実行時は権限の承認ダイアログが表示されます。承認してください。',
    ui.ButtonSet.OK
  );
}

// ============================================================
// 共通ユーティリティ
// ============================================================
function getCfgSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CFG_SHEET);
}

function getDataSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DATA_SHEET);
}

function getFolderId_() {
  const cfg = getCfgSheet_();
  const raw = String(cfg.getRange('B11').getValue() || '').trim();
  if (!raw || raw.startsWith('（') || raw.length < 10) {
    throw new Error('「AI設定」シートB11に、生徒ワークシートを集めたフォルダのIDを入力してください。');
  }
  return raw;
}

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new Error('スクリプト プロパティに GEMINI_API_KEY が設定されていません。'
      + '「拡張機能」→「Apps Script」→「プロジェクトの設定」から設定してください。');
  }
  return key;
}

function getModel_() {
  const cfg = getCfgSheet_();
  const v = String(cfg.getRange('B3').getValue() || '').trim();
  return v || 'gemini-2.0-flash';
}

function getTemperature_() {
  const cfg = getCfgSheet_();
  const v = cfg.getRange('B4').getValue();
  return (typeof v === 'number') ? v : 0.1;
}

// ============================================================
// Step1：回答を収集
// ============================================================
function collectOnly() {
  const ui = SpreadsheetApp.getUi();
  let folderId;
  try {
    folderId = getFolderId_();
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
    return;
  }

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

  const records = [];
  while (files.hasNext()) {
    const file = files.next();
    try {
      const record = readStudentFile_(file);
      if (record) records.push(record);
    } catch (e) {
      Logger.log('読み込み失敗: ' + file.getName() + ' / ' + e.message);
    }
    Utilities.sleep(500); // API制限対策
  }

  records.sort((a, b) => {
    if (a.klass !== b.klass) return a.klass - b.klass;
    return a.number - b.number;
  });

  const sheet = getDataSheet_();
  const existing = getExistingGradesMap_(sheet);

  let row = CONFIG.DATA_START;
  records.forEach((rec) => {
    if (row > CONFIG.DATA_END) return;
    const key = rec.klass + '-' + rec.number;
    const prev = existing[key] || {};
    sheet.getRange(row, CONFIG.COL.CLASS).setValue(rec.klass);
    sheet.getRange(row, CONFIG.COL.NUMBER).setValue(rec.number);
    sheet.getRange(row, CONFIG.COL.NAME).setValue(rec.name);
    sheet.getRange(row, CONFIG.COL.Q1).setValue(rec.q1);
    sheet.getRange(row, CONFIG.COL.Q2).setValue(rec.q2);
    sheet.getRange(row, CONFIG.COL.ESSAY).setValue(rec.essay);
    sheet.getRange(row, CONFIG.COL.SELF_EVAL).setValue(rec.selfEval);
    sheet.getRange(row, CONFIG.COL.DIFF).setValue(rec.diff);
    sheet.getRange(row, CONFIG.COL.REFLECT).setValue(rec.reflect);
    // 採点済みの結果は保持する（再収集で消さない）
    if (prev.grade) sheet.getRange(row, CONFIG.COL.GRADE).setValue(prev.grade);
    if (prev.comment) sheet.getRange(row, CONFIG.COL.COMMENT).setValue(prev.comment);
    if (prev.teacher) sheet.getRange(row, CONFIG.COL.TEACHER).setValue(prev.teacher);
    row++;
  });

  ui.alert('収集完了', records.length + '件の回答を収集しました。', ui.ButtonSet.OK);
}

function getExistingGradesMap_(sheet) {
  const map = {};
  const lastRow = Math.min(sheet.getLastRow(), CONFIG.DATA_END);
  if (lastRow < CONFIG.DATA_START) return map;
  const values = sheet.getRange(CONFIG.DATA_START, 1, lastRow - CONFIG.DATA_START + 1,
    CONFIG.TOTAL_COLS).getValues();
  values.forEach((v) => {
    const klass = v[CONFIG.COL.CLASS - 1];
    const number = v[CONFIG.COL.NUMBER - 1];
    if (klass === '' || number === '') return;
    map[klass + '-' + number] = {
      grade: v[CONFIG.COL.GRADE - 1],
      comment: v[CONFIG.COL.COMMENT - 1],
      teacher: v[CONFIG.COL.TEACHER - 1],
    };
  });
  return map;
}

function readStudentFile_(file) {
  const ss = SpreadsheetApp.open(file);
  const infoSheet = ss.getSheetByName(CONFIG.STUDENT.INFO_SHEET);
  const answerSheet = ss.getSheetByName(CONFIG.STUDENT.ANSWER_SHEET);
  if (!infoSheet || !answerSheet) return null;

  const klass = infoSheet.getRange(CONFIG.STUDENT.CLASS_CELL).getValue();
  const number = infoSheet.getRange(CONFIG.STUDENT.NUMBER_CELL).getValue();
  const name = infoSheet.getRange(CONFIG.STUDENT.NAME_CELL).getValue();
  if (!klass || !number) return null;

  return {
    klass: Number(klass),
    number: Number(number),
    name: String(name || ''),
    q1: String(answerSheet.getRange(CONFIG.STUDENT.Q1_CELL).getValue() || ''),
    q2: String(answerSheet.getRange(CONFIG.STUDENT.Q2_CELL).getValue() || ''),
    essay: String(answerSheet.getRange(CONFIG.STUDENT.ESSAY_CELL).getValue() || ''),
    selfEval: String(answerSheet.getRange(CONFIG.STUDENT.SELF_EVAL_CELL).getValue() || ''),
    diff: String(answerSheet.getRange(CONFIG.STUDENT.DIFF_CELL).getValue() || ''),
    reflect: String(answerSheet.getRange(CONFIG.STUDENT.REFLECT_CELL).getValue() || ''),
  };
}

// ============================================================
// Step2：一括自動採点
// ============================================================
function gradeOnly() {
  const ui = SpreadsheetApp.getUi();
  let apiKey;
  try {
    apiKey = getApiKey_();
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
    return;
  }

  const sheet = getDataSheet_();
  const lastRow = Math.min(sheet.getLastRow(), CONFIG.DATA_END);
  let gradedCount = 0;
  let errorCount = 0;

  for (let row = CONFIG.DATA_START; row <= lastRow; row++) {
    const essay = String(sheet.getRange(row, CONFIG.COL.ESSAY).getValue() || '').trim();
    const existingGrade = String(sheet.getRange(row, CONFIG.COL.GRADE).getValue() || '').trim();
    if (!essay) continue;
    if (existingGrade) continue; // 採点済みはスキップ

    try {
      const result = callGeminiForGrading_(essay, apiKey);
      sheet.getRange(row, CONFIG.COL.GRADE).setValue(result.grade);
      sheet.getRange(row, CONFIG.COL.COMMENT).setValue(result.comment);
      gradedCount++;
    } catch (e) {
      Logger.log('採点失敗（行' + row + '）: ' + e.message);
      sheet.getRange(row, CONFIG.COL.COMMENT).setValue('採点エラー: ' + e.message);
      errorCount++;
    }
    Utilities.sleep(1500); // API制限対策
  }

  ui.alert('採点完了', gradedCount + '件を採点しました。'
    + (errorCount > 0 ? '（エラー ' + errorCount + '件）' : ''), ui.ButtonSet.OK);
}

function buildGradingPrompt_(essayText) {
  return [
    'あなたは中学校国語の教師です。',
    '以下のルーブリックに基づいて、中学3年生が書いた「自己PR文」を評価してください。',
    'この文章は、高校入試の自己評価資料の自己PR欄（200〜300字）として書かれたものです。',
    '',
    '【ルーブリック】',
    'S：具体的な体験と、そこから得た学びの両方が、独自性・説得力をもって明確に書かれている。',
    'A：具体的な体験と、そこから得た学びの両方が明確に書かれている。',
    'B：具体的な体験、学びのどちらか一方のみが書かれている。',
    'C：体験も学びも抽象的・一般的な表現にとどまっている。',
    'D：題材が定まっておらず、自己PR文として内容が成立していない。',
    '',
    '評価のポイント：',
    '1. 具体的な体験（エピソード）が書かれているか',
    '2. その体験から得た学び・成長した力が、体験と結びつけて書かれているか',
    '',
    '※ 1と2の両方に具体的に触れていなければ、A以上にはなりません。',
    '※ 文章の巧拙（表現の上手い下手）そのものではなく、内容の具体性・結びつきを重視して評価してくだ' +
    'さい。',
    '',
    '【生徒の自己PR文】',
    essayText,
    '',
    '以下のJSON形式のみで回答してください。他の文章は一切含めないでください。',
    '{"grade":"S","comment":"（100字以内で、良い点と改善点を中学3年生にわかりやすく）"}',
  ].join('\n');
}

function callGeminiForGrading_(essayText, apiKey) {
  const model = getModel_();
  const temperature = getTemperature_();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model
    + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [{ parts: [{ text: buildGradingPrompt_(essayText) }] }],
    generationConfig: {
      temperature: temperature,
      responseMimeType: 'application/json',
    },
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('APIエラー(' + code + '): ' + response.getContentText());
  }

  const json = JSON.parse(response.getContentText());
  const text = json.candidates && json.candidates[0] && json.candidates[0].content
    && json.candidates[0].content.parts && json.candidates[0].content.parts[0].text;
  if (!text) throw new Error('APIレスポンスの形式が不正です。');

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    const match = text.match(/\{[\s\S]*\}/);
    if (!match) throw new Error('JSON形式の応答が取得できませんでした: ' + text);
    parsed = JSON.parse(match[0]);
  }

  const grade = String(parsed.grade || '').toUpperCase().trim();
  if (['S', 'A', 'B', 'C', 'D'].indexOf(grade) === -1) {
    throw new Error('不正な評価値です: ' + grade);
  }
  return { grade: grade, comment: String(parsed.comment || '') };
}

// ============================================================
// Step3：AI評価を生徒に返却
// ============================================================
function returnOnly() {
  const ui = SpreadsheetApp.getUi();
  let folderId;
  try {
    folderId = getFolderId_();
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
    return;
  }

  const sheet = getDataSheet_();
  const lastRow = Math.min(sheet.getLastRow(), CONFIG.DATA_END);
  const folder = DriveApp.getFolderById(folderId);

  // クラス+番号 → 教師最終評価（未入力ならAI評価） のマップを作成
  const gradeMap = {};
  for (let row = CONFIG.DATA_START; row <= lastRow; row++) {
    const klass = sheet.getRange(row, CONFIG.COL.CLASS).getValue();
    const number = sheet.getRange(row, CONFIG.COL.NUMBER).getValue();
    if (!klass || !number) continue;
    const teacherGrade = String(sheet.getRange(row, CONFIG.COL.TEACHER).getValue() || '').trim();
    const aiGrade = String(sheet.getRange(row, CONFIG.COL.GRADE).getValue() || '').trim();
    const finalGrade = teacherGrade || aiGrade;
    if (!finalGrade) continue;
    gradeMap[klass + '-' + number] = finalGrade;
  }

  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  let returnedCount = 0;
  while (files.hasNext()) {
    const file = files.next();
    try {
      const ss = SpreadsheetApp.open(file);
      const infoSheet = ss.getSheetByName(CONFIG.STUDENT.INFO_SHEET);
      const answerSheet = ss.getSheetByName(CONFIG.STUDENT.ANSWER_SHEET);
      if (!infoSheet || !answerSheet) continue;

      const klass = infoSheet.getRange(CONFIG.STUDENT.CLASS_CELL).getValue();
      const number = infoSheet.getRange(CONFIG.STUDENT.NUMBER_CELL).getValue();
      const key = klass + '-' + number;
      if (gradeMap[key]) {
        answerSheet.getRange(CONFIG.STUDENT.AI_EVAL_CELL).setValue(gradeMap[key]);
        returnedCount++;
      }
    } catch (e) {
      Logger.log('返却失敗: ' + file.getName() + ' / ' + e.message);
    }
    Utilities.sleep(500);
  }

  ui.alert('返却完了', returnedCount + '件のAI評価を返却しました。', ui.ButtonSet.OK);
}

// ============================================================
// 採点結果クリア
// ============================================================
function clearGradingResults() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('確認', 'AI評価・AIコメント・教師最終評価を全てクリアします。よろしいですか？',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const sheet = getDataSheet_();
  const lastRow = Math.min(Math.max(sheet.getLastRow(), CONFIG.DATA_START), CONFIG.DATA_END);
  const numRows = lastRow - CONFIG.DATA_START + 1;
  if (numRows <= 0) return;

  sheet.getRange(CONFIG.DATA_START, CONFIG.COL.GRADE, numRows, 1).clearContent();
  sheet.getRange(CONFIG.DATA_START, CONFIG.COL.COMMENT, numRows, 1).clearContent();
  sheet.getRange(CONFIG.DATA_START, CONFIG.COL.TEACHER, numRows, 1).clearContent();

  ui.alert('クリア完了', '採点結果をクリアしました。', ui.ButtonSet.OK);
}
