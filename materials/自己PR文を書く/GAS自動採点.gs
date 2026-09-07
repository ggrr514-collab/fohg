/**
 * 「自分の歩みを言葉にする ―自己PR文を書こう―」AI自動採点スクリプト
 *
 * 教員用評価シート（本ファイルをコンテナバインドしたスプレッドシート）に設置する。
 * 生徒は Google Classroom 経由で配布された「生徒用ワークシート」（Googleスプレッドシート化した
 * 生徒用ワークシート.xlsx）に記入し、教員が指定フォルダにそれらを集約している前提。
 *
 * 本単元は「書くこと」（第1時・自己PR文）と「話すこと・聞くこと」（第2時・聞き取り評価）の複合
 * 単元のため、2種類の評価対象をそれぞれAIが採点する。
 *
 * 運用フロー：
 *   第1時 → 生徒が自己PR文を執筆
 *   第2時 → 生徒がグループ発表を聞き、聞き取り評価①〜③を記入
 *   単元終了後 → 教員が【Step1：回答を収集】（クラス・番号順にソートして「回答入力」シートへ集約）
 *            → 教員が【Step2：一括自動採点】
 *                ・自己PR文（F列）→ G・H列（AI評価・AIコメント）
 *                ・聞き取り評価①〜③（J〜L列）→ M・N列（AI評価・AIコメント）
 *            → 教員がAIコメントを確認し、I列（自己PR文）・O列（聞き取り評価）に最終評価を入力
 *
 * 生徒への個別返却フローは設けない。AI評価は教員用シートで完結し、教師が最終チェックしたうえで
 * 指導に生かす。
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
    LISTEN1_NAME_CELL: 'B31',
    LISTEN1_COMMENT_CELL: 'A33',
    LISTEN2_NAME_CELL: 'B35',
    LISTEN2_COMMENT_CELL: 'A37',
    LISTEN3_NAME_CELL: 'B39',
    LISTEN3_COMMENT_CELL: 'A41',
    REFLECT_CELL: 'A46',
  },
  COL: {
    CLASS: 1, NUMBER: 2, NAME: 3,
    Q1: 4, Q2: 5, ESSAY: 6,
    ESSAY_GRADE: 7, ESSAY_COMMENT: 8, ESSAY_TEACHER: 9,
    LISTEN1: 10, LISTEN2: 11, LISTEN3: 12,
    LISTEN_GRADE: 13, LISTEN_COMMENT: 14, LISTEN_TEACHER: 15,
    REFLECT: 16,
  },
  TOTAL_COLS: 16,
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
    '② 「AI設定」シートのB12に、生徒ワークシートを集めたGoogle DriveフォルダのIDを入力してください。\n' +
    '③ Step1→Step2の順に実行し、AIコメントを確認したうえで、教師最終評価（I列・O列）を入力して' +
    'ください。\n' +
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
  const raw = String(cfg.getRange('B12').getValue() || '').trim();
  if (!raw || raw.startsWith('（') || raw.length < 10) {
    throw new Error('「AI設定」シートB12に、生徒ワークシートを集めたフォルダのIDを入力してください。');
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
    sheet.getRange(row, CONFIG.COL.LISTEN1).setValue(rec.listen1);
    sheet.getRange(row, CONFIG.COL.LISTEN2).setValue(rec.listen2);
    sheet.getRange(row, CONFIG.COL.LISTEN3).setValue(rec.listen3);
    sheet.getRange(row, CONFIG.COL.REFLECT).setValue(rec.reflect);
    // 採点済みの結果は保持する（再収集で消さない）
    if (prev.essayGrade) sheet.getRange(row, CONFIG.COL.ESSAY_GRADE).setValue(prev.essayGrade);
    if (prev.essayComment) sheet.getRange(row, CONFIG.COL.ESSAY_COMMENT).setValue(prev.essayComment);
    if (prev.essayTeacher) sheet.getRange(row, CONFIG.COL.ESSAY_TEACHER).setValue(prev.essayTeacher);
    if (prev.listenGrade) sheet.getRange(row, CONFIG.COL.LISTEN_GRADE).setValue(prev.listenGrade);
    if (prev.listenComment) {
      sheet.getRange(row, CONFIG.COL.LISTEN_COMMENT).setValue(prev.listenComment);
    }
    if (prev.listenTeacher) sheet.getRange(row, CONFIG.COL.LISTEN_TEACHER).setValue(prev.listenTeacher);
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
      essayGrade: v[CONFIG.COL.ESSAY_GRADE - 1],
      essayComment: v[CONFIG.COL.ESSAY_COMMENT - 1],
      essayTeacher: v[CONFIG.COL.ESSAY_TEACHER - 1],
      listenGrade: v[CONFIG.COL.LISTEN_GRADE - 1],
      listenComment: v[CONFIG.COL.LISTEN_COMMENT - 1],
      listenTeacher: v[CONFIG.COL.LISTEN_TEACHER - 1],
    };
  });
  return map;
}

function formatListenEntry_(name, comment) {
  name = String(name || '').trim();
  comment = String(comment || '').trim();
  if (!comment) return '';
  return name ? (name + '：' + comment) : comment;
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

  const s = CONFIG.STUDENT;
  return {
    klass: Number(klass),
    number: Number(number),
    name: String(name || ''),
    q1: String(answerSheet.getRange(s.Q1_CELL).getValue() || ''),
    q2: String(answerSheet.getRange(s.Q2_CELL).getValue() || ''),
    essay: String(answerSheet.getRange(s.ESSAY_CELL).getValue() || ''),
    listen1: formatListenEntry_(
      answerSheet.getRange(s.LISTEN1_NAME_CELL).getValue(),
      answerSheet.getRange(s.LISTEN1_COMMENT_CELL).getValue()),
    listen2: formatListenEntry_(
      answerSheet.getRange(s.LISTEN2_NAME_CELL).getValue(),
      answerSheet.getRange(s.LISTEN2_COMMENT_CELL).getValue()),
    listen3: formatListenEntry_(
      answerSheet.getRange(s.LISTEN3_NAME_CELL).getValue(),
      answerSheet.getRange(s.LISTEN3_COMMENT_CELL).getValue()),
    reflect: String(answerSheet.getRange(s.REFLECT_CELL).getValue() || ''),
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
    // ① 自己PR文（書くこと）
    const essay = String(sheet.getRange(row, CONFIG.COL.ESSAY).getValue() || '').trim();
    const existingEssayGrade = String(sheet.getRange(row, CONFIG.COL.ESSAY_GRADE).getValue() || '')
      .trim();
    if (essay && !existingEssayGrade) {
      try {
        const result = callGeminiForGrading_(buildEssayGradingPrompt_(essay), apiKey);
        sheet.getRange(row, CONFIG.COL.ESSAY_GRADE).setValue(result.grade);
        sheet.getRange(row, CONFIG.COL.ESSAY_COMMENT).setValue(result.comment);
        gradedCount++;
      } catch (e) {
        Logger.log('自己PR文の採点失敗（行' + row + '）: ' + e.message);
        sheet.getRange(row, CONFIG.COL.ESSAY_COMMENT).setValue('採点エラー: ' + e.message);
        errorCount++;
      }
      Utilities.sleep(1500); // API制限対策
    }

    // ② 聞き取り評価（聞くこと）
    const listenTexts = [
      String(sheet.getRange(row, CONFIG.COL.LISTEN1).getValue() || '').trim(),
      String(sheet.getRange(row, CONFIG.COL.LISTEN2).getValue() || '').trim(),
      String(sheet.getRange(row, CONFIG.COL.LISTEN3).getValue() || '').trim(),
    ].filter((t) => t);
    const existingListenGrade = String(sheet.getRange(row, CONFIG.COL.LISTEN_GRADE).getValue() || '')
      .trim();
    if (listenTexts.length > 0 && !existingListenGrade) {
      try {
        const result = callGeminiForGrading_(buildListeningGradingPrompt_(listenTexts), apiKey);
        sheet.getRange(row, CONFIG.COL.LISTEN_GRADE).setValue(result.grade);
        sheet.getRange(row, CONFIG.COL.LISTEN_COMMENT).setValue(result.comment);
        gradedCount++;
      } catch (e) {
        Logger.log('聞き取り評価の採点失敗（行' + row + '）: ' + e.message);
        sheet.getRange(row, CONFIG.COL.LISTEN_COMMENT).setValue('採点エラー: ' + e.message);
        errorCount++;
      }
      Utilities.sleep(1500); // API制限対策
    }
  }

  ui.alert('採点完了', gradedCount + '件を採点しました。'
    + (errorCount > 0 ? '（エラー ' + errorCount + '件）' : ''), ui.ButtonSet.OK);
}

function buildEssayGradingPrompt_(essayText) {
  return [
    'あなたは中学校国語の教師です。',
    '以下のルーブリックに基づいて、中学3年生が書いた「自己PR文」を評価してください。',
    'この文章は、高校入試の自己評価資料の自己PR欄（300〜500字）として書かれたものです。',
    '',
    '【ルーブリック（書くこと）】',
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

function buildListeningGradingPrompt_(listenTexts) {
  const joined = listenTexts.map((t, i) => '(' + (i + 1) + ') ' + t).join('\n');
  return [
    'あなたは中学校国語の教師です。',
    '以下は、ある生徒がグループ内の発表（自己PR文の発表）を聞いて書いた「聞き取り評価」の記述です'
    + '（発表者ごとに1件ずつ）。',
    'この生徒自身の発表内容は評価の対象ではありません。あくまで、この生徒が「聞き手」として書いた'
    + '記述の質を、以下のルーブリックに基づいて評価してください。',
    '',
    '【ルーブリック（聞くこと）】',
    'S：3つの観点（体験の具体性・学びの独自性・応答からの理解の深まり）すべてについて、話し手の言葉'
    + 'を具体的な根拠として挙げながら評価し、自分なりの気づきや考えも述べている。',
    'A：3つの観点のうち少なくとも2つについて、話し手の言葉を具体的な根拠として挙げながら評価してい'
    + 'る。',
    'B：観点には触れているが、根拠がやや一般的（「わかりやすかった」「よかった」など）にとどまって'
    + 'いる。',
    'C：観点に基づかない感想（「面白かった」「すごいと思った」など）のみで、評価として成立していな'
    + 'い。',
    'D：記述がない、または発表内容と無関係な記述にとどまっている。',
    '',
    '評価のポイント：',
    '1. 話し手の言葉を引用・要約するなど、具体的な根拠を伴って評価しているか',
    '2. 単なる感想（「よかった」「面白かった」等）で終わっていないか',
    '',
    '※ 具体的な根拠を伴わない感想のみの記述はB以下と評価してください。',
    '',
    '【この生徒が書いた聞き取り評価】',
    joined,
    '',
    '以下のJSON形式のみで回答してください。他の文章は一切含めないでください。',
    '{"grade":"S","comment":"（100字以内で、良い点と改善点を中学3年生にわかりやすく）"}',
  ].join('\n');
}

function callGeminiForGrading_(promptText, apiKey) {
  const model = getModel_();
  const temperature = getTemperature_();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model
    + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
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
// 採点結果クリア
// ============================================================
function clearGradingResults() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('確認',
    '自己PR文・聞き取り評価それぞれのAI評価・AIコメント・教師最終評価を全てクリアします。'
    + 'よろしいですか？',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const sheet = getDataSheet_();
  const lastRow = Math.min(Math.max(sheet.getLastRow(), CONFIG.DATA_START), CONFIG.DATA_END);
  const numRows = lastRow - CONFIG.DATA_START + 1;
  if (numRows <= 0) return;

  [CONFIG.COL.ESSAY_GRADE, CONFIG.COL.ESSAY_COMMENT, CONFIG.COL.ESSAY_TEACHER,
   CONFIG.COL.LISTEN_GRADE, CONFIG.COL.LISTEN_COMMENT, CONFIG.COL.LISTEN_TEACHER].forEach((col) => {
    sheet.getRange(CONFIG.DATA_START, col, numRows, 1).clearContent();
  });

  ui.alert('クリア完了', '採点結果をクリアしました。', ui.ButtonSet.OK);
}
