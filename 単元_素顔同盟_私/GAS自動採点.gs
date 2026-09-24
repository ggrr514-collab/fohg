/**
 * 中学3年 国語　単元「『私』を消すもの、『私』を支えるもの」
 * ――「素顔同盟」（すやまたけし）／「私」（三崎亜記）を批判的に読む
 *
 * 批評文（800字）のAI一次採点システム
 *   ステップ1：生徒のワークシートを集めて「回答入力」に転記
 *   ステップ2：批評文をAIが採点（思考・判断・表現 S/A/B/C/D）
 *   ステップ3：AI評価とコメントを生徒のワークシートに返却
 *
 * 【使う前の準備】
 *   1. 拡張機能 → Apps Script でこのコードを貼る
 *   2. メニュー「単元：素顔同盟・私」→「APIキーを登録」でGemini APIキーを入れる
 *      （APIキーはコードに直接書かない。スクリプト プロパティに保存される）
 *   3.「AI設定」シートのB11に、生徒の提出物が入ったDriveフォルダのIDを入れる
 */

const CONFIG = {
  // ---- 生徒用ワークシートのセル位置（生徒用ワークシート.xlsx と一致）----
  STUDENT: {
    INFO_SHEET:      '単元ガイド',
    CLASS_CELL:      'A2',
    NUMBER_CELL:     'B2',
    NAME_CELL:       'D2',
    ANSWER_SHEET:    '発問・回答',
    Q1_CELL:         'A6',    // 発問①（知技・語感）
    Q2_CELL:         'A10',   // 発問②（思判表・仮定判断）
    ESSAY_CELL:      'A25',   // 批評文（800字）★AI採点対象
    AI_EVAL_CELL:    'C29',   // ←返却（評価）
    AI_COMMENT_CELL: 'C30',   // ←返却（コメント）
    SELF_EVAL_CELL:  'C31',
    DIFF_CELL:       'A34',
    REFLECT_CELL:    'A38',
  },

  // ---- 教員用評価シート「回答入力」の列（12列）----
  COL: {
    CLASS: 1, NUMBER: 2, NAME: 3,
    Q1: 4, Q2: 5, ESSAY: 6,
    GRADE: 7, COMMENT: 8, TEACHER: 9,
    SELF_EVAL: 10, DIFF: 11, REFLECT: 12,
  },
  TOTAL_COLS: 12,

  CFG_SHEET:  'AI設定',
  DATA_SHEET: '回答入力',
  DATA_START: 4,
  DATA_END:   43,

  FILEMAP_SHEET: '_ファイル対応',   // 返却先を覚えておく作業用シート（非表示）

  CFG_CELL: {
    MODEL:  'B3',
    TEMP:   'B4',
    FOLDER: 'B11',
  },

  SLEEP_COLLECT: 500,
  SLEEP_GRADE:   1500,
};

/* ============================================================
 *  メニュー
 * ============================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('単元：素顔同盟・私')
    .addItem('① 生徒の回答を集める', 'step1_collect')
    .addItem('② 批評文をAIで採点する', 'step2_grade')
    .addItem('③ 評価を生徒に返す', 'step3_return')
    .addSeparator()
    .addItem('APIキーを登録', 'setApiKey')
    .addItem('採点結果をクリア（再採点用）', 'clearGrades')
    .addToUi();
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Gemini APIキーを入力してください',
                        'キーはスクリプト プロパティに保存され、シートには残りません。',
                        ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) { ui.alert('キーが空です。'); return; }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  ui.alert('APIキーを登録しました。');
}

/* ============================================================
 *  ステップ1：回答を集める
 * ============================================================ */
function step1_collect() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(CONFIG.CFG_SHEET);
  const data = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!cfg || !data) { ui.alert('「AI設定」または「回答入力」シートが見つかりません。'); return; }

  const folderId = String(cfg.getRange(CONFIG.CFG_CELL.FOLDER).getValue()).trim();
  if (!folderId || folderId.indexOf('（') === 0) {
    ui.alert('「AI設定」シートのB11に、提出物フォルダのIDを入力してください。');
    return;
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    ui.alert('フォルダを開けませんでした。IDを確認してください。\n' + e.message);
    return;
  }

  const S = CONFIG.STUDENT;
  const map = getFileMapSheet_();
  map.clear();
  map.getRange(1, 1, 1, 3).setValues([['クラス', '出席番号', 'ファイルID']]);

  const rows = [];
  const skipped = [];
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    let sheetId = null;
    let temp = null;
    try {
      const mime = file.getMimeType();
      if (mime === MimeType.GOOGLE_SHEETS) {
        sheetId = file.getId();
      } else if (mime === MimeType.MICROSOFT_EXCEL ||
                 mime === MimeType.MICROSOFT_EXCEL_LEGACY) {
        // xlsxのままでは読めないのでスプレッドシートに変換してから読む。
        // ※ サービス「Drive API」を有効にしておくこと（詳細なGoogleサービス）
        temp = Drive.Files.copy(
          { title: '[一時] ' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS },
          file.getId());
        sheetId = temp.id;
      } else {
        skipped.push(file.getName() + '（対象外の形式）');
        continue;
      }

      const wb = SpreadsheetApp.openById(sheetId);
      const info = wb.getSheetByName(S.INFO_SHEET);
      const ans  = wb.getSheetByName(S.ANSWER_SHEET);
      if (!info || !ans) { skipped.push(file.getName() + '（シート名が違う）'); continue; }

      const klass  = info.getRange(S.CLASS_CELL).getValue();
      const number = info.getRange(S.NUMBER_CELL).getValue();
      const name   = info.getRange(S.NAME_CELL).getValue();
      if (!klass && !number && !name) { skipped.push(file.getName() + '（氏名等が未入力）'); continue; }

      const row = new Array(CONFIG.TOTAL_COLS).fill('');
      row[CONFIG.COL.CLASS - 1]     = klass;
      row[CONFIG.COL.NUMBER - 1]    = number;
      row[CONFIG.COL.NAME - 1]      = name;
      row[CONFIG.COL.Q1 - 1]        = ans.getRange(S.Q1_CELL).getDisplayValue();
      row[CONFIG.COL.Q2 - 1]        = ans.getRange(S.Q2_CELL).getDisplayValue();
      row[CONFIG.COL.ESSAY - 1]     = ans.getRange(S.ESSAY_CELL).getDisplayValue();
      row[CONFIG.COL.SELF_EVAL - 1] = ans.getRange(S.SELF_EVAL_CELL).getDisplayValue();
      row[CONFIG.COL.DIFF - 1]      = ans.getRange(S.DIFF_CELL).getDisplayValue();
      row[CONFIG.COL.REFLECT - 1]   = ans.getRange(S.REFLECT_CELL).getDisplayValue();
      rows.push({ row: row, klass: klass, number: number, fileId: file.getId() });

    } catch (e) {
      skipped.push(file.getName() + '（' + e.message + '）');
    } finally {
      if (temp) { try { DriveApp.getFileById(temp.id).setTrashed(true); } catch (e2) {} }
      Utilities.sleep(CONFIG.SLEEP_COLLECT);
    }
  }

  if (!rows.length) {
    ui.alert('読み取れる提出物がありませんでした。\n' +
             (skipped.length ? 'スキップ：\n' + skipped.join('\n') : ''));
    return;
  }

  // クラス→出席番号の順に並べる
  rows.sort(function (a, b) {
    const ka = Number(a.klass) || 0, kb = Number(b.klass) || 0;
    if (ka !== kb) return ka - kb;
    return (Number(a.number) || 0) - (Number(b.number) || 0);
  });

  const n = Math.min(rows.length, CONFIG.DATA_END - CONFIG.DATA_START + 1);
  data.getRange(CONFIG.DATA_START, 1,
                CONFIG.DATA_END - CONFIG.DATA_START + 1, CONFIG.TOTAL_COLS).clearContent();
  data.getRange(CONFIG.DATA_START, 1, n, CONFIG.TOTAL_COLS)
      .setValues(rows.slice(0, n).map(function (r) { return r.row; }));

  map.getRange(2, 1, n, 3).setValues(
    rows.slice(0, n).map(function (r) { return [r.klass, r.number, r.fileId]; }));

  let msg = n + '人分の回答を集めました。';
  if (rows.length > n) msg += '\n※ ' + (rows.length - n) + '人分が枠（40人）を超えたため入りませんでした。';
  if (skipped.length)  msg += '\n\nスキップしたファイル：\n' + skipped.join('\n');
  ui.alert(msg);
}

function getFileMapSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.FILEMAP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.FILEMAP_SHEET);
    sh.hideSheet();
  }
  return sh;
}

/* ============================================================
 *  ステップ2：AI採点
 * ============================================================ */
function step2_grade() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(CONFIG.CFG_SHEET);
  const data = ss.getSheetByName(CONFIG.DATA_SHEET);

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) { ui.alert('先にメニューの「APIキーを登録」を実行してください。'); return; }

  const model = String(cfg.getRange(CONFIG.CFG_CELL.MODEL).getValue() || 'gemini-2.0-flash').trim();
  const temp  = Number(cfg.getRange(CONFIG.CFG_CELL.TEMP).getValue());
  const temperature = isNaN(temp) ? 0.1 : temp;

  const rowCount = CONFIG.DATA_END - CONFIG.DATA_START + 1;
  const values = data.getRange(CONFIG.DATA_START, 1, rowCount, CONFIG.TOTAL_COLS).getValues();

  let done = 0, skip = 0, fail = 0;
  const errors = [];

  for (let i = 0; i < rowCount; i++) {
    const r = CONFIG.DATA_START + i;
    const essay = String(values[i][CONFIG.COL.ESSAY - 1] || '').trim();
    const grade = String(values[i][CONFIG.COL.GRADE - 1] || '').trim();

    if (!essay) { continue; }
    if (grade)  { skip++; continue; }   // 採点済みはスキップ

    try {
      const result = callGemini_(apiKey, model, temperature, buildPrompt_(essay));
      data.getRange(r, CONFIG.COL.GRADE).setValue(result.grade);
      data.getRange(r, CONFIG.COL.COMMENT).setValue(result.comment);
      done++;
    } catch (e) {
      fail++;
      errors.push(r + '行目：' + e.message);
      data.getRange(r, CONFIG.COL.COMMENT).setValue('【採点エラー】' + e.message);
    }
    Utilities.sleep(CONFIG.SLEEP_GRADE);
  }

  let msg = '採点しました：' + done + '件\n採点済みでスキップ：' + skip + '件';
  if (fail) msg += '\nエラー：' + fail + '件\n' + errors.join('\n');
  ui.alert(msg);
}

function clearGrades() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('AI評価（G列）とAIコメント（H列）を消して再採点できるようにします。よろしいですか？',
                       ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DATA_SHEET);
  const n = CONFIG.DATA_END - CONFIG.DATA_START + 1;
  data.getRange(CONFIG.DATA_START, CONFIG.COL.GRADE, n, 2).clearContent();
  ui.alert('クリアしました。');
}

/* ------------------------------------------------------------
 *  採点プロンプト
 * ---------------------------------------------------------- */
function buildPrompt_(essayText) {
  return [
'あなたは中学校国語の教師です。',
'以下のルーブリックに基づいて、中学3年生が書いた批評文を評価してください。',
'',
'【教材】「素顔同盟」（すやまたけし）と「私」（三崎亜記）の2作品',
'　・「素顔同盟」…仮面着用が法制化された社会。語り手「僕」は制度を疑う側で、最後に川を上流へ歩きだす。',
'　・「私」…住民データが「私」を証明する社会。語り手は制度を運用する側の市役所職員で、',
'　　　　　　自分のデータが二重になり一方を消されても「どちらが消えようが、同じ『私』なのだ。',
'　　　　　　何の問題もない」と受け入れてしまう。',
'',
'【単元の目標】二つの作品を批判的に読み、それぞれの文章に表れているものの見方や考え方について考える。',
'',
'【場面設定】この批評文の読み手は、まだ「素顔同盟」も「私」も読んでいない隣のクラスの友達である。',
'　読んでいない相手に伝わるよう、本文からの引用が必要である。',
'',
'【課題】二つの作品は、「私が私であること」を何が支え、何が奪うと考えているか。',
'　二つの作品のものの見方・考え方を検討したうえで、自分の考えを書く。（800字）',
'',
'【必須条件】',
'　条件① 両作品からそれぞれ一か所以上引用し、その作品が「私が私であること」について',
'　　　　 何を支え・何が奪うと考えているかを示している',
'　条件② 少なくとも一方の作品の見方に対して、「しかし」「とはいえ」で始まる疑いの一文がある',
'　条件③ 話し合いで自分の考えが動いた発言に触れ、自分の生活や今の社会と結んでいる',
'',
'【ルーブリック】（思考・判断・表現／5段階）',
'S：条件①②を満たし、さらに協議で自分の考えが動いた地点を示して自分の生活・社会と結んでいる。',
'　　加えて、自分の評価の前提や限界にも触れている。',
'　　（例「私は最初、僕に共感していた。しかしそれは、僕が私たちと同じ側にいるように書かれていたからだ」）',
'A：条件①②の両方を満たしている。両作品から引用し、それぞれの見方のちがいを捉え、',
'　　少なくとも一方の見方に「しかし／とはいえ」で疑いを差しはさんでいる。',
'B：条件①②のどちらか一方のみ。両作品を引用しているが疑いがない、',
'　　または疑いはあるが引用が一方の作品にしかない。',
'C：両作品に触れているが、引用も疑いもなく、あらすじや感想が中心。',
'D：一方の作品にしか触れていない、または未完。',
'',
'評価のポイント：',
'1. 条件①（両作品からの引用があり、それが根拠になっているか）を満たしているか',
'2. 条件②（一方の作品の見方に「しかし／とはいえ」で疑いを差しはさんでいるか）を満たしているか',
'3. 条件③（協議で考えが動いた発言に触れ、自分の生活・社会と結んでいるか）に達しているか',
'4. まだ読んでいない友達に、二つの作品の中身が伝わる書き方になっているか',
'',
'※ 1と2の両方に触れていなければ、A以上にはなりません。',
'※ 「素顔同盟」の語り手「僕」に共感しているだけで、その見方を疑っていない答案は条件②を満たしません。',
'※ あらすじの要約が長く、自分の考えが薄い答案はC以下です。',
'',
'【生徒の批評文】',
essayText,
'',
'以下のJSON形式のみで回答してください。前後に説明文をつけないでください。',
'{"grade":"S","comment":"...（100字以内。よい点を1つと、次に伸ばす点を1つ。中学3年生にわかりやすい言葉で）"}'
  ].join('\n');
}

/* ------------------------------------------------------------
 *  Gemini API 呼び出し
 * ---------------------------------------------------------- */
function callGemini_(apiKey, model, temperature, prompt) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/'
            + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(apiKey);
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: temperature,
      responseMimeType: 'application/json',
    },
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) throw new Error('API応答 ' + code + '：' + body.slice(0, 200));

  let text;
  try {
    text = JSON.parse(body).candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error('応答の形式が想定外：' + body.slice(0, 200));
  }

  const parsed = parseJsonLoose_(text);
  const grade = String(parsed.grade || '').trim().toUpperCase();
  if (['S', 'A', 'B', 'C', 'D'].indexOf(grade) < 0) {
    throw new Error('評価が S/A/B/C/D 以外：' + text.slice(0, 120));
  }
  return { grade: grade, comment: String(parsed.comment || '').trim() };
}

function parseJsonLoose_(text) {
  try { return JSON.parse(text); } catch (e) {}
  const m = String(text).match(/\{[\s\S]*\}/);   // ```json などで囲まれた場合
  if (m) { try { return JSON.parse(m[0]); } catch (e2) {} }
  throw new Error('JSONとして読めない：' + String(text).slice(0, 120));
}

/* ============================================================
 *  ステップ3：生徒に返却
 * ============================================================ */
function step3_return() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(CONFIG.DATA_SHEET);
  const map = ss.getSheetByName(CONFIG.FILEMAP_SHEET);
  if (!map || map.getLastRow() < 2) {
    ui.alert('返却先の情報がありません。先に「① 生徒の回答を集める」を実行してください。');
    return;
  }

  const S = CONFIG.STUDENT;
  const mapRows = map.getRange(2, 1, map.getLastRow() - 1, 3).getValues();
  const idByKey = {};
  mapRows.forEach(function (r) { idByKey[String(r[0]) + '|' + String(r[1])] = r[2]; });

  const n = CONFIG.DATA_END - CONFIG.DATA_START + 1;
  const values = data.getRange(CONFIG.DATA_START, 1, n, CONFIG.TOTAL_COLS).getValues();

  let done = 0, fail = 0;
  const errors = [];

  for (let i = 0; i < n; i++) {
    const klass  = values[i][CONFIG.COL.CLASS - 1];
    const number = values[i][CONFIG.COL.NUMBER - 1];
    // 教師が最終評価を入れていればそれを、無ければAI評価を返す
    const grade   = String(values[i][CONFIG.COL.TEACHER - 1] || values[i][CONFIG.COL.GRADE - 1] || '').trim();
    const comment = String(values[i][CONFIG.COL.COMMENT - 1] || '').trim();
    if (!grade) continue;

    const fileId = idByKey[String(klass) + '|' + String(number)];
    if (!fileId) {
      fail++; errors.push(klass + '組' + number + '番：提出ファイルが見つかりません');
      continue;
    }

    try {
      const wb = SpreadsheetApp.openById(fileId);
      const ans = wb.getSheetByName(S.ANSWER_SHEET);
      if (!ans) throw new Error('「' + S.ANSWER_SHEET + '」シートがありません');
      ans.getRange(S.AI_EVAL_CELL).setValue(grade);
      ans.getRange(S.AI_COMMENT_CELL).setValue(comment);
      SpreadsheetApp.flush();
      done++;
    } catch (e) {
      fail++;
      errors.push(klass + '組' + number + '番：' + e.message);
    }
    Utilities.sleep(CONFIG.SLEEP_COLLECT);
  }

  let msg = done + '人に評価を返しました。';
  if (fail) msg += '\n返せなかった：' + fail + '件\n' + errors.join('\n');
  ui.alert(msg);
}
