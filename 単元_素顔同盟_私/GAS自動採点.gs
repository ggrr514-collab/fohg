/**
 * 中学3年 国語　単元「『私』を消すもの、『私』を支えるもの」
 * ――「素顔同盟」（すやまたけし）／「私」（三崎亜記）を批判的に読む
 *
 * 単元のまとめは【グループディスカッション】。
 * 話し合いそのものは教員用「協議見取り」シートで教師が評価する。
 * このスクリプトは、話し合いを終えて生徒が書いた300字のAI一次採点を行う。
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
    ESSAY_CELL:      'A25',   // ディスカッションを終えて（300字）★AI採点対象
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
    .addItem('② 300字をAIで採点する', 'step2_grade')
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

  const rows = [];
  const skipped = [];
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    let sheetId = null;
    let temp = null;
    try {
      const mime = file.getMimeType();
      let isSheet = false;
      if (mime === MimeType.GOOGLE_SHEETS) {
        sheetId = file.getId();
        isSheet = true;
      } else if (mime === MimeType.MICROSOFT_EXCEL ||
                 mime === MimeType.MICROSOFT_EXCEL_LEGACY) {
        // xlsxのままでは読めないので、一時的にスプレッドシートへ変換して読む。
        // ※ サービス「Drive API（v2）」を有効にしておくこと（エディタ左の「サービス」から追加）
        //    v2の記法なのでキーは title。v3を追加した場合は name に読み替える。
        // ※ 読み取りはできるが【返却（ステップ3）はできない】。返却まで行うなら
        //    Googleスプレッドシート形式で配付・提出させること。
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
      if (!klass || !number) {
        skipped.push(file.getName() + '（組または出席番号が未入力。返却先を特定できません）');
        continue;
      }

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
      rows.push({ row: row, klass: klass, number: number,
                  fileId: file.getId(), isSheet: isSheet });

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

  const rowCount = CONFIG.DATA_END - CONFIG.DATA_START + 1;
  const n = Math.min(rows.length, rowCount);

  // 生徒は返却を受け取ってから自己評価・評価差考察・ふりかえりを書くので、
  // 教師は「収集→採点→返却→再収集」を行うことになる。その再収集で
  // AI評価(G)・AIコメント(H)・教師最終評価(I) を消さないよう、クラス|番号 をキーに退避する。
  const keep = {};
  const prev = data.getRange(CONFIG.DATA_START, 1, rowCount, CONFIG.TOTAL_COLS).getValues();
  prev.forEach(function (r) {
    const key = String(r[CONFIG.COL.CLASS - 1]) + '|' + String(r[CONFIG.COL.NUMBER - 1]);
    if (key === '|') return;
    const g = r[CONFIG.COL.GRADE - 1], h = r[CONFIG.COL.COMMENT - 1], t = r[CONFIG.COL.TEACHER - 1];
    if (g || h || t) keep[key] = [g, h, t];
  });

  data.getRange(CONFIG.DATA_START, 1, rowCount, CONFIG.TOTAL_COLS).clearContent();
  data.getRange(CONFIG.DATA_START, 1, n, CONFIG.TOTAL_COLS)
      .setValues(rows.slice(0, n).map(function (r) { return r.row; }));

  // 退避しておいた G/H/I を、行の順番ではなく クラス|番号 で書き戻す
  let restored = 0;
  rows.slice(0, n).forEach(function (r, idx) {
    const key = String(r.klass) + '|' + String(r.number);
    if (keep[key]) {
      data.getRange(CONFIG.DATA_START + idx, CONFIG.COL.GRADE, 1, 3).setValues([keep[key]]);
      restored++;
    }
  });

  map.clear();
  map.getRange(1, 1, 1, 4).setValues([['クラス', '出席番号', 'ファイルID', '形式']]);
  map.getRange(2, 1, n, 4).setValues(
    rows.slice(0, n).map(function (r) {
      return [r.klass, r.number, r.fileId, r.isSheet ? 'スプレッドシート' : 'xlsx'];
    }));

  let msg = n + '人分の回答を集めました。';
  if (restored) msg += '\n※ ' + restored + '人分の採点結果・教師評価を保持しました。';
  if (rows.length > n) msg += '\n※ ' + (rows.length - n) + '人分が枠（40人）を超えたため入りませんでした。';
  const xlsxCount = rows.slice(0, n).filter(function (r) { return !r.isSheet; }).length;
  if (xlsxCount) {
    msg += '\n※ ' + xlsxCount + '人分がxlsx形式です。採点はできますが【ステップ3の返却ができません】。'
         + 'Googleスプレッドシート形式で再提出させてください。';
  }
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
  if (!cfg || !data) { ui.alert('「AI設定」または「回答入力」シートが見つかりません。'); return; }

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
'以下のルーブリックに基づいて、中学3年生が書いた文章を評価してください。',
'',
'【この単元のまとめは「グループディスカッション」である】',
'　話し合いそのものは教師が観察して評価する。あなたが評価するのは、',
'　話し合いを終えて生徒が書いた300字の記述である。',
'　したがって「もっと詳しく書くべき」などの分量の指摘より、',
'　話し合いで自分の読みがどう動いたかが書けているかを見ること。',
'',
'【教材】「素顔同盟」（すやまたけし）と「私」（三崎亜記）の2作品',
'　・「素顔同盟」…仮面着用が法制化された社会。語り手「僕」は制度を疑う側で、最後に川を上流へ歩きだす。',
'　・「私」…住民データが「私」を証明する社会。語り手は制度を運用する側の市役所職員で、',
'　　　　　　自分のデータが二重になり一方を消されても「どちらが消えようが、同じ『私』なのだ。',
'　　　　　　何の問題もない」と受け入れてしまう。',
'',
'【話し合いの論題（3つ）】',
'　論題1 仮面を捨てた先に、自由はあるか（素顔同盟）',
'　論題2 「何の問題もない」と言う語り手は、何に気づいていないか（私）',
'　論題3 「私」が「私」であることを支えるのは何か（両作品）',
'',
'【単元の目標】自分で決めた観点から二つの作品を批判的に読み、文章に表れたものの見方や',
'　考え方について、根拠を示してグループで話し合う。',
'',
'【重要】読みの観点は生徒が自分で決めている。次の9観点から選ぶ決まりで、必修は「批評・評価」。',
'　登場人物／設定／構成／視点・語り／表現の工夫／題名／主題／批評・評価／自分に生かす',
'　★どの観点を選んだかで評価を変えないこと。「別の観点にすべき」という指摘もしないこと。',
'　　評価するのは、その観点から本文の根拠を示して吟味できているかである。',
'',
'【300字に入れるべき条件】',
'　条件① 自分の観点と、本文の根拠を示して主張したことが書けている',
'　条件② 他の観点からの指摘を受けて、自分の考えをどう見直したかが書けている',
'　　　　（「しかし」「とはいえ」などで、自分の読みに疑いを差しはさんでいる）',
'　条件③ 自分の観点では見えなかったことに触れ、話し合いを経た今の考えを述べている',
'',
'【ルーブリック】（思考・判断・表現／5段階）',
'S：条件①②に加え、条件③（自分の観点では見えなかったことに触れ、今の考えを述べている）まで達している。',
'　　（例「私は『表現の工夫』という観点で読んだので、笑顔の描かれ方ばかり見ていた。班の人の',
'　　　『題名』という観点を聞いて、選べるかどうかを見ていなかったと気づいた」）',
'A：条件①②の両方。自分の観点と本文の根拠で主張し、他の観点を受けて考えを見直している。',
'B：条件①②の一方だけ。主張はあるが見直しがない、または見直しはあるが観点と根拠が示されていない。',
'C：話し合いの感想だけ（「楽しかった」「いろいろな意見が出た」）。主張も見直しもない。',
'D：記述が乏しい、または一方の作品にしか触れていない。',
'',
'評価のポイント：',
'1. 条件①（自分の観点が示され、本文の根拠とともに主張が書けているか）',
'2. 条件②（他の人の観点を受けて、自分の考えが動いたことが具体的に書けているか。',
'　 だれの・どの観点からの指摘かが分かるとよい）',
'3. 条件③（自分の観点の限界に触れ、今の考えを述べているか）に達しているか',
'',
'※ 1と2の両方がなければ、A以上にはなりません。',
'※ 「意見が変わった」とだけ書いてあり、何をどう見直したのかが無い場合は条件②を未達とする。',
'※ 観点の種類そのものを評価の材料にしないこと。観点の名前を書いただけで本文と',
'　 結びついていなければ条件①は未達とする。',
'',
'【生徒の記述（ディスカッションを終えて・300字）】',
essayText,
'',
'以下のJSON形式のみで回答してください。前後に説明文をつけないでください。',
'{"grade":"S","comment":"...（100字以内。よい点を1つと、次に伸ばす点を1つ。中学3年生にわかりやすい言葉で。',
'　生徒が選んだ観点を尊重し、別の観点をすすめる書き方はしない）"}'
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
  const mapRows = map.getRange(2, 1, map.getLastRow() - 1, 4).getValues();
  const idByKey = {};
  mapRows.forEach(function (r) {
    idByKey[String(r[0]) + '|' + String(r[1])] = { id: r[2], isSheet: (r[3] === 'スプレッドシート') };
  });

  const n = CONFIG.DATA_END - CONFIG.DATA_START + 1;
  const values = data.getRange(CONFIG.DATA_START, 1, n, CONFIG.TOTAL_COLS).getValues();

  let done = 0, fail = 0;
  const errors = [];

  for (let i = 0; i < n; i++) {
    const klass  = values[i][CONFIG.COL.CLASS - 1];
    const number = values[i][CONFIG.COL.NUMBER - 1];
    // 教師が最終評価を入れていればそれを、無ければAI評価を返す
    const aiGrade      = String(values[i][CONFIG.COL.GRADE - 1] || '').trim();
    const teacherGrade = String(values[i][CONFIG.COL.TEACHER - 1] || '').trim();
    const grade = teacherGrade || aiGrade;
    if (!grade) continue;

    // 採点エラーの文言は生徒に返さない
    let comment = String(values[i][CONFIG.COL.COMMENT - 1] || '').trim();
    if (comment.indexOf('【採点エラー】') === 0) comment = '';
    // 先生が評価を上書きした場合、AIコメントと食い違って見えないように注記する
    if (teacherGrade && aiGrade && teacherGrade !== aiGrade) {
      comment = '（評価は先生が確定しています。以下はAIのコメントです）\n' + comment;
    }

    const entry = idByKey[String(klass) + '|' + String(number)];
    if (!entry) {
      fail++; errors.push(klass + '組' + number + '番：提出ファイルが見つかりません');
      continue;
    }
    if (!entry.isSheet) {
      fail++;
      errors.push(klass + '組' + number + '番：xlsx形式のため返却できません。'
                + 'Googleスプレッドシート形式で再提出させてください');
      continue;
    }

    try {
      const wb = SpreadsheetApp.openById(entry.id);
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
