const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  Footer, PageNumber, PageOrientation, VerticalAlign,
} = require('docx');
const fs = require('fs');

const MIN = 'Yu Mincho', GOT = 'Yu Gothic';
const NAVY = '1F3864', CRIM = 'C00000', GRAYT = '595959';
const W = 9638;           // A4縦 本文幅（余白2cm）
const WL = 14570;         // A4横 本文幅

const f = (n) => ({ name: n, eastAsia: n });

// **太字** を解釈して TextRun 配列を返す
function runs(text, o = {}) {
  const base = { size: o.size || 21, font: f(o.font || MIN), color: o.color };
  const out = [];
  text.split(/(\*\*[^*]+\*\*)/).forEach((seg) => {
    if (!seg) return;
    const b = seg.startsWith('**') && seg.endsWith('**');
    out.push(new TextRun(Object.assign({}, base, {
      text: b ? seg.slice(2, -2) : seg,
      bold: b || !!o.bold,
    })));
  });
  return out;
}

const p = (text, o = {}) => new Paragraph({
  children: runs(text, o),
  spacing: { before: o.before === undefined ? 40 : o.before,
             after: o.after === undefined ? 80 : o.after, line: o.line || 300 },
  alignment: o.align,
  indent: o.indent,
});

const h1 = (text) => new Paragraph({
  children: runs(text, { size: 26, font: GOT, bold: true, color: NAVY }),
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 320, after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: NAVY, space: 4 } },
});

const h2 = (text) => new Paragraph({
  children: runs(text, { size: 23, font: GOT, bold: true, color: NAVY }),
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 100 },
});

// 発話ブロック（左罫線＋淡い地色）。pos: 'first' | 'mid' | 'last' | 'only'
function say(speaker, text, pos = 'only') {
  const kids = [];
  if (speaker) {
    kids.push(new TextRun({ text: speaker, bold: true, size: 21,
                            font: f(GOT), color: NAVY }));
    kids.push(new TextRun({ text: '　', size: 21, font: f(MIN) }));
  }
  kids.push(...runs(text));
  return new Paragraph({
    children: kids,
    spacing: {
      before: (pos === 'first' || pos === 'only') ? 140 : 0,
      after:  (pos === 'last'  || pos === 'only') ? 140 : 0,
      line: 300,
    },
    indent: { left: 400, right: 200 },
    border: { left: { style: BorderStyle.SINGLE, size: 20, color: NAVY, space: 10 } },
    shading: { type: ShadingType.CLEAR, fill: 'F4F6FA', color: 'auto' },
  });
}

// 注記ブロック（薄いグレー地・細い左罫線）
function note(lines) {
  return lines.map((t, i) => new Paragraph({
    children: runs(t, { size: 19, color: GRAYT }),
    spacing: { before: i === 0 ? 120 : 0, after: i === lines.length - 1 ? 120 : 0, line: 280 },
    indent: { left: 340, right: 200 },
    border: { left: { style: BorderStyle.SINGLE, size: 10, color: 'B4B4B4', space: 8 } },
    shading: { type: ShadingType.CLEAR, fill: 'F7F7F5', color: 'auto' },
  }));
}

// 生徒の記述ブロック（枠囲み）
function box(lines) {
  return lines.map((t, i) => new Paragraph({
    children: runs(t, { size: 21 }),
    spacing: { before: i === 0 ? 140 : 0, after: i === lines.length - 1 ? 140 : 0, line: 320 },
    indent: { left: 280, right: 280 },
    border: {
      left:   { style: BorderStyle.SINGLE, size: 8, color: '9A7B3F', space: 10 },
      right:  { style: BorderStyle.SINGLE, size: 8, color: '9A7B3F', space: 10 },
      top:    i === 0 ? { style: BorderStyle.SINGLE, size: 8, color: '9A7B3F', space: 6 } : undefined,
      bottom: i === lines.length - 1 ? { style: BorderStyle.SINGLE, size: 8, color: '9A7B3F', space: 6 } : undefined,
    },
    shading: { type: ShadingType.CLEAR, fill: 'FFFDF5', color: 'auto' },
  }));
}

function table(widths, head, rows, o = {}) {
  const sz = o.size || 18;
  const cell = (txt, wd, isHead, align) => new TableCell({
    width: { size: wd, type: WidthType.DXA },
    shading: isHead ? { type: ShadingType.CLEAR, fill: NAVY, color: 'auto' }
                    : (o.zebra ? { type: ShadingType.CLEAR, fill: 'FFFFFF', color: 'auto' } : undefined),
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: [new Paragraph({
      children: runs(txt, { size: sz, font: isHead ? GOT : MIN,
                            bold: isHead, color: isHead ? 'FFFFFF' : undefined }),
      alignment: align || (isHead ? AlignmentType.CENTER : AlignmentType.LEFT),
      spacing: { before: 0, after: 0, line: 270 },
    })],
  });
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: {
      top:    { style: BorderStyle.SINGLE, size: 6, color: '8C8C8C' },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: '8C8C8C' },
      left:   { style: BorderStyle.SINGLE, size: 6, color: '8C8C8C' },
      right:  { style: BorderStyle.SINGLE, size: 6, color: '8C8C8C' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
      insideVertical:   { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
    },
    rows: [
      new TableRow({ tableHeader: true,
        children: head.map((t, i) => cell(t, widths[i], true)) }),
      ...rows.map((r) => new TableRow({
        children: r.map((t, i) => cell(t, widths[i], false,
          (o.center || []).includes(i) ? AlignmentType.CENTER : undefined)) })),
    ],
  });
}

const footer = new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: ['— ', PageNumber.CURRENT, ' —'],
                             size: 18, font: f(MIN), color: GRAYT })],
  })],
});

// ========================= 本文 =========================
const body = [];

body.push(new Paragraph({
  children: runs('授業イメージ　―― 6人班の10分を、そのまま再現する',
                 { size: 30, font: GOT, bold: true, color: NAVY }),
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 60 },
}));
body.push(new Paragraph({
  children: runs('中学3年 国語「素顔同盟」・「私」　批判的に読む単元（全7時間）',
                 { size: 19, color: GRAYT }),
  alignment: AlignmentType.CENTER,
  spacing: { after: 220 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 8 } },
}));

body.push(p('「観点がちがうと、同じ論題がどう変わるか」を具体で示す資料。第4時の指導や、研究協議の説明にそのまま使える。'));
body.push(...note([
  '引用のページ・行は教科書に合わせて差し替えてください。',
  '論題は3つとも**二つの作品にまたがる**問いです。',
  '　論題1　仮面とデータは、同じものか',
  '　論題2　訴えた人は、救われたか',
  '　論題3　この社会を続けさせているのは、誰か',
  '以下は**論題1**のラウンドを再現したものです。',
]));

// --- 1 ---
body.push(h1('1.　6人の観点（第2時に決めたもの）'));
body.push(p('全員「批評・評価」が必修。それに加えて、**残り8つの中から各自が1つ選んだ**。班の中で観点が重ならないように組んである。'));
body.push(table([620, 1180, 2600, 5238],
  ['', '観点', '選んだ理由', 'その観点で二作から見えたこと'],
  [
    ['**A**', '**表現の工夫**', '「笑顔」という言葉の出方が気になった。仮面がなぜ「笑顔」の形なのか',
     '【素顔同盟】「その無個性な笑顔はみんなと同じなのだ」。先生も「仮面に笑顔を浮かべ、熱弁をふるっている」。**笑顔が中身と切り離されている**／【私】「解決」「誠意」「模範的」がカギカッコ付き。**言葉だけが正しい**'],
    ['**B**', '視点・語り', '「僕」の言うことをそのまま信じていいのか気になった',
     '全部「僕」の目。「彼らは仮面の下で、どんな顔をしているのだろう」と言うだけで**確かめていない**。友人は「笑顔のおかげで、僕たちはけんかをしないですんでいるんだろ」と答えている'],
    ['**C**', '設定', 'なぜ川で隔てられているのか気になった',
     '「この橋のない川を隔てて」「自然保護区は荒らされてはならない聖域だった」。川岸は「コンクリートで固められ」、イチョウは「等間隔に並んでいて」。**分け方がきれいすぎる**'],
    ['**D**', '題名', '「素顔同盟」という名前が変だと思った',
     '仮面を外した人の集まりなのに「同盟」。同盟は仲間の約束。**約束があるなら、そこにもルールがある**'],
    ['**E**', '構成', '最後の一行で急に動き出すのが気になった',
     '朝→授業→帰り道→夜→次の日→数週間。**ずっと僕は見ているだけ**。最後だけ「ためらいもなく、その川を上流に向かって歩きだした」'],
    ['**F**', '自分に生かす', '自分も本当の気持ちを言わないことがあるから',
     '僕は「先生の今の話、おかしいと思わない？」と聞いたのに、注意されたらしょんぼりして一人で帰る。彼女に声もかけられない。「結局、勇気がなかったのだ」'],
  ], { center: [0] }));
body.push(p('**同じ二作を読んでいるのに、目に入っているものが6通りちがう。**これが議論の燃料になる。', { before: 160 }));
body.push(...note(['観点は二作の**両方**に向ける。3つの論題がどちらも二作にまたがるので、片方だけ見ていると議論に入れない。']));

// --- 2 ---
body.push(h1('2.　ディスカッション例（第5時・ラウンド1／10分）'));
body.push(p('**論題1「仮面とデータは、同じものか」**'));
body.push(p('提案者＝A（表現の工夫）　／　引用チェック係＝C　／　疑い係＝D　／　記録係＝E', { size: 19, color: GRAYT }));
body.push(...note(['**この時間のはじめに、Aさんは自分が論題1の提案者だと知らされた。**3つとも準備してきている。']));

body.push(h2('0:00〜2:00　提案'));
body.push(say('A', '「私は『表現の工夫』の観点です。主張は、**仮面とデータは同じではない**です。', 'first'));
body.push(say(null, '　根拠は二つ。『素顔同盟』の仮面は『その無個性な笑顔はみんなと同じなのだ』と書かれていて、**言葉で形が説明されている**。つまり外から見える。だから彼女は『両手で仮面を覆うと、そっとそれを外した』ことができました。', 'mid'));
body.push(say(null, '　『私』のほうは、大事な言葉が全部カギカッコに入っています。『解決』『誠意』『模範的』『正常な状態』。**言葉だけが正しくて中身が空っぽ**という書き方です。データも同じで、見えないし触れない。消すのも山中さんという他人の手です。', 'mid'));
body.push(say(null, '　自分で外せるものと、自分では触れないもの。だから同じではありません。」', 'last'));

body.push(h2('2:00〜8:00　検討'));
body.push(say('C（引用チェック係）', '「『自分では触れない』って、本文のどこ？」', 'first'));
body.push(say('A', '「『すまないが、そちらのデータを復元して、今のデータを削除してもらえないだろうか』——語り手が自分では操作せず、山中さんに頼んでいます。」', 'mid'));
body.push(say('C', '「あ、ちゃんとある。私は『設定』の観点だけど、そこは近い。『素顔同盟』は『この橋のない川を隔てて』と、境界がはっきり見える。『私』は市役所と図書館が地続きで、境界が見えない。**見える壁と見えない壁**だと思う。」', 'last'));

body.push(say('D（疑い係・題名）', '「**しかし**、私は同じだと思う。『素顔同盟』の仮面は『法令化され、制度として確立された』と本文にある。『私』のデータも役所の制度。**どちらも制度が動かしている**点は同じだよね。外せるかどうかは、制度の強さの違いだけじゃない？」', 'first'));
body.push(say('A', '「……でも、外せたかどうかは結果として大きい気が。」', 'mid'));
body.push(say('D', '「彼女は外せたけど、そのあと**どうなったか本文に書いてある？**」', 'mid'));
body.push(say('A', '「……書いてないです。仮面が川に浮いていただけ。」', 'mid'));
body.push(say('D', '「隔離されたかもしれないよね。だったら『外せる』も見かけだけかもしれない。」', 'last'));

body.push(say('B（視点・語り）', '「私はそこで気になることがあって。『私』は語り手が**自分では気づいていない**でしょ。『どちらが消えようが、同じ「私」なのだ。何の問題もない』で終わる。『素顔同盟』の僕は気づいている。**同じ制度でも、当人が気づいているかどうかが違う。**だから私は、仮面とデータは同じでも、**人間の側が違う**と思う。」', 'first'));
body.push(say('F（自分に生かす）', '「私も二つを比べて思ったんだけど、仮面って学校の『とりあえず笑っとく』に近い。データは成績とか出席番号に近い。**自分で外せるのは前者だけ**っていうAさんの区別は、自分のことで考えると分かる。」', 'mid'));
body.push(say('E（記録係・構成）', '「構成で見ると、『素顔同盟』は最後に動いて終わる、『私』は最後に止まって終わる。**同じものを描いても、終わらせ方が逆。**」', 'last'));

body.push(h2('8:00〜10:00　提案者の応答と到達点'));
body.push(say('A', '「……私は『表現の工夫』で、**書かれ方のちがい**ばかり見ていました。', 'first'));
body.push(say(null, '　でもDさんの『どちらも制度が動かしている』と、『外せたあと本文に書いてない』で変わりました。私は『外せた』を根拠にしていたけど、**外したあとどうなったかは本文にない**。', 'mid'));
body.push(say(null, '　だから主張を直します。**『制度としては同じ。ちがうのは、外せるように見えるかどうかだけ』**です。」', 'last'));
body.push(say('到達点（記録係E）', '「仮面とデータは制度としては同じ。ちがうのは**当人に見えるかどうか**と、**当人が気づいているかどうか**。外せるかどうかは本文が答えていない。」', 'only'));

// --- 3 ---
body.push(h1('3.　この10分で、何が起きたか'));
body.push(table([2400, 7238], ['仕掛け', '実際に効いたところ'], [
  ['**論題が二作にまたがる**', 'Aは冒頭から両作を引いた。**片方だけでは主張が立たない形**になっている'],
  ['**引用チェック係**', 'Cの「本文のどこ？」にAは即答できた。準備してきた効果'],
  ['**疑い係**', 'Dが「外せたあとは本文に書いてあるか」と突いて、Aの根拠が崩れた'],
  ['**異観点編成**', 'B（語り）は「人間の側の違い」、E（構成）は「終わらせ方が逆」を出した'],
  ['**当日に論題を知る**', 'Aは3つ準備していたので、指名されても2分で話せた'],
]));
body.push(p('Aは「表現の工夫」、Dは「題名」、Fは「自分に生かす」——**どれも9観点の中のありふれた選択肢**である。それでも6人の目に入るものはこれだけ違う。**どの観点を選んだかで評価は変えない**——この単元でいちばん守ってほしい原則。', { before: 160 }));

// --- 4 ---
body.push(h1('4.　第6時（2周目）で同じ論題1に戻ると'));
body.push(p('提案者はD（題名）に交代。**1周目の到達点から始まる**ので、議論の位置が上がる。'));
body.push(say('D', '「1周目で『制度としては同じ。ちがうのは見えるかどうか』まで出ました。私は題名の観点から、その先を問います。', 'first'));
body.push(say(null, '　『素顔同盟』は**仮面を捨てた人の集まりに「同盟」という名前がついている**。『私』のほうは題名が「私」で、**かぎかっこが無い**。本文では何度も『私』とカギカッコ付きで書かれるのに、題名だけ素の「私」なんです。', 'mid'));
body.push(say(null, '　つまり作者は、**制度に名前をつけることと、名前を奪うこと**の両方を書いている。仮面とデータのちがいは、**名前が残るかどうか**ではないか。」', 'last'));
body.push(p('1周目は「同じか、ちがうか」、2周目は「**では何がちがうのか**」。同じ論題でも問いの階層が上がる。これが2周する意味。'));
body.push(...note(['論題2・3のラウンドも同じ形で進む。論題2では「女性は満足して帰った。あれは救いか」が、論題3では「山中は気づいているのに止めなかった」が争点になりやすい。']));

// --- 5 ---
body.push(h1('5.　Aさんの「見直し」200字（第7時）'));
body.push(...box([
  '私は「表現の工夫」の観点で、仮面とデータは同じではないと主張した。仮面は「その無個性な笑顔は」と形が書かれ外から見えるが、データは「解決」「誠意」のようにカギカッコの中にあって見えないからだ。',
  '　**しかし**、論題1でDさんが「題名」の観点から「外せたあとどうなったかは本文に書いてない」と言った。**私は書かれ方のちがいばかり見ていて、書かれていないことを見ていなかった。**',
  '　いま私は、二つは制度としては同じで、ちがうのは当人に見えるかどうかだけだと考えている。',
]));
body.push(p('**評価：S**（ただし当日の発言が○であることが条件）', { before: 160 }));
body.push(table([2700, 3700, 3238], ['条件', '該当箇所', '見るところ'], [
  ['① 3論題を観点と根拠でまとめた', '「考えをまとめる」シート D〜F列', 'AIが一次採点'],
  ['② 当日、観点と根拠で発言した', '冒頭2分で両作から引用。Cの確認にも即答', '**教師が「協議見取り」H列に○**'],
  ['③ 見えなかったこと＋今の考え', '「書かれていないことを見ていなかった」「いま私は〜」', 'AIが一次採点'],
]));
body.push(h2('対照：同じ話し合いに出ていたBさんの200字（B評価）'));
body.push(...box([
  '話し合いで、Aさんが仮面とデータは同じではないと言った。Dさんは制度としては同じだと言った。**しかし**私はどちらとも言えないと思った。本文には書いていないことが多いからだ。いろいろな見方があると分かった。',
]));
body.push(p('→　「しかし」はあるが、**自分の観点と本文の根拠で主張した部分がない**。さらに記述側も論題3が空欄だった。**B止まり。**'));
body.push(...note([
  '**指導のポイント**：Bさんは話し合いでは「人間の側が違う」という鋭い発言をしている（H列＝○）。**発言は条件②を満たしているのに、条件①の記述が足りない。**',
  '逆のパターンもある——記述は3つそろっているのに当日は「私も同じです」で終わる生徒。**どちらもB。**この構造を第4時に生徒へ説明しておくと、準備と発言の両方に力が入る。',
]));

// ========== 横向きセクション：協議見取り ==========
const land = [];
land.push(h1('6.　教師の「協議見取り」記入例'));
land.push(table([770, 500, 1400, 900, 4200, 700, 5400, 700],
  ['氏名', '班', '選んだ観点', '発表論題', '当日の発言の要点', '条件②', '見えなかったこと／述べ直した箇所', '態度'],
  [
    ['A', '3', '表現の工夫', '1', '両作から引用。仮面は見える／データは見えない', '**○**',
     '書かれていないことを見ていなかった／「制度としては同じ」に修正', '**A**'],
    ['B', '3', '視点・語り', '2', '「女性は納得したが、督促状は一字一句変わらない」', '**○**',
     '記述の論題3が空欄', '**B**'],
    ['D', '3', '題名', '1(2周目)', '「名前が残るかどうか」。1周目のAの根拠を突いた', '**○**',
     '疑い係として議論を動かした', '**A**'],
    ['F', '3', '自分に生かす', '3', '「山中は気づいているのに止めていない」', '**○**',
     '「みんなのせい」から本文の人物に戻した', '**A**'],
  ], { center: [0, 1, 3, 5, 7] }));
land.push(p('**発話は残らないので、この表が条件②と態度の唯一の証拠になります。**条件②の列が空だと、記述がどれだけ良くても**A評価をつける根拠がありません**（条件①だけではB止まり）。', { before: 180 }));
land.push(p('全発言を聞き取る必要はありません。見るのは**提案者の2分**だけで足ります——「自分の観点を言ったか」「本文を引いたか」の2点です。提案者以外の生徒は、その人が提案者になるラウンドで見ればよい（6ラウンドで全員が1回提案者になります）。'));
land.push(...note([
  'Bさんの行が示すとおり、**発言が○でも記述が足りなければB**です。逆に記述が3つそろっていても発言が△ならBです。',
  'AIは条件②の列を見ていないので、**最終評価は必ず「回答入力」J列で確定**してください。',
]));

const doc = new Document({
  styles: { default: { document: { run: { font: f(MIN), size: 21 } } } },
  sections: [
    {
      properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
      footers: { default: footer },
      children: body,
    },
    {
      properties: {
        page: {
          size: { orientation: PageOrientation.LANDSCAPE },
          margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 },
        },
      },
      footers: { default: footer },
      children: land,
    },
  ],
});

Packer.toBuffer(doc).then((b) => {
  fs.writeFileSync('/home/user/fohg/単元_素顔同盟_私/授業イメージ_ディスカッション例.docx', b);
  console.log('written', b.length, 'bytes');
});
