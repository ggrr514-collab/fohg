# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "/home/user/fohg/単元_素顔同盟_私/生徒用ワークシート.xlsx"

NAVY=PatternFill('solid',fgColor='1F3864'); BLUE=PatternFill('solid',fgColor='D6E4F0')
LBLUE=PatternFill('solid',fgColor='E8F0FE'); ORANGE_F=PatternFill('solid',fgColor='FFF2CC')
GREEN_F=PatternFill('solid',fgColor='D5F5E3'); GRAY=PatternFill('solid',fgColor='F2F2F2')
INPUT=PatternFill('solid',fgColor='FFFFF0'); AI_TIP=PatternFill('solid',fgColor='EBF5FB')
PINK=PatternFill('solid',fgColor='FADBD8'); GREYH=PatternFill('solid',fgColor='808080')
GREYL=PatternFill('solid',fgColor='EFEFEF')

wrap=Alignment(wrap_text=True,vertical='top'); wrapc=Alignment(wrap_text=True,vertical='center')
center=Alignment(horizontal='center',vertical='center',wrap_text=True)
thin=Side(style='thin',color='BFBFBF'); box=Border(left=thin,right=thin,top=thin,bottom=thin)

F_TITLE=Font(size=15,bold=True,color='FFFFFF'); F_HEAD=Font(size=10,bold=True,color='FFFFFF')
F_SEC=Font(size=11,bold=True,color='1F3864'); F_LBL=Font(size=10,bold=True,color='333333')
F_BODY=Font(size=10); F_SMALL=Font(size=9,color='555555')
F_RED=Font(size=10,bold=True,color='C00000'); F_TIP=Font(size=9,color='1A5276',italic=True)
F_GOAL=Font(size=11,bold=True,color='1F3864'); F_Q=Font(size=11,bold=True,color='1F3864')

GOAL = ("自分で決めた観点から二つの作品を批判的に読みながら、三つの論題について、"
        "文章に表れているものの見方や考え方について考えることができる。")
CRIT_S = ("「読むこと」において、文章を批判的に読みながら、"
          "文章に表れているものの見方や考え方について考えている。")
CRIT_A = ("積極的に文章を批判的に読み、学習課題に沿って、"
          "考えたことを互いに伝え合おうとしている。")
CRIT_K = ("語句の意味や働きに注意して文章を読むことを通して、"
          "語感を磨き語彙を豊かにしている。")

TOPICS = [
    ('論題1', '仮面とデータは、同じものか', '「私」を消しているもの'),
    ('論題2', '訴えた人は、救われたか', '仮面を外した彼女／データを消された女性'),
    ('論題3', 'この社会を続けさせているのは、誰か', '友人（素顔同盟）／山中（私）'),
]

def mg(ws,r1,c1,r2,c2,val,font=None,fill=None,align=None,border=False,h=None):
    ws.cell(row=r1,column=c1).value=val
    if r1!=r2 or c1!=c2: ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
    for r in range(r1,r2+1):
        for c in range(c1,c2+1):
            cell=ws.cell(row=r,column=c)
            if font: cell.font=font
            if fill: cell.fill=fill
            if align: cell.alignment=align
            if border: cell.border=box
    if h: ws.row_dimensions[r1].height=h
    return ws.cell(row=r1,column=c1)

def tip(ws,r,text,lc=4):
    mg(ws,r,1,r,lc,'🤖 '+text,F_TIP,AI_TIP,wrapc,h=26); return r+1
def sec(ws,r,text,fill=BLUE,lc=4):
    mg(ws,r,1,r,lc,text,F_SEC,fill,wrapc,h=20); return r+1
def rows2(ws,r,pairs,lc=4,fill=GRAY,h=22,split=2):
    for a,b in pairs:
        mg(ws,r,1,r,split,a,F_LBL,fill,wrapc,border=True)
        mg(ws,r,split+1,r,lc,b,F_BODY,None,wrapc,border=True,h=h); r+=1
    return r
def counter(ws,r,col,n):
    c=ws.cell(row=r,column=col)
    c.value='=IF(A%d="","0/%d字",LEN(A%d)&"/%d字")'%(r,n,r,n)
    c.font=F_SMALL; c.alignment=center; c.border=box

wb=openpyxl.Workbook()

# ===================== Sheet1 単元ガイド =====================
ws=wb.active; ws.title='単元ガイド'
for col,w in zip('ABCDEF',[15,13,13,20,15,15]): ws.column_dimensions[col].width=w
mg(ws,1,1,1,6,'「私」を消すもの、「私」を支えるもの',F_TITLE,NAVY,center,h=30)
c=ws.cell(row=2,column=1); c.number_format='0"組"'; c.fill=LBLUE
c.font=Font(size=12,bold=True); c.alignment=center; c.border=box
c=ws.cell(row=2,column=2); c.number_format='0"番"'; c.fill=LBLUE
c.font=Font(size=12,bold=True); c.alignment=center; c.border=box
mg(ws,2,3,2,3,'名前',F_LBL,GRAY,center,border=True)
mg(ws,2,4,2,6,None,Font(size=12,bold=True),LBLUE,center,border=True,h=26)
dvk=DataValidation(type='whole',operator='between',formula1='1',formula2='20',allow_blank=True,
                   showErrorMessage=True,errorTitle='組',error='半角数字で')
dvn=DataValidation(type='whole',operator='between',formula1='1',formula2='50',allow_blank=True,
                   showErrorMessage=True,errorTitle='番',error='半角数字で')
ws.add_data_validation(dvk); ws.add_data_validation(dvn); dvk.add(ws['A2']); dvn.add(ws['B2'])
mg(ws,3,1,3,6,'※ 組・番は半角数字　／　書きこむのは「考えをまとめる」シート1枚だけ',F_RED,None,wrap)
mg(ws,5,1,5,6,'■ 単元目標',F_HEAD,NAVY,wrapc,h=20)
mg(ws,6,1,7,6,GOAL,F_GOAL,ORANGE_F,wrapc)
ws.row_dimensions[6].height=22; ws.row_dimensions[7].height=22

r=9
r=sec(ws,r,'■ 批判の目　―― 全員おなじ。鵜呑みにせず、根拠を確かめる',lc=6)
r=rows2(ws,r,[('① 誰の目か','語り手に見えていないものは何か'),
              ('② 前提は何か','疑わずに通している「当たり前」は何か'),
              ('③ 何と何の対立か','その対立は作者が作った枠ではないか'),
              ('④ 結末は何を残したか','どんな言葉で価値づけているか')],lc=6)
r+=1
r=sec(ws,r,'■ 読みの観点　―― 一人ずつちがう。自分で決める',PINK,lc=6)
mg(ws,r,1,r,6,'必修は「批評・評価」だけ。あと1〜2個を、下の「批評・評価」以外の8つから自分で選ぶ。（複数も可）',
   F_RED,None,wrapc,h=22); r+=1
for a,b in [('第1層\n何が描かれているか','登場人物　　設定　　構成　　視点・語り'),
            ('第2層\nどう描かれているか','表現の工夫　　題名　　主題'),
            ('第3層\n自分はどう考えるか','批評・評価（必修）　　自分に生かす')]:
    mg(ws,r,1,r,1,a,F_LBL,GRAY,center,border=True)
    mg(ws,r,2,r,6,b,F_BODY,None,wrapc,border=True,h=30); r+=1
r+=1
mg(ws,r,1,r,6,'★ 論題は班でおなじ。ちがうのは入口（観点）。だから話がかみ合いながら深くなる。',
   F_RED,None,wrapc,h=22); r+=2

r=sec(ws,r,'■ 話し合いの論題（3つ）　―― 3つとも、二つの作品にまたがる問い',lc=6)
mg(ws,r,1,r,6,'★ 当日、どの論題で発表するか知らされる。3つすべて準備しておく。',F_RED,PINK,wrapc,h=22); r+=1
for n,t,sub in TOPICS:
    mg(ws,r,1,r,1,n,F_LBL,GRAY,center,border=True)
    mg(ws,r,2,r,5,t,Font(size=10,bold=True),None,wrapc,border=True)
    mg(ws,r,6,r,6,sub,F_SMALL,None,center,border=True,h=26); r+=1
mg(ws,r,1,r,6,'5時＝1周目（出し合う）／6時＝2周目（別の観点をふまえて問い直す）',F_SMALL,GREEN_F,wrapc,h=20); r+=1
mg(ws,r,1,r,6,'3つとも、片方の作品だけでは答えられない。必ず両方から根拠を出すこと。',F_RED,None,wrapc,h=20); r+=2

r=sec(ws,r,'■ 授業の流れ（全7時間）',lc=6)
for t,a in [('1','「素顔同盟」「私」を読む／批判の目を知る／観点を仮決め'),
            ('2','「素顔同盟」を吟味／観点を決める／比較表の「素顔同盟」側'),
            ('3','「私」を吟味／比較表の「私」側／論題1をまとめる'),
            ('4','論題2・3の主張と引用を決める（250字にするのは家庭学習）'),
            ('5','話し合い1周目　★はじめに発表する人と論題が知らされる'),
            ('6','話し合い2周目　／　聞いて考えたことを書く'),
            ('7','まとめのディスカッション／聞いて考えたことを仕上げる')]:
    mg(ws,r,1,r,1,t+'時',F_LBL,GRAY,center,border=True)
    mg(ws,r,2,r,6,a,F_BODY,None,wrapc,border=True,h=22); r+=1
r+=1

r=sec(ws,r,'■ 1ラウンド10分の中身　／　役割は毎回交代',ORANGE_F,lc=6)
r=rows2(ws,r,[('0〜2分','提案者が話す（観点＋主張＋本文の根拠2つ）'),
              ('2〜8分','検討。ほかの人は全員、質問か反論を1回以上'),
              ('8〜10分','提案者が答える。到達点を1文にする')],lc=6)
r=rows2(ws,r,[('引用チェック係','「本文のどこ？」を必ず1回きく'),
              ('疑い係','「しかし」「とはいえ」で始める発言を必ず1回'),
              ('記録係','主張と到達点を1行ずつ書く')],lc=6,fill=PINK)
r+=1

r=sec(ws,r,'■ 7時：まとめのディスカッション',GREEN_F,lc=6)
r=rows2(ws,r,[('パネル①（12分）','訴えた人は、救われたか'),
              ('パネル②（12分）','この社会を続けさせているのは、誰か'),
              ('聞く人の仕事','考えが動いた発言を「考えをまとめる」シート④に書く')],lc=6)

# ===================== Sheet2 評価基準 =====================
ws=wb.create_sheet('評価基準')
for col,w in zip('ABCD',[8,12,58,16]): ws.column_dimensions[col].width=w
mg(ws,1,1,1,4,'評価基準',F_TITLE,NAVY,center,h=28)
mg(ws,2,1,2,4,'単元目標：'+GOAL,F_GOAL,ORANGE_F,wrapc,h=34)
mg(ws,3,1,3,4,'この単元で評価するのは【思考・判断・表現】と【主体的に学習に取り組む態度】の2つ。'
               '知識・技能は指導のみで、評価は別の単元で行う。',F_RED,None,wrapc,h=20)

r=5
mg(ws,r,1,r,4,'【知識・技能】　(1)イ　語感を磨き語彙を豊かにする　＜本単元では評価しない＞',
   Font(size=10,bold=True,color='FFFFFF'),GREYH,wrapc,h=20); r+=1
mg(ws,r,1,r,4,'★ この単元では【指導だけ】を行い、評価はしない。第3時に「私」のカギカッコ付きの語'
               '（「解決」「誠意」「模範的」「正常な状態」）を扱う。評価は別の単元または定期テストで行う。',
   F_RED,PINK,wrapc,h=32); r+=1
mg(ws,r,1,r,1,'評価規準',F_LBL,GRAY,center,border=True)
mg(ws,r,2,r,4,CRIT_K,Font(size=10,bold=True,color='595959'),GREYL,wrapc,border=True,h=28); r+=1
mg(ws,r,1,r,1,'段階',F_HEAD,NAVY,center,border=True)
mg(ws,r,2,r,4,'判断の目安',F_HEAD,NAVY,center,border=True); r+=1
for g,t,h in [
    ('A','選んだ語の働きを、語り手の立場（制度を運用する側）と結びつけて説明している。'
         '例「『解決』とは相手が満足することであって、問題が解決することではない、と語り手自身が知っている」',30),
    ('B','その語が「言葉だけは立派で中身が空」と示すことを説明し、自分の経験を挙げている',22),
    ('C','語を挙げるだけ、または経験を挙げるだけ',20),
]:
    mg(ws,r,1,r,1,g,Font(size=12,bold=True,color='808080'),GRAY,center,border=True)
    mg(ws,r,2,r,4,t,Font(size=10,color='595959'),None,wrapc,border=True,h=h); r+=1
r+=1

mg(ws,r,1,r,4,'【思考・判断・表現】　C 読むこと(1)イ　＜この単元で評価する＞',F_HEAD,NAVY,wrapc,h=20); r+=1
mg(ws,r,1,r,1,'評価規準',F_LBL,GRAY,center,border=True)
mg(ws,r,2,r,4,CRIT_S,Font(size=10,bold=True,color='1F3864'),LBLUE,wrapc,border=True,h=34); r+=1
mg(ws,r,1,r,4,'見るもの：3つの論題の記述　＋　当日の発言　＋　聞いて考えたこと（3論題分）',F_SMALL,None,wrapc,h=18); r+=1
mg(ws,r,1,r,4,'★ AとBの分かれ目 ★　条件①（3つの論題すべてを、観点と本文の根拠でまとめた）と'
               '条件②（当日、自分の観点から根拠を示して発言した）の【両方】が必要',
   F_RED,PINK,wrapc,h=32); r+=1
mg(ws,r,1,r,1,'段階',F_HEAD,NAVY,center,border=True)
mg(ws,r,2,r,4,'判断の目安',F_HEAD,NAVY,center,border=True); r+=1
for g,t,h in [
    ('S','①②に加え、3論題とも「聞いて考えたこと」を書き、自分の観点では見えなかったことに触れて'
         '今の考えを述べている',30),
    ('A','①②の両方。3つの論題すべてを観点と本文の根拠でまとめ、各論題で両作品に触れ、'
         '当日も自分の観点から根拠を示して発言した',30),
    ('B','①②の一方だけ。準備はあるが発言が根拠に届かない、または発言はよいが3つそろっていない',28),
    ('C','準備が1〜2論題だけで本文の根拠がない、または発言していない',22),
    ('D','準備がほとんどない',20),
]:
    mg(ws,r,1,r,1,g,Font(size=12,bold=True,color='C00000' if g in 'CD' else '1F3864'),
       GRAY,center,border=True)
    mg(ws,r,2,r,4,t,F_BODY,None,wrapc,border=True,h=h); r+=1
mg(ws,r,1,r,4,'※ どの観点を選んだかでは評価しない。どの観点でも、根拠を示して吟味できていればよい。',
   F_RED,None,wrapc,h=20); r+=1
mg(ws,r,1,r,4,'※ AIの評価は「記述」だけを見た暫定。当日の発言は先生が加えて確定する。',F_SMALL,None,wrapc,h=18); r+=2

mg(ws,r,1,r,4,'【主体的に学習に取り組む態度】　＜この単元で評価する＞',F_HEAD,NAVY,wrapc,h=20); r+=1
mg(ws,r,1,r,1,'評価規準',F_LBL,GRAY,center,border=True)
mg(ws,r,2,r,4,CRIT_A,Font(size=10,bold=True,color='1F3864'),LBLUE,wrapc,border=True,h=30); r+=1
mg(ws,r,1,r,4,'見るもの：観点と選んだ理由・記録係シート・聞いて考えたこと・ふりかえり',F_SMALL,None,wrapc,h=18); r+=1
mg(ws,r,1,r,1,'段階',F_HEAD,NAVY,center,border=True)
mg(ws,r,2,r,4,'判断の目安',F_HEAD,NAVY,center,border=True); r+=1
for g,t,h in [
    ('A','自分が選んだ観点で二作を読み通そうとし、他の観点からの指摘を受けて'
         '自分の読みを見直し、書き直している',28),
    ('B','自分が選んだ観点で読み、話し合いに参加して自分の考えを述べている',22),
    ('C','観点が決まらない、または話し合いに参加していない',22),
]:
    mg(ws,r,1,r,1,g,Font(size=12,bold=True,color='1F3864'),GRAY,center,border=True)
    mg(ws,r,2,r,4,t,F_BODY,None,wrapc,border=True,h=h); r+=1
r+=1
mg(ws,r,1,r,4,'【観点別評価への読みかえ】　指導要録はA・B・Cの3段階',F_HEAD,NAVY,wrapc,h=20); r+=1
mg(ws,r,1,r,2,'このシートの段階',F_HEAD,NAVY,center,border=True)
mg(ws,r,3,r,4,'観点別学習状況の評価',F_HEAD,NAVY,center,border=True); r+=1
for a,b in [('S ・ A','A（十分満足できる）'),('B','B（おおむね満足できる）'),('C ・ D','C（努力を要する）')]:
    mg(ws,r,1,r,2,a,Font(size=11,bold=True),GRAY,center,border=True)
    mg(ws,r,3,r,4,b,F_BODY,None,center,border=True,h=22); r+=1
mg(ws,r,1,r,4,'※ 生徒への返却はS〜Dのまま。指導要録へはこの表で読みかえる。'
               '知識・技能と態度はもともとA・B・Cなので、そのまま転記する。',F_SMALL,None,wrapc,h=18)

# ===================== Sheet3 考えをまとめる ★GAS参照・書きこむのはここだけ =====
ws=wb.create_sheet('考えをまとめる')
for col,w in zip('ABCD',[32,32,32,12]): ws.column_dimensions[col].width=w
mg(ws,1,1,1,4,'考えをまとめる　―― 書きこむのはこのシートだけ',F_TITLE,NAVY,center,h=28)
mg(ws,2,1,2,4,'目標：'+GOAL,F_GOAL,ORANGE_F,wrapc,h=34)
K={}

# ① 観点を決める
r=4
r=sec(ws,r,'① 自分の観点を決める（1時のおわり〜2時）',PINK)
mg(ws,r,1,r,4,'必修は「批評・評価」。あと1〜2個を残り8つから選ぶ。'
               '（登場人物／設定／構成／視点・語り／表現の工夫／題名／主題／自分に生かす）',
   F_RED,None,wrapc,h=28); r+=1
mg(ws,r,1,r,1,'私の観点（批評・評価＋1〜2個）',F_LBL,GRAY,wrapc,border=True)
mg(ws,r,2,r,4,None,Font(size=11,bold=True),INPUT,wrapc,border=True,h=30); K['LENS']=r; r+=1
mg(ws,r,1,r,1,'選んだ理由',F_LBL,GRAY,wrapc,border=True)
mg(ws,r,2,r,4,None,F_BODY,INPUT,wrap,border=True,h=44); K['REASON']=r; r+=1
mg(ws,r,1,r,1,'2時：その観点で「素顔同盟」から見えたこと',F_LBL,GRAY,wrapc,border=True)
mg(ws,r,2,r,4,None,F_BODY,INPUT,wrap,border=True,h=56); K['SAW1']=r; r+=1
mg(ws,r,1,r,1,'3時：その観点で「私」から見えたこと',F_LBL,GRAY,wrapc,border=True)
mg(ws,r,2,r,4,None,F_BODY,INPUT,wrap,border=True,h=56); K['SAW2']=r; r+=1
mg(ws,r,1,r,4,'2時のおわりに先生が見ます。観点がばらけるように班を組みます。',F_SMALL,GREEN_F,wrapc,h=18); r+=2

# ② 二作比較表
r=sec(ws,r,'② 二作比較表（2時・3時）　★の4つは必ず書く。3つの論題の材料になる',BLUE)
mg(ws,r,1,r,1,'',F_HEAD,NAVY,center,border=True)
mg(ws,r,2,r,2,'素顔同盟',F_HEAD,NAVY,center,border=True)
mg(ws,r,3,r,3,'私',F_HEAD,NAVY,center,border=True)
mg(ws,r,4,r,4,'気づいたこと',F_HEAD,NAVY,center,border=True); r+=1
for item in ['★「私」を消すものは何か（→論題1）','★訴えた人は誰で、どうなったか（→論題2）',
             '★疑わない人は誰か（→論題3）','★引用したい一文',
             '語り手はどんな立場か','結末で語り手は何をするか']:
    mg(ws,r,1,r,1,item,F_LBL,GRAY,wrapc,border=True)
    for c in (2,3,4): mg(ws,r,c,r,c,None,F_BODY,INPUT,wrap,border=True)
    ws.row_dimensions[r].height=42; r+=1
r+=1

# ③ 3つの論題について考えをまとめる
r=sec(ws,r,'③ 3つの論題について、考えをまとめる（3時・4時）',ORANGE_F)
mg(ws,r,1,r,4,'★ 当日、どの論題で発表するか知らされる。だから3つすべて準備しておく。',F_RED,PINK,wrapc,h=20); r+=1
mg(ws,r,1,r,4,'【3つとも同じ3点を書く】① 批評・評価として、どう評価するか（主張1文）　'
               '② 自分の観点から見えたこと　③ 本文の根拠（引用）　★引用は【両方の作品】から',
   Font(size=10,italic=True),None,wrapc,h=32); r+=1
for i,(n,t,sub) in enumerate(TOPICS,start=1):
    mg(ws,r,1,r,4,'【%s】%s　―― %s'%(n,t,sub),F_Q,GRAY,wrapc,h=26); r+=1
    mg(ws,r,1,r,3,None,F_BODY,INPUT,wrap,border=True,h=140); counter(ws,r,4,250)
    K['T%d'%i]=r; r+=1
r=tip(ws,r,'3つ書けたら【壁打ち】へ。根拠の弱いところと、出そうな反論だけもらう。観点は変えず、答えは書かせない。')
r+=1

# ④ 聞いて考えたこと（3論題分）
r=sec(ws,r,'④ ディスカッションを聞いて考えたこと（5時〜7時）　★3つとも書く',PINK)
mg(ws,r,1,r,4,'**自分の観点以外の観点**からの発言を受けて、考えたことを書く。'
               'だれの、どの観点からの発言かを必ず入れること。7時のパネルで聞いたことも足す。'.replace('**',''),
   F_RED,None,wrapc,h=30); r+=1
for i,(n,t,sub) in enumerate(TOPICS,start=1):
    mg(ws,r,1,r,4,'【%s】%s　―― ほかの観点を受けて考えたこと'%(n,t),F_Q,GRAY,wrapc,h=24); r+=1
    mg(ws,r,1,r,3,None,F_BODY,INPUT,wrap,border=True,h=110); counter(ws,r,4,200)
    K['R%d'%i]=r; r+=1
r=tip(ws,r,'3つ書けたら【見直しチェック】へ。3点のチェックだけもらう。書き直しはしてもらわない。')
r+=1

# ⑤ 当日メモ・記録係
r=sec(ws,r,'⑤ 当日メモと記録係シート（5時・6時）',BLUE)
mg(ws,r,1,r,2,'私が発表する論題（当日に知らされる）',F_LBL,GRAY,wrapc,border=True)
mg(ws,r,3,r,4,None,Font(size=12,bold=True),INPUT,center,border=True,h=24); K['DAY']=r; r+=1
mg(ws,r,1,r,1,'ラウンド',F_HEAD,NAVY,center,border=True)
mg(ws,r,2,r,2,'提案者の主張',F_HEAD,NAVY,center,border=True)
mg(ws,r,3,r,3,'到達点（1文）',F_HEAD,NAVY,center,border=True)
mg(ws,r,4,r,4,'観点',F_HEAD,NAVY,center,border=True); r+=1
for lab in ['5時 論題1','5時 論題2','5時 論題3','6時 論題1','6時 論題2','6時 論題3']:
    mg(ws,r,1,r,1,lab,F_LBL,GRAY,center,border=True)
    for c in (2,3,4): mg(ws,r,c,r,c,None,F_BODY,INPUT,wrap,border=True)
    ws.row_dimensions[r].height=34; r+=1
r+=1

# ⑥ ふりかえり
r=sec(ws,r,'⑥ ふりかえり（7時）',GREEN_F)
mg(ws,r,1,r,4,'この単元で、自分の読み方はどこが変わったか　150字',F_LBL,GRAY,wrapc,h=20); r+=1
mg(ws,r,1,r,3,None,F_BODY,INPUT,wrap,border=True,h=90); counter(ws,r,4,150)
K['REFLECT']=r

print('CELLS  LENS=C%d REASON=B%d SAW1=B%d SAW2=B%d'%(K['LENS'],K['REASON'],K['SAW1'],K['SAW2']))
print('       T1=A%d T2=A%d T3=A%d'%(K['T1'],K['T2'],K['T3']))
print('       R1=A%d R2=A%d R3=A%d'%(K['R1'],K['R2'],K['R3']))
print('       DAY=C%d REFLECT=A%d'%(K['DAY'],K['REFLECT']))

# ===================== Sheet4 AIプロンプト =====================
ws=wb.create_sheet('AIプロンプト')
ws.column_dimensions['A'].width=92
mg(ws,1,1,1,1,'AIプロンプト（コピーして使う）',F_TITLE,NAVY,center,h=28)
mg(ws,2,1,2,1,'名前は書かない／まず自分で書く／書かせるのではなく、ヒントをもらう',F_RED,PINK,wrapc,h=22)

P1='''中学3年生にわかる言葉で答えてください。

「素顔同盟」（すやまたけし）と「私」（三崎亜記）について、3つの論題でグループ討論をします。
論題1 仮面とデータは、同じものか
論題2 訴えた人は、救われたか
論題3 この社会を続けさせているのは、誰か

私は「批評・評価」と、自分で決めた観点から答えます。
下が3つの論題への私の考えです。次の2点だけ教えてください。
1 本文の根拠が弱いところ（どの論題の、どこか）
2 班の人から出そうな反論

★私の観点を別の観点に変える提案はしないでください。この観点のまま深めるヒントを。
★答えを代わりに書かないでください。300字以内で。

【私の観点】
★ここに貼る★
【論題1への考え】
★ここに貼る★
【論題2への考え】
★ここに貼る★
【論題3への考え】
★ここに貼る★'''
mg(ws,4,1,4,1,'【壁打ち】3つの考えを書いたあと（4時）',F_HEAD,NAVY,wrapc,h=20)
mg(ws,5,1,5,1,P1,Font(size=10,name='Consolas'),LBLUE,wrap,h=330)

P2='''中学3年生にわかる言葉で答えてください。

グループ討論を聞いて考えたことを、3つの論題それぞれについて書きました。
次をチェックしてください。
□ ① だれの、どの観点からの発言かが書けているか（3つとも）
□ ② それが【自分の観点とはちがう観点】からの発言になっているか
□ ③ その発言を受けて、自分の考えがどう動いたかが書けているか
□ 「いろいろな意見が出た」だけの感想になっていないか

書き直した文章は出さないでください。ヒントだけ、300字以内で。

【私の観点】
★ここに貼る★
【論題1：聞いて考えたこと】
★ここに貼る★
【論題2：聞いて考えたこと】
★ここに貼る★
【論題3：聞いて考えたこと】
★ここに貼る★'''
mg(ws,7,1,7,1,'【見直しチェック】聞いて考えたことを書いたあと（7時）',F_HEAD,NAVY,wrapc,h=20)
mg(ws,8,1,8,1,P2,Font(size=10,name='Consolas'),LBLUE,wrap,h=300)

wb.save(OUT)
print('saved',OUT)
