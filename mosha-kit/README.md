# 模写300 — おえかきお父さんの模写ドリル（制作キット）

「画風を選んで、300枚模写する。」有料ワークブックの制作一式。
企画・競合分析・練習設計は `../docs/mosha-kit/企画提案書.md`。

## 構成

| パス | 中身 |
|---|---|
| `content/*.json` | 各パックの内容（1枚ごとのタイトル・今日のポイント・チェック項目・グリッド・目安時間） |
| `samples/<pack>/NNN.svg` | **お手本（AI 作成の手描き風線画）。** `gen-samples.mjs`（基礎 01〜26）と `gen-samples-ai.mjs`（基礎 27〜49・発展 01〜30）で生成 |
| `samples/<pack>/NNN.png` | 手描きのお手本を置く場所。同じ番号の SVG より **PNG が優先**される |
| `template/sheet.css` | ワークシートの見た目（A4） |
| `build/build.mjs` | JSON → HTML → PDF/PNG |
| `build/gen-samples.mjs` | 基礎の SVG お手本を生成 |
| `extras/extras.html` + `build/build-extras.mjs` | 同梱物（使い方・30日カレンダー・比較シート・ルーブリック・診断表） |
| `lp/index.html` | 専用ページ（診断つき）。別リポジトリに移して GitHub Pages で公開する想定 |
| `booth/商品ページ文面.md` | BOOTH の商品説明文 |
| `dist/` | 出力（PDF・PNG） |

## お手本について

現状はすべて AI 作成の線画（SVG）。1点ずつ差し替えたい時は、同じ番号の PNG を置くだけでよい（PNG が優先）。
差し替えた分だけ「手描き」と言える。商品説明の「お手本について」は、実態に合わせて書き換えること。

## 手描きお手本の入れ方

1. A4 の紙に描く（枠いっぱいでなく、周囲に余白を残す）。ペン入れ済みの線画が望ましい
2. スキャン or スマホで撮影 → 傾き補正 → 白背景で書き出し（PNG、長辺 1600px 以上）
3. `samples/basic/027.png` のように、**シート番号と同じ名前**で置く
4. ビルドすると、お手本枠に表示され、なぞり枠には自動で淡色化（22%・グレー）されて入る

置いていない番号は「お手本（手描き）をここに」の点線枠で出力されるので、途中でもビルドできる。

## ビルド

```bash
cd mosha-kit
node build/gen-samples.mjs            # 基礎 01〜26 の SVG お手本（初回のみ）
node build/gen-samples-ai.mjs         # 基礎 27〜49・発展 01〜30 の SVG お手本（初回のみ）
node build/contact-sheet.mjs basic    # お手本一覧を dist/basic/samples.png に出して確認
node build/build.mjs basic            # dist/basic/基礎ドリル.pdf
node build/build.mjs basic --all-png  # + dist/basic/png/001.png … （タブレット用）
node build/build.mjs dev-anime
node build/build-extras.mjs           # dist/extras/同梱物.pdf
PNG_NO=27 node build/build.mjs basic --png   # 27枚目だけ PNG で確認
```

Node 22 と Playwright（Chromium）が必要。日本語フォントは IPA ゴシック（`fc-list | grep IPA` で確認）。

## 内容の編集

`content/<pack>.json` の1枚分:

```json
{ "no": 27, "kind": "motif", "phase": "グリッド模写", "style": "リアル",
  "title": "りんご（線画）", "motif": "りんご線画（手描き）",
  "grid": "4x4",            // "none" | "2lines" | "3x3" | "4x4" | "6x6"
  "memory": false,          // 記憶で描く枠
  "trace": true,            // なぞる枠（false で「見て描く」1枠のみ）
  "ratioMemo": false,       // 「縦:横= : 」欄
  "darkMark": false,        // 「いちばん暗い所に×」（明暗課題）
  "palette": ["赤","紺","生成り"],  // 3色限定課題
  "point": "各マスの中で「線がどこを通るか」だけ見る。",
  "steps": ["マスごとに輪郭が通る点", "点を結ぶ", "へたとくぼみ"],
  "checks": ["各マスで線の通過点は合っているか", "全体を見て縦横比は合っているか", "線は1回で引けたか"] }
```

`kind` は `line`（引き直し禁止の注記）／`motif`（①外形②かたまり③細部の手順）／`shade`（明暗）／`color`。
パック全体の `repeats` に反復シートの番号を入れると、進捗マスが破線になる。

## 新しいパックを足す

1. `content/dev-yuru.json` のように JSON を作る（`id` はファイル名と同じ）
2. `samples/dev-yuru/` にお手本を置く
3. `node build/build.mjs dev-yuru`
4. BOOTH の既存商品に「ファイルの追加」でアップロード。`lp/index.html` の `CONFIG` にリンクと画風の状態を追記

## 発売前チェック

- [ ] 全番号にお手本が入っている（点線枠が残っていない）。差し替えた手描きがあれば PNG で上書き済み
- [ ] 家庭用プリンタで白黒印刷して、なぞり線が見える
- [ ] `lp/index.html` の `CONFIG.booth` を実際の URL に差し替えた
- [ ] BOOTH の説明文の「お手本について」が実態（AI 作成／手描き）と一致している。「再DL可」「再配布禁止」が入っている
- [ ] 無料お試し版（8枚）を 0 円商品として出した
