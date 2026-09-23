// 模写ドリル ワークシート生成
// 使い方: node build/build.mjs <pack-id> [--png] [--all-png]
//   content/<pack-id>.json を読み、dist/<pack-id>/<title>.pdf を出力する
//   --png     1 枚目（PNG_NO=n で n 枚目）を PNG 出力（確認用）
//   --all-png 全シートを dist/<pack-id>/png/ に 1 枚ずつ PNG 出力（商品画像・SNS 用の確認画像）
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packId = process.argv[2];
const wantPng = process.argv.includes('--png');
const wantAllPng = process.argv.includes('--all-png');
if (!packId) { console.error('usage: node build/build.mjs <pack-id> [--png]'); process.exit(1); }

const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'content', `${packId}.json`), 'utf8'));
const css = fs.readFileSync(path.join(ROOT, 'template', 'sheet.css'), 'utf8');

const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function sampleMarkup(sheet, forTrace) {
  // sample 未指定なら samples/<pack>/NNN.(png|jpg|jpeg|svg) を順に探す。手描きスキャン(png)が SVG より優先される
  let rel = sheet.sample;
  if (!rel) {
    const base = `samples/${packId}/${String(sheet.no).padStart(3, '0')}`;
    rel = ['png', 'jpg', 'jpeg', 'svg'].map(e => `${base}.${e}`).find(r => fs.existsSync(path.join(ROOT, r))) || `${base}.png`;
  }
  const abs = path.join(ROOT, rel);
  if (fs.existsSync(abs)) {
    if (abs.endsWith('.svg')) return fs.readFileSync(abs, 'utf8');
    const ext = path.extname(abs).slice(1).toLowerCase();
    const mime = ext === 'jpg' ? 'image/jpeg' : `image/${ext}`;
    const b64 = fs.readFileSync(abs).toString('base64');
    return `<img src="data:${mime};base64,${b64}" alt="">`;
  }
  return `<div class="placeholder"><b>お手本（手描き）</b>${esc(sheet.motif)}<br><span style="font-size:6.5pt">${esc(rel)}</span></div>`;
}

function gridClass(g) {
  if (!g || g === 'none') return 'grid frame';
  if (g === '2lines') return 'grid g2 frame';
  const n = String(g).split('x')[0];
  return `grid g${n} frame`;
}

function progress(sheet, total, repeats) {
  let cells = '';
  for (let i = 1; i <= total; i++) {
    const cls = ['c', i < sheet.no ? 'done' : '', i === sheet.no ? 'now' : '', repeats.includes(i) ? 'rep' : ''].filter(Boolean).join(' ');
    cells += `<span class="${cls}"></span>`;
  }
  return cells;
}

function renderSheet(sheet, pack) {
  const t = Object.assign({ trace: 3, draw: 10, memory: 5 }, pack.time || {}, sheet.time || {});
  const isLine = sheet.kind === 'line';
  const hasTrace = sheet.trace !== false;
  const hasMemory = !!sheet.memory;
  const checks = (sheet.checks || []).map((c, i) => `<li class="${i === 0 ? 'first' : ''}">${esc(c)}</li>`).join('');
  const steps = sheet.kind === 'motif' || sheet.kind === 'shade' || sheet.kind === 'color'
    ? `<div class="steps"><div class="s"><b>① 外形</b>${esc(sheet.steps?.[0] || 'いちばん外のシルエットと縦横比')}</div><div class="s"><b>② かたまり</b>${esc(sheet.steps?.[1] || '大きなパーツの位置と大きさ')}</div><div class="s"><b>③ 細部</b>${esc(sheet.steps?.[2] || '最後に細かい線')}</div></div>`
    : '';
  const rules = [];
  if (isLine) rules.push('<div class="rule warn">引き直し禁止。1本の線を1回で。ずれても消さない。</div>');
  if (sheet.ratioMemo) rules.push('<div class="ratio">描く前に測る： 縦 : 横 = <span></span> : <span></span></div>');
  if (sheet.darkMark) rules.push('<div class="rule">塗る前に、いちばん暗い所に × を付ける。次に、いちばん明るい所に ○。</div>');
  if (sheet.palette) rules.push(`<div class="rule">使う色は3つだけ： ${sheet.palette.map(esc).join(' ／ ')}<br>主 ＝ ______　従 ＝ ______　差し色 ＝ ______</div>`);
  if (sheet.rule) rules.push(`<div class="rule">${esc(sheet.rule)}</div>`);

  const gc = gridClass(sheet.grid);
  const gridLabel = sheet.grid && sheet.grid !== 'none' ? (sheet.grid === '2lines' ? '基準線2本' : 'グリッド ' + sheet.grid) : 'グリッドなし';
  const traceBox = hasTrace ? `
    <div class="box trace"><div class="lab">なぞる<span class="t">1回だけ・同じマスを見ながら</span></div><div class="time">目安 ${t.trace}分</div>
      <div class="inner">${sampleMarkup(sheet, true)}<div class="${gc}"></div></div></div>` : '';
  const drawBox = `
    <div class="box draw"><div class="lab">見て描く<span class="t">${esc(gridLabel)}</span></div><div class="time">目安 ${t.draw}分</div>
      <div class="inner"><div class="${gc}"></div></div></div>`;
  const memBox = hasMemory ? `
    <div class="box mem"><div class="lab">記憶で描く<span class="t">お手本を折って隠す</span></div><div class="time">目安 ${t.memory}分</div><div class="grid frame"></div></div>` : '';

  return `
<section class="sheet">
  <div class="hd">
    <div><div class="pack">${esc(pack.title)}</div><div class="no">${String(sheet.no).padStart(2, '0')}<small>/ ${pack.sheets.length}</small></div></div>
    <div class="title">${esc(sheet.title)}</div>
    <div class="tags"><span class="tag fill">${esc(pack.level)}</span><span class="tag">${esc(sheet.style || '共通')}</span>${sheet.phase ? `<span class="tag">${esc(sheet.phase)}</span>` : ''}</div>
    <div class="date">日付<span></span></div>
  </div>
  <div class="r1">
    <div class="box sample"><div class="lab">お手本<span class="t">${esc(gridLabel)}</span></div><div class="inner">${sampleMarkup(sheet, false)}<div class="${gc}"></div></div></div>
    <div class="point">
      <div class="p"><b>今日のポイント（これだけ盗む）</b>${esc(sheet.point)}</div>
      ${steps}
      ${rules.join('')}
    </div>
  </div>
  <div class="r2 ${hasTrace ? '' : 'wide'}" style="${hasTrace ? '' : 'grid-template-columns:1fr'}">
    ${traceBox}${drawBox}
  </div>
  <div class="r3 ${hasMemory ? '' : 'nomem'}">
    ${memBox}
    <div class="box checks"><div class="lab">描いたら比べる<span class="t">お手本と並べて</span></div>
      <ol>${checks}</ol>
      <div class="compare">比べ方：並べる → 離れて見る → 紙を裏から透かして重ねる。ずれた所に ○ を付けてから直す。</div>
    </div>
  </div>
  <div class="ft">
    <div class="note"><span class="l">今日の気づき（1行・空欄なら進捗を塗れない）</span><span class="u"></span></div>
    <div class="prog"><span class="l">進捗</span><div class="cells">${progress(sheet, pack.sheets.length, pack.repeats || [])}</div></div>
    <div class="foot"><span>${esc(pack.brand || '')}　${esc(pack.title)}　｜　A4・白黒・「実際のサイズ」で印刷</span><span>${esc(pack.copyright || '')}　※購入者本人の練習用。再配布禁止。</span></div>
  </div>
</section>`;
}

const html = `<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"><title>${esc(manifest.title)}</title><style>${css}</style></head><body>${manifest.sheets.map(s => renderSheet(s, manifest)).join('')}</body></html>`;

const outDir = path.join(ROOT, 'dist', packId);
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, 'sheets.html'), html);

const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const browser = await chromium.launch();
const page = await browser.newPage();
await page.setContent(html, { waitUntil: 'load' });
const pdfPath = path.join(outDir, `${manifest.title}.pdf`);
await page.pdf({ path: pdfPath, format: 'A4', printBackground: true, preferCSSPageSize: true });
if (wantPng) {
  await page.setViewportSize({ width: 794, height: 1123 });
  const n = Number(process.env.PNG_NO || 1);
  await page.evaluate(i => { document.querySelectorAll('.sheet').forEach((el, k) => { el.style.display = (k === i - 1) ? '' : 'none'; }); }, n);
  await page.screenshot({ path: path.join(outDir, `preview-${String(n).padStart(3, '0')}.png`), fullPage: false });
}
if (wantAllPng) {
  // 確認・商品画像用: 1枚ずつ PNG
  const pngDir = path.join(outDir, 'png');
  fs.mkdirSync(pngDir, { recursive: true });
  await page.setViewportSize({ width: 794, height: 1123 });
  for (let i = 1; i <= manifest.sheets.length; i++) {
    await page.evaluate(k => { document.querySelectorAll('.sheet').forEach((el, j) => { el.style.display = (j === k - 1) ? '' : 'none'; }); }, i);
    await page.screenshot({ path: path.join(pngDir, `${String(i).padStart(3, '0')}.png`), scale: 'css', clip: { x: 0, y: 0, width: 794, height: 1123 } });
  }
  console.log(`png: ${manifest.sheets.length} files in ${pngDir}`);
}
await browser.close();
console.log(`ok: ${pdfPath} (${manifest.sheets.length} sheets)`);
