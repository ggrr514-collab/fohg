// 同梱物（使い方・カレンダー・比較シート・ルーブリック・診断表）を PDF にする
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const browser = await chromium.launch();
const page = await browser.newPage();
await page.goto('file://' + path.join(ROOT, 'extras', 'extras.html'), { waitUntil: 'load' });
fs.mkdirSync(path.join(ROOT, 'dist', 'extras'), { recursive: true });
const out = path.join(ROOT, 'dist', 'extras', '同梱物.pdf');
await page.pdf({ path: out, format: 'A4', printBackground: true, preferCSSPageSize: true });
if (process.argv.includes('--png')) {
  await page.setViewportSize({ width: 794, height: 1123 });
  const n = Number(process.env.PNG_NO || 1);
  await page.evaluate(i => { document.querySelectorAll('.pg').forEach((el, k) => { el.style.display = (k === i - 1) ? '' : 'none'; }); }, n);
  await page.screenshot({ path: path.join(ROOT, 'dist', 'extras', `preview-${n}.png`) });
}
await browser.close();
console.log('ok:', out);
