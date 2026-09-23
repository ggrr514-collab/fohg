// お手本一覧（確認用）: node build/contact-sheet.mjs <pack> → dist/<pack>/samples.png
import fs from 'node:fs'; import path from 'node:path'; import { fileURLToPath } from 'node:url'; import { createRequire } from 'node:module';
const { chromium } = createRequire(import.meta.url)('/opt/node22/lib/node_modules/playwright');
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pack = process.argv[2];
const dir = path.join(ROOT, 'samples', pack);
const files = fs.readdirSync(dir).filter(f => /\.(svg|png|jpg)$/.test(f)).sort();
const cells = files.map(f => { const p = path.join(dir, f); const inner = f.endsWith('.svg') ? fs.readFileSync(p, 'utf8') : `<img src="data:image/${path.extname(f).slice(1)};base64,${fs.readFileSync(p).toString('base64')}">`; return `<div class="c"><div class="i">${inner}</div><div class="n">${f}</div></div>`; }).join('');
const html = `<html><head><meta charset="utf-8"><style>body{margin:0;background:#fff;font-family:sans-serif}.g{display:grid;grid-template-columns:repeat(6,190px);gap:8px;padding:8px}.c{border:1px solid #ddd;border-radius:6px;padding:6px}.i{width:176px;height:176px}.i svg,.i img{width:100%;height:100%}.n{font-size:11px;color:#666;text-align:center}</style></head><body><div class="g">${cells}</div></body></html>`;
const b = await chromium.launch(); const pg = await b.newPage({ viewport: { width: 1220, height: 800 } });
await pg.setContent(html); const out = path.join(ROOT, 'dist', pack, 'samples.png'); fs.mkdirSync(path.dirname(out), { recursive: true });
await pg.screenshot({ path: out, fullPage: true }); await b.close(); console.log('ok:', out, files.length, 'files');
