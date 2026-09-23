// 基礎ドリルの線・図形・シルエット課題のお手本 SVG を生成する
// 手描きのお手本（motif 課題）は samples/basic/0NN.png を置けば自動で優先される
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(ROOT, 'samples', 'basic');
fs.mkdirSync(out, { recursive: true });

const S = 'stroke="#111" stroke-width="2.2" fill="none" stroke-linecap="round" stroke-linejoin="round"';
const wrap = body => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">${body}</svg>`;
const rows = (n, f) => Array.from({ length: n }, (_, i) => f(i)).join('');

const gen = {
  1: () => wrap(rows(6, i => `<line x1="12" y1="${18 + i * 13}" x2="88" y2="${18 + i * 13}" ${S}/>`)),
  2: () => wrap(rows(6, i => `<line x1="${18 + i * 13}" y1="12" x2="${18 + i * 13}" y2="88" ${S}/>`)),
  3: () => wrap(rows(5, i => `<line x1="${10 + i * 12}" y1="88" x2="${34 + i * 12}" y2="12" ${S}/>`) + rows(3, i => `<line x1="${8 + i * 14}" y1="12" x2="${22 + i * 14}" y2="40" ${S}/>`)),
  4: () => wrap(`<circle cx="12" cy="20" r="2.5" fill="#111"/><circle cx="88" cy="30" r="2.5" fill="#111"/><line x1="12" y1="20" x2="88" y2="30" ${S}/>
    <circle cx="15" cy="55" r="2.5" fill="#111"/><circle cx="85" cy="85" r="2.5" fill="#111"/><line x1="15" y1="55" x2="85" y2="85" ${S}/>
    <circle cx="20" cy="88" r="2.5" fill="#111"/><circle cx="80" cy="50" r="2.5" fill="#111"/><line x1="20" y1="88" x2="80" y2="50" ${S}/>`),
  5: () => wrap(rows(4, i => `<path d="M10 ${22 + i * 18} q 10 -12 20 0 t 20 0 t 20 0 t 20 0" ${S}/>`)),
  6: () => wrap(`<path d="M10 50 A 40 40 0 0 1 90 50" ${S}/><path d="M20 80 A 30 30 0 0 1 80 80" ${S}/><path d="M30 25 A 20 20 0 0 0 70 25" ${S}/>`),
  7: () => wrap(`<ellipse cx="50" cy="28" rx="36" ry="12" ${S}/><ellipse cx="50" cy="58" rx="28" ry="16" ${S}/><ellipse cx="50" cy="86" rx="40" ry="8" ${S}/>`),
  8: () => wrap(`<circle cx="30" cy="30" r="18" ${S}/><circle cx="70" cy="30" r="12" ${S}/><circle cx="50" cy="70" r="24" ${S}/>`),
  9: () => wrap(`<rect x="10" y="14" width="34" height="34" ${S}/><rect x="54" y="14" width="36" height="20" ${S}/><rect x="10" y="58" width="22" height="34" ${S}/><rect x="42" y="58" width="48" height="30" ${S}/>`),
  10: () => wrap(`<path d="M50 12 L88 82 L12 82 Z" ${S}/><path d="M28 30 L46 60 L10 60 Z" ${S}/>`),
  11: () => wrap(`<circle cx="50" cy="66" r="24" ${S}/><circle cx="50" cy="30" r="15" ${S}/><circle cx="45" cy="27" r="1.6" fill="#111"/><circle cx="55" cy="27" r="1.6" fill="#111"/><path d="M50 32 l-4 3 h8 z" fill="#111"/><rect x="38" y="10" width="24" height="6" ${S}/><rect x="42" y="2" width="16" height="9" ${S}/>`),
  12: () => wrap(`<path d="M50 88 C 10 60, 10 22, 32 20 C 42 19, 50 30, 50 34 C 50 30, 58 19, 68 20 C 90 22, 90 60, 50 88 Z" ${S}/>`),
  13: () => wrap(`<path d="M50 10 L59 36 L88 36 L65 52 L74 80 L50 63 L26 80 L35 52 L12 36 L41 36 Z" ${S}/>`),
  14: () => wrap(`<rect x="12" y="20" width="28" height="42" ${S}/><text x="26" y="76" font-size="7" text-anchor="middle" fill="#111">2 : 3</text><rect x="50" y="14" width="42" height="28" ${S}/><text x="71" y="56" font-size="7" text-anchor="middle" fill="#111">3 : 2</text><rect x="50" y="64" width="42" height="14" ${S}/><text x="71" y="90" font-size="7" text-anchor="middle" fill="#111">1 : 3</text>`),
  15: () => wrap(`<rect x="16" y="24" width="40" height="40" ${S}/><circle cx="62" cy="58" r="22" ${S}/><path d="M60 14 L84 14 L72 40 Z" ${S}/>`),
  16: () => wrap(`<rect x="30" y="30" width="40" height="40" transform="rotate(20 50 50)" ${S}/><rect x="12" y="10" width="20" height="20" transform="rotate(-30 22 20)" ${S}/><path d="M70 72 L92 78 L80 96 Z" ${S}/>`),
  17: () => wrap(`<path d="M50 24 C 38 12, 14 18, 14 46 C 14 70, 34 92, 50 84 C 66 92, 86 70, 86 46 C 86 18, 62 12, 50 24 Z" fill="#111"/><path d="M50 24 C 52 16, 56 12, 60 10" ${S}/>`),
  18: () => wrap(`<path d="M50 14 C 42 14, 40 30, 38 42 C 34 56, 16 62, 18 78 C 20 92, 80 92, 82 78 C 84 62, 66 56, 62 42 C 60 30, 58 14, 50 14 Z" fill="#111"/>`),
  19: () => wrap([0, 90, 180, 270].map(a => `<path d="M50 50 C 40 38, 20 34, 20 48 C 20 58, 40 62, 50 50 Z" fill="#111" transform="rotate(${a} 50 50) translate(0 -2)"/>`).join('') + `<path d="M50 52 L 58 88" stroke="#111" stroke-width="4" fill="none"/>`),
  20: () => wrap(`<path d="M20 26 H 68 V 84 H 20 Z" fill="#111"/><path d="M68 40 C 88 36, 90 66, 68 70" stroke="#111" stroke-width="8" fill="none"/>`),
  21: () => wrap(`<path d="M38 8 H 62 V 14 H 58 V 24 C 68 28, 72 38, 72 48 V 90 H 28 V 48 C 28 38, 32 28, 42 24 V 14 H 38 Z" fill="#111"/>`),
  22: () => wrap(`<path d="M30 30 L 26 12 L 40 22 C 46 20, 54 20, 60 22 L 74 12 L 70 30 C 82 42, 80 60, 66 68 L 70 90 H 30 L 34 68 C 20 60, 18 42, 30 30 Z" fill="#111"/>`),
  23: () => wrap(`<circle cx="50" cy="62" r="26" fill="#111"/><circle cx="60" cy="30" r="16" fill="#111"/><path d="M76 30 L 88 34 L 76 38 Z" fill="#111"/><path d="M40 86 L 36 96 M46 88 L 46 96 M56 88 L 56 96 M62 86 L 66 96" stroke="#111" stroke-width="3" fill="none"/>`),
  24: () => wrap(`<path d="M50 8 C 30 12, 14 30, 16 50 C 18 66, 30 74, 46 74 V 92 H 54 V 74 C 70 74, 82 66, 84 50 C 86 30, 70 12, 50 8 Z" fill="#111"/>`),
  25: () => wrap(`<path d="M12 62 C 12 52, 20 50, 26 50 L 34 36 H 66 L 74 50 C 84 50, 90 54, 90 62 V 72 H 12 Z" fill="#111"/><circle cx="28" cy="74" r="8" fill="#111"/><circle cx="72" cy="74" r="8" fill="#111"/><circle cx="28" cy="74" r="3" fill="#fff"/><circle cx="72" cy="74" r="3" fill="#fff"/>`),
  26: () => wrap(`<path d="M8 50 C 8 20, 92 20, 92 50 Z" fill="#111"/><rect x="30" y="50" width="40" height="40" fill="#111"/><path d="M42 90 V 72 C 42 64, 58 64, 58 72 V 90 Z" fill="#fff"/><rect x="66" y="26" width="8" height="12" fill="#111"/>`),
};

for (const [no, f] of Object.entries(gen)) {
  fs.writeFileSync(path.join(out, `${String(no).padStart(3, '0')}.svg`), f());
}
console.log(`generated ${Object.keys(gen).length} sample svgs in samples/basic`);
