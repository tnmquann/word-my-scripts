// Generate ribbon icons (PNG) from inline SVG into assets/icons/.
// Run: npm run icons   (output is committed; only re-run after changing designs here)
//
// Style: minimal line icons — single accent colour, 2px strokes on a 32px grid, round caps,
// no text glyphs (fonts render inconsistently at 16px).
import sharp from "sharp";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const OUT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "assets", "icons");
const SIZES = [16, 32, 80];

const C = "#2B579A"; // Word blue
const S = `fill="none" stroke="${C}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"`;
const page = `<path d="M9 4h10l6 6v18H9z" ${S}/><path d="M19 4v6h6" ${S}/>`;

const ICONS = {
  // App logo: solid tile with a ">_" prompt
  addin: `<rect x="2" y="2" width="28" height="28" rx="7" fill="${C}"/>
          <path d="M9.5 11l5 5-5 5" fill="none" stroke="#fff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M17 21.5h6" stroke="#fff" stroke-width="2.6" stroke-linecap="round"/>`,
  // ĐGNL menu: page with a check
  dgnl: `${page}<path d="M12.5 18.5l2.8 2.8 5.2-5.6" ${S}/>`,
  // Chuẩn hoá Toán: sigma
  chuanHoaToan: `<path d="M23 7H9l8 9-8 9h14" ${S}/>`,
  // Fix tab câu hỏi: aligned list with hanging markers
  fixTabNumbering: `<path d="M6 9h2M6 16h2M6 23h2" ${S}/><path d="M13 9h13M13 16h13M13 23h13" ${S}/>`,
  // Gạch chân đáp án: letter A + underline
  gachDapAnNguVan: `<path d="M10 21l6-15 6 15M12.4 15.5h7.2" ${S}/><path d="M8 26h16" ${S}/>`,
  // Fix break dòng: two lines joined across a dashed break
  removeBreakEng: `<path d="M6 8h20M6 24h20" ${S}/><path d="M6 16h3M12.5 16h3M19 16h3" ${S}/><path d="M26 12v8" ${S}/>`,
  // Xoá ĐACT: page with minus
  xoaDactDgnl: `${page}<path d="M12.5 19h8" ${S}/>`,
  // Tác giả: person
  author: `<circle cx="16" cy="11" r="5" ${S}/><path d="M6.5 27c1.5-5 5-7.5 9.5-7.5s8 2.5 9.5 7.5" ${S}/>`,
  // Kiểm tra cập nhật: circular arrow
  update: `<path d="M25 16a9 9 0 1 1-2.6-6.4" ${S}/><path d="M23 4.5v5.5h-5.5" ${S}/>`,
};

await fs.mkdir(OUT, { recursive: true });
for (const [name, body] of Object.entries(ICONS)) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">${body}</svg>`;
  const sizes = name === "addin" ? [32, 64, 80] : SIZES;
  for (const s of sizes) {
    await sharp(Buffer.from(svg), { density: (72 * s) / 32 })
      .resize(s, s)
      .png()
      .toFile(path.join(OUT, `${name}-${s}.png`));
  }
}
console.log(`Icons written to ${OUT}`);
