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
  // CTPT 2018 menu: open book
  ctpt2018: `<path d="M16 9c-2.5-2-6-2.5-10-2v17c4-.5 7.5 0 10 2 2.5-2 6-2.5 10-2V7c-4-.5-7.5 0-10 2z" ${S}/><path d="M16 9v17" ${S}/>`,
  // Tạo bảng đáp án: answer grid with a check in the first answer cell
  taoBangDapAn: `<rect x="5" y="7" width="22" height="18" rx="2" ${S}/><path d="M5 16h22M12.5 7v18M19.5 7v18" ${S}/>
         <path d="M6.8 20.6l1.6 1.6 2.8-3" ${S}/>`,
  // Canh giữa Table: small table between two margin lines
  canhGiuaTable: `<path d="M4 5v22M28 5v22" ${S}/><rect x="9" y="10" width="14" height="12" rx="1" ${S}/><path d="M9 16h14M16 10v12" ${S}/>`,
  // Chuyển STT Text sang Auto: typed line -> numbered list (arrow)
  sttTextToAuto: `<path d="M5 8h9M5 13h7" ${S}/><path d="M17 10.5h6M20.5 8l2.5 2.5-2.5 2.5" ${S}/><circle cx="7" cy="22" r="1.5" ${S}/><circle cx="7" cy="28" r="1.5" ${S}/><path d="M12 22h15M12 28h15" ${S}/>`,
  // Equations menu: radical sign
  equations: `<path d="M4 18h4l4 8 7-21h9" ${S}/>`,
  // LaTeX -> Equation: braces, arrow, radical
  latexToEquation: `<path d="M6 9c-1.5 0-2 1-2 3s-.5 3-2 4c1.5 1 2 2 2 4s.5 3 2 3" ${S}/>
         <path d="M9.5 16h5M12.5 13.5L15 16l-2.5 2.5" ${S}/><path d="M18 18h1.5l2 4 3-11H30" ${S}/>`,
  // Equation -> LaTeX: radical, arrow, braces
  equationToLatex: `<path d="M2 18h1.5l2 4 3-11H13" ${S}/>
         <path d="M16 16h5M19 13.5l2.5 2.5L19 18.5" ${S}/><path d="M27 9c1.5 0 2 1 2 3s.5 3 2 4c-1.5 1-2 2-2 4s-.5 3-2 3" ${S}/>`,
  // Chuẩn hoá BigKey menu: key
  bigkey: `<circle cx="10" cy="16" r="5" ${S}/><path d="M15 16h13M23 16v4M27 16v3" ${S}/>`,
  // Chuẩn hoá BigKey: question line + four option columns
  chuanHoaBigKey: `<path d="M5 8h22" ${S}/><path d="M5 15h4M13 15h4M21 15h4M5 22h4M13 22h4M21 22h4" ${S}/>`,
  // Check lỗi đề: magnifier over a line
  checkLoiDe: `<circle cx="14" cy="14" r="7" ${S}/><path d="M19 19l7 7M11 14h6" ${S}/>`,
  // Xoá cần thiết menu: eraser
  xoa: `<path d="M18 6l8 8-11 11H9l-4-4z" ${S}/><path d="M12 12l8 8M15 25h12" ${S}/>`,
  // Xoá trích nguồn: bracket tag struck through
  xoaTrichNguon: `<path d="M10 8H7v16h3M22 8h3v16h-3" ${S}/><path d="M12 20l8-8" ${S}/>`,
  // Xoá lời giải: page with lines and a cross
  xoaLoiGiai: `${page}<path d="M13 16l6 6M19 16l-6 6" ${S}/>`,
  // Xoá tác giả: person with minus
  xoaTacGia: `<circle cx="13" cy="11" r="5" ${S}/><path d="M4 27c1.5-5 4.5-7.5 9-7.5 2.2 0 4 .6 5.5 1.7" ${S}/><path d="M21 24h8" ${S}/>`,
  // Xoá mức độ… (task pane): level bars
  xoaMucDo: `<path d="M7 26v-6M13 26v-10M19 26v-14M25 26V6" ${S}/>`,
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
