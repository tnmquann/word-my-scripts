// Generate ribbon icons (PNG) from inline SVG into assets/icons/.
// Run: npm run icons   (output is committed; only re-run after changing designs here)
import sharp from "sharp";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const OUT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "assets", "icons");
const SIZES = [16, 32, 80];

// All designs drawn on a 32x32 canvas.
const page = (fill = "#fff", stroke = "#5b6b7f") =>
  `<path d="M7 3h12l6 6v20H7z" fill="${fill}" stroke="${stroke}" stroke-width="1.6" stroke-linejoin="round"/>
   <path d="M19 3v6h6" fill="none" stroke="${stroke}" stroke-width="1.6" stroke-linejoin="round"/>`;
const lines = (ys, x1 = 10, x2 = 22, c = "#9aa7b5") =>
  ys.map((y) => `<path d="M${x1} ${y}H${x2}" stroke="${c}" stroke-width="1.8" stroke-linecap="round"/>`).join("");
const text = (s, x, y, size, color, weight = 700) =>
  `<text x="${x}" y="${y}" font-family="Segoe UI, Arial, sans-serif" font-size="${size}" font-weight="${weight}" fill="${color}" text-anchor="middle">${s}</text>`;

const ICONS = {
  // Add-in + group/menu: exam sheet with check marks
  addin: `<rect x="2" y="2" width="28" height="28" rx="6" fill="#2b6cb0"/>
          ${text("Đ", 16, 23, 19, "#fff")}`,
  dgnl: `${page("#eef4fb", "#2b6cb0")}
         <path d="M10 13l2 2 3.5-3.5" fill="none" stroke="#2f9e44" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
         ${lines([13.5], 17, 22, "#2b6cb0")}
         <path d="M10 20l2 2 3.5-3.5" fill="none" stroke="#2f9e44" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
         ${lines([20.5], 17, 22, "#2b6cb0")}`,
  chuanHoaToan: `${page()}${text("∑", 16, 25, 17, "#7048e8")}`,
  fixTabNumbering: `${page()}
         ${text("1.", 12, 15, 8, "#2f9e44")}${lines([12.5], 16, 22)}
         ${text("2.", 12, 23, 8, "#2f9e44")}${lines([20.5], 16, 22)}`,
  gachDapAnNguVan: `${page()}${text("A", 16, 22, 15, "#e03131")}
         <path d="M10 25.5H22" stroke="#e03131" stroke-width="2.2" stroke-linecap="round"/>`,
  removeBreakEng: `${page()}${lines([10], 10, 17)}
         <path d="M9 17h3M14 17h4M20 17h3" stroke="#f08c00" stroke-width="2" stroke-linecap="round"/>
         ${text("En", 16, 27, 8, "#f08c00")}`,
  xoaDactDgnl: `${page()}${lines([10, 14])}
         <circle cx="21" cy="22" r="6.5" fill="#e03131"/>
         <path d="M18.4 19.4l5.2 5.2M23.6 19.4l-5.2 5.2" stroke="#fff" stroke-width="2" stroke-linecap="round"/>`,
  // Author / about: person badge
  author: `<circle cx="16" cy="16" r="14" fill="#2b6cb0"/>
         <circle cx="16" cy="12.5" r="5" fill="#fff"/>
         <path d="M7.5 25.5c1.6-4.6 5-6.6 8.5-6.6s6.9 2 8.5 6.6" fill="#fff"/>`,
  // Check update: circular arrows
  update: `<circle cx="16" cy="16" r="14" fill="#2f9e44"/>
         <path d="M22.5 13.5A7 7 0 0 0 10 12" fill="none" stroke="#fff" stroke-width="2.4" stroke-linecap="round"/>
         <path d="M9.5 18.5A7 7 0 0 0 22 20" fill="none" stroke="#fff" stroke-width="2.4" stroke-linecap="round"/>
         <path d="M23.5 8.5v5.5H18" fill="none" stroke="#fff" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>
         <path d="M8.5 23.5V18H14" fill="none" stroke="#fff" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>`,
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
