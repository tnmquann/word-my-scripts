// Kiểm tra cập nhật.
// - Phiên bản manifest đang cài: manifest gắn ?mv=<Version> vào URL của FunctionFile (commands.html).
// - Phiên bản mới nhất: version.json trên GitHub Pages (không cache).
// - Cài đặt: mở giao thức myscripts:update -> updater.ps1 (do install.ps1 đăng ký) hỏi user rồi cài.

const UPDATE_URI = "myscripts:update";
let cached = null;

const parse = (v) => String(v || "0").split(".").map((n) => parseInt(n, 10) || 0);
function newer(a, b) {
  const x = parse(a);
  const y = parse(b);
  for (let i = 0; i < Math.max(x.length, y.length); i++) {
    if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0);
  }
  return false;
}

/** { installed, latest, available } — latest = null nếu không kiểm tra được */
export async function getUpdateInfo(force = false) {
  if (cached && !force) return cached;
  const mv = new URLSearchParams(location.search).get("mv") || "";
  const installed = mv.split(".").slice(0, 3).join(".") || null;
  let latest = null;
  try {
    const res = await fetch(new URL(`version.json?t=${Date.now()}`, location.href).toString(), { cache: "no-store" });
    if (res.ok) latest = (await res.json()).version || null;
  } catch (e) {
    console.warn("update check failed", e);
  }
  cached = { installed, latest, available: !!(installed && latest && newer(latest, installed)) };
  return cached;
}

/** Mở trình cập nhật (updater.ps1 qua giao thức myscripts:). Trả về false nếu không mở được. */
export function launchUpdater() {
  try {
    if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
      Office.context.ui.openBrowserWindow(UPDATE_URI);
    } else {
      window.open(UPDATE_URI);
    }
    return true;
  } catch (e) {
    console.warn("launchUpdater failed", e);
    return false;
  }
}

/** Command "Kiểm tra cập nhật" */
export async function checkUpdate() {
  const info = await getUpdateInfo(true);
  if (!info.latest) throw new Error("Không kiểm tra được bản cập nhật (mất mạng?).");
  if (!info.installed) return `Phiên bản mới nhất: v${info.latest}.`;
  return info.available
    ? `Có bản mới: v${info.latest} (đang dùng v${info.installed}).\nBấm "Cập nhật" để cài.`
    : `Bạn đang dùng bản mới nhất (v${info.installed}).`;
}
