// Nạp bundle lệnh (features-<hash>.js) và tự cập nhật khi có bản mới — không cần khởi động lại Word.
//
// Word giữ FunctionFile (commands.html) trong bộ nhớ suốt phiên làm việc, nên trước đây bản vá
// chỉ có hiệu lực sau khi mở lại Word. Giờ mỗi lần bấm nút: đọc version.json (no-store) -> nếu
// tên bundle features khác bản đang chạy thì nạp bundle mới, rồi mới chạy lệnh.

const CHECK_TIMEOUT_MS = 3000;
let loadedFile = null; // tên bundle features đang chạy
let inflight = null;

function loadScript(file) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = new URL(file, location.href).toString();
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Không tải được ${file}`));
    document.head.appendChild(s);
  });
}

async function latestFeaturesFile() {
  const ctrl = typeof AbortController !== "undefined" ? new AbortController() : null;
  const timer = ctrl ? setTimeout(() => ctrl.abort(), CHECK_TIMEOUT_MS) : null;
  try {
    const res = await fetch(new URL(`version.json?t=${Date.now()}`, location.href).toString(), {
      cache: "no-store",
      signal: ctrl ? ctrl.signal : undefined,
    });
    if (!res.ok) return null;
    const info = await res.json();
    return info.features || null;
  } catch (e) {
    console.warn("features check failed (offline?) — using loaded bundle", e);
    return null;
  } finally {
    if (timer) clearTimeout(timer);
  }
}

/** Ghi nhận bundle đã được commands.html nạp sẵn (<meta name="features">). */
export function initFeatures(initialFile) {
  loadedFile = initialFile;
}

/** Trước mỗi lệnh: nạp bundle mới nếu đã deploy bản mới. Trả về registry groups. */
export async function ensureLatestFeatures() {
  if (!inflight) {
    inflight = (async () => {
      const latest = await latestFeaturesFile();
      if (latest && latest !== loadedFile) {
        try {
          await loadScript(latest);
          console.info(`features updated: ${loadedFile} -> ${latest}`);
          loadedFile = latest;
        } catch (e) {
          console.warn(e);
        }
      }
    })().finally(() => {
      inflight = null;
    });
  }
  await inflight;
  return (window.__MY_SCRIPTS_FEATURES__ || { groups: {} }).groups;
}
