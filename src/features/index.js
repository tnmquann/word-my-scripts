// Bundle "features": toàn bộ các nhóm lệnh. Build riêng (features-<hash>.js) để runtime có thể
// nạp bản mới mỗi khi bấm nút mà không cần khởi động lại Word (xem src/runtime/loader.js).
// Thêm nhóm mới: import rồi thêm vào GROUPS.
import dgnl from "./dgnl/index.js";
import ctpt2018 from "./ctpt2018/index.js";
import equations from "./equations/index.js";
import bigkey from "./bigkey/index.js";
import xoa from "./xoa/index.js";

const GROUPS = { dgnl, ctpt2018, equations, bigkey, xoa };

// Bundle mới nhất ghi đè bản trước; runtime luôn đọc window.__MY_SCRIPTS_FEATURES__
window.__MY_SCRIPTS_FEATURES__ = { groups: GROUPS };
