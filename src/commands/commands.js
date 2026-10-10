// Entry của FunctionFile: đăng ký các ribbon command (ExecuteFunction).
// Tên action phải khớp <FunctionName> trong manifest/manifest.template.xml.

import { makeCommand } from "../core/runCommand.js";
import { chuanHoaToan } from "../functions/dgnl/chuanHoaToan.js";
import { fixTabNumbering } from "../functions/dgnl/fixTabNumbering.js";
import { gachDapAnNguVan } from "../functions/dgnl/gachDapAnNguVan.js";
import { removeBreakEng } from "../functions/dgnl/removeBreakEng.js";
import { xoaDactDgnl } from "../functions/dgnl/xoaDactDgnl.js";

const COMMANDS = {
  chuanHoaToan: makeCommand("Chuẩn hoá Toán ĐGNL", chuanHoaToan),
  fixTabNumbering: makeCommand("Fix lỗi tab câu hỏi", fixTabNumbering),
  gachDapAnNguVan: makeCommand("Gạch chân đáp án Ngữ Văn", gachDapAnNguVan),
  removeBreakEng: makeCommand("Fix lỗi break dòng T.Anh", removeBreakEng),
  xoaDactDgnl: makeCommand("Xoá ĐACT ĐGNL", xoaDactDgnl),
};

Office.onReady(() => {
  for (const [name, fn] of Object.entries(COMMANDS)) Office.actions.associate(name, fn);
});
