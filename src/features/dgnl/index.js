// Nhóm ĐGNL (V-ACT): các command trên ribbon MY_SCRIPTS → ĐGNL.
// Key = <FunctionName> trong manifest; label = tiêu đề hộp thoại kết quả.
import { chuanHoaToan } from "./chuanHoaToan.js";
import { fixTabNumbering } from "./fixTabNumbering.js";
import { gachDapAnNguVan } from "./gachDapAnNguVan.js";
import { removeBreakEng } from "./removeBreakEng.js";
import { xoaDactDgnl } from "./xoaDactDgnl.js";

export default {
  chuanHoaToan: { label: "Chuẩn hoá Toán ĐGNL", run: chuanHoaToan },
  fixTabNumbering: { label: "Fix lỗi tab câu hỏi", run: fixTabNumbering },
  gachDapAnNguVan: { label: "Gạch chân đáp án Ngữ Văn", run: gachDapAnNguVan },
  removeBreakEng: { label: "Fix lỗi break dòng T.Anh", run: removeBreakEng },
  xoaDactDgnl: { label: "Xoá ĐACT ĐGNL", run: xoaDactDgnl },
};
