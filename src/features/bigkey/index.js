// Nhóm Chuẩn hoá BigKey (port từ BigKeyUI.dotm → nhóm "Chuẩn hóa").
// Key = <FunctionName> trong manifest; label = tiêu đề hộp thoại kết quả.
import { chuanHoaBigKeyDe, chuanHoaBigKeyBoiDen } from "./chuanHoaBigKey.js";
import { checkLoiDe } from "./checkLoiDe.js";

export default {
  chuanHoaBigKeyDe: { label: "Chuẩn hoá BigKey", run: chuanHoaBigKeyDe },
  chuanHoaBigKeyBoiDen: { label: "Chuẩn hoá BigKey (vùng chọn)", run: chuanHoaBigKeyBoiDen },
  checkLoiDe: { label: "Check lỗi đề", run: checkLoiDe },
};
