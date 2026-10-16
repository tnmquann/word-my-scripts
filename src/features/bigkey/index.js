// Nhóm Chuẩn hoá BigKey (port từ BigKeyUI.dotm → nhóm "Chuẩn hóa").
// Key = <FunctionName> trong manifest; label = tiêu đề hộp thoại kết quả.
// Các công cụ hỗ trợ (ID, phương án, đáp án, lời giải, số thứ tự, hình ảnh) chạy trong task pane
// src/ui/bigkey (gọi thẳng tools.js) — Office JS không có menu lồng nhau như BigKey.
import { chuanHoaBigKeyDe, chuanHoaBigKeyBoiDen } from "./chuanHoaBigKey.js";
import { checkLoiDe } from "./checkLoiDe.js";

export default {
  chuanHoaBigKeyDe: { label: "Chuẩn hoá BigKey", run: chuanHoaBigKeyDe },
  chuanHoaBigKeyBoiDen: { label: "Chuẩn hoá BigKey (vùng chọn)", run: chuanHoaBigKeyBoiDen },
  checkLoiDe: { label: "Check lỗi đề", run: checkLoiDe },
};
