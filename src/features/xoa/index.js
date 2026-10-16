// Nhóm Xoá cần thiết (port từ NHOM-WORD-BIEN-SOAN.dotm). Các lệnh của form "Xóa mức độ"
// chạy trong task pane src/ui/xoa (gọi thẳng các hàm trong xoaCanThiet.js).
// Key = <FunctionName> trong manifest; label = tiêu đề hộp thoại kết quả.
import { xoaTrichNguon, xoaLoiGiai, xoaTacGia } from "./xoaCanThiet.js";

export default {
  xoaTrichNguon: { label: "Xoá trích nguồn", run: xoaTrichNguon },
  xoaLoiGiai: { label: "Xoá lời giải", run: xoaLoiGiai },
  xoaTacGia: { label: "Xoá tác giả", run: xoaTacGia },
};
