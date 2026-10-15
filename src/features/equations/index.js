// Nhóm Equations: chuyển LaTeX <-> equation Word (ribbon MY_SCRIPTS → Equations).
// Mỗi nút hỏi phạm vi bằng popup (Vùng chọn / Toàn bộ tài liệu / Huỷ).
// Key = <FunctionName> trong manifest; label = tiêu đề hộp thoại kết quả.
import { latexToEquation, equationToLatex } from "./latexEquations.js";

export default {
  latexToEquation: { label: "LaTeX → Equation", run: latexToEquation },
  equationToLatex: { label: "Equation → LaTeX", run: equationToLatex },
};
