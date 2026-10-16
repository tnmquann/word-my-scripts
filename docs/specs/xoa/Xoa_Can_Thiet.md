# Xoá cần thiết

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → Xoá cần thiết → **Xoá trích nguồn** / **Xoá lời giải** / **Xoá tác giả** / **Xoá mức độ, khung…** (task pane) |
| Actions | `xoaTrichNguon`, `xoaLoiGiai`, `xoaTacGia`; pane: `xoaTheoMucDo`, `xoaID`, `xoaHuongSuyLuan`, `themDongCham`, `taoKhung`, `huyKhung` |
| Modules | `src/features/xoa/xoaCanThiet.js`; pane `src/ui/xoa/xoa.html` + `xoa.js` |
| Origin | *Xóa cần thiết* group of `NHOM-WORD-BIEN-SOAN.dotm`: `xoa_can_thiet.bas` (Xoa_trich_nguon, Xoa_loi_giai, xoa_tac_gia) and the *Xóa mức độ* form `HO_tro` |
| API sets | WordApi 1.3. *Tạo khung* / *Huỷ khung* need WordApiDesktop 1.3 (paragraph borders and shading); on Word on the web the buttons are disabled with a warning. |
| Destructive | Yes, except the frame tools. Ctrl+Z restores; run on a copy. |
| Idempotent | Yes. A second run finds nothing left to delete. |

## Purpose

Strip what a student version of an exam must not contain, and prepare it for printing.

## How blocks are found

The VBA used chains of Find/Replace with temporary marker characters (`#`, `~`, `#$`), which broke on documents that already contained those characters. This port works on paragraphs:

- A **question** is a paragraph outside a table whose text (list number included) starts with `Câu N.` or `Câu N:`.
- A **block** starts at its marker paragraph and ends before the next question or **boundary**: `CĐ1…`, `PHẦN …`, `BẢNG ĐÁP ÁN`, or a group lead-in (`Questions 46-52: …`, `Read the following…`, `Mark the letter…`, `Đọc đoạn văn…`, `Dựa vào…`, `Sử dụng các thông tin…`, `…trả lời các câu hỏi từ X đến Y`). Without the lead-in boundary, deleting the solution of the last question of a group would also delete the next reading passage.
- Tables inside a block are deleted with it; deletion runs from the end of the document (fast batched path; if it fails, the document is re-scanned and blocks are deleted one by one).

## Ribbon commands

| Command | Deletes |
| --- | --- |
| Xoá trích nguồn | The `[…]` / `(…)` tags right after `Câu N.` (source, ID), e.g. `Câu 1. [2D1-3] (Sở Hà Nội) …` → `Câu 1. …`. |
| Xoá lời giải | Every block from `Lời giải` / `Hướng dẫn giải` to the next question or boundary. |
| Xoá tác giả | Paragraphs that start with `Tác giả`, and inline groups `[(Tác giả …)]` / `(Tác giả …)`. |

## Task pane *Xoá mức độ, khung…*

Replaces the `HO_tro` UserForm (VBA forms cannot exist in Office JS).

| Control | Behaviour |
| --- | --- |
| Mức 1–4 + *Xoá các mức đã chọn* | Deletes whole questions whose ID ends in `-k]` (e.g. `[2D1-3]` is level 3), then renumbers typed `Câu N.` labels consecutively from the first question. |
| Xoá lời giải | Same as the ribbon command. |
| Xoá hướng suy luận | Deletes `Phân tích hướng dẫn giải` / `Phân tích lời giải` blocks, up to *Lời giải*. |
| Xoá mã ID | Deletes `[2D1-3]`, `[1H2-1]`, `[D1-2]`, `[H3-4]`-style tags anywhere. |
| Dòng chấm (n, default 7) | Deletes the solutions, then adds n dotted lines (grey, 1.5 line spacing) after each question; after the table when a question ends in one. |
| Tạo khung | Borders every question (stem and options, up to *Lời giải*) with one of the 8 styles of the form (Đôi, Chấm, Gạch ngắn, Gạch dài, Gạch chấm, Gạch chấm chấm, Sóng đôi, Gạch chấm nổi), border colour blue `#0000FF` or pink `#FF00FF`, fill pale yellow `#FAF0E2`, light green `#CCFFCC` or none, and a 1.75 cm hanging indent so consecutive paragraphs merge into one frame. |
| Huỷ khung | Removes those borders and the fill. |

Dropped from the VBA form: re-typing the option labels inside *Tạo khung* (use **Chuẩn hoá BigKey** for option layout).

## Test checklist

Run on a **copy** of a real exam.

- [ ] Xoá lời giải on an English exam with reading passages: every `Questions 46-52: …` lead-in and passage stays.
- [ ] Xoá lời giải on a Math exam with tables inside the solution: the tables are deleted, tables in the stem stay.
- [ ] Xoá trích nguồn: `Câu 1. [2D1-3] (Sở Hà Nội) Cho…` → `Câu 1. Cho…`; a stem that starts with `(x)` after text is unchanged.
- [ ] Xoá tác giả: a `Tác giả: …` line is deleted; `Tác giảng viên…` is not.
- [ ] Mức 3 + 4: those questions disappear and the rest are numbered without gaps.
- [ ] Dòng chấm 7: the solutions are replaced by 7 dotted lines under each question.
- [ ] Tạo khung (Word desktop): one frame per question with the chosen style, colour and fill; Huỷ khung removes it. On Word on the web the buttons are disabled with a warning.

Automated: `tests/xoa-helpers.mjs [export.json …]` (block finding, tags, author lines; with exports it checks that no solution block crosses a question or lead-in).
