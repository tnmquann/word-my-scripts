Attribute VB_Name = "MathTools"
' MY_SCRIPTS Desktop - Công cụ MathType (port từ S_MATH.bas của NHOM-WORD-BIEN-SOAN.dotm).
' Nguồn UTF-8: scripts/build-companion.ps1 đổi chuỗi tiếng Việt thành ChrW() trước khi nạp vào VBA.
' Spec: docs/specs/companion/MathType_Tools.md
Option Explicit

' (hàm, không phải Const: Const không chứa được ChrW() sau khi build)
Private Function APP_TITLE() As String
    APP_TITLE = "MY_SCRIPTS - Công cụ MathType"
End Function

' ---------------------------------------------------------------- ribbon callbacks

Public Sub MS_MathTypeHoa(ByVal control As IRibbonControl)
    MathTypeHoa
End Sub

Public Sub MS_LechMathType(ByVal control As IRibbonControl)
    LechMathType
End Sub

Public Sub MS_LoiEquation(ByVal control As IRibbonControl)
    LoiEquation
End Sub

Public Sub MS_LechPicMath(ByVal control As IRibbonControl)
    LechPicMath
End Sub

' ---------------------------------------------------------------- MathType hoá
' Bản cũ: xoá MỌI dấu cách trong vùng chọn (hỏng "\alpha x" -> "\alphax"), Cut rồi SendKeys vào cửa sổ MathType.
' Bản này: dùng lệnh Toggle TeX của MathType (MTCommand_TeXToggle) — đổi mọi $…$, $$…$$, \[…\], \(…\) trong
' vùng chọn thành đối tượng MathType; vùng chọn chưa có dấu $ thì tự bọc $…$. Không SendKeys -> chạy được trên Mac.
Public Sub MathTypeHoa()
    Dim r As Range, t As String
    If Selection.Type = wdSelectionIP Then
        If MsgBox("Chưa bôi đen công thức." & vbCr & vbCr & _
                  "Chuyển MỌI công thức LaTeX có dấu $…$ / \[…\] trong toàn bộ tài liệu thành MathType?", _
                  vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
        ActiveDocument.Content.Select
    Else
        Set r = Selection.Range
        ' bỏ dấu cách / xuống dòng ở hai đầu (không đụng dấu cách bên trong công thức)
        Do While r.End > r.Start And InStr(" " & vbTab & vbCr & Chr(11) & Chr(160), Right$(r.Text, 1)) > 0
            r.MoveEnd wdCharacter, -1
        Loop
        Do While r.End > r.Start And InStr(" " & vbTab & Chr(160), Left$(r.Text, 1)) > 0
            r.MoveStart wdCharacter, 1
        Loop
        If r.End <= r.Start Then Exit Sub
        t = r.Text
        If InStr(t, "$") = 0 And InStr(t, "\[") = 0 And InStr(t, "\(") = 0 Then
            r.InsertBefore "$"
            r.InsertAfter "$"
        End If
        r.Select
    End If
    If Not RunMathType("MTCommand_TeXToggle") Then Exit Sub
End Sub

' ---------------------------------------------------------------- Lệch MathType
' Bản cũ: Selection.ClearFormatting (mất cả đậm/màu/căn lề của vùng chọn) rồi mở Format Equations.
' Bản này: chỉ đưa vị trí nâng/hạ chữ (Font.Position) của các đối tượng MathType trong vùng chọn
' (hoặc cả tài liệu nếu chưa chọn) về 0, rồi mở hộp thoại Format Equations của MathType để định dạng lại.
Public Sub LechMathType()
    Dim r As Range, s As InlineShape
    If Selection.Type = wdSelectionIP Then Set r = ActiveDocument.Content Else Set r = Selection.Range
    For Each s In r.InlineShapes
        If IsMathType(s) Then
            If s.Range.Font.Position <> 0 Then s.Range.Font.Position = 0
        End If
    Next
    RunMathType "MTCommand_FormatEqns"
End Sub

' ---------------------------------------------------------------- Lỗi Equation
' Equation Word (OMML) hiển thị sai/lệch khi font không phải Cambria Math (thường do copy từ nguồn khác).
Public Sub LoiEquation()
    Dim m As OMath, n As Long
    For Each m In ActiveDocument.OMaths
        If m.Range.Font.Name <> "Cambria Math" Then m.Range.Font.Name = "Cambria Math": n = n + 1
    Next
    MsgBox "Đã đặt font Cambria Math cho " & n & " / " & ActiveDocument.OMaths.Count & " equation.", vbInformation, APP_TITLE
End Sub

' ---------------------------------------------------------------- Lệch Pic-Math
' Hạ ảnh công thức (hoặc MathType) cho thẳng dòng chữ: Position = -cao/2 + 3 (ảnh thấp) hoặc + 3,5.
' Bản cũ: lấy chiều cao ẢNH ĐẦU TIÊN rồi đặt lại toàn bộ font của cả vùng chọn (mất đậm, màu, cỡ chữ…).
' Bản này: chỉnh từng ảnh theo chiều cao của chính nó, không đụng định dạng chữ xung quanh.
Public Sub LechPicMath()
    Dim r As Range, s As InlineShape, n As Long, h As Single
    If Selection.Type = wdSelectionIP Then
        MsgBox "Hãy bôi đen đoạn có ảnh công thức cần chỉnh.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    Set r = Selection.Range
    For Each s In r.InlineShapes
        If s.Type = wdInlineShapePicture Or s.Type = wdInlineShapeLinkedPicture Or IsMathType(s) Then
            h = s.Height
            If h >= 14.4 Then s.Range.Font.Position = CLng(-h / 2 + 3.5) Else s.Range.Font.Position = CLng(-h / 2 + 3)
            n = n + 1
        End If
    Next
    If n = 0 Then MsgBox "Vùng chọn không có ảnh / công thức nào.", vbExclamation, APP_TITLE
End Sub

' ---------------------------------------------------------------- helpers

Private Function IsMathType(ByVal s As InlineShape) As Boolean
    On Error Resume Next
    If s.Type = wdInlineShapeEmbeddedOLEObject Then IsMathType = (InStr(1, s.OLEFormat.ProgID, "Equation.DSMT", vbTextCompare) = 1)
End Function

' Gọi macro của MathType (MathType Commands *.dotm). False + thông báo nếu MathType chưa cài / chưa nạp.
Private Function RunMathType(ByVal macro As String) As Boolean
    On Error GoTo NoMT
    Application.Run macro
    RunMathType = True
    Exit Function
NoMT:
    MsgBox "Không gọi được MathType (" & macro & ")." & vbCr & vbCr & _
           "Hãy cài MathType và bật tab MathType trong Word, rồi thử lại." & vbCr & _
           "Chi tiết: " & Err.Description, vbExclamation, APP_TITLE
End Function
