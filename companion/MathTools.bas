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

' menu MathType Color: tag của nút = tên màu TeX (black = bỏ màu)
Public Sub MS_MathColor(ByVal control As IRibbonControl)
    MathColor control.Tag
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

' ---------------------------------------------------------------- MathType Color
' Port từ menu MathType Color của BigKey (w_MathTypeColor.bas, Trương Minh Sang). Mỗi công thức: Toggle TeX ->
' bỏ mọi \color{…} -> chèn \color{màu} sau dấu mở ($, $$, \[, \() -> Toggle TeX lại. Đen = chỉ bỏ màu.
' Phạm vi: các MathType trong vùng chọn; vùng chọn không có MathType (hoặc chưa chọn) -> cả tài liệu.
' Bản cũ: Find/Replace trên CẢ tài liệu (tô màu cả $…$ chưa là MathType, cả khi chỉ chọn một công thức),
' không bỏ \color{white}, đổi \[…\] thành $…$, đếm mọi đối tượng OLE là công thức, gọi UndoClear.
Public Sub MathColor(ByVal mau As String)
    Dim n As Long, t As Single
    t = Timer
    If mau = "black" Then mau = ""
    Application.ScreenUpdating = False
    If Selection.Type <> wdSelectionIP Then n = MathColorRange(Selection.Range, mau)
    If n = 0 Then
        n = MathColorRange(ActiveDocument.Content, mau)
        Application.ScreenUpdating = True
        If n = 0 Then
            MsgBox "Không có công thức MathType nào để đổi màu.", vbExclamation, APP_TITLE
        ElseIf n > 0 Then
            Selection.HomeKey wdStory
            MsgBox "Đã đổi màu " & n & " công thức MathType (" & Format(Timer - t, "0.0") & " giây).", vbInformation, APP_TITLE
        End If
    End If
    Application.ScreenUpdating = True
End Sub

' Đổi màu mọi MathType trong r; trả về số công thức, -1 nếu không gọi được MathType.
' Không có MsgBox (trừ lỗi MathType) -> test gọi thẳng qua Application.Run.
Public Function MathColorRange(ByVal r As Range, ByVal mau As String) As Long
    Dim eq As New Collection, s As InlineShape, i As Long, t As String, p As Long
    For Each s In r.InlineShapes
        If IsMathType(s) Then eq.Add s
    Next
    For i = eq.Count To 1 Step -1   ' từ cuối lên: các công thức phía trước không bị dời
        eq(i).Select
        If Not RunMathType("MTCommand_TeXToggle") Then MathColorRange = -1: Exit Function
        ' sau Toggle TeX, Selection là đúng đoạn TeX vừa tạo
        p = Selection.Start
        t = ColorTeX(Selection.Text, mau)
        Selection.Range.Text = t
        ActiveDocument.Range(p, p + Len(t)).Select
        If Not RunMathType("MTCommand_TeXToggle") Then MathColorRange = -1: Exit Function
        MathColorRange = MathColorRange + 1
    Next
End Function

' TeX của một công thức: bỏ mọi \color{…} / \color[model]{…}, rồi chèn \color{mau} sau dấu mở.
Public Function ColorTeX(ByVal t As String, ByVal mau As String) As String
    Dim i As Long, j As Long, k As Long
    i = InStr(t, "\color")
    Do While i > 0
        j = i + 6
        If Mid$(t, j, 1) = "[" Then j = InStr(j, t, "]") + 1
        k = 0
        If j > 1 And Mid$(t, j, 1) = "{" Then k = InStr(j, t, "}")
        If k > 0 Then t = Left$(t, i - 1) & Mid$(t, k + 1) Else i = i + 1
        i = InStr(i, t, "\color")
    Loop
    If mau = "" Then
    ElseIf Left$(t, 2) = "$$" Or Left$(t, 2) = "\[" Or Left$(t, 2) = "\(" Then
        t = Left$(t, 2) & "\color{" & mau & "}" & Mid$(t, 3)
    ElseIf Left$(t, 1) = "$" Then
        t = "$\color{" & mau & "}" & Mid$(t, 2)
    End If
    ColorTeX = t
End Function

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
