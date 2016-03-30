Private Sub ForwardCell()
    If ActiveCell.MergeCells Then
        If ActiveCell.MergeArea.End(xlToRight).Column = Columns.Count Then Exit Sub
    End If
    If ActiveCell.Column <> Columns.Count Then ActiveCell.Offset(0, 1).Activate
End Sub

Private Sub BackwardCell()
    If ActiveCell.Column <> 1 Then ActiveCell.Offset(0, -1).Activate
End Sub

Private Sub PrevLineCell()
    If ActiveCell.Row <> 1 Then ActiveCell.Offset(-1, 0).Activate
End Sub

Private Sub NextLineCell()
    If ActiveCell.MergeCells Then
        If ActiveCell.MergeArea.End(xlDown).Row = Rows.Count Then Exit Sub
    End If
    If ActiveCell.Row <> Rows.Count Then ActiveCell.Offset(1, 0).Activate
End Sub

Private Sub NextSheet()
    If ActiveSheet.Name = Sheets(Sheets.Count).Name Then
        MsgBox "last sheet"
    Else
        ActiveSheet.Next.Activate
    End If
End Sub

Private Sub PrevSheet()
    If ActiveSheet.Name = Sheets(1).Name Then
         MsgBox "first sheet"
     Else
         ActiveSheet.Previous.Activate
     End If
End Sub

Private Sub GoToLine()
    Dim LineNum As Long
    On Error GoTo endSub
    LineNum = CLng(Application.InputBox("GoToLineNumber"))
    If LineNum > 0 Then
        ActiveWindow.ScrollRow = LineNum
        Cells(LineNum, Selection(1).Column).Select
    End If
endSub:
End Sub

Private Sub ScrollDown()
    ScrollDownAction (40)
End Sub

Private Sub ScrollUp()
    ScrollUpAction (40)
End Sub

Private Sub ScrollHalfDown()
    ScrollDownAction (20)
End Sub

Private Sub ScrollHalfUp()
    ScrollUpAction (20)
End Sub

Private Sub ScrollDownAction(ByVal MoveRowNum As Long)
    CurrentRowNum = ActiveCell.Row
    MaxRow = Rows.Count
    If MaxRow < (CurrentRowNum + MoveRowNum) Then
        ActiveWindow.ScrollRow = MaxRow
        Cells(MaxRow, Selection(1).Column).Select
    Else
        ActiveCell.Offset(MoveRowNum, 0).Activate
    End If
End Sub

Private Sub ScrollUpAction(ByVal MoveRowNum As Long)
    CurrentRowNum = ActiveCell.Row
    MinRow = 1
    If MinRow > (CurrentRowNum - MoveRowNum) Then
        ActiveWindow.ScrollRow = MinRow
        Cells(MinRow, Selection(1).Column).Select
    Else
        ActiveCell.Offset(-MoveRowNum, 0).Activate
    End If
End Sub

Sub NormalMode()
    With Application
        ' add
        .OnKey "^{h}", "BackwardCell"
        .OnKey "^{j}", "NextLineCell"
        .OnKey "^{k}", "PrevLineCell"
        .OnKey "^{l}", "ForwardCell"
        .OnKey "^{TAB}", "NextSheet"
        .OnKey "+^{TAB}", "PrevSheet"
        .OnKey "^{g}", "GoToLine"
        .OnKey "^{f}", "ScrollDown"
        .OnKey "^{b}", "ScrollUp"
        .OnKey "^{u}", "ScrollHalfUp"
        .OnKey "^{d}", "ScrollHalfDown"
        .OnKey "^{i}", "InsertMode"
        ' delete
        .OnKey "+{ESC}"
    End With
End Sub

Sub InsertMode()
    With Application
        ' delete
        .OnKey "^{h}"
        .OnKey "^{j}"
        .OnKey "^{k}"
        .OnKey "^{l}"
        .OnKey "^{TAB}"
        .OnKey "+^{TAB}"
        .OnKey "^{g}"
        .OnKey "^{f}"
        .OnKey "^{b}"
        .OnKey "^{u}"
        .OnKey "^{d}"
        .OnKey "^{i}"
        ' add
        .OnKey "+{ESC}", "NormalMode"
    End With
End Sub

