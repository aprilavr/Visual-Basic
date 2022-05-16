Sub goToParagraph()
'used to navigate user to specific paragraph number, such as paragraph of missing parentheses as output by error macro
    Application.ScreenUpdating = False
Dim i As Long, j As Long, k As Long, Rng As Range
With ActiveDocument
  j = .Range.Paragraphs.Count
  i = InputBox("Enter paragraph number: ")
  If i < 1 Then
    k = 0
  ElseIf i > j Then
    MsgBox ("Paragraph not found, " & j & " paragraphs in this document.")
    Exit Sub
  Else
    k = i
  End If
  Set Rng = .Range(0, 0)
  With Rng
    .MoveEnd wdParagraph, k - 1
    .Collapse wdCollapseEnd
    If i < 1 Then
    ElseIf i > j Then
      .Start = ActiveDocument.Range.End
    Else
      .MoveEnd wdParagraph, 1
    End If
    .Select
  End With
End With
Application.ScreenUpdating = True
End Sub