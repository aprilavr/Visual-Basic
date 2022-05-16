Sub checkMatchingParentheses(m As Long, strArr2() As String)
    'check for matching sets of opening/closing parentheses and brackets and identify location of unmatched
    Dim OpenChar, CloseChar, paragrIP As String
    Dim errMsg As String
    Dim CheckP() As Boolean
    Dim paragrCounter, i, j As Integer
    Dim leftParenthCounter, rightParenthCounter As Integer

    paragrCounter = ActiveDocument.Paragraphs.Count
    leftParenthCounter = 0
    rightParenthCounter = 0
    ReDim CheckP(paragrCounter)

    i = 0
    While i < 2
        If i = 0 Then
            OpenChar = "("
            CloseChar = ")"
        Else
            OpenChar = "["
            CloseChar = "]"
        End If

        For j = 1 To paragrCounter
            CheckP(j) = False
            paragrIP = ActiveDocument.Paragraphs(j).Range.Text
            If Len(paragrIP) <> 0 Then
                leftParenthCounter = CountChars(paragrIP, OpenChar)
                rightParenthCounter = CountChars(paragrIP, CloseChar)
                If leftParenthCounter <> rightParenthCounter Then CheckP(j) = True
            End If
        Next j

        For j = paragrCounter To 1 Step -1
            If CheckP(j) Then
                If i = 0 Then
                    errMsg = "Uneven number of parentheses in paragraph number " & j
                Else
                    errMsg = "Uneven number of brackets in paragraph number " & j
                End If
                Call addToMissingItems(m, strArr2(), errMsg)
            End If
        Next j
        i = i + 1
    Wend
End Sub

Private Function CountChars(A As String, ByVal B As String) As Integer
    Dim Count, Found As Integer

    Count = 0
    Found = InStr(A, B)
    While Found <> 0
        Count = Count + 1
        Found = InStr(Found + 1, A, B)
    Wend
    CountChars = Count
End Function