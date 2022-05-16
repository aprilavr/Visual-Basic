Function createErrorDocument(strArr2() As String, errArr() As String, e As Long, m As Long)
'compile all missing items and errors into text document for user reference
    Dim ThisDoc, ThatDoc As Document
    Dim r As Range
    Dim headerText As String
    Dim i As Integer
    
    Set ThisDoc = ActiveDocument
    Set r = ThisDoc.Range
    Set ThatDoc = Documents.Add
    
    If m > 0 Then ThatDoc.Range.InsertAfter "Missing items: " & vbCrLf & Chr(149) & Join(strArr2, vbCrLf & Chr(149) & " ") & vbCrLf
        
    If e > 0 Then ThatDoc.Range.InsertAfter "Errors: " & vbCrLf & Chr(149) & Join(errArr, vbCrLf & Chr(149) & " ")
    
    'format text for readability
    i = 0
    While i < 2
        If i = 0 Then
            headerText = "Missing items:"
        Else
            headerText = "Errors:"
        End If
        
        With ThatDoc.Range.Find
            .Text = headerText
            .Replacement.Font.Bold = True
            .Replacement.Font.AllCaps = True
            .Replacement.Text = "^&"
            .Format = True
            .Execute Replace:=wdReplaceAll
        End With
        
        i = i + 1
    Wend
    
    
        
    
End Function