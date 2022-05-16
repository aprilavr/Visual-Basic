'Functions to add missing items and errors to arrays for output to user via CreateErrorDocument
Function addToErrorList(e As Long, errorArr() As String, errMsg As String, r As Range)
    'add errors to array for later output
    'capture selection to be highlighted and page/line number
    Dim errorText As String
    errorText = "Error Found: " & errMsg & vbCrLf & "Location: Page " & r.Information(wdActiveEndPageNumber) & " Line " & r.Information(wdFirstCharacterLineNumber)
    ReDim Preserve errorArr(e)
    errorArr(e) = errorText
    e = e + 1
    errorText = " "
End Function

Function addToMissingItems(m As Long, strArr2() As String, msngMsg As String)
    'add missing items to array for later output
    ReDim Preserve strArr2(m)
    strArr2(m) = msngMsg
    m = m + 1
End Function