Private Sub sendEmails_Click()
'capture SAIB info and applicability from user input
    Dim filePath As String
    Dim SAIBsubject As String
    Dim applicability As String
    Dim SAIBmake As String
    Dim SAIBmodel As String
    Dim i, r, j As Integer
    Dim saibNumber As String
    Dim saibDate As Date
    Dim APIarray
    Dim documents As Object
    Dim summary As Object
    Dim makeCount, modelCount, subtypeCount As Integer
    

    
    saibNumber = numberSAIB
    saibDate = issueDate
    If saibNumber = "" Or saibDate = Null Then
        MsgBox ("Please complete all fields.")
        Exit Sub
    End If

    
    'Call saibAPIcall(saibNumber, saibDate)
    'APIarray(0) > documents
    'APIarray(1) > summary
    APIarray = saibAPIcall(saibNumber, saibDate)
    Set summary = APIarray(1)
    Set documents = APIarray(0)
    Debug.Print (summary.count)
    For i = 1 To summary("count")
        If documents(i)("drs:documentNumber") = saibNumber Then
           Debug.Print documents(i)("drs:documentNumber")
           'get subject
           SAIBsubject = documents(i)("drs:title")
           'get make(s)
           makeCount = documents(i)("drs:saibMake").count
           If makeCount = 1 Then
                SAIBmake = documents(i)("drs:saibMake")(1)
            ElseIf makeCount > 1 Then
                For r = 1 To makeCount
                    If r < makeCount Then
                        SAIBmake = SAIBmake + documents(i)("drs:saibMake")(r) + ", "
                    Else
                        SAIBmake = SAIBmake + documents(i)("drs:saibMake")(r)
                    End If
                Next r
            Else
                '
            End If

           'get model(s)
           modelCount = documents(i)("drs:saibModel").count
           If modelCount = 1 Then
                SAIBmodel = documents(i)("drs:saibModel")(1)
            ElseIf modelCount > 1 Then
                For r = 1 To modelCount
                    If r < modelCount Then
                        SAIBmodel = SAIBmodel + documents(i)("drs:saibModel")(r) + ", "
                    Else
                        SAIBmodel = SAIBmodel + documents(i)("drs:saibModel")(r)
                    End If
                Next r
            Else
                '
            End If
           'get product subtype
           subtypeCount = documents(i)("drs:productSubType").count
           If subtypeCount = 0 Then
                applicability = "All"
            ElseIf subtypeCount = 1 Then
                applicability = documents(i)("drs:productSubType")(1)
            ElseIf subtypeCount > 1 Then
                If subtypeCount > 3 Then
                    applicability = "All"
                Else
                    For r = 1 To subtypeCount
                        If r < subtypeCount Then
                            applicability = applicability + documents(i)("drs:productSubType")(r) + ", "
                        Else
                            applicability = applicability + documents(i)("drs:productSubType")(r)
                        End If
                    Next r
                End If
            Else
                '
            End If
        End If
    Next i
    
    
    
    Call sendSAIBemails(saibNumber, SAIBsubject, SAIBmake, SAIBmodel, applicability, saibDate)
    
End Sub

Function saibAPIcall(saibNum As String, inputDate As Date) As Variant
    Dim todayDate As String
    Dim apiURL As String
    Dim json As Object
    Dim documents As Object
    Dim summary As Object
    Dim count, r As Integer
    Dim APIarray
    
    todayDate = Format(inputDate, "YYYY-MM-DD")
    'generate URL
    apiURL = "https://drs.faa.gov/api/drs/data-pull/SAIB?docLastModifiedDate=" & todayDate & "T00:00:00.000Z"
    'curl --location --request GET
    Set HTTPReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTPReq.Open "GET", apiURL, False
    'API key sanitized for portfolio use
    HTTPReq.setRequestHeader "x-api-key", ""
    HTTPReq.Send
    Set json = JsonConverter.ParseJson(HTTPReq.responseText)
    Set documents = json("documents")
    Set summary = json("summary")

    APIarray = Array(documents, summary)

    saibAPIcall = APIarray
    

End Function



Private Sub openEmails_Click()
'opens db table containing email addresses to be edited by user
    DoCmd.OpenForm "emailTable", acFormDS
        
End Sub

Private Sub clearButton_Click()
'clear form data on click of clear button
    numberSAIB = ""
    subject = ""
    modelList = ""
    largeAirplane = False
    smallAirplane = False
    rotorcraft = False
    smallLarge = False
    issueDate = Null
    
End Sub

'Created 042522 by April VanRavenswaay, Lib_email modules are legacy code but needed for SendMail function to send emails
Sub sendSAIBemails(saibNumber As String, SAIBsubject As String, SAIBmakes As String, SAIBmodels As String, applic As String, saibDate As Date)
    'for testing
    Dim subject As String
    Dim body As String
    Dim filePath As String
    Dim sendTo As String
    Dim smallLarge As Boolean
    Dim rs As DAO.Recordset
        
    smallLarge = False
    

'get emails from table and build SQL statement based on applicability
    sql = "SELECT EmailAddress FROM emailTable WHERE EmailAddress Is Not Null"
    If applic = "All" Then
    Else
        sql = sql & " AND Applicability='All'"
        If InStr(applic, "Small/Large") Then
            smallLarge = True
            sql = sql & " OR Applicability='Small Airplane' OR Applicability='Large Airplane'"  'Send to All, Large, and Small
        End If
        If InStr(applic, "Large") Then
            If smallLarge = False Then sql = sql & " OR Applicability='Large Airplane'" 'Send to Large and All"
        End If
        If InStr(applic, "Small") Then
            If smallLarge = False Then sql = sql & " OR Applicability='Small Airplane'"
        End If
        If InStr(applic, "Rotor") Then sql = sql & " OR Applicability='Rotorcraft'"
    End If
    
    sql = sql & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    
'Email Addresses
   sendTo = ""
   Do Until rs.EOF
      sendTo = sendTo & rs.Fields(0) & ", "
      rs.MoveNext
   Loop
   rs.Close
   Set rs = Nothing
   
   subject = "SAIB " & saibNumber
   
      'Body
   body = "On " & saibDate & ", we published SAIB " & saibNumber & " to http://drs.faa.gov. " & "It is also attached to this email." & _
      vbCrLf & vbCrLf & "Subject:" & _
      vbCrLf & SAIBsubject & vbCrLf & vbCrLf & "Applicability:" & _
      vbCrLf & "Makes: " & SAIBmakes & _
      vbCrLf & "Models: " & SAIBmodels
   'filePath sanitized for portfolio use
   filePath = "" + saibNumber + "\" + saibNumber + ".pdf"
    'validation that file exists
    If Dir(filePath) = "" Then
        MsgBox (saibNumber & ".pdf not found. Please check the directory.")
        Exit Sub
    End If
    
    
    SendMail subject, body, sendTo, , , filePath, True
End Sub

