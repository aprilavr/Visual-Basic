Option Compare Database
'Created 042522 by April VanRavenswaay, Lib_email modules are legacy code but needed for SendMail function to send emails
Sub sendSAIBemails(saibNumber As String, SAIBsubject As String, SAIBmodels As String, applic As String)
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
        If InStr(applic, "SmallLarge") Then
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
   body = "On " & Date & ", we published SAIB " & saibNumber & " to http://rgl.faa.gov and http://drs.faa.gov. " & "It is also attached to this email." & _
      vbCrLf & vbCrLf & "Subject:" & _
      vbCrLf & SAIBsubject & vbCrLf & vbCrLf & "Applicability:" & _
      vbCrLf & SAIBmodels
   
'sanitized for use in portfolio
   filePath = "insert file path here" + saibNumber + ".pdf"
        
    SendMail subject, body, sendTo, , , filePath, True
End Sub