'captures user input from Access form to be used for generating notification emails based on applicability
Private Sub sendEmails_Click()
'capture SAIB info and applicability from user input
    Dim filePath As String
    Dim SAIBsubject As String
    Dim SAIBmodels As String
    Dim applicability As String
    Dim applicArray As Variant
    Dim i, applicCounter As Integer
    Dim saibNumber As String
    
    
    saibNumber = numberSAIB
    SAIBsubject = subject
    SAIBmodels = modelList
    
    applicArray = Array(largeAirplane, smallAirplane, rotorcraft, smallLarge)
    
    i = 0
    applicCounter = 0
    
    While i < 4
        If applicArray(i) = True Then
            applicCounter = applicCounter + 1
        End If
        i = i + 1
    Wend
    'capture applicability from user input
    If applicCounter = 0 Or applicCounter = 4 Then
        applicability = "All"
    Else
        If largeAirplane = True Then applicability = "Large"
        If smallAirplane = True Then applicability = applicability + ", Small"
        If rotorcraft = True Then applicability = applicability + ", Rotor"
        If smallLarge = True Then applicability = applicability + ", SmallLarge"
    End If
    
    Call sendSAIBemails(saibNumber, SAIBsubject, SAIBmodels, applicability)
    
End Sub

Private Sub openEmails_Click()
'opens db table containing email addresses to be edited by user
    DoCmd.OpenForm "emailTable", acFormDS
        
End Sub
