VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScopeOfWork 
   Caption         =   "Scope of Work"
   ClientHeight    =   12960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "ScopeOfWork.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScopeOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strTFolder As String

Private Sub cbEmail_Click()
    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem
    Dim strTo, strCC, strSubject, strBody As String
    'Dim strVerb As String
    Dim strCompany, strPath As String
    'Dim vTemp As Variant
    Dim iNumber, iEmailSelf As Integer
    Dim iMTEMC, iOther, iDWG As Integer
    Dim iRevision As Integer
    
    On Error Resume Next
    
    Me.Hide
    
    strPath = ThisDrawing.Path
    strCompany = ""
    If InStr(UCase(strPath), "UNITED") Then strCompany = "UTC "
    If InStr(UCase(strPath), "LORETTO") Then strCompany = "Loretto "
    
    If Right(tbNumber.Value, 1) = "*" Then
        strSubject = "<Project Update> " & strCompany & tbNumber.Value & " - " & " Scope Revised"
        strBody = "<b>" & tbNumber.Value & " " & tbTitle.Value & "</b> " & " has been revised." & vbCr
    Else
        strSubject = "<Project Update> " & strCompany & tbNumber.Value & " - " & " Planning Review Complete"
        strBody = "<b>" & tbNumber.Value & " " & tbTitle.Value & "</b> " & " is ready for the next stage." & vbCr
        'strBody = strBody & vbCr & "# LU = UNKNOWN"
        tbNumber.Value = tbNumber.Value & "*"
    End If
    
    'strSubject = "<Project Update> " & strCompany & tbNumber.Value & " " & tbTitle.Value & " - " & " Planning Review Complete"
    strTo = "jon.wilburn@integrity-us.com;adam.kemper@integrity-us.com;rich.taylor@integrity-us.com"

    Load ProjectStatusEmail
    
    ProjectStatusEmail.Caption = "Planning"
    ProjectStatusEmail.tbBody.Value = Replace(strBody, "<br>", vbCr)
    ProjectStatusEmail.tbSubject.Value = strSubject
    ProjectStatusEmail.lbCC.AddItem "Wade.Hampton@integrity-us.com"
    ProjectStatusEmail.show
    
    strBody = Replace(ProjectStatusEmail.tbBody.Value, vbCr, "<br>")
    If strBody = "" Then
        MsgBox "Email Cancelled."
        Me.show
        Exit Sub
    End If
    
    If ProjectStatusEmail.cbCcSelf.Value = True Then
        iEmailSelf = 1
    Else
        iEmailSelf = 0
    End If
    
    For n = 0 To ProjectStatusEmail.lbCC.ListCount - 1
        'If ProjectStatusEmail.lbCC.Selected(n) = True Then
            strCC = strCC & ProjectStatusEmail.lbCC.List(n) & ";"
        'End If
    Next n
    
    Unload ProjectStatusEmail
    
    Set objOutlook = New Outlook.Application
    Set objMail = objOutlook.CreateItem(olMailItem)
    
    objMail.To = strTo
    If iEmailSelf = 1 Then strCC = strCC & objOutlook.Session.CurrentUser.Address
    objMail.CC = strCC  'objOutlook.Session.CurrentUser.Address
    objMail.Subject = strSubject
    objMail.HTMLBody = strBody
    
    'objMail.Display
    objMail.Send
    Me.show
End Sub

Private Sub cbFind_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If cbFind.Value = "" Then Exit Sub
    
    cbReplaceText.Clear
    cbReplaceText.MatchRequired = False
    
    Dim vLine, vItem As Variant
    
    For i = 0 To lbVariables.ListCount - 1
        If lbVariables.List(i, 0) = cbFind.Value Then
            If lbVariables.List(i, 0) = "<Empty>" Then Exit Sub
            GoTo Found_Line
        End If
    Next i
    
Found_Line:
    'vLine = Split(lbVariables.List(i, 1), ",")
    'If UBound(vLine) = 0 Then Exit Sub
    
    vItem = Split(lbVariables.List(i, 1), ",")
    For i = 0 To UBound(vItem)
        cbReplaceText.AddItem vItem(i)
    Next i
    
    If UBound(vItem) > 0 Then cbReplaceText.MatchRequired = True
End Sub

Private Sub cbJobType_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'If Not tbSOW.Value = "" Then Exit Sub
    If strTFolder = "" Then Exit Sub
    If cbJobType.Value = "" Then Exit Sub
    
    If Len(tbSOW.Value) > 1 Then
        Dim result As Integer
    
        result = MsgBox("Overwrite Notes and Scope?", vbYesNo, "Overwrite Existing")
        If result = vbNo Then
            Exit Sub
        End If
    End If
    
    Dim strPath, strLine As String
    Dim fName As String
    Dim iStatus As Integer
    Dim vLine, vTemp As Variant
    Dim vVar, vCond As Variant
    
    tbSpecial.Value = ""
    tbSOW.Value = ""
    
    lbVariables.Clear
    lbCond.Clear
    
    iStatus = 0
    strPath = Replace(strTFolder, "*.*", "") & cbJobType.Value & ".swt"
    
    fName = Dir(strPath)
    If fName = "" Then
        Exit Sub
    End If
    
    Open strPath For Input As #1
    
    While Not EOF(1)
        Line Input #1, strLine
        
        If InStr(strLine, ">>") > 0 And Not iStatus = 3 Then
            vLine = Split(strLine, ">>")
            For j = 0 To UBound(vLine) - 1
            
                vTemp = Split(vLine(j), "<<")
                vVar = Split(vTemp(1), "/")
            
                'If lbVariables.ListCount < 0 Then
                    For i = 0 To lbVariables.ListCount - 1
                        If lbVariables.List(i, 0) = vVar(0) Then GoTo Skip_Adding_Var
                    Next i
                'End If
            
                lbVariables.AddItem vVar(0)
                If UBound(vVar) > 0 Then
                    lbVariables.List(lbVariables.ListCount - 1, 1) = Replace(vVar(1), ">>", "")
                Else
                    lbVariables.List(lbVariables.ListCount - 1, 1) = "<Empty>"
                End If
        
Skip_Adding_Var:
            Next j
            
            If UBound(vLine) > 2 Then
                For i = 1 To UBound(vLine) - 1
                    vTemp = Split(vLine(i), "<<")
                    vVar = Split(vTemp(1), "/")
            
                    'For j = 0 To lbVariables.ListCount - 1
                        'If lbVariables.List(j, 0) = vVar(0) Then GoTo Skip_Adding_This
                    'Next j
            
                    lbVariables.AddItem vVar(0)
                    If UBound(vVar) > 0 Then
                        lbVariables.List(lbVariables.ListCount - 1, 1) = Replace(vVar(1), ">>", "")
                    Else
                        lbVariables.List(lbVariables.ListCount - 1, 1) = "<Empty>"
                    End If
Skip_Adding_This:
                    'cbFind.AddItem "<<" & vTemp(1) & ">>"
                Next i
            End If
        End If
        
        If Left(strLine, 2) = "[[" Then
            Select Case Left(strLine, 3)
                Case "[[I"
                    iStatus = 1
                Case "[[S"
                    iStatus = 2
                Case "[[C"
                    iStatus = 3
            End Select
        Else
            Select Case iStatus
                Case Is = 1
                    If tbSpecial.Value = "" Then
                        tbSpecial.Value = strLine
                    Else
                        tbSpecial.Value = tbSpecial.Value & vbCr & strLine
                    End If
                Case Is = 2
                    If tbSOW.Value = "" Then
                        tbSOW.Value = strLine
                    Else
                        tbSOW.Value = tbSOW.Value & vbCr & strLine
                    End If
                Case Is = 3
                    '<---------------------------------------------------------------------
                    vCond = Split(strLine, "=")
                    lbCond.AddItem vCond(0)
                    lbCond.List(lbCond.ListCount - 1, 1) = vCond(1)
                    lbCond.List(lbCond.ListCount - 1, 2) = vCond(2)
            End Select
        End If
    Wend
    
    Close #1
    
    For i = 0 To lbVariables.ListCount - 1
        cbFind.AddItem lbVariables.List(i, 0)
    Next i
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbReplaceAll_Click()
    If cbReplaceText.Value = "" Then Exit Sub
    
    cbReplaceText.MatchRequired = False
    
    Dim strTemp, strFind As String
    
    For i = 0 To lbVariables.ListCount - 1
        If lbVariables.List(i, 0) = cbFind.Value Then GoTo Found_Var
    Next i
    
Found_Var:
    
    strFind = "<<" & lbVariables.List(i, 0) '& "/" & lbVariables.List(i, 1) & ">>"
    If lbVariables.List(i, 1) = "<Empty>" Then
        strFind = strFind & ">>"
    Else
        strFind = strFind & "/" & lbVariables.List(i, 1) & ">>"
    End If
    
    strTemp = Replace(tbSpecial.Value, strFind, "<b>" & cbReplaceText.Value & "</b>")
    tbSpecial.Value = strTemp
    
    strTemp = Replace(tbSOW.Value, strFind, "<b>" & cbReplaceText.Value & "</b>")
    tbSOW.Value = strTemp
    
    For n = 0 To lbCond.ListCount - 1
        If lbCond.List(n, 0) = lbVariables.List(i, 0) Then
            'MsgBox "Found Variable"
            
            If lbCond.List(n, 1) = cbReplaceText.Value Then
                'MsgBox "Found Condition." & vbCr & lbCond.List(n, 2)
                strFind = "{{" & cbFind.Value & "}}"
                
                strTemp = Replace(tbSpecial.Value, strFind, "<b>" & lbCond.List(n, 2) & "</b>")
                tbSpecial.Value = strTemp
                
                strTemp = Replace(tbSOW.Value, strFind, "<b>" & lbCond.List(n, 2) & "</b>")
                tbSOW.Value = strTemp
                
                If InStr(lbCond.List(n, 2), ">>") > 0 Then
                    vLine = Split(lbCond.List(n, 2), ">>")
                    vTemp = Split(vLine(0), "<<")
                    vVar = Split(vTemp(1), "/")
            
                    lbVariables.AddItem vVar(0)
                    If UBound(vVar) > 0 Then
                        lbVariables.List(lbVariables.ListCount - 1, 1) = Replace(vVar(1), ">>", "")
                    Else
                        lbVariables.List(lbVariables.ListCount - 1, 1) = "<Empty>"
                    End If
                End If
            End If
        End If
    Next n
    
    lbVariables.RemoveItem i
    
    cbFind.Clear
    For j = 0 To lbVariables.ListCount - 1
        cbFind.AddItem lbVariables.List(j, 0)
    Next j
    cbFind.Value = ""
    
    cbReplaceText.Clear
    cbReplaceText.Value = ""
End Sub

Private Sub cbUpdate_Click()
    'Dim objSS As AcadSelectionSet
    'Dim objSOW As AcadBlockReference
    'Dim vAttList As Variant
    'Dim filterType, filterValue As Variant
    'Dim grpCode(0) As Integer
    'Dim grpValue(0) As Variant
    
    'GoTo Skip_Block

'    grpCode(0) = 2
'    grpValue(0) = "scope of work"
'    filterType = grpCode
'    filterValue = grpValue
'
'    On Error Resume Next
'    Err = 0
'
'    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
'    If Not Err = 0 Then
'        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
'        Err = 0
'    End If
'
'    objSS.Select acSelectionSetAll, , , filterType, filterValue
'    If objSS.count < 1 Then GoTo Exit_Sub
'
'    For Each objBlock In objSS
'        vAttList = objBlock.GetAttributes
'
'        vAttList(0).TextString = tbNumber.Value
'        vAttList(1).TextString = tbTitle.Value
'        vAttList(2).TextString = cbJobType.Value & "^" & tbDescription.Value
'        vAttList(3).TextString = tbStart.Value
'        vAttList(4).TextString = tbEnd.Value
'        vAttList(5).TextString = tbWirecenter.Value
'        vAttList(6).TextString = cbStructure.Value
'        vAttList(7).TextString = cbPlanner.Value
'        vAttList(8).TextString = tbDate.Value
'        vAttList(9).TextString = tbReferences.Value
'        vAttList(10).TextString = tbSpecial.Value
'        vAttList(11).TextString = tbSOW.Value
'        vAttList(12).TextString = tbRevision.Value
'    Next objBlock
    
'Skip_Block:
    
    'Dim strFileName As String
    'Dim strDWGName As String
    'Dim strPath As String
    'Dim vTemp As Variant
    
    'strPath = ThisDrawing.Path & "\"
    'vTemp = Split(ThisDrawing.Name, " ")
    'strDWGName = Left(ThisDrawing.Name, 9)
    'strFileName = strPath & vTemp(0) & " Scope of Work.txt"

    'Open strFileName For Output As #1
    
    'If tbNumber.Value = "" Then tbNumber.Value = "<Blank>"
    'Print #1, "Job Number:" & vbTab & tbNumber.Value
    'If tbTitle.Value = "" Then tbTitle.Value = "<Blank>"
    'Print #1, "Job Title:" & vbTab & tbTitle.Value
    'If cbJobType.Value = "" Then cbJobType.Value = "<Blank>"
    'Print #1, "Job Type:" & vbTab & cbJobType.Value
    'If tbDescription.Value = "" Then tbDescription.Value = "<Blank>"
    'Print #1, "Job Description:" & vbTab & tbDescription.Value & vbCr
    'If tbWirecenter.Value = "" Then tbWirecenter.Value = "<Blank>"
    'Print #1, "Wirecenter:" & vbTab & tbWirecenter.Value
    'If cbStructure.Value = "" Then cbStructure.Value = "<Blank>"
    'Print #1, "Structure:" & vbTab & cbStructure.Value & vbCr
    'If tbStart.Value = "" Then tbStart.Value = "<Blank>"
    'Print #1, "Handoff:" & vbTab & tbStart.Value
    'If tbEnd.Value = "" Then tbEnd.Value = "<Blank>"
    'Print #1, "Due Date:" & vbTab & tbEnd.Value & vbCr
    'Print #1, "Planner:" & vbTab & cbPlanner.Value
    'Print #1, "Approved:" & vbTab & tbDate.Value & vbCr
    'Print #1, "References:" & vbCr & tbReferences.Value & vbCr
    'Print #1, "Notes:" & vbCr & tbSpecial.Value & vbCr
    'Print #1, "Scope of Work:" & vbCr & tbSOW.Value & vbCr
    'Print #1, "Revisions:" & vbCr & tbRevision.Value & vbCr
    
    'Close #1
    
    Call CreateHTML
    
    If tbEnd.Value = "<Blank>" Then
        Frame1.ForeColor = &HFF
        Frame1.BorderColor = &HFF
        Label7.ForeColor = &HFF
    Else
        'Exit Sub
        If CDate(tbEnd.Value) <= Date Then
            tbEnd.ForeColor = &HFF
            Frame1.ForeColor = &HFF
            Frame1.BorderColor = &HFF
            Label7.ForeColor = &HFF
        Else
            tbEnd.ForeColor = &H80000012
            Frame1.ForeColor = &H80000012
            Frame1.BorderColor = &H80000012
            Label7.ForeColor = &H80000012
        End If
    End If
    
    'If tbEnd.Value = "" Then
        'Frame1.ForeColor = &HFF
        'Frame1.BorderColor = &HFF
        'Label7.ForeColor = &HFF
    'End If
    
    If tbNumber.Value = "" Then
        Label1.ForeColor = &HFF
    Else
        Label1.ForeColor = &H80000012
    End If
    
    If tbSOW.Value = "" Then
        Label9.ForeColor = &HFF
    Else
        Label9.ForeColor = &H80000012
    End If
Exit_Sub:
    MsgBox "Done"
End Sub

Private Sub Label1_Click()
    Dim strTemp As String
    Dim vTemp As Variant
    
    strTemp = Replace(UCase(ThisDrawing.Name), ".DWG", "")
    vTemp = Split(strTemp, " ")
    tbNumber.Value = vTemp(0)
    
    If UBound(vTemp) = 0 Then
        tbTitle.Value = ""
    Else
        strTemp = ""
        For i = 1 To UBound(vTemp)
            If strTemp = "" Then
                strTemp = vTemp(i)
            Else
                strTemp = strTemp & " " & vTemp(i)
            End If
        Next i
        tbTitle.Value = strTemp
    End If
End Sub

Private Sub Label16_Click()
    tbDate.Value = Date
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbVariables.ColumnCount = 2
    lbVariables.ColumnWidths = "144;144"
    
    lbCond.ColumnCount = 3
    lbCond.ColumnWidths = "84;84;120"
    
    cbStructure.AddItem ""
    cbStructure.AddItem "96 FDH"
    cbStructure.AddItem "144 FDH"
    cbStructure.AddItem "288 FDH"
    cbStructure.AddItem "432 FDH"
    cbStructure.AddItem "576 FDH"
    cbStructure.AddItem "CO Panel"
    cbStructure.AddItem "RST Panel"
    
    cbF1.AddItem ""
    cbF1.AddItem "NA"
    cbF1.AddItem "Unknown"
    
    cbF2.AddItem ""
    cbF2.AddItem "NA"
    cbF2.AddItem "Unknown"
    
    'cbJobType.AddItem "DSPLIT"
    'cbJobType.AddItem "FDH-ALL"
    'cbJobType.AddItem "FDH-F1"
    'cbJobType.AddItem "FDH-F2"
    'cbJobType.AddItem "TRUNK"
    'cbJobType.AddItem "BUS"
    
    cbPlanner.AddItem "Adam Kemper"
    cbPlanner.AddItem "Byron Auer"
    cbPlanner.AddItem "Franklin Angulo"
    cbPlanner.AddItem "Matt Snyder"
    
    Call RetreiveData
    If tbNumber.Value = "<Blank>" Or tbNumber.Value = "" Then GoTo Skip_Block_Entry
    
'    Dim objSS As AcadSelectionSet
'    Dim objSOW As AcadBlockReference
'    Dim vAttList As Variant
'    Dim filterType, filterValue As Variant
'    Dim grpCode(0) As Integer
'    Dim grpValue(0) As Variant
'
'    grpCode(0) = 2
'    grpValue(0) = "scope of work"
'    filterType = grpCode
'    filterValue = grpValue
'
'    On Error Resume Next
'    Err = 0
'
'    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
'    If Not Err = 0 Then
'        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
'        Err = 0
'    End If
'
'    objSS.Select acSelectionSetAll, , , filterType, filterValue
'
'    If objSS.count < 1 Then
'        tbDescription.Value = "No Scope of Work found"
'        Exit Sub
'    End If
'
'    Dim vLine, vItem As Variant
'
'    For Each objBlock In objSS
'        vAttList = objBlock.GetAttributes
'
'        vLine = Split(vAttList(2).TextString, "^")
'        If UBound(vLine) > 0 Then
'            tbNumber.Value = vAttList(0).TextString
'            tbTitle.Value = vAttList(1).TextString
'
'            cbJobType.Value = vLine(0)
'            tbDescription.Value = vLine(1)
'            tbStart.Value = vAttList(3).TextString
'            tbEnd.Value = vAttList(4).TextString
'            tbWirecenter.Value = vAttList(5).TextString
'            cbStructure.Value = vAttList(6).TextString
'            cbPlanner.Value = vAttList(7).TextString
'            tbDate.Value = vAttList(8).TextString
'            tbReferences.Value = vAttList(9).TextString
'            tbSpecial.Value = vAttList(10).TextString
'            tbSOW.Value = vAttList(11).TextString
'            tbRevision.Value = vAttList(12).TextString
'        Else
'            tbNumber.Value = vAttList(0).TextString
'            tbTitle.Value = vAttList(1).TextString
'            tbDescription.Value = vAttList(2).TextString
'            tbStart.Value = vAttList(3).TextString
'            tbEnd.Value = vAttList(4).TextString
'            tbWirecenter.Value = vAttList(5).TextString
'            cbStructure.Value = vAttList(6).TextString
'            cbF1.Value = vAttList(7).TextString
'            cbF2.Value = vAttList(8).TextString
'            tbReferences.Value = vAttList(9).TextString
'            tbSpecial.Value = vAttList(10).TextString
'            tbSOW.Value = vAttList(11).TextString
'            tbRevision.Value = vAttList(12).TextString
'        End If
'
'    Next objBlock
    
    If tbEnd.Value = "ASAP" Then GoTo Skip_Block_Entry
    
    If tbEnd.Value = "<Blank>" Then
        Frame1.ForeColor = &HFF
        Frame1.BorderColor = &HFF
        Label7.ForeColor = &HFF
    Else
        If CDate(tbEnd.Value) <= Date Then
            tbEnd.ForeColor = &HFF
            Frame1.ForeColor = &HFF
            Frame1.BorderColor = &HFF
            Label7.ForeColor = &HFF
        End If
    End If
    
    If tbNumber.Value = "" Then
        Label1.ForeColor = &HFF
    Else
        Label1.ForeColor = &H80000012
    End If
    
    If tbSOW.Value = "" Then
        Label9.ForeColor = &HFF
    Else
        Label9.ForeColor = &H80000012
    End If
    
Skip_Block_Entry:
    
    Dim strPath, strFile As String
    Dim vTemp As Variant
    
    strTFolder = ""
    
    strPath = ThisDrawing.Path
    'MsgBox strPath
    vTemp = Split(strPath, "Dropbox")
    If UBound(vTemp) < 1 Then
        Exit Sub
    End If
    
    strTFolder = vTemp(0) & "\Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\1-JOBS\PLANNING\SCOPE TEMPLATES\*.*"
    
    strFile = Dir$(strTFolder)
    
    Do While strFile <> ""
        If InStr(strFile, ".swt") Then
            cbJobType.AddItem Replace(strFile, ".swt", "")
        End If
        strFile = Dir$
    Loop
    
    Call GetVariableList
End Sub

Private Sub CreateHTML()
    Dim strFileName As String
    Dim strDWGName As String
    Dim strPath As String
    Dim vTemp As Variant
    Dim vLine As Variant
    
    On Error Resume Next
    
    strPath = ThisDrawing.Path & "\"
    vTemp = Split(ThisDrawing.Name, " ")
    'strDWGName = Left(ThisDrawing.Name, 9)
    'strDWGName = vTemp(0)
    strFileName = strPath & vTemp(0) & " Scope of Work.htm"

    Open strFileName For Output As #1
    Err = 0
    
    Print #1, "<!--"
    Print #1, ">" & vbTab & "Number" & vbTab & tbNumber.Value
    Print #1, ">" & vbTab & "Title" & vbTab & tbTitle.Value
    Print #1, ">" & vbTab & "Desc" & vbTab & tbDescription.Value
    Print #1, ">" & vbTab & "WC" & vbTab & tbWirecenter.Value
    Print #1, ">" & vbTab & "Type" & vbTab & cbJobType.Value
    Print #1, ">" & vbTab & "Start" & vbTab & tbStart.Value
    Print #1, ">" & vbTab & "End" & vbTab & tbEnd.Value
    Print #1, ">" & vbTab & "Planner" & vbTab & cbPlanner.Value
    Print #1, ">" & vbTab & "Approved" & vbTab & tbDate.Value
    Print #1, ">" & vbTab & "Reference"
    Print #1, tbReferences.Value
    Print #1, ">" & vbTab & "Notes"
    Print #1, tbSpecial.Value
    Print #1, ">" & vbTab & "Scope"
    Print #1, tbSOW.Value
    Print #1, ">" & vbTab & "Revisions"
    Print #1, tbRevision.Value
    Print #1, "-->"
    If Not Err = 0 Then
        MsgBox "Error saving Data!"
        Close #1
        Exit Sub
    End If
    
    Print #1, "<!doctype html>"
    Print #1, "<html><head><style>"
    Print #1, "h2 {text-align: center;}"
    Print #1, ".border {border: 1px solid black;}"
    Print #1, ".borderAll {border: 1px solid black; width: 1024px;}"
    Print #1, ".borderLeft {border-left: 1px solid black;}"
    Print #1, "th, td {padding: 4px;}"
    Print #1, "</style></head>"
    
    Print #1, "<body><h2>" & tbNumber.Value & " SCOPE OF WORK</h2>"
    
    Print #1, "<table class=""borderAll""><tr>"
    Print #1, "<td>Job Number </td><td><b>" & tbNumber.Value & "</b></td></tr>"
    Print #1, "<td>Job Title </td><td><b>" & tbTitle.Value & "</b></td></tr>"
    Print #1, "<td>Job Type </td><td><b>" & cbJobType.Value & "</b></td></tr>"
    Print #1, "<td>Job Description </td><td><b>" & tbDescription.Value & "</b></td></tr>"
    Print #1, "<td>Wirecenter </td><td><b>" & tbWirecenter.Value & "</b></td></tr>"
    Print #1, "</table><br>"
    
    Print #1, "<table class=""borderAll""><tr>"
    Print #1, "<td>Handoff</td><td><b>" & tbStart.Value & "</b></td>"
    Print #1, "<td class="" borderLeft"">Planner </td><td><b>" & cbPlanner.Value & "</b></td></tr>"
    Print #1, "<tr><td>Due Date</td><td><b>" & tbEnd.Value & "</b></td>"
    Print #1, "<td class=""borderLeft"">Approved </td><td><b>" & tbDate.Value & "</b></td></tr>"
    Print #1, "</table><br>"
    
    Print #1, "<table class=""borderAll""><tr>"
    Print #1, "<th class=""border"">Reference Jobs</th></tr>"
    Print #1, "<tr><td>" & Replace(tbReferences.Value, vbCr, "<br>") & "</td></tr>"
    Print #1, "</table><br>"
    Print #1, "<table class=""borderAll""><tr>"
    Print #1, "<th class=""border"">Notes</th></tr>"
    Print #1, "<tr><td>" '& Replace(tbSpecial.Value, vbCr, "<br>") & "</td></tr>"
    'Print #1, "<tr><td>" & Replace(tbSpecial.Value, vbCr, "<br>") & "</td></tr>"
    vLine = Split(tbSpecial.Value, vbCr)
    For i = 0 To UBound(vLine)
        vLine(i) = Replace(vLine(i), vbLf, "")
        If vLine(i) = "<table>" Then
            Print #1, vLine(i)
            i = i + 1
            vLine(i) = Replace(vLine(i), vbLf, "")
            While Not Left(vLine(i), 3) = "</t"
                vLine(i) = Replace(vLine(i), "<tr>", "<tr><td>")
                vLine(i) = Replace(vLine(i), " == ", "</td><td>")
                vLine(i) = Replace(vLine(i), " = ", "</td><td class=""border"">")
                vLine(i) = vLine(i) & "</td></tr>"
                Print #1, vLine(i)
            
                i = i + 1
                vLine(i) = Replace(vLine(i), vbLf, "")
                'MsgBox "Left(vLine(i), 4)= " & Left(vLine(i), 4)
            Wend
            Print #1, vLine(i) & "<br>"
            i = i + 1
        End If
        Print #1, vLine(i) & "<br>"
    Next i
    Print #1, "</td></tr>"
    Print #1, "</table><br>"
    Print #1, "<table class=""borderAll""><tr>"
    
    Print #1, "<th class=""border"">Scope of Work</th></tr>"
    Print #1, "<tr><td>" '& Replace(tbSOW.Value, vbCr, "<br>") & "</td></tr>"
    vLine = Split(tbSOW.Value, vbCr)
    For i = 0 To UBound(vLine)
        vLine(i) = Replace(vLine(i), vbLf, "")
        If vLine(i) = "<table>" Then
            Print #1, vLine(i)
            i = i + 1
            vLine(i) = Replace(vLine(i), vbLf, "")
            While Not Left(vLine(i), 3) = "</t"
                vLine(i) = Replace(vLine(i), "<tr>", "<tr><td>")
                vLine(i) = Replace(vLine(i), " = ", "</td><td class=""border"">")
                vLine(i) = vLine(i) & "</td></tr>"
                Print #1, vLine(i)
            
                i = i + 1
                vLine(i) = Replace(vLine(i), vbLf, "")
                'MsgBox "Left(vLine(i), 4)= " & Left(vLine(i), 4)
            Wend
            Print #1, vLine(i) & "<br>"
            i = i + 1
        End If
        Print #1, vLine(i) & "<br>"
    Next i
    Print #1, "</td></tr>"
    Print #1, "</table><br>"
    
    Print #1, "<table class=""borderAll""><tr>"
    Print #1, "<th class=""border"">Revisions</th></tr>"
    Print #1, "<tr><td>" & Replace(tbRevision.Value, vbCr, "<br>") & "</td></tr>"
    Print #1, "</table></body></html>"
    Print #1, ""
    If Not Err = 0 Then
        MsgBox "Error creating Scope!"
    End If
    
    Close #1
End Sub

Private Sub RetreiveData()
    Dim strFileName As String
    Dim strPath As String
    Dim strLine As String
    Dim vTemp As Variant
    Dim vLine As Variant
    Dim fName As String
    Dim iCount As Integer
    
    'On Error Resume Next
    
    iCount = 0
    tbReferences.Value = ""
    tbSpecial.Value = ""
    tbSOW.Value = ""
    tbRevision.Value = ""
    
    strPath = ThisDrawing.Path & "\"
    vTemp = Split(ThisDrawing.Name, " ")
    strFileName = strPath & vTemp(0) & " Scope of Work.htm"
    
    'strFileName = LabelFileName.Caption & cbFileList.Value
    
    fName = Dir(strFileName)
    If fName = "" Then
        MsgBox "No Scope of Work File found."
        'Call GetDataFromBlock
        Exit Sub
    End If
    
    Open strFileName For Input As #2
    
    While Not EOF(2)
        Line Input #2, strLine
        vLine = Split(strLine, vbTab)
        
        If vLine(0) = ">" Then
Check_vLine1:
            'MsgBox strLine
            If UBound(vLine) > 1 Then
                If vLine(2) = "" Then vLine(2) = "<Blank>"
            End If
            
            Select Case vLine(1)
                Case "Number"
                    tbNumber.Value = vLine(2)
                Case "Title"
                    tbTitle.Value = vLine(2)
                Case "Desc"
                    tbDescription.Value = vLine(2)
                Case "WC"
                    tbWirecenter.Value = vLine(2)
                Case "Type"
                    cbJobType.Value = vLine(2)
                Case "Start"
                    tbStart.Value = vLine(2)
                Case "End"
                    tbEnd.Value = vLine(2)
                Case "Planner"
                    cbPlanner.Value = vLine(2)
                Case "Approved"
                    tbDate.Value = vLine(2)
                Case "Reference"
Continue_Ref:
                    Line Input #2, strLine
                    If strLine = "" Then GoTo Continue_Ref
                    
                    vLine = Split(strLine, vbTab)
                    If vLine(0) = ">" Then GoTo Check_vLine1
        
                    If tbReferences.Value = "" Then
                        tbReferences.Value = strLine
                    Else
                        tbReferences.Value = tbReferences.Value & vbCr & strLine
                    End If
                    
                    GoTo Continue_Ref
                Case "Notes"
Continue_Notes:
                    Line Input #2, strLine
                    If strLine = "" Then GoTo Continue_Notes
                    
                    vLine = Split(strLine, vbTab)
                    If vLine(0) = ">" Then GoTo Check_vLine1
                    
                    If tbSpecial.Value = "" Then
                        tbSpecial.Value = strLine
                    Else
                        tbSpecial.Value = tbSpecial.Value & vbCr & strLine
                    End If
                    
                    GoTo Continue_Notes
                Case "Scope"
Continue_SOW:
                    Line Input #2, strLine
                    If strLine = "" Then GoTo Continue_SOW
                    
                    vLine = Split(strLine, vbTab)
                    If vLine(0) = ">" Then GoTo Check_vLine1
                    
                    If tbSOW.Value = "" Then
                        tbSOW.Value = strLine
                    Else
                        tbSOW.Value = tbSOW.Value & vbCr & strLine
                    End If
                        
                    GoTo Continue_SOW
                Case "Revisions"
Continue_Revisions:
                    Line Input #2, strLine
                    If strLine = "" Then GoTo Continue_Revisions
                    If strLine = "-->" Then GoTo Exit_Sub
                        
                    vLine = Split(strLine, vbTab)
                    
                    If vLine(0) = ">" Then GoTo Check_vLine1
                    
                    If tbRevision.Value = "" Then
                        tbRevision.Value = strLine
                    Else
                        tbRevision.Value = tbRevision.Value & vbCr & strLine
                    End If
                        
                    GoTo Continue_Revisions
                Case "-->"
                    GoTo Exit_Sub
            End Select
        End If
    Wend
    
Exit_Sub:
    Close #2
End Sub

Private Sub GetDataFromBlock()
    Dim objSS As AcadSelectionSet
    Dim objSOW As AcadBlockReference
    Dim vAttList As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim strTemp As String

    grpCode(0) = 2
    grpValue(0) = "scope of work"
    filterType = grpCode
    filterValue = grpValue

    On Error Resume Next
    Err = 0

    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If

    objSS.Select acSelectionSetAll, , , filterType, filterValue

    If objSS.count < 1 Then
        MsgBox "No Scope of Work Block found"
        Exit Sub
    End If

    If objSS.count < 1 Then Exit Sub

    Dim vLine, vItem As Variant

    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes

        If vAttList(2).TextString = "" Then Exit Sub
        
        For k = 0 To 12
            If vAttList(k).TextString = "" Then vAttList(k).TextString = "<Empty>"
        Next k
        
        objBlock.Update

        vLine = Split(vAttList(2).TextString, "^")
        
        If UBound(vLine) > 0 Then
        strTemp = vAttList(0).TextString
        MsgBox strTemp
            tbNumber.Value = strTemp    'vAttList(0).TextString
        Exit Sub
            tbTitle.Value = vAttList(1).TextString

            cbJobType.Value = vLine(0)
            tbDescription.Value = vLine(1)
            tbStart.Value = vAttList(3).TextString
            tbEnd.Value = vAttList(4).TextString
            tbWirecenter.Value = vAttList(5).TextString
            cbStructure.Value = vAttList(6).TextString
            cbPlanner.Value = vAttList(7).TextString
            tbDate.Value = vAttList(8).TextString
            tbReferences.Value = vAttList(9).TextString
            tbSpecial.Value = vAttList(10).TextString
            tbSOW.Value = vAttList(11).TextString
            tbRevision.Value = vAttList(12).TextString
        Else
            tbNumber.Value = vAttList(0).TextString
            tbTitle.Value = vAttList(1).TextString
            tbDescription.Value = vAttList(2).TextString
            tbStart.Value = vAttList(3).TextString
            tbEnd.Value = vAttList(4).TextString
            tbWirecenter.Value = vAttList(5).TextString
            cbStructure.Value = vAttList(6).TextString
            cbF1.Value = vAttList(7).TextString
            cbF2.Value = vAttList(8).TextString
            tbReferences.Value = vAttList(9).TextString
            tbSpecial.Value = vAttList(10).TextString
            tbSOW.Value = vAttList(11).TextString
            tbRevision.Value = vAttList(12).TextString
        End If

    Next objBlock
End Sub

Private Sub GetVariableList()
    If cbJobType.Value = "" Then Exit Sub
    
    Dim vText, vLine, vTemp As Variant
    Dim vVar, vCond As Variant
    Dim strLine As String
    
    lbVariables.Clear
    lbCond.Clear
    
    strLine = Replace(tbSpecial.Value, vbLf, "")
    vText = Split(strLine, vbCr)
    
    For n = 0 To UBound(vText)
        If InStr(vText(n), ">>") > 0 And Not iStatus = 3 Then
            vLine = Split(vText(n), ">>")
            For j = 0 To UBound(vLine) - 1
            
                vTemp = Split(vLine(j), "<<")
                vVar = Split(vTemp(1), "/")
            
                'If lbVariables.ListCount < 0 Then
                    For i = 0 To lbVariables.ListCount - 1
                        If lbVariables.List(i, 0) = vVar(0) Then GoTo Skip_Adding_Var
                    Next i
                'End If
            
                lbVariables.AddItem vVar(0)
                If UBound(vVar) > 0 Then
                    lbVariables.List(lbVariables.ListCount - 1, 1) = Replace(vVar(1), ">>", "")
                Else
                    lbVariables.List(lbVariables.ListCount - 1, 1) = "<Empty>"
                End If
        
Skip_Adding_Var:
            Next j
            
            If UBound(vLine) > 2 Then
                For i = 1 To UBound(vLine) - 1
                    vTemp = Split(vLine(i), "<<")
                    vVar = Split(vTemp(1), "/")
            
                    lbVariables.AddItem vVar(0)
                    If UBound(vVar) > 0 Then
                        lbVariables.List(lbVariables.ListCount - 1, 1) = Replace(vVar(1), ">>", "")
                    Else
                        lbVariables.List(lbVariables.ListCount - 1, 1) = "<Empty>"
                    End If
Skip_Adding_This:
                    'cbFind.AddItem "<<" & vTemp(1) & ">>"
                Next i
            End If
        End If
    Next n
    
    For i = 0 To lbVariables.ListCount - 1
        cbFind.AddItem lbVariables.List(i, 0)
    Next i
End Sub

