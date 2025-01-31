VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectStatus 
   Caption         =   "Project Status"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "ProjectStatus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strFileName As String
Dim strProjectName As String
Dim iEmailSent As Integer

Private Sub cbGetWorkload_Click()
    Me.Hide
        Load zzzWorkload
            zzzWorkload.show
        Unload zzzWorkload
    Me.show
End Sub

Private Sub CommandButton1_Click()
    Dim vTemp As Variant
    Dim strReport, strTemp As String
    Dim strCompleted, strDate As String
    Dim iSendEmail As Integer
    
    Me.Hide
    
    strCompleted = ""
    
    vTemp = Split(Me.Caption, " ")
    strTemp = vTemp(0)
    strReport = "Project:" & vbTab & strTemp & vbCr
    
    strReport = strReport & "Setup:" & vbTab
    If Not tbPSI.Value = "" Then
        strReport = strReport & tbPSI.Value & vbTab & tbPSS.Value & vbTab & tbPSC.Value & vbTab & tbPSH.Value
        If Not tbPSC.Value = "" Then
            If tbPSC.Enabled = True Then strCompleted = "Setup"
        End If
    End If
    
    strReport = strReport & vbCr & "BDK:" & vbTab
    If Not tbBI.Value = "" Then
        strReport = strReport & tbBI.Value & vbTab & tbBS.Value & vbTab & tbBC.Value & vbTab & tbBH.Value
        If Not tbBC.Value = "" Then
            If tbBC.Enabled = True Then strCompleted = "BDK"
        End If
    End If
    
    strReport = strReport & vbCr & "Planning:" & vbTab
    If Not tbPLI.Value = "" Then
        strReport = strReport & tbPLI.Value & vbTab & tbPLS.Value & vbTab & tbPLC.Value & vbTab & tbPLH.Value
        If Not tbPLC.Value = "" Then
            If tbPLC.Enabled = True Then strCompleted = "Planning"
        End If
    End If
    
    strReport = strReport & vbCr & "Prepare:" & vbTab
    If Not tbCOI.Value = "" Then
        strReport = strReport & tbCOI.Value & vbTab & tbCOS.Value & vbTab & tbCOC.Value & vbTab & tbCOH.Value
        If Not tbCOC.Value = "" Then
            If tbCOC.Enabled = True Then strCompleted = "Prepare Job"
        End If
    End If
    
    strReport = strReport & vbCr & "Field:" & vbTab
    If Not tbFI.Value = "" Then
        strReport = strReport & tbFI.Value & vbTab & tbFS.Value & vbTab & tbFC.Value & vbTab & tbFH.Value
        If Not tbFC.Value = "" Then
            If tbFC.Enabled = True Then strCompleted = "Fielding"
        End If
    End If
    
    strReport = strReport & vbCr & "Cleanup:" & vbTab
    If Not tbCUI.Value = "" Then
        strReport = strReport & tbCUI.Value & vbTab & tbCUS.Value & vbTab & tbCUC.Value & vbTab & tbCUH.Value
        If Not tbCUC.Value = "" Then
            If tbCUC.Enabled = True Then strCompleted = "Cleanup"
        End If
    End If
    
    strReport = strReport & vbCr & "Cut Out:" & vbTab
    If Not tbCDI.Value = "" Then
        strReport = strReport & tbCDI.Value & vbTab & tbCDS.Value & vbTab & tbCDC.Value & vbTab & tbCDH.Value
        If Not tbCDC.Value = "" Then
            If tbCDC.Enabled = True Then strCompleted = "Cut Out"
        End If
    End If
    
    strReport = strReport & vbCr & "Make Ready:" & vbTab
    If Not tbMRI.Value = "" Then
        strReport = strReport & tbMRI.Value & vbTab & tbMRS.Value & vbTab & tbMRC.Value & vbTab & tbMRH.Value
        If Not tbMRC.Value = "" Then
            If tbMRC.Enabled = True Then strCompleted = "Make Ready"
        End If
    End If
    
    strReport = strReport & vbCr & "Permits:" & vbTab
    If Not tbPI.Value = "" Then
        strReport = strReport & tbPI.Value & vbTab & tbPS.Value & vbTab & tbPC.Value & vbTab & tbPH.Value
        If Not tbPC.Value = "" Then
            If tbPC.Enabled = True Then strCompleted = "Permits"
        End If
    End If
    
    strReport = strReport & vbCr & "Counts:" & vbTab
    If Not tbCntI.Value = "" Then
        strReport = strReport & tbCntI.Value & vbTab & tbCntS.Value & vbTab & tbCntC.Value & vbTab & tbCntH.Value
        If Not tbCntC.Value = "" Then
            If tbCntC.Enabled = True Then strCompleted = "Counts"
        End If
    End If
    
    strReport = strReport & vbCr & "Units:" & vbTab
    If Not tbUI.Value = "" Then
        strReport = strReport & tbUI.Value & vbTab & tbUS.Value & vbTab & tbUC.Value & vbTab & tbUH.Value
        If Not tbUC.Value = "" Then
            If tbUC.Enabled = True Then strCompleted = "Units"
        End If
    End If
    
    strReport = strReport & vbCr & "QC:" & vbTab
    If Not tbQCI.Value = "" Then
        strReport = strReport & tbQCI.Value & vbTab & tbQCS.Value & vbTab & tbQCC.Value & vbTab & tbQCH.Value
        If Not tbQCC.Value = "" Then
            If tbQCC.Enabled = True Then strCompleted = "QC"
        End If
    End If
    
    strReport = strReport & vbCr & "Reports:" & vbTab
    If Not tbRI.Value = "" Then
        strReport = strReport & tbRI.Value & vbTab & tbRS.Value & vbTab & tbRC.Value & vbTab & tbRH.Value
        If Not tbRC.Value = "" Then
            If tbRC.Enabled = True Then strCompleted = "Reports"
        End If
    End If
    
  On Error Resume Next
    
    If Not strCompleted = "" Then
        SendCompleteEmail (strCompleted)
    End If
    
    If iEmailSent = 0 Then
        Dim result As Integer
    
        result = MsgBox("Update without email?", vbYesNo, "Update without Email")
        If result = vbNo Then
            MsgBox "Update Cancelled"
            Me.show
            Exit Sub
        Else
            strCompleted = ""
        End If
    End If
    
    Open strFileName For Output As #3
    
    Print #3, UCase(strReport)
    
    Close #3
    
    Call RefreshProjectForm
    
    If Not strCompleted = "" Then
        MsgBox "Updated and Email sent for " & strCompleted
    Else
        MsgBox "Updated"
    End If
    Me.show
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Private Sub Frame1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbFI.Enabled = True
    tbFS.Enabled = True
    tbFC.Enabled = True
End Sub

Private Sub Frame10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbCDI.Enabled = True
    tbCDS.Enabled = True
    tbCDC.Enabled = True
End Sub

Private Sub Frame2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbCOI.Enabled = True
    tbCOS.Enabled = True
    tbCOC.Enabled = True
End Sub

Private Sub Frame3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbCUI.Enabled = True
    tbCUS.Enabled = True
    tbCUC.Enabled = True
End Sub

Private Sub Frame4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbMRI.Enabled = True
    tbMRS.Enabled = True
    tbMRC.Enabled = True
End Sub

Private Sub Frame5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbPI.Enabled = True
    tbPS.Enabled = True
    tbPC.Enabled = True
End Sub

Private Sub Frame6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbCntI.Enabled = True
    tbCntS.Enabled = True
    tbCntC.Enabled = True
End Sub

Private Sub Frame7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbUI.Enabled = True
    tbUS.Enabled = True
    tbUC.Enabled = True
End Sub

Private Sub Frame8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbQCI.Enabled = True
    tbQCS.Enabled = True
    tbQCC.Enabled = True
End Sub

Private Sub Frame9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbRI.Enabled = True
    tbRS.Enabled = True
    tbRC.Enabled = True
End Sub

Private Sub Label41_Click()
    If tbCOI.Value = "" Then Exit Sub
    If Not tbCOS.Value = "" Then Exit Sub
    tbCOS.Value = Date
    
    'SendCompleteEmail ("Prepare Assigned")
    'MsgBox "Email sent to Project Management"
End Sub

Private Sub Label42_Click()
    'Add Fielding Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbFH.Value = "" Then
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbFH.Value)
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        'strDate = ProjectStatusHours.tbDate.Value
        'strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbFH.Value = dOldHours + dNewHours
    
End Sub

Private Sub Label44_Click()
    If Not tbFC.Value = "" Then Exit Sub
    tbFC.Value = Date
End Sub

Private Sub Label45_Click()
    If Not tbFS.Value = "" Then Exit Sub
    If tbFI.Value = "" Then Exit Sub
    tbFS.Value = Date
    SendCompleteEmail ("Fielding Assigned")
    MsgBox "Email sent to Project Management"
End Sub

Private Sub Label46_Click()
    'Add Cleanup Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbCUH.Value = "" Then    '<---------------------------------------------Change
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbCUH.Value)    '<---------------------------------------------Change
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbCUH.Value = dOldHours + dNewHours    '<---------------------------------------------Change
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label48_Click()
    If Not tbCUC.Value = "" Then Exit Sub
    tbCUC.Value = Date
End Sub

Private Sub Label49_Click()
    If Not tbCUS.Value = "" Then Exit Sub
    tbCUS.Value = Date
End Sub

Private Sub Label5_Click()
    If Not tbCOC.Value = "" Then Exit Sub
    tbCOC.Value = Date
End Sub

Private Sub Label50_Click()
    'Add Cleanup Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbMRH.Value = "" Then    '<---------------------------------------------Change
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbMRH.Value)    '<---------------------------------------------Change
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbMRH.Value = dOldHours + dNewHours    '<---------------------------------------------Change
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label52_Click()
    If Not tbMRC.Value = "" Then Exit Sub
    tbMRC.Value = Date
End Sub

Private Sub Label53_Click()
    If Not tbMRS.Value = "" Then Exit Sub
    tbMRS.Value = Date
End Sub

Private Sub Label54_Click()
    'Add Cleanup Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbPH.Value = "" Then    '<---------------------------------------------Change
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbPH.Value)    '<---------------------------------------------Change
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbPH.Value = dOldHours + dNewHours    '<---------------------------------------------Change
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label56_Click()
    If Not tbPC.Value = "" Then Exit Sub
    tbPC.Value = Date
End Sub

Private Sub Label57_Click()
    If Not tbPS.Value = "" Then Exit Sub
    tbPS.Value = Date
End Sub

Private Sub Label58_Click()
    'Add Cleanup Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbCntH.Value = "" Then    '<---------------------------------------------Change
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbCntH.Value)    '<---------------------------------------------Change
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbCntH.Value = dOldHours + dNewHours    '<---------------------------------------------Change
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label6_Click()
    'Add Prepare Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbCOH.Value = "" Then
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbCOH.Value)
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbCOH.Value = dOldHours + dNewHours
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label60_Click()
    If Not tbCntC.Value = "" Then Exit Sub
    tbCntC.Value = Date
End Sub

Private Sub Label61_Click()
    If Not tbCntS.Value = "" Then Exit Sub
    tbCntS.Value = Date
End Sub

Private Sub Label62_Click()
    'Add Cleanup Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbUH.Value = "" Then    '<---------------------------------------------Change
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbUH.Value)    '<---------------------------------------------Change
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbUH.Value = dOldHours + dNewHours    '<---------------------------------------------Change
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label64_Click()
    If Not tbUC.Value = "" Then Exit Sub
    tbUC.Value = Date
End Sub

Private Sub Label65_Click()
    If Not tbUS.Value = "" Then Exit Sub
    tbUS.Value = Date
End Sub

Private Sub Label66_Click()
    'Add Cleanup Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbQCH.Value = "" Then    '<---------------------------------------------Change
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbQCH.Value)    '<---------------------------------------------Change
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbQCH.Value = dOldHours + dNewHours    '<---------------------------------------------Change
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label68_Click()
    If Not tbQCC.Value = "" Then Exit Sub
    tbQCC.Value = Date
End Sub

Private Sub Label69_Click()
    If Not tbQCS.Value = "" Then Exit Sub
    tbQCS.Value = Date
End Sub

Private Sub Label70_Click()
    'Add Cleanup Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbRH.Value = "" Then    '<---------------------------------------------Change
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbRH.Value)    '<---------------------------------------------Change
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbRH.Value = dOldHours + dNewHours    '<---------------------------------------------Change
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label72_Click()
    If Not tbRC.Value = "" Then Exit Sub
    tbRC.Value = Date
End Sub

Private Sub Label73_Click()
    If Not tbRS.Value = "" Then Exit Sub
    tbRS.Value = Date
End Sub

Private Sub Label74_Click()
    If Not tbCDC.Value = "" Then Exit Sub
    tbCDC.Value = Date
End Sub

Private Sub Label75_Click()
    'Add Cleanup Hours
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbCDH.Value = "" Then    '<---------------------------------------------Change
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbCDH.Value)    '<---------------------------------------------Change
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbCDH.Value = dOldHours + dNewHours    '<---------------------------------------------Change
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label77_Click()
    If Not tbCDS.Value = "" Then Exit Sub
    tbCDS.Value = Date
End Sub

Private Sub Label78_Click()
    Dim dNewHours, dOldHours, dTotal As Double
    Dim strDate As String
    
    If tbPLH.Value = "" Then
        dOldHours = 0#
    Else
        dOldHours = CDbl(tbPLH.Value)
    End If
    
    strJob = strProjectName
    
    Load ProjectStatusHours
        ProjectStatusHours.tbJob.Value = strJob
        ProjectStatusHours.show
        
        dNewHours = CDbl(ProjectStatusHours.tbHours.Value)
        strDate = ProjectStatusHours.tbDate.Value
        strJob = ProjectStatusHours.tbJob.Value
        'MsgBox strDate & " - " & dNewHours & " Hours"
    Unload ProjectStatusHours
    
    tbPLH.Value = dOldHours + dNewHours
    
    Call AddHoursToTimesheet(CStr(strJob), CStr(strDate), CDbl(dNewHours))
End Sub

Private Sub Label80_Click()
    If Not tbPLC.Value = "" Then Exit Sub
    tbPLC.Value = Date
End Sub

Private Sub Label81_Click()
    If tbPLI.Value = "" Then Exit Sub
    If Not tbPLS.Value = "" Then Exit Sub
    tbPLS.Value = Date
    
    SendCompleteEmail ("Planning Assigned")
    MsgBox "Email sent to Project Management"
End Sub

Private Sub Label84_Click()
    If Not tbPSC.Value = "" Then Exit Sub
    tbPSC.Value = Date
    
    Frame13.Enabled = True
End Sub

Private Sub Label85_Click()
    If tbPSI.Value = "" Then Exit Sub
    If Not tbPSS.Value = "" Then Exit Sub
    tbPSS.Value = Date
End Sub

Private Sub Label88_Click()
    If Not tbBC.Value = "" Then Exit Sub
    tbBC.Value = Date
    
    Frame11.Enabled = True
End Sub

Private Sub Label89_Click()
    If tbBI.Value = "" Then Exit Sub
    If Not tbBS.Value = "" Then Exit Sub
    tbBS.Value = Date
End Sub

Private Sub UserForm_Initialize()
    Dim strProject As String
    Dim vName As Variant
    
    strProject = ThisDrawing.Name
    vName = Split(strProject, " ")
    strProjectName = vName(0)
    Me.Caption = strProjectName & " Project Status"
    
    Call RefreshProjectForm
End Sub

Private Sub SendCompleteEmail(strStatus As String)
    If Not cbEmails.Value Then Exit Sub
    
    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem
    Dim strTo, strCC, strSubject, strBody As String
    Dim strVerb As String
    Dim strCompany, strPath As String
    Dim vTemp As Variant
    Dim iNumber, iEmailSelf As Integer
    Dim iMTEMC, iOther, iDWG As Integer
    
    On Error Resume Next
    
    vTemp = Split(Me.Caption, " ")
    
    If Right(strStatus, 1) = "s" Then
        strVerb = "have"
    Else
        strVerb = "has"
    End If
    
    strPath = ThisDrawing.Path
    strCompany = ""
    If InStr(UCase(strPath), "UNITED") Then strCompany = "UTC "
    If InStr(UCase(strPath), "LORETTO") Then strCompany = "Loretto "
    
    'strTo = "jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com;jeremy.pafford@integrity-us.com"
    strSubject = "<Project Update> " & strCompany & strProjectName & " - " & strStatus & " Complete"
    strBody = "<b>" & strStatus & "</b> " & strVerb & " been completed and ready for the next stage."
    strTo = "jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
    strCC = ""
    
    Load ProjectStatusEmail
    
    'Me.Hide

    Select Case strStatus
        Case "Units", "Counts"
            strTo = strTo & ";daniel.campbell@integrity-us.com"
        'Case "Prepare Job", "Fielding"
            'strTo = "jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com;wade.hampton@integrity-us.com"
        Case "QC"
            strBody = GetQCData(strBody)
        Case "Cleanup"
            strBody = GetCleanupData(strBody)
            ProjectStatusEmail.lbCC.AddItem "Ronn.Elliott@Integrity-US.com"
        Case "Cut Out"
            ''If UCase(tbCOI.Value) = "TT" Then
                ''strBody = "Ronn, This one is now ready for you.  <span style='font-family:Segoe UI Emoji,sans-serif'>&#128522;</span><br><br>Tara"
            ''End If
            'strTo = "jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
            Err = 0
            iNumber = GetLastDWG()
            If Not Err = 0 Then iNumber = 0
            strBody = strBody & "<br><b>" & iNumber & " Drawings</b>"
        'Case "Make Ready"
            'strTo = "jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
        Case "Permits"
            'strTo = "jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
            If tbPI.Value = "N/A" Then
                strSubject = "<Project Update> " & strCompany & vTemp(0) & " Permits - Not Required"
                strBody = "<b>Permits are not required for this Job.</b>"
            End If
        Case "Setup"
            strTo = "adam.kemper@integrity-us.com;jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
            strSubject = "<Project Update> " & strCompany & vTemp(0) & " Planning Setup completed by " & tbPSI.Value
            strBody = "<b>BDK</b> is ready to be assigned."
            ProjectStatusEmail.lbCC.AddItem "Byron.Auer@Integrity-US.com"
            ProjectStatusEmail.lbCC.AddItem "Franklin.Angulo@Integrity-US.com"
            ProjectStatusEmail.lbCC.AddItem "Matt.Snyder@Integrity-US.com"
        Case "BDK"
            strTo = "adam.kemper@integrity-us.com;jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
            strSubject = "<Project Update> " & strCompany & vTemp(0) & " BDK is completed by " & tbBI.Value
            strBody = "<b>Planning</b> is ready to be assigned."
            ProjectStatusEmail.lbCC.AddItem "Byron.Auer@Integrity-US.com"
            ProjectStatusEmail.lbCC.AddItem "Franklin.Angulo@Integrity-US.com"
            ProjectStatusEmail.lbCC.AddItem "Matt.Snyder@Integrity-US.com"
        Case "Planning"
            strTo = "adam.kemper@integrity-us.com;jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
            strSubject = "<Project Update> " & strCompany & vTemp(0) & " Planning Review has completed by " & tbPLI.Value
            strBody = "<b>Prepare</b> is ready to be assigned."
            ProjectStatusEmail.lbCC.AddItem "Wade.Hampton@Integrity-US.com"
        Case "Planning Assigned"
            strTo = "adam.kemper@integrity-us.com;jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
            strSubject = "<Project Update> " & strCompany & vTemp(0) & " Planning task assigned to " & tbPLI.Value
            strBody = "<b>" & tbPLI.Value & "</b> has been assigned " & vTemp(0) & " to plan."
        Case "Fielding Assigned"
            'strTo = "jon.wilburn@integrity-us.com;rich.taylor@integrity-us.com"
            strSubject = "<Project Update> " & strCompany & vTemp(0) & " Fielding task assigned to " & tbFI.Value
            strBody = "<b>" & tbFI.Value & "</b> has been assigned " & vTemp(0) & " to field."
        Case "Reports"
            strBody = strBody & vbCr & vbCr & "LU/BU: 0/0" & vbCr & "DWG: 0" & vbCr & "Route Footage: 0" & vbCr
            strBody = strBody & "JU Company:" & vbTab & "Unknown" & vbCr & "xx Total Poles" & vbCr & "xx Poles to OCALC"
            strBody = strBody & vbCr & vbCr & "City Permit:" & vbTab & "Unknown" & vbCr & "TDOT Permit:" & vbTab & "Unknown"
            strBody = strBody & vbCr & "RR Permit:" & vbTab & "Unknown" & vbCr & vbCr & "Make Ready Required:" & vbCr & "None"
            strBody = strBody & vbCr & vbCr & "Notes:" & vbCr & "None"
    End Select
    
    strBody = strBody & vbCr & vbCr & "<b>THIS JOB IS USING ENGINEERING 2.0</b>"
    
    ProjectStatusEmail.tbBody.Value = Replace(strBody, "<br>", vbCr)
    ProjectStatusEmail.tbSubject.Value = strSubject
    ProjectStatusEmail.show
    'Me.show
    
    strBody = Replace(ProjectStatusEmail.tbBody.Value, vbCr, "<br>")
    If strBody = "" Then
        iEmailSent = 0
        Exit Sub
    Else
        iEmailSent = 1
    End If
    
    If ProjectStatusEmail.cbCcSelf.Value = True Then
        iEmailSelf = 1
    Else
        iEmailSelf = 0
    End If
    
    If ProjectStatusEmail.lbCC.ListCount > -1 Then
        For n = 0 To ProjectStatusEmail.lbCC.ListCount - 1
            strCC = strCC & ProjectStatusEmail.lbCC.List(n) & ";"
        Next n
    End If
    
    Unload ProjectStatusEmail
    
    Set objOutlook = New Outlook.Application
    Set objMail = objOutlook.CreateItem(olMailItem)
    
    objMail.To = strTo  '"rich.taylor@integrity-us.com"
    'objMail.To = objOutlook.Session.CurrentUser.Address
    If iEmailSelf = 1 Then strCC = strCC & objOutlook.Session.CurrentUser.Address
    objMail.CC = strCC  'objOutlook.Session.CurrentUser.Address
    objMail.Subject = strSubject
    objMail.HTMLBody = strBody
    
    'objMail.Display
    objMail.Send
End Sub

Private Sub RefreshProjectForm()
    Dim strProject As String
    Dim vName As Variant
    
    Dim str As String
    
    Dim strDWGPath As String
    Dim strDWGName As String
    Dim vItem As Variant
    Dim fName As String
    'Dim iCount As Integer
    
    'strProject = ThisDrawing.Name
    'vName = Split(strProject, " ")
    strFileName = ThisDrawing.Path & "\" & strProjectName & " Project Status.pjs"
    
    fName = Dir(strFileName)
    If fName = "" Then
        Exit Sub
    End If
    
    Open strFileName For Input As #2
    
    While Not EOF(2)
        Input #2, str
        
        vName = Split(str, vbTab)
        
        Select Case vName(0)
            Case "SETUP:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbPSI.Value = vName(1)
                    tbPSI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbPSS.Value = vName(2)
                    If Not vName(2) = "" Then tbPSS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbPSC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbPSC.Enabled = False
                        Frame12.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbPSH.Value = vName(4)
            Case "BDK:"
                If UBound(vName) > 0 Then
                    Frame13.Enabled = True
                    If vName(1) = "" Then GoTo Next_line
                    tbBI.Value = vName(1)
                    tbBI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbBS.Value = vName(2)
                    If Not vName(2) = "" Then tbBS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbBC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbBC.Enabled = False
                        Frame13.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbBH.Value = vName(4)
            Case "PLANNING:"
                If UBound(vName) > 0 Then
                    Frame11.Enabled = True
                    If vName(1) = "" Then GoTo Next_line
                    tbPLI.Value = vName(1)
                    tbPLI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbPLS.Value = vName(2)
                    If Not vName(2) = "" Then tbPLS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbPLC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbPLC.Enabled = False
                        Frame11.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbPLH.Value = vName(4)
            Case "PREPARE:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbCOI.Value = vName(1)
                    tbCOI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbCOS.Value = vName(2)
                    If Not vName(2) = "" Then tbCOS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbCOC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbCOC.Enabled = False
                        Frame2.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbCOH.Value = vName(4)
            Case "FIELD:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbFI.Value = vName(1)
                    tbFI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbFS.Value = vName(2)
                    If Not vName(2) = "" Then tbFS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbFC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbFC.Enabled = False
                        Frame1.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbFH.Value = vName(4)
            Case "CLEANUP:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbCUI.Value = vName(1)
                    tbCUI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbCUS.Value = vName(2)
                    If Not vName(2) = "" Then tbCUS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbCUC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbCUC.Enabled = False
                        Frame3.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbCUH.Value = vName(4)
            Case "CUT OUT:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbCDI.Value = vName(1)
                    tbCDI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbCDS.Value = vName(2)
                    If Not vName(2) = "" Then tbCDS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbCDC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbCDC.Enabled = False
                        Frame10.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbCDH.Value = vName(4)
            Case "MAKE READY:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbMRI.Value = vName(1)
                    tbMRI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbMRS.Value = vName(2)
                    If Not vName(2) = "" Then tbMRS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbMRC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbMRC.Enabled = False
                        Frame4.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbMRH.Value = vName(4)
            Case "PERMITS:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbPI.Value = vName(1)
                    tbPI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbPS.Value = vName(2)
                    If Not vName(2) = "" Then tbPS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbPC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbPC.Enabled = False
                        Frame5.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbPH.Value = vName(4)
            Case "COUNTS:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbCntI.Value = vName(1)
                    tbCntI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbCntS.Value = vName(2)
                    If Not vName(2) = "" Then tbCntS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbCntC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbCntC.Enabled = False
                        Frame6.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbCntH.Value = vName(4)
            Case "UNITS:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbUI.Value = vName(1)
                    tbUI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbUS.Value = vName(2)
                    If Not vName(2) = "" Then tbUS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbUC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbUC.Enabled = False
                        Frame7.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbUH.Value = vName(4)
            Case "QC:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbQCI.Value = vName(1)
                    tbQCI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbQCS.Value = vName(2)
                    If Not vName(2) = "" Then tbQCS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbQCC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbQCC.Enabled = False
                        Frame8.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbQCH.Value = vName(4)
            Case "REPORTS:"
                If UBound(vName) > 0 Then
                    If vName(1) = "" Then GoTo Next_line
                    tbRI.Value = vName(1)
                    tbRI.Enabled = False
                End If
                If UBound(vName) > 1 Then
                    tbRS.Value = vName(2)
                    If Not vName(2) = "" Then tbRS.Enabled = False
                End If
                If UBound(vName) > 2 Then
                    tbRC.Value = vName(3)
                    If Not vName(3) = "" Then
                        tbRC.Enabled = False
                        Frame9.ForeColor = &H8000000D
                    End If
                End If
                If UBound(vName) > 3 Then tbRH.Value = vName(4)
        End Select
Next_line:
    Wend
    
    Close #2
End Sub

Private Sub AddHoursToTimesheet(strJob As String, strDate As String, dHours As Double)
    Dim excelApp As Excel.Application
    Dim ws As Excel.Worksheet
    Dim strPath As String
    Dim iRow, iColumn As Integer
    Dim fName As String
    Dim strComputer, strFile As String
    
    strComputer = Environ$("Username")
    
    strFile = "C:\Users\" & strComputer & "\Documents\Work\Timesheets\Active Timesheet.xlsx"
    
    fName = Dir(strFile)
    If fName = "" Then
        MsgBox "There is no Timesheet ready."
        Exit Sub
    End If
    
    On Error Resume Next
    
    Set excelApp = CreateObject("Excel.Application")
    'excelApp.Visible = True
    excelApp.Visible = False
    
    'excelApp.Workbooks.Add
    Set ws = excelApp.ActiveWorkbook.Sheets("Sheet1")
    'Set wb = Nothing
    
    Err = 0
    Set wb = Excel.Workbooks.Open(strFile)
    If Not Err = 0 Then
        MsgBox "Error opening file"
        Exit Sub
    End If
    iRow = 5
    
    While iRow < 60
        For iColumn = 2 To 7
            If InStr(1, Sheets("Sheet1").Cells(iRow, iColumn).Value, strDate, vbTextCompare) > 0 Then GoTo Exit_While
        Next iColumn
        iRow = iRow + 1
    Wend
Exit_While:
    
    'MsgBox "Cell(" & iRow & " , " & iColumn & ")"
    
    iRow = iRow + 2
    While iRow < 60
        If InStr(1, Sheets("Sheet1").Cells(iRow, 1).Value, strJob, vbTextCompare) > 0 Then GoTo Exit_While2
        If Sheets("Sheet1").Cells(iRow, 1).Value = "" Then
            Sheets("Sheet1").Cells(iRow, 1).Value = strJob
            GoTo Exit_While2
        End If
        
        iRow = iRow + 1
    Wend
Exit_While2:
    
    If iRow < 60 Then
        Sheets("Sheet1").Cells(iRow, iColumn).Value = Sheets("Sheet1").Cells(iRow, iColumn).Value + dHours
    End If
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    excelApp.Quit
End Sub

Private Function GetLastDWG()
    Dim filterType, filterValue As Variant  '
    Dim grpCode(0) As Integer               '
    Dim grpValue(0) As Variant              '
    Dim objSS7 As AcadSelectionSet          '
    Dim objDWG As AcadBlockReference        '
    Dim vAttList, vTemp As Variant          '
    Dim iCount As Integer
    
    iCount = 0
    
    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS7 = ThisDrawing.SelectionSets.Add("objSS7")
    objSS7.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objDWG In objSS7
        vAttList = objDWG.GetAttributes
            vTemp = Split(vAttList(0).TextString, " ")
            If vTemp(0) = "DWG" Then
                If UBound(vTemp) > 0 Then
                    If CInt(vTemp(1)) > iCount Then
                        iCount = CInt(vTemp(1))
                    End If
                End If
            End If
    Next objDWG
    
    GetLastDWG = iCount
End Function

Private Function GetCleanupData(strText As String)
    Dim objSSTemp2 As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim entItem As AcadEntity
    Dim obrTemp As AcadBlockReference
    Dim vAttList, vTemp As Variant
    Dim iMTEMC, iOther, iUTC As Integer
    Dim iDWG As Integer
    
    iMTEMC = 0
    iOther = 0
    iUTC = 0
    iDWG = 0

    grpCode(0) = 2
    grpValue(0) = "iPole"
    filterType = grpCode
    filterValue = grpValue
  
    Set objSSTemp2 = ThisDrawing.SelectionSets.Add("objSSTemp2")
    objSSTemp2.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each obrTemp In objSSTemp2
        vAttList = obrTemp.GetAttributes
        If Not vAttList(0).TextString = "POLE" Then
            Select Case obrTemp.Layer
                Case "Integrity Poles-Power"
                    iMTEMC = iMTEMC + 1
                Case "Integrity Poles-Other"
                    iOther = iOther + 1
                Case "Integrity Poles-UTC"
                    iUTC = iUTC + 1
            End Select
        End If
    Next obrTemp
    
    objSSTemp2.Clear

    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    objSSTemp2.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each obrTemp In objSSTemp2
        vAttList = obrTemp.GetAttributes
        If obrTemp.Layer = "Integrity Sheets" Then
            vTemp = Split(vAttList(0).TextString, " ")
            If UBound(vTemp) > 0 Then
                If iDWG < CInt(vTemp(1)) Then iDWG = CInt(vTemp(1))
            End If
        End If
    Next obrTemp
    
    objSSTemp2.Clear
    objSSTemp2.Delete
    
    strText = strText & "<br>" & iMTEMC & " Power Poles<br>" & iUTC & " UTC Poles<br>" & iOther & " Other Poles<br>"
    strText = strText & "<br>" & iDWG & " Drawings<br>"
    
    GetCleanupData = strText
End Function

Private Function GetQCData(strText As String)
    Dim objSSTemp2 As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim entItem As AcadEntity
    'Dim obrTemp As AcadBlockReference
    Dim objText As AcadText
    Dim objMText As AcadMText
    Dim vTemp As Variant
    Dim strTemp As String
    Dim iOCALC As Integer
    
    iOCALC = 0

    grpCode(0) = 8
    grpValue(0) = "Integrity Poles-Other"
    filterType = grpCode
    filterValue = grpValue
  
    Set objSSTemp2 = ThisDrawing.SelectionSets.Add("objSSTemp2")
    objSSTemp2.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entItem In objSSTemp2
        If TypeOf entItem Is AcadText Then
            Set objText = entItem
            vTemp = Split(objText.TextString, " ")
            If UBound(vTemp) > 0 Then
                If vTemp(1) = "POLES" Then
                    iOCALC = CInt(vTemp(0))
                    GoTo Exit_For
                End If
            End If
        End If
        
        If TypeOf entItem Is AcadMText Then
            Set objMText = entItem
            vTemp = Split(objMText.TextString, " ")
            If UBound(vTemp) > 0 Then
                If vTemp(1) = "POLES" Then
                    iOCALC = CInt(vTemp(0))
                    GoTo Exit_For
                End If
            End If
        End If
    Next entItem
Exit_For:
    
    strText = strText & "<br>" & iOCALC & " Poles to OCALC"
    GetQCData = strText
End Function
