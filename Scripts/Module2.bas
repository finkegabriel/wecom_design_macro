Attribute VB_Name = "Module2"

Public Sub LaunchMainMenu()
    Dim strPath As String
    
    strPath = LCase(ThisDrawing.Application.vbe.ActiveVBProject.Filename)
    
    If InStr(strPath, "integrity") = 0 Then Exit Sub
    
    'MsgBox ThisDrawing.Path & vbCr & ThisDrawing.Name
    If InStr(ThisDrawing.Name, "Drawing") > 0 Then
        If InStr(LCase(ThisDrawing.Path), "dropbox") = 0 Then
            Load ProjectOpen
                ProjectOpen.show
            Unload ProjectOpen
            
            Exit Sub
        End If
    End If
    
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objText As AcadText
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim iVersion As Integer
    
    iVersion = 1
    
    grpCode(0) = 8
    grpValue(0) = "Integrity Border Info"
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Add("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadText Then
            Set objText = objEntity
            If objText.TextString = "THIS JOB IS USING ENGINEERING 2.0" Then
                iVersion = 2
                GoTo Found_Version
            End If
        End If
    Next objEntity
    
Found_Version:
    
    objSS.Clear
    objSS.Delete
    
    Select Case iVersion
        Case Is = 1
            Application.LoadDVB "C:\Integrity\VBA\UTC.dvb"
            Application.RunMacro "AddModule.startMainForm"
        Case Is = 2
            Load MainForm
                MainForm.show
            Unload MainForm
    End Select
End Sub

Public Sub startMainForm()
    Dim strPath As String
    
    strPath = LCase(ThisDrawing.Application.vbe.ActiveVBProject.Filename)
    
    'MsgBox strPath
    If InStr(strPath, "integrity") = 0 Then Exit Sub
    
    Load MainForm
    MainForm.show
    Unload MainForm
End Sub

Public Sub United_Request()
    
End Sub
