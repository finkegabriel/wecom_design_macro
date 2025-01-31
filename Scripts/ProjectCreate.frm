VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectCreate 
   Caption         =   "Create Project Folders"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895.001
   OleObjectBlob   =   "ProjectCreate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCreateProject_Click()
    Call UpdateFolderNames
End Sub

Private Sub cbType_Change()
    Select Case cbType.Value
        Case "All"
            For i = 0 To lbFolders.ListCount - 1
                lbFolders.Selected(i) = True
            Next i
            
            'For i = 0 To lbFiles.ListCount - 1
                'lbFiles.Selected(i) = True
            'Next i
        Case "All Folders"
            For i = 0 To lbFolders.ListCount - 1
                If InStr(lbFolders.List(i, 1), "\") > 0 Then lbFolders.Selected(i) = True
            Next i
            
            'For i = 0 To lbFiles.ListCount - 1
                'lbFiles.Selected(i) = False
            'Next i
        Case "None"
            For i = 0 To lbFolders.ListCount - 1
                lbFolders.Selected(i) = False
            Next i
            
            'For i = 0 To lbFiles.ListCount - 1
                'lbFiles.Selected(i) = False
            'Next i
    End Select
End Sub

Private Sub lbFolders_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call FindMotherFolders
End Sub

Private Sub tbNumber_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    tbPath.Value = ""
    tbFolder.Value = ""
    
    If tbNumber.Value = "" Then Exit Sub
    
    Call GetFolderFile
    
    Dim vLine As Variant
    Dim strPath, strFolder, strDWG, strDWL As String
    Dim strAll, strNumber, strTemp As String
    Dim strUser As String
    
    strUser = Environ("USERNAME")
    tbNumber.Value = UCase(tbNumber.Value)
    strNumber = tbNumber.Value
    
    vLine = Split(strNumber, "20")
    Select Case vLine(0)
        Case ""
            strPath = "C:\Users\" & strUser & "\Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\1-JOBS\"
            vLine = Split(strNumber, "-")
            
            Select Case vLine(0)
                Case "2019"
                    strPath = strPath & "1-2019 JOBS\"
                Case "2020"
                    strPath = strPath & "2-2020 JOBS\"
                Case "2021"
                    strPath = strPath & "3-2021 JOBS\"
                Case "2022"
                    strPath = strPath & "4-2022 JOBS\"
                Case "2023"
                    strPath = strPath & "5-2023 JOBS\"
            End Select
        Case "L"
            strPath = "C:\Users\" & strUser & "\Dropbox\LORETTO TEL & KCW SHARED FOLDER\01 - JOBS\"
            strNumber = Replace(strNumber, "L", "")
            vLine = Split(strNumber, "-")
            
            Select Case vLine(0)
                Case "2019"
                    strPath = strPath & "2019\"
                Case "2020"
                    strPath = strPath & "2020\"
                Case "2021"
                    strPath = strPath & "2021\"
                Case "2022"
                    strPath = strPath & "2022\"
                Case "2023"
                    strPath = strPath & "2023\"
            End Select
        Case "MAS"
            strPath = "C:\Users\" & strUser & "\Dropbox\MASTEC JOB INFORMATION\1 - JOBS\"
            strNumber = Replace(strNumber, "MAS", "")
            vLine = Split(strNumber, "-")
            
            Select Case vLine(0)
                Case "2019"
                    strPath = strPath & "2019\"
                Case "2020"
                    strPath = strPath & "2020\"
                Case "2021"
                    strPath = strPath & "1-2021 JOBS\"
                Case "2022"
                    strPath = strPath & "2-2022 JOBS\"
                Case "2023"
                    strPath = strPath & "3-2023 JOBS\"
            End Select
        
    End Select
    
    tbPath.Value = strPath
    
    strFolder = GetFolderName(CStr(strPath), CStr(strNumber))
    If strFolder = "<not found>" Then
        tbFolder.Value = ""
        tbTitle.SetFocus
        Exit Sub
    End If
    
    tbTitle.Value = Right(strFolder, Len(strFolder) - Len(tbNumber.Value))
    strFolder = strFolder & "\"
    tbFolder.Value = strFolder
    
    Call UpdateFolderNames
    Call FindIfFolderExists
    
    'For i = 0 To lbFolders.ListCount - 1
        'lbFolders.List(i) = Replace(lbFolders.List(i), "<<pn>>", tbNumber.Value)
        'lbFolders.List(i) = Replace(lbFolders.List(i), "<<pt>>", tbNumber.Value)
    'Next i
    cbCreateProject.SetFocus
End Sub

Private Sub tbTitle_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not tbFolder.Value = "" Then Exit Sub
    If tbNumber.Value = "" Then Exit Sub
    If tbTitle.Value = "" Then Exit Sub
    
    Dim strLine, strFolder, fName As String
    
    strFolder = tbNumber.Value & " " & tbTitle.Value & "\"
    tbFolder.Value = strFolder
    
    strFolder = tbPath.Value & strFolder
    'strLine = Dir(strFolder, vbDirectory)
    
    Call UpdateFolderNames
    
    cbCreateProject.Enabled = True
    cbCreateProject.SetFocus
    
    
End Sub

Private Sub UserForm_Initialize()
    'lbFiles.ColumnCount = 2
    'lbFiles.ColumnWidths = "192;156"
    
    lbFolders.ColumnCount = 2
    lbFolders.ColumnWidths = "72;354"
    
    cbType.AddItem "All"
    cbType.AddItem "All Folders"
    cbType.AddItem "None"
    
    Call GetFolderFile
End Sub

Private Sub GetFolderFile()
    Dim vTemp As Variant
    Dim strUser, strPath As String
    Dim fName, strLine As String
    Dim iStatus As Integer
    
    lbFolders.Clear
    'lbFiles.Clear
    
    iStatus = 0
    strUser = Environ("USERNAME")
    
    strPath = "C:\Users\" & strUser & "\Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\References\"
    strPath = strPath & "Project Folders.txt"
    
    fName = Dir(strPath)
    If fName = "" Then
        Exit Sub
    End If
    
    Open strPath For Input As #1
    
    While Not EOF(1)
        Line Input #1, strLine
        
        Select Case strLine
            Case "{{FOLDERS}}"
                iStatus = 1
            Case "{{FILES}}"
                iStatus = 2
            Case "{{NOTES}}"
                iStatus = 3
            Case ""
                'Do nothing
            Case Else
                Select Case iStatus
                    Case Is = 1
                        strLine = Replace(strLine, vbTab, "   ")
                        lbFolders.AddItem ""
                        lbFolders.List(lbFolders.ListCount - 1, 1) = strLine
                    Case Is = 2
                        'vTemp = Split(strLine, vbTab)
                        'lbFiles.AddItem vTemp(1)
                        'lbFiles.List(lbFiles.ListCount - 1, 1) = vTemp(0)
                End Select
        End Select
    Wend
    
    Close #1
End Sub

Private Function GetFolderName(strFolder As String, strNumber As String)
    Dim vDirectory As Variant
    Dim strFile As String
    Dim strResult As String
    Dim vTemp As Variant
    
    strResult = "<not found>"
    vDirectory = Dir(strFolder, vbDirectory)
    
    'MsgBox vDirectory
    
    Do While vDirectory <> ""
        If InStr(vDirectory, strNumber) > 0 Then
            strResult = CStr(vDirectory)
            GoTo Exit_Function
        End If
        vDirectory = Dir$
    Loop
Exit_Function:
    GetFolderName = strResult
End Function

Private Sub UpdateFolderNames()
    If tbNumber.Value = "" Then Exit Sub
    If tbTitle.Value = "" Then Exit Sub
    
    For i = 0 To lbFolders.ListCount - 1
        lbFolders.List(i, 1) = Replace(lbFolders.List(i, 1), "<<pn>>", tbNumber.Value)
        lbFolders.List(i, 1) = Replace(lbFolders.List(i, 1), "<<pt>>", tbTitle.Value)
    Next i
    
    'For i = 0 To lbFiles.ListCount - 1
        'lbFiles.List(i, 0) = Replace(lbFiles.List(i, 0), "<<pn>>", tbNumber.Value)
        'lbFiles.List(i, 0) = Replace(lbFiles.List(i, 0), "<<pt>>", tbTitle.Value)
        
        'lbFiles.List(i, 1) = Replace(lbFiles.List(i, 1), "<<pn>>", tbNumber.Value)
        'lbFiles.List(i, 1) = Replace(lbFiles.List(i, 1), "<<pt>>", tbTitle.Value)
    'Next i
End Sub

Private Sub FindMotherFolders()
    If lbFolders.ListCount < 2 Then Exit Sub
    
    Dim iSubfolder, iFolder As Integer
    
    For i = lbFolders.ListCount - 1 To 1 Step -1
        If lbFolders.Selected(i) = False Then GoTo Next_line
        
        If Left(lbFolders.List(i, 1), 9) = "         " Then
            iSubfolder = 3
        ElseIf Left(lbFolders.List(i, 1), 6) = "      " Then
            iSubfolder = 2
        ElseIf Left(lbFolders.List(i, 1), 3) = "   " Then
            iSubfolder = 1
        Else
            GoTo Next_line
        End If
        
        For j = i - 1 To 0 Step -1
            If Left(lbFolders.List(j, 1), 6) = "      " Then
                iFolder = 2
            ElseIf Left(lbFolders.List(j, 1), 3) = "   " Then
                iFolder = 1
            Else
                iFolder = 0
            End If
            
            If iFolder < iSubfolder Then
                lbFolders.Selected(j) = True
                i = j + 1
                GoTo Next_line
            End If
        Next j
Next_line:
    Next i
End Sub

Private Sub FindIfFolderExists()
    If lbFolders.ListCount < 1 Then Exit Sub
    
    Dim strMainFolder, strSub(3) As String
    Dim fName, strLine As String
    Dim iSelected, iSubfolder  As Integer
    
    strMainFolder = tbPath.Value & tbFolder.Value
    strSubFolders = ""
    
    For i = 0 To lbFolders.ListCount - 1
        'If InStr(lbFolders.List(i, 1), "*F*") > 0 Then GoTo Next_Line
        
        If Left(lbFolders.List(i, 1), 9) = "         " Then
            strSub(3) = Right(lbFolders.List(i, 1), Len(lbFolders.List(i, 1)) - 9)
            strLine = strMainFolder & strSub(0) & "\" & strSub(1) & "\" & strSub(2) & "\" & strSub(3)
            If InStr(strLine, "*F* ") > 0 Then
                strLine = Replace(strLine, "*F* ", "")
            End If
        ElseIf Left(lbFolders.List(i, 1), 6) = "      " Then
            strSub(2) = Right(lbFolders.List(i, 1), Len(lbFolders.List(i, 1)) - 6)
            strLine = strMainFolder & strSub(0) & "\" & strSub(1) & "\" & strSub(2)
            If InStr(strLine, "*F* ") > 0 Then
                strLine = Replace(strLine, "*F* ", "")
            End If
        ElseIf Left(lbFolders.List(i, 1), 3) = "   " Then
            strSub(1) = Right(lbFolders.List(i, 1), Len(lbFolders.List(i, 1)) - 3)
            strLine = strMainFolder & strSub(0) & "\" & strSub(1)
            If InStr(strLine, "\") > 0 Then
                strLine = Replace(strLine, "*F* ", "")
            End If
        Else
            strSub(0) = lbFolders.List(i, 1)
            strLine = strMainFolder & strSub(0)
            If InStr(strLine, "*F* ") > 0 Then
                strLine = Replace(strLine, "*F* ", "")
            End If
        End If
    
        'MsgBox strLine
        fName = Dir(strLine, vbDirectory)
        
        If Not fName = "" Then
            lbFolders.List(i, 0) = "Y"
        End If
    Next i
End Sub

'Private Sub CheckIfProjectExists()

'End Sub
