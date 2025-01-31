VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OpenProject 
   Caption         =   "Open File(s)"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "OpenProject.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OpenProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAllDWGs_Click()
    If cbAllDWGs.Value = False Then cbAllFiles.Value = False
End Sub

Private Sub cbAllFiles_Click()
    If cbAllFiles.Value = True Then cbAllDWGs.Value = True
End Sub

Private Sub cbCleanupTemp_Click()
    If InStr(tbCopyFolder.Value, "Dropbox") > 0 Then Exit Sub
    
    tbPath.Value = ""
    tbFolder.Value = ""
    lbDrawings.Clear
    lbSubFolders.Clear
    tbNumber.Value = ""
    cbAllDWGs.Value = True
    cbDeleteFile.Enabled = True
    
    Dim vLine As Variant
    Dim strPath, strFolder, strDWG, strDWL As String
    Dim strAll, strNumber, strTemp As String
    Dim strUser As String
    Dim iIndex As Integer
    
    strPath = "C:\Integrity\"
    strFolder = "Temp\"
    
    tbPath.Value = strPath
    tbFolder.Value = strFolder
    
    Call GetSubFolders
    
    If lbSubFolders.ListCount > 0 Then
        iIndex = 1
        While iIndex < lbSubFolders.ListCount
            Call GetSubSubFolders(lbSubFolders.List(iIndex))
            
            iIndex = iIndex + 1
        Wend
        
        Call SortSubFolders
    End If
    
    Call GetDrawingNames(CStr(strPath), CStr(strFolder), "")
    If Not lbSubFolders.ListCount < 0 Then
        For i = 0 To lbSubFolders.ListCount - 1
            Call GetDrawingNames(CStr(strPath), CStr(strFolder), lbSubFolders.List(i) & "\")
        Next i
    End If
    
    If lbDrawings.ListCount > -1 Then
        For i = 0 To lbDrawings.ListCount - 1
            strTemp = strPath & strFolder & lbSubFolders.List(lbSubFolders.ListIndex) & "\"
            strDWG = CheckIfOpen(CStr(strTemp), CStr(lbDrawings.List(i, 1)))
            
            lbDrawings.List(i, 2) = strDWG
        Next i
    End If
    
End Sub

Private Sub cbDeleteFile_Click()
    If lbDrawings.ListIndex < 0 Then Exit Sub
    
    Dim iIndex As Integer
    
    iIndex = lbDrawings.ListIndex
    If Not lbDrawings.List(iIndex, 2) = "" Then
        If iIndex = lbDrawings.ListCount - 1 Then
            If Not iIndex = 0 Then
                iIndex = iIndex - 1
                lbDrawings.Selected(iIndex) = True
                lbDrawings.ListIndex = iIndex
            End If
        Else
            iIndex = iIndex + 1
            lbDrawings.Selected(iIndex) = True
            lbDrawings.ListIndex = iIndex
        End If
        
        Exit Sub
    End If
    
    Dim strFile As String
    Dim result As Integer
    
    strFile = tbPath.Value & tbFolder.Value
    If Not lbDrawings.List(iIndex, 0) = "" Then strFile = strFile & lbDrawings.List(iIndex, 0) & "\"
    strFile = strFile & lbDrawings.List(iIndex, 1)
    
    'MsgBox strFile
    
    result = MsgBox("Are you sure you want to delete" & vbCr & strFile & "?", vbYesNo, "Delete File!")
    If result = vbYes Then
        Kill strFile
        lbDrawings.RemoveItem iIndex
    End If
    
Exit_Sub:
    
End Sub

Private Sub cbGetInfo_Click()
    tbPath.Value = ""
    tbFolder.Value = ""
    lbDrawings.Clear
    lbSubFolders.Clear
    cbDeleteFile.Enabled = False
    'cbAllFiles.Value = False
    
    If tbNumber.Value = "" Then Exit Sub
    
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
    'strFolder = GetFolderName(CStr(strPath), CStr(tbNumber.Value))
    If strFolder = "<not found>" Then
        tbFolder.Value = strFolder
        Exit Sub
    End If
    
    strFolder = strFolder & "\"
    tbFolder.Value = strFolder
    'strDWG = GetDrawingName(CStr(strPath), CStr(strFolder), CStr(strNumber))
    Call GetSubFolders
    
    If lbSubFolders.ListCount > 1 Then
        iIndex = 1
        While iIndex < lbSubFolders.ListCount
            Call GetSubSubFolders(lbSubFolders.List(iIndex))
            
            iIndex = iIndex + 1
        Wend
        
        Call SortSubFolders
    End If
    
    Call GetDrawingNames(CStr(strPath), CStr(strFolder), "")
    
    If cbAllFiles.Value = False Then
        If lbSubFolders.ListCount > 1 Then
            For i = 1 To lbSubFolders.ListCount - 1
                Call GetDrawingNames(CStr(strPath), CStr(strFolder), lbSubFolders.List(i) & "\")
            Next i
        End If
    End If
    
    If lbDrawings.ListCount > -1 Then
        For i = 0 To lbDrawings.ListCount - 1
            If InStr(lbDrawings.List(i, 1), ".dwg") > 0 Then
                strTemp = strPath & strFolder & lbDrawings.List(i, 0) & "\"
                strDWG = CheckIfOpen(CStr(strTemp), CStr(lbDrawings.List(i, 1)))
                
                lbDrawings.List(i, 2) = strDWG
            End If
        Next i
        
        vLine = Split(tbFolder.Value, "\")
        strTemp = LCase(vLine(UBound(vLine) - 1)) & ".dwg"
        
        For i = 0 To lbDrawings.ListCount - 1
            If LCase(lbDrawings.List(i, 1)) = strTemp Then
                lbDrawings.Selected(i) = True
                If lbDrawings.List(i, 0) = "" Then GoTo Exit_Sub
                If InStr(lbDrawings.List(i, 0), "CONSTRUCTION DRAWINGS") > 0 Then GoTo Exit_Sub
            End If
        Next i
    End If
Exit_Sub:
    
End Sub

Private Sub cbCopyDWG_Click()
    Dim iIndex As Integer
    
    iIndex = lbDrawings.ListIndex
    If iIndex < 0 Then Exit Sub
    
    Dim fso As Object
    Dim strFileFrom, strFileTo As String
    
    strFileFrom = tbPath.Value & tbFolder.Value
    If Not lbDrawings.List(iIndex, 0) = "" Then
        strFileFrom = strFileFrom & lbDrawings.List(iIndex, 0) & "\"
    End If
    strFileFrom = strFileFrom & lbDrawings.List(iIndex, 1)
    strFileTo = tbCopyFolder & lbDrawings.List(iIndex, 1)
    
    'MsgBox "Copying from:" & vbCr & vbCr & strFileFrom & vbCr & vbCr & "Copying to:" & vbCr & vbCr & strFileTo
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Call fso.CopyFile(strFileFrom, strFileTo)
    
    If Right(strFileTo, 4) = ".dwg" Then
        ThisDrawing.Application.Documents.Open strFileTo
    Else
        CreateObject("Shell.Application").Open CVar(strFileTo)
    End If
End Sub

Private Sub cbOpen_Click()
    If lbDrawings.ListIndex < 0 Then Exit Sub
    If Not lbDrawings.List(lbDrawings.ListIndex, 2) = "" Then Exit Sub
    
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & lbDrawings.List(lbDrawings.ListIndex, 0)
    If Not Right(strFileName, 1) = "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & lbDrawings.List(lbDrawings.ListIndex, 1)
    
    If Right(strFileName, 4) = ".dwg" Then
        ThisDrawing.Application.Documents.Open strFileName
    Else
        CreateObject("Shell.Application").Open CVar(strFileName)
    End If
End Sub

Private Sub cbReadOnly_Click()
    If lbDrawings.ListIndex < 0 Then Exit Sub
    
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & lbDrawings.List(lbDrawings.ListIndex, 0)
    If Not Right(strFileName, 1) = "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & lbDrawings.List(lbDrawings.ListIndex, 1)
    
    ThisDrawing.Application.Documents.Open strFileName, True
End Sub

Private Sub lbDrawings_Click()
    cbReadOnly.Enabled = True
    cbOpen.Enabled = True
    cbCopyDWG.Enabled = True
    
    If Not lbDrawings.List(lbDrawings.ListIndex, 2) = "" Then cbOpen.Enabled = False
End Sub

Private Sub lbDrawings_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & lbDrawings.List(lbDrawings.ListIndex, 0)
    If Not Right(strFileName, 1) = "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & lbDrawings.List(lbDrawings.ListIndex, 1)
    
    If Right(strFileName, 4) = ".dwg" Then
        Select Case lbDrawings.List(lbDrawings.ListIndex, 2)
            Case ""
                ThisDrawing.Application.Documents.Open strFileName
            Case Else
                'MsgBox strFileName
                ThisDrawing.Application.Documents.Open strFileName, True
                MsgBox lbDrawings.List(lbDrawings.ListIndex, 1) & vbCr & "was opened in ReadOnly."
        End Select
        'ThisDrawing.Application.Documents.Open strFileName
    Else
        CreateObject("Shell.Application").Open CVar(strFileName)
    End If
End Sub

Private Sub lbDrawings_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & lbDrawings.List(lbDrawings.ListIndex, 0)
    If Not Right(strFileName, 1) = "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & lbDrawings.List(lbDrawings.ListIndex, 1)
    
    Select Case KeyCode
        Case vbKeyReturn
            If Right(strFileName, 4) = ".dwg" Then
                Select Case lbDrawings.List(lbDrawings.ListIndex, 2)
                    Case ""
                        ThisDrawing.Application.Documents.Open strFileName
                    Case Else
                        ThisDrawing.Application.Documents.Open strFileName, True
                        MsgBox lbDrawings.List(lbDrawings.ListIndex, 1) & vbCr & "was opened in ReadOnly."
                End Select
            Else
                CreateObject("Shell.Application").Open CVar(strFileName)
            End If
    End Select
End Sub

Private Sub lbSubFolders_Click()
    Dim strPath, strFolder As String
    Dim iIndex As Integer
    
    lbDrawings.Clear
    iIndex = lbSubFolders.ListIndex
    'For i = 0 To lbSubFolders.ListCount - 1
        'MsgBox lbSubFolders.List(i) & vbCr & lbSubFolders.Selected(i)
        'If lbSubFolders.Selected(i) = True Then
            'iIndex = i
            'GoTo Found_Selected
        'End If
    'Next i
    
'Found_Selected:
    'MsgBox iIndex
    If iIndex < 0 Then Exit Sub
    
    strPath = tbPath.Value
    strFolder = tbFolder.Value
    
    Call GetDrawingNames(CStr(strPath), CStr(strFolder), lbSubFolders.List(iIndex) & "\")
    
    If lbDrawings.ListCount > -1 Then
        For i = 0 To lbDrawings.ListCount - 1
            If InStr(lbDrawings.List(i, 1), ".dwg") > 0 Then
                strTemp = strPath & strFolder & lbDrawings.List(i, 0) & "\"
                strDWG = CheckIfOpen(CStr(strTemp), CStr(lbDrawings.List(i, 1)))
                
                lbDrawings.List(i, 2) = strDWG
            End If
        Next i
        
        vLine = Split(tbFolder.Value, "\")
        strTemp = LCase(vLine(UBound(vLine) - 1)) & ".dwg"
        
        For i = 0 To lbDrawings.ListCount - 1
            If LCase(lbDrawings.List(i, 1)) = strTemp Then
                lbDrawings.Selected(i) = True
                GoTo Exit_Sub
            End If
        Next i
    End If
Exit_Sub:
    
End Sub

Private Sub lbSubFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strPath, strFolder As String
    Dim iIndex As Integer
    
    lbDrawings.Clear
    iIndex = lbSubFolders.ListIndex
    'For i = 0 To lbSubFolders.ListCount - 1
        'MsgBox lbSubFolders.List(i) & vbCr & lbSubFolders.Selected(i)
        'If lbSubFolders.Selected(i) = True Then
            'iIndex = i
            'GoTo Found_Selected
        'End If
    'Next i
    
'Found_Selected:
    'MsgBox iIndex
    If iIndex < 0 Then Exit Sub
    
    strPath = tbPath.Value
    strFolder = tbFolder.Value
    
    Call GetDrawingNames(CStr(strPath), CStr(strFolder), lbSubFolders.List(iIndex) & "\")
    
    If lbDrawings.ListCount > -1 Then
        For i = 0 To lbDrawings.ListCount - 1
            If InStr(lbDrawings.List(i, 1), ".dwg") > 0 Then
                strTemp = strPath & strFolder & lbDrawings.List(i, 0) & "\"
                strDWG = CheckIfOpen(CStr(strTemp), CStr(lbDrawings.List(i, 1)))
                
                lbDrawings.List(i, 2) = strDWG
            End If
        Next i
        
        vLine = Split(tbFolder.Value, "\")
        strTemp = LCase(vLine(UBound(vLine) - 1)) & ".dwg"
        
        For i = 0 To lbDrawings.ListCount - 1
            If LCase(lbDrawings.List(i, 1)) = strTemp Then
                lbDrawings.Selected(i) = True
                GoTo Exit_Sub
            End If
        Next i
    End If
Exit_Sub:
    
End Sub

Private Sub UserForm_Initialize()
    lbDrawings.ColumnCount = 3
    lbDrawings.ColumnWidths = "180;264;42"
    
    tbNumber.SetFocus
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

Private Sub GetDrawingNames(strPath As String, strFolder As String, strSub As String)
    Dim vDirectory As Variant
    Dim strTemp, strFile, strDWG As String
    Dim strResult As String
    Dim vTemp As Variant
    
    strTemp = strPath & strFolder & strSub
    If cbAllFiles.Value = False Then strTemp = strTemp & "*.dwg"
    
    strFile = Dir(strTemp, vbHidden)
    
    Do While strFile <> ""
        If cbAllDWGs.Value = False Then
            If InStr(strFile, tbNumber.Value) > 0 Then
                If strSub = "" Then
                    lbDrawings.AddItem strSub
                Else
                    lbDrawings.AddItem Left(strSub, Len(strSub) - 1)
                End If
                lbDrawings.List(lbDrawings.ListCount - 1, 1) = strFile
            End If
        Else
            'lbDrawings.AddItem Replace(strSub, "\", "")
            If strSub = "" Then
                lbDrawings.AddItem strSub
            Else
                lbDrawings.AddItem Left(strSub, Len(strSub) - 1)
            End If
            lbDrawings.List(lbDrawings.ListCount - 1, 1) = strFile
        End If
        
        strFile = Dir()
    Loop
End Sub

Private Sub GetSubFolders()
    Dim strFolder, strFile As String
    Dim strTemp As String
    
    lbSubFolders.Clear
    lbSubFolders.AddItem ""
    strFolder = tbPath.Value & tbFolder.Value
    
    strFile = Dir(strFolder, vbDirectory)
    
    Do While strFile <> ""
        strTemp = strFolder & strFile
        'MsgBox strFile & vbCr & GetAttr(strTemp)
        
        If GetAttr(strTemp) = 48 Or GetAttr(strTemp) = 16 Then
            If Not strFile = "." Then
                If Not strFile = ".." Then lbSubFolders.AddItem strFile
            End If
        End If
        
        strFile = Dir()
    Loop
End Sub

Private Sub GetSubSubFolders(strSub As String)
    Dim strFolder, strFile As String
    Dim strTemp As String
    Dim vCount As Variant
    
    If Left(strSub, 1) = "\" Then strSub = Right(strSub, Len(strSub) - 1)
    
    strFolder = tbPath.Value & tbFolder.Value & strSub & "\"
    vCount = Split(strFolder, "\")
    If UBound(vCount) > 10 Then Exit Sub
    
    strFile = Dir(strFolder, vbDirectory)
    
    Do While strFile <> ""
        strTemp = strFolder & strFile
        'MsgBox strTemp & vbCr & GetAttr(strTemp)
        'Exit Sub
        
        If Len(strFile) > 70 Then GoTo Next_strFile
        
        If GetAttr(strTemp) = 48 Or GetAttr(strTemp) = 16 Then
            If Not strFile = "." Then
                If Not strFile = ".." Then lbSubFolders.AddItem strSub & "\" & strFile
            End If
        End If
        
Next_strFile:
        strFile = Dir()
    Loop
End Sub

Private Sub SortSubFolders()
    Dim a, b As Integer
    Dim iCount As Integer
    Dim strAtt As String
    'Dim strAtt(0 To 2) As String
    
    iCount = lbSubFolders.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbSubFolders.List(b, 0) > lbSubFolders.List(b + 1, 0) Then
                'If Not Err = 0 Then
                    'MsgBox "Error sorting list"
                    'lbSubFolders.Selected(b) = True
                    'lbSubFolders.ListIndex = b
                    'Exit Sub
                'End If
                
                strAtt = lbSubFolders.List(b + 1, 0)
                lbSubFolders.List(b + 1, 0) = lbSubFolders.List(b, 0)
                lbSubFolders.List(b, 0) = strAtt
            End If
        Next b
    Next a
End Sub

Private Function CheckIfOpen(strFolder As String, strNumber As String)
    If InStr(strNumber, ".dwg") = 0 Then Exit Function
    
    Dim strFile, strDWL As String
    Dim strLine As String
    Dim vTemp As Variant
    
    'strFolder = tbPath.Value & tbFolder.Value
    strDWL = Replace(strNumber, ".dwg", ".dwl")
    
    strFile = Dir$(strFolder, vbHidden)
    
    Do While strFile <> ""
        If strFile = strDWL Then GoTo Found_File
        
        strFile = Dir$()
    Loop
    
    CheckIfOpen = ""
    
    Exit Function
Found_File:
    
    strDWL = strFolder & strFile
    
    Open strDWL For Input As #1
    
    Line Input #1, strLine
    vTemp = Split(strLine, vbLf)
    vTemp(0) = Replace(vTemp(0), vbCr, "")
    
    Close #1
    
    Select Case LCase(vTemp(0))
        Case "integrity"
            strLine = "Dylan Spears"
        Case "integrity.tab"
            strLine = "Jeremy Pafford"
        Case "integrity.1"
            strLine = "Ronn Elliott"
        Case "integrity.2"
            strLine = "Rich Taylor"
        Case "integrity.3"
            strLine = "Jason Pafford"
        Case "integrity.4"
            strLine = "Byron Auer"
        Case "integrity.5"
            strLine = "Adam Kemper"
        Case "integrity.6"
            strLine = "Jon Wilburn"
        Case "integrity.7"
            strLine = "Tara Taylor"
        Case "integrity8"
            strLine = "Franklin Angulo"
        Case "integrity9"
            strLine = "Wade Hampton"
        Case "integrity10"
            strLine = "Sam Jackson"
        'Case "integrity11"
            'strLine = "A Ghost?"
        Case "integrity12"
            strLine = "Daniel Campbell"
        Case "integrity13"
            strLine = "Nick Lockyear"
        'Case "integrity14"
            'strLine = "A Ghost?"
        Case "integrity.15"
            strLine = "Jay Penny"
        Case "integrity16"
            strLine = "Drew Curtis"
        Case Else
            strLine = LCase(vTemp(0)) & " ?"
    End Select
    
    CheckIfOpen = strLine
End Function

