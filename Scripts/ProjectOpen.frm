VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectOpen 
   Caption         =   "Open Project"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "ProjectOpen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCleanupTemp_Click()
    If InStr(tbCopyFolder.Value, "Dropbox") > 0 Then Exit Sub
    cbDeleteFile.Enabled = True
    
    tbPath.Value = ""
    tbFolder.Value = ""
    lbFiles.Clear
    lbSubFolders.Clear
    tbNumber.Value = ""
    cbAllDWGs.Value = True
    'cbDeleteFile.Enabled = True
    
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
    
    lbSubFolders.Selected(0) = True
    lbSubFolders.ListIndex = 0
    
    'Call GetDrawingNames(CStr(strPath), CStr(strFolder), "")
    
    'If lbSubFolders.ListCount > 0 Then
        'For i = 1 To lbSubFolders.ListCount - 1
            'Call GetDrawingNames(CStr(strPath), CStr(strFolder), lbSubFolders.List(i) & "\")
        'Next i
    'End If
    
    If lbFiles.ListCount > -1 Then
        For i = 0 To lbFiles.ListCount - 1
            If InStr(lbFiles.List(i, 0), ".dwg") > 0 Then
                strTemp = strPath & strFolder
                strDWG = CheckIfOpen(CStr(strTemp), CStr(lbFiles.List(i, 0)))
                
                lbFiles.List(i, 1) = strDWG
            End If
        Next i
    End If
    
    lbFiles.SetFocus
End Sub

Private Sub GetDrawingNames(strPath As String, strFolder As String, strSub As String)
    Dim vDirectory As Variant
    Dim strTemp, strFile, strDWG As String
    Dim strResult As String
    Dim vTemp As Variant
    
    strTemp = strPath & strFolder & strSub
    'If cbAllFiles.Value = False Then strTemp = strTemp & "*.dwg"
    
    strFile = Dir(strTemp, vbHidden)
    
    Do While strFile <> ""
        'If cbAllDWGs.Value = False Then
            'If InStr(strFile, tbNumber.Value) > 0 Then
                'If strSub = "" Then
                    'lbFiles.AddItem strSub
                'Else
                    'lbFiles.AddItem Left(strSub, Len(strSub) - 1)
                'End If
                'lbFiles.List(lbFiles.ListCount - 1, 1) = strFile
            'End If
        'Else
            ''If strSub = "" Then
                ''lbFiles.AddItem strSub
            ''Else
                ''lbFiles.AddItem Left(strSub, Len(strSub) - 1)
            ''End If
            ''lbFiles.List(lbFiles.ListCount - 1, 1) = strFile
            lbFiles.AddItem strFile
            lbFiles.List(lbFiles.ListCount - 1, 1) = ""
        'End If
        
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
        
        If GetAttr(strTemp) = 48 Or GetAttr(strTemp) = 16 Then
            If Not strFile = "." Then
                If Not strFile = ".." Then lbSubFolders.AddItem strSub & "\" & strFile
            End If
        End If
        
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

Private Sub cbCopyDWG_Click()
    Dim iIndex As Integer
    
    iIndex = lbFiles.ListIndex
    If iIndex < 0 Then Exit Sub
    
    Dim fso As Object
    Dim strFileFrom, strFileTo As String
    
    strFileFrom = tbPath.Value & tbFolder.Value
    If Not lbSubFolders.List(lbSubFolders.ListIndex) = "" Then
        strFileFrom = strFileFrom & lbSubFolders.List(lbSubFolders.ListIndex) & "\"
    End If
    strFileFrom = strFileFrom & lbFiles.List(iIndex, 0)
    strFileTo = tbCopyFolder & lbFiles.List(iIndex, 0)
    
    'MsgBox "Copying from:" & vbCr & vbCr & strFileFrom & vbCr & vbCr & "Copying to:" & vbCr & vbCr & strFileTo
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Call fso.CopyFile(strFileFrom, strFileTo)
    
    If Right(strFileTo, 4) = ".dwg" Then
        ThisDrawing.Application.Documents.Open strFileTo
    Else
        CreateObject("Shell.Application").Open CVar(strFileTo)
    End If
End Sub

Private Sub cbCreateProject_Click()
    cbDeleteFile.Enabled = False
    
    Me.Hide
    
    Load ProjectCreate
        If Not tbNumber.Value = "" Then
            ProjectCreate.tbNumber.Value = tbNumber.Value
            ProjectCreate.tbTitle.SetFocus
        End If
        
        ProjectCreate.show
    Unload ProjectCreate
    
    Me.Hide
End Sub

Private Sub cbDeleteFile_Click()
    If lbFiles.ListIndex < 0 Then Exit Sub
    
    Dim iIndex As Integer
    
    iIndex = lbFiles.ListIndex
    If Not lbFiles.List(iIndex, 1) = "" Then
        If iIndex = lbFiles.ListCount - 1 Then
            If Not iIndex = 0 Then
                iIndex = iIndex - 1
                lbFiles.Selected(iIndex) = True
                lbFiles.ListIndex = iIndex
            End If
        Else
            iIndex = iIndex + 1
            lbFiles.Selected(iIndex) = True
            lbFiles.ListIndex = iIndex
        End If
        
        Exit Sub
    End If
    
    Dim strFile As String
    Dim result As Integer
    
    strFile = tbPath.Value & tbFolder.Value & lbSubFolders.List(lbSubFolders.ListIndex)
    If Not Right(strFile, 1) = "\" Then strFile = strFile & "\"
    'If Not lbFiles.List(iIndex, 0) = "" Then strFile = strFile & lbFiles.List(iIndex, 0) & "\"
    strFile = strFile & lbFiles.List(iIndex, 0)
    
    'MsgBox strFile
    'Exit Sub
    
    result = MsgBox("Are you sure you want to delete" & vbCr & strFile & "?", vbYesNo, "Delete File!")
    If result = vbYes Then
        Kill strFile
        lbFiles.RemoveItem iIndex
    End If
    
Exit_Sub:
    
End Sub

Private Sub cbGetInfo_Click()
    Dim strIntegrityPath As String
    strIntegrityPath = "C:\Integrity\"
    
    ' Check if Integrity folder exists
    If Dir(strIntegrityPath, vbDirectory) = "" Then
        MsgBox "Integrity folder not found at " & strIntegrityPath, vbExclamation
        Exit Sub
    End If
    
    ' Set initial directory and search
    Dim strDrawingPath As String
    strDrawingPath = strIntegrityPath & Me.tbNumber.Text & ".dwg"
    
    ' Check if file exists
    If Dir(strDrawingPath) <> "" Then
        ThisDrawing.Application.Documents.Open strDrawingPath
    Else
        MsgBox "Drawing not found: " & strDrawingPath, vbExclamation
    End If
    
    On Error GoTo 0
End Sub

Private Sub cbReadOnly_Click()
    If lbFiles.ListIndex < 0 Then Exit Sub
    
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & lbSubFolders.List(lbSubFolders.ListIndex)
    If Not Right(strFileName, 1) = "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & lbFiles.List(lbDrawings.ListIndex, 0)
    
    ThisDrawing.Application.Documents.Open strFileName, True
End Sub

Private Sub lbFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & lbSubFolders.List(lbSubFolders.ListIndex)
    If Not Right(strFileName, 1) = "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & lbFiles.List(lbFiles.ListIndex, 0)
    
    If Right(strFileName, 4) = ".dwg" Then
        Select Case lbFiles.List(lbFiles.ListIndex, 1)
            Case ""
                ThisDrawing.Application.Documents.Open strFileName
            Case Else
                'MsgBox strFileName
                ThisDrawing.Application.Documents.Open strFileName, True
                MsgBox lbFiles.List(lbFiles.ListIndex, 0) & vbCr & "was opened in ReadOnly."
        End Select
        'ThisDrawing.Application.Documents.Open strFileName
    Else
        CreateObject("Shell.Application").Open CVar(strFileName)
    End If
End Sub

Private Sub lbFiles_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & lbSubFolders.List(lbSubFolders.ListIndex)
    If Not Right(strFileName, 1) = "\" Then strFileName = strFileName & "\"
    
    Select Case KeyCode
        Case vbKeyReturn
            strFileName = strFileName & lbFiles.List(lbFiles.ListIndex, 0)
            If Right(strFileName, 4) = ".dwg" Then
                Select Case lbFiles.List(lbFiles.ListIndex, 1)
                    Case ""
                        ThisDrawing.Application.Documents.Open strFileName
                    Case Else
                        ThisDrawing.Application.Documents.Open strFileName, True
                        MsgBox lbFiles.List(lbFiles.ListIndex, 0) & vbCr & "was opened in ReadOnly."
                End Select
            Else
                CreateObject("Shell.Application").Open CVar(strFileName)
            End If
    End Select
End Sub

Private Sub lbSubFolders_Click()
    Dim strPath, strFolder As String
    Dim iIndex As Integer
    
    lbFiles.Clear
    iIndex = lbSubFolders.ListIndex
    
    strPath = tbPath.Value
    strFolder = tbFolder.Value
    
    lbFiles.Clear
    
    Call GetDrawingNames(CStr(strPath), CStr(strFolder), lbSubFolders.List(iIndex) & "\")
    
    If lbFiles.ListCount > -1 Then
        For i = 0 To lbFiles.ListCount - 1
            If InStr(lbFiles.List(i, 0), ".dwg") > 0 Then
                strTemp = strPath & strFolder & lbSubFolders.List(lbSubFolders.ListIndex) & "\"
                strDWG = CheckIfOpen(CStr(strTemp), CStr(lbFiles.List(i, 0)))
                
                lbFiles.List(i, 1) = strDWG
            End If
        Next i
        
        vLine = Split(tbFolder.Value, "\")
        strTemp = LCase(vLine(UBound(vLine) - 1)) & ".dwg"
        
        For i = 0 To lbFiles.ListCount - 1
            If LCase(lbFiles.List(i, 1)) = strTemp Then
                lbFiles.Selected(i) = True
                GoTo Exit_Sub
            End If
        Next i
    End If
Exit_Sub:
    
    lbFiles.SetFocus
End Sub

Private Sub UserForm_Initialize()
    lbFiles.ColumnCount = 2
    lbFiles.ColumnWidths = "264;78"
    
    tbNumber.SetFocus
End Sub

Private Function GetFolderName(strFolder As String, strNumber As String)
    Dim vDirectory As Variant
    Dim strFile As String
    Dim strResult As String
    Dim vTemp As Variant
    
    strResult = "<not found>"
    vDirectory = Dir(strFolder, vbDirectory)
    
    Do While vDirectory <> ""
        If InStr(vDirectory, strNumber) > 0 Then
            strResult = CStr(vDirectory)
            Exit Do
        End If
        vDirectory = Dir$
    Loop
    GetFolderName = strResult
End Function

