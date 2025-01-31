VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OpenDrawing 
   Caption         =   "Open Drawing"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   OleObjectBlob   =   "OpenDrawing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OpenDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGetInfo_Click()
    tbPath.Value = ""
    tbFolder.Value = ""
    tbDrawingName.Value = ""
    tbOComputer.Value = ""
    tbOName.Value = ""
    
    Dim vLine As Variant
    Dim strPath, strFolder, strDWG, strDWL As String
    Dim strAll, strNumber, strTemp As String
    Dim strUser As String
    
    strUser = Environ("USERNAME")
    tbNumber.Value = UCase(tbNumber.Value)
    strNumber = tbNumber.Value
    'strNumber = Replace(tbNumber.Value, "L", "")
    'strNumber = Replace(strNumber, "MAS", "")
    
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
    
    'strAll = strPath & strFolder
    strFolder = GetFolderName(CStr(strPath), CStr(strNumber))
    'strFolder = GetFolderName(CStr(strPath), CStr(tbNumber.Value))
    If strFolder = "<not found>" Then
        tbFolder.Value = strFolder
        Exit Sub
    End If
    
    strFolder = strFolder & "\"
    tbFolder.Value = strFolder
    strDWG = GetDrawingName(CStr(strPath), CStr(strFolder), CStr(strNumber))
    
    'tbFolder.Value = strFolder
    tbDrawingName.Value = strDWG
    
    Call CheckIfOpen
    
    If tbDrawingName.Value = "<not found>" Then
        cbOpen.Enabled = False
        cbReadOnly.Enabled = False
    Else
        cbOpen.Enabled = True
        cbReadOnly.Enabled = True
        
        cbOpen.SetFocus
    End If
    
    If Not tbOComputer.Value = "<none>" Then
        cbOpen.Enabled = False
        cbReadOnly.SetFocus
    End If
End Sub

Private Sub cbOpen_Click()
    If tbDrawingName.Value = "<not found>" Then Exit Sub
    
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & tbDrawingName.Value
    
    ThisDrawing.Application.Documents.Open strFileName, True
    
    Me.Hide
End Sub

Private Sub cbReadOnly_Click()
    If tbDrawingName.Value = "<not found>" Then Exit Sub
    
    Dim strFileName As String
    
    strFileName = tbPath.Value & tbFolder.Value & tbDrawingName.Value
    
    ThisDrawing.Application.Documents.Open strFileName, True
    
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
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

Private Function GetDrawingName(strPath As String, strFolder As String, strNumber As String)
    Dim vDirectory As Variant
    Dim strTemp, strFile As String
    Dim strResult As String
    Dim vTemp As Variant
    
    strTemp = strPath & strFolder
    strResult = "<not found>"
    vDirectory = Dir(strTemp, vbDirectory)
    
    'MsgBox vDirectory
    
    Do While vDirectory <> ""
        If InStr(vDirectory, strNumber) > 0 Then
            If Right(vDirectory, 4) = ".dwg" Then
                strResult = CStr(vDirectory)
                GoTo Exit_Function
            End If
        End If
        vDirectory = Dir$
    Loop
    
    strFolder = strFolder & tbNumber.Value & " CONSTRUCTION DRAWINGS\"
    tbFolder.Value = strFolder
    strTemp = strPath & strFolder
    vDirectory = Dir(strTemp, vbDirectory)
    
    'MsgBox vDirectory
    
    Do While vDirectory <> ""
        If InStr(vDirectory, strNumber) > 0 Then
            If Right(vDirectory, 4) = ".dwg" Then
                strResult = CStr(vDirectory)
                GoTo Exit_Function
            End If
        End If
        vDirectory = Dir$
    Loop
    
    
Exit_Function:
    GetDrawingName = strResult
End Function

Private Sub CheckIfOpen()
    Dim strFolder, strFile, strDWL As String
    Dim strLine As String
    Dim vTemp As Variant
    
    strFolder = tbPath.Value & tbFolder.Value
    strDWL = Replace(tbDrawingName.Value, ".dwg", ".dwl")
    
    strFile = Dir$(strFolder, vbHidden)
    
    Do While strFile <> ""
        If strFile = strDWL Then GoTo Found_File
        
        strFile = Dir$
    Loop
    
    tbOComputer.Value = "<none>"
    Exit Sub
Found_File:
    
    strDWL = strFolder & strFile
    
    Open strDWL For Input As #1
    
    Line Input #1, strLine
    vTemp = Split(strLine, vbLf)
    vTemp(0) = Replace(vTemp(0), vbCr, "")
    
    tbOComputer.Value = vTemp(0)
    Close #1
    
    Select Case LCase(vTemp(0))
        Case "integrity"
            tbOName.Value = "Dylan Spears"
        Case "integrity.tab"
            tbOName.Value = "Jeremy Pafford"
        Case "integrity.1"
            tbOName.Value = "Ronn Elliott"
        Case "integrity.2"
            tbOName.Value = "Rich Taylor"
        Case "integrity.3"
            tbOName.Value = "Jason Pafford"
        Case "integrity.4"
            tbOName.Value = "Byron Auer"
        Case "integrity.5"
            tbOName.Value = "Adam Kemper"
        Case "integrity.6"
            tbOName.Value = "Jon Wilburn"
        Case "integrity.7"
            tbOName.Value = "Tara Taylor"
        Case "integrity.8"
            tbOName.Value = "Franklin Angulo"
        Case "integrity9"
            tbOName.Value = "Wade Hampton"
        Case "integrity10"
            tbOName.Value = "Sam Jackson"
        Case "integrity11"
            tbOName.Value = "A Ghost?"
        Case "integrity12"
            tbOName.Value = "Daniel Campbell"
        Case "integrity13"
            tbOName.Value = "Nick Lockyear"
        Case "integrity14"
            tbOName.Value = "A Ghost?"
        Case "integrity15"
            tbOName.Value = "Jay Penny"
        Case "integrity16"
            tbOName.Value = "Drew Curtis"
        Case Else
            tbOName.Value = "Unknown"
    End Select
End Sub
