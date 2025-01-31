VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateMCL 
   Caption         =   "Modify MCL File"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150.001
   OleObjectBlob   =   "CreateMCL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateMCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbClearPanel_Click()
    lbCounts.Clear
End Sub

Private Sub cbClearSelected_Click()
    If lbCounts.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbCounts.ListCount - 1
        If lbCounts.Selected(i) = True Then
            lbCounts.List(i, 1) = "<>"
            lbCounts.List(i, 2) = "<>"
            lbCounts.List(i, 3) = "<>"
            lbCounts.List(i, 4) = "<>"
            lbCounts.List(i, 5) = "<>"
        End If
    Next i
End Sub

Private Sub cbCreatePanel_Click()
    If tbF1Name.Value = "" Then Exit Sub
    
    'Label42.Enabled = False
    'Label43.Enabled = False
    'tbStartFiber.Enabled = False
    'tbEndFiber.Enabled = False
    'cbCreatePanel.Enabled = False
    'lbCounts.Clear
    
    Dim iStart, iEnd As Integer
    
    iStart = CInt(tbStartFiber.Value)
    iEnd = CInt(tbEndFiber.Value)
    
    For i = iStart To iEnd
        lbCounts.AddItem i
        lbCounts.List(lbCounts.ListCount - 1, 1) = "<>"
        lbCounts.List(lbCounts.ListCount - 1, 2) = "<>"
        lbCounts.List(lbCounts.ListCount - 1, 3) = "<>"
        lbCounts.List(lbCounts.ListCount - 1, 4) = "<>"
        lbCounts.List(lbCounts.ListCount - 1, 5) = "<>"
    Next i
End Sub

Private Sub cbMarkSelected_Click()
    If lbCounts.ListCount < 1 Then Exit Sub
    If tbMark.Value = "" Then Exit Sub
    
    For i = 0 To lbCounts.ListCount - 1
        If lbCounts.Selected(i) = True Then
            lbCounts.List(i, 1) = tbMark.Value
            lbCounts.List(i, 2) = tbMark.Value
            lbCounts.List(i, 3) = tbMark.Value
            lbCounts.List(i, 4) = tbMark.Value
            lbCounts.List(i, 5) = tbMark.Value
        End If
    Next i
End Sub

Private Sub cbMoveSelected_Click()
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim iCount As Integer
    Dim strList(5) As String
    
    iCount = lbCounts.ListCount - 1
    For i = iCount To 0 Step -1
        If lbCounts.Selected(i) = True Then
            strList(0) = lbCounts.List(i, 0)
            strList(1) = lbCounts.List(i, 1)
            strList(2) = lbCounts.List(i, 2)
            strList(3) = lbCounts.List(i, 3)
            strList(4) = lbCounts.List(i, 4)
            strList(5) = lbCounts.List(i, 5)
            
            lbCounts.Selected(i) = False
            lbCounts.RemoveItem i
            lbCounts.AddItem strList(0), 0
            lbCounts.List(0, 1) = strList(1)
            lbCounts.List(0, 2) = strList(2)
            lbCounts.List(0, 3) = strList(3)
            lbCounts.List(0, 4) = strList(4)
            lbCounts.List(0, 5) = strList(5)
            
            i = i + 1
        End If
    Next i
End Sub

Private Sub cbOpenMCL_Click()
    If tbF1Name.Value = "" Then Exit Sub
    lbCounts.Clear
    
    Dim strFileName As String
    Dim vName, vLine, vItem As Variant
    Dim vFile As Variant
    Dim fName As String
    
    vName = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vName(0) & " Counts -" & tbF1Name.Value & ".mcl"
    
    fName = Dir(strFileName)
    If fName = "" Then
        'Label42.Enabled = True
        'Label43.Enabled = True
        'tbStartFiber.Enabled = True
        'tbEndFiber.Enabled = True
        'cbCreatePanel.Enabled = True
        
        tbStartFiber.SetFocus
        Exit Sub
    End If
    
    Open strFileName For Input As #1
    vFile = Split(Input$(LOF(1), 1), vbCrLf)
    Close #1
    
    For i = 1 To UBound(vFile)
        If vFile(i) = "" Then GoTo Next_I
        vLine = Split(vFile(i), vbTab)
        
        lbCounts.AddItem vLine(0)
        lbCounts.List(lbCounts.ListCount - 1, 1) = vLine(1)
        lbCounts.List(lbCounts.ListCount - 1, 2) = vLine(2)
        lbCounts.List(lbCounts.ListCount - 1, 3) = vLine(3)
        lbCounts.List(lbCounts.ListCount - 1, 4) = vLine(4)
        lbCounts.List(lbCounts.ListCount - 1, 5) = vLine(5)
Next_I:
    Next i
End Sub

Private Sub cbRemoveSelected_Click()
    If lbCounts.ListCount < 1 Then Exit Sub
    
    For i = lbCounts.ListCount - 1 To 0 Step -1
        If lbCounts.Selected(i) = True Then lbCounts.RemoveItem i
    Next i
End Sub

Private Sub cbSavePanel_Click()
    If tbF1Name.Value = "" Then Exit Sub
    
    Dim strFileName As String
    Dim vName As Variant
    Dim fName, strLine As String
    
    vName = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vName(0) & " Counts -" & tbF1Name.Value & ".mcl"
    
    Open strFileName For Output As #1
    
    Print #1, tbF1Name.Value & " FIBER COUNTS"
    For i = 0 To lbCounts.ListCount - 1
        strLine = lbCounts.List(i, 0) & vbTab & lbCounts.List(i, 1) & vbTab & lbCounts.List(i, 2) & vbTab & lbCounts.List(i, 3)
        strLine = strLine & vbTab & lbCounts.List(i, 4) & vbTab & lbCounts.List(i, 5)
        Print #1, strLine
    Next i
    
    Close #1
End Sub

Private Sub tbF1Name_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    tbF1Name.Value = UCase(tbF1Name.Value)
End Sub

Private Sub UserForm_Initialize()
    lbCounts.Clear
    lbCounts.ColumnCount = 6
    lbCounts.ColumnWidths = "30;72;36;144;48;108"
    
End Sub
