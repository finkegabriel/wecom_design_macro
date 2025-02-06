VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertCustomers 
   Caption         =   "Convert Customers"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "ConvertCustomers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConvertCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAdd_Click()
    If tbText.Value = "" Then Exit Sub
    If cbType.Value = "" Then Exit Sub
    
    lbList.AddItem UCase(tbText.Value)
    lbList.List(lbList.ListCount - 1, 1) = cbType.Value
    
    tbText.Value = ""
    cbType.Value = ""
    
    tbText.SetFocus
End Sub

Private Sub cbConvert_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iCount As Integer
    
    On Error Resume Next
    
    iCount = 0
    
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    MsgBox "Found:  " & objSS.count
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        For i = 0 To lbList.ListCount - 1
            If InStr(UCase(vAttList(3).TextString), lbList.List(i, 0)) > 0 Then
                vAttList(0).TextString = lbList.List(i, 1)
                
                Select Case Left(lbList.List(i, 1), 1)
                    Case "R"
                        vAttList(5).TextString = ""
                    Case "E"
                        vAttList(5).TextString = "X"
                    Case Else
                        vAttList(5).TextString = Left(lbList.List(i, 1), 1)
                End Select
                
                objBlock.Update
                
                iCount = iCount + 1
                
                GoTo Next_objBlock
            End If
        Next i
Next_objBlock:
    Next objBlock
    
    MsgBox "Converted  " & iCount & "  customers"
    lbList.Clear
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub lbList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbList.ListCount < 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            lbList.RemoveItem lbList.ListIndex
    End Select
End Sub

Private Sub lbNotes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbNotes.ListIndex < 0 Then Exit Sub
    
    tbText.Value = lbNotes.List(lbNotes.ListIndex, 1)
End Sub

Private Sub lbNotes_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbNotes.ListIndex < 0 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            lbNotes.RemoveItem lbNotes.ListIndex
            
            tbListcount.Value = lbNotes.ListCount
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbNotes.ColumnCount = 2
    lbNotes.ColumnWidths = "36;100"
    
    lbList.ColumnCount = 2
    lbList.ColumnWidths = "72;66"
    
    cbType.AddItem "BUSINESS"
    cbType.AddItem "CHURCH"
    cbType.AddItem "MDU"
    cbType.AddItem "RESIDENCE"
    cbType.AddItem "SCHOOL"
    cbType.AddItem "TRAILER"
    cbType.AddItem "EXTENSION"
    
    'Dim vDwgLL, vDwgUR As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAtt As Variant
    Dim strLine As String
    
    'vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    'vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    tbBlocks.Value = objSS.count
    If objSS.count < 1 Then GoTo Exit_Sub
    
    For Each objBlock In objSS
        vAtt = objBlock.GetAttributes
        strLine = vAtt(3).TextString
        
        If lbNotes.ListCount > 0 Then
            For i = 0 To lbNotes.ListCount - 1
                If lbNotes.List(i, 1) = strLine Then GoTo Found_Note
            Next i
        End If
        
        Select Case vAtt(5).TextString
            Case ""
                lbNotes.AddItem "R"
            Case Else
                lbNotes.AddItem vAtt(5).TextString
        End Select
        
        lbNotes.List(lbNotes.ListCount - 1, 1) = strLine
        
Found_Note:
    Next objBlock
    
    tbListcount.Value = lbNotes.ListCount
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub
