VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zzzzPoleCOMMupdate 
   Caption         =   "Pole Attachment Aliases"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "zzzzPoleCOMMupdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zzzzPoleCOMMupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objBlock As AcadBlockReference

Private Sub cbAdd_Click()
    If cbAdd.Caption = "Add" Then
        lbList.AddItem UCase(tbAlias.Value)
        lbList.List(lbList.ListCount - 1, 1) = UCase(tbCompany.Value)
    Else
        lbList.List(lbList.ListIndex, 0) = UCase(tbAlias.Value)
        lbList.List(lbList.ListIndex, 1) = UCase(tbCompany.Value)
    
        cbAdd.Caption = "Add"
    
        cbUpdateAlias.Enabled = True
        cbConvert.Enabled = True
        cbUpdateBlock.Enabled = True
    End If
End Sub

Private Sub cbConvert_Click()
    MsgBox "This is not completed." & vbCr & "May not be on here anyway."
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbUpdateAlias_Click()
    If lbList.ListCount > 8 Then
        MsgBox "Too many Aliases(" & lbList.ListCount & ") For the Pole Attributes." & vbCr & "First 8 will be used."
    End If
    
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objPole As AcadBlockReference
    Dim vAttList, vTemp As Variant
    Dim i, iMax As Integer
    
    On Error Resume Next
    
    iMax = lbList.ListCount - 1
    If iMax > 7 Then iMax = 7
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    If objSS.count < 1 Then GoTo Exit_Sub
    'If objSS.count < 1 Then Me.Hide
    
    For Each objPole In objSS
        vAttList = objPole.GetAttributes
        
        For i = 0 To iMax
            'If lbList.ListCount - 1 > i Then GoTo Next_objPole
            If vAttList(16 + i).TextString = "" Then vAttList(16 + i).TextString = lbList.List(i, 0) & "="
        Next i
'Next_objPole:
    Next objPole
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub

Private Sub cbUpdateBlock_Click()
    Dim vAttList As Variant
    Dim i, iMax, iAtt As Integer
    
    vAttList = objBlock.GetAttributes
    
    For i = 0 To 19
        vAttList(i).TextString = "<blank>"
        i = i + 1
        vAttList(i).TextString = "<blank>"
    Next i
    
    iMax = lbList.ListCount - 1
    If iMax > 7 Then iMax = 7
    
    iAtt = 0
    
    For i = 0 To iMax
        vAttList(iAtt).TextString = lbList.List(i, 0)
        iAtt = iAtt + 1
        vAttList(iAtt).TextString = lbList.List(i, 1)
        iAtt = iAtt + 1
    Next i
    
    objBlock.Update
End Sub

Private Sub lbList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbAlias.Value = lbList.List(lbList.ListIndex, 0)
    tbCompany.Value = lbList.List(lbList.ListIndex, 1)
    
    cbAdd.Caption = "Update Line"
    
    cbUpdateAlias.Enabled = False
    cbConvert.Enabled = False
    cbUpdateBlock.Enabled = False
    
    tbAlias.SetFocus
End Sub

Private Sub lbList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbList.ListCount < 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            lbList.RemoveItem lbList.ListIndex
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbList.ColumnCount = 2
    lbList.ColumnWidths = "72;66"
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim vAttList As Variant
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "Attachment alias"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    If objSS.count < 1 Then GoTo Exit_Sub
    
    Set objBlock = objSS.Item(0)
    vAttList = objBlock.GetAttributes
    
    For i = 0 To 19
        If vAttList(i).TextString = vAttList(i + 1).TextString Then
            i = i + 1
        Else
            lbList.AddItem vAttList(i).TextString
            i = i + 1
            lbList.List(lbList.ListCount - 1, 1) = vAttList(i).TextString
        End If
    Next i
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub
