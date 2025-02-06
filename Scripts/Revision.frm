VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Revision 
   Caption         =   "Revision"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "Revision.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Revision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSS As AcadSelectionSet
    
Private Sub cbAddNote_Click()
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    
    On Error Resume Next
    
    Me.Hide
    
Place_Another:
    
    vReturnPnt = ThisDrawing.Utility.GetPoint(, vbCr & "Place note:")
    If Not Err = 0 Then GoTo Exit_Sub
    
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(vReturnPnt, "Revision", 1#, 1#, 1#, 0#)
    objBlock.Layer = "Integrity Notes"
    
    vAttList = objBlock.GetAttributes
    vAttList(0).TextString = cbNumber.Value
    vAttList(1).TextString = tbDate.Value
    vAttList(2).TextString = UCase(tbInitials.Value)
    vAttList(3).TextString = UCase(tbNote.Value)
    objBlock.Update
    
    GoTo Place_Another
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbNumber_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If cbNumber.ListCount < 1 Then Exit Sub
    
    For i = 0 To cbNumber.ListCount - 1
        If cbNumber.Value = cbNumber.List(i) Then GoTo Found_Revision
    Next i
    
    cbUpdateNote.Enabled = False
    Exit Sub
    
Found_Revision:
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = cbNumber.Value Then
            tbDate.Value = vAttList(1).TextString
            tbInitials.Value = vAttList(2).TextString
            tbNote.Value = vAttList(3).TextString
    
            cbUpdateNote.Enabled = True
            cbAddNote.Enabled = False
            cbNumber.Enabled = False
            Exit Sub
        End If
    Next objBlock
    
End Sub

Private Sub cbUpdateNote_Click()
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = cbNumber.Value Then
            vAttList(1).TextString = tbDate.Value
            vAttList(2).TextString = UCase(tbInitials.Value)
            vAttList(3).TextString = UCase(tbNote.Value)
    
            objBlock.Update
        End If
    Next objBlock
    
    cbUpdateNote.Enabled = False
    cbAddNote.Enabled = True
    cbNumber.Enabled = True
    
    cbNumber.Value = ""
    tbDate.Value = ""
    tbInitials.Value = ""
    tbNote.Value = ""
End Sub

Private Sub Label2_Click()
    tbDate.Value = Date
End Sub

Private Sub tbNote_Change()
    tbNote.Value = UCase(tbNote.Value)
End Sub

Private Sub UserForm_Deactivate()
    objSS.Clear
    objSS.Delete
End Sub

Private Sub UserForm_Initialize()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strNumber As String
    Dim iNumber, iNext As Integer
    
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    grpCode(0) = 2
    grpValue(0) = "Revision"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    iNumber = 0
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        strNumber = vAttList(0).TextString
        iNext = CInt(strNumber)
        If iNext > iNumber Then iNumber = iNext
        
        If cbNumber.ListCount < 1 Then
            cbNumber.AddItem strNumber
        Else
            For i = 0 To cbNumber.ListCount - 1
                If cbNumber.List(i) = strNumber Then GoTo Next_objBlock
            Next i
            
            cbNumber.AddItem vAttList(0).TextString
        End If
Next_objBlock:
    Next objBlock
    
    cbNumber.Value = iNumber + 1
End Sub
