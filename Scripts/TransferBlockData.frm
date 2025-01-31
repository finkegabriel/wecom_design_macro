VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferBlockData 
   Caption         =   "Transfer Block Data"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000.001
   OleObjectBlob   =   "TransferBlockData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransferBlockData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objFrom As AcadBlockReference
Dim objTo As AcadBlockReference

Private Sub cbGetFrom_Click()
    Dim objEntity As AcadEntity
    Dim vAttList, vReturnPnt As Variant
    
    Me.Hide
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block to transfer FROM:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objFrom = objEntity
    vAttList = objFrom.GetAttributes
    If Not Err = 0 Then GoTo Exit_Sub
    
    For i = 0 To UBound(vAttList)
        lbFrom.AddItem vAttList(i).TagString
        lbFrom.List(i, 1) = vAttList(i).TextString
    Next i
    
    LabelFrom.Caption = LabelFrom.Caption & ": " & objFrom.Name
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block to transfer TO:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objTo = objEntity
    vAttList = objTo.GetAttributes
    If Not Err = 0 Then GoTo Exit_Sub
    
    For i = 0 To UBound(vAttList)
        lbTo.AddItem vAttList(i).TagString
        lbTo.List(i, 1) = vAttList(i).TextString
    Next i
    
    LabelTo.Caption = LabelTo.Caption & ": " & objTo.Name
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbUpdate_Click()
    If lbFrom.ListCount < 1 Then Exit Sub
    If lbTo.ListCount < 1 Then Exit Sub
    
    Dim strLine As String
    
    strLine = ""
    For i = 0 To lbTo.ListCount - 1
        If lbTo.Selected(i) = True Then
            If InStr(lbTo.List(i, 1), "{{") > 0 Then
                If strLine = "" Then
                    strLine = lbTo.List(i, 1) & ">" & i
                Else
                    strLine = strLine & ":" & lbTo.List(i, 1) & ">" & i
                End If
            End If
        End If
    Next i
    
    If strLine = "" Then Exit Sub
    strLine = Replace(strLine, "{{", "")
    strLine = Replace(strLine, "}}", "")
    
    Dim objEntity As AcadEntity
    Dim vAttFrom, vAttTo As Variant
    Dim vReturnPnt, vLine, vItem As Variant
    
    Me.Hide
    On Error Resume Next
    
Next_Set:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block to transfer FROM:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objFrom = objEntity
    vAttFrom = objFrom.GetAttributes
    If Not Err = 0 Then GoTo Exit_Sub
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block to transfer TO:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objTo = objEntity
    vAttTo = objTo.GetAttributes
    If Not Err = 0 Then GoTo Exit_Sub
    
    vLine = Split(strLine, ":")
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ">")
        vAttTo(CInt(vItem(1))).TextString = vAttFrom(CInt(vItem(0))).TextString
    Next i
    
    GoTo Next_Set
    
Exit_Sub:
    Me.show
End Sub

Private Sub lbTo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbFrom.ListCount < 1 Then Exit Sub
    If lbTo.ListCount < 1 Then Exit Sub
    
    'Dim strLine As String
    Dim iIndex As Integer
    
    'strLine = ""
    For i = 0 To lbTo.ListCount - 1
        If lbTo.Selected(i) = True Then
            For j = 0 To lbFrom.ListCount - 1
                If lbFrom.Selected(j) = True Then
                    lbTo.List(i, 1) = "{{" & j & "}}"
                    
                    If j < lbFrom.ListCount - 1 Then
                        lbFrom.Selected(j) = False
                        lbFrom.Selected(j + 1) = True
                    End If
                    
                    Exit Sub
                End If
            Next j
        End If
    Next i
    
End Sub

Private Sub UserForm_Initialize()
    lbFrom.ColumnCount = 2
    lbFrom.ColumnWidths = "72;138"
    
    lbTo.ColumnCount = 2
    lbTo.ColumnWidths = "72;138"
End Sub
