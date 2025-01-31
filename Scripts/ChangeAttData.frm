VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeAttData 
   Caption         =   "Change Attribute Properties"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "ChangeAttData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeAttData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRunAll_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    Dim strLine As String
    Dim dRotation As Double
    'Dim vLine As Variant
    
    If Not tbRotation.Value = "" Then dRotation = CDbl(tbRotation.Value) * 3.14159265359 / 180
    Me.Hide
    
    On Error Resume Next
Get_Another:
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Callout:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    If objBlock.HasAttributes = False Then GoTo Get_Another
    
    vAttList = objBlock.GetAttributes
    
    For i = 0 To UBound(vAttList)
        If Not tbRotation.Value = "" Then vAttList(i).Rotation = dRotation
        If Not tbHeight.Value = "" Then vAttList(i).Height = CDbl(tbHeight.Value)
        Select Case cbJustification.Value
            Case "Aligned"
                vAttList(i).Alignment = acAlignmentAligned
            Case "BottomCenter"
                vAttList(i).Alignment = acAlignmentBottomCenter
            Case "BottomLeft"
                vAttList(i).Alignment = acAlignmentBottomLeft
            Case "BottomRight"
                vAttList(i).Alignment = acAlignmentBottomRight
            Case "Center"
                vAttList(i).Alignment = acAlignmentCenter
            Case "Fit"
                vAttList(i).Alignment = acAlignmentFit
            Case "Left"
                vAttList(i).Alignment = acAlignmentLeft
            Case "Middle"
                vAttList(i).Alignment = acAlignmentMiddle
            Case "MiddleCenter"
                vAttList(i).Alignment = acAlignmentMiddleCenter
            Case "MiddleLeft"
                vAttList(i).Alignment = acAlignmentMiddleLeft
            Case "MiddleRight"
                vAttList(i).Alignment = acAlignmentMiddleRight
            Case "Right"
                vAttList(i).Alignment = acAlignmentRight
            Case "TopCenter"
                vAttList(i).Alignment = acAlignmentTopCenter
            Case "TopLeft"
                vAttList(i).Alignment = acAlignmentTopLeft
            Case "TopRight"
                vAttList(i).Alignment = acAlignmentTopRight
        End Select
        vAttList(i).Backward = cbBackwards.Value
        vAttList(i).UpsideDown = cbUpsideDown.Value
    Next i
    
    If Not tbFirstAtt.Value = "" Then vAttList(0).TextString = tbFirstAtt.Value
    objBlock.Update
    
    GoTo Get_Another
Exit_Sub:
    Me.show
End Sub

Private Sub cbRunCustomer_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    Dim strLine As String
    Dim dRotation As Double
    'Dim vLine As Variant
    
    If Not tbRotation.Value = "" Then dRotation = CDbl(tbRotation.Value) * 3.14159265359 / 180
    Me.Hide
    
    On Error Resume Next
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Callout:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    If objBlock.HasAttributes = False Then GoTo Get_Another
    
    vAttList = objBlock.GetAttributes
    
    If Not tbRotation.Value = "" Then vAttList(1).Rotation = dRotation
    If Not tbHeight.Value = "" Then vAttList(1).Height = CDbl(tbHeight.Value)
    Select Case cbJustification.Value
        Case "Aligned"
            vAttList(1).Alignment = acAlignmentAligned
        Case "BottomCenter"
            vAttList(1).Alignment = acAlignmentBottomCenter
        Case "BottomLeft"
            vAttList(1).Alignment = acAlignmentBottomLeft
        Case "BottomRight"
            vAttList(1).Alignment = acAlignmentBottomRight
        Case "Center"
            vAttList(1).Alignment = acAlignmentCenter
        Case "Fit"
            vAttList(1).Alignment = acAlignmentFit
        Case "Left"
            vAttList(1).Alignment = acAlignmentLeft
        Case "Middle"
            vAttList(1).Alignment = acAlignmentMiddle
        Case "MiddleCenter"
            vAttList(1).Alignment = acAlignmentMiddleCenter
        Case "MiddleLeft"
            vAttList(1).Alignment = acAlignmentMiddleLeft
        Case "MiddleRight"
            vAttList(1).Alignment = acAlignmentMiddleRight
        Case "Right"
            vAttList(1).Alignment = acAlignmentRight
        Case "TopCenter"
            vAttList(1).Alignment = acAlignmentTopCenter
        Case "TopLeft"
            vAttList(1).Alignment = acAlignmentTopLeft
        Case "TopRight"
            vAttList(1).Alignment = acAlignmentTopRight
    End Select
    vAttList(1).Backward = cbBackwards.Value
    vAttList(1).UpsideDown = cbUpsideDown.Value
    If Not tbFirstAtt.Value = "" Then vAttList(1).TextString = tbFirstAtt.Value
    objBlock.Update
    
    GoTo Get_Another
Exit_Sub:
    Me.show
End Sub

Private Sub cbRunFirst_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    Dim strLine As String
    Dim dRotation As Double
    'Dim vLine As Variant
    
    If Not tbRotation.Value = "" Then dRotation = CDbl(tbRotation.Value) * 3.14159265359 / 180
    Me.Hide
    
    On Error Resume Next
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Callout:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    If objBlock.HasAttributes = False Then GoTo Get_Another
    
    vAttList = objBlock.GetAttributes
    
    If Not tbRotation.Value = "" Then vAttList(0).Rotation = dRotation
    If Not tbHeight.Value = "" Then vAttList(0).Height = CDbl(tbHeight.Value)
    Select Case cbJustification.Value
        Case "Aligned"
            vAttList(0).Alignment = acAlignmentAligned
        Case "BottomCenter"
            vAttList(0).Alignment = acAlignmentBottomCenter
        Case "BottomLeft"
            vAttList(0).Alignment = acAlignmentBottomLeft
        Case "BottomRight"
            vAttList(0).Alignment = acAlignmentBottomRight
        Case "Center"
            vAttList(0).Alignment = acAlignmentCenter
        Case "Fit"
            vAttList(0).Alignment = acAlignmentFit
        Case "Left"
            vAttList(0).Alignment = acAlignmentLeft
        Case "Middle"
            vAttList(0).Alignment = acAlignmentMiddle
        Case "MiddleCenter"
            vAttList(0).Alignment = acAlignmentMiddleCenter
        Case "MiddleLeft"
            vAttList(0).Alignment = acAlignmentMiddleLeft
        Case "MiddleRight"
            vAttList(0).Alignment = acAlignmentMiddleRight
        Case "Right"
            vAttList(0).Alignment = acAlignmentRight
        Case "TopCenter"
            vAttList(0).Alignment = acAlignmentTopCenter
        Case "TopLeft"
            vAttList(0).Alignment = acAlignmentTopLeft
        Case "TopRight"
            vAttList(0).Alignment = acAlignmentTopRight
    End Select
    vAttList(0).Backward = cbBackwards.Value
    vAttList(0).UpsideDown = cbUpsideDown.Value
    If Not tbFirstAtt.Value = "" Then vAttList(0).TextString = tbFirstAtt.Value
    objBlock.Update
    
    GoTo Get_Another
Exit_Sub:
    Me.show
End Sub

Private Sub UserForm_Initialize()
    cbJustification.AddItem "Aligned"
    cbJustification.AddItem "BottomCenter"
    cbJustification.AddItem "BottomLeft"
    cbJustification.AddItem "BottomRight"
    cbJustification.AddItem "Center"
    cbJustification.AddItem "Fit"
    cbJustification.AddItem "Left"
    cbJustification.AddItem "Middle"
    cbJustification.AddItem "MiddleCenter"
    cbJustification.AddItem "MiddleLeft"
    cbJustification.AddItem "MiddleRight"
    cbJustification.AddItem "Right"
    cbJustification.AddItem "TopCenter"
    cbJustification.AddItem "TopLeft"
    cbJustification.AddItem "TopRight"
    
    'cbJustification.Value = "MiddleCenter"
End Sub
