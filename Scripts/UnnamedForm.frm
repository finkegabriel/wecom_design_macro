VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnnamedForm 
   Caption         =   "Place LL Note"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   OleObjectBlob   =   "UnnamedForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnnamedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGetPoint_Click()
    Dim vCoords As Variant
    
    On Error Resume Next
    Me.Hide
    
    vCoords = ThisDrawing.Utility.GetPoint(, vbCr & "Pick Point:")
    If Not Err = 0 Then GoTo Exit_Sub
    
    tbTN83F.Value = vCoords(1) & "," & vCoords(0)
    
    Call GetLL
Exit_Sub:
    Me.show
End Sub

Private Sub cbGoogleMaps_Click()
    If tbLL.Value = "" Then Exit Sub
    If tbLL.Value = "error" Then Exit Sub
    If InStr(tbLL.Value, ",") < 1 Then Exit Sub
    
    Dim strURL As String
    
    strURL = "http://www.google.com/maps/place/" & tbLL.Value
    Shell ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -url " & strURL)
End Sub

Private Sub cbPlaceNote_Click()
    Dim objMText As AcadMText
    Dim vBasePnt, vLL As Variant
    Dim strLine As String
    Dim dInsert(2) As Double
    Dim dScale As Double
    
    On Error Resume Next
    Me.Hide
    
    vBasePnt = ThisDrawing.Utility.GetPoint(, "Select Note Placement: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    dInsert(0) = vBasePnt(0)
    dInsert(1) = vBasePnt(1)
    dInsert(2) = 0#
    
    dScale = CDbl(tbTextHeight.Value)
    
    strLine = Replace(tbNote.Value, vbLf, "")
    strLine = Replace(strLine, vbCr, "\P")
    
    If InStr(strLine, "{Lat}") > 0 Then
        vLL = Split(tbLL.Value, ",")
        
        strLine = Replace(strLine, "{Lat}", vLL(0))
        strLine = Replace(strLine, "{Long}", vLL(1))
    End If
    
    If InStr(strLine, "{N}") > 0 Then
        vLL = Split(tbTN83F.Value, ",")
        
        strLine = Replace(strLine, "{N}", vLL(0))
        strLine = Replace(strLine, "{E}", vLL(1))
    End If
    
    If InStr(strLine, "{D}") > 0 Then
        strLine = Replace(strLine, "{D}", Date)
    End If

    Set objMText = ThisDrawing.ModelSpace.AddMText(dInsert, 0, strLine)
    objMText.Layer = cbNoteLayer.Value
    objMText.Height = dScale
    objMText.InsertionPoint = dInsert
    Select Case cbJust.Value
        Case "TL"
            objMText.AttachmentPoint = acAttachmentPointTopLeft
        Case "TC"
            objMText.AttachmentPoint = acAttachmentPointTopCenter
        Case "TR"
            objMText.AttachmentPoint = acAttachmentPointTopRight
        Case "ML"
            objMText.AttachmentPoint = acAttachmentPointMiddleLeft
        Case "MC"
            objMText.AttachmentPoint = acAttachmentPointMiddleCenter
        Case "MR"
            objMText.AttachmentPoint = acAttachmentPointMiddleRight
        Case "BL"
            objMText.AttachmentPoint = acAttachmentPointBottomLeft
        Case "BC"
            objMText.AttachmentPoint = acAttachmentPointBottomCenter
        Case "BR"
            objMText.AttachmentPoint = acAttachmentPointBottomRight
    End Select
    If cbMask.Value = True Then objMText.BackgroundFill = True
    
    objMText.Update
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbZoomCoordinates_Click()
    If tbTN83F.Value = "" Then Exit Sub
    If InStr(tbTN83F.Value, ",") < 1 Then Exit Sub
    
    Dim vCoords As Variant
    'Dim dLL(2), dUR(2) As Double
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    Dim strLine As String
    
    vCoords = Split(tbTN83F.Value, ",")
    
    Me.Hide
    
    viewCoordsB(0) = CDbl(vCoords(1)) - 300
    viewCoordsB(1) = CDbl(vCoords(0)) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = viewCoordsB(0) + 600
    viewCoordsE(1) = viewCoordsB(1) + 600
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Me.show
End Sub

Private Sub tbTN83F_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If InStr(tbTN83F.Value, ",") < 1 Then Exit Sub
    
    Dim vCoords As Variant
    Dim dN, dE As Double
    
    On Error Resume Next
    
    vCoords = Split(tbTN83F.Value, ",")
    dN = CDbl(vCoords(0))
    dE = CDbl(vCoords(1))
    
    If Not Err = 0 Then
        tbLL.Value = "error"
        Exit Sub
    End If
    
    Call GetLL
End Sub

Private Sub UserForm_Initialize()
    cbJust.AddItem "TL"
    cbJust.AddItem "TC"
    cbJust.AddItem "TR"
    cbJust.AddItem "ML"
    cbJust.AddItem "MC"
    cbJust.AddItem "MR"
    cbJust.AddItem "BL"
    cbJust.AddItem "BC"
    cbJust.AddItem "BR"
    cbJust.Value = "TL"
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        cbNoteLayer.AddItem objLayer.Name
    Next objLayer
    cbNoteLayer.Value = "Integrity Planning"
    
    Dim strNote As String
    
    strNote = "{Lat}itude  {Long}itude  {N}orth  {E}ast  {D}ate"
    tbNote.ControlTipText = strNote
    
    tbTN83F.SetFocus
End Sub

Private Sub GetLL()
    Dim vCoords As Variant
    Dim vLL As Variant
    Dim dN, dE As Double
    
    vCoords = Split(tbTN83F.Value, ",")
    dN = CDbl(vCoords(0))
    dE = CDbl(vCoords(1))
    
    vLL = TN83FtoLL(CDbl(dN), CDbl(dE))
    tbLL.Value = vLL(0) & "," & vLL(1)
End Sub

Private Function TN83FtoLL(dNorth As Double, dEast As Double)
    Dim dLat, dLong, dDLat As Double
    Dim dDiffE, dEast0 As Double
    Dim dDiffN, dNorth0 As Double
    Dim dR, dCAR, dCA, dU, dK As Double
    Dim LL(2) As Double
    
    dEast = dEast * 0.3048006096
    dNorth = dNorth * 0.3048006096
    
    dDiffE = dEast - 600000
    dDiffN = dNorth - 166504.1691
    
    dR = 8842127.1422 - dDiffN
    dCAR = Atn(dDiffE / dR)
    dCA = dCAR * 180 / 3.14159265359
    dLong = -86 + dCA / 0.585439726459
    
    dU = dDiffN - dDiffE * Tan(dCAR / 2)
    dDLat = dU * (0.00000901305249 + dU * (-6.77268E-15 + dU * (-3.72351E-20 + dU * -9.2828E-28)))
    dLat = 35.8340607459 + dDLat
    
    dK = 0.999948401424 + (1.23188E-14 * dU * dU) + (4.54E-22 * dU * dU * dU)
    
    LL(0) = dLat
    LL(1) = dLong
    LL(2) = dK
    
    TN83FtoLL = LL
End Function

Private Function LLtoTN83F(dLat As Double, dLong As Double)
    Dim dDLat As Double
    Dim dEast, dDiffE, dEast0 As Double
    Dim dNorth, dDiffN, dNorth0 As Double
    Dim dU, dR, dCA, dK As Double
    Dim NE(2) As Double
    
    dDLat = dLat - 35.8340607459
    dU = dDLat * (110950.2019 + dDLat * (9.25072 + dDLat * (5.64572 + dDLat * 0.017374)))
    dR = 8842127.1422 - dU
    dCA = ((86 + dLong) * 0.585439726459) * 3.14159265359 / 180
    
    dDiffE = dR * Sin(dCA)
    dDiffN = dU + dDiffE * Tan(dCA / 2)
    
    dEast = (dDiffE + 600000) / 0.3048006096
    dNorth = (dDiffN + 166504.1691) / 0.3048006096
    
    dK = 0.999948401424 + (1.23188E-14 * dU * dU) + (4.54E-22 * dU * dU * dU)
    
    NE(0) = dNorth
    NE(1) = dEast
    NE(2) = dK
    
    LLtoTN83F = NE
End Function
