VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlaceSPIDAPoles 
   Caption         =   "Place Poles from SPIDAmin"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10080
   OleObjectBlob   =   "PlaceSPIDAPoles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlaceSPIDAPoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGetList_Click()
    If tbPasted.Value = "" Then Exit Sub
    If cbLat.Value = "" Then Exit Sub
    If cbLong.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vTemp As Variant
    Dim strLine As String
    Dim iIndex As Integer
    Dim iH, iC, iOwner, iYear, iLat, iLong, iFID As Integer
    
    iH = -1: iC = -1: iOwner = -1: iYear = -1
    iLat = -1: iLong = -1: iFID = -1
    
    If Not cbH.Value = "" Then
        vTemp = Split(cbH.Value, " - ")
        iH = CInt(vTemp(0))
    End If
    
    If Not cbC.Value = "" Then
        vTemp = Split(cbC.Value, " - ")
        iC = CInt(vTemp(0))
    End If
    
    If Not cbOwner.Value = "" Then
        vTemp = Split(cbOwner.Value, " - ")
        iOwner = CInt(vTemp(0))
    End If
    
    If Not cbYear.Value = "" Then
        vTemp = Split(cbYear.Value, " - ")
        iYear = CInt(vTemp(0))
    End If
    
    vTemp = Split(cbLat.Value, " - ")
    iLat = CInt(vTemp(0))
    
    vTemp = Split(cbLong.Value, " - ")
    iLong = CInt(vTemp(0))
    
    If Not cbFID.Value = "" Then
        vTemp = Split(cbFID.Value, " - ")
        iFID = CInt(vTemp(0))
    End If
    
    lbPoles.Clear
    
    strLine = Replace(tbPasted.Value, vbLf, "")
    vLine = Split(strLine, vbCr)
    
    For i = 0 To UBound(vLine)
        If InStr(vLine(i), ",") < 1 Then GoTo Next_line
        
        vItem = Split(vLine(i), ",")
        'If UBound(vItem) < 1 Then GoTo Next_Line
        If Not vItem(1) = "POLE" Then GoTo Next_line
        
        For j = 0 To UBound(vItem)
            vItem(j) = Replace(vItem(j), ";;", ",")
        Next j
        
        lbPoles.AddItem "POLE"
        iIndex = lbPoles.ListCount - 1
        If Not iH = -1 Or iC = -1 Then lbPoles.List(iIndex, 1) = vItem(iH) & "-" & vItem(iC)
        If Not iOwner = -1 Then lbPoles.List(iIndex, 2) = vItem(iOwner)
        If Not iYear = -1 Then lbPoles.List(iIndex, 3) = vItem(iYear)
        lbPoles.List(iIndex, 4) = vItem(iLat) & "," & vItem(iLong)
        If Not iFID = -1 Then lbPoles.List(iIndex, 5) = vItem(iFID)
        
Next_line:
    Next i
    
    tbListCount.Value = lbPoles.ListCount
End Sub

Private Sub cbPlace_Click()
    If lbPoles.ListCount < 1 Then Exit Sub
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vCoords, vNE As Variant
    Dim strLayer As String
    Dim dCoords(2) As Double
    
    For i = 0 To lbPoles.ListCount - 1
        vCoords = Split(lbPoles.List(i, 4), ",")
        vNE = LLtoTN83F(CDbl(vCoords(0)), CDbl(vCoords(1)))
        dCoords(0) = CDbl(vNE(1))
        dCoords(1) = CDbl(vNE(0))
        dCoords(2) = 0#
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dCoords, "sPole", 1#, 1#, 1#, 0#)
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = lbPoles.List(i, 0)
        vAttList(1).TextString = "INCOMPLETE"
        If Not lbPoles.List(i, 2) = "" Then vAttList(2).TextString = lbPoles.List(i, 2)
        If Not lbPoles.List(i, 1) = "" Then vAttList(5).TextString = lbPoles.List(i, 1)
        If Not lbPoles.List(i, 3) = "" Then vAttList(6).TextString = lbPoles.List(i, 3)
        If Not lbPoles.List(i, 4) = "" Then vAttList(7).TextString = lbPoles.List(i, 4)
        If Not lbPoles.List(i, 5) = "" Then vAttList(24).TextString = "FID_" & lbPoles.List(i, 5)
        
        Select Case lbPoles.List(i, 2)
            Case "MTEMC", "NES"
                objBlock.Layer = "Integrity Poles-Power"
            Case Else
                objBlock.Layer = "Integrity Poles-Other"
        End Select
        
        objBlock.Update
    Next i
    
    MsgBox "Done."
End Sub

Private Sub tbPasted_Change()
    If tbPasted.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vTemp As Variant
    Dim strLine As String
    Dim iIndex As Integer
    
    lbPoles.Clear
    
    tbPasted.Value = Replace(tbPasted.Value, """", "")
    strLine = Replace(tbPasted.Value, vbLf, "")
    vItem = Split(strLine, vbCr)
    vLine = Split(vItem(0), ",")
    
    For i = 0 To UBound(vLine)
        strLine = i & " - " & vLine(i)
        
        cbH.AddItem strLine
        cbC.AddItem strLine
        cbOwner.AddItem strLine
        cbYear.AddItem strLine
        cbLat.AddItem strLine
        cbLong.AddItem strLine
        cbFID.AddItem strLine
    Next i
    
    Select Case cbCompany.Value
        Case "MTE"
            If cbH.ListCount > 16 Then
                cbH.Value = cbH.List(5)
                cbC.Value = cbC.List(4)
                cbOwner.Value = cbOwner.List(12)
                cbYear.Value = cbYear.List(16)
                cbLat.Value = cbLat.List(2)
                cbLong.Value = cbH.List(3)
                cbFID.Value = cbH.List(8)
            End If
        Case Else
            If cbH.ListCount > 16 Then
                cbH.Value = cbH.List(11)
                cbC.Value = cbC.List(10)
                cbOwner.Value = cbOwner.List(12)
                cbYear.Value = ""
                cbLat.Value = cbLat.List(2)
                cbLong.Value = cbH.List(3)
                cbFID.Value = cbH.List(0)
            End If
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbPoles.ColumnCount = 6
    lbPoles.ColumnWidths = "48;72;60;48;144;114"
    
    cbCompany.AddItem "MTE"
    cbCompany.AddItem "NES"
    cbCompany.Value = "NES"
End Sub

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
