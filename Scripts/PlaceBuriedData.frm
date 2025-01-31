VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlaceBuriedData 
   Caption         =   "Place Buried Plant Data"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10680
   OleObjectBlob   =   "PlaceBuriedData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlaceBuriedData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objBlock As AcadBlockReference
Dim strRoute As String
Dim iNumber As Integer
    
Private Sub cbAddUnit_Click()
    If cbAddUnit.Caption = "Update" Then
        lbUnits.List(lbUnits.ListIndex, 0) = tbUnit.Value
        lbUnits.List(lbUnits.ListIndex, 1) = tbQuantity.Value
        lbUnits.List(lbUnits.ListIndex, 2) = tbUNote.Value
        
        cbAddUnit.Caption = "Add Unit"
        GoTo Exit_Sub
    End If
    
    lbUnits.AddItem tbUnit.Value
    lbUnits.List(lbUnits.ListCount - 1, 1) = tbQuantity.Value
    lbUnits.List(lbUnits.ListCount - 1, 2) = tbUNote.Value
    
Exit_Sub:
    
    tbUnit.Value = ""
    tbQuantity.Value = ""
    tbUNote.Value = ""
End Sub

Private Sub cbGetBlock_Click()
    Dim objObject As AcadObject
    Dim vBasePnt, vAttList As Variant
    Dim vLine, vItem, vTemp As Variant
    Dim strTemp, strType, strGround As String
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objObject, vBasePnt, "Select Buried Plant: "
    If TypeOf objObject Is AcadBlockReference Then
        Set objBlock = objObject
    Else
        MsgBox "Not a valid object."
        Me.show
        Exit Sub
    End If
    
    cbStatus.Value = "PROPOSED"
    If InStr(LCase(objBlock.Layer), "existing") > 0 Then cbStatus.Value = "EXISTING"
    If InStr(LCase(objBlock.Layer), "future") > 0 Then cbStatus.Value = "FUTURE"
    
    Select Case objBlock.Name
        Case "sPed", "sHH", "sFP", "sPanel", "sMH"
        Case Else
            MsgBox "Not a valid block."
            Me.show
            Exit Sub
    End Select
    
    'tbNumber.Value = ""
    'tbStatus.Value = ""
    'cbType.Value = ""
    'tbLL.Value = ""
    'cbGround.Value = ""
    
    strType = cbType.Value
    strGround = cbGround.Value
    
    vAttList = objBlock.GetAttributes
    
    Select Case vAttList(0).TextString
        Case "PED", "HH", "FP", "PANEL", "MH"
            tbNumber.Value = strRoute & iNumber
        Case Else
            tbNumber.Value = vAttList(0).TextString
        
            Dim v00, v01 As Variant
            Dim strLine As String
            Dim iLen As Integer
    
            strLine = tbNumber.Value
    
            If Right(strLine, 1) = "X" Then
                strRoute = strLine
                iNumber = 1
            Else
                v01 = Split(strLine, "/")
                v00 = Split(v01(UBound(v01)), "L")
                v01 = Split(v00(UBound(v00)), "R")
                v00 = Split(v01(UBound(v01)), "-")
                strTemp = v00(UBound(v00))
        
                iLen = Len(strLine) - Len(strTemp)
        
                strRoute = Left(strLine, iLen)
                iNumber = Int(strTemp) + 1
            End If
    End Select
    
    tbStatus.Value = vAttList(1).TextString
    
    cbType.Value = vAttList(2).TextString
    
    If vAttList(3).TextString = "" Then
        Dim dN, dE As Double
        Dim vLL As Variant
        
        dE = objBlock.InsertionPoint(0)
        dN = objBlock.InsertionPoint(1)
        vLL = TN83FtoLL(CDbl(dN), CDbl(dE))
        
        tbLL.Value = vLL(0) & "," & vLL(1)
    Else
        tbLL.Value = vAttList(3).TextString
    End If
    
    cbGround.Value = vAttList(4).TextString
    
    '<---------------------------------------------------------------- Add Units
    
    lbUnits.Clear
    lbCables.Clear
    lbSplices.Clear
    
    If Not vAttList(7).TextString = "" Then
        vAttList(7).TextString = Replace(vAttList(7).TextString, vbLf, "")
        vLine = Split(vAttList(7).TextString, ";;")
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), "=")
            lbUnits.AddItem vItem(0)
            If InStr(vItem(1), "  ") > 0 Then
                vTemp = Split(vItem(1), "  ")
                lbUnits.List(lbUnits.ListCount - 1, 1) = vTemp(0)
                lbUnits.List(lbUnits.ListCount - 1, 2) = vTemp(1)
            Else
                lbUnits.List(lbUnits.ListCount - 1, 1) = vItem(1)
                lbUnits.List(lbUnits.ListCount - 1, 2) = ""
            End If
        Next i
    End If
    
    If cbType.Value = "" Then cbType.Value = strType
    If cbGround.Value = "" Then cbGround.Value = strGround
    
    
    'att 25
    If Not vAttList(5).TextString = "" Then
        vAttList(5).TextString = Replace(vAttList(5).TextString, vbLf, "")
        vLine = Split(vAttList(5).TextString, vbCr)
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), " / ")
            lbCables.AddItem vItem(0)
            lbCables.List(lbCables.ListCount - 1, 1) = vItem(1)
        Next i
    End If
    
    'att 26
    If Not vAttList(6).TextString = "" Then
        vAttList(6).TextString = Replace(vAttList(6).TextString, vbLf, "")
        'If InStr(vAttList(6).TextString, " + ") > 0 Then vAttList(6).TextString = Replace(vAttList(6).TextString, " + ", vbCr)
        
        vLine = Split(vAttList(6).TextString, vbCr)
        For i = 0 To UBound(vLine)
            lbSplices.AddItem vLine(i)
        Next i
    End If
    
    cbUpdateBlock.Enabled = True
    cbPlaceData.SetFocus
    
    Me.show
End Sub

Private Sub cbGetCables_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vBasePnt As Variant
    Dim strLine As String
    
    Me.Hide
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Get Cable Callout: "
    
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox "Invalid selection"
        GoTo Exit_Sub
    End If
    
    Set objBlock = objEntity
    
    If Not objBlock.Name = "CableCounts" Then
        MsgBox "Invalid block"
        Me.show
        Exit Sub
    End If
    
    vAttList = objBlock.GetAttributes
    
    lbCables.AddItem vAttList(1).TextString
    lbCables.List(lbCables.ListCount - 1, 1) = Replace(vAttList(0).TextString, "\P", " + ")
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbGetClosures_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vBasePnt As Variant
    Dim vLine As Variant
    Dim strLine As String
    
    Me.Hide
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Get Closure Callout: "
    
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox "Invalid selection"
        GoTo Exit_Sub
    End If
    
    Set objBlock = objEntity
    
    If Not objBlock.Name = "terminal" Then
        MsgBox "Invalid block"
        GoTo Exit_Sub
    End If
    
    vAttList = objBlock.GetAttributes
    
    strLine = Replace(vAttList(0).TextString, "\P", " + ")
    lbSplices.AddItem strLine
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbGetUnits_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vLine, vItem As Variant
    Dim strTemp, strLine As String
    Dim iFeet, iInch, iTotal, iRL As Integer
    
    Me.Hide
    
  On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        
    objSS.SelectOnScreen
    For Each objEntity In objSS
        If Not TypeOf objEntity Is AcadBlockReference Then GoTo Next_objEntity
        
        Set objBlock = objEntity
        If objBlock.Name = "pole_unit" Then
            vAttList = objBlock.GetAttributes
            
            vLine = Split(vAttList(3).TextString, "=")
            lbUnits.AddItem vLine(0), 0
            
            vLine(1) = Replace(vLine(1), "  ", " ")
            vItem = Split(vLine(1), " ")
            lbUnits.List(0, 1) = vItem(0)
            If UBound(vItem) < 1 Then
                lbUnits.List(0, 2) = ""
            Else
                lbUnits.List(0, 2) = vItem(1)
            End If
        End If
Next_objEntity:
    Next objEntity
    
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbGround_Change()
    If cbGround.Value = "" Then Exit Sub
    Select Case cbStatus.Value
        Case "EXISTING", "FUTURE", ""
            Exit Sub
    End Select
    
    Dim strTemp As String
    
    If lbUnits.ListCount > -1 Then
        For i = 0 To lbUnits.ListCount - 1
            If InStr(lbUnits.List(i, 0), "+BM2(") > 0 Then
                lbUnits.RemoveItem i
                GoTo Exit_For
            End If
            If lbUnits.List(i, 0) = "+BM2A" Then
                lbUnits.RemoveItem i
                GoTo Exit_For
            End If
        Next i
    End If
Exit_For:

    Select Case cbGround.Value
        Case "TGB"
            strTemp = "+BM2(5/8)(8)"
        Case "MGN", "MGNV", "OTHER"
            strTemp = "+BM2A"
        Case Else
            Exit Sub
    End Select
        
    lbUnits.AddItem strTemp
    lbUnits.List(lbUnits.ListCount - 1, 1) = 1
End Sub

Private Sub cbPlaceData_Click()
    Dim insertionPnt(0 To 2) As Double
    Dim lwpPnt(0 To 3) As Double
    Dim dScale As Double
    Dim objPoint As AcadEntity
    Dim blockRefObj As AcadBlockReference
    Dim obrUnit As AcadBlockReference
    Dim lineObj As AcadLWPolyline
    Dim returnPoint, attList As Variant
    Dim attArray() As Variant
    Dim str, str1, str2, str3 As String
    Dim strArray() As String
    Dim strPoleNum, strPosition As String
    Dim strBlock, strLine As String
    Dim n, counter As Integer
    Dim layerObj As AcadLayer
    Dim vTemp As Variant
    Dim dLat, dLong As Double
    
  On Error Resume Next
  
    If cbScale.Value = "" Then
        dScale = 1#
    Else
        dScale = CDbl(cbScale.Value) / 100
    End If
    
    strPoleNum = tbNumber.Value
    str = "pole_info"
    
    Select Case cbStatus.Value
        Case "FUTURE"
            strLayer = "Integrity Future"
        Case "EXISTING"
            strLayer = "Integrity Existing"
        Case Else
            strLayer = "Integrity Proposed"
    End Select
    
    Me.Hide

    returnPoint = ThisDrawing.Utility.GetPoint(, "Select point:")
    n = 0
    For Each Item In returnPoint
        insertionPnt(n) = Item
        n = n + 1
    Next Item
    
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        
    attList = blockRefObj.GetAttributes
    attList(0).TextString = strPoleNum
    attList(1).TextString = "0.0"
    attList(2).TextString = strPoleNum
    attList(3).TextString = ""
    
    blockRefObj.Layer = strLayer
    
    blockRefObj.Update
    insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    
    If Not cbType.Value = "" Then
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        
        attList = blockRefObj.GetAttributes
        attList(0).TextString = strPoleNum
        attList(1).TextString = "0.1"
        attList(2).TextString = cbType.Value
        attList(3).TextString = ""
    
        blockRefObj.Layer = strLayer
        
        blockRefObj.Update
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    End If
    
    If Not cbGround.Value = "" Then
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        
        If cbGround.Value = "NONE" Then cbGround.Value = "NO GRD"
        
        attList = blockRefObj.GetAttributes
        attList(0).TextString = strPoleNum
        attList(1).TextString = "0.2"
        attList(2).TextString = cbGround.Value
        attList(3).TextString = ""
    
        blockRefObj.Layer = strLayer
        
        blockRefObj.Update
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    End If
    
    If cbLL.Value = True Then
        If Not tbLL.Value = "" Then
            vTemp = Split(tbLL.Value, ",")
            dLat = CLng(CDbl(vTemp(0)) * 10000000) / 10000000
            dLong = CLng(CDbl(vTemp(1)) * 10000000) / 10000000
            
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        
            attList = blockRefObj.GetAttributes
            attList(0).TextString = strPoleNum
            attList(1).TextString = "0.99"
            attList(2).TextString = dLat & "," & dLong
            attList(3).TextString = ""
    
            blockRefObj.Layer = strLayer
        
            blockRefObj.Update
            insertionPnt(1) = insertionPnt(1) - (9 * dScale)
        End If
    End If
    
    lwpPnt(0) = insertionPnt(0) - (4 * dScale)
    lwpPnt(1) = insertionPnt(1) + (7 * dScale)
    lwpPnt(2) = lwpPnt(0) + (100 * dScale)
    lwpPnt(3) = lwpPnt(1)
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpPnt)
    
    lineObj.Layer = blockRefObj.Layer
    
    lineObj.Update
    
    insertionPnt(1) = insertionPnt(1) - (1 * dScale)
    
    str = "pole_unit"
    If cbAddUnits.Value = False Then
        Me.show
        Exit Sub
    End If
    
    If Not strLayer = "Integrity Future" Then strLayer = "Integrity Pole-Units"
    
    For i = 0 To (lbUnits.ListCount - 1)
        Set obrUnit = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        
        strLine = lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
        If Not lbUnits.List(i, 2) = "" Then strLine = strLine & "  " & lbUnits.List(i, 2)
        
        attList = obrUnit.GetAttributes
        attList(0).TextString = strPoleNum
        attList(1).TextString = "2"     '<------------------ add i if needed
        attList(2).TextString = "N/A"
        attList(3).TextString = strLine
        obrUnit.Layer = strLayer
        obrUnit.Update
        
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    Next i
    
    Me.show
End Sub

Private Sub cbPlacePlant_Click()
    If tbNumber.Value = "" Then
        MsgBox "Need a Buried Plant Number."
        Exit Sub
    End If
    
    If cbType.Value = "" Then
        MsgBox "Need a Type."
        Exit Sub
    End If
    
    If cbGround.Value = "" Then
        MsgBox "Need a Ground Source."
        Exit Sub
    End If
    
    Dim objNew As AcadBlockReference
    Dim strBlock, strRoute, strShort As String
    Dim strCables, strUnits, strSplices As String
    Dim strChar, strLayer As String
    Dim vInsertPnt, vAttList As Variant
    Dim vSlash, vDash, vL, vR As Variant
    Dim dScale As Double
    
    dScale = CDbl(cbScale.Value) / 100
    
    tbNumber.Value = UCase(tbNumber.Value)
    
    vSlash = Split(tbNumber.Value, "/")
    vDash = Split(vSlash(UBound(vSlash)), "-")
    vL = Split(vDash(UBound(vDash)), "L")
    vR = Split(vL(UBound(vL)), "R")
    strShort = vR(UBound(vR))
    strRoute = Left(tbNumber.Value, Len(tbNumber.Value) - Len(strShort))
    
    If lbCables.ListCount > 0 Then
        strCables = lbCables.List(0, 0) & " / " & lbCables.List(0, 1)
        If lbCables.ListCount > 1 Then
            For i = 1 To lbCables.ListCount - 1
                strCables = strCables & vbCr & lbCables.List(i, 0) & " / " & lbCables.List(i, 1)
            Next i
        End If
    Else
        strCables = ""
    End If
    
    If lbUnits.ListCount > 0 Then
        strUnits = lbUnits.List(0, 0) & "=" & lbUnits.List(0, 1)
        If Not lbUnits.List(0, 2) = "" Then strUnits = strUnits & "  " & lbUnits.List(0, 2)
        If lbUnits.ListCount > 1 Then
            For i = 1 To lbUnits.ListCount - 1
                strUnits = strUnits & ";;" & lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
                If Not lbUnits.List(i, 2) = "" Then strUnits = strUnits & "  " & lbUnits.List(i, 2)
            Next i
        End If
    Else
        strUnits = ""
    End If
    
    If lbSplices.ListCount > 0 Then
        strSplices = lbSplices.List(0)
        If lbSplices.ListCount > 1 Then
            For i = 1 To lbSplices.ListCount - 1
                strSplices = strSplices & vbCr & lbSplices.List(i)
            Next i
        End If
    Else
        strSplices = ""
    End If
    
    strBlock = "sPed"
    
    If InStr(cbType.Value, "UHF") > 0 Then strBlock = "sHH"
    If InStr(cbType.Value, "BHF") > 0 Then strBlock = "sHH"
    If InStr(cbType.Value, "FP") > 0 Then strBlock = "sFP"
    If InStr(cbType.Value, "FIBER") > 0 Then strBlock = "sPanel"
    If InStr(cbType.Value, "BUDI") > 0 Then strBlock = "sPanel"
    
    Me.Hide

On Error Resume Next
    
Add_Another:
    
    vInsertPnt = ThisDrawing.Utility.GetPoint(, vbCr & "Selct Insertion Point: ")
    If Not Err = 0 Then
        Me.show
        Exit Sub
    End If
    
    Set objNew = ThisDrawing.ModelSpace.InsertBlock(vInsertPnt, strBlock, dScale, dScale, dScale, 0#)
    vAttList = objNew.GetAttributes
    
    Dim dN, dE As Double
    Dim vLL As Variant
        
    dE = objNew.InsertionPoint(0)
    dN = objNew.InsertionPoint(1)
    vLL = TN83FtoLL(CDbl(dN), CDbl(dE))
        
    tbLL.Value = vLL(0) & "," & vLL(1)
    
    vAttList(0).TextString = tbNumber.Value
    
    Select Case tbStatus.Value
        Case ""
            Case "EXISTING"
    End Select
    If Not tbStatus.Value = "" Then vAttList(1).TextString = tbStatus.Value
    vAttList(2).TextString = cbType.Value
    vAttList(3).TextString = tbLL.Value
    vAttList(4).TextString = cbGround.Value
    vAttList(5).TextString = strCables
    vAttList(6).TextString = strSplices
    vAttList(7).TextString = strUnits
    
    If strBlock = "sFP" Then vAttList(1).TextString = strShort
    
    Select Case cbStatus.Value
        Case "EXISTING"
            strLayer = "Integrity Existing-Buried"
        Case "FUTURE"
            strLayer = "Integrity Future"
        Case Else
            strLayer = "Integrity Proposed-Buried"
    End Select
    objNew.Layer = strLayer
    
    objNew.Update
    
    If IsNumeric(strShort) Then
        strShort = CInt(strShort) + 1
        tbNumber.Value = strRoute & strShort
    Else
        strChar = Chr(Asc(Right(strShort, 1)) + 1)
        strShort = Left(strShort, Len(strShort) - 1) & strChar
        tbNumber.Value = strRoute & strShort
    End If
    
    GoTo Add_Another
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRoatateAtt_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    'Dim vLine As Variant
    
    Me.Hide
    
    On Error Resume Next
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Callout:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    
    Select Case objBlock.Name
        Case "sPole", "sPed", "sHH"
        Case Else
            GoTo Exit_Sub
    End Select
    
    vAttList = objBlock.GetAttributes
    
    For i = 0 To UBound(vAttList)
        vAttList(i).Rotation = 0#
    Next i
    objBlock.Update
    
    GoTo Get_Another
Exit_Sub:
    Me.show
End Sub

Private Sub cbType_Change()
    If cbType.Value = "" Then Exit Sub
    Select Case cbStatus.Value
        Case "EXISTING", "FUTURE", ""
            Exit Sub
    End Select
    
    Dim strTemp As String
    
    strTemp = "+" & cbType.Value
    
    If lbUnits.ListCount > -1 Then
        For i = 0 To lbUnits.ListCount - 1
            If InStr(lbUnits.List(i, 0), "+BDO") > 0 Then
                lbUnits.RemoveItem i
                GoTo Exit_For
            End If
            If InStr(lbUnits.List(i, 0), "+BHF") > 0 Then
                lbUnits.RemoveItem i
                GoTo Exit_For
            End If
            If InStr(lbUnits.List(i, 0), "+UHF") > 0 Then
                lbUnits.RemoveItem i
                GoTo Exit_For
            End If
            If InStr(lbUnits.List(i, 0), "+FP") > 0 Then
                lbUnits.RemoveItem i
                GoTo Exit_For
            End If
            If InStr(lbUnits.List(i, 0), "+FIBER") > 0 Then
                lbUnits.RemoveItem i
                GoTo Exit_For
            End If
            If InStr(lbUnits.List(i, 0), "+BUDI") > 0 Then
                lbUnits.RemoveItem i
                GoTo Exit_For
            End If
        Next i
    End If
Exit_For:
        
    lbUnits.AddItem strTemp
    lbUnits.List(lbUnits.ListCount - 1, 1) = 1
End Sub

Private Sub cbUpdateBlock_Click()
    Dim vAttList As Variant
    Dim strLine As String
    
    vAttList = objBlock.GetAttributes
    
    vAttList(0).TextString = tbNumber.Value
    If Not tbStatus.Value = "" Then vAttList(1).TextString = tbStatus.Value
    If Not cbType.Value = "" Then vAttList(2).TextString = cbType.Value
    If Not tbLL.Value = "" Then vAttList(3).TextString = tbLL.Value
    If Not cbGround.Value = "" Then vAttList(4).TextString = cbGround.Value
    
    If lbUnits.ListCount > 0 Then
        strLine = lbUnits.List(0, 0) & "=" & lbUnits.List(0, 1)
        If Not lbUnits.List(0, 2) = "" Then strLine = strLine & "  " & lbUnits.List(0, 2)
        
        If lbUnits.ListCount > 1 Then
            For i = 1 To lbUnits.ListCount - 1
                strLine = strLine & ";;" & lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
                If Not lbUnits.List(i, 2) = "" Then strLine = strLine & "  " & lbUnits.List(i, 2)
            Next i
        End If
        
        vAttList(7).TextString = strLine
    End If
    
    strLine = ""
    
    If lbCables.ListCount > 0 Then
        For i = 0 To lbCables.ListCount - 1
            If strLine = "" Then
                strLine = lbCables.List(i, 0) & " / " & lbCables.List(i, 1)
            Else
                strLine = strLine & vbCr & lbCables.List(i, 0) & " / " & lbCables.List(i, 1)
            End If
        Next i
        
        vAttList(5).TextString = strLine
    End If
    
    strLine = ""
    
    If lbSplices.ListCount > 0 Then
        For i = 0 To lbSplices.ListCount - 1
            If strLine = "" Then
                strLine = lbSplices.List(i, 0)
            Else
                strLine = strLine & " + " & lbSplices.List(i, 0)
            End If
        Next i
        
        vAttList(6).TextString = strLine
    End If
    
    objBlock.Update
    
    'Dim v00, v01 As Variant
    'Dim strTemp As String
    'Dim iLen As Integer
    
    'strLine = tbNumber.Value
    
    'If Right(strLine, 1) = "X" Then
        'strRoute = strLine
        'iNumber = 1
    'Else
        'v01 = Split(strLine, "/")
        'v00 = Split(v01(UBound(v01)), "L")
        'v01 = Split(v00(UBound(v00)), "R")
        'v00 = Split(v01(UBound(v01)), "-")
        'strTemp = v00(UBound(v00))
        
        'iLen = Len(strLine) - Len(strTemp)
        
        'strRoute = Left(strLine, iLen)
        'iNumber = Int(strTemp) + 1
    'End If
End Sub

Private Sub Label13_Click()
    Dim objEntity As AcadEntity
    Dim objTemp As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Plant Number: "
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objTemp = objEntity
        vAttList = objTemp.GetAttributes
        
        tbNumber.Value = vAttList(0).TextString
    End If
    
    Me.show
End Sub

Private Sub Label38_Click()
    If cbUpdateBlock.Enabled = False Then Exit Sub
    
    Dim dN, dE As Double
    Dim vLL As Variant
        
    dE = objBlock.InsertionPoint(0)
    dN = objBlock.InsertionPoint(1)
    vLL = TN83FtoLL(CDbl(dN), CDbl(dE))
        
    tbLL.Value = vLL(0) & "," & vLL(1)
End Sub

Private Sub lbCables_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbSplices.ListCount < 1 Then Exit Sub
    
    Dim objCO As AcadBlockReference
    'Dim objEntity As AcadEntity
    Dim returnPnt As Variant
    Dim vBlockCoords(0 To 2) As Double
    Dim lwpCoords(0 To 3) As Double
    Dim dPrevious(0 To 2) As Double
    Dim dOrigin(0 To 2) As Double
    Dim dScale As Double
    Dim lineObj As AcadLWPolyline
    Dim objCircle As AcadCircle
    Dim objCircle2 As AcadCircle
    Dim objText As AcadText
    Dim objText2 As AcadText
    Dim vLine, vCounts, vTemp As Variant
    Dim strLetter, strCounts, strHO1 As String
    Dim strSegment As String
    Dim iHO1, iIndex As Integer
    
  'On Error Resume Next
    iIndex = lbCables.ListIndex
    
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
    Me.Hide
    returnPnt = ThisDrawing.Utility.GetPoint(, "Select Point: ")
    
    lwpCoords(0) = returnPnt(0)
    lwpCoords(1) = returnPnt(1)
    dPrevious(0) = returnPnt(0)
    dPrevious(1) = returnPnt(1)
    dPrevious(2) = 0#
    
    returnPnt = ThisDrawing.Utility.GetPoint(dPrevious, "Select Point: ")
    vBlockCoords(0) = returnPnt(0)
    vBlockCoords(1) = returnPnt(1)
    vBlockCoords(2) = returnPnt(2)
    lwpCoords(2) = vBlockCoords(0)
    lwpCoords(3) = vBlockCoords(1)
    
    If cbPlacement.Value = "Leader" Then
        Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
        lineObj.Layer = "Integrity Proposed"
        lineObj.Update
    Else
        Set objCircle = ThisDrawing.ModelSpace.AddCircle(dPrevious, 8)
        objCircle.Layer = "Integrity Proposed"
        objCircle.Update
        Set objCircle2 = ThisDrawing.ModelSpace.AddCircle(returnPnt, 8)
        objCircle2.Layer = "Integrity Proposed"
        objCircle2.Update
        
        strLetter = UCase(ThisDrawing.Utility.GetString(0, "Enter Callout Letter:"))
        Set objText = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8)
        Set objText2 = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8)
        objText.Layer = "Integrity Proposed"
        objText.Alignment = acAlignmentMiddle
        objText.TextAlignmentPoint = dPrevious
        objText2.Layer = "Integrity Proposed"
        objText2.Alignment = acAlignmentMiddle
        objText2.TextAlignmentPoint = vBlockCoords
        objText.Update
        objText2.Update
        
        vBlockCoords(0) = vBlockCoords(0) + 8
        lwpCoords(0) = 0
    End If
    
    dScale = CDbl(cbScale.Value) / 100
    If dScale = 0 Then dScale = 0.75
    
    vLine = Split(lbCables.List(iIndex, 0), ": ")
    strSegment = tbNumber.Value & ": " & vLine(0)
    strHO1 = vLine(1)
    
    vLine = Split(lbCables.List(iIndex, 1), " + ")
    'vCounts = Split(vLine(0), "] ")
    vTemp = Split(vLine(0), ": ")
    strCounts = vTemp(0) & ": " & vTemp(1)
    
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vLine)
            vTemp = Split(vLine(i), ": ")
            strCounts = strCounts & "\P" & vTemp(0) & ": " & vTemp(1)
        Next i
    End If
    
    'MsgBox strSegment & vbCr & strHO1 & vbCr & strCounts
    
        
    Set objCO = ThisDrawing.ModelSpace.InsertBlock(vBlockCoords, "Callout", dScale, dScale, dScale, 0)
    
    objCO.Layer = "Integrity Proposed"
    attItem = objCO.GetAttributes
    attItem(0).TextString = strSegment
    attItem(1).TextString = strHO1
    attItem(2).TextString = strCounts
    
    If cbPlacement.Value = "Leader" Then
        If lwpCoords(2) < lwpCoords(0) Then
            vBlockCoords(0) = vBlockCoords(0) - (75 * dScale)
            objCO.InsertionPoint = vBlockCoords
        End If
    End If
    objCO.Update
    
    Me.show
End Sub

Private Sub lbSplices_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbSplices.ListCount < 1 Then Exit Sub
    
    Dim objCO As AcadBlockReference
    'Dim objEntity As AcadEntity
    Dim returnPnt As Variant
    Dim vBlockCoords(0 To 2) As Double
    Dim lwpCoords(0 To 3) As Double
    Dim dPrevious(0 To 2) As Double
    Dim dOrigin(0 To 2) As Double
    Dim dScale As Double
    Dim lineObj As AcadLWPolyline
    'Dim objCircle As AcadCircle
    'Dim objCircle2 As AcadCircle
    'Dim objText As AcadText
    'Dim objText2 As AcadText
    Dim vLine, vCounts, vTemp As Variant
    Dim strLetter, strCounts, strHO1 As String
    Dim strSegment As String
    Dim iHO1 As Integer
    
  'On Error Resume Next
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
    Me.Hide
    'returnPnt = ThisDrawing.Utility.GetPoint(, "Select Point: ")
    returnPnt = objBlock.InsertionPoint
    
    lwpCoords(0) = returnPnt(0)
    lwpCoords(1) = returnPnt(1)
    dPrevious(0) = returnPnt(0)
    dPrevious(1) = returnPnt(1)
    dPrevious(2) = 0#
    
    returnPnt = ThisDrawing.Utility.GetPoint(dPrevious, "Select Point: ")
    vBlockCoords(0) = returnPnt(0)
    vBlockCoords(1) = returnPnt(1)
    vBlockCoords(2) = returnPnt(2)
    lwpCoords(2) = vBlockCoords(0)
    lwpCoords(3) = vBlockCoords(1)
    
    'If cbPlacement.Value = "Leader" Then
        Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
        lineObj.Layer = "Integrity Proposed"
        lineObj.Update
    'Else
        'Set objCircle = ThisDrawing.ModelSpace.AddCircle(dPrevious, 8)
        'objCircle.Layer = "Integrity Proposed"
        'objCircle.Update
        'Set objCircle2 = ThisDrawing.ModelSpace.AddCircle(returnPnt, 8)
        'objCircle2.Layer = "Integrity Proposed"
        'objCircle2.Update
        
        'strLetter = UCase(ThisDrawing.Utility.GetString(0, "Enter Callout Letter:"))
        'Set objText = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8)
        'Set objText2 = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8)
        'objText.Layer = "Integrity Proposed"
        'objText.Alignment = acAlignmentMiddle
        'objText.TextAlignmentPoint = dPrevious
        'objText2.Layer = "Integrity Proposed"
        'objText2.Alignment = acAlignmentMiddle
        'objText2.TextAlignmentPoint = vBlockCoords
        'objText.Update
        'objText2.Update
        
        'vBlockCoords(0) = vBlockCoords(0) + 8
        'lwpCoords(0) = 0
    'End If
    
    dScale = CDbl(cbScale.Value) / 100
    If dScale = 0 Then dScale = 0.75
    
    vLine = Split(lbSplices.List(lbSplices.ListIndex), "] ")
    strSegment = tbNumber.Value & ": " & Replace(vLine(0), "[", "")
    
    vLine = Split(lbSplices.List(lbSplices.ListIndex), " + ")
    vCounts = Split(vLine(0), "] ")
    vTemp = Split(vCounts(1), ": ")
    strCounts = vTemp(0) & ": " & vTemp(1)
    
    vCounts = Split(vTemp(1), "-")
    
    If UBound(vCounts) = 0 Then
        iHO1 = 1
    Else
        iHO1 = CInt(vCounts(1)) - CInt(vCounts(0)) + 1
    End If
    
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vCounts)
            vTemp = Split(vLine(i), ": ")
            strCounts = strCounts & "\P" & vTemp(0) & ": " & vTemp(1)
            
            vCounts = Split(vTemp(1), "-")
    
            If UBound(vCounts) = 0 Then
                iHO1 = iHO1 + 1
            Else
                iHO1 = iHO1 + CInt(vCounts(1)) - CInt(vCounts(0)) + 1
            End If
        Next i
    End If
    
    strHO1 = "+HO1B=" & iHO1
    
    'MsgBox strSegment & vbCr & strHO1 & vbCr & strCounts
    
        
    Set objCO = ThisDrawing.ModelSpace.InsertBlock(vBlockCoords, "Callout", dScale, dScale, dScale, 0)
    
    objCO.Layer = "Integrity Proposed"
    attItem = objCO.GetAttributes
    attItem(0).TextString = strSegment
    attItem(1).TextString = strHO1
    attItem(2).TextString = strCounts
    
    If lwpCoords(2) < lwpCoords(0) Then
        vBlockCoords(0) = vBlockCoords(0) - (75 * dScale)
        objCO.InsertionPoint = vBlockCoords
    End If
    objCO.Update
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbUnits.ColumnCount = 3
    lbUnits.ColumnWidths = "96;36;48"
    
    lbCables.ColumnCount = 2
    lbCables.ColumnWidths = "96;285"
    
    cbStatus.AddItem "EXISTING"
    cbStatus.AddItem "PROPOSED"
    cbStatus.AddItem "FUTURE"
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    cbType.AddItem "BDO3"
    cbType.AddItem "BDO5"
    cbType.AddItem "BDO7"
    cbType.AddItem "BDO(OP10106)"
    cbType.AddItem "BDO(OP12126)"
    cbType.AddItem "BHF(30X48X36)T"
    cbType.AddItem "UHF(17X30X18)"
    cbType.AddItem "UHF(24X36X24)"
    cbType.AddItem "UHF(30X48X24)T"
    cbType.AddItem "UHF(30X48X36)"
    cbType.AddItem "UHF(48X96X48)TIER22"
    cbType.AddItem "FP"
    cbType.AddItem "BUDI BOX"
    cbType.AddItem "FIBER PANEL"
    'cbType.AddItem "FP (Future)"
    
    cbGround.AddItem "TGB"
    cbGround.AddItem "MGN"
    cbGround.AddItem "MGNV"
    cbGround.AddItem "OTHER"
    cbGround.AddItem "NO GRD"
    'cbGround.Value = "NONE"
    
    cbPlacement.AddItem "Leader"
    cbPlacement.AddItem "Away"
    cbPlacement.Value = "Leader"
    
    strRoute = "01A/"
    iNumber = 1
    
    cbGetBlock.SetFocus
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
