Private Sub cbADSS_Click()
    If cbADSS.Value = True Then
        cbGuy.Value = False
        cbGuy.Enabled = False
        
        tbLashedCables.Value = ""
        tbLashedCables.Enabled = False
    Else
        cbGuy.Enabled = True
        If tbStrandCable.Value = "" Then
            cbGuy.Value = False
        Else
            cbGuy.Value = True
        End If
        
        tbLashedCables.Enabled = True
    End If
End Sub

Private Sub cbPlace_Click()
    On Error Resume Next
    
    ' Ensure layer 0 exists and is current
    Dim layerObj As AcadLayer
    If Not LayerExists("0") Then
        Set layerObj = ThisDrawing.Layers.Add("0")
    End If
    ThisDrawing.ActiveLayer = ThisDrawing.Layers.Item("0")
    
    ' Set annotation properties
    With VsagAnno
        .Layer = "0"
        .Visible = True
        .color = acByLayer
        .Update
    End With
    
    ' Force drawing update
    ThisDrawing.ActiveViewport.ZoomExtents
    ThisDrawing.Regen acAllViewports
    ThisDrawing.Application.Update
    
    On Error GoTo 0
    
    ' Continue with existing code
    Dim objEntity As AcadEntity
    Dim objPole As AcadBlockReference
    Dim objPrevious As AcadBlockReference
    Dim vAttItem, vBasePnt, vTemp As Variant
    Dim vAttPrevious As Variant
    Dim dPnt1(0 To 2), dPnt2(0 To 2), dPnt3(0 To 2) As Double
    Dim X(0 To 2), Y(0 To 2), Z(0 To 2) As Double
    Dim dInsertPnt(0 To 2), dLeaderPnt(0 To 3) As Double
    Dim dTextPnt(0 To 2) As Double
    Dim xDiff, yDiff, dDist As Double
    Dim dRotate, dScale, distance As Double
    'Dim dNewVertex(0 To 1) As Double
    Dim iCounter, iTemp As Integer
    Dim objLine As AcadLWPolyline
    Dim objSpan As AcadBlockReference
    Dim strLayer, strTextLayer, strLine, strTemp As String
    Dim strCables As String
    Dim strResult As String
    Dim dInfo(0 To 6) As Double
    Dim vInfo As Variant
    
    iCounter = 0
    dScale = CDbl(cbScale.Value) / 100
    
    Select Case cbStatus.Value
        Case "Proposed"
            strLayer = 0
            strTextLayer = 0
        Case "Existing"
            strLayer = 0
            strTextLayer = 0
        Case "Removed"
            strLayer = 0
            strTextLayer = 0
    End Select
    
    Me.Hide
    
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "From Pole: "
    If Not Err = 0 Then GoTo Exit_While
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objPrevious = objEntity
    Else
        GoTo Exit_While
    End If
    
    X(1) = objPrevious.InsertionPoint(0)
    Y(1) = objPrevious.InsertionPoint(1)
    Z(1) = 0#
    
Add_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "To Pole: "
    If Not Err = 0 Then GoTo Exit_While
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objPole = objEntity
    Else
        GoTo Exit_While
    End If
    
    X(0) = objPole.InsertionPoint(0)
    Y(0) = objPole.InsertionPoint(1)
    Z(0) = 0#
    
    If iCounter = 0 Then
        X(2) = X(0)
        Y(2) = Y(0)
        Z(2) = 0#
    End If
    
    dLeaderPnt(0) = X(1)
    dLeaderPnt(1) = Y(1)
    dLeaderPnt(2) = X(0)
    dLeaderPnt(3) = Y(0)
    
    If cbCable.Value = True Then
        Set objLine = ThisDrawing.ModelSpace.AddLightWeightPolyline(dLeaderPnt)
        objLine.Layer = strLayer
        objLine.Update
    End If
    
    dTextPnt(0) = (dLeaderPnt(0) + dLeaderPnt(2)) / 2
    dTextPnt(1) = (dLeaderPnt(1) + dLeaderPnt(3)) / 2
    dTextPnt(2) = 0#
      
    xDiff = dLeaderPnt(2) - dLeaderPnt(0)
    yDiff = dLeaderPnt(3) - dLeaderPnt(1)
        
    If (yDiff = 0) Then
        dRotate = 0
    ElseIf (xDiff = 0) Then
        dRotate = Pi / 2
    Else
        dRotate = Atn(yDiff / xDiff)
    End If
    distance = Int(Math.Sqr(xDiff ^ 2 + yDiff ^ 2)) + 1
    
    strCables = ""
    If Not tbStrandCable.Value = "" Then
        If cbADSS.Value = False Then
            strCables = "CO(" & tbStrandCable.Value & ")M"
        Else
            strCables = "ADSS(" & tbStrandCable.Value & ")"
        End If
    End If
    
    If Not tbLashedCables.Value = "" Then
        vTemp = Split(tbLashedCables.Value, " ")
        For i = 0 To UBound(vTemp)
            If strCables = "" Then
                strCables = "CO(" & vTemp(i) & ")E"
            Else
                strCables = strCables & " CO(" & vTemp(i) & ")E"
            End If
        Next i
    End If
    'saginput_Change()"

    Set objSpan = ThisDrawing.ModelSpace.InsertBlock(dTextPnt, "cable_span", dScale, dScale, dScale, dRotate)
    objSpan.Layer = "0"
    
    vAttItem = objSpan.GetAttributes
    VsagAnno = saginput * distance
    
    'For Each att In vAttItem
     '   att.Layer = "0"
       '  att.StyleName = "STANDARD"
       ' att.Height = 2.5
    'Next att
    
    vAttItem(0).TextString = ""
    vAttItem(1).TextString = strCables
    vAttItem(2).TextString = distance + VsagAnno & "'"
    vAttItem(3).TextString = "INCOMPLETE"
    objSpan.Update
    
    If cbGuy.Value = True Then
        If iCounter = 0 Then
            strTemp = PlaceGuy((X(1)), (Y(1)), (X(0)), (Y(0)), (X(0)), (Y(0)))
            'strTemp = "+PM2A=1;;"
        Else
            strTemp = PlaceGuy((X(1)), (Y(1)), (X(0)), (Y(0)), (X(2)), (Y(2)))
            'strTemp = ""
        End If
            
        If Not strTemp = "" Then
            vAttPrevious = objPrevious.GetAttributes
            
            If vAttPrevious(27).TextString = "" Then
                vAttPrevious(27).TextString = strTemp
            Else
                vAttPrevious(27).TextString = vAttPrevious(27).TextString & ";;" & strTemp
            End If
        End If
    End If
    iCounter = iCounter + 1
    
    strResult = ThisDrawing.Utility.GetString(0, "Need a Closure [y/n]:(n)")
    If Err = 0 Then
        If LCase(strResult) = "y" Then
            dInfo(0) = distance
            dInfo(1) = dScale
            dInfo(2) = dRotate
            dInfo(3) = X(0)
            dInfo(4) = Y(0)
            dInfo(5) = X(1)
            dInfo(6) = Y(1)
            
            vInfo = dInfo
            Call PlaceClosure(vInfo)
            
            If strLine = "" Then
                strLine = "+CO(" & tbCoilCable.Value & ")E=100  LOOP;;+HACO(" & tbCoilCable.Value & ")=1  " & cbClosure.Value
            Else
                strLine = "+CO(" & tbCoilCable.Value & ")E=100  LOOP;;+HACO(" & tbCoilCable.Value & ")=1  " & cbClosure.Value & ";;" & strLine
            End If
        End If
    Else
        Err = 0
    End If
    
    If Not tbLashedCables.Value = "" Then
        vTemp = Split(tbLashedCables.Value, " ")
        For i = 0 To UBound(vTemp)
            If strLine = "" Then
                strLine = "+CO(" & vTemp(i) & ")E=" & distance & "'"
            Else
                strLine = "+CO(" & vTemp(i) & ")E=" & distance & ";;" & strLine
            End If
        Next i
    End If
    
    vAttList = objPole.GetAttributes
    
    If Not tbStrandCable.Value = "" Then
        If strLine = "" Then
            If cbADSS.Value = False Then
                strLine = "+CO(" & tbStrandCable.Value & ")6M-EHS=" & distance
            Else
                strLine = "+ADSS(" & tbStrandCable.Value & ")=" & distance
            End If
        Else
            If cbADSS.Value = False Then
                strLine = "+CO(" & tbStrandCable.Value & ")6M-EHS=" & distance & ";;" & strLine
            Else
                strLine = "+ADSS(" & tbStrandCable.Value & ")" & distance & ";;" & strLine
            End If
        End If
        
        If cbADSS.Value = False Then
            Select Case Left(UCase(vAttList(8).TextString), 1)
                Case "M", "m"
                    strLine = strLine & ";;+PM2A=1"
            End Select
        End If
    End If
    
    If vAttList(27).TextString = "" Then
        vAttList(27).TextString = strLine
    Else
        vAttList(27).TextString = vAttList(27).TextString & ";;" & strLine
    End If
    
    objPole.Update
    
    Err = 0
    strLine = ""
    iCounter = iCounter + 1
    
    X(2) = X(1)
    Y(2) = Y(1)
    Z(2) = Z(1)
    
    X(1) = X(0)
    Y(1) = Y(0)
    Z(1) = Z(0)
    
    Set objPrevious = objPole
    
    GoTo Add_Another    '<--------------------------------------------------------------Get Another Pole
    
Exit_While:
    If cbADSS.Value = True Then iCounter = 0
    
    If Not iCounter = 0 Then
        If cbGuy Then
            iTemp = PlaceGuy((X(1)), (Y(1)), (X(2)), (Y(2)), (X(2)), (Y(2)))
        
            strLine = ";;+PE1-3G=1;;+PF1-5A=1;;+PM11=1"
        End If
        
        vAttList = objPole.GetAttributes
        vAttList(27).TextString = vAttList(27).TextString & strLine
        objPole.Update
    End If
    
    Me.show

    ' Switch to correct layout and make active
    Dim oLayout As AcadLayout
    Set oLayout = ThisDrawing.Layouts.Item(8)
    ThisDrawing.ActiveLayout = oLayout

    ' Get viewport and ensure it's on
    Dim oViewport As AcadViewport
    Set oViewport = oLayout.Block.Item(1)
    oViewport.On = True

    ' Set viewport properties
    With oViewport
        .CustomScale = 1    ' or desired scale
        .Display = True
        .Layer = "0"
    End With

    ' Ensure layers are visible
    ' ThisDrawing.ActiveViewport.ZoomExtents
    ThisDrawing.Regen acAllViewports

    ' Update display
    ThisDrawing.Application.Update
    ' ThisDrawing.Application.ZoomExtents

    ThisDrawing.Regen True

    ' Layer setup
    strTextLayer = "0"
    ThisDrawing.ActiveLayer = ThisDrawing.Layers.Item(strTextLayer)

    ' Force updates
    ThisDrawing.Regen acActiveViewport
    ThisDrawing.Application.ZoomExtents
    ThisDrawing.Application.Update
    
    ' Position text
    dInsertPnt(0) = dInsertPnt(0) + (15 * dScale * dDiff(0) / dDiff(2))
    dInsertPnt(1) = dInsertPnt(1) + (15 * dScale * dDiff(1) / dDiff(2))
    
    ' Force updates
    ThisDrawing.ActiveLayer = ThisDrawing.Layers.Item(strTextLayer)
    ThisDrawing.Regen acActiveViewport
    ThisDrawing.Application.Update
    
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description
    End If
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub


Private Function PlaceGuy(X0 As Double, Y0 As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
    Dim startPnt(0 To 2), midPnt(0 To 2), endPnt(0 To 2) As Double
    Dim firstAng, endAng, defAng, rotateAng, ancAng, pull, tempPull As Double
    Dim dRotate, Pi, dTemp As Double
    Dim returnPoint As Variant
    Dim insertPnt(0 To 2) As Double
    Dim n, result, iDirection As Integer  'iDirection 0=L 1=R
    Dim str, str2, strTemp As String
    Dim layerObj As AcadLayer
    Dim strPull, strResult As String
    Dim vPoleCoords As Variant
    Dim entPole As AcadObject
    Dim obrGP As AcadBlockReference
    Dim attItem, basePnt As Variant
    Dim iCount As Integer
    
    Pi = 3.14159265359
    
'On Error GoTo Exit_Sub
    
    If (X1 - X0) = 0 Then
        If (Y1 - Y0) < 0 Then
            firstAng = 3 * Pi / 2
        Else
            firstAng = Pi / 2
        End If
    Else
        firstAng = Math.Atn((Y1 - Y0) / (X1 - X0))
        If (X1 - X0) < 0 Then
            firstAng = Pi + firstAng
        ElseIf (Y1 - Y0) < 0 Then
            firstAng = 2 * Pi + firstAng
        End If
    End If
    
    If (X2 - X0) = 0 Then
        If (Y2 - Y0) < 0 Then
            endAng = 3 * Pi / 2
        Else
            endAng = Pi / 2
        End If
    Else
        endAng = Math.Atn((Y2 - Y0) / (X2 - X0))
        If (X2 - X0) < 0 Then
            endAng = Pi + endAng
        ElseIf (Y2 - Y0) < 0 Then
            endAng = 2 * Pi + endAng
        End If
    End If
    
    Select Case Abs(endAng - firstAng)
        Case Is = 0
            rotateAng = firstAng + Pi
        Case Is < Pi
            rotateAng = (endAng + firstAng) / 2 + Pi
        Case Else
            rotateAng = (endAng + firstAng) / 2
    End Select
    
    If rotateAng > (2 * Pi) Then rotateAng = rotateAng - (2 * Pi)
    If rotateAng < 0 Then rotateAng = rotateAng + (2 * Pi)
    
    defAng = firstAng + Pi - endAng
    pull = Math.Round(Math.Abs(100 * Math.Sin(defAng / 2)))
    
    If pull < 3 Then
        PlaceGuy = ""
        Exit Function
    End If
    
    If (rotateAng > (Pi / 2)) And (rotateAng <= (3 * Pi / 2)) Then
        str = "ExGuyOL"
        str2 = "ExAncOL"
    Else
        str = "ExGuyOR"
        str2 = "ExAncOR"
    End If
    
    'insertPnt = ThisDrawing.Utility.GetPoint(, "Place Guy:")
    insertPnt(0) = X0
    insertPnt(1) = Y0
    insertPnt(2) = 0#
    
    If pull >= 50 Then  '<------------------------------------- Get strPull
        strPull = "DE"
        
        dTemp = Abs(firstAng - endAng)
        
        If dTemp < 0.0002 Then
            Call AddAnchorGuy(insertPnt, CStr(str2), CStr(str), CStr(strPull), CDbl(rotateAng))
        
            strTemp = "+PE1-3G=1;;+PF1-5A=1;;+PM11=1"
        Else
            Call AddAnchorGuy(insertPnt, CStr(str2), CStr(str), CStr(strPull), (firstAng + Pi))
            Call AddAnchorGuy(insertPnt, CStr(str2), CStr(str), CStr(strPull), (endAng + Pi))
        
            strTemp = "+PM2A=1;;+PE1-3G=2;;+PF1-5A=2;;+PM11=2"
        End If
    Else
        strPull = pull
        Call AddAnchorGuy(insertPnt, CStr(str2), CStr(str), CStr(strPull), CDbl(rotateAng))
        
        strTemp = "+PE1-3G=1;;+PF1-5A=1;;+PM11=1"
    End If
    
Exit_Sub:
    PlaceGuy = strTemp
End Function

Private Sub AddAnchorGuy(vInsert As Variant, strAnc As String, strGuy As String, strPull As String, dRotate As Double)
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim dScale As Double
    Dim dInsert(2) As Double
    
    dInsert(0) = vInsert(0)
    dInsert(1) = vInsert(1)
    dInsert(2) = vInsert(2)
    dScale = CDbl(cbScale.Value) / 100
    
    'MsgBox strGuy & vbCr & dScale & vbCr & dRotate
    
    'Set layerObj = ThisDrawing.Layers.Add("Integrity Proposed")
    ThisDrawing.ActiveLayer = layerObj
        
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dInsert, strGuy, dScale, dScale, dScale, dRotate)
    objBlock.Layer = "Integrity Proposed"
    vAttList = objBlock.GetAttributes
    vAttList(2).TextString = "PE1-3G"
    objBlock.Update

    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dInsert, strAnc, dScale, dScale, dScale, dRotate)
    objBlock.Layer = "Integrity Proposed"
    vAttList = objBlock.GetAttributes
    vAttList(0).TextString = "PF1-5A"
    If strPull = "DE" Then
        vAttList(1).TextString = "P= DE"
    Else
        vAttList(1).TextString = "P= " & strPull
    End If
    objBlock.Update
End Sub

Private Sub sagfactor_Click()

End Sub

Private Sub saginput_Change()
   ' Declare sagValue outside the error handling block so it's accessible even if an error occurs
    Dim sagValue As Double

    On Error Resume Next ' Turn on error handling

    ' Attempt to convert the input to a value
    sagValue = Val(Me.saginput.Text)

    ' Check for errors *immediately* after the conversion attempt
    If Err.Number <> 0 Then  ' An error occurred during conversion
        MsgBox "Invalid input. Please enter a number.", vbExclamation, "Input Error"
        Me.saginput.Text = ""  ' Clear the invalid input (optional)
        sagValue = 0 ' Or some other default value if needed
        Err.Clear  ' Clear the error object (important!)
        Exit Sub ' Exit the subroutine to prevent further errors
    End If

    On Error GoTo 0 ' Turn off error handling (good practice)

    ' ... (Your code to use sagValue - it will now be a valid number or 0) ...

    'Example:
    result = sagValue * distance ' Your calculation

End Sub

Private Sub tbCoilCable_Change()
    If tbCoilCable.Value = "" Then
        cbClosure.Value = ""
        Exit Sub
    End If
    
    Dim strPath As String
    
    strPath = LCase(ThisDrawing.Path)
    
    If InStr(strPath, "united") > 0 Then
        Select Case CInt(tbCoilCable.Value)
            Case Is < 49
                cbClosure.Value = "FOSC A"
            Case Is > 96
                cbClosure.Value = "FOSC D"
            Case Else
                cbClosure.Value = "FOSC B"
        End Select
    Else
        Select Case CInt(tbCoilCable.Value)
            Case Is < 25
                cbClosure.Value = "G4"
            Case Is > 144
                cbClosure.Value = "G6"
            Case Else
                cbClosure.Value = "G5"
        End Select
    End If
End Sub

Private Sub tbStrandCable_Change()
    If tbStrandCable.Value = "" Then
        tbCoilCable.Value = ""
        cbGuy.Value = False
        Exit Sub
    End If
    
    tbCoilCable.Value = tbStrandCable.Value
    cbGuy.Value = True
End Sub

Private Sub UserForm_Initialize()
    cbClosure.AddItem ""
    cbClosure.AddItem "#1413"
    cbClosure.AddItem "FOSC A"
    cbClosure.AddItem "FOSC B"
    cbClosure.AddItem "FOSC D"
    cbClosure.AddItem "G4"
    cbClosure.AddItem "G5"
    cbClosure.AddItem "G6"
    cbClosure.AddItem "G6S"
    
    cbCoilType.AddItem ""
    cbCoilType.AddItem "FEED"
    cbCoilType.AddItem "12F"
    cbCoilType.AddItem "24F"
    cbCoilType.AddItem "36F"
    cbCoilType.AddItem "48F"
    cbCoilType.AddItem "72F"
    cbCoilType.AddItem "96F"
    cbCoilType.AddItem "144F"
    cbCoilType.AddItem "216F"
    cbCoilType.AddItem "288F"
    cbCoilType.AddItem "432F"
    
    cbClosureType.AddItem ""
    cbClosureType.AddItem "Aerial"
    cbClosureType.AddItem "Buried"
    cbClosureType.AddItem "UG"
    cbClosureType.Value = "Aerial"
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    cbStatus.AddItem "Proposed"
    cbStatus.AddItem "Existing"
    cbStatus.AddItem "Removed"
    cbStatus.Value = "Proposed"
End Sub

Private Sub PlaceClosure(vInfo As Variant)
    Dim dDist, dScale, dRotate As Double
    Dim dInsertPnt(2) As Double
    Dim dBackPnt(2) As Double
    
    dDist = vInfo(0)
    dScale = vInfo(1)
    dRotate = vInfo(2)
    
    dInsertPnt(0) = vInfo(3)
    dInsertPnt(1) = vInfo(4)
    dInsertPnt(2) = 0#
    
    dBackPnt(0) = vInfo(5)
    dBackPnt(1) = vInfo(6)
    dBackPnt(2) = 0#
    
    
    
    Dim obrMClosure As AcadBlockReference
    Dim obrMCoil As AcadBlockReference
    'Dim objBlock, obrGP3 As AcadBlockReference
    'Dim vReturnPnt, vEndPnt As Variant
    Dim vAttList As Variant
    Dim dDiff(0 To 3) As Double
    Dim strBlockName, strTermType As String
    'Dim layerObj As AcadLayer
    'Dim returnPnt, vPoleCoords As Variant
    Dim str, strLayer As String
    Dim vBlockCoords(0 To 2) As Double
    Dim i As Integer

    Me.Hide

    strTermType = ""

  On Error Resume Next

    str = ""
    Err = 0

    dDiff(0) = dBackPnt(0) - dInsertPnt(0)
    dDiff(1) = dBackPnt(1) - dInsertPnt(1)
    dDiff(2) = Math.Sqr(dDiff(0) * dDiff(0) + (dDiff(1) * dDiff(1)))

    dInsertPnt(0) = dInsertPnt(0) + 15 * dScale * dDiff(0) / dDiff(2)
    dInsertPnt(1) = dInsertPnt(1) + 15 * dScale * dDiff(1) / dDiff(2)
    dInsertPnt(2) = 0#

    strLayer = 0  ' Change layer to 0 instead of "Integrity Map Splices-Aerial"

    strBlockName = "Map splice"
    Set obrMClosure = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, strBlockName, dScale, dScale, dScale, dRotate)
    obrMClosure.Layer = strLayer
    vAttList = obrMClosure.GetAttributes
    If Not cbClosure.Value = "" Then vAttList(0).TextString = cbClosure.Value
    obrMClosure.Update

place_coil:

    dInsertPnt(0) = dInsertPnt(0) + 15 * dScale * dDiff(0) / dDiff(2)
    dInsertPnt(1) = dInsertPnt(1) + 15 * dScale * dDiff(1) / dDiff(2)

    If dDiff(0) > 0 Then
        dInsertPnt(0) = dInsertPnt(0) + 9 * dScale * dDiff(1) / dDiff(2)
        dInsertPnt(1) = dInsertPnt(1) - 9 * dScale * dDiff(0) / dDiff(2)
    Else
        dInsertPnt(0) = dInsertPnt(0) - 9 * dScale * dDiff(1) / dDiff(2)
        dInsertPnt(1) = dInsertPnt(1) + 9 * dScale * dDiff(0) / dDiff(2)
    End If

    strLayer = 0  ' Change layer to 0 instead of "Integrity Map Coils-Aerial"

    Set obrMCoil = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "Map coil", dScale, dScale, dScale, 0#)
    obrMCoil.Layer = strLayer
    vAttList = obrMCoil.GetAttributes
    If Not tbCoil.Value = "" Then vAttList(0).TextString = tbCoil.Value & "'"
    vAttList(1).TextString = tbCoilCable.Value & "F"
    obrMCoil.Update

    ' Force redraw
    ThisDrawing.Regen True

Exit_Sub:
    
End Sub

Private Function LayerExists(layerName As String) As Boolean
    Dim lay As AcadLayer
    For Each lay In ThisDrawing.Layers
        If lay.Name = layerName Then
            LayerExists = True
            Exit Function
        End If
    Next lay
    LayerExists = False
End Function

