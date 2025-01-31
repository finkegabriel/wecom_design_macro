VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} aaPlaceCable 
   Caption         =   "Place Cable 2.1"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   OleObjectBlob   =   "aaPlaceCable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "aaPlaceCable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pi = 3.14159265358979

Private Sub PlaceCable()
    If tbNumber.Value = "" Then Exit Sub
    
    Dim objEntity As AcadEntity
    Dim objLine As AcadLWPolyline
    Dim objSpan As AcadBlockReference
    Dim objPole As AcadBlockReference
    Dim objPrevious As AcadBlockReference
    Dim strLayer, strTextLayer, strAnchorLayer As String
    Dim strCable, strCounts, strLine As String
    Dim strResult, strNote As String
    Dim strRoute As String
    Dim iNumber, iIncrement As Integer
    Dim vBasePnt As Variant
    Dim vAttItem, vAttPrevious As Variant
    Dim vInfo, vAttList As Variant
    Dim X(0 To 2), Y(0 To 2), Z(0 To 2) As Double
    Dim dInsertPnt(0 To 2), dLeaderPnt(0 To 3) As Double
    Dim dTextPnt(0 To 2) As Double
    Dim dInfo(0 To 6) As Double
    Dim xDiff, yDiff, dDist As Double
    Dim dRotate, dScale, distance As Double
    'Dim dScale As Double
    Dim iCounter As Integer
    
    strLayer = "Integrity " & cbStatus.Value
    strTextLayer = "Integrity " & cbStatus.Value
    strAnchorLayer = "Integrity " & cbStatus.Value
    If cbType.Value = "CO" Then
            If cbSuffix.Value = "E" Then strAnchorLayer = "Integrity Existing"
    End If
    
    strCable = cbType.Value & "(" & cbSize.Value & ")"
    If Not cbSuffix.Value = "" Then strCable = strCable & cbSuffix.Value
    
    strCounts = Replace(tbCounts.Value, vbLf, "")
    strCounts = Replace(strCounts, vbCr, " + ")
    
    strRoute = UCase(tbRoute.Value)
    iNumber = CInt(tbNumber.Value)
    iIncrement = CInt(tbIncrement.Value)
    
    Me.Hide
    
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "From Pole: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objPrevious = objEntity
    Else
        GoTo Exit_Sub
    End If
    
    X(1) = objPrevious.InsertionPoint(0)
    Y(1) = objPrevious.InsertionPoint(1)
    Z(1) = 0#
    
    X(2) = 0#
    Y(2) = 0#
    Z(2) = 0#
    
    dScale = 1#
    iCounter = 0
    
Add_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "To Pole: "
    If Not Err = 0 Then GoTo Stop_Placing
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objPole = objEntity
    Else
        GoTo Stop_Placing
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
    
    Set objSpan = ThisDrawing.ModelSpace.InsertBlock(dTextPnt, "cable_span", dScale, dScale, dScale, dRotate)
    objSpan.Layer = strTextLayer
    
    vAttItem = objSpan.GetAttributes
    
    vAttItem(0).TextString = ""
    vAttItem(1).TextString = strCable
    vAttItem(2).TextString = distance & "'"
    objSpan.Update
    
    vAttPrevious = objPrevious.GetAttributes
    
    If cbGuy.Value = True Or cbExistGuys.Value = True Then
        If iCounter = 0 Then
            strTemp = PlaceGuy((X(1)), (Y(1)), (X(0)), (Y(0)), (X(0)), (Y(0)))
            'vAttPrevious(0).TextString = strRoute & iNumber
            'strTemp = "+PM2A=1;;"
        Else
            strTemp = PlaceGuy((X(1)), (Y(1)), (X(0)), (Y(0)), (X(2)), (Y(2)))
            'strTemp = ""
        End If
        
        If cbExistGuys.Value = True Then strTemp = ""
            
        If Not strTemp = "" Then
            If vAttPrevious(27).TextString = "" Then
                vAttPrevious(27).TextString = strTemp
            Else
                vAttPrevious(27).TextString = vAttPrevious(27).TextString & ";;" & strTemp
            End If
        End If
    End If
    objPrevious.Update
    'iCounter = 1
    
    strLine = ""
    
    strResult = ThisDrawing.Utility.GetString(0, "Need a Closure [y/n]:(n)")
    If Err = 0 Then
        If LCase(Left(strResult, 1)) = "y" Then
            dInfo(0) = distance
            dInfo(1) = dScale
            dInfo(2) = dRotate
            dInfo(3) = X(0)
            dInfo(4) = Y(0)
            dInfo(5) = X(1)
            dInfo(6) = Y(1)
            
            vInfo = dInfo
            Call PlaceClosure(vInfo)
            
            strLine = "+" & cbType.Value & "(" & cbSize.Value & ")" & cbSuffix.Value & "=" & tbCoil.Value & "  LOOP;;"
            Select Case cbType.Value
                Case "CO", "ADSS"
                    strLine = strLine & "+HACO("
                Case Else
                    strLine = strLine & "+HBFO("
            End Select
            strLine = strLine & cbSize.Value & ")=1  " & cbClosure.Value
        End If
    Else
        Err = 0
    End If
    
    If strLine = "" Then
        strLine = "+" & cbType.Value & "(" & cbSize.Value & ")" & cbSuffix.Value & "=" & distance & ";;"
    Else
        strLine = "+" & cbType.Value & "(" & cbSize.Value & ")" & cbSuffix.Value & "=" & distance & ";;" & strLine
    End If
    
    strNote = strRoute & " - " & (iNumber) & ";;" & FixStructureNumber(vAttPrevious(0).TextString) & ";;" & distance & ";;"
    strNote = strNote & tbTrunk.Value & ":0:0;; ;; ;; ;; ;;" & strCable & " / " & strCounts
       
    vAttList = objPole.GetAttributes
    
    If cbType.Value = "CO" Then
        Select Case Left(UCase(vAttList(8).TextString), 1)
            Case "M", "m"
                strLine = strLine & ";;+PM2A=1"
        End Select
    End If
    
    vAttList(0).TextString = strRoute & iNumber
    
    If vAttList(27).TextString = "" Then
        vAttList(27).TextString = strLine
    Else
        vAttList(27).TextString = vAttList(27).TextString & ";;" & strLine
    End If
    
    vAttList(28).TextString = strNote
    
    objPole.Update
    
    Err = 0
    strLine = ""
    iCounter = 1
    iNumber = iNumber + iIncrement
    
    X(2) = X(1)
    Y(2) = Y(1)
    Z(2) = Z(1)
    
    X(1) = X(0)
    Y(1) = Y(0)
    Z(1) = Z(0)
    
    Set objPrevious = objPole
    
    GoTo Add_Another    '<--------------------------------------------------------------Get Another Pole
    
Stop_Placing:
    If Not iCounter = 0 Then
        If cbGuy.Value = True Or cbExistGuys.Value = True Then
            strTemp = PlaceGuy((X(1)), (Y(1)), (X(2)), (Y(2)), (X(2)), (Y(2)))
        
            strLine = ";;+PE1-3G=1;;+PF1-5A=1;;+PM11=1"
        End If
        
        'strNote = strRoute & " - " & (iNumber + iIncrement) & ";;" & strRoute & " - " & iNumber & ";;" & distance & ";;"
        'strNote = strNote & " ;; ;; ;; ;; ;;" & strCable
        
        iNumber = iNumber + iIncrement
        
        vAttList = objPole.GetAttributes
        'vAttList(0).TextString = strRoute & iNumber
        
        If vAttList(27).TextString = "" Then
            vAttList(27).TextString = strLine
        Else
            vAttList(27).TextString = vAttList(27).TextString & ";;" & strLine
        End If
        
        'vAttList(28).TextString = strNote
        
        objPole.Update
    End If
    
Exit_Sub:
    
    tbNumber.Value = iNumber
    
    Me.show
End Sub

Private Sub cbExistGuys_Click()
    If cbExistGuys.Value = True Then cbGuy.Value = False
End Sub

Private Sub cbGuy_Click()
    If cbGuy.Value = True Then cbExistGuys.Value = False
End Sub

Private Sub cbPlace_Click()
    Call PlaceCable
    
    'Save data to pole attribute (28)
    'Save data to other attribute (8) when ready and created
    
End Sub

Private Sub cbSize_Change()
    Dim strPath As String
    
    strPath = LCase(ThisDrawing.Path)
    
    If InStr(strPath, "united") > 0 Then
        Select Case CInt(cbSize.Value)
            Case Is < 49
                cbClosure.Value = "FOSC A"
            Case Is > 96
                cbClosure.Value = "FOSC D"
            Case Else
                cbClosure.Value = "FOSC B"
        End Select
    Else
        Select Case CInt(cbSize.Value)
            Case Is < 25
                cbClosure.Value = "G4"
            Case Is > 144
                cbClosure.Value = "G6"
            Case Else
                cbClosure.Value = "G5"
        End Select
    End If
    
    Call UpdateCounts
End Sub

Private Sub cbSuffix_Change()
    If cbType.Value = "CO" Then
        If cbSuffix.Value = "E" Then
            cbGuy.Value = False
            cbExistGuys.Value = True
        Else
            cbGuy.Value = True
            cbExistGuys.Value = False
        End If
    Else
        cbGuy.Value = False
    End If
End Sub

Private Sub cbSuffix_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    cbSuffix.Value = UCase(cbSuffix.Value)
End Sub

Private Sub cbType_Change()
    cbSuffix.Clear
    cbExistGuys.Value = False
    
    Select Case cbType.Value
        Case "CO"
            cbSuffix.AddItem "E"
            cbSuffix.AddItem "6M"
            cbSuffix.AddItem "6M-EHS"
            cbSuffix.AddItem "10M"
            cbSuffix.AddItem "16M"
            cbSuffix.Value = "6M-EHS"
            
            'cbClosure.Value = "Aerial"
            cbExistGuys.Value = False
            cbGuy.Value = True
            tbCoil.Value = "100"
        Case "ADSS"
            'cbClosure.Value = "Aerial"
            cbGuy.Value = False
            tbCoil.Value = "50"
        Case "BFO"
            'cbClosure.Value = "Buried"
            cbGuy.Value = False
            tbCoil.Value = "20"
        Case Else
            'cbClosure.Value = "Buried"
            cbGuy.Value = False
            tbCoil.Value = "50"
    End Select
End Sub

Private Sub Label49_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttLst, vReturnPnt As Variant
    
    On Error Resume Next
    
    Me.Hide
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Structure:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    Select Case objBlock.Name
        Case "sPole", "sPed", "sHH", "sPanel"
        Case Else
            GoTo Exit_Sub
    End Select
    
    vAttList = objBlock.GetAttributes
    
    tbRoute.Value = vAttList(0).TextString
Exit_Sub:
    Me.show
End Sub

Private Sub tbF1_Change()
    If cbSize.Value = "" Then Exit Sub
    
    Call UpdateCounts
End Sub

Private Sub tbTrunk_Change()
    If cbSize.Value = "" Then Exit Sub
    
    Call UpdateCounts
End Sub

Private Sub UserForm_Initialize()
    cbType.AddItem "ADSS"
    cbType.AddItem "BFO"
    cbType.AddItem "CO"
    cbType.AddItem "UO"
    cbType.Value = "CO"
    
    cbSize.AddItem "12"
    cbSize.AddItem "24"
    cbSize.AddItem "36"
    cbSize.AddItem "48"
    cbSize.AddItem "72"
    cbSize.AddItem "96"
    cbSize.AddItem "144"
    cbSize.AddItem "216"
    cbSize.AddItem "288"
    cbSize.AddItem "360"
    cbSize.AddItem "432"
    cbSize.AddItem "576"
    cbSize.AddItem "864"
    cbSize.Value = "24"
    
    cbStatus.AddItem "Existing"
    cbStatus.AddItem "Proposed"
    cbStatus.AddItem "Future"
    cbStatus.AddItem "Remove"
    cbStatus.Value = "Proposed"
    
    cbClosure.AddItem ""
    cbClosure.AddItem "#1413"
    cbClosure.AddItem "FOSC A"
    cbClosure.AddItem "FOSC B"
    cbClosure.AddItem "FOSC D"
    cbClosure.AddItem "G4"
    cbClosure.AddItem "G5"
    cbClosure.AddItem "G6"
    cbClosure.AddItem "G6S"
    
    tbRoute.Value = "QQ/"
    tbRoute.SetFocus
    
    tbCounts.Value = "F2: 1-24"
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
    dScale = 1#
        
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dInsert, strGuy, dScale, dScale, dScale, dRotate)
    If cbExistGuys.Value = True Then
        objBlock.Layer = "Integrity Existing"
    Else
        objBlock.Layer = "Integrity " & cbStatus.Value
    End If
    vAttList = objBlock.GetAttributes
    vAttList(2).TextString = "PE1-3G"
    objBlock.Update

    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dInsert, strAnc, dScale, dScale, dScale, dRotate)
    If cbExistGuys.Value = True Then
        objBlock.Layer = "Integrity Existing"
    Else
        objBlock.Layer = "Integrity " & cbStatus.Value
    End If
    vAttList = objBlock.GetAttributes
    vAttList(0).TextString = "PF1-5A"
    If strPull = "DE" Then
        vAttList(1).TextString = "P= DE"
    Else
        vAttList(1).TextString = "P= " & strPull
    End If
    objBlock.Update
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

    strLayer = "Integrity Map Splices-Aerial"

    strBlockName = "Map splice"
    Set obrMClosure = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, strBlockName, dScale, dScale, dScale, dRotate)
    obrMClosure.Layer = strLayer
    vAttList = obrMClosure.GetAttributes

    If Not cbClosure.Value = "" Then vAttList(0).TextString = cbClosure.Value
    'vAttList(1).TextString = str
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

    strLayer = "Integrity Map Coils-Aerial"   'Replace(strLayer, "Splices", "Coils")

    Set obrMCoil = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "Map coil", dScale, dScale, dScale, 0#)
    obrMCoil.Layer = strLayer
    vAttList = obrMCoil.GetAttributes
    If Not tbCoil.Value = "" Then vAttList(0).TextString = tbCoil.Value & "'"
    vAttList(1).TextString = tbCoilCable.Value & "F"
    obrMCoil.Update

Exit_Sub:
    
End Sub

Private Sub UpdateCounts()
    Dim iSize, iCount As Integer
    Dim iTrunk, iF1, iF2 As Integer
    Dim strCount As String
    
    iCount = 0
    iSize = CInt(cbSize.Value)
    If tbTrunk.Value = "" Then
        iTrunk = 0
    Else
        iTrunk = CInt(tbTrunk.Value)
    End If
    If tbF1.Value = "" Then
        iF1 = 0
    Else
        iF1 = CInt(tbF1.Value)
    End If
    
    strCounts = ""
    If iTrunk > 0 Then
        If Not iTrunk < iSize Then
            strCounts = "TRUNK: 1-" & iSize
            GoTo Full
        Else
            strCounts = "TRUNK: 1-" & iTrunk & vbCr
            iCount = iTrunk
        End If
    End If
    If iF1 > 0 Then
        If Not iCount + iF1 < iSize Then
            strCounts = strCounts & "F1: " & (iCount + 1) & "-" & iSize
            GoTo Full
        Else
            strCounts = strCounts & "F1: " & (iCount + 1) & "-" & (iCount + iF1) & vbCr
            iCount = iCount + iF1
        End If
    End If
    strCounts = strCounts & "F2: " & (iCount + 1) & "-" & iSize
    
Full:
    
    tbCounts.Value = strCounts
End Sub

Private Function FixStructureNumber(strLine As String)
    Dim vL, vR As Variant
    Dim strNumber, strResult As String
    
    vL = Split(strLine, "L")
    vR = Split(vL(UBound(vL)), "R")
    strNumber = vR(UBound(vR))
    
    strResult = Left(strLine, Len(strLine) - Len(strNumber)) & " - " & strNumber
    
    FixStructureNumber = strResult
End Function
