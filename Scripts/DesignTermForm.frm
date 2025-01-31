VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DesignTermForm 
   Caption         =   "Design Terminals"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   OleObjectBlob   =   "DesignTermForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DesignTermForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim iListIndex As Integer
Dim Pi As Double


Private Sub cbClosureType_Change()
    Select Case cbClosureType.Value
        Case "Aerial"
            tbCoil.Value = "100"
        Case "Buried"
            tbCoil.Value = "20"
    End Select
End Sub

Private Sub cbCoilType_Change()
    If InStr(UCase(ThisDrawing.Path), "LORETTO") > 0 Then
        Select Case Left(cbCoilType.Value, 3)
            Case "12F", "24F"
                cbClosure.Value = "G4"
            Case "36F", "48F", "72F", "96F", "144"
                cbClosure.Value = "G5"
            Case Else
                cbClosure.Value = "G6"
        End Select
    Else
        Select Case Left(cbCoilType.Value, 3)
            Case "12F", "24F", "36F", "48F"
                cbClosure.Value = "FOSC A"
            Case "72F", "96F"
                cbClosure.Value = "FOSC B"
            Case Else
                cbClosure.Value = "FOSC D"
        End Select
    End If
    'If lbWLimits.ListCount > 0 Then cbClosure.Value = "#1413"
End Sub

Private Sub cbConvert_Click()
    Dim obrMClosure As AcadBlockReference
    Dim obrMCoil As AcadBlockReference
    Dim objBlock, obrGP3 As AcadBlockReference
    Dim objEntity As AcadEntity
    Dim vReturnPnt, vEndPnt As Variant
    Dim vAttList As Variant
    Dim dDiff(0 To 3) As Double
    Dim dInsertPnt(0 To 2) As Double
    Dim dDist As Double
    Dim strBlockName, strTermType As String
    Dim layerObj As AcadLayer
    Dim returnPnt, vPoleCoords As Variant
    Dim str, strLayer As String
    Dim vBlockCoords(0 To 2) As Double
    Dim dScale As Double
    Dim i As Integer

    Me.Hide

    strTermType = ""

  On Error Resume Next

    dScale = CDbl(cbScale.Value) / 100

    str = ""
    'For i = 0 To lbWLimits.ListCount - 1
    '    str = str & lbWLimits.List(i) & "::"
    'Next i
    Err = 0

    vReturnPnt = ThisDrawing.Utility.GetPoint(, "Start Pole/Point:")
    vEndPnt = ThisDrawing.Utility.GetPoint(, "Direction of Placement:")
    
    If Not Err = 0 Then GoTo Exit_Sub

    dDiff(0) = vEndPnt(0) - vReturnPnt(0)
    dDiff(1) = vEndPnt(1) - vReturnPnt(1)
    dDiff(2) = Math.Sqr(dDiff(0) * dDiff(0) + (dDiff(1) * dDiff(1)))

    If cbClosureType.Value = "Buried" Then
        dInsertPnt(0) = vReturnPnt(0) + 7.25 * dScale * dDiff(0) / dDiff(2)
        dInsertPnt(1) = vReturnPnt(1) + 7.25 * dScale * dDiff(1) / dDiff(2)
        dInsertPnt(2) = 0#
    Else
        dInsertPnt(0) = vReturnPnt(0) + 15 * dScale * dDiff(0) / dDiff(2)
        dInsertPnt(1) = vReturnPnt(1) + 15 * dScale * dDiff(1) / dDiff(2)
        dInsertPnt(2) = 0#
    End If

    If dDiff(0) = 0 Then
        dRotate = Pi / 2
    ElseIf dDiff(1) = 0 Then
        dRotate = 0#
    Else
        dRotate = Math.Atn(dDiff(1) / dDiff(0))
        If dDiff(0) < 0 Then
            dRotate = Pi + dRotate
        ElseIf dDiff(1) < 0 Then
            dRotate = 2 * Pi + dRotate
        End If
    End If

    If dRotate > Pi Then dRotate = dRotate - Pi
    If dRotate > Pi / 2 Then dRotate = Pi + dRotate

    If Not cbPlaceClosure Then GoTo place_coil

    Select Case cbClosureType.Value
        Case "Aerial"
            strLayer = "Integrity Map Splices-Aerial"
        Case "Buried"
            strLayer = "Integrity Map Splices-Buried"
        Case "UG"
            strLayer = "Integrity Map Splices-UG"
    End Select
    'Set layerObj = ThisDrawing.Layers.Add("Integrity Map Splices-Aerial")
    'ThisDrawing.ActiveLayer = layerObj

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

    If obCTop Then
        If dDiff(0) > 0 Then
            dInsertPnt(0) = dInsertPnt(0) - 9 * dScale * dDiff(1) / dDiff(2)
            dInsertPnt(1) = dInsertPnt(1) + 9 * dScale * dDiff(0) / dDiff(2)
        Else
            dInsertPnt(0) = dInsertPnt(0) + 9 * dScale * dDiff(1) / dDiff(2)
            dInsertPnt(1) = dInsertPnt(1) - 9 * dScale * dDiff(0) / dDiff(2)
        End If
    Else
        If dDiff(0) > 0 Then
            dInsertPnt(0) = dInsertPnt(0) + 9 * dScale * dDiff(1) / dDiff(2)
            dInsertPnt(1) = dInsertPnt(1) - 9 * dScale * dDiff(0) / dDiff(2)
        Else
            dInsertPnt(0) = dInsertPnt(0) - 9 * dScale * dDiff(1) / dDiff(2)
            dInsertPnt(1) = dInsertPnt(1) + 9 * dScale * dDiff(0) / dDiff(2)
        End If
    End If

    If Not cbPlaceCoil Then GoTo Exit_Sub

    'Set layerObj = ThisDrawing.Layers.Add("Integrity Map Coils-Aerial")
    'ThisDrawing.ActiveLayer = layerObj
    
    strLayer = Replace(strLayer, "Splices", "Coils")

    Set obrMCoil = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "Map coil", dScale, dScale, dScale, 0#)
    obrMCoil.Layer = strLayer
    vAttList = obrMCoil.GetAttributes
    If Not tbCoil.Value = "" Then vAttList(0).TextString = tbCoil.Value & "'"
    If Not cbCoilType.Value = "" Then vAttList(1).TextString = cbCoilType.Value
    obrMCoil.Update

Exit_Sub:

    Dim strLine, strTemp As String
    Dim iAtt As Integer
    
    strLine = ""
    
    If cbPlaceCoil.Value = True Then
        Select Case cbClosureType.Value
            Case "UG"
                strLine = "+UO(" & Replace(cbCoilType.Value, "F", "") & ")=" & tbCoil.Value & "'  LOOP"
            Case "Buried"
                strLine = "+BFO(" & Replace(cbCoilType.Value, "F", "") & ")=" & tbCoil.Value & "'  LOOP"
            Case Else
                strLine = "+CO(" & Replace(cbCoilType.Value, "F", "") & ")E=" & tbCoil.Value & "'  LOOP"
        End Select
    End If
    
    If cbPlaceClosure.Value = True Then
        Select Case cbClosureType.Value
            Case "UG"
                strTemp = "+HBFO(" & Replace(cbCoilType.Value, "F", "") & ")=1  " & cbClosure.Value
            Case "Buried"
                strTemp = "+HBFO(" & Replace(cbCoilType.Value, "F", "") & ")=1  " & cbClosure.Value
            Case Else
                strTemp = "+HACO(" & Replace(cbCoilType.Value, "F", "") & ")=1  " & cbClosure.Value
        End Select
        
        If strLine = "" Then
            strLine = strTemp
        Else
            If Not strTemp = "" Then strLine = strLine & ";;" & strTemp
        End If
    End If
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Structure to add Units:"
    If Not Err = 0 Then GoTo Tidy_Up
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox "Not a Block"
        GoTo Tidy_Up
    End If
    
    Set objBlock = objEntity
    
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 27
        Case "sPED", "sHH"
            iAtt = 7
        Case Else
            MsgBox "Not a valid Block"
            GoTo Tidy_Up
    End Select
    
    vAttList = objBlock.GetAttributes
    
    If vAttList(iAtt).TextString = "" Then
        vAttList(iAtt).TextString = strLine
    Else
        vAttList(iAtt).TextString = vAttList(iAtt).TextString & ";;" & strLine
    End If
    
    objBlock.Update
    
Tidy_Up:

    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub Label25_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
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
    
    tbCoil.Value = "100"
    
    Pi = 3.14150265359
    
    'cbUP.Caption = Chr(225)
    'cbDOWN.Caption = Chr(226)
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    'cbScale.Value = MainForm.cbScale.Value
    'If cbScale.Value = "" Then cbScale.Value = "75"
End Sub
