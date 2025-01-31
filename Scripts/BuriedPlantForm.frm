VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuriedPlantForm 
   Caption         =   "Buried Plant"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3615
   OleObjectBlob   =   "BuriedPlantForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BuriedPlantForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Pi As Double

Private Sub cbPlaceBM61_Click()
    'Dim lineObj2 As AcadLWPolyline
    Dim objLWP As AcadLWPolyline
    Dim returnPnt1, returnPnt2 As Variant
    Dim lwpCoords(0 To 3) As Double
    Dim objOffset1, objOffset2 As Variant
    
  Me.Hide
    
    returnPnt1 = ThisDrawing.Utility.GetPoint(, "From point:")
    returnPnt2 = ThisDrawing.Utility.GetPoint(, "To point:")
    
    lwpCoords(0) = returnPnt1(0)
    lwpCoords(1) = returnPnt1(1)
    lwpCoords(2) = returnPnt2(0)
    lwpCoords(3) = returnPnt2(1)
    
    Set objLWP = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
    objLWP.Update
    'Set lineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
    'lineObj2.Update
    
    'objOffset1 = lineObj2.Offset(1)
    'objOffset2 = lineObj2.Offset(-1)
  Me.show
End Sub

Private Sub cbPlaceUD_Click()
    Dim insertPnt(0 To 2), leaderPnt(0 To 3) As Double
    Dim newVertex(0 To 1) As Double
    Dim objBlock As AcadBlockReference
    Dim lineObj As AcadLWPolyline
    Dim returnPnt, attList As Variant
    Dim attItem, basePnt As Variant
    Dim str, str1 As String
    Dim strDistance As String
    'Dim strArray() As String
    Dim dScale, dDistance, dRotate As Double
    Dim dOffset As Double
    Dim xDiff, yDiff, zDiff, dDist As Double
    Dim vEOBL, vEOBR As Variant
    Dim vReturnPnt, vEndPnt As Variant
    Dim iCounter As Integer
    
    Me.Hide
  On Error Resume Next
    
    vReturnPnt = ThisDrawing.Utility.GetPoint(, "Start Point: ")
    
    dScale = cbScale.Value / 100
    leaderPnt(0) = vReturnPnt(0)
    leaderPnt(1) = vReturnPnt(1)
    
    Err = 0
    vEndPnt = ThisDrawing.Utility.GetPoint(vReturnPnt, "Next Point: ")
    If Not Err = 0 Then GoTo Exit_Sub
    vReturnPnt = vEndPnt
    
    leaderPnt(2) = vEndPnt(0)
    leaderPnt(3) = vEndPnt(1)
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(leaderPnt)
    Select Case cbUDStatus.Value
        Case "New"
            lineObj.Layer = "Integrity Proposed"
        Case "Existing"
            lineObj.Layer = "Integrity Existing-Buried"
        Case "Future"
            lineObj.Layer = "Integrity Future"
    End Select
    
    lineObj.Linetype = "DASHED"
    lineObj.Update
    
    iCounter = 2
    Err = 0
    
    While Err = 0
        vEndPnt = ThisDrawing.Utility.GetPoint(vReturnPnt, "Next Point: ")
        If Not Err = 0 Then GoTo Exit_While
        vReturnPnt = vEndPnt
        
        newVertex(0) = vEndPnt(0)
        newVertex(1) = vEndPnt(1)
        
        lineObj.AddVertex iCounter, newVertex
        lineObj.Update
        iCounter = iCounter + 1
    Wend
Exit_While:

    Err = 0
    
    dOffset = 2 * dScale
    vEOBL = lineObj.Offset(dOffset)
    vEOBR = lineObj.Offset(-dOffset)
    
    lineObj.Linetype = "HIDDEN"
    'lineObj.Layer = objBlock.Layer
    lineObj.Update
    dDistance = CInt(lineObj.Length)
    
    vReturnPnt = ThisDrawing.Utility.GetPoint(, "Callout Point: ")
    vEndPnt = ThisDrawing.Utility.GetPoint(vReturnPnt, "Direction of Rotation: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    insertPnt(0) = vReturnPnt(0)
    insertPnt(1) = vReturnPnt(1)
    insertPnt(2) = 0#
    
    xDiff = vEndPnt(0) - vReturnPnt(0)
    yDiff = vEndPnt(1) - vReturnPnt(1)
    
    dRotate = Atn(yDiff / xDiff)
    
    Err = 0
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(vReturnPnt, "cable_span", dScale, dScale, dScale, dRotate)
    If Not Err = 0 Then MsgBox "Error: Placing Block"
    attList = objBlock.GetAttributes
    
    strDistance = ThisDrawing.Utility.GetString(0, "Enter Distance(" & dDistance & "'): ")
    
    If strDistance = "" Then
        str1 = cbUDType.Value & "=" & dDistance & "'"
    Else
        str1 = cbUDType.Value & "=" & strDistance
        If Not Right(str1, 1) = "'" Then str1 = str1 & "'"
    End If
    
    attList(2).TextString = str1
    
    objBlock.Layer = lineObj.Layer
    objBlock.Update
    
    If Not cbPlaceCbl Then lineObj.Delete
    lineObj.Update
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbType_Change()
    If cbType.Value = "FP" Then
        tbLetter.Enabled = True
    Else
        tbLetter.Value = ""
        tbLetter.Enabled = False
    End If
End Sub

Private Sub cpPlaceFP_Click()
    Dim objBlock As AcadBlockReference
    Dim returnPnt, attList As Variant
    Dim vLevel1, vLevel2, vLevel3 As Variant
    Dim strBlock, strNumber, strRoute As String
    Dim dScale As Double
    Dim iNumber As Integer
    'Dim objInfo As AcadBlockReference
    'Dim str As String
    'Dim insertPnt(0 To 2), leaderPnt(0 To 3) As Double
    'Dim item, item2 As Variant
    'Dim layerObj As AcadLayer
    'Dim lwpPnt(0 To 3) As Double
    'Dim lineObj As AcadLWPolyline
    
  On Error Resume Next
  
    If cbType.Value = "" Then GoTo Exit_Sub
    If tbPoleNum.Value = "" Then GoTo Exit_Sub
    
    Select Case Left(cbType.Value, 2)
        Case "FP"
            strBlock = "dFP"
        Case "BD"
            strBlock = "PED"
        Case "BU"
            strBlock = "Map splice"
        Case Else
            strBlock = "dHH"
    End Select
    
    Me.Hide
   
    dScale = CDbl(cbScale.Value) / 100
    If Err <> 0 Then dScale = 0.75
    Err = 0
    
    returnPnt = ThisDrawing.Utility.GetPoint(, "Place " & cbType.Value & ": ")
    
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(returnPnt, strBlock, dScale, dScale, dScale, 0#)
    attList = objBlock.GetAttributes
    
    vLevel1 = Split(tbPoleNum.Value, "/")
    vLevel2 = Split(vLevel1(UBound(vLevel1)), "L")
    vLevel3 = Split(vLevel2(UBound(vLevel2)), "R")
    iNumber = CInt(vLevel3(UBound(vLevel3)))
    
    Select Case iNumber
        Case Is < 10
            strRoute = Left(tbPoleNum.Value, Len(tbPoleNum.Value) - 1)
        Case Is < 100
            strRoute = Left(tbPoleNum.Value, Len(tbPoleNum.Value) - 2)
        Case Else
            strRoute = Left(tbPoleNum.Value, Len(tbPoleNum.Value) - 3)
    End Select
    
    If cbType.Value = "FP" Then
        strNumber = iNumber & tbLetter.Value
        tbLetter.Value = Chr(Asc(tbLetter.Value) + 1)
    Else
        strNumber = tbPoleNum.Value
        tbPoleNum.Value = strRoute & (iNumber + 1)
    End If
    
    attList(0).TextString = strNumber
    attList(1).TextString = cbType.Value
    If cbHousingStatus.Value = "Future" Then attList(1).TextString = attList(1).TextString & " (FUTURE)"
    objBlock.Update
    
    Select Case cbHousingStatus.Value
        Case "New"
            objBlock.Layer = "Integrity Proposed"
        Case "Existing"
            objBlock.Layer = "Integrity Existing-Buried"
        Case "Future"
            objBlock.Layer = "Integrity Future"
    End Select
    
    objBlock.Update

Exit_Sub:
    Me.show
End Sub

Private Sub Label1_Click()
    Me.Hide
    Call Get_Pole
    
    If cbType.Value = "FP" Then tbLetter.Value = "A"
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    cbType.AddItem "BHF(30X48X36)T"
    cbType.AddItem "UHF(24X36X24)"
    cbType.AddItem "UHF(30X48X36)"
    cbType.AddItem "UHF(17X30X18)"
    cbType.AddItem "FP"
    cbType.AddItem "BDO3"
    cbType.AddItem "BDO5"
    cbType.AddItem "BDO7"
    cbType.AddItem "BUDI"
    cbType.AddItem "FDH"

    cbUDType.AddItem "BM60(1.5)D"
    cbUDType.AddItem "BM60(1.25)D"
    cbUDType.AddItem "UD(1X1-2)"
    cbUDType.AddItem "UD(1X1-4)"
    cbUDType.AddItem "UD(1X1-1.25)"
    cbUDType.AddItem "UD(1X2-2)"
    cbUDType.AddItem "UD(1X2-4)"
    'Call Get_Pole
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    cbHousingStatus.AddItem "Existing"
    cbHousingStatus.AddItem "New"
    cbHousingStatus.AddItem "Future"
    cbHousingStatus.Value = "New"
    
    cbUDStatus.AddItem "Existing"
    cbUDStatus.AddItem "New"
    cbUDStatus.AddItem "Future"
    cbUDStatus.Value = "New"
    
    Pi = 3.14159265359
End Sub

Private Sub Get_Pole()
    Dim entPole As AcadObject
    Dim obrGP As AcadBlockReference
    Dim attItem, basePnt As Variant

  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Block: "
    If TypeOf entPole Is AcadBlockReference Then
        Set obrGP = entPole
    Else
        MsgBox "Not a valid block."
        Exit Sub
    End If
    
    attItem = obrGP.GetAttributes
    tbPoleNum.Value = attItem(0).TextString
End Sub
