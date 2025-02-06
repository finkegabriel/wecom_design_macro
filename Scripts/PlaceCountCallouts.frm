VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlaceCountCallouts 
   Caption         =   "Place Counts Callouts"
   ClientHeight    =   7170
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   8040
   OleObjectBlob   =   "PlaceCountCallouts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlaceCountCallouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objStructure As AcadBlockReference
Dim strMainCC, strMainSC As String

Private Sub cbBreakLines_Click()
    Me.Hide
    
    ThisDrawing.SendCommand "_break" & vbCr
    Me.show
End Sub

Private Sub cbConsolidateCable_Click()
    If tbCableCounts.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vCounts As Variant
    Dim strLine As String
    Dim strPrevious, strCurrent, strRef As String
    Dim iCurrent As Integer
    Dim iStart, iEnd As Integer
    Dim iTempS, iTempE As Integer
    
    iCurrent = 0
    strRef = ""
    
    strLine = Replace(tbCableCounts.Value, vbLf, "")
    strLine = Replace(strLine, vbTab, " ")
    vLine = Split(strLine, vbCr)
    
    vItem = Split(vLine(0), ": ")
    If UBound(vItem) > 1 Then
        vCounts = Split(vItem(1), "-")
        iStart = CInt(vCounts(0))
        If UBound(vCounts) = 0 Then
            iEnd = iStart
        Else
            iEnd = CInt(vCounts(1))
        End If
        strPrevious = vItem(0)
        strRef = vItem(2)
    'Else
        '
    End If
    
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vLine)
            If vLine(i) = "" Then GoTo Next_line
            
            vItem = Split(vLine(i), ": ")
            vCounts = Split(vItem(1), "-")
            iTempS = CInt(vCounts(0))
            If UBound(vCounts) = 0 Then
                iTempE = iTempS
            Else
                iTempE = CInt(vCounts(1))
            End If
            strCurrent = vItem(0)
            
            If strCurrent = strPrevious Then
                If iTempS = iEnd + 1 Then
                    iEnd = iTempE
                    vLine(iCurrent) = strPrevious & ": " & iStart & "-" & iEnd & ": " & strRef
                    vLine(i) = ""
                Else
                    iStart = iTempS
                    iEnd = iTempE
                    strRef = vItem(2)
                End If
            Else
                iCurrent = i
                iStart = iTempS
                iEnd = iTempE
                strPrevious = strCurrent
                strRef = vItem(2)
            End If
Next_line:
        Next i
    End If
    
    strLine = ""
    For i = 0 To UBound(vLine)
        If Not vLine(i) = "" Then
            If strLine = "" Then
                strLine = vLine(i)
            Else
                strLine = strLine & vbCr & vLine(i)
            End If
        End If
    Next i
    
    strLine = Replace(strLine, " ", vbTab)
    tbCableCounts.Value = strLine
End Sub

Private Sub cbGet_Click()
    Dim objEntity As AcadEntity
    Dim vAttList, vReturnPnt As Variant
    Dim vLine, vItem, vCounts As Variant
    Dim vCable, vSplice As Variant
    Dim strText, strBack As String
    Dim strCables, strSplices, strTemp As String
    Dim result, iAtt As Integer
    Dim bTest As Boolean
    
    bTest = False
    
    If Not tbCableCounts.Value = strMainCC Then bTest = True
    If Not tbClosure.Value = strMainSC Then bTest = True
    If strMainCC = "empty" Then
        If strMainSC = "empty" Then bTest = False
    End If
    If bTest = True Then
        result = MsgBox("Save count changes to Block?", vbYesNo, "Save Changes")
        If result = vbYes Then Call UpdateBlock(CStr(tbPosition.Value))
    End If
    
    Me.Hide
    
    On Error Resume Next
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Structure:"
    If Not Err = 0 Then
        MsgBox Err.Description
        GoTo Exit_Sub
    End If
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox objEntity.ObjectName
        GoTo Exit_Sub
    End If
    
    Set objStructure = objEntity
    
    Select Case objStructure.Name
        Case "sPole"
            iAtt = 25
            tbType.Value = "POLE"
        Case "sPed"
            iAtt = 5
            tbType.Value = "PED"
        Case "sHH"
            iAtt = 5
            tbType.Value = "HH"
        Case "sMH"
            iAtt = 5
            tbType.Value = "MH"
        Case "sPanel"
            iAtt = 5
            tbType.Value = "PANEL"
        Case "sFP"
            iAtt = 5
            tbType.Value = "FP"
        Case Else
            GoTo Exit_Sub
    End Select
    
    cbScale.Value = GetScale(objStructure.InsertionPoint) * 100
    
    vAttList = objStructure.GetAttributes
    
    'MsgBox vAttList(iAtt).TextString
    strCables = Replace(vAttList(iAtt).TextString, vbLf, "")
    strCables = Replace(strCables, "\P", vbCr)
    vCable = Split(strCables, vbCr)
    
    cbPosition.Clear
    For i = 0 To UBound(vCable)
        vCounts = Split(vCable(i), ": ")
        cbPosition.AddItem vCounts(0)
    Next i
    cbPosition.Value = cbPosition.List(0)
        
    Call GetPositionData(cbPosition.Value)
    
    If tbNumber.Value = "" Then tbNumber.Value = vAttList(0).TextString
        
Exit_Sub:
    
    Me.show
End Sub

Private Sub cbGetCCallout_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    
    Dim vLine As Variant
    Dim strLine, strSource As String
    Dim strTemp As String
    
    Me.Hide
    
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Callout:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    
    Select Case objBlock.Name
        Case "Callout"
            vAttList = objBlock.GetAttributes
            
            strLine = UCase(vAttList(2).TextString)
            If strLine = "" Then GoTo Exit_Sub
            
            tbCableType.Value = vAttList(1).TextString
            vLine = Split(vAttList(0).TextString, ": ")
            tbPosition.Value = vLine(1)
        Case "CableCounts"
            vAttList = objBlock.GetAttributes
            
            strLine = UCase(vAttList(0).TextString)
            If strLine = "" Then GoTo Exit_Sub
            
            tbCableType.Value = vAttList(1).TextString
            tbPosition.Value = "A1"
        Case Else
            GoTo Exit_Sub
    End Select
    
    strLine = Replace(strLine, "\P", vbCr)
    strLine = Replace(strLine, ") ", ")" & vbCr)
    strLine = Replace(strLine, ": ", ":" & vbTab)
    
    'GoTo Skip_This
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Source Block:"
    If Err = 0 Then
        If Not TypeOf objEntity Is AcadBlockReference Then
            strSource = "UNKNOWN"
            GoTo No_Error
        End If
        
        Set objBlock = objEntity
        
        Select Case objBlock.Name
            Case "sPole", "sPed", "sHH", "sPanel", "sMH"
                vAttList = objBlock.GetAttributes
                strSource = vAttList(0).TextString
                GoTo No_Error
            Case Else
                strSource = "UNKNOWN"
                GoTo No_Error
        End Select
    End If
    
    Err = 0
    strSource = "UNKNOWN"
    
No_Error:
    
    vLine = Split(strLine, vbCr)
    For i = 0 To UBound(vLine)
        strTemp = Replace(vLine(i), ")", "")
        strTemp = strTemp & ":" & vbTab & strSource
        
        If Left(strTemp, 1) = "(" Then strTemp = strTemp & ")"
        
        vLine(i) = strTemp
    Next i
    
    strLine = vLine(0)
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vLine)
            strLine = strLine & vbCr & vLine(i)
        Next i
    End If
    
Skip_This:
    
    tbCableCounts.Value = strLine
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbPan_Click()
    Dim objEntity As AcadEntity
    Dim vReturn As Variant
    
    On Error Resume Next
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturn, vbCr & "Right Click to Exit:"
    
    Me.show
End Sub

Private Sub cbPlaceCblCallout_Click()
    If tbCableCounts.Value = "" Then Exit Sub
    
    Dim objCallout As AcadBlockReference
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
    'Dim str1, str2, str3, strLetter As String
    'Dim lowNum, highNum, strArray() As String
    'Dim mainStatus, tempStatus As String
    Dim vLine, vItem, vTemp As Variant
    
    Dim strAtt0, strAtt1, strAtt2 As String
    Dim strLayer As String
    
    'vLine = Split(lbCables.List(iIndex, 0), ": ")
    strAtt0 = tbNumber.Value & ": " & tbPosition.Value
    strAtt1 = tbCableType.Value
    strAtt2 = Replace(tbCableCounts.Value, vbCr, "\P")
    strAtt2 = Replace(strAtt2, vbLf, "")
    strAtt2 = Replace(strAtt2, vbTab, " ")
    
    vLine = Split(strAtt2, "\P")
    
    For i = 0 To UBound(vLine)
        If vLine(i) = "" Then GoTo Next_line
        
        vItem = Split(vLine(i), ": ")
        vTemp = Split(vItem(1), "-")
        If UBound(vTemp) > 0 Then
            If vTemp(0) = vTemp(1) Then vItem(1) = vTemp(0)
        End If
        vLine(i) = vItem(0) & ": " & vItem(1)
        If Left(vLine(i), 1) = "(" Then vLine(i) = vLine(i) & ")"
Next_line:
    Next i
    
    Dim strTemp As String
    
    strTemp = vLine(0)
    For i = 1 To UBound(vLine)
        strTemp = strTemp & vbCr & vLine(i)
    Next i
    
    'MsgBox strTemp
    
    vLine = Split(strTemp, vbCr)
    strAtt2 = vLine(0)
    If Right(strAtt2, 1) = ")" Then strAtt2 = strAtt2 & " "
    
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vLine)
            
            If Right(strAtt2, 2) = ") " Then
                strAtt2 = strAtt2 & vLine(i)
            Else
                strAtt2 = strAtt2 & "\P" & vLine(i)
            End If
            
            If Right(strAtt2, 1) = ")" Then strAtt2 = strAtt2 & " "
        Next i
    End If
    
    If Left(tbCableType.Value, 2) = "CO" Then
        strLayer = "Integrity Proposed-Aerial"
    Else
        strLayer = "Integrity Proposed-Buried"
    End If
    
  'On Error Resume Next
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
    
    dScale = CDbl(cbScale.Value) / 100
    
    If cbPlacement.Value = "Leader" Then
        Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
        lineObj.Layer = strLayer
        lineObj.Update
    Else
        Set objCircle = ThisDrawing.ModelSpace.AddCircle(dPrevious, 8 * dScale)
        objCircle.Layer = strLayer
        objCircle.Update
        Set objCircle2 = ThisDrawing.ModelSpace.AddCircle(returnPnt, 8 * dScale)
        objCircle2.Layer = strLayer
        objCircle2.Update
        
        strLetter = UCase(ThisDrawing.Utility.GetString(0, "Enter Callout Letter:"))
        Set objText = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8 * dScale)
        Set objText2 = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8 * dScale)
        objText.Layer = strLayer
        objText.Alignment = acAlignmentMiddle
        objText.TextAlignmentPoint = dPrevious
        objText2.Layer = strLayer
        objText2.Alignment = acAlignmentMiddle
        objText2.TextAlignmentPoint = vBlockCoords
        objText.Update
        objText2.Update
        
        vBlockCoords(0) = vBlockCoords(0) + 8 * dScale
        lwpCoords(0) = 0
    End If
    
    'dScale = 0.75
        
    Set objCallout = ThisDrawing.ModelSpace.InsertBlock(vBlockCoords, "Callout", dScale, dScale, dScale, 0)
    attItem = objCallout.GetAttributes
    
    attItem(0).TextString = strAtt0
    attItem(1).TextString = strAtt1
    attItem(2).TextString = strAtt2
    
    objCallout.Layer = strLayer
    
    If lwpCoords(2) < lwpCoords(0) Then
        vBlockCoords(0) = vBlockCoords(0) - (75 * dScale)
        objCallout.InsertionPoint = vBlockCoords
    End If
    objCallout.Update
    
    Me.show
End Sub

Private Sub cbPlaceClosureCallout_Click()
    If tbClosure.Value = "" Then Exit Sub
    
    Dim objCO As AcadBlockReference
    Dim returnPnt, vAttList As Variant
    Dim vBlockCoords(0 To 2) As Double
    Dim lwpCoords(0 To 3) As Double
    Dim dPrevious(0 To 2) As Double
    Dim dOrigin(0 To 2) As Double
    Dim dScale As Double
    Dim lineObj As AcadLWPolyline
    Dim vLine, vCounts, vTemp, vItem As Variant
    Dim strCounts, strHO1 As String
    Dim strSegment As String
    Dim iHO1 As Integer
    
  'On Error Resume Next
    Select Case Left(tbCableType.Value, 2)
        Case "CO"
            strHO1 = "+HO1A="
        Case Else
            strHO1 = "+HO1B="
    End Select
    
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
    Me.Hide
    returnPnt = objStructure.InsertionPoint
    
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
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
    lineObj.Layer = "Integrity Proposed"
    lineObj.Update
    
    'dScale = 0.75
    dScale = CDbl(cbScale.Value) / 100
    
    strSegment = tbNumber.Value & ": " & tbPosition.Value
    
    strCounts = Replace(tbClosure.Value, vbLf, "")
    strCounts = Replace(strCounts, vbTab, " ")
    vLine = Split(strCounts, vbCr)
    vTemp = Split(vLine(0), ": ")
    vItem = Split(vTemp(1), "-")
    If UBound(vItem) > 0 Then
        If vItem(0) = vItem(1) Then vTemp(1) = vItem(0)
    End If
    strCounts = vTemp(0) & ": " & vTemp(1)
    
    If InStr(vTemp(1), " FUTURE") < 1 Then
        vCounts = Split(vTemp(1), "-")
    
        If UBound(vCounts) = 0 Then
            iHO1 = 1
        Else
            iHO1 = CInt(vCounts(1)) - CInt(vCounts(0)) + 1
        End If
    End If
    
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vLine)
            If vLine(i) = "" Then GoTo Next_line
            
            vTemp = Split(vLine(i), ": ")
            vItem = Split(vTemp(1), "-")
            If UBound(vItem) > 0 Then
                If vItem(0) = vItem(1) Then vTemp(1) = vItem(0)
            End If
            strCounts = strCounts & "\P" & vTemp(0) & ": " & vTemp(1)
    
            If InStr(vTemp(1), " FUTURE") < 1 Then
                vCounts = Split(vTemp(1), "-")
    
                If UBound(vCounts) = 0 Then
                    iHO1 = iHO1 + 1
                Else
                    iHO1 = iHO1 + CInt(vCounts(1)) - CInt(vCounts(0)) + 1
                End If
            End If
            
Next_line:
        Next i
    End If
    
    strHO1 = strHO1 & iHO1
        
    Set objCO = ThisDrawing.ModelSpace.InsertBlock(vBlockCoords, "Callout", dScale, dScale, dScale, 0)
    
    objCO.Layer = "Integrity Proposed"
    vAttItem = objCO.GetAttributes
    vAttItem(0).TextString = strSegment
    vAttItem(1).TextString = strHO1
    vAttItem(2).TextString = strCounts
    
    If lwpCoords(2) < lwpCoords(0) Then
        vBlockCoords(0) = vBlockCoords(0) - (75 * dScale)
        objCO.InsertionPoint = vBlockCoords
    End If
    objCO.Update
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub SortWL()
    Dim strText As String
    Dim vLine, vB, vB1 As Variant
    Dim iCount As Integer
    Dim strAtt As String
    
    strText = Replace(tbWL.Value, vbLf, "")
    vLine = Split(strText, vbCr)
    iCount = UBound(vLine)
    
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            'vB = Split(vLine(b), " ")
            'vB1 = Split(vLine(b + 1), " ")
            vB = Split(vLine(b), vbTab)
            vB1 = Split(vLine(b + 1), vbTab)
            
                'MsgBox vB(1) & vbCr & vB1(1)
            
            If vB(0) > vB1(0) Then
            'If CInt(vB(0)) > CInt(vB1(0)) Then
                strAtt = vLine(b + 1)
                
                vLine(b + 1) = vLine(b)
                
                vLine(b) = strAtt
            End If
        Next b
    Next a
    
    strText = vLine(0)
    
    For i = 1 To UBound(vLine)
        If Not vLine(i) = vLine(i - 1) Then
            strText = strText & vbCr & vLine(i)
        End If
    Next i
    
    tbWL.Value = strText
End Sub

Private Sub cbSort_Click()
    Dim vLine, vItem, vCounts As Variant
    Dim vTemp, vItem1, vCounts1 As Variant
    Dim strLine As String
    Dim iFiber, iFiber1 As Integer
    Dim iCount As Integer
    
    strLine = Replace(tbClosure.Value, vbLf, "")
    strLine = Replace(strLine, vbTab, " ")
    vLine = Split(strLine, vbCr)
    
    iCount = UBound(vLine)
    If iCount = 0 Then Exit Sub
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            vItem = Split(vLine(b), ": ")
            vItem1 = Split(vLine(b + 1), ": ")
            If vItem(0) > vItem1(0) Then
                strLine = vLine(b + 1)
                vLine(b + 1) = vLine(b)
                vLine(b) = strLine
            ElseIf vItem(0) = vItem1(0) Then
                vCounts = Split(vItem(1), "-")
                vTemp = Split(vCounts(1), " ")
                iFiber = CInt(vTemp(0))
                vCounts1 = Split(vItem1(1), "-")
                vTemp = Split(vCounts1(1), " ")
                iFiber1 = CInt(vTemp(0))
                
                If iFiber > iFiber1 Then
                    strLine = vLine(b + 1)
                    vLine(b + 1) = vLine(b)
                    vLine(b) = strLine
                End If
            End If
        Next b
    Next a
    
    strLine = vLine(0)
    
    For i = 1 To iCount
        strLine = strLine & vbCr & vLine(i)
    Next i
    
    strLine = Replace(strLine, " ", vbTab)
    tbClosure.Value = strLine
End Sub

Private Sub cbSwitchPositions_Click()
    If cbPosition.Value = tbPosition.Value Then Exit Sub
    
    'tbCableCounts.Value = ""
    'tbClosure.Value = ""
    
    Call GetPositionData(cbPosition.Value)
End Sub

Private Sub cbTest_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    'Dim vLine As Variant
    
    Me.Hide
    
    On Error Resume Next
    
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
    
    vAttList(0).Rotation = 0#
    objBlock.Update
Exit_Sub:
    Me.show
End Sub

Private Sub cbUpdate_Click()
    Call UpdateBlock(CStr(tbPosition.Value))
    
    For i = 0 To cbPosition.ListCount - 1
        If cbPosition.Value = tbPosition.Value Then GoTo Found_Position
    Next i
    
    cbPosition.AddItem tbPosition.Value
    cbPosition.Value = tbPosition.Value
    
Found_Position:
    
End Sub

Private Sub cbUpdateCblCallout_Click()
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim strSearch As String
    
    strSearch = tbNumber.Value & ": " & tbPosition.Value

    grpCode(0) = 2
    grpValue(0) = "Callout"
    filterType = grpCode
    filterValue = grpValue
    
  On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        If vAttList(0).TextString = strSearch Then
            If InStr(vAttList(1).TextString, "+HO1") < 1 Then GoTo Found_Callout
        End If
    Next objBlock
    
    MsgBox "No Callout found for structure."
    GoTo Exit_Sub
    
Found_Callout:
    
    Dim strAtt0, strAtt1, strAtt2 As String
    Dim vLine As Variant
    
    'strAtt0 = tbNumber.Value & ": " & tbPosition.Value
    strAtt1 = tbCableType.Value
    strAtt2 = Replace(tbCableCounts.Value, vbCr, "\P")
    strAtt2 = Replace(strAtt2, vbLf, "")
    strAtt2 = Replace(strAtt2, vbTab, " ")
    
    vLine = Split(strAtt2, "\P")
    
    For i = 0 To UBound(vLine)
        If vLine(i) = "" Then GoTo Next_line
        
        vItem = Split(vLine(i), ": ")
        vLine(i) = vItem(0) & ": " & vItem(1)
        If Left(vLine(i), 1) = "(" Then vLine(i) = vLine(i) & ")"
Next_line:
    Next i
    
    Dim strTemp As String
    
    strTemp = vLine(0)
    For i = 1 To UBound(vLine)
        strTemp = strTemp & vbCr & vLine(i)
    Next i
    
    'MsgBox strTemp
    
    vLine = Split(strTemp, vbCr)
    strAtt2 = vLine(0)
    If Right(strAtt2, 1) = ")" Then strAtt2 = strAtt2 & " "
    
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vLine)
            
            If Right(strAtt2, 2) = ") " Then
                strAtt2 = strAtt2 & vLine(i)
            Else
                strAtt2 = strAtt2 & "\P" & vLine(i)
            End If
            
            If Right(strAtt2, 1) = ")" Then strAtt2 = strAtt2 & " "
        Next i
    End If
    
    vAttList(1).TextString = strAtt1
    vAttList(2).TextString = strAtt2
    objBlock.Update
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub

Private Sub cbUpdateClosureCallout_Click()
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim strSearch As String
    
    strSearch = tbNumber.Value & ": " & tbPosition.Value

    grpCode(0) = 2
    grpValue(0) = "Callout"
    filterType = grpCode
    filterValue = grpValue
    
  On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        If vAttList(0).TextString = strSearch Then
            If InStr(vAttList(1).TextString, "+HO1") > 0 Then GoTo Found_Callout
        End If
    Next objBlock
    
    MsgBox "No Callout found for structure."
    GoTo Exit_Sub
    
Found_Callout:
    
    Dim vLine, vTemp As Variant
    Dim strCounts, strHO1 As String
    Dim strSegment As String
    Dim iHO1 As Integer
    
    Select Case Left(tbCableType.Value, 2)
        Case "CO"
            strHO1 = "+HO1A="
        Case Else
            strHO1 = "+HO1B="
    End Select
    
    strCounts = Replace(tbClosure.Value, vbLf, "")
    strCounts = Replace(strCounts, vbTab, " ")
    vLine = Split(strCounts, vbCr)
    vTemp = Split(vLine(0), ": ")
    strCounts = vTemp(0) & ": " & vTemp(1)
    
    If InStr(vTemp(1), " FUTURE") < 1 Then
        vCounts = Split(vTemp(1), "-")
    
        If UBound(vCounts) = 0 Then
            iHO1 = 1
        Else
            iHO1 = CInt(vCounts(1)) - CInt(vCounts(0)) + 1
        End If
    End If
    
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vLine)
            If vLine(i) = "" Then GoTo Next_line
            
            vTemp = Split(vLine(i), ": ")
            strCounts = strCounts & "\P" & vTemp(0) & ": " & vTemp(1)
    
            If InStr(vTemp(1), " FUTURE") < 1 Then
                vCounts = Split(vTemp(1), "-")
    
                If UBound(vCounts) = 0 Then
                    iHO1 = iHO1 + 1
                Else
                    iHO1 = iHO1 + CInt(vCounts(1)) - CInt(vCounts(0)) + 1
                End If
            End If
            
Next_line:
        Next i
    End If
    
    strHO1 = strHO1 & iHO1
    
    vAttList(1).TextString = strHO1
    vAttList(2).TextString = strCounts
    
    objBlock.Update
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub

Private Sub Label23_Click()
    If tbCableCounts.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vCount As Variant
    Dim strLine As String
    Dim iStart, iEnd, iCount As Integer
    
    iCount = 0
    strLine = Replace(tbCableCounts.Value, vbLf, "")
    strLine = Replace(strLine, vbTab, " ")
    
    vLine = Split(strLine, vbCr)
    For i = 0 To UBound(vLine)
        If vLine(i) = "" Then GoTo Next_line
        If Left(vLine(i), 1) = "(" Then GoTo Next_line
        
        vItem = Split(vLine(i), ": ")
        vCounts = Split(vItem(1), "-")
        
        If UBound(vCounts) = 0 Then
            iCount = iCount + 1
        Else
            iStart = CInt(vCounts(0))
            iEnd = CInt(vCounts(1))
            
            If iStart = iEnd Then
                iCount = iCount + 1
            Else
                iCount = iCount + 1 + iEnd - iStart
            End If
        End If
Next_line:
    Next i
    
    tbTotal.Value = iCount
End Sub

Private Sub tbCableCounts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
    
    Load PlaceCableCCounts
    PlaceCableCCounts.show
    Unload PlaceCableCCounts
    
    Me.show
End Sub

Private Sub tbTotal_Enter()
    Dim vLine, vItem, vCount As Variant
    Dim strLine As String
    Dim iStart, iEnd, iCount As Integer
    
    iCount = 0
    strLine = Replace(tbCableCounts.Value, vbLf, "")
    strLine = Replace(strLine, vbTab, " ")
    
    vLine = Split(strLine, vbCr)
    For i = 0 To UBound(vLine)
        If vLine(i) = "" Then GoTo Next_line
        
        vItem = Split(vLine(i), ": ")
        vCounts = Split(vItem(1), "-")
        
        If UBound(vCounts) = 0 Then
            iCount = iCount + 1
        Else
            iStart = CInt(vCounts(0))
            iEnd = CInt(vCounts(1))
            
            If iStart = iEnd Then
                iCount = iCount + 1
            Else
                iCount = iCount + 1 + iEnd - iStart
            End If
        End If
Next_line:
    Next i
    
    tbTotal.Value = iCount
End Sub

Private Sub UserForm_Initialize()
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    cbPlacement.AddItem "Leader"
    cbPlacement.AddItem "Away"
    cbPlacement.Value = "Leader"
    
    strMainCC = "empty"
    strMainSC = "empty"
End Sub

Private Sub UpdateBlock(strPosition As String)
    Dim vAttList As Variant
    Dim vCable, vSplice As Variant
    Dim strCables, strSplices, strTemp As String
    Dim strText, strCable, strClosure As String
    Dim iAtt As Integer
    
    Select Case objStructure.Name
        Case "sPole"
            iAtt = 25
        Case Else
            iAtt = 5
    End Select
    
    If tbCableCounts.Value = "" Then
        strCable = ""
        strText = ""
    Else
        strCable = tbPosition.Value & ": " & tbCableType.Value & " / "
        strText = Replace(tbCableCounts.Value, vbLf, "")
        strText = Replace(strText, vbCr, " + ")
        strText = Replace(strText, vbTab, " ")
    
        If Right(strText, 3) = " + " Then strText = Left(strText, Len(strText) - 3)
        If Left(strText, 3) = " + " Then strText = Right(strText, Len(strText) - 3)
        strCable = strCable & strText
    End If
    
    If tbClosure.Value = "" Then
        strClosure = ""
    Else
        strClosure = Replace(tbClosure.Value, vbLf, "")
        strClosure = Replace(strClosure, vbCr, " + ")
        strClosure = Replace(strClosure, vbTab, " ")
        strClosure = "[" & tbPosition.Value & "] " & strClosure
        If Right(strClosure, 3) = " + " Then strClosure = Left(strClosure, Len(strClosure) - 3)
    End If
    
    
    vAttList = objStructure.GetAttributes
    If vAttList(iAtt).TextString = "" Then
        strCables = strCable
        GoTo Found_Cable
    End If
    
    'MsgBox vAttList(iAtt).TextString
    strCables = Replace(vAttList(iAtt).TextString, vbLf, "")
    If Left(strCables, 1) = vbCr Then strCables = Right(strCables, Len(strCables) - 1)
    strCables = Replace(strCables, "\P", vbCr)
    'MsgBox strCables
    
    If strCables = "" Then
        strCables = strCable
        GoTo Found_Cable
    End If
    
    vCable = Split(strCables, vbCr)
    
    For i = 0 To UBound(vCable)
        vItem = Split(vCable(i), ": ")
        If vItem(0) = strPosition Then
            vCable(i) = strCable
            
            strCables = ""
            For j = 0 To UBound(vCable)
                If Not vCable(j) = "" Then
                    If strCables = "" Then
                        strCables = vCable(j)
                    Else
                        strCables = strCables & vbCr & vCable(j)
                    End If
                End If
            Next j
            GoTo Found_Cable
        End If
    Next i
    
    If strCables = "" Then
        strCables = strCable
    Else
        strCables = strCables & vbCr & strCable
    End If
    
Found_Cable:
    If Not strCable = "" Then vAttList(iAtt).TextString = strCables
    
Cables_Empty:
    
    If strClosure = "" Then GoTo No_Closure
        
    strSplices = Replace(vAttList(iAtt + 1).TextString, vbLf, "")
    strSplices = Replace(strSplices, "\P", vbCr)
    
    'If strSplices = "" Then GoTo Splices_Empty
    vSplice = Split(strSplices, vbCr)
    
    For i = 0 To UBound(vSplice)
        vItem = Split(vSplice(i), "] ")
        strTemp = Replace(vItem(0), "[", "")
        
        If strTemp = strPosition Then
            vSplice(i) = strClosure
            
            strSplices = ""
            For j = 0 To UBound(vSplice)
                If Not vSplice(j) = "" Then
                    If strSplices = "" Then
                        strSplices = vSplice(j)
                    Else
                        strSplices = strSplices & vbCr & vSplice(j)
                    End If
                End If
            Next j
            GoTo Found_Splice
        End If
    Next i
    
    If strSplices = "" Then
        strSplices = strClosure
    Else
        strSplices = strSplices & vbCr & strClosure
    End If
    
    
Found_Splice:
    vAttList(iAtt + 1).TextString = strSplices
    
Splices_Empty:
    
No_Closure:
    objStructure.Update
    
    strMainCC = tbCableCounts.Value
    strMainSC = tbClosure.Value
End Sub

Private Sub GetPositionData(strPosition As String)
    Dim vAttList As Variant
    Dim vLine, vItem, vCounts As Variant
    Dim vCable, vSplice As Variant
    Dim strText, strBack As String
    Dim strCables, strSplices, strTemp As String
    Dim iAtt As Integer
    
    On Error Resume Next
    
    tbCableCounts.Value = ""
    tbClosure.Value = ""
    tbWL.Value = ""
    
    Select Case objStructure.Name
        Case "sPole"
            iAtt = 25
            tbType.Value = "POLE"
        Case "sPed"
            iAtt = 5
            tbType.Value = "PED"
        Case "sHH"
            iAtt = 5
            tbType.Value = "HH"
        Case "sPanel"
            iAtt = 5
            tbType.Value = "PANEL"
        Case "sMH"
            iAtt = 5
            tbType.Value = "MH"
        Case "sFP"
            iAtt = 5
            tbType.Value = "FP"
        Case Else
            MsgBox objStructure.Name & vbCr & "Why?"
            GoTo Exit_Sub
    End Select
    
    vAttList = objStructure.GetAttributes
    
    strCables = Replace(vAttList(iAtt).TextString, vbLf, "")
    strCables = Replace(strCables, "\P", vbCr)
    vCable = Split(strCables, vbCr)
    
    For i = 0 To UBound(vCable)
        vItem = Split(vCable(i), ": ")
        If vItem(0) = strPosition Then GoTo Found_Cable
    Next i
    
    'MsgBox "No Cable found at that Position."
    GoTo Find_Splice
    
Found_Cable:
    vLine = Split(vCable(i), " / ")
    vItem = Split(vLine(0), ": ")
    tbNumber.Value = vAttList(0).TextString
    tbPosition.Value = vItem(0)
    tbCableType.Value = vItem(1)
            
    strText = Replace(vLine(1), " + ", vbCr)
    strText = Replace(strText, " ", vbTab)
    tbCableCounts.Value = strText
    
Find_Splice:
            
    strSplices = Replace(vAttList(iAtt + 1).TextString, vbLf, "")
    
    If Left(strSplices, 1) = vbCrLf Then strSplices = Right(strSplices, Len(strSplices) - 1)
    If Left(strSplices, 1) = vbCr Then strSplices = Right(strSplices, Len(strSplices) - 1)
    If Left(strSplices, 1) = vbLf Then strSplices = Right(strSplices, Len(strSplices) - 1)
    
    If Right(strSplices, 1) = vbCrLf Then strSplices = Left(strSplices, Len(strSplices) - 1)
    If Right(strSplices, 1) = vbCr Then strSplices = Left(strSplices, Len(strSplices) - 1)
    If Right(strSplices, 1) = vbLf Then strSplices = Left(strSplices, Len(strSplices) - 1)
    
    strSplices = Replace(strSplices, "\P", vbCr)
    
    vAttList(iAtt + 1).TextString = strSplices
    objStructure.Update
    
    vSplice = Split(strSplices, vbCr)
    
    For i = 0 To UBound(vSplice)
        If vSplice(i) = "" Then GoTo Next_I
        
        vItem = Split(vSplice(i), "] ")
        strTemp = Replace(vItem(0), "[", "")
        If tbPosition.Value = "" Then tbPosition.Value = strTemp
        If strTemp = strPosition Then GoTo Found_Splice
Next_I:
    Next i
    
    'MsgBox "No Splice found at that Position."
    GoTo Get_Wiring_Limits
    
Found_Splice:
    
    If Not vSplice(0) = "" Then
        For i = 0 To UBound(vSplice)
            vCounts = Split(vSplice(i), "] ")
            strTemp = Replace(vCounts(0), "[", "")
            If strTemp = strPosition Then
                strText = Replace(vCounts(1), " + ", vbCr)
                strText = Replace(strText, " ", vbTab)
                tbClosure.Value = strText
                
                GoTo Get_Wiring_Limits
            End If
        Next i
    End If
    
Get_Wiring_Limits:

    tbWL.Value = ""
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    
    grpCode(0) = 2
    grpValue(0) = "Customer,SG"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        Select Case objBlock.Name
            Case "SG"
                If vAttList(2).TextString = "" Then GoTo Next_objBlock
                vLine = Split(vAttList(2).TextString, " - ")
                strText = vLine(1) & vbTab & "SG - " & vAttList(1).TextString
            Case Else
                If vAttList(4).TextString = "" Then GoTo Next_objBlock
                vLine = Split(vAttList(4).TextString, " - ")
                strText = vLine(1) & vbTab & vAttList(1).TextString & " " & vAttList(2).TextString
        End Select
        
        If vLine(0) = tbNumber.Value Then
            'vItem = Split(vLine(1), ": ")
            'strText = Replace(vItem(1), ")", "")
            
            If tbWL.Value = "" Then
                'tbWL.Value = strText & vbTab & vAttList(1).TextString & " " & vAttList(2).TextString
                tbWL.Value = strText
            Else
                'tbWL.Value = tbWL.Value & vbCr & strText & vbTab & vAttList(1).TextString & " " & vAttList(2).TextString
                tbWL.Value = tbWL.Value & vbCr & strText
            End If
            
        End If
Next_objBlock:
    Next objBlock
    
    Call SortWL
    
Clear_objSS:
    objSS.Clear
    objSS.Delete
        
Exit_Sub:
    cbUpdate.Enabled = True
    
    strMainCC = tbCableCounts.Value
    strMainSC = tbClosure.Value
End Sub

Private Function GetScale(vCoords As Variant)
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim dMinX, dMaxX, dMinY, dMaxY As Double
    Dim dScale As Double
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        dScale = objBlock.XScaleFactor
        dMinX = objBlock.InsertionPoint(0)
        dMaxX = dMinX + (1652 * dScale)
        
        dMinY = objBlock.InsertionPoint(1)
        dMaxY = dMinY + (1052 * dScale)
        
        If vCoords(0) > dMinX And vCoords(0) < dMaxX Then
            If vCoords(1) > dMinY And vCoords(1) < dMaxY Then
                dScale = dScale * 1.333333333
                GoTo Exit_Sub
            End If
        End If
    Next objBlock
    
    dScale = 1#
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    'MsgBox dScale
    GetScale = dScale
End Function
