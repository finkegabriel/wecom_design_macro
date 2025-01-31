Attribute VB_Name = "Module1"
Public Sub AddLLtoPoles()
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim i As Integer
    
    Dim dN, dE As Double
    Dim vLL As Variant
    
    i = 0

    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
  On Error Resume Next
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
  
    Err = 0
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    If Not Err = 0 Then
        MsgBox "Error: " & Err.Number & vbCr & Err.Description
        Exit Sub
    End If
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        'If vAttList(0).TextString = "POLE" Then GoTo Next_objBlock
        'If vAttList(0).TextString = "" Then GoTo Next_objBlock
        
        dE = objBlock.InsertionPoint(0)
        dN = objBlock.InsertionPoint(1)
        vLL = TN83FtoLL(CDbl(dN), CDbl(dE))
        
        vAttList(7).TextString = vLL(0) & "," & vLL(1)
        objBlock.Update
        i = i + 1
Next_objBlock:
    Next objBlock
    
    MsgBox "Lat/Long added to " & i & " poles."
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

Public Sub TransferPoleToSheets()
    Dim filterType, filterValue As Variant  '
    Dim grpCode(0) As Integer               '
    Dim grpValue(0) As Variant              '
    Dim objSS7 As AcadSelectionSet          '
    Dim SSobj3 As AcadSelectionSet          '
    Dim objMapPole As AcadBlockReference    '
    Dim objDwgPole As AcadBlockReference    '
    Dim vAttMap, vAttDwg As Variant         '
    Dim vPnt1, vPnt2 As Variant             '
    Dim strPoleNum As String                '
    Dim iTest, iCount  As Integer                   '
    Dim vTemp As Variant
    Dim iHeight As Integer
    Dim strTemp, strPole As String
    Dim strHC As String
    Dim vBasePnt As Variant
    Dim iAttPole, iPower As Integer
    Dim iCATV, iAtt, iTDS, iUTC As Integer
    Dim iXO, iZAYO, iCLEC, iTelco As Integer
    Dim iCity, iTraffic, iOHG As Integer
    Dim strMR As String
    
    On Error GoTo Exit_Sub
    
    iCount = 0
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get DWG BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get DWG UR Corner: ")
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS7 = ThisDrawing.SelectionSets.Add("objSS7")
    objSS7.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get MAP BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get MAP UR Corner: ")
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set SSobj3 = ThisDrawing.SelectionSets.Add("SSobj3")
    SSobj3.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    If objSS7.count < 1 Then GoTo Exit_Sub
    If SSobj3.count < 1 Then GoTo Exit_Sub

    For Each objDwgPole In objSS7
        vAttDwg = objDwgPole.GetAttributes
        strPoleNum = vAttDwg(0).TextString
        
        If strPoleNum = "POLE" Then GoTo Next_objDwgPole
        If strPoleNum = "" Then GoTo Next_objDwgPole
        iTest = 0
        
        For Each objMapPole In SSobj3
            vAttMap = objMapPole.GetAttributes
            
            If strPoleNum = vAttMap(0).TextString Then
                'If vAttDwg(2).TextString = "" Then vAttDwg(2).TextString = vAttMap(2).TextString
                'If vAttDwg(5).TextString = "" Then vAttDwg(5).TextString = vAttMap(5).TextString
                'vAttDwg(6).TextString = vAttMap(6).TextString
                vAttDwg(7).TextString = vAttMap(7).TextString
                
                objDwgPole.Update
                iTest = 1
                iCount = iCount + 1
                GoTo Next_objDwgPole
            End If
        Next objMapPole
Next_objDwgPole:
        If iTest = 0 And Not strPoleNum = "POLE" Then MsgBox strPoleNum & " was not found"
    Next objDwgPole
    
    '<---------------------------------------------------------------------------------------------------------
Exit_Sub:
    If Err <> 0 Then
        MsgBox "Error:" & vbCr & Err.Number & vbCr & Err.Description
    Else
        MsgBox "Done." & vbCr & iCount & " poles updated."
    End If
    
    objSS7.Clear
    objSS7.Delete

    SSobj3.Clear
    SSobj3.Delete
End Sub

Public Sub AddMatchlines()
    Dim objLWP As AcadLWPolyline
    Dim objText As AcadText
    Dim objBlock As AcadBlockReference
    Dim basePnt, returnPnt As Variant
    Dim vLWPO, vCoords, vTemp As Variant
    Dim lwpCoords(0 To 3) As Double
    Dim dTextInsert(0 To 2) As Double
    Dim dMidPnt(0 To 2) As Double
    Dim xDiff, yDiff, zDiff As Double
    Dim dRotate, dScale, dXScale, dTScale As Double
    Dim strLine, strTemp As String
    Dim iStatus As Integer
    
    On Error Resume Next
    
    While Err = 0
        returnPnt = ThisDrawing.Utility.GetPoint(, "Select First Point: ")
        lwpCoords(0) = returnPnt(0)
        lwpCoords(1) = returnPnt(1)
        
        vCoords = returnPnt
        
        If Not Err = 0 Then GoTo Exit_Sub
        
        returnPnt = ThisDrawing.Utility.GetPoint(, "Select Second Point: ")
        lwpCoords(2) = returnPnt(0)
        lwpCoords(3) = returnPnt(1)
        
        dMidPnt(0) = (lwpCoords(0) + lwpCoords(2)) / 2
        dMidPnt(1) = (lwpCoords(1) + lwpCoords(3)) / 2
        dMidPnt(2) = 0#
        
        strLine = UCase(ThisDrawing.Utility.GetString(0, "Matches Drawing #: "))
        
        If Not Err = 0 Then GoTo Exit_Sub
        
        strTemp = GetScale(vCoords)
        vTemp = Split(strTemp, ";;")
        strLayer = vTemp(0)
        dScale = CDbl(vTemp(1))
        dTScale = 8 * (dScale / 1.3333)
        
        If InStr(strLine, "R") > 0 Then
            strLine = Replace(strLine, "R", "")
            iStatus = 1
        Else
            iStatus = 0
        End If
        
        Select Case Len(strLine)
            Case Is = 1
                strLine = "SEE DWG 00" & strLine
            Case Is = 2
                strLine = "SEE DWG 0" & strLine
            Case Is = 3
                strLine = "SEE DWG " & strLine
        End Select
        
        xDiff = lwpCoords(2) - lwpCoords(0)
        yDiff = lwpCoords(3) - lwpCoords(1)
        zDiff = Sqr((xDiff * xDiff) + (yDiff * yDiff))
            
        Select Case xDiff
            Case Is < 0
                If yDiff = 0 Then
                    dRotate = 0#
                ElseIf yDiff < 0 Then
                    dRotate = Atn(yDiff / xDiff)
                Else
                    dRotate = Atn(yDiff / xDiff) '+ 3.14159265359
                End If
                
                If iStatus = 1 Then
                    dMidPnt(0) = dMidPnt(0) - 13.5 * yDiff / zDiff * dScale
                    dMidPnt(1) = dMidPnt(1) + 13.5 * xDiff / zDiff * dScale
                End If
                
                dTextInsert(0) = dMidPnt(0) - 13.5 * yDiff / zDiff * dScale
                dTextInsert(1) = dMidPnt(1) + 13.5 * xDiff / zDiff * dScale
            Case Is = 0
                dRotate = 1.570796327
                
                If yDiff > 0 Then
                    If iStatus = 1 Then
                        dMidPnt(0) = dMidPnt(0) - 7.5 * yDiff / zDiff * dScale
                        dMidPnt(1) = dMidPnt(1) + 7.5 * xDiff / zDiff * dScale
                    End If
                    
                    dTextInsert(0) = dMidPnt(0) - 7.5 * yDiff / zDiff * dScale
                    dTextInsert(1) = dMidPnt(1) + 7.5 * xDiff / zDiff * dScale
                Else
                    If iStatus = 1 Then
                        dMidPnt(0) = dMidPnt(0) - 13.5 * yDiff / zDiff * dScale
                        dMidPnt(1) = dMidPnt(1) + 13.5 * xDiff / zDiff * dScale
                    End If
                    
                    dTextInsert(0) = dMidPnt(0) - 13.5 * yDiff / zDiff * dScale
                    dTextInsert(1) = dMidPnt(1) + 13.5 * xDiff / zDiff * dScale
                End If
            Case Is > 0
                If yDiff = 0 Then
                    dRotate = 0#
                ElseIf yDiff > 0 Then
                    dRotate = Atn(yDiff / xDiff)
                Else
                    dRotate = Atn(yDiff / xDiff) '+ 3.14159265359
                End If
                
                If iStatus = 1 Then
                    dMidPnt(0) = dMidPnt(0) - 7.5 * yDiff / zDiff * dScale
                    dMidPnt(1) = dMidPnt(1) + 7.5 * xDiff / zDiff * dScale
                End If
                
                dTextInsert(0) = dMidPnt(0) - 7.5 * yDiff / zDiff * dScale
                dTextInsert(1) = dMidPnt(1) + 7.5 * xDiff / zDiff * dScale
        End Select
        
        dTextInsert(2) = 0#
        
        'If xDiff < 0 Then
            'dTextInsert(0) = dTextInsert(0) - 13.5 * yDiff / zDiff
            'dTextInsert(1) = dTextInsert(1) + 13.5 * xDiff / zDiff
        'Else
            'dTextInsert(0) = dTextInsert(0) - 7.5 * yDiff / zDiff
            'dTextInsert(1) = dTextInsert(1) + 7.5 * xDiff / zDiff
        'End If
        
        If iStatus = 0 Then
            lwpCoords(2) = lwpCoords(0) + (zDiff + 10) * xDiff / zDiff
            lwpCoords(3) = lwpCoords(1) + (zDiff + 10) * yDiff / zDiff
            lwpCoords(0) = lwpCoords(2) - (zDiff + 20) * xDiff / zDiff
            lwpCoords(1) = lwpCoords(3) - (zDiff + 20) * yDiff / zDiff
        
            Set objLWP = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
            objLWP.Layer = "Integrity Sheets"
            objLWP.Update
        
            vLWPO = objLWP.Offset(-5)
        Else
            dXScale = dScale * ((zDiff + 20) / 100)
            
            Err = 0
            Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dMidPnt, "RefLine", dXScale, dScale, dScale, dRotate)
            If Not Err = 0 Then
                MsgBox "Block may not be in drawing."
                GoTo Exit_Sub
            End If
            objBlock.Layer = strLayer
            objBlock.Update
        End If
            
        Set objText = ThisDrawing.ModelSpace.AddText(strLine, dTextInsert, dTScale)
        objText.Layer = "Integrity Sheets"
        objText.Alignment = acAlignmentCenter
        objText.TextAlignmentPoint = dTextInsert
        objText.Rotation = dRotate
        objText.Update
    Wend
Exit_Sub:
End Sub

Private Function GetScale(vCoords As Variant)
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strLine As String
    Dim dMinX, dMaxX, dMinY, dMaxY As Double
    Dim dScale As Double
    
    On Error Resume Next
    
    strLine = "none;;"
    
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
                strLine = objBlock.Layer & ";;"
                GoTo Exit_Sub
            End If
        End If
    Next objBlock
    
    dScale = 1#
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    strLine = strLine & dScale
    GetScale = strLine
End Function

