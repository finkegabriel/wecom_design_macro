VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TDOT 
   Caption         =   "Permit Drawings"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   OleObjectBlob   =   "TDOT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TDOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbAddSheets_Click() '<------------------------------------------------ Add Permit Layer if it doesn't exist
    Dim entPole As AcadObject
    Dim objLayer As AcadLayer
    Dim obrGSS As AcadBlockReference
    Dim obrSSInfo As AcadBlockReference
    Dim obrSSDWG1 As AcadBlockReference
    Dim attItem, attItem1, basePnt, returnPnt As Variant
    Dim vEmail As Variant
    Dim insertPnt(0 To 2) As Double
    Dim dScale As Double
    Dim iDWGNum, iNextNum As Long
    Dim str, str2 As String
    Dim strType, strLayer As String
    
    If cbDWG.Value = "" Then Exit Sub
    
    Me.Hide
  On Error Resume Next
    
    Call CreateLayer("Integrity Permits-" & cbDWG.Value)
  
    Err = 0
  
    iNextNum = 1
    iDWGNum = 1
    strType = cbDWG.Value
    
While Err = 0
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Staking Sheet: "
    If TypeOf entPole Is AcadBlockReference Then
        Set obrGSS = entPole
    Else
        GoTo Exit_Sub
    End If
    
    If Not Err = 0 Then GoTo Exit_Sub
    If Not obrGSS.Name = "SS-11x17" Then GoTo Exit_Sub
    
    obrGSS.Layer = "Integrity Permits-" & cbDWG.Value
    attItem = obrGSS.GetAttributes
    
    Err = 0
    iDWGNum = ThisDrawing.Utility.GetInteger(vbCr & "Enter Drawing Number(" & iNextNum & "): ")
    If Not Err = 0 Then iDWGNum = iNextNum
    iNextNum = iDWGNum + 1
    
    Select Case iDWGNum
        Case Is < 10
            attItem(0).TextString = strType & " 00" & iDWGNum
        Case Is > 99
            attItem(0).TextString = strType & " " & iDWGNum
        Case Else
            attItem(0).TextString = strType & " 0" & iDWGNum
    End Select
    
    obrGSS.Update
    dScale = obrGSS.XScaleFactor
    
    lbSheets.AddItem attItem(0).TextString & vbTab & dScale & vbTab & obrGSS.Name & vbTab & obrGSS.InsertionPoint(0) & vbTab & obrGSS.InsertionPoint(1)
  
    insertPnt(0) = obrGSS.InsertionPoint(0) + 20 * dScale
    insertPnt(1) = obrGSS.InsertionPoint(1) + 20 * dScale
    insertPnt(2) = obrGSS.InsertionPoint(2)
    
    Set obrSSInfo = ThisDrawing.ModelSpace.InsertBlock(insertPnt, "ss info", dScale, dScale, dScale, 0#)
    obrSSInfo.Layer = "Integrity Permits-" & cbDWG.Value
    
    attItem = obrSSInfo.GetAttributes
    attItem(0).TextString = tbProject.Value
    attItem(1).TextString = tbWO.Value
    attItem(2).TextString = cbExchange.Value
    attItem(3).TextString = tbRST.Value
    attItem(4).TextString = cbCounty.Value
    attItem(5).TextString = tbCity.Value
    attItem(6).TextString = iDWGNum
    attItem(7).TextString = tbTotalDWG.Value
    obrSSInfo.Update
    
    If iDWGNum = 1 Then
        insertPnt(0) = obrGSS.InsertionPoint(0) + 20 * dScale
        insertPnt(1) = obrGSS.InsertionPoint(1) + 280 * dScale
        insertPnt(2) = obrGSS.InsertionPoint(2)
    
        Set obrSSDWG1 = ThisDrawing.ModelSpace.InsertBlock(insertPnt, "ss dwg1", dScale, dScale, dScale, 0#)
        obrSSDWG1.Layer = "Integrity Permits-" & cbDWG.Value
    
        attItem1 = obrSSDWG1.GetAttributes
        
        vEmail = Split(cbDesigner.Value, " ")
        If UBound(vEmail) = 1 Then
            str2 = LCase(vEmail(0)) & "." & LCase(vEmail(1)) & "@integrity-us.com"
        Else
            str2 = "NA"
        End If
        str = cbDesigner.Value & "\P" & str2 & "\P" & tbPhoneNumber.Value & "\P"
        str = str & "730 Middle Tennessee Blvd - Suite 6\PMurfreesboro, TN 37130"
        attItem1(0).TextString = str
        obrSSDWG1.Update
    End If
    Err = 0
Wend
    'layerObj.Delete
    
Exit_Sub:
    Set layerObj = ThisDrawing.Layers.Add("0")
    ThisDrawing.ActiveLayer = layerObj
    Me.show
End Sub

Private Sub PlaceSSBlock(strName As String, vInsertPnt As Variant, strLayer As String)
    Dim obrGSS As AcadBlockReference
    Dim obrSSInfo As AcadBlockReference
    Dim obrSSDWG1 As AcadBlockReference
    Dim attItem, attItem1 As Variant
    Dim vEmail, vTemp As Variant
    Dim insertPnt(0 To 2) As Double
    Dim dScale As Double
    Dim str, str2 As String
    Dim iDWG As Integer
    
  On Error Resume Next
    
    'Call CreateLayer("Integrity Permits-" & cbDWG.Value)
  
    Err = 0
    
    vTemp = Split(strName, " ")
    iDWG = CInt(vTemp(1))
    
    insertPnt(0) = vInsertPnt(0)
    insertPnt(1) = vInsertPnt(1)
    insertPnt(2) = vInsertPnt(2)
    dScale = CInt(cbScale.Value) / 100 'obrGSS.XScaleFactor
    
    If Not Err = o Then Exit Sub
    
    
    Set obrGSS = ThisDrawing.ModelSpace.InsertBlock(insertPnt, "SS-11x17", dScale, dScale, dScale, 0#)
    obrGSS.Layer = strLayer
    attItem = obrGSS.GetAttributes
    attItem(0).TextString = strName
    obrGSS.Update
    
    insertPnt(0) = obrGSS.InsertionPoint(0) + 20 * dScale
    insertPnt(1) = obrGSS.InsertionPoint(1) + 20 * dScale
    insertPnt(2) = obrGSS.InsertionPoint(2)
    
    Set obrSSInfo = ThisDrawing.ModelSpace.InsertBlock(insertPnt, "ss info", dScale, dScale, dScale, 0#)
    obrSSInfo.Layer = strLayer
    
    attItem = obrSSInfo.GetAttributes
    attItem(0).TextString = tbProject.Value
    attItem(1).TextString = tbWO.Value
    attItem(2).TextString = cbExchange.Value
    attItem(3).TextString = tbRST.Value
    attItem(4).TextString = cbCounty.Value
    attItem(5).TextString = tbCity.Value
    attItem(6).TextString = iDWG
    attItem(7).TextString = tbTotalDWG.Value
    obrSSInfo.Update
    
    If iDWG = 1 Then
        insertPnt(0) = obrGSS.InsertionPoint(0) + 20 * dScale
        insertPnt(1) = obrGSS.InsertionPoint(1) + 280 * dScale
        insertPnt(2) = obrGSS.InsertionPoint(2)
    
        Set obrSSDWG1 = ThisDrawing.ModelSpace.InsertBlock(insertPnt, "ss dwg1", dScale, dScale, dScale, 0#)
        obrSSDWG1.Layer = strLayer
    
        attItem1 = obrSSDWG1.GetAttributes
        
        vEmail = Split(cbDesigner.Value, " ")
        If UBound(vEmail) = 1 Then
            str2 = LCase(vEmail(0)) & "." & LCase(vEmail(1)) & "@integrity-us.com"
        Else
            str2 = "NA"
        End If
        str = cbDesigner.Value & "\P" & str2 & "\P" & tbPhoneNumber.Value & "\P"
        str = str & "730 Middle Tennessee Blvd - Suite 6\PMurfreesboro, TN 37130"
        attItem1(0).TextString = str
        obrSSDWG1.Update
    End If
               
        str = strName & vbTab & obrGSS.XScaleFactor
        str = str & vbTab & obrGSS.InsertionPoint(0) & vbTab & obrGSS.InsertionPoint(1)
        lbSheets.AddItem str
End Sub

Private Sub cbArray_Click()
    Dim objSSBlock As AcadBlockReference
    Dim vReturnPnt, vAttList As Variant
    Dim vSSData, vDWGNum As Variant
    Dim strLast As String
    Dim strLayer As String
    Dim str As String
    Dim dInsertPnt(0 To 2) As Double
    Dim dScale, dTest As Double
    Dim iDWG, iLastDWG, iTest As Integer
    
    If cbDWG.Value = "" Then Exit Sub
    If tbTotalDWG.Value = "" Then Exit Sub
    iLastDWG = CInt(tbTotalDWG.Value)
    
    'MsgBox lbSheets.ListCount & vbCr & lbSheets.List(0) & "--"
    'Exit Sub
    
    If lbSheets.ListCount = 1 Then
        If lbSheets.List(0) = "" Then lbSheets.Clear
    End If
    
    Me.Hide
    
    If lbSheets.ListCount < 1 Then
        Call CreateLayer("Integrity Permits-" & cbDWG.Value)
        cbDWG.AddItem cbDWG.Value
        vReturnPnt = ThisDrawing.Utility.GetPoint(, "Insertion Point of " & cbDWG.Value & ":")
        iDWG = 1
        dScale = CDbl(cbScale.Value) / 100
        dInsertPnt(0) = vReturnPnt(0)
        dInsertPnt(1) = vReturnPnt(1)
    Else
        strLast = lbSheets.List(lbSheets.ListCount - 1)
        vSSData = Split(strLast, vbTab)
        vDWGNum = Split(vSSData(0), " ")
        iDWG = CInt(vDWGNum(1)) + 1
        
        If iDWG > iLastDWG Then GoTo Exit_Sub2
        
        dScale = CDbl(vSSData(1))
        dTest = iDWG / 10
        dTest = (dTest - CInt(dTest)) * 10
        If dTest = 1 Then
            dInsertPnt(0) = CDbl(vSSData(2)) + (2000 * dScale)
            dInsertPnt(1) = CDbl(vSSData(3)) + (10800 * dScale)
        Else
            dInsertPnt(0) = CDbl(vSSData(2))
            dInsertPnt(1) = CDbl(vSSData(3)) - (1200 * dScale)
        End If
    End If
    dInsertPnt(2) = 0#
    
    If cbDWG.Value = "DWG" Then
        strLayer = "Integrity Sheets"
    Else
        strLayer = "Integrity Permits-" & cbDWG.Value
    End If
    
    For i = iDWG To iLastDWG
        strLast = cbDWG.Value
        Select Case i
            Case Is < 10
                strLast = strLast & " 00" & i
            Case Is > 99
                strLast = strLast & " " & i
            Case Else
                strLast = strLast & " 0" & i
        End Select
        
        Call PlaceSSBlock(strLast, dInsertPnt, strLayer)
         
        'Set objSSBlock = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "SS-11x17", dScale, dScale, dScale, 0#)
        'objSSBlock.Layer = strLayer
        
        'vAttList = objSSBlock.GetAttributes
        
        
        'vAttList(0).TextString = strLast
        'objSSBlock.Update
        
        dTest = i / 10
        dTest = (dTest - CInt(dTest)) * 10
        'MsgBox dTest
        If dTest = 0 Then
            dInsertPnt(0) = dInsertPnt(0) + (2000 * dScale)
            dInsertPnt(1) = dInsertPnt(1) + (10800 * dScale)
        Else
            dInsertPnt(1) = dInsertPnt(1) - (1200 * dScale)
        End If
    Next i
    
    Call SortLBItems
Exit_Sub2:
    Call UpdateAllDWGs
    Me.show
End Sub

Private Sub cbCreateDWG_Click()
    Me.Hide
        Load CreateTDOT
            CreateTDOT.show
        Unload CreateTDOT
    Me.show
    
    Exit Sub
    
    
    
    Dim objSS4 As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim returnPnt As Variant
    Dim dCoords(0 To 8) As Double
    Dim dStartPnt(0 To 2) As Double
    Dim dEndPnt(0 To 2) As Double
    Dim dLL(0 To 2) As Double
    Dim dUR(0 To 2) As Double
    Dim dTo(0 To 2) As Double
    Dim dScale As Double
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim strCommand As String
    
    'If cbDWG.Value = "DWG" Then Exit Sub
    
  On Error Resume Next
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, returnPnt, "Select DWG: "
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
        If Not objBlock.Name = "SS-11x17" Then
            MsgBox "Not a valid DWG."
            Me.show
            Exit Sub
        End If
        
        Set obrBlock = objEntity
        dCoords(0) = objBlock.InsertionPoint(0)
        dCoords(1) = objBlock.InsertionPoint(1)
        dCoords(2) = objBlock.InsertionPoint(2)
        'dLL(0) = objBlock.InsertionPoint(0)
        'dLL(1) = objBlock.InsertionPoint(1)
        'dLL(2) = objBlock.InsertionPoint(2)
        
        dScale = objBlock.XScaleFactor
        
        dCoords(3) = dCoords(0) + (1652 * dScale)
        dCoords(4) = dCoords(1) + (1052 * dScale)
        dCoords(5) = 0#
        'dUR(0) = dStartPnt(0) + (1652 * dScale)
        'dUR(1) = dStartPnt(1) + (1052 * dScale)
        'dUR(2) = 0#
    Else
        MsgBox "Not a valid DWG."
        Exit Sub
    End If
    
    ThisDrawing.Utility.GetEntity objEntity, returnPnt, "Select TDOT: "
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
        If Not objBlock.Name = "SS-11x17" Then
            MsgBox "Not a valid DWG."
            Me.show
            Exit Sub
        End If
        
        Set obrBlock = objEntity
        dCoords(6) = objBlock.InsertionPoint(0)
        dCoords(7) = objBlock.InsertionPoint(1)
        dCoords(8) = objBlock.InsertionPoint(2)
        'dTO(0) = objBlock.InsertionPoint(0)
        'dTO(1) = objBlock.InsertionPoint(1)
        'dTO(2) = objBlock.InsertionPoint(2)
    Else
        MsgBox "Not a valid Permit."
        Exit Sub
    End If
    
    If Not Err = 0 Then
        Me.show
        Exit Sub
    End If
    
'    grpCode(0) = 8
'    grpValue(0) = strLayer  '"Integrity Roads-EOP"
'    filterType = grpCode
'    filterValue = grpValue
    
'    Set objSS4 = ThisDrawing.SelectionSets.Add("objSS3")
'    objSS4.Select acSelectionSetWindow, dLL, dUR, filterType, filterValue
    
'    strCommand = "COPY" & vbCr & "P" & vbCr & vbCr
'    strCommand = strCommand & dLL(0) & "," & dLL(1) & ",0" & vbCr
'    strCommand = strCommand & dTO(0) & "," & dTO(1) & ",0" & vbCr & vbCr
    
'    ThisDrawing.SendCommand strCommand
    
'    objSS4.Clear
'    objSS4.Delete
    
    Call CopyDWGtoPermit(dCoords, "Integrity Roads-EOP,Integrity Roads-Clearance,Roads_MText")
    Call CopyDWGtoPermit(dCoords, "Integrity Poles-Others,Integrity Poles-Power,Integrity Poles-UTC")
    Call CopyDWGtoPermit(dCoords, "Integrity Cable,Integrity Cable-Aerial,Integrity Cable-Aerial Text")
    Call CopyDWGtoPermit(dCoords, "Parcels,Integrity Cable-Buried,Integrity Cable-Buried Text")
    Call CopyDWGtoPermit(dCoords, "Integrity Building,Integrity Business,Integrity Building Misc")
    Call CopyDWGtoPermit(dCoords, "Integrity Building-BUS,Integrity Building-SCH,Integrity Building-CHU")
    Call CopyDWGtoPermit(dCoords, "Integrity Building-RES,Integrity Building-TRL,Integrity Building-MDU")
    'Call CopyDWGtoPermit(dCoords, "Integrity Roads-Clearance")
    'Call CopyDWGtoPermit(dCoords, "Roads_MText")
    'Call CopyDWGtoPermit(dCoords, "Integrity Poles-Power")
    'Call CopyDWGtoPermit(dCoords, "Integrity Cable-Aerial")
    'Call CopyDWGtoPermit(dCoords, "Integrity Cable-Aerial Text")
    'Call CopyDWGtoPermit(dCoords, "Integrity Cable-Buried")
    'Call CopyDWGtoPermit(dCoords, "Integrity Cable-Buried Text")
    'Call CopyDWGtoPermit(dCoords, "Integrity Building-SCH")
    'Call CopyDWGtoPermit(dCoords, "Integrity Building-TRL")
    'Call CopyDWGtoPermit(dCoords, "Integrity Building-CHU")
    'Call CopyDWGtoPermit(dCoords, "Integrity Building-MDU")
    'Call CopyDWGtoPermit(dCoords, "Integrity Business")
    'Call CopyDWGtoPermit(dCoords, "Integrity Building Misc")
    'Call CopyDWGtoPermit(dCoords, "Integrity Sheets")
    
    Me.show
End Sub

Private Sub cbDWG_Change()
    Call FillOutDrawingList
End Sub

Private Sub GetDWGList(strLayer As String)
    Dim objSS3 As AcadSelectionSet
    Dim entBlock As AcadObject
    Dim obrTemp As AcadBlockReference
    Dim attItem2, vTemp As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim str As String
    
  On Error Resume Next
    lbSheets.Clear
    
    grpCode(0) = 8
    grpValue(0) = "Integrity Sheets"    'strLayer  '"Integrity Permits"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS3.Select acSelectionSetAll, , , filterType, filterValue
    
    If Not Err = 0 Then MsgBox "Error"
    
    For Each entBlock In objSS3
        Set obrTemp = entBlock
        
        If obrTemp.Name = "SS-11x17" Then
            attItem2 = obrTemp.GetAttributes
            'vTemp = Split(attItem2(0).TextString, " ")
               
            str = attItem2(0).TextString & vbTab & obrTemp.XScaleFactor
            str = str & vbTab & obrTemp.InsertionPoint(0) & vbTab & obrTemp.InsertionPoint(1)
            lbSheets.AddItem str
        End If
Next_entBlock:
    Next entBlock
    
    Call SortLBItems
Exit_Sub:
    objSS3.Delete
End Sub

Private Sub cbHighlight_Click()
    Dim objBlock As AcadBlockReference
    Dim objEntity As AcadEntity
    Dim vAttList, vBasePnt As Variant
    Dim objMText As AcadMText
    Dim strMText As String
    Dim dInsertPnt(0 To 2) As Double
    
    Me.Hide
    On Error Resume Next
    ThisDrawing.SetVariable "DIMTFILLCOLOR", 50
    Err = 0
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, vbCr & "Select Pole Text:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    If Not Err = 0 Then GoTo Exit_Sub
    
    vAttList = objBlock.GetAttributes
    strMText = vAttList(0).TextString
    vAttList(0).TextString = ""
    objBlock.Update
    
    dInsertPnt(0) = vBasePnt(0)
    dInsertPnt(1) = vBasePnt(1)
    dInsertPnt(2) = 0#
    
    Set objMText = ThisDrawing.ModelSpace.AddMText(dInsertPnt, 15, strMText)

    objMText.Layer = "Integrity Poles-Power"
    objMText.Height = 6
    objMText.AttachmentPoint = acAttachmentPointMiddleCenter
    objMText.InsertionPoint = dInsertPnt
    'objMText.Rotation = dRotate
    objMText.BackgroundFill = True
    'objMText.BackgroundFill = acYellow
    'objMText.Highlight
    objMText.Update
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbMatchlines_Click()
    Me.Hide
    Call Module1.AddMatchlines
    Me.show
    Exit Sub
    
    Dim objLWP As AcadLWPolyline
    Dim objText As AcadText
    Dim basePnt, returnPnt As Variant
    Dim vLWPO As Variant
    Dim lwpCoords(0 To 3) As Double
    Dim dTextInsert(0 To 2) As Double
    Dim xDiff, yDiff, zDiff As Double
    Dim dRaotate As Double
    Dim strLine As String
    Dim strLayer As String
    
    If cbDWG.Value = "" Then Exit Sub
    
  On Error Resume Next
    
    Me.Hide
    
    If cbDWG.Value = "DWG" Then
        strLayer = "Integrity Sheets"
    Else
        strLayer = "Integrity Permits-" & cbDWG.Value
        Call CreateLayer(strLayer)
    End If
    
    While Err = 0
        returnPnt = ThisDrawing.Utility.GetPoint(, "Select Mid Point: ")
        lwpCoords(0) = returnPnt(0)
        lwpCoords(1) = returnPnt(1)
        
        If Not Err = 0 Then GoTo Exit_Sub
        
        returnPnt = ThisDrawing.Utility.GetPoint(, "Select Point: ")
        lwpCoords(2) = returnPnt(0)
        lwpCoords(3) = returnPnt(1)
        
        dTextInsert(0) = (lwpCoords(0) + lwpCoords(2)) / 2
        dTextInsert(1) = (lwpCoords(1) + lwpCoords(3)) / 2
        dTextInsert(2) = 0#
        
        If Not Err = 0 Then GoTo Exit_Sub
        
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
            Case Is = 0
                dRotate = 1.570796327
            Case Is > 0
                If yDiff = 0 Then
                    dRotate = 0#
                ElseIf yDiff > 0 Then
                    dRotate = Atn(yDiff / xDiff)
                Else
                    dRotate = Atn(yDiff / xDiff) '+ 3.14159265359
                End If
        End Select
        
        If xDiff < 0 Then
            dTextInsert(0) = dTextInsert(0) - 13.5 * yDiff / zDiff
            dTextInsert(1) = dTextInsert(1) + 13.5 * xDiff / zDiff
        Else
            dTextInsert(0) = dTextInsert(0) - 7.5 * yDiff / zDiff
            dTextInsert(1) = dTextInsert(1) + 7.5 * xDiff / zDiff
        End If
        
        lwpCoords(2) = lwpCoords(0) + (zDiff + 10) * xDiff / zDiff
        lwpCoords(3) = lwpCoords(1) + (zDiff + 10) * yDiff / zDiff
        lwpCoords(0) = lwpCoords(2) - (zDiff + 20) * xDiff / zDiff
        lwpCoords(1) = lwpCoords(3) - (zDiff + 20) * yDiff / zDiff
        
        Set objLWP = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
        objLWP.Layer = strLayer
        objLWP.Update
        
        vLWPO = objLWP.Offset(-5)
        
        strLine = ThisDrawing.Utility.GetString(0, "Matches Drawing #: ")
        
        If Not Err = 0 Then GoTo Exit_Sub
        
        Select Case Len(strLine)
            Case Is = 1
                strLine = "SEE DWG 00" & strLine
            Case Is = 2
                strLine = "SEE DWG 0" & strLine
            Case Is = 3
                strLine = "SEE DWG " & strLine
        End Select
            
        Set objText = ThisDrawing.ModelSpace.AddText(strLine, dTextInsert, 6)
        objText.Layer = strLayer
        objText.Alignment = acAlignmentCenter
        objText.TextAlignmentPoint = dTextInsert
        objText.Rotation = dRotate
        objText.Update
    Wend
Exit_Sub:
    Me.show
End Sub

Private Sub cbPrintSelected_Click()
    Dim strArraySheets() As String
    Dim attTotalLine, attTotalSplit As Variant
    Dim str, str1 As String
    Dim strDWG As String
    Dim dScale As Double
    Dim dLL(0 To 1), dUR(0 To 1) As Double
    Dim strPlot(0 To 24) As String
    Dim i As Integer
    Dim iCount As Integer
    Dim iResult As Integer
    Dim strCommand As String
    Dim strPath As String
    Dim strFileName As String
    Dim strArray As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    strPath = ThisDrawing.Path & "\"
    
    iCount = 0
    
    strPlot(0) = "-plot"
    strPlot(1) = "y"
    strPlot(2) = ""
    'strPlot(3) = "Microsoft Print to PDF"
    strPlot(4) = "TABLOID"
    strPlot(5) = "i" '"i"
    strPlot(6) = "LANDSCAPE"
    strPlot(7) = "n"
    strPlot(8) = "w"
    'strPlot(9) = ""
    'strPlot(10) = ""
    'strPlot(11) = "" '"1=75"
    strPlot(12) = "c"
    strPlot(13) = "y"
    strPlot(14) = ""    '"United.ctb"
    strPlot(15) = "y"
    strPlot(16) = "a"
    strPlot(17) = "n"
    strPlot(18) = "n"
    strPlot(19) = "y"
    'strPlot(20) = ""
    strPlot(21) = ""
    strPlot(22) = ""
    strPlot(23) = ""
    strPlot(24) = ""
    
    ThisDrawing.SetVariable "FILEDIA", 0
    ThisDrawing.SetVariable "CMDDIA", 0
    
    For i = 0 To lbSheets.ListCount - 1
        If lbSheets.Selected(i) Then
            iCount = iCount + 1
            
            str = lbSheets.List(i)
            attTotalSplit = Split(str, vbTab)
            strDWG = attTotalSplit(0)
            dScale = CDbl(attTotalSplit(1))
            dLL(0) = CDbl(attTotalSplit(2))
            dLL(1) = CDbl(attTotalSplit(3))
            dUR(0) = dLL(0) + (1652 * dScale)
            dUR(1) = dLL(1) + (1052 * dScale)
    
            viewCoordsB(0) = dLL(0)
            viewCoordsB(1) = dLL(1)
            viewCoordsB(2) = 0#
            viewCoordsE(0) = dUR(0)
            viewCoordsE(1) = dUR(1)
            viewCoordsE(2) = 0#
    
            ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
            
            strPlot(9) = (dLL(0) - (1 * dScale)) & "," & (dLL(1) - (1 * dScale))
            strPlot(10) = (dUR(0) + (1 * dScale)) & "," & (dUR(1) + (1 * dScale))
            strPlot(11) = "1=" & CStr((dScale + 0.02) * 100)
    
            strFileName = """" & strPath & "Individual Sheets\" & strDWG & ".pdf"""
            strFileName = Replace(strFileName, "/", "\")
            'strFileName = Replace(strFileName, "\", "/")
            strPlot(20) = strFileName
            'MsgBox strFileName
            'Exit Sub
            
            strCommand = ""
            For j = 0 To 19
                If strPlot(j) = "" Then
                    strCommand = strCommand & vbCr
                Else
                    strCommand = strCommand & strPlot(j) & vbCr
                End If
            Next j
            
            ThisDrawing.SendCommand strCommand
            
            Select Case iCount
                Case Is = 2, Is = 10, Is = 20
                    iResult = MsgBox("Continue Plotting?", vbYesNo, "Continue Plotting")
                        If iResult = vbNo Then GoTo Exit_Next
            End Select
        End If
    Next i
Exit_Next:
    
    ThisDrawing.SetVariable "FILEDIA", 1
    ThisDrawing.SetVariable "CMDDIA", 1
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbSelectAll_Click()
    For i = 0 To lbSheets.ListCount - 1
        If lbSheets.List(i) = "" Then GoTo Next_I
        'lbSheets.List(i).Enabled = True
        lbSheets.Selected(i) = True
Next_I:
    Next i
End Sub

Private Sub cbSelectNone_Click()
    For i = 0 To lbSheets.ListCount - 1
        lbSheets.Selected(i) = False
    Next i
End Sub

Private Sub Label3_Click()
    Dim objObject As AcadObject
    Dim obrGP2 As AcadBlockReference
    Dim basePnt, attItem As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objObject, basePnt, "Select SS Info: "
    If TypeOf objObject Is AcadBlockReference Then
        Set obrGP2 = objObject
    Else
        Me.show
        Exit Sub
    End If
    attItem = obrGP2.GetAttributes
    
    tbProject.Value = attItem(0).TextString
    tbWO.Value = attItem(1).TextString
    cbExchange.Value = attItem(2).TextString
    tbRST.Value = attItem(3).TextString
    cbCounty.Value = attItem(4).TextString
    tbCity.Value = attItem(5).TextString
    tbTotalDWG = attItem(7).TextString
    
    Me.show
End Sub

Private Sub lbSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim str As String
    Dim strArray As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    str = lbSheets.List(lbSheets.ListIndex)
    strArray = Split(str, vbTab)
    
    viewCoordsB(0) = strArray(2)
    viewCoordsB(1) = strArray(3)
    viewCoordsB(2) = 0#
    viewCoordsE(0) = viewCoordsB(0) + 1650 * CDbl(strArray(1))
    viewCoordsE(1) = viewCoordsB(1) + 1050 * CDbl(strArray(1))
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    Me.show
End Sub

Private Sub UserForm_Initialize()
    cbExchange.AddItem "FLOR"
    cbExchange.AddItem "INTE"
    cbExchange.AddItem "SMYR"
    cbExchange.AddItem "ALMA"
    cbExchange.AddItem "MRBO"
    cbExchange.AddItem "AIR"
    cbExchange.AddItem "ROCK"
    cbExchange.AddItem "CHRI"
    cbExchange.AddItem "KITT"
    cbExchange.AddItem "CAIT"
    cbExchange.AddItem "LAK"
    cbExchange.AddItem "LAS"
    cbExchange.AddItem "LEB"
    cbExchange.AddItem "MAR"
    cbExchange.AddItem "NOR"
    cbExchange.AddItem "SOU"
    cbExchange.AddItem "VES"
    cbExchange.AddItem "WAT"
    cbExchange.AddItem "Belfast"
    cbExchange.AddItem "Bell Buckle"
    cbExchange.AddItem "Chapel Hill"
    cbExchange.AddItem "Chapel Hill-CLEC"
    cbExchange.AddItem "College Grove"
    cbExchange.AddItem "College Grove-CLEC"
    cbExchange.AddItem "Flat Creek"
    cbExchange.AddItem "Fosterville"
    cbExchange.AddItem "Franklin-CLEC"
    cbExchange.AddItem "Nolensville"
    cbExchange.AddItem "Nolensville-CLEC"
    cbExchange.AddItem "Shelbyville-CLEC"
    cbExchange.AddItem "Triune"
    cbExchange.AddItem "Unionville"
    
    cbCounty.AddItem "Bedford"
    cbCounty.AddItem "Davidson"
    cbCounty.AddItem "Marshall"
    cbCounty.AddItem "Maury"
    cbCounty.AddItem "Rutherford"
    cbCounty.AddItem "Williamson"
    cbCounty.AddItem "Wilson"
    
    cbDesigner.AddItem "Dylan Spears"
    cbDesigner.AddItem "Jeremy Pafford"
    cbDesigner.AddItem "Ronn Elliott"
    cbDesigner.AddItem "Rich Taylor"
    cbDesigner.AddItem "Jason Pafford"
    cbDesigner.AddItem "Adam Kemper"
    cbDesigner.AddItem "Byron Auer"
    cbDesigner.AddItem "Jon Wilburn"
    cbDesigner.AddItem "Franklin Angulo"
    cbDesigner.AddItem "Jay Penny"
    cbDesigner.AddItem "Sam Jackson"
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "75"
    
    Dim objSS3 As AcadSelectionSet
    Dim entBlock As AcadObject
    Dim obrTemp As AcadBlockReference
    Dim attItem2, vTemp As Variant
    Dim mode As Integer
    'Dim str As String
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim str, strList As String
    'Dim vTemp As Variant
    
  On Error Resume Next
  
    strList = ""
    
    str = ThisDrawing.Name
    vTemp = Split(str, " ")
    tbWO.Value = vTemp(0)
    
    Select Case UBound(vTemp)
        Case Is = 1
            str = vTemp(1)
        Case Is > 1
            str = vTemp(1)
            For i = 2 To UBound(vTemp)
                str = str & " " & vTemp(i)
            Next i
    End Select
    tbProject.Value = str
    
    'grpCode(0) = 8
    'grpValue(0) = "Integrity Permits"
    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS3.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entBlock In objSS3
        Set obrTemp = entBlock
        attItem2 = obrTemp.GetAttributes
        
        vTemp = Split(attItem2(0).TextString, " ")
        If vTemp(0) = "dwg" Then GoTo Next_entBlock
        If Asc(vTemp(0)) > 47 And Asc(vTemp(0)) < 58 Then GoTo Next_entBlock
        
        If InStr(strLine, vTemp(0)) = 0 Then
            strLine = strLine & vTemp(0) & "<>"
        End If
Next_entBlock:
    Next entBlock
    
    cbDWG.AddItem ""
    vTemp = Split(strLine, "<>")
    For i = LBound(vTemp) To UBound(vTemp)
        If Not vTemp(i) = "" Then
            If vTemp(i) = "DWG" Then
                cbDWG.AddItem vTemp(i), 1
                'cbDWG.Value = "DWG"
            Else
                cbDWG.AddItem vTemp(i)
            End If
        'If vTemp(i) = "DWG" Then
            'cbDWG.Value = "DWG"
            'Call GetDWGList("Integrity Sheets")
        End If
    Next i
    
    Call FillOutDrawingList
    'Call SortLBItems
Exit_Sub:
    objSS3.Delete
End Sub

Private Sub cbUpdateAllSS_Click()
    Dim objSS3 As AcadSelectionSet
    Dim entBlock As AcadObject
    Dim obrTemp As AcadBlockReference
    Dim attItem2 As Variant
    Dim mode As Integer
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim str As String
    
  On Error Resume Next
    
    grpCode(0) = 8
    If cbDWG.Value = "DWG" Then
        grpValue(0) = "Integrity Sheets"
    Else
        grpValue(0) = "Integrity Permits-" & cbDWG.Value
    End If
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS3.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entBlock In objSS3
        Set obrTemp = entBlock
        
        If obrTemp.Name = "SS Info" Then
            attItem2 = obrTemp.GetAttributes
            attItem2(0).TextString = tbProject.Value
            attItem2(1).TextString = tbWO.Value
            attItem2(2).TextString = cbExchange.Value
            attItem2(3).TextString = tbRST.Value
            attItem2(4).TextString = cbCounty.Value
            attItem2(5).TextString = tbCity.Value
            attItem2(7).TextString = tbTotalDWG.Value
            
            obrTemp.Update
        End If
    Next entBlock
    
    Call SortLBItems
Exit_Sub:
    objSS3.Delete
End Sub

Private Sub cbLineCallout_Click()
    Dim objEntity As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objPolyline As AcadPolyline
    Dim objLine As AcadLine
    Dim objMText As AcadMText
    Dim objText As AcadText
    Dim vBasePnt As Variant
    Dim vStartPnt, vEndPnt As Variant
    Dim vCoords As Variant
    Dim dRotate, dScale, Pi As Double
    Dim dInsertPnt(0 To 2) As Double
    Dim dStartPnt(0 To 2) As Double
    Dim dEndPnt(0 To 2) As Double
    Dim dDiff(0 To 2) As Double
    Dim dToText(0 To 2) As Double
    Dim dFromText(0 To 2) As Double
    Dim dTemp, dDistance As Double
    Dim strLine, strLayer As String
    
    On Error Resume Next
    
    Me.Hide
  
    Pi = 3.14159265359
    strLine = ""
    
    While Err = 0
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Get Line or Polyline"
    If Not Err = 0 Then
        Me.show
        Exit Sub
    End If
    
    If Not Err = 0 Then
        Me.show
        Exit Sub
    End If
    
    Select Case objEntity.ObjectName
        Case "AcDbLine"
            Set objLine = objEntity
            'iLength = CInt(objLine.Length)
            vStartPnt = objLine.StartPoint
            vEndPnt = objLine.EndPoint
            
            dDiff(0) = vEndPnt(0) - vStartPnt(0)
            dDiff(1) = vEndPnt(1) - vStartPnt(1)
            dDiff(2) = Sqr((dDiff(0) * dDiff(0)) + (dDiff(1) * dDiff(1)))
            
            If vBasePnt(0) > vStartPnt(0) Then
                dToText(0) = vBasePnt(0) - vStartPnt(0)
                dToText(1) = vBasePnt(1) - vStartPnt(1)
            Else
                dToText(0) = vStartPnt(0) - vBasePnt(0)
                dToText(1) = vStartPnt(1) - vBasePnt(1)
            End If
            dToText(2) = Sqr((dToText(0) * dToText(0)) + (dToText(1) * dToText(1)))
            
            dScale = Abs(dToText(2) / dDiff(2))
            dInsertPnt(0) = vStartPnt(0) + (dDiff(0) * dScale)
            dInsertPnt(1) = vStartPnt(1) + (dDiff(1) * dScale)
            dInsertPnt(2) = 0#
            
            If dToText(0) = 0 Then
                dRotate = Pi / 2
            Else
                dRotate = Atn(dToText(1) / dToText(0))
            End If
            
            strLayer = objLine.Layer
            Select Case strLayer
                Case "Parcels"
                    strLine = "R/W"
                Case "Integrity Roads-EOP"
                    strLine = "EOP"
                Case "Integrity Roads-CL"
                    strLine = "C/L"
                Case "Integrity Cable-Aerial", "Integrity Cable"
                    strLine = "CABLE"
                Case Else
                    strLine = tbLTText.Value
            End Select
        Case "AcDbPolyline"
            Set objLWP = objEntity
            vCoords = objLWP.Coordinates
            
            For i = 3 To UBound(vCoords) Step 2
                dDiff(0) = vCoords(i - 1) - vCoords(i - 3)
                dDiff(1) = vCoords(i) - vCoords(i - 2)
                dDiff(2) = Sqr((dDiff(0) * dDiff(0)) + (dDiff(1) * dDiff(1)))
            
                dToText(0) = vBasePnt(0) - vCoords(i - 3)
                dToText(1) = vBasePnt(1) - vCoords(i - 2)
                dToText(2) = Sqr((dToText(0) * dToText(0)) + (dToText(1) * dToText(1)))
                
                If dDiff(2) > dToText(2) Then
                    dFromText(0) = vBasePnt(0) - vCoords(i - 1)
                    dFromText(1) = vBasePnt(1) - vCoords(i)
                    dFromText(2) = Sqr((dFromText(0) * dFromText(0)) + (dFromText(1) * dFromText(1)))
                    
                    dTemp = (dFromText(2) + dToText(2)) / dDiff(2)
                    If dTemp < 1.01 Then
                        GoTo Exit_First_Next
                    End If
                End If
            Next i
Exit_First_Next:
            
            dScale = Abs(dToText(2) / dDiff(2))
            dInsertPnt(0) = vCoords(i - 3) + (dDiff(0) * dScale)
            dInsertPnt(1) = vCoords(i - 2) + (dDiff(1) * dScale)
            dInsertPnt(2) = 0#
            
            If dToText(0) = 0 Then
                dRotate = Pi / 2
            Else
                dRotate = Atn(dToText(1) / dToText(0))
            End If
            
            strLayer = objLWP.Layer
            Select Case strLayer
                Case "Parcels"
                    strLine = "R/W"
                Case "Integrity Roads-EOP"
                    strLine = "EOP"
                Case "Integrity Roads-CL"
                    strLine = "C/L"
                Case "Integrity Cable-Aerial", "Integrity Cable", "IS_Fiber"
                    strLine = "CABLE"
                Case Else
                    strLine = tbLTText.Value
            End Select
    End Select
            
'    Set objText = ThisDrawing.ModelSpace.AddText(strLine, dInsertPnt, 4.5)
'    objText.Layer = strLayer
'    objText.Alignment = acAlignmentMiddle
'    objText.TextAlignmentPoint = dInsertPnt
'    objText.Rotation = dRotate
'    objText.Update

    dScale = CInt(cbScale.Value) / 100

    Set objMText = ThisDrawing.ModelSpace.AddMText(dInsertPnt, 15, strLine)
    objMText.Layer = strLayer
    objMText.Height = 6 * dScale
    objMText.AttachmentPoint = acAttachmentPointMiddleCenter
    objMText.InsertionPoint = dInsertPnt
    objMText.Rotation = dRotate
    objMText.BackgroundFill = True
    objMText.Update
    
    Wend
    
    Me.show
End Sub

Private Sub SortLBItems()
    Dim iCount As Integer
    Dim vItem, vItem1 As Variant
    Dim strTemp As String
    
    iCount = lbSheets.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            vItem = Split(lbSheets.List(b), vbTab)
            vItem1 = Split(lbSheets.List(b + 1), vbTab)
            If vItem(0) > vItem1(0) Then
                
                strTemp = lbSheets.List(b + 1)
                lbSheets.List(b + 1) = lbSheets.List(b)
                lbSheets.List(b) = strTemp
            End If
        Next b
    Next a
End Sub

Private Sub CreateLayer(strLayer As String)
    Dim objLayer As AcadLayer
    
    On Error Resume Next
    
    Set objLayer = ThisDrawing.Layers.Add(strLayer)
    
    If Not Err = 0 Then Set objLayer = ThisDrawing.Layers.Item(strLayer)
    
    objLayer.color = acGreen
End Sub

Private Sub CopyDWGtoPermit(vCoords As Variant, strLayer As String)
    Dim objSS4 As AcadSelectionSet
    Dim entBlock As AcadEntity
    Dim objEMove As AcadEntity
    Dim dLL(0 To 2) As Double
    Dim dUR(0 To 2) As Double
    Dim dTo(0 To 2) As Double
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim strCommand As String
    
  'On Error Resume Next
    
    dLL(0) = vCoords(0)
    dLL(1) = vCoords(1)
    dLL(2) = vCoords(2)
    
    dUR(0) = vCoords(3)
    dUR(1) = vCoords(4)
    dUR(2) = vCoords(5)
    
    dTo(0) = vCoords(6)
    dTo(1) = vCoords(7)
    dTo(2) = vCoords(8)
    
    grpCode(0) = 8
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS4 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS4.Select acSelectionSetWindow, dLL, dUR, filterType, filterValue
    
    If objSS4.count = 0 Then GoTo Exit_Sub
    
    strCommand = "COPY" & vbCr & "P" & vbCr & vbCr
    strCommand = strCommand & dLL(0) & "," & dLL(1) & ",0" & vbCr
    strCommand = strCommand & dTo(0) & "," & dTo(1) & ",0" & vbCr & vbCr
    
    ThisDrawing.SendCommand strCommand
    
    
Exit_Sub:
    objSS4.Clear
    objSS4.Delete
End Sub

Private Sub FillOutDrawingList()
    Dim objSS3 As AcadSelectionSet
    Dim entBlock As AcadObject
    Dim obrTemp As AcadBlockReference
    Dim attItem2, vTemp As Variant
    Dim mode As Integer
    'Dim str As String
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim str, strList As String
    Dim strLayer As String
    
  On Error Resume Next
    lbSheets.Clear
    
    If cbDWG.Value = "DWG" Then
        strLayer = "Integrity Sheets"
    Else
        strLayer = "Integrity Permits-" & cbDWG.Value
    End If
    
    grpCode(0) = 8
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS3.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entBlock In objSS3
        If Not TypeOf entBlock Is AcadBlockReference Then GoTo Next_entBlock
        Set obrTemp = entBlock
        attItem2 = obrTemp.GetAttributes
        
        Select Case obrTemp.Name
            Case "SS Info"
                tbProject.Value = attItem2(0).TextString
                tbWO.Value = attItem2(1).TextString
                cbExchange.Value = attItem2(2).TextString
                tbRST.Value = attItem2(3).TextString
                cbCounty.Value = attItem2(4).TextString
                tbCity.Value = attItem2(5).TextString
                If tbTotalDWG.Value < attItem2(7).TextString Then tbTotalDWG.Value = attItem2(7).TextString
            Case "SS-11x17"
                vTemp = Split(attItem2(0).TextString, " ")
            
                'If Not vTemp(0) = cbDWG.Value Then GoTo Next_entBlock
               
                str = attItem2(0).TextString & vbTab & obrTemp.XScaleFactor
                str = str & vbTab & obrTemp.InsertionPoint(0) & vbTab & obrTemp.InsertionPoint(1)
                lbSheets.AddItem str
        End Select
Next_entBlock:
    Next entBlock
    
    Call SortLBItems
Exit_Sub:
    objSS3.Delete
End Sub

Private Sub UpdateAllDWGs()
    Dim objSS3 As AcadSelectionSet
    Dim entBlock As AcadObject
    Dim obrTemp As AcadBlockReference
    Dim attItem2 As Variant
    Dim mode As Integer
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim str As String
    
  On Error Resume Next
    
    grpCode(0) = 8
    If cbDWG.Value = "DWG" Then
        grpValue(0) = "Integrity Sheets"
    Else
        grpValue(0) = "Integrity Permits-" & cbDWG.Value
    End If
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS3.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entBlock In objSS3
        Set obrTemp = entBlock
        
        If obrTemp.Name = "SS Info" Then
            attItem2 = obrTemp.GetAttributes
            attItem2(0).TextString = tbProject.Value
            attItem2(1).TextString = tbWO.Value
            attItem2(2).TextString = cbExchange.Value
            attItem2(3).TextString = tbRST.Value
            attItem2(4).TextString = cbCounty.Value
            attItem2(5).TextString = tbCity.Value
            attItem2(7).TextString = tbTotalDWG.Value
            
            obrTemp.Update
        End If
    Next entBlock
    
    Call SortLBItems
Exit_Sub:
    objSS3.Delete
End Sub
