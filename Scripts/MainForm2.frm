VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm2 
   Caption         =   "Integrity Menu 2.0"
   ClientHeight    =   12150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18495
   OleObjectBlob   =   "MainForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim strUsageReport As String

Private Sub AddCLR_Click()
    Me.Hide
        Load AddClearances
        AddClearances.show
        Unload AddClearances
    Me.show
End Sub

Private Sub cbAddBldgs_Click()
    Me.Hide
        Load AddBLDGForm
        AddBLDGForm.show
        Unload AddBLDGForm
    Me.show
End Sub

Private Sub cbAddExistingGuys_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim objAnchor As AcadBlockReference
    Dim vCoords As Variant
    Dim dRotate, dCoords(2) As Double
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "ExGuyOR,ExGuyOL"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        dRotate = objBlock.Rotation
        vCoords = objBlock.InsertionPoint
        
        Select Case objBlock.Name
            Case "ExGuyOR"
                dRotate = dRotate - (3.14159265359 / 2)
            Case "ExGuyOL"
                dRotate = dRotate + (3.14159265359 / 2)
        End Select
        
        dCoords(0) = vCoords(0) + (15 * Cos(dRotate))
        dCoords(1) = vCoords(1) + (15 * Sin(dRotate))
        dCoords(2) = 0#
        
        Set objAnchor = ThisDrawing.ModelSpace.InsertBlock(dCoords, "Existing_Guys", 1#, 1#, 1#, objBlock.Rotation)
        objAnchor.Layer = "Integrity Existing-Guys"
        objAnchor.Update
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete

End Sub

Private Sub cbAddOHG_Click()
    Dim entPole As AcadObject
    Dim obrGP As AcadBlockReference
    Dim obrTemp As AcadBlockReference
    Dim attItem, attList, basePnt As Variant
    Dim lwpObj As AcadLWPolyline
    Dim lwpCoords(0 To 3) As Double
    Dim strFromPole, strToPole As String
    Dim vCoords As Variant
    Dim dFromCoords(0 To 2), dToCoords(0 To 2) As Double
    Dim dInsertionPnt(0 To 2) As Double
    Dim dTemp, Pi  As Double
    Dim dX, dY, dRotate, dScale, dDistance As Double
    Dim strBlock, strRUS, strTemp As String
    Dim layerObj As AcadLayer
    
    Me.Hide
  On Error Resume Next
    Pi = 3.14159265359
  
   dScale = CDbl(MainForm.cbScale.Value) / 100
   If Err <> 0 Then dScale = 1#
  
    ThisDrawing.Utility.GetEntity entPole, basePnt, "From Pole: "
    If TypeOf entPole Is AcadBlockReference Then
        Set obrGP = entPole
    Else
        MsgBox "Not a valid pole."
        Exit Sub
    End If
    
    dFromCoords(0) = obrGP.InsertionPoint(0)
    dFromCoords(1) = obrGP.InsertionPoint(1)
    dFromCoords(2) = 0#
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "To Pole: "
    If TypeOf entPole Is AcadBlockReference Then
        Set obrGP = entPole
    Else
        MsgBox "Not a valid pole."
        Exit Sub
    End If
    
    dToCoords(0) = obrGP.InsertionPoint(0)
    dToCoords(1) = obrGP.InsertionPoint(1)
    dToCoords(2) = 0#
    
    dInsertionPnt(0) = (dFromCoords(0) + dToCoords(0)) / 2
    dInsertionPnt(1) = (dFromCoords(1) + dToCoords(1)) / 2
    dInsertionPnt(2) = 0#
    
    Select Case (dToCoords(0) - dFromCoords(0))
        Case Is < 0
            dX = (dFromCoords(0) - dToCoords(0))
            dY = (dFromCoords(1) - dToCoords(1))
            dRotate = Atn(dY / dX) + Pi
            strBlock = "ohgL"
        Case Is > 0
            dX = (dToCoords(0) - dFromCoords(0))
            dY = (dToCoords(1) - dFromCoords(1))
            dRotate = Atn(dY / dX)
            strBlock = "ohgR"
        Case Else
            dRotate = 0
            strBlock = "ohgR"
    End Select
    
    dDistance = Sqr((dX * dX) + (dY * dY))
    dDistance = Round(dDistance)
    
    Set layerObj = ThisDrawing.Layers.Add("Integrity Guys")
    ThisDrawing.ActiveLayer = layerObj
    
    Set obrTemp = ThisDrawing.ModelSpace.InsertBlock(dInsertionPnt, strBlock, dScale, dScale, dScale, dRotate)

    attList = obrTemp.GetAttributes
    
    attList(0).TextString = dDistance & "'"
    attList(1).TextString = "PE2-3G"
    
    lwpCoords(0) = dFromCoords(0)
    lwpCoords(1) = dFromCoords(1)
    lwpCoords(2) = dToCoords(0)
    lwpCoords(3) = dToCoords(1)
    
    Set lwpObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
    lwpObj.Update
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, vbCr & "Select block to add units:"
    
    Set obrTemp = entPole
    
    If obrTemp.Name = "sPole" Then
        attList = obrTemp.GetAttributes
            
        If attList(27).TextString = "" Then
            attList(27).TextString = "+PE2-3G=" & dDistance & "'"
        Else
            attList(27).TextString = attList(27).TextString & ";;+PE2-3G=" & dDistance & "'"
        End If
        
        obrTemp.Update
    End If
    
    Me.show
End Sub

Private Sub cbAddSheets_Click()
    'Dim layerObj As AcadLayer
  'On Error Resume Next
    'Set layerObj = ThisDrawing.Layers.Add("Integrity Sheets")
    'ThisDrawing.ActiveLayer = layerObj
    
    Dim entPole As AcadObject
    Dim obrGSS As AcadBlockReference
    Dim obrSSInfo As AcadBlockReference
    Dim obrSSDWG1 As AcadBlockReference
    Dim attItem, attItem1, basePnt, returnPnt As Variant
    Dim vEmail As Variant
    Dim insertPnt(0 To 2) As Double
    Dim dScale As Double
    Dim iDWGNum, iNextNum As Long
    Dim str, str2 As String
    
    Me.Hide
  On Error Resume Next
  
    iNextNum = 1
    
While Err = 0
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Staking Sheet: "
    If TypeOf entPole Is AcadBlockReference Then
        Set obrGSS = entPole
    Else
        GoTo Exit_Sub
        'Me.show
        'Exit Sub
    End If
    
    If Not Err = 0 Then GoTo Exit_Sub
    If Not obrGSS.Name = "SS-11x17" Then GoTo Exit_Sub
    
    obrGSS.Layer = "Integrity Sheets"
    attItem = obrGSS.GetAttributes
    iDWGNum = ThisDrawing.Utility.GetInteger(vbCr & "Enter Drawing Number(" & iNextNum & "): ")
    If Err < 0 Then iDWGNum = iNextNum
    iNextNum = iDWGNum + 1
    
    Select Case iDWGNum
        Case Is < 10
            attItem(0).TextString = "DWG 00" & iDWGNum
        Case Is > 99
            attItem(0).TextString = "DWG " & iDWGNum
        Case Else
            attItem(0).TextString = "DWG 0" & iDWGNum
    End Select
    
    obrGSS.Update
    dScale = obrGSS.XScaleFactor
    
    lbSheets.AddItem iDWGNum & vbTab & dScale & vbTab & obrGSS.Name & vbTab & obrGSS.InsertionPoint(0) & vbTab & obrGSS.InsertionPoint(1)
  
    insertPnt(0) = obrGSS.InsertionPoint(0) + 20 * dScale
    insertPnt(1) = obrGSS.InsertionPoint(1) + 20 * dScale
    insertPnt(2) = obrGSS.InsertionPoint(2)
    
    Set obrSSInfo = ThisDrawing.ModelSpace.InsertBlock(insertPnt, "ss info", dScale, dScale, dScale, 0#)
    obrSSInfo.Layer = "Integrity Sheets"
    
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
        obrSSDWG1.Layer = "Integrity Sheets"
    
        attItem1 = obrSSDWG1.GetAttributes
        
        vEmail = Split(cbDesigner.Value, " ")
        str2 = LCase(vEmail(0)) & "." & LCase(vEmail(1)) & "@integrity-us.com"
        str = cbDesigner.Value & "\P" & str2 & "\P" & tbPhoneNumber.Value & "\P"
        'Select Case cbDesigner.Value
        '    Case "Dylan Spears"
        '        str = "Dylan Spears\Pdylan.spears@integrity-us.com\P" & tbPhoneNumber.Value & "\P"
        '    Case "Jeremy Pafford"
        '        str = "Jeremy Pafford\Pjeremy.pafford@integrity-us.com\P(931)698-1992\P"
        '    Case "Ronn Elliott"
        '        str = "Ronn Elliott\Pronn.elliott@integrity-us.com\P(615)419-5421\P"
        '    Case "Rich Taylor"
        '        str = "Rich Taylor\Prich.taylor@integrity-us.com\P(615)785-2032\P"
        '    Case "Ronn Elliott"
        '        str = "Jason Pafford\Pjason.pafford@integrity-us.com\P(931)209-3269\P"
        'End Select
        str = str & "730 Middle Tennessee Blvd - Suite 6\PMurfreesboro, TN 37130"
        attItem1(0).TextString = str
        obrSSDWG1.Update
    End If
    Err = 0
Wend
    layerObj.Delete
    
Exit_Sub:
    Set layerObj = ThisDrawing.Layers.Add("0")
    ThisDrawing.ActiveLayer = layerObj
    Me.show
End Sub

Private Sub cbAddTerminals_Click()
    Me.Hide
    Load DesignTermForm
    DesignTermForm.show
    Unload DesignTermForm
    Me.show
End Sub

Private Sub cbAddUnits_Click()
    Me.Hide
        Load AddRemoveUnits
        AddRemoveUnits.show
        Unload AddRemoveUnits
    Me.show
End Sub

Private Sub cbAnchorGuys_Click()
    Dim startPnt(0 To 2), midPnt(0 To 2), endPnt(0 To 2) As Double
    Dim firstAng, endAng, defAng, rotateAng, ancAng, pull, tempPull As Double
    Dim X(0 To 2), Y(0 To 2), Z(0 To 2) As Double
    Dim dScale As Double
    Dim objBlock As AcadBlockReference
    Dim objBlock2 As AcadBlockReference
    Dim objTemp As AcadBlockReference
    Dim returnPoint, vAttList, insertPnt As Variant
    Dim n, result, iDirection As Integer  'iDirection 0=L 1=R
    Dim str, str2 As String
    'Dim layerObj As AcadLayer
    Dim strPull, strResult As String
    Dim vPoleCoords As Variant
    Dim objEntity As AcadEntity
    'Dim obrGP As AcadBlockReference
    Dim attItem, basePnt As Variant
    Dim Pi As Double
    
    Pi = 3.14159265359
    Me.Hide
On Error Resume Next

Place_Anchor:
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, basePnt, "Select Pole: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
    Else
        'MsgBox "Not a valid pole."
        Me.show
        Exit Sub
    End If
    vPoleCoords = objBlock.InsertionPoint
    
    X(0) = vPoleCoords(0)
    Y(0) = vPoleCoords(1)
    
    ThisDrawing.Utility.GetEntity objEntity, basePnt, "Get Backspan Pole: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock2 = objEntity
    Else
        'MsgBox "Not a valid pole."
        Me.show
        Exit Sub
    End If
    vPoleCoords = objBlock2.InsertionPoint
    
    X(1) = vPoleCoords(0)
    Y(1) = vPoleCoords(1)
    
    ThisDrawing.Utility.GetEntity objEntity, basePnt, "Get Frontspan Pole: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock2 = objEntity
    Else
        'MsgBox "Not a valid pole."
        Me.show
        Exit Sub
    End If
    vPoleCoords = objBlock2.InsertionPoint
    
    
    X(2) = vPoleCoords(0)
    Y(2) = vPoleCoords(1)
    
    If (X(1) - X(0)) = 0 Then
        If (Y(1) - Y(0)) < 0 Then
            firstAng = 3 * Pi / 2
        Else
            firstAng = Pi / 2
        End If
    Else
        firstAng = Math.Atn((Y(1) - Y(0)) / (X(1) - X(0)))
        If (X(1) - X(0)) < 0 Then
            firstAng = Pi + firstAng
        ElseIf (Y(1) - Y(0)) < 0 Then
            firstAng = 2 * Pi + firstAng
        End If
    End If
    
    If (X(2) - X(0)) = 0 Then
        If (Y(2) - Y(0)) < 0 Then
            firstAng = 3 * Pi / 2
        Else
            firstAng = Pi / 2
        End If
    Else
        endAng = Math.Atn((Y(2) - Y(0)) / (X(2) - X(0)))
        If (X(2) - X(0)) < 0 Then
            endAng = Pi + endAng
        ElseIf (Y(2) - Y(0)) < 0 Then
            endAng = 2 * Pi + endAng
        End If
    End If
    
    rotateAng = (endAng + firstAng) / 2
    If Math.Abs(endAng - firstAng) < Pi Then
        rotateAng = rotateAng - Pi
    End If
        
    If rotateAng < 0 Then rotateAng = rotateAng + (2 * Pi)
    If rotateAng > (2 * Pi) Then rotateAng = rotateAng - (2 * Pi)
    
    defAng = firstAng + Pi - endAng
    pull = Math.Round(Math.Abs(100 * Math.Sin(defAng / 2)))
    
    If pull >= 50 Then  '<------------------------------------- Get strPull
        strPull = "DE"
    Else
        strPull = pull
    End If
    
    If (rotateAng > (Pi / 2)) And (rotateAng <= (3 * Pi / 2)) Then
        iDirection = 0
    Else
        iDirection = 1
    End If
    
    insertPnt = ThisDrawing.Utility.GetPoint(, "Place Guy:")
    
    dScale = 1#
    'dScale = CDbl(MainForm.cbScale.Value) / 100
    
    If iDirection = 0 Then
        str = "ExGuyOL"
    Else
        str = "ExGuyOR"
    End If
    
    'Set layerObj = ThisDrawing.Layers.Add("Integrity Guys")
    'ThisDrawing.ActiveLayer = layerObj
    
    Err = 0
    
    Set objTemp = ThisDrawing.ModelSpace.InsertBlock(insertPnt, str, dScale, dScale, dScale, rotateAng)
    objTemp.Layer = "Integrity Proposed"
    vAttList = objTemp.GetAttributes
   
    vAttList(1).TextString = " "
    vAttList(2).TextString = "PE1-3G"
    objTemp.Update
    If Not Err = 0 Then MsgBox "error" & vbCr & Err.Description

    If iDirection = 0 Then
        str = "ExAncOL"
    Else
        str = "ExAncOR"
    End If

    Set objTemp = ThisDrawing.ModelSpace.InsertBlock(insertPnt, str, dScale, dScale, dScale, rotateAng)
    objTemp.Layer = "Integrity Proposed"
    vAttList = objTemp.GetAttributes
    
    vAttList(0).TextString = "PF1-5A"
    'If strPull = "DE" Then
        'vAttList(1).TextString = "P= DE"
    'Else
        vAttList(1).TextString = "P= " & strPull
    'End If
    objTemp.Update
    
    vAttList = objBlock.GetAttributes
    
    If vAttList(27).TextString = "" Then
        vAttList(27).TextString = "PE1-3G=1;;PF1-5A=1;;PM11=1"
    Else
        vAttList(27).TextString = vAttList(27).TextString & ";;+PE1-3G=1;;+PF1-5A=1;;+PM11=1"
    End If
    
    objBlock.Update
    
    GoTo Place_Anchor
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbAttachAlias_Click()
    Me.Hide
        Load AttachmentAlias
            AttachmentAlias.show
        Unload AttachmentAlias
    Me.show
End Sub

Private Sub cbAVCounts_Click()
    Me.Hide
        Load CountsForm
            CountsForm.show
        Unload CountsForm
    Me.show
End Sub

Private Sub cbBlockAtt_Click()
    Me.Hide
        Load ChangeAttData
        ChangeAttData.show
        Unload ChangeAttData
    Me.show
End Sub

Private Sub cbBore_Click()
    Me.Hide
        Load BuriedPlantForm
        BuriedPlantForm.show
        Unload BuriedPlantForm
    Me.show
End Sub

Private Sub cbCableCO_Click()
    Me.Hide
        Load CableCalloutForm
            CableCalloutForm.show
        Unload CableCalloutForm
    Me.show
End Sub

Private Sub cbCalcMR_Click()
    Me.Hide
    Load MRWorksheet
    MRWorksheet.show
    Unload MRWorksheet
    Me.show
End Sub

Private Sub cbCleanup_Click()
    Me.Hide
    Load CleanupForm
    CleanupForm.show
    Unload CleanupForm
    Me.show
End Sub

Private Sub cbClosureLocations_Click()
    Me.Hide
        Load ClosureLocations
        ClosureLocations.show
        Unload ClosureLocations
    Me.show
End Sub

Private Sub cbConvertBlocksBlock_Click()
    Me.Hide
        Load ConvertChangeBlocks
        ConvertChangeBlocks.show
        Unload ConvertChangeBlocks
    Me.show
End Sub

Private Sub cbConvertCustomer_Click()
    Me.Hide
        Load ConvertCustomers
            ConvertCustomers.show
        Unload ConvertCustomers
    Me.show
End Sub

Private Sub cbConvertGPS_Click()
    Me.Hide
        Load ConvertGPS
        ConvertGPS.show
        Unload ConvertGPS
    Me.show
End Sub

Private Sub cbConvertLayers_Click()
    Me.Hide
        Load ConvertBlockLayers
        ConvertBlockLayers.show
        Unload ConvertBlockLayers
    Me.show
End Sub

Private Sub cbConvertMisc_Click()
    Me.Hide
        Load ConvertBlocks
        ConvertBlocks.show
        Unload ConvertBlocks
    Me.show
End Sub

Private Sub cbConvertNES_Click()
    Me.Hide
        Load CreateNESPoles
        CreateNESPoles.show
        Unload CreateNESPoles
    Me.show
End Sub

Private Sub cbConvertText_Click()
    Me.Hide
        Load ConvertBlockText
        ConvertBlockText.show
        Unload ConvertBlockText
    Me.show
End Sub

Private Sub cbCustomers_Click()
'    Dim SSobj2 As AcadSelectionSet
'    Dim objBLDG As AcadBlockReference
'    Dim objNewBLDG As AcadBlockReference
'    Dim attList, attNewList As Variant
'    Dim filterType, filterValue As Variant
'    Dim grpCode(0) As Integer
'    Dim grpValue(0) As Variant
'    Dim entBlock As AcadEntity
'    Dim iCount As Integer
'    Dim str1, str2, strBlock, strLayer As String
'    Dim strTemp As String
'    Dim dInsertPnt(0 To 2) As Double
'    Dim dScale As Double
'
'    Me.hide
'    iCount = 0
'
'  On Error Resume Next
'
'    strTemp = "Integrity Building-BUS"
'    'strTemp = "Integrity Building, Integrity Building-RES, Integrity Building-BUS, Integrity Building-MDU, "
'    'strTemp = strTemp & "Integrity Building-TRL, Integrity Building-SCH, Integrity Building-CHU"
'
'    grpCode(0) = 8
'    grpValue(0) = strTemp
'    'grpCode(1) = 8
'    'grpValue(1) = "Integrity Building-RES"
'    filterType = grpCode
'    filterValue = grpValue
'
'    Set SSobj2 = ThisDrawing.SelectionSets.Add("SSobj2")
'    SSobj2.Select acSelectionSetAll, , , filterType, filterValue
'
'    MsgBox SSobj2.count
'
'    For Each entBlock In SSobj2
'        If Not entBlock.ObjectName = "AcDbBlockReference" Then GoTo Next_entBlock
'        Set objBLDG = entBlock
'        attList = objBLDG.GetAttributes
'        If Not UBound(attList) = 2 Then GoTo Next_entBlock
'
'        str1 = Left(attList(2).TextString, 3)
'        If str1 = "" Then GoTo Next_entBlock
'
'        str2 = UCase(attList(2).TextString)
'        If Len(str2) > 3 Then str2 = Right(str2, Len(str2) - 3)
'        If Len(str2) = 0 Then str2 = " "
'        If Left(str2, 1) = " " Then str2 = Right(str2, Len(str2) - 1)
'
'        Select Case str1
'            Case "res"
'                If objBLDG.Name = "RES" Then GoTo Next_entBlock
'                strBlock = "RES"
'                strLayer = "Integrity Building-RES"
'            Case "bus"
'                If objBLDG.Name = "BUSINESS" Then GoTo Next_entBlock
'                strBlock = "BUSINESS"
'                strLayer = "Integrity Building-BUS"
'            Case "trl"
'                If objBLDG.Name = "TRLR" Then GoTo Next_entBlock
'                strBlock = "TRLR"
'                strLayer = "Integrity Building-TRL"
'            Case "mdu"
'                If objBLDG.Name = "MDU" Then GoTo Next_entBlock
'                strBlock = "MDU"
'                strLayer = "Integrity Building-MDU"
'            Case "ext"
'                If objBLDG.Name = "EXTENTION" Then GoTo Next_entBlock
'                strBlock = "EXTENTION"
'                strLayer = "Integrity Building Misc"
'            Case "sch"
'                If objBLDG.Name = "SCHOOL" Then GoTo Next_entBlock
'                strBlock = "SCHOOL"
'                strLayer = "Integrity Building-SCH"
'            Case "chu"
'                MsgBox "CHURCH" & vbCr & str1 & vbCr & str2
'                If objBLDG.Name = "CHURCH" Then GoTo Next_entBlock
'                strBlock = "CHURCH"
'                strLayer = "Integrity Building-CHU"
'            Case Else
'                GoTo Next_entBlock
'        End Select
'
'        'str2 = "Old Block: " & objBLDG.Name & vbCr & "Data: " & attList(2).TextString & vbCr
'        'str2 = str2 & "New Block: " & strBlock
'        'MsgBox str2
'
'        'str2 = UCase(attList(2).TextString)
'        'str2 = Right(str2, Len(str2) - 3)
'
'        dInsertPnt(0) = objBLDG.InsertionPoint(0)
'        dInsertPnt(1) = objBLDG.InsertionPoint(1)
'        dInsertPnt(2) = 0#
'
'        dScale = objBLDG.XScaleFactor
'
'        Set objNewBLDG = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, strBlock, dScale, dScale, dScale, 0#)
'        objNewBLDG.Layer = strLayer
'        attNewList = objNewBLDG.GetAttributes
'
'        attNewList(0).TextString = attList(0).TextString
'        attNewList(1).TextString = attList(1).TextString
'        attNewList(2).TextString = str2
'
'        objNewBLDG.Update
'        objBLDG.Delete
'        iCount = iCount + 1
'Next_entBlock:
'    Next entBlock
'
'    SSobj2.Clear
'    SSobj2.Delete
'
'    MsgBox "Converted Buildings: " & iCount
'
'    Me.show
End Sub

Private Sub cbDBKF_Click()
    Me.Hide
        Load dbLoss
            dbLoss.show
        Unload dbLoss
    Me.show
End Sub

Private Sub cbDeleteDrop_Click()
    Me.Hide
        Call DeleteDropCallouts
    Me.show
End Sub

Private Sub cbDropCallouts_Click()
    Me.Hide
        Call DropCallout
    Me.show
End Sub

Private Sub cbExchange_Change()
    Select Case cbExchange.Value
        Case "Belfast"
            cbCounty.Value = "Marshall"
            tbCity.Value = "Belfast"
        Case "Bell Buckle"
            cbCounty.Value = "Bedford"
            tbCity.Value = "Bell Buckle"
        Case "Chapel Hill"
            cbCounty.Value = "Marshall"
            tbCity.Value = "Chapel Hill"
        Case "Chapel Hill-CLEC"
            cbCounty.Value = "Marshall"
            tbCity.Value = " "
        Case "College Grove"
            cbCounty.Value = "Williamson"
            tbCity.Value = "College Grove"
        Case "College Grove-CLEC"
            cbCounty.Value = "Williamson"
            tbCity.Value = " "
        Case "Flat Creek"
            cbCounty.Value = "Bedford"
            tbCity.Value = "Flat Creek"
        Case "Fosterville"
            cbCounty.Value = "Rutherford"
            tbCity.Value = "Fosterville"
        Case "Franklin-CLEC"
            cbCounty.Value = "Williamson"
            tbCity.Value = "Franklin"
        Case "Lebanon-CLEC"
            cbCounty.Value = "Wilson"
            tbCity.Value = "Lebanon"
        Case "Murfreesboro-CLEC"
            cbCounty.Value = "Rutherford"
            tbCity.Value = "Murfreesboro"
        Case "Nolensville"
            cbCounty.Value = "Williamson"
            tbCity.Value = "Nolensville"
        Case "Nolensville-CLEC"
            cbCounty.Value = "Williamson"
            tbCity.Value = "Nolensville"
        Case "Shelbyville-CLEC"
            cbCounty.Value = "Bedford"
            tbCity.Value = "Shelbyville"
        Case "Smyrna-CLEC"
            cbCounty.Value = "Rutherford"
            tbCity.Value = "Smyrna"
        Case "Triune"
            cbCounty.Value = "Williamson"
            tbCity.Value = "Triune"
        Case "Unionville"
            cbCounty.Value = "Bedford"
            tbCity.Value = "Unionville"
    End Select
End Sub

Private Sub cbExtraHeights_Click()
    Me.Hide
        Load ExtraHeightForm
        ExtraHeightForm.show
        Unload ExtraHeightForm
    Me.show
End Sub

Private Sub cbGetLengths_Click()
    Me.Hide
        Load GetCableLengths
        GetCableLengths.show
        Unload GetCableLengths
    Me.show
End Sub

Private Sub cbGetSpansReel_Click()
    Me.Hide
        Load GetSpansReel
        GetSpansReel.show
        Unload GetSpansReel
    Me.show
End Sub

Private Sub cbHousing_Click()
    Me.Hide
        Load PlaceBuriedData
        PlaceBuriedData.show
        Unload PlaceBuriedData
    Me.show
End Sub

Private Sub cbLoadBlocks_Click()
    Me.Hide
        Load LoadBlocks
        LoadBlocks.show
        Unload LoadBlocks
    Me.show
End Sub

Private Sub cbMaptrim_Click()
    Dim objSSM1 As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objEntTemp As AcadEntity
    Dim objRemove(0) As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objLWPBorder As AcadLWPolyline
    Dim objSSBlock As AcadBlockReference
    Dim objLayer As AcadLayer
    'Dim strLayer As String
    Dim strCommand As String
    Dim strHandle As String
    Dim strRemove As String
    Dim vReturnPnt As Variant
    Dim vCoords As Variant
    Dim dLL(0 To 2) As Double
    Dim dUR(0 To 2) As Double
    Dim dFrom(0 To 2) As Double
    Dim dTo(0 To 2) As Double
    Dim dDiff(0 To 1) As Double
    Dim dScale As Double
    
  On Error Resume Next
    'Select boundary & find drawing number
    
    Me.Hide
    
    Err = 0
    Set objSSM1 = ThisDrawing.SelectionSets.Add("objSSM1")
    If Not Err = 0 Then
        Set objSSM1 = ThisDrawing.SelectionSets.Item("objSSM1")
        Err = 0
    End If
    
  While Err = 0
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Trim Border: "
    If Not objEntity.ObjectName = "AcDbPolyline" Then
        Me.show
        Exit Sub
    End If
    
    Set objLWP = objEntity
    vCoords = objLWP.Coordinates
    
    dLL(0) = vCoords(0)
    dLL(1) = vCoords(1)
    dUR(0) = vCoords(0)
    dUR(1) = vCoords(1)
    
    For i = 2 To UBound(vCoords)
        If vCoords(i) < dLL(0) Then
            dLL(0) = vCoords(i)
        Else
            If vCoords(i) > dUR(0) Then
                dUR(0) = vCoords(i)
            End If
        End If
        
        i = i + 1
        
        If vCoords(i) < dLL(1) Then
            dLL(1) = vCoords(i)
        Else
            If vCoords(i) > dUR(1) Then
                dUR(1) = vCoords(i)
            End If
        End If
    Next i
    
    dFrom(0) = (dLL(0) + dUR(0)) / 2
    dFrom(1) = (dLL(1) + dUR(1)) / 2
    
    'Get coordinates to DWG Sheet and draw LWP (Save strHandle)
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select DWG Border: "
    If TypeOf objEntity Is AcadBlockReference Then
        Set objSSBlock = objEntity
    Else
        MsgBox "Not a valid entity."
        Exit Sub
    End If
    
    dScale = objSSBlock.XScaleFactor
    strRemove = objSSBlock.Layer
    
    dTo(0) = objSSBlock.InsertionPoint(0) + (825 * dScale)
    dTo(1) = objSSBlock.InsertionPoint(1) + (525 * dScale)
    dTo(2) = 0#
    
    dDiff(0) = dTo(0) - dFrom(0)
    dDiff(1) = dTo(1) - dFrom(1)
    
    Set objLWPBorder = ThisDrawing.ModelSpace.AddLightWeightPolyline(vCoords)
    objLWPBorder.Closed = True
    objLWPBorder.Layer = objLWP.Layer
    objLWPBorder.Move dFrom, dTo
    objLWPBorder.Update
    strHandle = objLWPBorder.Handle
    
    'Get objects crossing within dimensions of DWG and copy to DWG
    
    objSSM1.Select acSelectionSetCrossing, dLL, dUR
    
    If objSSM1.count = 0 Then GoTo Exit_Sub
    
    strCommand = "COPY" & vbCr & "P" & vbCr & vbCr
    strCommand = strCommand & dFrom(0) & "," & dFrom(1) & ",0" & vbCr
    strCommand = strCommand & dTo(0) & "," & dTo(1) & ",0" & vbCr & vbCr
    
    ThisDrawing.SetVariable "CMDDIA", 0
    ThisDrawing.SendCommand strCommand
    
    objSSM1.Clear
    
    Set objLayer = ThisDrawing.Layers(strRemove)
    objLayer.Lock = True
    
    dLL(0) = dLL(0) + dDiff(0)
    dLL(1) = dLL(1) + dDiff(1)
    dUR(0) = dUR(0) + dDiff(0)
    dUR(1) = dUR(1) + dDiff(1)
    
    'Set objSSM1 = ThisDrawing.SelectionSets.Add("objSSM1")
    objSSM1.Select acSelectionSetCrossing, dLL, dUR
        
    For Each objEntTemp In objSSM1
        If objEntTemp.Layer = strRemove Then
            'MsgBox "Found an entity" & vbCr & objEntTemp.ObjectName
            Set objRemove(0) = objEntTemp
            objSSM1.RemoveItems objRemove
        End If
    Next objEntTemp
    
    strCommand = "_MAPTRIM" & vbCr & "S" & vbCr & "(handent """ & strHandle & """)" & vbCr
    strCommand = strCommand & "N" & vbCr & "Y" & vbCr & "P" & vbCr & vbCr & "O" & vbCr
    strCommand = strCommand & "Y" & vbCr & "Y" & vbCr & "R" & vbCr & "Y" & vbCr

    ThisDrawing.SendCommand strCommand
    
    objSSM1.Clear
    
    objLayer.Lock = False
  Wend
    
Exit_Sub:
    objSSM1.Delete
    
    ThisDrawing.SetVariable "CMDDIA", 1
    Me.show
End Sub

Private Sub cbMatchlines_Click()
    Me.Hide
    Call Module1.AddMatchlines
    Me.show
    Exit Sub
End Sub

Private Sub cbMissingCallouts_Click()
    Me.Hide
        Load FindMissingCallouts
            FindMissingCallouts.show
        Unload FindMissingCallouts
    Me.show
End Sub

Private Sub cbMRReport_Click()
    Me.Hide
        Load MRSheets
            MRSheets.show
        Unload MRSheets
    Me.show
End Sub

Private Sub cbOldMenu_Click()
    Me.Hide
        Load MainForm2
        MainForm2.show
        Unload MainForm2
    Me.show
End Sub

Private Sub cbMRReview_Click()
    Me.Hide
    Load MRReview
    MRReview.show
    Unload MRReview
    Me.show
End Sub

Private Sub cbODInquiry_Click()
    Me.Hide
        Load ObjectDataInquiry
        ObjectDataInquiry.show
        Unload ObjectDataInquiry
    Me.show
End Sub

Private Sub cbPlaceBuriedCbl_Click()
    Dim amap As AcadMap
    Dim ODRcs As ODRecords
    Dim tbl As ODTable
    Dim tbls As ODTables
    Dim boolVal As Boolean
    Dim entPole As AcadObject
    Dim objLWP As AcadLWPolyline
    Dim objLine As AcadLine
    Dim obrDrop As AcadBlockReference
    Dim attItem, basePnt As Variant
    Dim strType, strLength, strTemp As String
    Dim dRotate, dScale, Pi As Double
    Dim dTest As Double
    Dim dInsertPnt(0 To 2) As Double
    Dim dStartPnt(0 To 2) As Double
    Dim dEndPnt(0 To 2) As Double
    Dim dDiff(0 To 1) As Double
    Dim vStartPnt, vEndPnt As Variant
    Dim vCoords, vAttList, vList As Variant
    Dim layerObj As AcadLayer
    Dim strCableSize As String
    
    Me.Hide
  On Error Resume Next
    Set layerObj = ThisDrawing.Layers.Add("Integrity Cable-Buried Text")
    ThisDrawing.ActiveLayer = layerObj
    
    strCableSize = ThisDrawing.Utility.GetString(1, "Cable Size (Space between Cables): ")
    vList = Split(strCableSize, " ")
    strCableSize = ""
    For iTemp = LBound(vList) To UBound(vList)
        strCableSize = strCableSize & vList(iTemp) & "F "
    Next iTemp
  
    Pi = 3.14159265359
    'dScale = 1
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
    Set tbls = amap.Projects(ThisDrawing).ODTables
    
get_Drop:

    dScale = 1#  'cbScale.Value / 100
    strLength = ""
    
    Err = 0
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Buried Cable: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    Select Case entPole.ObjectName
        Case "AcDbLine"
            Set objLine = entPole
            strTemp = CInt(objLine.Length)
            vStartPnt = objLine.StartPoint
            vEndPnt = objLine.EndPoint
            
            Err = 0
            If tbls.count > 0 Then
                For Each tbl In tbls
                    If tbl.Name = "UGPrimary" Then GoTo Exit_For2
                Next
            End If
Exit_For2:
            
            Set ODRcs = tbl.GetODRecords
            boolVal = ODRcs.Init(entPole, True, False)
             
            strLength = ODRcs.Record.Item(35).Value
            'MsgBox Len(strLength) & vbCr & Err
            If Len(strLength) < 1 Then
                strLength = strTemp
            End If
            
        Case "AcDbPolyline"
            Set objLWP = entPole
            strTemp = CInt(objLWP.Length)
            MsgBox strTemp
            vCoords = objLWP.Coordinates
            dStartPnt(0) = vCoords(0)
            dStartPnt(1) = vCoords(1)
            dStartPnt(2) = 0#
            dEndPnt(0) = vCoords(2)
            dEndPnt(1) = vCoords(3)
            dEndPnt(2) = 0#
            
            dDiff(0) = dEndPnt(0) - dStartPnt(0)
            dDiff(1) = dEndPnt(1) - dStartPnt(1)
            dTemp = Sqr(dDiff(0) * dDiff(0) + dDiff(1) * dDiff(1))
            
            If UBound(vCoords) > 3 Then
                For i = 5 To UBound(vCoords)
                    dDiff(0) = vCoords(i - 1) - vCoords(i - 3)
                    dDiff(1) = vCoords(i) - vCoords(i - 2)
                    If dTemp < Sqr(dDiff(0) * dDiff(0) + dDiff(1) * dDiff(1)) Then
                        dStartPnt(0) = vCoords(i - 3)
                        dStartPnt(1) = vCoords(i - 2)
                        dEndPnt(0) = vCoords(i - 1)
                        dEndPnt(1) = vCoords(i)
                        dTemp = Sqr(dDiff(0) * dDiff(0) + dDiff(1) * dDiff(1))
                    End If
                Next i
            End If
            
            Err = 0
            If tbls.count > 0 Then
                For Each tbl In tbls
                    If tbl.Name = "UGPrimary" Then GoTo exit_for4
                Next
            End If
exit_for4:
            
            Set ODRcs = tbl.GetODRecords
            boolVal = ODRcs.Init(entPole, True, False)
               
            strLength = ODRcs.Record.Item(35).Value
            'MsgBox Len(strLength) & vbCr & Err
            If Len(strLength) < 1 Then
                strLength = strTemp
            End If
            
            vStartPnt = dStartPnt
            vEndPnt = dEndPnt
        Case Else
            Set layerObj = ThisDrawing.Layers.Add("0")
            ThisDrawing.ActiveLayer = layerObj
            Me.show
            Exit Sub
    End Select
    
    dInsertPnt(0) = (vStartPnt(0) + vEndPnt(0)) / 2
    dInsertPnt(1) = (vStartPnt(1) + vEndPnt(1)) / 2
    dInsertPnt(2) = 0#
    
    dDiff(0) = vEndPnt(0) - vStartPnt(0)
    dDiff(1) = vEndPnt(1) - vStartPnt(1)
    
    If dDiff(0) = 0 Then
        dRotate = 0#
    Else
        If dDiff(1) = 0 Then
            dRotate = Pi / 2
        Else
            dRotate = Atn(dDiff(1) / dDiff(0))
        End If
    End If
    
    If dRotate > (Pi / 2) And dRotate < (1.5 * Pi) Then dRotate = dRotate - Pi
    
    'ThisDrawing.ActiveLayer = layerObj
    
    Set obrDrop = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "Cable_span", dScale, dScale, dScale, dRotate)
    vAttList = obrDrop.GetAttributes
    vAttList(1).TextString = strCableSize
    vAttList(2).TextString = strLength & "'"
    obrDrop.Update
    
    GoTo get_Drop
Exit_Sub:
    Me.show
End Sub

Private Sub cbPlaceCable_Click()
    Me.Hide
    
        Load PlaceAerialCable
        PlaceAerialCable.show
        Unload PlaceAerialCable
    
    Me.show
End Sub

Private Sub cbPlaceFromSpida_Click()
    Me.Hide
        Load PlaceSPIDAPoles
        PlaceSPIDAPoles.show
        Unload PlaceSPIDAPoles
    Me.show
End Sub

Private Sub cbPlaceOwner_Click()
    Me.Hide
            Load PoleOwnerNote
            PoleOwnerNote.show
            Unload PoleOwnerNote
    Me.show
End Sub

Private Sub cbPlacePoleData_Click()
    Me.Hide
            Load PlacePoleData
            PlacePoleData.show
            Unload PlacePoleData
    Me.show
End Sub

Private Sub cbPonVHLE_Click()
    Me.Hide
            Load PonVHLE
            PonVHLE.show
            Unload PonVHLE
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
    Dim strCommand As String
    Dim strPath As String
    Dim strFileName As String
    'Dim str As String
    Dim strArray As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    strPath = ThisDrawing.Path & "\"
    
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
    strPlot(14) = "United.ctb"
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
            strPlot(20) = strFileName
            'MsgBox strFileName
            
            strCommand = ""
            For j = 0 To 19
                If strPlot(j) = "" Then
                    strCommand = strCommand & vbCr
                Else
                    strCommand = strCommand & strPlot(j) & vbCr
                End If
            Next j
            
            ThisDrawing.SendCommand strCommand
            
            'str1 = strDWG & vbCr & "Scale= " & dScale & vbCr
            'str1 = str1 & "LLx " & dLL(0) & vbCr & "LLy " & dLL(1) & vbCr
            'str1 = str1 & "URx " & dUR(0) & vbCr & "URy " & dUR(1)
            
            'MsgBox str1
            
            'ThisDrawing.ActiveLayout.SetWindowToPlot dLL, dUR
            'ThisDrawing.ActiveLayout.PlotType = acWindow
            'ThisDrawing.ActiveLayout.ConfigName "Microsoft Print to PDF.pc3"
            'ThisDrawing.Plot.PlotToDevice
        End If
    Next i
    
    ThisDrawing.SetVariable "FILEDIA", 1
    ThisDrawing.SetVariable "CMDDIA", 1
    
End Sub

Private Sub cbProjectStatus_Click()
    Me.Hide
            Load ProjectStatus
            ProjectStatus.show
            Unload ProjectStatus
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRearrageCounts_Click()
    Me.Hide
        Load RearrangeCounts
            RearrangeCounts.show
        Unload RearrangeCounts
    Me.show
End Sub

Private Sub cbRenumber_Click()
    Me.Hide
        Load RenumberPoles
            RenumberPoles.show
        Unload RenumberPoles
    Me.show
End Sub

Private Sub cbReplaceCounts_Click()
    Me.Hide
        Load ReplaceCounts
            ReplaceCounts.show
        Unload ReplaceCounts
    Me.show
End Sub

Private Sub cbResizeBlocks_Click()
    Me.Hide
        Load ResizeBlocks
            ResizeBlocks.show
        Unload ResizeBlocks
    Me.show
End Sub

Private Sub cbSave_Click()
    ThisDrawing.SendCommand "QSAVE" & vbCr
End Sub

Private Sub cbScope_Click()
    Me.Hide
            Load ScopeOfWork
            ScopeOfWork.show
            Unload ScopeOfWork
    Me.show
End Sub

Private Sub cbSendCommand_Click()
    Dim strCommand As String
    
    Me.Hide
    
    strCommand = ThisDrawing.Utility.GetString(True, "Enter AutoCAD Command or Alias: ")
    
    ThisDrawing.SendCommand strCommand & vbCr
    
    Me.show
End Sub

Private Sub cbSpidaminSearch_Click()
    Me.Hide
        Load SPIDAminSearch
            SPIDAminSearch.show
        Unload SPIDAminSearch
    Me.show
End Sub

Private Sub cbTabGeneric_Click()
    Me.Hide
        Load TabGeneric
            TabGeneric.show
        Unload TabGeneric
    Me.show
End Sub

Private Sub cbTDOT_Click()
    Me.Hide
        Load TDOT
            TDOT.show
        Unload TDOT
    Me.show
End Sub

Private Sub cbTDS_Click()
    Me.Hide
        Load TDSMainForm
            TDSMainForm.show
        Unload TDSMainForm
    Me.show
End Sub

Private Sub cbTemplates_Click()
    Me.Hide
        Load TemplateEditor
            TemplateEditor.show
        Unload TemplateEditor
    Me.show
End Sub

Private Sub cbTicket_Click()
    Me.Hide
        Load TroubleTicket
            TroubleTicket.show
        Unload TroubleTicket
    Me.show
End Sub

Private Sub cbTrim_Click()
'    Dim SSobj2 As AcadSelectionSet
'    Dim objTrim As AcadBlockReference
'    Dim attList As Variant
'    Dim filterType, filterValue As Variant
'    Dim grpCode(0) As Integer
'    Dim grpValue(0) As Variant
'    Dim entBlock As AcadEntity
'    Dim iCount As Integer
'
'    Me.hide
'    iCount = 0
'
'    grpCode(0) = 8
'    grpValue(0) = "_Annotations"
'    filterType = grpCode
'    filterValue = grpValue
'
'    Set SSobj2 = ThisDrawing.SelectionSets.Add("SSobj2")
'    SSobj2.Select acSelectionSetAll, , , filterType, filterValue
'
'    For Each entBlock In SSobj2
'        If Not entBlock.ObjectName = "AcDbBlockReference" Then GoTo Next_entBlock
'        Set objTrim = entBlock
'        If Not objTrim.Name = "__Trim" Then GoTo Next_entBlock
'
'        attList = objTrim.GetAttributes
'        attList(0).TextString = "T.T=" & attList(0).TextString & "'"
'        objTrim.Layer = "Integrity Notes"
'        objTrim.Update
'
'        iCount = iCount + 1
'Next_entBlock:
'    Next entBlock
'
'    SSobj2.Clear
'    SSobj2.Delete
'
'    MsgBox "Trim Blocks: " & iCount
'    Me.show
End Sub

Private Sub cbTrace_Click()
    Me.Hide
        Load TraceForm
            TraceForm.show
        Unload TraceForm
    Me.show
End Sub

Private Sub cbUnits_Click()
    Me.Hide
        Load UnitForm
        UnitForm.show
        Unload UnitForm
    Me.show
End Sub

Private Sub cbUnitTab_Click()
    Me.Hide
        Load TabGeneric
        TabGeneric.show
        Unload TabGeneric
    Me.show
End Sub

Private Sub cbUpdateAllSS_Click()
    Dim objSS1 As AcadSelectionSet
    Dim entBlock As AcadObject
    Dim obrTemp As AcadBlockReference
    Dim attItem2 As Variant
    Dim mode As Integer
    'Dim str As String
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim str As String
    
  On Error Resume Next
    
    'str = ThisDrawing.Name
    'tbWO.Value = Left(str, 8)
    'str = Right(str, (Len(str) - 9))
    'tbProject.Value = Left(str, (Len(str) - 4))
    
    grpCode(0) = 8
    grpValue(0) = "Integrity Sheets"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS1 = ThisDrawing.SelectionSets.Add("objSS1")
    objSS1.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entBlock In objSS1
        'If TypeOf entBlock Is AcadBlockReference Then
            Set obrTemp = entBlock
        'Else
            'MsgBox "Error" & Err
            'GoTo exit_sub
        'End If
        
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
    objSS1.Delete
End Sub

Private Sub cbUnnamedForm_Click()
    Me.Hide
        Load UnnamedForm
        UnnamedForm.show
        Unload UnnamedForm
    Me.show
End Sub

Private Sub cbUpdateAncGuy_Click()
    Me.Hide
        Call UpdateAnchors
    Me.show
End Sub

Private Sub cbValidate_Click()
    Me.Hide
        Load ValidateFielding
        ValidateFielding.show
        Unload ValidateFielding
    Me.show
End Sub

Private Sub cbValidateAttach_Click()
    Me.Hide
        Load ValidateAttachments
        ValidateAttachments.show
        Unload ValidateAttachments
    Me.show
End Sub

Private Sub cbValidateCounts_Click()
    Me.Hide
        Load ValidateCounts
        ValidateCounts.show
        Unload ValidateCounts
    Me.show
End Sub

Private Sub cbValidateCustomers_Click()
    Me.Hide
        Load ValidateCustomers
        ValidateCustomers.show
        Unload ValidateCustomers
    Me.show
End Sub

Private Sub cbValidateHO1_Click()
    Me.Hide
        Load ValidateHO1
        ValidateHO1.show
        Unload ValidateHO1
    Me.show
End Sub

Private Sub cbValidateMCL_Click()
    Me.Hide
        Load ValidateMCL
        ValidateMCL.show
        Unload ValidateMCL
    Me.show
End Sub

Private Sub cbValidateML_Click()
    Me.Hide
        Load ValidateML
        ValidateML.show
        Unload ValidateML
    Me.show
End Sub

Private Sub cbVerifyAsFielded_Click()
    Me.Hide
        'Load CheckFieldedPoles
        'CheckFieldedPoles.show
        'Unload CheckFieldedPoles
        Load ValidateAttachments
        ValidateAttachments.show
        Unload ValidateAttachments
    Me.show
End Sub

Private Sub cbVerifyAttCallouts_Click()
    Me.Hide
        Load VerifyAttachmentCallouts
            VerifyAttachmentCallouts.show
        Unload VerifyAttachmentCallouts
    Me.show
End Sub

Private Sub cbVerifyUnits_Click()
    Me.Hide
        Load VerifyUnits
            VerifyUnits.show
        Unload VerifyUnits
    Me.show
End Sub

Private Sub cbVHLE_Click()
    Me.Hide
    Load DesignVHLE
    DesignVHLE.show
    Unload DesignVHLE
    Me.show
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
    Load OtherSubsPoles
    OtherSubsPoles.show
    Unload OtherSubsPoles
    Me.show
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
    Load zMisc
    zMisc.show
    Unload zMisc
    Me.show
End Sub

Private Sub CommandButton3_Click()
    Me.Hide
        Load PlaceCountCallouts
        PlaceCountCallouts.show
        Unload PlaceCountCallouts
    Me.show
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

Private Sub Label5_Click()
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
    
    cbExchange.Value = attItem(2).TextString
    tbRST.Value = attItem(3).TextString
    cbCounty.Value = attItem(4).TextString
    tbCity.Value = attItem(5).TextString
    
    Me.show
End Sub

Private Sub CommandButton4_Click()
    Me.Hide
        Load TransferToMap
        TransferToMap.show
        Unload TransferToMap
    Me.show
End Sub

Private Sub CommandButton5_Click()
    Me.Hide
        Load ReplaceF1
            ReplaceF1.show
        Unload ReplaceF1
    Me.show
End Sub

Private Sub ConvertSPOle_Click()
    Me.Hide
        Load ConvertToSPole
        ConvertToSPole.show
        Unload ConvertToSPole
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

Private Sub SortLBItems()
    Dim strArrayList(), strArraySorted() As String
    Dim attArray, attItem As Variant
    Dim str1, str2, strItem As String
    Dim i, iDWGNum, test1 As Integer
    
  On Error Resume Next
    test1 = 0
    ReDim strArraySorted(0 To lbSheets.ListCount)
    
    For i = 0 To lbSheets.ListCount - 1
        str1 = lbSheets.List(i)
        attArray = Split(str1, vbTab)
        str2 = Right(attArray(0), 3)
        
        If str2 = "" Then GoTo Next_I
        If str2 = "MAP" Then
            iDWGNum = 0
        Else
            iDWGNum = CInt(str2)
        End If
        
        If iDWGNum > test1 Then test1 = iDWGNum
        
        strArraySorted(iDWGNum) = str1
Next_I:
    Next i
    
    lbSheets.Clear
    For i = 0 To test1
        lbSheets.AddItem strArraySorted(i)
    Next i
End Sub

Private Sub LDwgScale_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    
    Dim objSS1 As AcadSelectionSet
    Dim entBlock As AcadObject
    Dim obrTemp As AcadBlockReference
    Dim attItem2, vTemp As Variant
    Dim mode As Integer
    'Dim str As String
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim str As String
    
  On Error Resume Next
    
    str = ThisDrawing.Name
    tbWO.Value = Left(str, 8)
    str = Right(str, (Len(str) - 9))
    tbProject.Value = Left(str, (Len(str) - 4))
    
    grpCode(0) = 8
    grpValue(0) = "Integrity Sheets"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS1 = ThisDrawing.SelectionSets.Add("objSS1")
    objSS1.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entBlock In objSS1
        'If TypeOf entBlock Is AcadBlockReference Then
            Set obrTemp = entBlock
        'Else
            'MsgBox "Error" & Err
            'GoTo exit_sub
        'End If
        
        If obrTemp.Name = "SS-11x17" Then
            attItem2 = obrTemp.GetAttributes
            vTemp = Split(attItem2(0).TextString, " ")
            'If Not (vTemp(0) = "DWG") Then GoTo Next_entBlock
            Select Case vTemp(0)
                Case "DWG"
                Case "MAP"
                Case "PMT"
                    GoTo Next_entBlock
                'Case Else
                '    GoTo Next_entBlock
            End Select
            'If Not vTemp(0) = "DWG" Then
            '    If vTemp(0) = "MAP" Or vTemp(1) = "MAP" Then
            '    GoTo Next_entBlock
            '
            'End If
                
            str = attItem2(0).TextString & vbTab & obrTemp.XScaleFactor
            str = str & vbTab & obrTemp.InsertionPoint(0) & vbTab & obrTemp.InsertionPoint(1)
            lbSheets.AddItem str
        End If
Next_entBlock:
    Next entBlock
    
    Call SortLBItems
Exit_Sub:
    objSS1.Delete
End Sub

Private Sub UpdateAnchors()
    Dim dTotal, dFeet, dInch As Double
    Dim iFeet, iInch As Integer
    Dim entPole As AcadObject
    Dim obrGP As AcadBlockReference
    Dim attItem, basePnt, vList As Variant
    Dim insertPnt As Variant
    Dim strLH, strPull As String
    Dim vList2 As Variant
    Dim strTemp, strBlockName As String
    Dim vLH As Variant
    Dim dScale, dRotate As Double
    Dim layerObj As AcadLayer
    
    'Me.hide
On Error Resume Next

update_block:
    Err = 0
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Guy or Anchor: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If TypeOf entPole Is AcadBlockReference Then
        Set obrGP = entPole
    Else
        GoTo Exit_Sub
    End If
    
    attItem = obrGP.GetAttributes
    
    Select Case obrGP.Name
        Case "ExGuyOR", "ExGuyOL"
            Load ConUpdateGuy
            
            If Not attItem(2).TextString = "" Then
                vLH = Split(attItem(2).TextString, "/")
                ConUpdateGuy.tbLead.Value = vLH(0)
                ConUpdateGuy.tbHeight.Value = vLH(1)
            End If
            
            ConUpdateGuy.show
            
            attItem(2).TextString = ConUpdateGuy.tbLead.Value & "/" & ConUpdateGuy.tbHeight.Value
            attItem(3).TextString = "PE1-3G"
            
            Unload ConUpdateGuy
        Case "ExAncOR", "ExAncOL"
            attItem(0).TextString = "PF1-5A"
    End Select
    obrGP.Update

    GoTo update_block
    
Exit_Sub:
End Sub

Private Sub DropCallout()
    Dim amap As AcadMap
    Dim ODRcs As ODRecords
    Dim tbl As ODTable
    Dim tbls As ODTables
    Dim boolVal As Boolean
    Dim entPole As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objPolyline As AcadPolyline
    Dim objLine As AcadLine
    Dim obrDrop As AcadBlockReference
    Dim objBlock As AcadBlockReference
    Dim attItem, vBasePnt As Variant
    Dim strType, strLength As String
    Dim strTemp As String
    Dim dRotate, dScale, Pi As Double
    Dim dTest, dDistScale As Double
    Dim dInsertPnt(0 To 2) As Double
    Dim dStartPnt(0 To 2) As Double
    Dim dEndPnt(0 To 2) As Double
    Dim dToText(0 To 2) As Double
    Dim dFromText(0 To 2) As Double
    Dim dDiff(0 To 2) As Double
    Dim vStartPnt, vEndPnt As Variant
    Dim vCoords, vAttList As Variant
    Dim layerObj As AcadLayer
    Dim iPosition, iLength As Integer
    
    'Me.Hide
  On Error Resume Next
  
    Pi = 3.14159265359
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
    Set tbls = amap.Projects(ThisDrawing).ODTables
    
get_Drop:

    dScale = cbScale.Value / 100
    If dScale = 0 Then dScale = 1#
    
    Err = 0
    
    ThisDrawing.Utility.GetEntity entPole, vBasePnt, "Select Drop: "
    
    If Not Err = 0 Then GoTo Exit_Sub
    
    Select Case entPole.ObjectName
        Case "AcDbLine"
            Set objLine = entPole
            iLength = CInt(objLine.Length)
            vStartPnt = objLine.StartPoint
            vEndPnt = objLine.EndPoint
            
            dDiff(0) = vEndPnt(0) - vStartPnt(0)
            dDiff(1) = vEndPnt(1) - vStartPnt(1)
            dDiff(2) = (dDiff(0) * dDiff(0)) + (dDiff(1) * dDiff(1))
            dDiff(2) = Sqr(dDiff(2))
            
            If vBasePnt(0) > vStartPnt(0) Then
                dToText(0) = vBasePnt(0) - vStartPnt(0)
                dToText(1) = vBasePnt(1) - vStartPnt(1)
            Else
                dToText(0) = vStartPnt(0) - vBasePnt(0)
                dToText(1) = vStartPnt(1) - vBasePnt(1)
            End If
            dToText(2) = Sqr((dToText(0) * dToText(0)) + (dToText(1) * dToText(1)))
            
            dDistScale = dToText(2) / dDiff(2)
            dInsertPnt(0) = vStartPnt(0) + (dDiff(0) * dDistScale)
            dInsertPnt(1) = vStartPnt(1) + (dDiff(1) * dDistScale)
            dInsertPnt(2) = 0#
            
            If dToText(0) = 0 Then
                dRotate = Pi / 2
            Else
                dRotate = Atn(dDiff(1) / dDiff(0))
            End If
            
            iPosition = 34
            
            Select Case objLine.Layer
                Case "Integrity Drops-Aerial", "OHSecondary", "OH_Secondary", "F-DROP-A-E"
                    strType = "SEAO"
                Case Else
                    strType = "SEBO"
            End Select
            
            strLength = iLength
            
        Case "AcDbPolyline"
            Set objLWP = entPole
            iLength = CInt(objLWP.Length)
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
            
            dDistScale = Abs(dToText(2) / dDiff(2))
            dInsertPnt(0) = vCoords(i - 3) + (dDiff(0) * dDistScale)
            dInsertPnt(1) = vCoords(i - 2) + (dDiff(1) * dDistScale)
            dInsertPnt(2) = 0#
            
            If dToText(0) = 0 Then
                dRotate = Pi / 2
            Else
                dRotate = Atn(dDiff(1) / dDiff(0))
            End If
            
            Select Case objLWP.Layer
                Case "Integrity Drops-Aerial", "OHSecondary", "OH_Secondary", "F-DROP-A-E"
                    strType = "SEAO"
                Case Else
                    strType = "SEBO"
            End Select
            
            strLength = iLength
            
        Case "AcDbBlockReference"
            Set objBlock = entPole
            
            Select Case objBlock.Name
                Case "Drop"
                    vAttList = objBlock.GetAttributes
                    strTemp = vAttList(0).TextString
                    If InStr(strTemp, "-") > 0 Then
                        vCoords = Split(strTemp, "-")
                        vCoords(0) = CInt(vCoords(0)) + 1
                        vAttList(0).TextString = vCoords(0) & "-" & vCoords(1)
                    Else
                        vAttList(0).TextString = "2-" & vAttList(0).TextString
                    End If
                    
                    objBlock.Update
                Case Else
                    Me.show
                    Exit Sub
            End Select
            
            GoTo get_Drop
        Case Else
            MsgBox entPole.ObjectName
            Me.show
            Exit Sub
    End Select

    Err = 0
    Set obrDrop = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "Drop", dScale, dScale, dScale, dRotate)
    If Not Err = 0 Then MsgBox Err.desription
    If strType = "SEAO" Then
        obrDrop.Layer = "Integrity Drops-Aerial Text"
    Else
        obrDrop.Layer = "Integrity Drops-Buried Text"
    End If
    
    vAttList = obrDrop.GetAttributes
    vAttList(0).TextString = strType
    vAttList(1).TextString = strLength & "'"
    obrDrop.Update
    
    GoTo get_Drop
Exit_Sub:

End Sub

Private Sub DeleteDropCallouts()
    Me.Hide
    
    Dim result As Integer
    
    result = MsgBox("Are you sure you want to delete all drop callouts?", vbYesNo, "Delete Drop Callouts")
    If result = vbNo Then
        Me.show
    End If
    
    Dim vDwgLL, vDwgUR As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAtt As Variant
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    
    grpCode(0) = 2
    grpValue(0) = "Drop"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    objSS.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    If objSS.count < 1 Then GoTo Exit_Sub
    
    For Each objBlock In objSS
        Select Case objBlock.Layer
            Case "Integrity Drops-Aerial Text", "Integrity Drops-Buried Text"
                objBlock.Delete
        End Select
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    ThisDrawing.Regen acAllViewports
    
    Me.show
End Sub

