VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PonVHLE 
   Caption         =   "VHLE - PON"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13890
   OleObjectBlob   =   "PonVHLE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PonVHLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbBGetBldg_Click()
    tbRes.Value = GetBlockCount("Integrity Building-RES")
    tbChurch.Value = GetBlockCount("Integrity Building-CHURCH") + GetBlockCount("Integrity Building-CHU")
    tbTRL.Value = GetBlockCount("Integrity Building-TRLR")
    tbMDU.Value = GetBlockCount("Integrity Building-MDU")
    tbBus.Value = GetBlockCount("Integrity Building-BUS")
    tbSchool.Value = GetBlockCount("Integrity Building-SCH")
    tbSG.Value = GetBlockCount("Integrity SmartGrid")
    
    tbRESTot.Value = CInt(tbRes.Value) + CInt(tbChurch.Value) + CInt(tbSG.Value)
    If cbTRL Then
        tbRESTot.Value = CInt(tbRESTot.Value) + CInt(tbTRL.Value)
    End If
    If cbMDU Then
        tbRESTot.Value = CInt(tbRESTot.Value) + CInt(tbMDU.Value)
    End If
    
    tbBUSTot.Value = CInt(tbBus.Value) + CInt(tbSchool.Value)
    tbMeters.Value = CInt(tbRESTot.Value) + CInt(tbBUSTot.Value)
End Sub

Private Sub cbClearForm_Click()
    Call ClearAll
End Sub

Private Sub cbGetInfo_Click()
    Dim iRESTotal, iBUSTotal, iTotal As Integer
    Dim lAC, lACoil, lBC, lBCoil, lATotal, lBTotal As Long
    
    iRes = 0: iChurch = 0: iTRL = 0: iMDU = 0: iBus = 0: iSchool = 0
    lAC = 0: lACoil = 0: lATotal = 0
    lBC = 0: lBCoil = 0: lBTotal = 0
        
    tbRes.Value = GetBlockCount("Integrity Building-RES")
    tbChurch.Value = GetBlockCount("Integrity Building-CHURCH") + GetBlockCount("Integrity Building-CHU")
    tbTRL.Value = GetBlockCount("Integrity Building-TRLR")
    tbMDU.Value = GetBlockCount("Integrity Building-MDU")
    tbBus.Value = GetBlockCount("Integrity Building-BUS")
    tbSchool.Value = GetBlockCount("Integrity Building-SCH")
    tbSG.Value = GetBlockCount("Integrity SmartGrid")
    
    tbRESTot.Value = CInt(tbRes.Value) + CInt(tbChurch.Value) + CInt(tbSG.Value)
    If cbTRL Then
        tbRESTot.Value = CInt(tbRESTot.Value) + CInt(tbTRL.Value)
    End If
    If cbMDU Then
        tbRESTot.Value = CInt(tbRESTot.Value) + CInt(tbMDU.Value)
    End If
    
    tbBUSTot.Value = CInt(tbBus.Value) + CInt(tbSchool.Value)
    tbMeters.Value = CInt(tbRESTot.Value) + CInt(tbBUSTot.Value)
    
    Call GetCables
    Call GetCoils
    
    tbACTotal.Value = CLng(tbAC.Value) + CLng(tbACE.Value) + CLng(tbACoil.Value)
    tbBCTotal.Value = CLng(tbBC.Value) + CLng(tbBCE.Value) + CLng(tbBCoil.Value)
End Sub

Private Sub cbGetPolygon_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim vAttList, vTemp As Variant
    Dim vCoords, vArray As Variant
    Dim dCoords() As Double
    Dim iRes, iBus, iSG, iTemp, iCounter As Integer
    Dim iChurch, iSchool, iTRL, iMDU As Integer
    Dim lADrop, lBDrop As Long
    Dim lRouteFeet, iPoles As Long
    Dim iRouteMile, iRouteKF As Double
    Dim iDistance, iGuy, iClosure As Integer
    Dim iRESTotal, iBUSTotal, iSpans As Integer
    Dim strUnit, strTemp As String
    Dim lACT, lACM, lACE, lACL As Long
    Dim lBCT, lBCM, lBCE, lBCL As Long
    Dim lUCT, lUCM, lUCE, lUCL As Long
    Dim iSkip As Integer
    
    iRes = 0: iChurch = 0: iTRL = 0: iMDU = 0
    iBus = 0: iSchool = 0: iSG = 0: iGuy = 0
    iClosure = 0: lADrop = 0: lBDrop = 0
    iPoles = 0: iSpans = 1
    
    lACT = 0: lACM = 0: lACE = 0: lACL = 0
    lBCT = 0: lBCM = 0: lBCE = 0: lBCL = 0
    lUCT = 0: lUCM = 0: lUCE = 0: lUCL = 0
    
    On Error Resume Next
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select PON Border: "
    If Not objEntity.ObjectName = "AcDbPolyline" Then
        MsgBox "Error: Invalid Selection."
        Me.show
        Exit Sub
    End If
    
    Call ClearAll
    
    Set objLWP = objEntity
    vCoords = objLWP.Coordinates
    
    iTemp = (UBound(vCoords) + 1) / 2 * 3 - 1
    ReDim dCoords(iTemp) As Double
    
    iCounter = 0
    For i = 0 To UBound(vCoords) Step 2
        dCoords(iCounter) = vCoords(i)
        iCounter = iCounter + 1
        dCoords(iCounter) = vCoords(i + 1)
        iCounter = iCounter + 1
        dCoords(iCounter) = 0#
        iCounter = iCounter + 1
    Next i
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    
    objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadBlockReference Then
            Set objBlock = objEntity
            If objBlock.Layer = "Integrity Existing" Then GoTo Next_objEntity
            vAttList = objBlock.GetAttributes
            
            Select Case objBlock.Name
                Case "Customer"
                    Select Case vAttList(5).TextString
                        Case "", "R"
                            iRes = iRes + 1
                        Case "B"
                            iBus = iBus + 1
                        Case "C"
                            iChurch = iChurch + 1
                        Case "M"
                            iMDU = iMDU + 1
                        Case "S"
                            iSchool = iSchool + 1
                        Case "T"
                            iTRL = iTRL + 1
                    End Select
                'Case "RES", "LOT"
                    'iRes = iRes + 1
                'Case "CHURCH"
                    'iChurch = iChurch + 1
                'Case "TRLR"
                    'iTRL = iTRL + 1
                'Case "MDU"
                    'iMDU = iMDU + 1
                'Case "BUSINESS"
                    'iBus = iBus + 1
                'Case "SCHOOL"
                    'iSchool = iSchool + 1
                Case "SG"
                    iSG = iSG + 1
                Case "cable_span"
                    'iSpans = iSpans + 1
                    
                    If CInt(Left(vAttList(2).TextString, 1)) = 0 Then
                        strTemp = vAttList(2).TextString
                        Select Case Left(strTemp, 1)
                            Case "+"
                                vTemp = Split(strTemp, "'")
                                iDistance = CInt(Replace(vTemp(0), "+", ""))
                                strUnit = "BFO(" & Replace(vAttList(1).TextString, "F", "") & ")D"
                                lBCE = lBCE + iDistance
                                Call AddUnitToTab("BM81", 1)
                            Case Else
                                vTemp = Split(strTemp, "=")
                                strUnit = vTemp(0)
                                iDistance = CInt(Replace(vTemp(1), "'", ""))
                        End Select
                        
                        Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                        GoTo Next_objEntity
                    End If
                    
                    iDistance = CInt(Replace(vAttList(2).TextString, "'", ""))
                    lRouteFeet = lRouteFeet + iDistance
                    
                    strTemp = Replace(vAttList(1).TextString, "F", "")
                    If strTemp = "" Then strTemp = "??"
                    vTemp = Split(strTemp, " ")
                    
                    iSkip = 0
                    If UBound(vTemp) > 0 Then
                        If vTemp(0) = "" Then
                            iSkip = 1
                        End If
                    End If
                    
                    Select Case objBlock.Layer
                        Case "Integrity Cable-UG Text", "Integrity Proposed-Buried"
                            strTemp = "UO("
                            strUnit = strTemp & vTemp(0) & ")"
                            If iSkip = 0 Then Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                            If iSkip = 0 Then lUCM = lUCM + iDistance
                            If iSkip = 0 Then Call AddUnitToTab("UD(1X1-4)", CInt(iDistance))
                        Case "Integrity Cable-Buried Text"
                            strTemp = "BFO("
                            strUnit = strTemp & vTemp(0) & ")"
                            If iSkip = 0 Then Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                            If iSkip = 0 Then lBCM = lBCM + iDistance
                        Case Else
                            iSpans = iSpans + 1
                            strTemp = "CO("
                            strUnit = strTemp & vTemp(0) & ")M"
                            If iSkip = 0 Then Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                            If iSkip = 0 Then lACM = lACM + iDistance
                    End Select
                    
                    If UBound(vTemp) = 0 Then GoTo Next_objEntity
                    
                    For n = 1 To UBound(vTemp)
                        Select Case strTemp
                            Case "UO("
                                strUnit = strTemp & vTemp(n) & ")"
                                Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                                lUCE = lUCE + iDistance
                                
                                If iSkip = 1 Then
                                    Call AddUnitToTab("RRUD", CInt(iDistance))
                                    iSkip = 0
                                End If
                            Case "BFO("
                                strUnit = strTemp & vTemp(n) & ")D"
                                Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                                lBCE = lBCE + iDistance
                            Case Else
                                strUnit = strTemp & vTemp(n) & ")E"
                                Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                                lACE = lACE + iDistance
                        End Select
                    Next n
                    
                Case "Map splice", "iClosure"
                    
                    iClosure = iClosure + 1
                    
                    Select Case objBlock.Layer
                        Case "Integrity Map Splices-Aerial"
                            strUnit = "HACO(" & vAttList(0).TextString & ")"
                        Case Else
                            strUnit = "HBFO(" & vAttList(0).TextString & ")"
                    End Select
                    
                    Call AddUnitToTab(CStr(strUnit), 1)
                    
                Case "Map coil"
                    'iClosure = iClosure + 1
                    
                    iDistance = CInt(Replace(vAttList(0).TextString, "'", ""))
                    
                    Select Case objBlock.Layer
                        Case "Integrity Map Coils-Aerial"
                            strUnit = "CO(" & Replace(vAttList(1).TextString, "F", "") & ") LOOP"
                            lACL = lACL + iDistance
                            Call AddUnitToTab("PM52", 1)
                        Case "Integrity Map Coils-Buried"
                            strUnit = "BFO(" & Replace(vAttList(1).TextString, "F", "") & ") LOOP"
                            lBCL = lBCL + iDistance
                        Case Else
                            strUnit = "UO(" & Replace(vAttList(1).TextString, "F", "") & ") LOOP"
                            lUCL = lUCL + iDistance
                    End Select
                    
                    'MsgBox strUnit & vbCr & iDistance
                    Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                    
                Case "Drop"
                    iDistance = CInt(Replace(vAttList(1).TextString, "'", ""))
                    Select Case vAttList(0).TextString
                        Case "SEBO"
                            lBDrop = lBDrop + iDistance
                            strUnit = "SEBO"
                        Case Else
                            lADrop = lADrop + iDistance
                            strUnit = "SEAO"
                    End Select
                    
                    Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                    
                Case "ohgL", "ohgR"
                    iSpans = iSpans + 1
                    
                    iDistance = CInt(Replace(vAttList(0).TextString, "'", ""))
                    strUnit = vAttList(1).TextString
                    
                    Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                    Call AddUnitToTab(CStr("{O} OHG"), 1)
                    
                Case "ExGuyOL", "ExGuyOR"
                    iGuy = iGuy + 1
                    
                    If UBound(vAttList) = 6 Then
                        strUnit = vAttList(3).TextString
                    Else
                        strUnit = vAttList(2).TextString
                    End If
                    
                    Call AddUnitToTab(CStr(strUnit), 1)
                    Call AddUnitToTab(CStr("PM11"), 1)
                    
                Case "ExAncOL", "ExAncOR"
                    If UBound(vAttList) = 4 Then
                        strUnit = vAttList(0).TextString
                    Else
                        strUnit = "PF"
                    End If
                    
                    Call AddUnitToTab(CStr(strUnit), 1)
                    
                Case "__Trim"
                    vTemp = Split(vAttList(0).TextString, "=")
                    iDistance = CInt(Replace(vTemp(1), "'", ""))
                    strUnit = "R3-5"
                    
                    Call AddUnitToTab(CStr(strUnit), CInt(iDistance))
                Case "iPole"
                    If vAttList(0).TextString = "" Then GoTo Next_objEntity
                    If vAttList(0).TextString = "POLE" Then GoTo Next_objEntity
                    
                    iPoles = iPoles + 1
                    
                    If vAttList(2).TextString = "" Then
                        strUnit = "{P} ??"
                    Else
                        strUnit = "{P} " & vAttList(2).TextString
                    End If
                    
                    Call AddUnitToTab(CStr(strUnit), 1)
                    
                    If Not vAttList(8).TextString = "" Then
                        Call AddUnitToTab("PM2A", 1)
                        
                        If vAttList(8).TextString = "B" Then
                            Call AddUnitToTab("PM2R", 1)
                        End If
                    End If
                Case "iLHH", "dHH"
                    strUnit = vAttList(1).TextString
                    Call AddUnitToTab(CStr(strUnit), 1)
                Case "iLPED", "PED"
                    strUnit = vAttList(1).TextString
                    Call AddUnitToTab(CStr(strUnit), 1)
                Case "FP", "dFP"
                    Call AddUnitToTab("FP", 1)
            End Select
Next_objEntity:
        Set objBlock = Nothing
        End If
    Next objEntity
    
    objSS.Clear
    objSS.Delete
    
    If iSpans = 1 Then
        iSpans = 0
    Else
        Call AddUnitToTab("{P}  Poles", CInt(iSpans))
    End If
    
    lACT = lACM + lACE + lACL
    lBCT = lBCM + lBCE + lBCL
    lUCT = lUCM + lUCE + lUCL
    
    tbACTotal.Value = lACT
    tbAC.Value = lACM
    tbACE.Value = lACE
    tbACoil.Value = lACL
    
    tbBCTotal.Value = lBCT
    tbBC.Value = lBCM
    tbBCE.Value = lBCE
    tbBCoil.Value = lBCL
    
    tbUCTotal.Value = lUCT
    tbUC.Value = lUCM
    tbUCE.Value = lUCE
    tbUCoil.Value = lUCL

    tbAD.Value = lADrop
    tbBD.Value = lBDrop
    
    tbRouteMiles.Value = CLng(lRouteFeet * 1000 / 5280) / 1000
    tbRouteKF.Value = lRouteFeet / 1000
    tbPoles.Value = iSpans
    'tbPoles.Value = iPoles
    
    tbTerminals.Value = iClosure
    tbGuys.Value = iGuy
        
    tbRes.Value = iRes
    tbChurch.Value = iChurch
    tbTRL.Value = iTRL
    tbMDU.Value = iMDU
    tbBus.Value = iBus
    tbSchool.Value = iSchool
    tbSG.Value = iSG
    
    iRESTotal = iRes + iChurch + iSG
    If cbTRL.Value Then iRESTotal = iRESTotal + iTRL
    If cbMDU.Value Then iRESTotal = iRESTotal + iMDU
    tbRESTot.Value = iRESTotal
    
    tbBUSTot.Value = CInt(tbBus.Value) + CInt(tbSchool.Value)
    tbMeters.Value = CInt(tbRESTot.Value) + CInt(tbBUSTot.Value)
    
    'Call GetCables
    'Call GetCoils
    
    'tbACTotal.Value = CLng(tbAC.Value) + CLng(tbACE.Value) + CLng(tbACoil.Value)
    'tbBCTotal.Value = CLng(tbBC.Value) + CLng(tbBCE.Value) + CLng(tbBCoil.Value)
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Function GetBlockCount(strLayer As String)
    Dim objSS6 As AcadSelectionSet
    Dim entItem As AcadEntity
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim iCount As Integer
    
    iCount = 0
    
    grpCode(0) = 8
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS6 = ThisDrawing.SelectionSets.Add("objSS6")
    objSS6.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entItem In objSS6
        If TypeOf entItem Is AcadBlockReference Then iCount = iCount + 1
    Next entItem
    
    'iCount = objSS6.Count
    
    objSS6.Clear
    objSS6.Delete
    
    GetBlockCount = iCount
End Function

Private Sub GetCables()
    Dim objSS6 As AcadSelectionSet
    Dim entItem As AcadEntity
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList, vTest As Variant
    Dim lALength, lELash, lBLength, lBSec As Long
    Dim iSpan As Integer
    Dim strSpan As String
    
    lALength = 0
    lELash = 0
    lBLength = 0
    lBSec = 0
    
    grpCode(0) = 2
    grpValue(0) = "cable_span"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS6 = ThisDrawing.SelectionSets.Add("objSS6")
    objSS6.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS6
        vAttList = objBlock.GetAttributes
        strSpan = vAttList(2).TextString
        If Right(strSpan, 1) = "'" Then strSpan = Left(strSpan, Len(strSpan) - 1)
        If strSpan = "" Then GoTo Next_Object
        iSpan = CInt(strSpan)
        
        Select Case objBlock.Layer
            Case "Integrity Cable-Aerial Text"
                lALength = lALength + iSpan
                
                vTest = Split(vAttList(1).TextString, "F ")
                lELash = lELash + (UBound(vTest) * iSpan)
            Case "Integrity Cable-Buried Text"
                lBLength = lBLength + iSpan
                
                vTest = Split(vAttList(1).TextString, "F ")
                lBSec = lBSec + (UBound(vTest) * iSpan)
        End Select
Next_Object:
    Next objBlock
    
    tbAC.Value = lALength
    tbACE.Value = lELash
    
    tbBC.Value = lBLength
    tbBCE.Value = lBSec
    
    objSS6.Clear
    objSS6.Delete
End Sub

Private Sub GetCoils()
    Dim objSS6 As AcadSelectionSet
    Dim entItem As AcadEntity
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList, vTest As Variant
    Dim lACoil, lBCoil As Long
    Dim iSpan As Integer
    Dim strSpan As String
    
    lACoil = 0
    lBCoil = 0
    
    grpCode(0) = 2
    grpValue(0) = "Map coil"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS6 = ThisDrawing.SelectionSets.Add("objSS6")
    objSS6.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS6
        vAttList = objBlock.GetAttributes
        strSpan = vAttList(0).TextString
        If Right(strSpan, 1) = "'" Then strSpan = Left(strSpan, Len(strSpan) - 1)
        iSpan = CInt(strSpan)
        
        Select Case iSpan
            Case 20, 25
                lBCoil = lBCoil + iSpan
            Case Else
                lACoil = lACoil + iSpan
        End Select
    Next objBlock
    
    tbACoil.Value = lACoil
    tbBCoil.Value = lBCoil
    
    objSS6.Clear
    objSS6.Delete
End Sub

Private Function GetTerminals(strLayer As String)
    Dim objSS6 As AcadSelectionSet
    Dim entItem As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim iCount As Integer
    
    iCount = 0
    
    grpCode(0) = 8
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS6 = ThisDrawing.SelectionSets.Add("objSS6")
    objSS6.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entItem In objSS6
        If TypeOf entItem Is AcadBlockReference Then
            Set objBlock = entItem
            If objBlock.Name = "Map splice" Then iCount = iCount + 1
        End If
    Next entItem
    
    'iCount = objSS6.Count
    
    objSS6.Clear
    objSS6.Delete
    
    GetTerminals = iCount
End Function

Private Sub ClearAll()
    tbACTotal.Value = ""
    tbAC.Value = ""
    tbACE.Value = ""
    tbACoil.Value = ""
    tbBCTotal.Value = ""
    tbBC.Value = ""
    tbBCE.Value = ""
    tbBCoil.Value = ""
    tbUCTotal.Value = ""
    tbUC.Value = ""
    tbUCE.Value = ""
    tbUCoil.Value = ""
    
    tbAD.Value = ""
    tbBD.Value = ""
    
    tbRouteMiles.Value = ""
    tbRouteKF.Value = ""
    tbPoles.Value = ""
    tbTerminals.Value = ""
    tbSplitters.Value = ""
    tbGuys.Value = ""
    
    tbRESTot.Value = ""
    tbRes.Value = ""
    tbChurch.Value = ""
    tbTRL.Value = ""
    tbMDU.Value = ""
    tbBUSTot.Value = ""
    tbBus.Value = ""
    tbSchool.Value = ""
    tbSG.Value = ""
    tbMeters.Value = ""
    
    lbTab.Clear
    tbReport.Value = ""
End Sub

Private Sub AddUnitToTab(strItem As String, iDist As Integer)
    If lbTab.ListCount = 0 Then
        GoTo Add_Item
    End If
    
    For i = 0 To lbTab.ListCount - 1
        If lbTab.List(i, 0) = strItem Then
            'MsgBox lbTab.List(i, 0) & vbCr & lbTab.List(i, 1) & vbCr & iDist
            lbTab.List(i, 1) = CInt(lbTab.List(i, 1)) + iDist
            Exit Sub
        End If
    Next i
    
Add_Item:
    lbTab.AddItem strItem
    lbTab.List(lbTab.ListCount - 1, 1) = iDist
    
End Sub

Private Sub cbSort_Click()
    Dim strTemp, strTotal As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    
    iCount = lbTab.ListCount - 1
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbTab.List(b, 0) > lbTab.List(b + 1, 0) Then
                strTemp = lbTab.List(b + 1, 0)
                strTotal = lbTab.List(b + 1, 1)
                
                lbTab.List(b + 1, 0) = lbTab.List(b, 0)
                lbTab.List(b + 1, 1) = lbTab.List(b, 1)
                
                lbTab.List(b, 0) = strTemp
                lbTab.List(b, 1) = strTotal
            End If
        Next b
    Next a
    
    'iOffset = 0
    'For c = 1 To lbTab.ListCount - 1
        'iIndex = c - iOffset
        'If lbTab.List(iIndex, 0) = lbTab.List(iIndex - 1, 0) Then
            'lbTab.List(iIndex - 1, 1) = CInt(lbTab.List(iIndex - 1, 1)) + CInt(lbTab.List(iIndex, 1))
            'lbTab.RemoveItem iIndex
            'iOffset = iOffset + 1
        'End If
    'Next c
    
    tbReport.Value = ""
    
    For c = 0 To lbTab.ListCount - 1
        If c = lbTab.ListCount - 1 Then
            tbReport.Value = tbReport.Value & lbTab.List(c, 0) & vbTab & lbTab.List(c, 1)
        Else
            tbReport.Value = tbReport.Value & lbTab.List(c, 0) & vbTab & lbTab.List(c, 1) & vbCr
        End If
        
        'Select Case Left(lbTab.List(c, 0), 3)
            'Case "CO("
                'Select Case Right(lbTab.List(c, 0), 1)
                    'Case "P"
                    'Case "E"
                    'Case "M"
                'End Select
            'Case "BFO"
            'Case "UO("
        'End Select
    Next c
End Sub

Private Sub UserForm_Initialize()
    lbTab.ColumnCount = 2
    lbTab.ColumnWidths = "120;74"
End Sub
