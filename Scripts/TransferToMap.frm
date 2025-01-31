VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferToMap 
   Caption         =   "Transfer Data"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "TransferToMap.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransferToMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vDwgLL, vDwgUR As Variant
Dim vMapLL, vMapUR As Variant

Private Sub cbGetBoundaries_Click()
    Me.Hide
    
    Call GetBoundaries
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbTransfer_Click()
    Dim iNumber As Integer
    iNumber = 0
    
    If cbPoles.Value = True Then Call TransferPoles
    Me.Repaint
    
    If cbPEDs.Value = True Then
        iNumber = TransferBuried(CStr("sPed"))
        LPed.Caption = iNumber
        Me.Repaint
        iNumber = 0
    End If
    
    If cbHH.Value = True Then
        iNumber = TransferBuried(CStr("sHH"))
        LHH.Caption = iNumber
        Me.Repaint
        iNumber = 0
    End If
    
    If cbMH.Value = True Then
        iNumber = TransferBuried(CStr("sMH"))
        LMH.Caption = iNumber
        Me.Repaint
        iNumber = 0
    End If
    
    If cbPanel.Value = True Then
        iNumber = TransferBuried(CStr("sPanel"))
        LPanel.Caption = iNumber
        Me.Repaint
        iNumber = 0
    End If
    
    If cbFP.Value = True Then
        iNumber = TransferBuried(CStr("sFP"))
        LFP.Caption = iNumber
        Me.Repaint
    End If
    
    If cbCustomer.Value = True Then Call TransferCustomers
End Sub

Private Sub cbTransferLatLong_Click()
    Dim iNumber As Integer
    Dim strLine As String
    strLine = ""
    iNumber = 0
    
    If cbPoles.Value = True Then Call AddLLtoPoles
    Me.Repaint
    
    If cbPEDs.Value = True Then
        strLine = AddLLToBuried(CStr("sPed"))
        LPed.Caption = strLine
        Me.Repaint
        strLine = ""
    End If
    
    If cbHH.Value = True Then
        strLine = AddLLToBuried(CStr("sHH"))
        LHH.Caption = strLine
        Me.Repaint
        strLine = ""
    End If
    
    If cbMH.Value = True Then
        strLine = AddLLToBuried(CStr("sMH"))
        LMH.Caption = strLine
        Me.Repaint
        strLine = ""
    End If
    
    If cbPanel.Value = True Then
        strLine = AddLLToBuried(CStr("sPanel"))
        LPanel.Caption = strLine
        Me.Repaint
        strLine = ""
    End If
    
    If cbFP.Value = True Then
        strLine = AddLLToBuried(CStr("sFP"))
        LFP.Caption = strLine
        Me.Repaint
    End If
    
    If cbCustomer.Value = True Then LCustomer.Caption = "x"
        
End Sub

Private Sub GetBoundaries()
    On Error Resume Next
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    
    vMapLL = ThisDrawing.Utility.GetPoint(, "Get MAP LL Corner: ")
    vMapUR = ThisDrawing.Utility.GetCorner(vMapLL, vbCr & "Get MAP UR Corner: ")
    
    If Not Err = 0 Then Exit Sub
    
    cbGetBoundaries.Enabled = False
    cbTransferLatLong.Enabled = True
    cbTransfer.Enabled = True
    Label2.Enabled = True
End Sub

Private Sub TransferPoles()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSDwg As AcadSelectionSet
    Dim objSSMap As AcadSelectionSet
    Dim objPoleDwg As AcadBlockReference
    Dim objPoleMap As AcadBlockReference
    Dim vAttDwg, vAttMap As Variant
    Dim strPoleNum As String
    Dim strDwgLatLong, strMapLatLong As String
    Dim iPoles, iData, iAttach As Integer
    Dim iCounts, iUnits As Integer
    Dim strAtt(27) As String
    
    For i = 0 To 27
        strAtt(i) = ""
    Next i
    
    iPoles = 0
    iData = 0
    iAttach = 0
    iCounts = 0
    iUnits = 0
        
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSDwg = ThisDrawing.SelectionSets.Add("objSSDwg")
    objSSDwg.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    If objSSDwg.count < 1 Then GoTo Exit_Sub
    
    Set objSSMap = ThisDrawing.SelectionSets.Add("objSSMap")
    objSSMap.Select acSelectionSetWindow, vMapLL, vMapUR, filterType, filterValue
    If objSSMap.count < 1 Then GoTo Sub_Exit_Sub
    
    For Each objPoleDwg In objSSDwg
        vAttDwg = objPoleDwg.GetAttributes
        If vAttDwg(7).TextString = "" Then GoTo Next_objDwgPole
        
        strPoleNum = vAttDwg(0).TextString
        strDwgLatLong = vAttDwg(7).TextString
        
        If strPoleNum = "POLE" Then GoTo Next_objDwgPole
        If strPoleNum = "" Then GoTo Next_objDwgPole
        'iTest = 0
        
        For Each objPoleMap In objSSMap
            vAttMap = objPoleMap.GetAttributes
            strMapLatLong = vAttMap(7).TextString
            
            If strMapLatLong = strDwgLatLong Then
                For i = 0 To 6
                    strAtt(i) = vAttDwg(i).TextString
                Next i
                
                For i = 8 To 27
                    strAtt(i) = vAttDwg(i).TextString
                Next i
                
                For i = 0 To 6
                    vAttMap(i).TextString = strAtt(i)
                Next i
                
                For i = 8 To 27
                    vAttMap(i).TextString = strAtt(i)
                Next i
                
                objPoleMap.Update
                iPoles = iPoles + 1
                GoTo Next_objDwgPole
            End If
        Next objPoleMap
Next_objDwgPole:
    Next objPoleDwg
    
Sub_Exit_Sub:
    objSSMap.Clear
    objSSMap.Delete
    
Exit_Sub:
    objSSDwg.Clear
    objSSDwg.Delete
    
    LPole.Caption = iPoles
End Sub

Private Function TransferBuried(strName As String)
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSDwg As AcadSelectionSet
    Dim objSSMap As AcadSelectionSet
    Dim objBlockDwg As AcadBlockReference
    Dim objBlockMap As AcadBlockReference
    Dim vAttDwg, vAttMap As Variant
    Dim strNumber, strOmit As String
    Dim strDwgLatLong, strMapLatLong As String
    Dim iItems As Integer
    Dim strAtt(7) As String
    
    For i = 0 To 7
        strAtt(i) = ""
    Next i
    
    iItems = 0
    strOmit = Replace(UCase(strName), "S", "")
        
    grpCode(0) = 2
    grpValue(0) = strName
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSDwg = ThisDrawing.SelectionSets.Add("objSSDwg")
    objSSDwg.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    
    Set objSSMap = ThisDrawing.SelectionSets.Add("objSSMap")
    objSSMap.Select acSelectionSetWindow, vMapLL, vMapUR, filterType, filterValue
    
    If objSSDwg.count < 1 Then GoTo Exit_Sub
    If objSSMap.count < 1 Then GoTo Exit_Sub
    
    For Each objBlockDwg In objSSDwg
        vAttDwg = objBlockDwg.GetAttributes
        
        strNumber = vAttDwg(0).TextString
        If strNumber = "" Then GoTo Next_objBlockDwg
        
        If strOmit = vAttDwg(0).TextString Then GoTo Next_objBlockDwg
        
        If vAttDwg(3).TextString = "" Then GoTo Next_objBlockDwg
        strDwgLatLong = vAttDwg(3).TextString
        
        For Each objBlockMap In objSSMap
            vAttMap = objBlockMap.GetAttributes
            strMapLatLong = vAttMap(3).TextString
            
            If strMapLatLong = strDwgLatLong Then
                For i = 0 To 7
                    strAtt(i) = vAttDwg(i).TextString
                Next i
                
                For i = 0 To 7
                    vAttMap(i).TextString = strAtt(i)
                Next i
                
                objBlockMap.Update
                iItems = iItems + 1
                GoTo Next_objBlockDwg
            End If
        Next objBlockMap
Next_objBlockDwg:
    Next objBlockDwg
    
Exit_Sub:
    objSSDwg.Clear
    objSSDwg.Delete
    
    objSSMap.Clear
    objSSMap.Delete
    
    TransferBuried = iItems
End Function

Private Sub TransferCustomers()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSDwg As AcadSelectionSet
    Dim objSSMap As AcadSelectionSet
    Dim objBlockDwg As AcadBlockReference
    Dim objBlockMap As AcadBlockReference
    Dim vAttDwg, vAttMap As Variant
    Dim strNumber, strOmit As String
    Dim strDwgLatLong, strMapLatLong As String
    Dim iItems As Integer
    Dim strAtt(5) As String
    
    For i = 0 To 5
        strAtt(i) = ""
    Next i
    
    iItems = 0
        
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSDwg = ThisDrawing.SelectionSets.Add("objSSDwg")
    objSSDwg.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    
    Set objSSMap = ThisDrawing.SelectionSets.Add("objSSMap")
    objSSMap.Select acSelectionSetWindow, vMapLL, vMapUR, filterType, filterValue
    If objSSDwg.count < 1 Then GoTo Exit_Sub
    If objSSMap.count < 1 Then GoTo Exit_Sub
    
    For Each objBlockDwg In objSSDwg
        vAttDwg = objBlockDwg.GetAttributes
        
        strNumber = vAttDwg(0).TextString
        If strNumber = "" Then GoTo Next_objBlockDwg
        
        For Each objBlockMap In objSSMap
            vAttMap = objBlockMap.GetAttributes
            
            If vAttMap(1).TextString = vAttDwg(1).TextString Then
                If vAttMap(2).TextString = vAttDwg(2).TextString Then
                    For i = 0 To 5
                        strAtt(i) = vAttDwg(i).TextString
                    Next i
                    
                    For i = 0 To 5
                        vAttMap(i).TextString = strAtt(i)
                    Next i
                
                    objBlockMap.Update
                    iItems = iItems + 1
                    GoTo Next_objBlockDwg
                End If
            End If
        Next objBlockMap
Next_objBlockDwg:
    Next objBlockDwg
    
Exit_Sub:
    objSSDwg.Clear
    objSSDwg.Delete
    
    objSSMap.Clear
    objSSMap.Delete
    
    LCustomer.Caption = iItems
End Sub

Private Sub AddLLtoPoles()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSDwg As AcadSelectionSet
    Dim objSSMap As AcadSelectionSet
    Dim objPoleDwg As AcadBlockReference
    Dim objPoleMap As AcadBlockReference
    Dim vAttDwg, vAttMap As Variant
    Dim iPoles, iXfer As Integer
    
    Dim dN, dE As Double
    Dim vLL As Variant
    
    iPoles = 0
    iXfer = 0
        
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSDwg = ThisDrawing.SelectionSets.Add("objSSDwg")
    objSSDwg.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    
    Set objSSMap = ThisDrawing.SelectionSets.Add("objSSMap")
    objSSMap.Select acSelectionSetWindow, vMapLL, vMapUR, filterType, filterValue
    
    If objSSMap.count < 1 Then GoTo Exit_Sub
    
    For Each objPoleMap In objSSMap
        vAttMap = objPoleMap.GetAttributes
        If vAttMap(0).TextString = "POLE" Then GoTo Next_objDwgPole
        If vAttMap(0).TextString = "" Then GoTo Next_objDwgPole
        
        If Not vAttMap(7).TextString = "" Then
            If cbOverwrite.Value = False Then GoTo Next_objDwgPole
        End If
        
        dE = objPoleMap.InsertionPoint(0)
        dN = objPoleMap.InsertionPoint(1)
        vLL = TN83FtoLL(CDbl(dN), CDbl(dE))
        
        vAttMap(7).TextString = vLL(0) & "," & vLL(1)
        objPoleMap.Update
        iPoles = iPoles + 1
        
        If cbIncludeDWG.Value = False Then GoTo Next_objDwgPole
        If objSSDwg.count < 1 Then GoTo Next_objDwgPole
        
        For Each objPoleDwg In objSSDwg
            vAttDwg = objPoleDwg.GetAttributes
            If vAttDwg(0).TextString = vAttMap(0).TextString Then
                vAttDwg(7).TextString = vAttMap(7).TextString
                objPoleDwg.Update
                iXfer = iXfer + 1
                GoTo Next_objDwgPole
            End If
            
        Next objPoleDwg
Next_objDwgPole:
    Next objPoleMap
    
Exit_Sub:
    objSSDwg.Clear
    objSSDwg.Delete
    
    objSSMap.Clear
    objSSMap.Delete
    
    LPole.Caption = iPoles & " / " & iXfer
End Sub

Private Function AddLLToBuried(strName As String)
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSDwg As AcadSelectionSet
    Dim objSSMap As AcadSelectionSet
    Dim objBlockDwg As AcadBlockReference
    Dim objBlockMap As AcadBlockReference
    Dim vAttDwg, vAttMap As Variant
    Dim strNumber, strOmit As String
    Dim strDwgLatLong, strMapLatLong As String
    Dim iItems, iXfer As Integer
    
    iItems = 0
    iXfer = 0
    strOmit = Replace(UCase(strName), "S", "")
        
    grpCode(0) = 2
    grpValue(0) = strName
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSDwg = ThisDrawing.SelectionSets.Add("objSSDwg")
    objSSDwg.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    
    Set objSSMap = ThisDrawing.SelectionSets.Add("objSSMap")
    objSSMap.Select acSelectionSetWindow, vMapLL, vMapUR, filterType, filterValue
    
    If objSSMap.count < 1 Then GoTo Exit_Sub
    
    For Each objBlockMap In objSSMap
        vAttMap = objBlockMap.GetAttributes
        If vAttMap(0).TextString = strOmit Then GoTo Next_objDwgPole
        If vAttMap(0).TextString = "" Then GoTo Next_objDwgPole
        
        If Not vAttMap(3).TextString = "" Then
            If cbOverwrite.Value = False Then GoTo Next_objDwgPole
        End If
        
        dE = objBlockMap.InsertionPoint(0)
        dN = objBlockMap.InsertionPoint(1)
        vLL = TN83FtoLL(CDbl(dN), CDbl(dE))
        
        vAttMap(3).TextString = vLL(0) & "," & vLL(1)
        objBlockMap.Update
        iItems = iItems + 1
        
        If cbIncludeDWG.Value = False Then GoTo Next_objDwgPole
        If objSSDwg.count < 1 Then GoTo Next_objDwgPole
        
        For Each objBlockDwg In objSSDwg
            vAttDwg = objBlockDwg.GetAttributes
            If vAttDwg(0).TextString = vAttMap(0).TextString Then
                vAttDwg(3).TextString = vAttMap(3).TextString
                objBlockDwg.Update
                iXfer = iXfer + 1
                GoTo Next_objDwgPole
            End If
            
        Next objBlockDwg
Next_objDwgPole:
    Next objBlockMap
    
Exit_Sub:
    objSSDwg.Clear
    objSSDwg.Delete
    
    objSSMap.Clear
    objSSMap.Delete
    
    AddLLToBuried = iItems & " / " & iXfer
End Function

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
