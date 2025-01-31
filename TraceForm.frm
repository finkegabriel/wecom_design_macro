VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TraceForm 
   Caption         =   "Trace Count"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   OleObjectBlob   =   "TraceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TraceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bResidence As Boolean

Private Sub cbGet_Click()
    If tbName.Value = "" Then Exit Sub
    If tbCount.Value = "" Then Exit Sub
    
    If bResidence = True Then
        Call CustomerTrace
        
        Exit Sub
    End If
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine, vCable, vBlock, vItems, vCounts, vTemp As Variant
    Dim strName, strCount, strCable, strCustomer, strTest, strTemp As String
    Dim iCount, iLow, iHigh, iAtt As Integer
    Dim bFeeder As Boolean
    Dim vPnt1, vPnt2 As Variant
    
    On Error Resume Next
    
    lbSpans.Clear
    bFeeder = False
    
    iCount = CInt(tbCount.Value)
    strName = tbName.Value
    
    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    Select Case cbWindow.Value
        Case "All"
            objSS.Select acSelectionSetAll, , , filterType, filterValue
        Case Else
            Me.Hide
            vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
            vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
            
            objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    End Select
    
Get_Feeder:
    Err = 0
    
    For Each objBlock In objSS
        If objBlock.Name = "sPole" Then
            iAtt = 25
        Else
            iAtt = 5
        End If
        
        vAttList = objBlock.GetAttributes
        
        If vAttList(iAtt).TextString = "" Then GoTo Next_objBlock
        If InStr(vAttList(iAtt).TextString, strName) < 1 Then GoTo Next_objBlock
        
        vLine = Split(vAttList(iAtt).TextString, vbCr)
        
        For i = 0 To UBound(vLine)
            If InStr(vLine(i), strName) < 1 Then GoTo Next_I
            
            vCable = Split(vLine(i), " / ")
            vBlock = Split(vCable(1), " + ")
            For j = 0 To UBound(vBlock)
                If InStr(vBlock(j), strName) < 1 Then GoTo Next_J
                
                vItems = Split(vBlock(j), ": ")
                vCounts = Split(vItems(1), "-")
                iLow = CInt(vCounts(0)) - 1
                If UBound(vCounts) > 0 Then
                    iHigh = CInt(vCounts(1)) + 1
                Else
                    iHigh = CInt(vCounts(0)) + 1
                End If
                
                If iCount > iLow And iCount < iHigh Then GoTo Found_Count
            
Next_J:
            Next j
            
Next_I:
        Next i
        
        GoTo Next_objBlock
                
Found_Count:
        If Not Err = 0 Then GoTo Next_objBlock
        
        lbSpans.AddItem objBlock.Name
        lbSpans.List(lbSpans.ListCount - 1, 1) = vAttList(0).TextString
        lbSpans.List(lbSpans.ListCount - 1, 2) = ""
        If UBound(vItems) > 1 Then lbSpans.List(lbSpans.ListCount - 1, 2) = vItems(2)
        lbSpans.List(lbSpans.ListCount - 1, 3) = ""
        lbSpans.List(lbSpans.ListCount - 1, 4) = ""
        lbSpans.List(lbSpans.ListCount - 1, 5) = ""
        lbSpans.List(lbSpans.ListCount - 1, 6) = ""
        lbSpans.List(lbSpans.ListCount - 1, 7) = ""
        lbSpans.List(lbSpans.ListCount - 1, 8) = objBlock.InsertionPoint(0) & "," & objBlock.InsertionPoint(1)
        
        vItems = Split(vCable(0), ": ")
        vCounts = Split(vItems(1), ")")
        strCable = vCounts(0) & ")"
        
        iAtt = iAtt + 1
        If Not vAttList(iAtt).TextString = "" Then
            vTemp = Split(vAttList(iAtt).TextString, "] ")
            vBlock = Split(vTemp(1), " + ")
            For i = 0 To UBound(vBlock)
                vItems = Split(vBlock(i), ": ")
                If vItems(0) = strName Then
                    vCounts = Split(vItems(1), "-")
                    iLow = CInt(vCounts(0)) - 1
                    If UBound(vCounts) > 0 Then
                        iHigh = CInt(vCounts(1)) + 1
                    Else
                        iHigh = CInt(vCounts(0)) + 1
                    End If
                
                    If iCount > iLow And iCount < iHigh Then
                        lbSpans.List(lbSpans.ListCount - 1, 5) = "1"
                        
                        If InStr(vTemp(1), tbName.Value) > 0 And bFeeder = True Then
                            lbSpans.List(lbSpans.ListCount - 1, 5) = "2"
                            lbSpans.List(lbSpans.ListCount - 1, 6) = tbSplitter.Value
                            If InStr(cbSplitter.Value, "Hub") > 0 Then lbSpans.List(lbSpans.ListCount - 1, 7) = "2"
                            
                            GoTo Done_With_Splices
                        End If
                    End If
                End If
Next_Splice:
            Next i
        End If
        
Done_With_Splices:
        
        iAtt = iAtt + 1
        If Not vAttList(iAtt).TextString = "" Then
            vLine = Split(vAttList(iAtt).TextString, ";;")
            For i = 0 To UBound(vLine)
                If InStr(vLine(i), strCable) > 0 Then
                    If InStr(vLine(i), "LOOP") > 0 Then
                        lbSpans.List(lbSpans.ListCount - 1, 4) = "100"
                    Else
                        vItems = Split(vLine(i), "=")
                        If InStr(vItems(0), "HA") > 0 Then GoTo Next_Unit
                        strTemp = Replace(vItems(1), "'", "")
                        If lbSpans.List(lbSpans.ListCount - 1, 3) = "" Then
                            lbSpans.List(lbSpans.ListCount - 1, 3) = strTemp
                        Else
                            lbSpans.List(lbSpans.ListCount - 1, 3) = CLng(lbSpans.List(lbSpans.ListCount - 1, 3)) + CLng(strTemp)
                        End If
                    End If
                End If
Next_Unit:
            Next i
        End If
        
        
        
        
Next_objBlock:
    Err = 0
    Next objBlock
    
    If InStr(strName, "-") > 0 Then
        vLine = Split(strName, "-")
        strName = vLine(0)
        iCount = CInt(vLine(1))
        bFeeder = True
        
        GoTo Get_Feeder
    End If
    
    '<------------------------ Get Customer with Count
    
    Err = 0
    
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
    objSS.Clear
    
    Select Case cbWindow.Value
        Case "All"
            objSS.Select acSelectionSetAll, , , filterType, filterValue
        Case Else
            objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    End Select
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        If vAttList(4).TextString = "" Then GoTo Next_Customer
        strCustomer = tbName.Value & ": " & tbCount.Value
        
        vLine = Split(vAttList(4).TextString, " - ")
        strTest = Replace(vLine(1), "(", "")
        strTest = Replace(strTest, ")", "")
        
        If strTest = strCustomer Then
            lbSpans.AddItem vAttList(0).TextString
            lbSpans.List(lbSpans.ListCount - 1, 1) = vAttList(1).TextString & " " & vAttList(2).TextString
            lbSpans.List(lbSpans.ListCount - 1, 2) = vLine(0)
            lbSpans.List(lbSpans.ListCount - 1, 3) = ""
            lbSpans.List(lbSpans.ListCount - 1, 4) = ""
            lbSpans.List(lbSpans.ListCount - 1, 5) = ""
            lbSpans.List(lbSpans.ListCount - 1, 6) = ""
            lbSpans.List(lbSpans.ListCount - 1, 7) = "1"
        End If
        
Next_Customer:
    Next objBlock
    
Exit_Sub:
    Call GetTotals
    Call OrderList
    
    objSS.Clear
    objSS.Delete
    
    If Not cbWindow.Value = "All" Then Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbSplitter_Change()
    Select Case cbSplitter.Value
        Case "16"
            tbSplitter.Value = "14.5"
            tbConnector.Value = "1"
        Case "32"
            tbSplitter.Value = "18.0"
            tbConnector.Value = "1"
        Case "64"
            tbSplitter.Value = "21.5"
            tbConnector.Value = "1"
        Case "16 Hub"
            tbSplitter.Value = "14.5"
            tbConnector.Value = "3"
        Case "32 Hub"
            tbSplitter.Value = "18.0"
            tbConnector.Value = "3"
        Case Else
            tbSplitter.Value = "0.0"
            tbConnector.Value = "1"
    End Select
End Sub

Private Sub Label11_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    Dim vLine, vTemp As Variant
    Dim strTest As String
    
    On Error Resume Next
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Customer:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    
    If Not objBlock.Name = "Customer" Then GoTo Exit_Sub
    
    vAttList = objBlock.GetAttributes
    If vAttList(4).TextString = "" Then GoTo Exit_Sub
    
    vLine = Split(vAttList(4).TextString, " - ")
        
    strTest = Replace(vLine(1), "(", "")
    strTest = Replace(strTest, ")", "")
    vTemp = Split(strTest, ": ")
    tbName.Value = vTemp(0)
    tbCount.Value = vTemp(1)
    
    GoTo Exit_Sub
    
    lbSpans.Clear
    
    lbSpans.AddItem vAttList(0).TextString
    lbSpans.List(lbSpans.ListCount - 1, 1) = vAttList(1).TextString & " " & vAttList(2).TextString
    lbSpans.List(lbSpans.ListCount - 1, 2) = vLine(0)
    lbSpans.List(lbSpans.ListCount - 1, 3) = ""
    lbSpans.List(lbSpans.ListCount - 1, 4) = ""
    lbSpans.List(lbSpans.ListCount - 1, 5) = ""
    lbSpans.List(lbSpans.ListCount - 1, 6) = ""
    lbSpans.List(lbSpans.ListCount - 1, 7) = "1"
    
    'bResidence = True
    
Exit_Sub:
    Me.show
End Sub

Private Sub Label6_Click()
    If lbSpans.ListIndex < 0 Then Exit Sub
    
    If lbSpans.List(lbSpans.ListIndex, 5) = "" Then
        lbSpans.List(lbSpans.ListIndex, 5) = "1"
    Else
        lbSpans.List(lbSpans.ListIndex, 5) = CInt(lbSpans.List(lbSpans.ListIndex, 5)) + 1
    End If
End Sub

Private Sub Label7_Click()
    If lbSpans.ListIndex < 0 Then Exit Sub
    
    If lbSpans.List(lbSpans.ListIndex, 5) = "" Then
        lbSpans.List(lbSpans.ListIndex, 5) = "2"
    lbSpans.List(lbSpans.ListIndex, 6) = tbSplitter.Value
            
    If InStr(cbSplitter.Value, "Hub") > 0 Then lbSpans.List(lbSpans.ListIndex, 7) = "2"
End Sub

Private Sub lbSpans_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vCoords, vAttList As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    Dim iIndex As Integer
    
    Me.Hide
    
    vCoords = Split(lbSpans.List(lbSpans.ListIndex, 8), ",")
    
    viewCoordsB(0) = vCoords(0) - 150
    viewCoordsB(1) = vCoords(1) - 150
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 150
    viewCoordsE(1) = vCoords(1) + 150
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Me.show
End Sub

Private Sub lbSpans_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyS
            lbSpans.List(lbSpans.ListIndex, 5) = "2"
            lbSpans.List(lbSpans.ListIndex, 6) = tbSplitter.Value
            
            If InStr(cbSplitter.Value, "Hub") > 0 Then lbSpans.List(lbSpans.ListIndex, 7) = "2"
        Case vbKeyX
            If lbSpans.List(lbSpans.ListIndex, 5) = "" Then
                lbSpans.List(lbSpans.ListIndex, 5) = "1"
            Else
                lbSpans.List(lbSpans.ListIndex, 5) = CInt(lbSpans.List(lbSpans.ListIndex, 5)) + 1
            End If
        Case vbKeyD
            lbSpans.List(lbSpans.ListIndex, 5) = ""
            lbSpans.List(lbSpans.ListIndex, 6) = ""
            lbSpans.List(lbSpans.ListIndex, 7) = ""
        Case vbKeyC
            If lbSpans.List(lbSpans.ListIndex, 7) = "" Then
                lbSpans.List(lbSpans.ListIndex, 7) = "1"
            Else
                lbSpans.List(lbSpans.ListIndex, 7) = CInt(lbSpans.List(lbSpans.ListIndex, 7)) + 1
            End If
    End Select
    
    Call GetTotals
End Sub

Private Sub tbDB1310_Change()
    If CDbl(tbDB1310.Value) > CDbl(tbMax.Value) Then
        Label19.ForeColor = &HFF&
    Else
        Label19.ForeColor = &H80000012
    End If
End Sub

Private Sub tbDB1550_Change()
    If CDbl(tbDB1550.Value) > CDbl(tbMax.Value) Then
        Label20.ForeColor = &HFF&
    Else
        Label20.ForeColor = &H80000012
    End If
End Sub

Private Sub UserForm_Initialize()
    lbSpans.ColumnCount = 9
    lbSpans.ColumnWidths = "72;120;120;36;36;36;36;36;6"
    
    cbSplitter.AddItem ""
    cbSplitter.AddItem "16"
    cbSplitter.AddItem "32"
    cbSplitter.AddItem "64"
    cbSplitter.AddItem "16 Hub"
    cbSplitter.AddItem "32 Hub"
    
    cbSplitter.Value = "32"
    tbSplitter.Value = "18.0"
    
    cbWindow.AddItem "All"
    cbWindow.AddItem "Window"
    cbWindow.Value = "All"
    
    bResidence = False
End Sub

Private Sub GetTotals()
    If lbSpans.ListCount < 1 Then Exit Sub
    
    Dim lSpans, lCoil As Long
    Dim iSplice, iConnect As Integer
    Dim dSplit As Double
    Dim d1310, d1550, dBoth, dTemp As Double
    
    lSpans = 0
    lCoil = 0
    
    iSplice = 0
    iSplit = 0
    iConnect = 0
    
    For i = 0 To lbSpans.ListCount - 1
        If Not lbSpans.List(i, 3) = "" Then lSpans = lSpans + CLng(lbSpans.List(i, 3))
        If Not lbSpans.List(i, 4) = "" Then lCoil = lCoil + CLng(lbSpans.List(i, 4))
        
        If Not lbSpans.List(i, 5) = "" Then iSplice = iSplice + CInt(lbSpans.List(i, 5))
        If Not lbSpans.List(i, 6) = "" Then dSplit = dSplit + CDbl(lbSpans.List(i, 6))
        If Not lbSpans.List(i, 7) = "" Then iConnect = iConnect + CInt(lbSpans.List(i, 7))
    Next i
    
    tbSpan.Value = lSpans
    tbCoiled.Value = lCoil
    
    tbSpliced.Value = iSplice
    tbSplit.Value = dSplit
    tbConnector.Value = iConnect
    
    dBoth = iSplice * CDbl(tbPerSplice.Value)
    dBoth = dBoth + dSplit
    dBoth = dBoth + iConnect * CDbl(tbPerConnector.Value)
    
    d1310 = dBoth + (lSpans + lCoil) / 1000 * CDbl(tbPer1310.Value)
    d1550 = dBoth + (lSpans + lCoil) / 1000 * CDbl(tbPer1550.Value)
    
    tbDB1310.Value = CLng(d1310 * 1000) / 1000
    tbDB1550.Value = CLng(d1550 * 1000) / 1000
    
    dTemp = CDbl(tbMax.Value) - dBoth
    tbMax1310.Value = CLng(dTemp / CDbl(tbPer1310.Value) * 1000) / 1000
    tbMax1550.Value = CLng(dTemp / CDbl(tbPer1550.Value) * 1000) / 1000
    
    tbListCount.Value = lbSpans.ListCount
End Sub

Private Sub OrderList()
    If lbSpans.ListCount < 2 Then Exit Sub
    
    Dim strPrevious, strLine As String
    Dim strList() As String
    Dim iCount As Integer
    Dim vLine As Variant
    
    iCount = lbSpans.ListCount - 1
    ReDim strList(iCount) As String
    
    Select Case lbSpans.List(iCount, 0)
        Case "sPole", "sHH", "sPed"
            'strPrevious = lbSpans.List(lbSpans.ListIndex, 2)
            Exit Sub
        Case Else
            strPrevious = lbSpans.List(iCount, 2)
            strLine = lbSpans.List(iCount, 0) & vbTab & lbSpans.List(iCount, 1) & vbTab & lbSpans.List(iCount, 2)
            strLine = strLine & vbTab & lbSpans.List(iCount, 3) & vbTab & lbSpans.List(iCount, 4) & vbTab & lbSpans.List(iCount, 5)
            strLine = strLine & vbTab & lbSpans.List(iCount, 6) & vbTab & lbSpans.List(iCount, 7) & vbTab & lbSpans.List(iCount, 8)
            strList(iCount) = strLine
            
            iCount = iCount - 1
    End Select
    
    While iCount > -1
        For i = 0 To lbSpans.ListCount - 1
            If lbSpans.List(i, 1) = strPrevious Then
                strPrevious = lbSpans.List(i, 2)
                strLine = lbSpans.List(i, 0) & vbTab & lbSpans.List(i, 1) & vbTab & lbSpans.List(i, 2)
                strLine = strLine & vbTab & lbSpans.List(i, 3) & vbTab & lbSpans.List(i, 4) & vbTab & lbSpans.List(i, 5)
                strLine = strLine & vbTab & lbSpans.List(i, 6) & vbTab & lbSpans.List(i, 7) & vbTab & lbSpans.List(i, 8)
                strList(iCount) = strLine
            
                iCount = iCount - 1
                GoTo Find_Next
            End If
        Next i
        
        MsgBox "Unable to find previous pole  " & strPrevious
        Exit Sub
Find_Next:
    Wend
    
    lbSpans.Clear
    
    For i = UBound(strList) To 0 Step -1
        vLine = Split(strList(i), vbTab)
        
        lbSpans.AddItem vLine(0), 0
        lbSpans.List(0, 1) = vLine(1)
        lbSpans.List(0, 2) = vLine(2)
        lbSpans.List(0, 3) = vLine(3)
        lbSpans.List(0, 4) = vLine(4)
        lbSpans.List(0, 5) = vLine(5)
        lbSpans.List(0, 6) = vLine(6)
        lbSpans.List(0, 7) = vLine(7)
        lbSpans.List(0, 8) = vLine(8)
    Next i
End Sub

Private Sub CustomerTrace()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strPrevious As String
    Dim iAtt As Integer
    'Dim vLine, vCable, vBlock, vItems, vCounts As Variant
    'Dim strName, strCount, strCable, strCustomer, strTest, strTemp As String
    'Dim iCount, iLow, iHigh, iAtt As Integer
    
    On Error Resume Next
    
    'iCount = CInt(tbCount.Value)
    'strName = tbName.Value
    
    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
Start_Over:
    Err = 0
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        If Not vAttList(0).TextString = strPrevious Then GoTo Next_objBlock
        
        
        
        
        
        
        
        
        
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    Call GetTotals
    
    objSS.Clear
    objSS.Delete
End Sub
