VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExtraHeightForm 
   Caption         =   "Extra Heights"
   ClientHeight    =   9120.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18690
   OleObjectBlob   =   "ExtraHeightForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExtraHeightForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iListIndex As Integer

Private Sub cbCompanies_Change()
    If cbCompanies.Value = "" Then Exit Sub
    If lbPoles.ListCount < 1 Then Exit Sub
    
    Dim vLine, vItem, vTemp As Variant
    Dim strCompany As String
    Dim iOwner, iMR As Integer
    Dim iTotal, iIndex As Integer
    Dim iCompany, iMRLoc, iMRTotal As Integer
    
    lbATT.Clear
    
    strCompany = cbCompanies.Value
    For i = 0 To lbPoles.ListCount - 1
        iOwner = 0
        iMR = 0
        
        If lbPoles.List(i, 1) = strCompany Then iOwner = 1
        
        vLine = Split(lbPoles.List(i, 5), " + ")
        For j = 0 To UBound(vLine)
            vItem = Split(vLine(j), "=")
            If vItem(0) = strCompany Then
                vTemp = Split(vItem(1), " ")
                For k = 0 To UBound(vTemp)
                    If InStr(vTemp(k), ")") > 0 Then
                        iMR = iMR + 1
                    End If
                Next k
            End If
        Next j
        
        iTotal = iOwner + iMR
        
        If iTotal > 0 Then
            lbATT.AddItem lbPoles.List(i, 0)
            iIndex = lbATT.ListCount - 1
            
            lbATT.List(iIndex, 1) = iOwner
            lbATT.List(iIndex, 2) = iMR
            lbATT.List(iIndex, 3) = i
        End If
    Next i
    
    iCompany = 0
    iMRLoc = 0
    iMRTotal = 0
    
    For i = 0 To lbATT.ListCount - 1
        iCompany = iCompany + CInt(lbATT.List(i, 1))
        If CInt(lbATT.List(i, 2)) > 0 Then
            iMRLoc = iMRLoc + 1
            iMRTotal = iMRTotal + CInt(lbATT.List(i, 2))
        End If
    Next i
    
    tbAttVisit.Value = lbATT.ListCount
    tbAtt.Value = iCompany
    tbATTMR.Value = iMRLoc
    tbATTMRTotal.Value = iMRTotal
End Sub

Private Sub cbGet_Click()
    Me.Hide
    Call GetPoles
    
    Call GetData
    Me.show
End Sub

Private Sub AddDataToTextBox(strCompany As String)
    Dim strAll, strTest As String
    Dim strOne, strTwo As String
    Dim vLine, vItem As Variant
    
    If lbMR.ListCount = 0 Then GoTo Add_To_List
    
    For i = 0 To lbMR.ListCount - 1
        If lbMR.List(i, 0) = "" Then GoTo Skip_This
        strTemp = lbMR.List(i, 0)
        
        If StrComp(strTemp, strCompany) = 0 Then
            lbMR.List(i, 1) = CInt(lbMR.List(i, 1)) + 1
            
            Exit Sub
        End If
Skip_This:
    Next i
    
Add_To_List:
    
    lbMR.AddItem strCompany
    lbMR.List(lbMR.ListCount - 1, 1) = 1
End Sub

Private Sub cbOwners_Change()
    If cbOwners.Value = "" Then Exit Sub
    If lbData.ListCount < 1 Then Exit Sub
    
    Dim vLine, vItem As Variant
    Dim strOwner, strHC As String
    Dim iTotal As Integer
    
    lbEHReport.Clear
    
    strOwner = cbOwners.Value
    For i = 0 To lbData.ListCount - 1
        If lbData.List(i, 0) = strOwner Then
            strHC = lbData.List(i, 2)
            If InStr(strHC, ")") > 0 Then
                vLine = Split(strHC, ")")
                strHC = vLine(1)
            End If
            
            If lbEHReport.ListCount > 0 Then
                For j = 0 To lbEHReport.ListCount - 1
                    If lbEHReport.List(j, 0) = strHC Then
                        lbEHReport.List(j, 1) = CInt(lbEHReport.List(j, 1)) + 1
                        GoTo Existing_Owner
                    End If
                Next j
            End If
            
            lbEHReport.AddItem strHC
            lbEHReport.List(lbEHReport.ListCount - 1, 1) = "1"
Existing_Owner:
            
        End If
    Next i
    
    iTotal = 0
    
    For i = 0 To lbEHReport.ListCount - 1
        iTotal = iTotal + CInt(lbEHReport.List(i, 1))
    Next i
    
    tbTotalEHCompany.Value = iTotal
    
    Call SortEH
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRun_Click()
    Dim i406, i405, i404, i403, i402, i401 As Integer
    Dim i456, i455, i454, i453, i452, i451 As Integer
    Dim i505, i504, i503, i502, i501 As Integer
    Dim i555, i554, i553, i552, i551 As Integer
    Dim i604, i603, i602, i601 As Integer
    Dim i653, i652, i651 As Integer
    Dim i702, i701 As Integer
    Dim vLine, vHC As Variant
    Dim strHC, strLine As String
    
    i406 = 0: i405 = 0: i404 = 0: i403 = 0: i402 = 0: i401 = 0
    i456 = 0: i455 = 0: i454 = 0: i453 = 0: i452 = 0: i451 = 0
    i505 = 0: i504 = 0: i503 = 0: i502 = 0: i501 = 0
    i555 = 0: i554 = 0: i553 = 0: i552 = 0: i551 = 0
    i604 = 0: i603 = 0: i602 = 0: i601 = 0
    i653 = 0: i652 = 0: i651 = 0
    i702 = 0: i701 = 0
    
    On Error Resume Next
    
    For n = 0 To lbData.ListCount - 1
        vLine = Split(lbData.List(n), vbTab)
        
        If Len(vLine(2)) > 6 Then
            vHC = Split(vLine(2), ") ")
            vLine(2) = vHC(UBound(vHC))
        End If
        
        vHC = Split(vLine(2), "-")
        Select Case Left(vHC(0), 1)
            Case "S", "C"
                vHC(0) = Right(vHC(0), 2)
                vHC(1) = 1
            'Case "("
            'Case Else
        End Select
        
        'If UBound(vHC) < 1 Then MsgBox vLine(2)
        
        strHC = vHC(0) & vHC(1)
        Select Case strHC
            Case "406"
                i406 = i406 + 1
            Case "405"
                i405 = i405 + 1
            Case "404"
                i404 = i404 + 1
            Case "403"
                i403 = i403 + 1
            Case "402"
                i402 = i402 + 1
            Case "401"
                i401 = i401 + 1
                
            Case "456"
                i456 = i456 + 1
            Case "455"
                i455 = i455 + 1
            Case "454"
                i454 = i454 + 1
            Case "453"
                i453 = i453 + 1
            Case "452"
                i452 = i452 + 1
            Case "451"
                i451 = i451 + 1
                
            Case "505"
                i505 = i505 + 1
            Case "504"
                i504 = i504 + 1
            Case "503"
                i503 = i503 + 1
            Case "502"
                i502 = i502 + 1
            Case "501"
                i501 = i501 + 1
                
            Case "555"
                i555 = i555 + 1
            Case "554"
                i554 = i554 + 1
            Case "553"
                i553 = i553 + 1
            Case "552"
                i552 = i552 + 1
            Case "551"
                i551 = i551 + 1
                
            Case "604"
                i604 = i604 + 1
            Case "603"
                i603 = i603 + 1
            Case "602"
                i602 = i602 + 1
            Case "601"
                i601 = i601 + 1
                
            Case "653"
                i653 = i653 + 1
            Case "652"
                i652 = i652 + 1
            Case "651"
                i651 = i651 + 1
                
            Case "702"
                i702 = i702 + 1
            Case "701"
                i701 = i701 + 1
        End Select
    Next n
    
    strLine = ""
    
    If Not i406 = 0 Then strLine = strLine & "40-6" & vbTab & i406 & vbCr
    If Not i405 = 0 Then strLine = strLine & "40-5" & vbTab & i405 & vbCr
    If Not i404 = 0 Then strLine = strLine & "40-4" & vbTab & i404 & vbCr
    If Not i403 = 0 Then strLine = strLine & "40-3" & vbTab & i403 & vbCr
    If Not i402 = 0 Then strLine = strLine & "40-2" & vbTab & i402 & vbCr
    If Not i401 = 0 Then strLine = strLine & "40-1" & vbTab & i401 & vbCr
    
    If Not i456 = 0 Then strLine = strLine & "45-6" & vbTab & i456 & vbCr
    If Not i455 = 0 Then strLine = strLine & "45-5" & vbTab & i455 & vbCr
    If Not i454 = 0 Then strLine = strLine & "45-4" & vbTab & i454 & vbCr
    If Not i453 = 0 Then strLine = strLine & "45-3" & vbTab & i453 & vbCr
    If Not i452 = 0 Then strLine = strLine & "45-2" & vbTab & i452 & vbCr
    If Not i451 = 0 Then strLine = strLine & "45-1" & vbTab & i451 & vbCr
    
    If Not i505 = 0 Then strLine = strLine & "50-5" & vbTab & i505 & vbCr
    If Not i504 = 0 Then strLine = strLine & "50-4" & vbTab & i504 & vbCr
    If Not i503 = 0 Then strLine = strLine & "50-3" & vbTab & i503 & vbCr
    If Not i502 = 0 Then strLine = strLine & "50-2" & vbTab & i502 & vbCr
    If Not i501 = 0 Then strLine = strLine & "50-1" & vbTab & i501 & vbCr
    
    If Not i555 = 0 Then strLine = strLine & "55-5" & vbTab & i555 & vbCr
    If Not i554 = 0 Then strLine = strLine & "55-4" & vbTab & i554 & vbCr
    If Not i553 = 0 Then strLine = strLine & "55-3" & vbTab & i553 & vbCr
    If Not i552 = 0 Then strLine = strLine & "55-2" & vbTab & i552 & vbCr
    If Not i551 = 0 Then strLine = strLine & "55-1" & vbTab & i551 & vbCr
    
    If Not i604 = 0 Then strLine = strLine & "60-4" & vbTab & i604 & vbCr
    If Not i603 = 0 Then strLine = strLine & "60-3" & vbTab & i603 & vbCr
    If Not i602 = 0 Then strLine = strLine & "60-2" & vbTab & i602 & vbCr
    If Not i601 = 0 Then strLine = strLine & "60-1" & vbTab & i601 & vbCr
    
    If Not i653 = 0 Then strLine = strLine & "65-3" & vbTab & i653 & vbCr
    If Not i652 = 0 Then strLine = strLine & "65-2" & vbTab & i652 & vbCr
    If Not i651 = 0 Then strLine = strLine & "65-1" & vbTab & i651 & vbCr
    
    If Not i702 = 0 Then strLine = strLine & "70-2" & vbTab & i702 & vbCr
    If Not i701 = 0 Then strLine = strLine & "70-1" & vbTab & i701 & vbCr
    
    tbReport.Value = strLine
End Sub

Private Sub cbTotalSelected_Click()
    If lbEHReport.ListCount < 0 Then Exit Sub
    
    Dim iTotal As Integer
    
    iTotal = 0
    
    For i = 0 To lbEHReport.ListCount - 1
        If lbEHReport.Selected(i) = True Then iTotal = iTotal + CInt(lbEHReport.List(i, 1))
    Next i
    
    tbTotalEHCompany.Value = iTotal
End Sub

Private Sub cbUpdate_Click()
    Dim strLine As String
    
    strLine = tbAttach.Value & vbTab & tbPole.Value & vbTab & tbHC.Value
    lbData.AddItem strLine, iListIndex
    lbData.RemoveItem (iListIndex + 1)
    
    cbUpdate.Enabled = False
    cbGet.Enabled = True
    cbRun.Enabled = True
    tbPole.Value = ""
    tbAttach.Value = ""
    tbHC.Value = ""
End Sub

Private Sub Label2_Click()
    Dim obrPole As AcadBlockReference
    Dim obrAttach As AcadBlockReference
    Dim entBlock As AcadEntity
    Dim objSSFix As AcadSelectionSet
    Dim strPole As String
    Dim vAttList, basePnt As Variant
    
    Me.Hide
    
  On Error Resume Next
    Err = 0
    
    ThisDrawing.Utility.GetEntity entBlock, basePnt, "Select Pole: "
    If TypeOf entBlock Is AcadBlockReference Then
        Set obrPole = entBlock
    Else
        MsgBox "Not a valid pole."
        Me.show
        Exit Sub
    End If
    
    vAttList = obrPole.GetAttributes
    
    If Not Err = 0 Then
        Me.show
        Exit Sub
    End If
    
    strPole = vAttList(0).TextString
    
  
    Set objSSFix = ThisDrawing.SelectionSets.Add("objSSFix")
    objSSFix.SelectOnScreen
    For Each entBlock In objSSFix
        If TypeOf entBlock Is AcadBlockReference Then
            Set obrAttach = entBlock
            
            vAttList = obrAttach.GetAttributes
            vAttList(0).TextString = strPole
            obrAttach.Update
            'If obrPole.Name = "pole_attach" Then
            '    attList = obrBlock.GetAttributes
            '    str1 = attList(2).TextString & vbTab & attList(3).TextString & vbTab & attList(4).TextString
            '    lbAttachments.AddItem str1
            '    If CDbl(attList(1).TextString) < dTest Then
            '        vCoordsAttach = obrBlock.InsertionPoint
            '        dTest = CDbl(attList(1).TextString)
            '        dPosition = dTest
            '    End If
            '    tbUnitedPole = attList(0).TextString
            '    cbScale.Value = obrBlock.XScaleFactor * 100
            'End If
        End If
    Next entBlock
    
    objSSFix.Clear
    objSSFix.Delete
    
    Me.show
End Sub

Private Sub lbATT_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim iIndex As Integer
    Dim strLine As String
    
    iIndex = CInt(lbATT.List(lbATT.ListIndex, 3))
    
    strLine = Replace(lbPoles.List(iIndex, 5), " + ", vbCr)
    
    MsgBox strLine
End Sub

Private Sub lbData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vLine As Variant
    
    cbUpdate.Enabled = True
    cbGet.Enabled = False
    cbRun.Enabled = False
    iListIndex = lbData.ListIndex
    
    vLine = Split(lbData.List(iListIndex), vbTab)
    tbPole.Value = vLine(1)
    tbAttach.Value = vLine(0)
    tbHC.Value = vLine(2)
End Sub

Private Sub lbPoles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vCoords, vAttList As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    vCoords = Split(lbPoles.List(lbPoles.ListIndex, 6), ",")
    
    viewCoordsB(0) = vCoords(0) - 200
    viewCoordsB(1) = vCoords(1) - 200
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 200
    viewCoordsE(1) = vCoords(1) + 200
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Me.show
End Sub

Private Sub lbPoles_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Dim iIndex As Integer
            Dim strLine As String
    
            iIndex = lbPoles.ListIndex
    
            strLine = Replace(lbPoles.List(iIndex, 5), " + ", vbCr)
    
            MsgBox strLine
            
            If Not iIndex = lbPoles.ListCount - 1 Then lbPoles.ListIndex = iIndex + 1
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbMR.ColumnCount = 2
    lbMR.ColumnWidths = "70;25"
    
    lbATT.ColumnCount = 4
    lbATT.ColumnWidths = "100;48;38;10"
    
    lbOwners.ColumnCount = 4
    lbOwners.ColumnWidths = "72;48;48;36"
    
    lbData.ColumnCount = 5
    lbData.ColumnWidths = "48;120;48;36;10"
    
    lbPoles.ColumnCount = 7
    lbPoles.ColumnWidths = "120;48;48;48;24;170;10"
    
    lbEHReport.ColumnCount = 2
    lbEHReport.ColumnWidths = "48;36"
    
    cbClient.AddItem "UTC"
    cbClient.AddItem "LOR"
    cbClient.AddItem "LTC"
    cbClient.AddItem "TDS"
    cbClient.AddItem ""
    cbClient.Value = "UTC"
End Sub

Private Sub GetPoles()
    Dim vPnt1, vPnt2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim filterType, filterValue As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vItem, vTemp As Variant
    Dim iHeight As Integer
    Dim iIndex As Integer
    Dim strMR, strPWR, strLine As String
    
    On Error Resume Next
    
    cbCompanies.Clear
    
    Err = 0
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then
        Me.show
        Exit Sub
    End If
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    'MsgBox objSS.count
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        If vAttList(0).TextString = "POLE" Then GoTo Next_objBlock
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        
        lbPoles.AddItem vAttList(0).TextString
        iIndex = lbPoles.ListCount - 1
        lbPoles.List(iIndex, 1) = vAttList(2).TextString
        lbPoles.List(iIndex, 2) = vAttList(5).TextString
        lbPoles.List(iIndex, 3) = vAttList(15).TextString
        lbPoles.List(iIndex, 4) = ""
        lbPoles.List(iIndex, 5) = ""
        lbPoles.List(iIndex, 6) = objBlock.InsertionPoint(0) & "," & objBlock.InsertionPoint(1)
        
        If cbCompanies.ListCount > 0 Then
            For j = 0 To cbCompanies.ListCount - 1
                If cbCompanies.List(j) = vAttList(2).TextString Then GoTo Found_Company
            Next j
        End If
        
        cbCompanies.AddItem vAttList(2).TextString
        
Found_Company:
        
        strPWR = ""
        For i = 9 To 14
            vItem = Split(vAttList(i).TextString, " ")
            For j = 0 To UBound(vItem)
                If InStr(vItem(j), ")") > 0 Then
                    If strPWR = "" Then
                        strPWR = "PWR=" & vItem(j)
                    Else
                        strPWR = strPWR & " " & vItem(j)
                    End If
                End If
            Next j
        Next i
        
        If strPWR = "" Then
            strLine = ""
        Else
            strLine = strPWR
        End If
        
        
        For i = 16 To 23
            strMR = ""
            vTemp = Split(UCase(vAttList(i).TextString), "=")
            If InStr(vTemp(0), cbClient.Value) > 0 Then
                lbPoles.List(iIndex, 4) = vTemp(1)
            End If
            
            If cbCompanies.ListCount > 0 Then
                For j = 0 To cbCompanies.ListCount - 1
                    If cbCompanies.List(j) = vTemp(0) Then
                        GoTo Found_Company2
                    End If
                Next j
            End If
        
            If InStr(vTemp(1), ")") > 0 Then cbCompanies.AddItem vTemp(0)
        
Found_Company2:
            
            vItem = Split(vTemp(1), " ")
            For j = 0 To UBound(vItem)
                If InStr(vItem(j), ")") > 0 Then
                    If strMR = "" Then
                        strMR = vTemp(0) & "=" & vItem(j)
                    Else
                        strMR = strMR & " " & vItem(j)
                    End If
                End If
                
                If InStr(vItem(j), "X") > 0 Then
                    If strMR = "" Then
                        strMR = vTemp(0) & "=" & vItem(j)
                    Else
                        strMR = strMR & " " & vItem(j)
                    End If
                End If
            Next j
            
            If Not strMR = "" Then
                If strLine = "" Then
                    strLine = strMR
                Else
                    strLine = strLine & " + " & strMR
                End If
            End If
        Next i
        
        lbPoles.List(iIndex, 5) = strLine
        
Next_objBlock:
    Next objBlock
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub GetData()
    Dim vLine, vItem, vData, vAttach As Variant
    Dim strLine, strOwner As String
    Dim iIndex, iHeight As Integer
    Dim iAttPole, iPower, iPWRPole As Integer
    Dim iCATV, iAtt, iATTCW, iTDS, iUTC, iDTC As Integer
    Dim iXO, iZAYO, iCLEC, iTelco As Integer
    Dim iCity, iTraffic, iOHG, iPoles As Integer
    Dim strMR As String
    
    On Error Resume Next
    
    lbMR.Clear
    lbOwners.Clear
    lbATT.Clear
    lbData.Clear
    tbPole.Value = ""
    tbAttach.Value = ""
    tbHC.Value = ""
    cbUpdate.Enabled = False
    cbOwners.Clear
    
    iHeight = 0
    iAttPole = 0: iPower = 0: iPWRPole = 0
    iCATV = 0: iAtt = 0: iATTCW = 0: iTDS = 0: iUTC = 0: iDTC = 0
    iXO = 0: iZAYO = 0: iCLEC = 0: iTelco = 0
    iCity = 0: iTraffic = 0: iOHG = 0: iPoles = 0
    
    For i = 0 To lbPoles.ListCount - 1
        ' Get Extra Height List data
        strOwner = lbPoles.List(i, 1)
        
        If lbOwners.ListCount > 0 Then
            For j = 0 To lbOwners.ListCount - 1
                If lbOwners.List(j, 0) = strOwner Then
                    iIndex = j
                    GoTo Existing_Owner
                End If
            Next j
        End If
            
        lbOwners.AddItem strOwner
        iIndex = lbOwners.ListCount - 1
        lbOwners.List(iIndex, 1) = "0"
        lbOwners.List(iIndex, 2) = "0"
        lbOwners.List(iIndex, 3) = "0"
Existing_Owner:
        
        'MsgBox "New: " & lbPoles.List(i, 3) & ">" & vbCr & "Exist: " & lbPoles.List(i, 4) & ">"
        
        If lbPoles.List(i, 4) = "" Then
            If Not lbPoles.List(i, 3) = "" Then
                lbOwners.List(iIndex, 3) = CInt(lbOwners.List(iIndex, 3)) + 1
            End If
        Else
            If lbPoles.List(i, 3) = "" Then
                lbOwners.List(iIndex, 1) = CInt(lbOwners.List(iIndex, 1)) + 1
            Else
                lbOwners.List(iIndex, 2) = CInt(lbOwners.List(iIndex, 2)) + 1
            End If
        End If
        
        If Not lbPoles.List(i, 3) = "" Then
            vAttach = Split(UCase(lbPoles.List(i, 3)), " ")
            
            For j = 0 To UBound(vAttach)
                strLine = Replace(vAttach(j), "T", "")
                vItem = Split(strLine, "-")
                
                iHeight = CInt(vItem(0)) * 12
                If UBound(vItem) > 0 Then iHeight = iHeight + CInt(vItem(1))
                If iHeight > 288 Then
                    lbData.AddItem lbPoles.List(i, 1)
                    lbData.List(lbData.ListCount - 1, 1) = lbPoles.List(i, 0)
                    lbData.List(lbData.ListCount - 1, 2) = lbPoles.List(i, 2)
                    lbData.List(lbData.ListCount - 1, 3) = lbPoles.List(i, 3)
                    GoTo Next_I
                End If
            Next j
        End If
Next_I:
    Next i
    
    If lbOwners.ListCount > 0 Then
        For i = 0 To lbOwners.ListCount - 1
            cbOwners.AddItem lbOwners.List(i, 0)
        Next i
    End If
    
    tbTotalEH.Value = lbData.ListCount
End Sub

Private Sub SortEH()
    Dim strTemp, strTotal As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 1) As String
    
    iCount = lbEHReport.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbEHReport.List(b, 0) > lbEHReport.List(b + 1, 0) Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbEHReport.Selected(b) = True
                    lbEHReport.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbEHReport.List(b + 1, 0)
                strAtt(1) = lbEHReport.List(b + 1, 1)
                
                lbEHReport.List(b + 1, 0) = lbEHReport.List(b, 0)
                lbEHReport.List(b + 1, 1) = lbEHReport.List(b, 1)
                
                lbEHReport.List(b, 0) = strAtt(0)
                lbEHReport.List(b, 1) = strAtt(1)
            End If
        Next b
    Next a
End Sub
