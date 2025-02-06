VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReplaceCounts 
   Caption         =   "Find / Replace Counts"
   ClientHeight    =   3591
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   OleObjectBlob   =   "ReplaceCounts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReplaceCounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbMove_Click()
    If tbMName.Value = "" Then Exit Sub
    If tbMFC.Value = "" Then Exit Sub
    If tbMTC.Value = "" Then Exit Sub
    If cbMSize.Value = "" Then Exit Sub
    If tbMFF.Value = "" Then Exit Sub
    If tbMTF.Value = "" Then Exit Sub
    
    '--------------------------------------------
    
    Dim strSize, strName As String
    Dim iFC, iTC As Integer
    Dim iFF, iTF As Integer
    
    strSize = "(" & cbMSize.Value & ")"
    strName = tbMName.Value
    
    iFC = CInt(tbMFC.Value)
    iTC = CInt(tbMTC.Value)
    iFF = CInt(tbMFF.Value)
    iTF = CInt(tbMTF.Value)
    
    '--------------------------------------------
    
    Dim vDwgLL, vDwgUR As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strLine, strResult As String
    Dim vLine, vItem, vCable As Variant
    
    Me.Hide
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    objSS.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    If objSS.count < 1 Then GoTo Exit_Sub
    
    'MsgBox objSS.count & "  CableCounts found."
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        strLine = vAttList(25).TextString
        If strLine = "" Then GoTo Next_objBlock
        If InStr(strLine, strSize) < 1 Then GoTo Next_objBlock
        If InStr(strLine, strName) < 1 Then GoTo Next_objBlock
        
        vLine = Split(strLine, vbCr)
        For i = 0 To UBound(vLine)
            If InStr(vLine(i), strName) > 0 Then
                vItem = Split(vLine(i), " / ")
                strResult = MoveCounts(CStr(vItem(1)))
                vLine(i) = vItem(0) & " / " & strResult
            End If
        Next i
        
        strResult = vLine(0)
        If UBound(vLine) > 0 Then
            For i = 1 To UBound(vLine)
                strResult = strResult & vbCr & vLine(i)
            Next i
        End If
        
        vAttList(25).TextString = strResult
        
        objBlock.Update
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbReplace_Click()
    Select Case cbOptions.Value
        Case "Single"
            Call ReplaceSingle
        Case "All sPole"
            Call ReplaceAllsPole
        Case "All Callouts"
            Call ReplaceAllCallouts
    End Select
End Sub

Private Sub Label7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ReplaceAllCallouts
End Sub

Private Sub tbMFF_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If tbMTC.Value = "" Then Exit Sub
    If tbMFC.Value = "" Then Exit Sub
    If tbMFF.Value = "" Then Exit Sub
    
    Dim iFC, iTC As Integer
    Dim iFF, iTF As Integer
    Dim iDiff, iTest As Integer
    Dim dTest, dResult As Double
    
    iFC = CInt(tbMFC.Value)
    iTC = CInt(tbMTC.Value)
    iDiff = iTC - iFC
    iFF = CInt(tbMFF.Value)
    iTF = iFF + iDiff
    
    iTest = Abs(iFF - iFC)
    dTest = (iTest / 12)
    dResult = dTest - Int(dTest)
    
    If Not dResult = 0# Then
        MsgBox "Off-Color"
        Exit Sub
    End If
    
    tbMTF.Value = iTF
End Sub

Private Sub tbMName_Change()
    If tbMName.Value = "" Then Exit Sub
    
    tbMName.Value = UCase(tbMName.Value)
End Sub

Private Sub tbMTC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If tbMFC.Value = "" Then Exit Sub
    
    If tbMFC.Value = "" Then tbMTC.Value = tbMFC.Value
End Sub

Private Sub UserForm_Initialize()
    cbOptions.AddItem "Single"
    cbOptions.AddItem "All sPole"
    cbOptions.AddItem "All Callouts"
    cbOptions.Value = "Single"
    
    cbMSize.AddItem ""
    cbMSize.AddItem "12"
    cbMSize.AddItem "24"
    cbMSize.AddItem "48"
    cbMSize.AddItem "72"
    cbMSize.AddItem "96"
    cbMSize.AddItem "144"
    cbMSize.AddItem "216"
    cbMSize.AddItem "288"
    cbMSize.AddItem "360"
    cbMSize.AddItem "432"
    cbMSize.AddItem "576"
    
    tbFName.SetFocus
End Sub

Private Sub ReplaceSingle()
    If tbFName.Value = "" Then Exit Sub
    If tbRName.Value = "" Then Exit Sub
    
    If tbFFrom.Value = "" Then Exit Sub
    If tbRFrom.Value = "" Then Exit Sub
    
    If tbFTo.Value = "" Then tbFTo.Value = tbFFrom.Value
    If tbRTo.Value = "" Then tbRTo.Value = tbRFrom.Value
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim vAttList As Variant
    Dim vLine, vItem, vCount As Variant
    Dim vTemp As Variant
    Dim strFind, strReplace As String
    Dim strLine, strInsert As String
    Dim iFFrom, iFTo, iRFrom, iRTo As Integer
    Dim iFFFiber, iRFFiber As Integer
    Dim iFind, iReplace As Integer
    Dim iCFrom, iCTo As Integer
    Dim iFTemp, iTTemp As Integer
    
    On Error Resume Next
    
    strFind = UCase(tbFName.Value)
    strReplace = UCase(tbRName.Value)
    
    iFFrom = CInt(tbFFrom.Value)
    If Not Err = 0 Then
        MsgBox "Find From is not an integer"
        Exit Sub
    End If
    
    iFTo = CInt(tbFTo.Value)
    If Not Err = 0 Then
        MsgBox "Find To is not an integer"
        Exit Sub
    End If
    
    iFind = iFTo - iFFrom
    
    iRFrom = CInt(tbRFrom.Value)
    If Not Err = 0 Then
        MsgBox "Replace From is not an integer"
        Exit Sub
    End If
    
    iRTo = CInt(tbRTo.Value)
    If Not Err = 0 Then
        MsgBox "Replace To is not an integer"
        Exit Sub
    End If
    
    iReplace = iRTo - iRFrom
    
    If Not iFind = iReplace Then
        strLine = "Number of Find counts (" & iFind & ") does not match the Replace counts (" & iReplace & ")"
        MsgBox strLine
        Exit Sub
    End If
    
    iFFFiber = iFFrom
    While iFFFiber > 12
        iFFFiber = iFFFiber - 12
    Wend
    
    iRFFiber = iRFrom
    While iRFFiber > 12
        iRFFiber = iRFFiber - 12
    Wend
    
    If Not iFFFiber = iRFFiber Then
        strLine = "Start Fiber for Find counts (" & iFFFiber & ") does not match the Replace counts (" & iRFFiber & ")"
        MsgBox strLine
        Exit Sub
    End If
    
    
    Me.Hide
    
Get_Entity:
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Callout or Pole: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Get_Entity
    Set objBlock = objEntity
    
    Select Case objBlock.Name
        Case "Callout"
            vAttList = objBlock.GetAttributes
            
            vLine = Split(UCase(vAttList(2).TextString), "\P")
            For i = 0 To UBound(vLine)
                vItem = Split(vLine(i), ": ")
                If vItem(0) = strFind Then
                    vCount = Split(vItem(1), "-")
                    iCFrom = CInt(vCount(0))
                    If UBound(vCount) = 0 Then
                        iCTo = CInt(vCount(0))
                    Else
                        iCTo = CInt(vCount(1))
                    End If
                    
                    Select Case iCFrom - iFFrom
                        Case Is < 0
                            If iCTo < iFTo Then GoTo Next_I1
                            
                            iTTemp = iFFrom - 1
                            iFTemp = iFTo + 1
                            strInsert = strFind & ": " & iCFrom
                            If Not iTTemp = iCFrom Then strInsert = strInsert & "-" & iTTemp
                            
                            strInsert = strInsert & "\P" & strReplace & ": " & iRFrom
                            If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                            
                            If Not iFTemp > iCTo Then
                                strInsert = strInsert & "\P" & strFind & ": " & iFTemp
                                If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                            End If
                            
                            vLine(i) = strInsert
                            
                            strLine = vLine(0)
                            If UBound(vLine) > 0 Then
                                For j = 1 To UBound(vLine)
                                    strLine = strLine & "\P" & vLine(j)
                                Next j
                            End If
                            
                            vAttList(2).TextString = strLine
                            objBlock.Update
                            
                            GoTo Get_Entity
                        Case Is = 0
                            If iCTo < iFTo Then GoTo Next_I1
                            
                            iFTemp = iFTo + 1
                            strInsert = strReplace & ": " & iRFrom
                            If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                            
                            If Not iFTemp > iCTo Then
                                strInsert = strInsert & "\P" & strFind & ": " & iFTemp
                                If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                            End If
                            
                            vLine(i) = strInsert
                            
                            strLine = vLine(0)
                            If UBound(vLine) > 0 Then
                                For j = 1 To UBound(vLine)
                                    strLine = strLine & "\P" & vLine(j)
                                Next j
                            End If
                            
                            vAttList(2).TextString = strLine
                            objBlock.Update
                            
                            GoTo Get_Entity
                        Case Is > 0
                            GoTo Next_I1
                    End Select
                End If
Next_I1:
            Next i
        Case "CableCounts"
            vAttList = objBlock.GetAttributes
            
            vLine = Split(UCase(vAttList(0).TextString), "\P")
            For i = 0 To UBound(vLine)
                vItem = Split(vLine(i), ": ")
                If vItem(0) = strFind Then
                    vCount = Split(vItem(1), "-")
                    iCFrom = CInt(vCount(0))
                    If UBound(vCount) = 0 Then
                        iCTo = CInt(vCount(0))
                    Else
                        iCTo = CInt(vCount(1))
                    End If
                    
                    Select Case iCFrom - iFFrom
                        Case Is < 0
                            If iCTo < iFTo Then GoTo Next_I
                            
                            iTTemp = iFFrom - 1
                            iFTemp = iFTo + 1
                            strInsert = strFind & ": " & iCFrom
                            If Not iTTemp = iCFrom Then strInsert = strInsert & "-" & iTTemp
                            
                            strInsert = strInsert & "\P" & strReplace & ": " & iRFrom
                            If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                            
                            If Not iFTemp > iCTo Then
                                strInsert = strInsert & "\P" & strFind & ": " & iFTemp
                                If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                            End If
                            
                            vLine(i) = strInsert
                            
                            strLine = vLine(0)
                            If UBound(vLine) > 0 Then
                                For j = 1 To UBound(vLine)
                                    strLine = strLine & "\P" & vLine(j)
                                Next j
                            End If
                            
                            vAttList(0).TextString = strLine
                            objBlock.Update
                            
                            GoTo Get_Entity
                        Case Is = 0
                            If iCTo < iFTo Then GoTo Next_I
                            
                            iFTemp = iFTo + 1
                            strInsert = strReplace & ": " & iRFrom
                            If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                            
                            If Not iFTemp > iCTo Then
                                strInsert = strInsert & "\P" & strFind & ": " & iFTemp
                                If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                            End If
                            
                            vLine(i) = strInsert
                            
                            strLine = vLine(0)
                            If UBound(vLine) > 0 Then
                                For j = 1 To UBound(vLine)
                                    strLine = strLine & "\P" & vLine(j)
                                Next j
                            End If
                            
                            vAttList(0).TextString = strLine
                            objBlock.Update
                            
                            GoTo Get_Entity
                        Case Is > 0
                            GoTo Next_I
                    End Select
                End If
Next_I:
            Next i
        Case "sPole"
            vAttList = objBlock.GetAttributes
            
            vLine = Split(UCase(vAttList(25).TextString), " / ")
            vTemp = Split(vLine(1), " + ")
            
            
            For i = 0 To UBound(vTemp)
                vItem = Split(vTemp(i), ": ")
                If vItem(0) = strFind Then
                    vCount = Split(vItem(1), "-")
                    iCFrom = CInt(vCount(0))
                    If UBound(vCount) = 0 Then
                        iCTo = CInt(vCount(0))
                    Else
                        iCTo = CInt(vCount(1))
                    End If
                    
                    Select Case iCFrom - iFFrom
                        Case Is < 0
                            If iCTo < iFTo Then GoTo Next_I1
                            
                            iTTemp = iFFrom - 1
                            iFTemp = iFTo + 1
                            strInsert = strFind & ": " & iCFrom
                            If Not iTTemp = iCFrom Then strInsert = strInsert & "-" & iTTemp
                            If UBound(vItem) > 1 Then strInsert = strInsert & ": " & vItem(2)
                            
                            strInsert = strInsert & " + " & strReplace & ": " & iRFrom
                            If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                            If UBound(vItem) > 1 Then strInsert = strInsert & ": " & vItem(2)
                            
                            If Not iFTemp > iCTo Then
                                strInsert = strInsert & " + " & strFind & ": " & iFTemp
                                If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                                If UBound(vItem) > 1 Then strInsert = strInsert & ": " & vItem(2)
                            End If
                            
                            vTemp(i) = strInsert
                            
                            strLine = vTemp(0)
                            If UBound(vTemp) > 0 Then
                                For j = 1 To UBound(vTemp)
                                    strLine = strLine & " + " & vTemp(j)
                                Next j
                            End If
                            
                            vAttList(25).TextString = vLine(0) & " / " & strLine
                            objBlock.Update
                            
                            GoTo Get_Entity
                        Case Is = 0
                            If iCTo < iFTo Then GoTo Next_I2
                            
                            iFTemp = iFTo + 1
                            strInsert = strReplace & ": " & iRFrom
                            If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                            If UBound(vItem) > 1 Then strInsert = strInsert & ": " & vItem(2)
                            
                            If Not iFTemp > iCTo Then
                                strInsert = strInsert & " + " & strFind & ": " & iFTemp
                                If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                                If UBound(vItem) > 1 Then strInsert = strInsert & ": " & vItem(2)
                            End If
                            
                            vTemp(i) = strInsert
                            
                            strLine = vTemp(0)
                            If UBound(vTemp) > 0 Then
                                For j = 1 To UBound(vTemp)
                                    strLine = strLine & " + " & vTemp(j)
                                Next j
                            End If
                            
                            vAttList(25).TextString = vLine(0) & " / " & strLine
                            objBlock.Update
                            
                            GoTo Get_Entity
                        Case Is > 0
                            GoTo Next_I2
                    End Select
                End If
Next_I2:
            Next i
        Case Else
    End Select
    
    GoTo Get_Entity
Exit_Sub:
    Me.show
End Sub

Private Sub ReplaceAllCallouts()
    If tbFName.Value = "" Then Exit Sub
    If tbRName.Value = "" Then Exit Sub
    
    If tbFFrom.Value = "" Then Exit Sub
    If tbRFrom.Value = "" Then Exit Sub
    
    If tbFTo.Value = "" Then tbFTo.Value = tbFFrom.Value
    If tbRTo.Value = "" Then tbRTo.Value = tbRFrom.Value
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim vAttList As Variant
    Dim vLine, vItem, vCount As Variant
    Dim vPole As Variant
    Dim strFind, strReplace As String
    Dim strLine, strInsert As String
    Dim iFFrom, iFTo, iRFrom, iRTo As Integer
    Dim iFFFiber, iRFFiber As Integer
    Dim iFind, iReplace As Integer
    Dim iCFrom, iCTo As Integer
    Dim iFTemp, iTTemp As Integer
    
    On Error Resume Next
    
    strFind = UCase(tbFName.Value)
    strReplace = UCase(tbRName.Value)
    
    iFFrom = CInt(tbFFrom.Value)
    If Not Err = 0 Then
        MsgBox "Find From is not an integer"
        Exit Sub
    End If
    
    iFTo = CInt(tbFTo.Value)
    If Not Err = 0 Then
        MsgBox "Find To is not an integer"
        Exit Sub
    End If
    
    iFind = iFTo - iFFrom
    
    iRFrom = CInt(tbRFrom.Value)
    If Not Err = 0 Then
        MsgBox "Replace From is not an integer"
        Exit Sub
    End If
    
    iRTo = CInt(tbRTo.Value)
    If Not Err = 0 Then
        MsgBox "Replace To is not an integer"
        Exit Sub
    End If
    
    iReplace = iRTo - iRFrom
    
    If Not iFind = iReplace Then
        strLine = "Number of Find counts (" & iFind & ") does not match the Replace counts (" & iReplace & ")"
        MsgBox strLine
        Exit Sub
    End If
    
    iFFFiber = iFFrom
    While iFFFiber > 12
        iFFFiber = iFFFiber - 12
    Wend
    
    iRFFiber = iRFrom
    While iRFFiber > 12
        iRFFiber = iRFFiber - 12
    Wend
    
    If Not iFFFiber = iRFFiber Then
        strLine = "Start Fiber for Find counts (" & iFFFiber & ") does not match the Replace counts (" & iRFFiber & ")"
        MsgBox strLine
        Exit Sub
    End If
    
    Dim vDwgLL, vDwgUR As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSDwg As AcadSelectionSet
    
    Me.Hide
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    
    grpCode(0) = 2
    grpValue(0) = "CableCounts"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSDwg = ThisDrawing.SelectionSets.Add("objSSDwg")
    objSSDwg.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    If objSSDwg.count < 1 Then GoTo Exit_Sub
    
    MsgBox objSSDwg.count & "  CableCounts found."
    
    For Each objBlock In objSSDwg
        vAttList = objBlock.GetAttributes
            
        vLine = Split(UCase(vAttList(0).TextString), "\P")
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), ": ")
            
            If vItem(0) = strFind Then
                vCount = Split(vItem(1), "-")
                iCFrom = CInt(vCount(0))
                If UBound(vCount) = 0 Then
                    iCTo = CInt(vCount(0))
                Else
                    iCTo = CInt(vCount(1))
                End If
                    
                Select Case iCFrom - iFFrom
                    Case Is < 0
                        If iCTo < iFTo Then GoTo Next_I
                            
                        iTTemp = iFFrom - 1
                        iFTemp = iFTo + 1
                        strInsert = strFind & ": " & iCFrom
                        If Not iTTemp = iCFrom Then strInsert = strInsert & "-" & iTTemp
                            
                        strInsert = strInsert & "\P" & strReplace & ": " & iRFrom
                        If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                            
                        If Not iFTemp > iCTo Then
                            strInsert = strInsert & "\P" & strFind & ": " & iFTemp
                            If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                        End If
                            
                        vLine(i) = strInsert
                            
                        strLine = vLine(0)
                        If UBound(vLine) > 0 Then
                            For j = 1 To UBound(vLine)
                                strLine = strLine & "\P" & vLine(j)
                            Next j
                        End If
                            
                        vAttList(0).TextString = strLine
                        objBlock.Update
                            
                        GoTo Next_objBlock
                    Case Is = 0
                        If iCTo < iFTo Then GoTo Next_I
                            
                        iFTemp = iFTo + 1
                        strInsert = strReplace & ": " & iRFrom
                        If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                            
                        If Not iFTemp > iCTo Then
                            strInsert = strInsert & "\P" & strFind & ": " & iFTemp
                            If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                        End If
                            
                        vLine(i) = strInsert
                            
                        strLine = vLine(0)
                        If UBound(vLine) > 0 Then
                            For j = 1 To UBound(vLine)
                                strLine = strLine & "\P" & vLine(j)
                            Next j
                        End If
                            
                        vAttList(0).TextString = strLine
                        objBlock.Update
                            
                        GoTo Next_objBlock
                    Case Is > 0
                        GoTo Next_I
                End Select
            End If
Next_I:
        Next i
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    objSSDwg.Clear
    objSSDwg.Delete
    
    Me.show
End Sub

Private Sub ReplaceAllsPole()
    If tbFName.Value = "" Then Exit Sub
    If tbRName.Value = "" Then Exit Sub
    
    If tbFFrom.Value = "" Then Exit Sub
    If tbRFrom.Value = "" Then Exit Sub
    
    If tbFTo.Value = "" Then tbFTo.Value = tbFFrom.Value
    If tbRTo.Value = "" Then tbRTo.Value = tbRFrom.Value
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim vAttList As Variant
    Dim vLine, vItem, vCount As Variant
    Dim vPole As Variant
    Dim strFind, strReplace As String
    Dim strLine, strInsert As String
    Dim iFFrom, iFTo, iRFrom, iRTo As Integer
    Dim iFFFiber, iRFFiber As Integer
    Dim iFind, iReplace As Integer
    Dim iCFrom, iCTo As Integer
    Dim iFTemp, iTTemp As Integer
    
    On Error Resume Next
    
    strFind = UCase(tbFName.Value)
    strReplace = UCase(tbRName.Value)
    
    iFFrom = CInt(tbFFrom.Value)
    If Not Err = 0 Then
        MsgBox "Find From is not an integer"
        Exit Sub
    End If
    
    iFTo = CInt(tbFTo.Value)
    If Not Err = 0 Then
        MsgBox "Find To is not an integer"
        Exit Sub
    End If
    
    iFind = iFTo - iFFrom
    
    iRFrom = CInt(tbRFrom.Value)
    If Not Err = 0 Then
        MsgBox "Replace From is not an integer"
        Exit Sub
    End If
    
    iRTo = CInt(tbRTo.Value)
    If Not Err = 0 Then
        MsgBox "Replace To is not an integer"
        Exit Sub
    End If
    
    iReplace = iRTo - iRFrom
    
    If Not iFind = iReplace Then
        strLine = "Number of Find counts (" & iFind & ") does not match the Replace counts (" & iReplace & ")"
        MsgBox strLine
        Exit Sub
    End If
    
    iFFFiber = iFFrom
    While iFFFiber > 12
        iFFFiber = iFFFiber - 12
    Wend
    
    iRFFiber = iRFrom
    While iRFFiber > 12
        iRFFiber = iRFFiber - 12
    Wend
    
    If Not iFFFiber = iRFFiber Then
        strLine = "Start Fiber for Find counts (" & iFFFiber & ") does not match the Replace counts (" & iRFFiber & ")"
        MsgBox strLine
        Exit Sub
    End If
    
    Dim vDwgLL, vDwgUR As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSDwg As AcadSelectionSet
    
    Me.Hide
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSDwg = ThisDrawing.SelectionSets.Add("objSSDwg")
    objSSDwg.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    If objSSDwg.count < 1 Then GoTo Exit_Sub
    
    MsgBox objSSDwg.count & "  CableCounts found."
    
    For Each objBlock In objSSDwg
        vAttList = objBlock.GetAttributes
        
        vPole = Split(UCase(vAttList(25).TextString), " / ")
        vLine = Split(vPole(1), " + ")
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), ": ")
            
            If vItem(0) = strFind Then
                strBack = vItem(2)
                vCount = Split(vItem(1), "-")
                iCFrom = CInt(vCount(0))
                If UBound(vCount) = 0 Then
                    iCTo = CInt(vCount(0))
                Else
                    iCTo = CInt(vCount(1))
                End If
                    
                Select Case iCFrom - iFFrom
                    Case Is < 0
                        If iCTo < iFTo Then GoTo Next_I
                            
                        iTTemp = iFFrom - 1
                        iFTemp = iFTo + 1
                        strInsert = strFind & ": " & iCFrom
                        If Not iTTemp = iCFrom Then strInsert = strInsert & "-" & iTTemp
                        strInsert = strInsert & ": " & strBack
                            
                        strInsert = strInsert & " + " & strReplace & ": " & iRFrom
                        If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                        strInsert = strInsert & ": " & strBack
                            
                        If Not iFTemp > iCTo Then
                            strInsert = strInsert & " + " & strFind & ": " & iFTemp
                            If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                            strInsert = strInsert & ": " & strBack
                        End If
                            
                        vLine(i) = strInsert
                            
                        strLine = vLine(0)
                        If UBound(vLine) > 0 Then
                            For j = 1 To UBound(vLine)
                                strLine = strLine & " + " & vLine(j)
                            Next j
                        End If
                            
                        vAttList(25).TextString = vPole(0) & " / " & strLine
                        objBlock.Update
                            
                        GoTo Next_objBlock
                    Case Is = 0
                        If iCTo < iFTo Then GoTo Next_I
                            
                        iFTemp = iFTo + 1
                        strInsert = strReplace & ": " & iRFrom
                        If Not iRFrom = iRTo Then strInsert = strInsert & "-" & iRTo
                        strInsert = strInsert & ": " & strBack
                            
                        If Not iFTemp > iCTo Then
                            strInsert = strInsert & " + " & strFind & ": " & iFTemp
                            If iCTo > iFTemp Then strInsert = strInsert & "-" & iCTo
                            strInsert = strInsert & ": " & strBack
                        End If
                            
                        vLine(i) = strInsert
                            
                        strLine = vLine(0)
                        If UBound(vLine) > 0 Then
                            For j = 1 To UBound(vLine)
                                strLine = strLine & " + " & vLine(j)
                            Next j
                        End If
                            
                        vAttList(25).TextString = vPole(0) & " / " & strLine
                        objBlock.Update
                            
                        GoTo Next_objBlock
                    Case Is > 0
                        GoTo Next_I
                End Select
            End If
Next_I:
        Next i
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    objSSDwg.Clear
    objSSDwg.Delete
    
    Me.show
End Sub

Private Function MoveCounts(strLine As String)
    Dim vLine, vItem, vCounts As Variant
    Dim strResult, strTemp As String
    Dim strName As String
    Dim strPrevious As String
    Dim iFC, iTC As Integer
    Dim iFF, iTF As Integer
    Dim iCFC, iCTC As Integer
    Dim iCFF, iCTF As Integer
    Dim iFFiber, iTFiber, iTemp As Integer
    
    Dim iSize As Integer
    Dim iFiber, iStart, iEnd As Integer
    Dim strCable() As String
    Dim iCount() As Integer
    Dim strSource() As String
    
    iSize = CInt(cbMSize.Value)
    ReDim strCable(iSize) As String
    ReDim iCount(iSize) As Integer
    ReDim strSource(iSize) As String
    iFiber = 1
    
    strName = tbMName.Value
    strResult = ""
    
    iFC = CInt(tbMFC.Value)
    iTC = CInt(tbMTC.Value)
    iFF = CInt(tbMFF.Value)
    iTF = CInt(tbMTF.Value)
    
    vLine = Split(strLine, " + ")
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        If vItem(0) = strName Then
            vCounts = Split(vlitem(1), "-")
            iCFC = CInt(vCounts(0))
            
            If UBound(vCounts) > 0 Then
                iCTC = CInt(vCounts(1))
            Else
                iCTC = iCFC
            End If
            
            If iCTC < iFC Then GoTo Next_I
            If iCFC > iTC Then GoTo Next_I
            
            If iCFC < iFC Then
                iCFC = iFC
                iCFF = iFF
            Else
                iCFF = iFF + (iCFC - iFC)
            End If
            
            If iCTC > iTC Then
                iCTC = iTC
                iCTF = iTF
            Else
                iCTF = iTF + (iCTC - iTC)
            End If
            
            If UBound(vItem) > 1 Then
                strSource(0) = vItem(2)
            Else
                strSource(0) = ""
            End If
Next_I:
    Next i
    
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        vCounts = Split(vlitem(1), "-")
        
        iStart = CInt(vCounts(0))
        If UBound(vCounts) = 0 Then
            iEnd = iStart
        Else
            iEnd = CInt(vCounts(1))
        End If
        
        For j = iStart To iEnd
            If vItem(0) = strName Then
                If j > iFF - 1 And j < iTF + 1 Then
                    strCable(iFiber) = "XD"
                    
                    If UBound(vItem) < 2 Then
                        strSource(iFiber) = ""
                    Else
                        strSource(iFiber) = vItem(2)
                    End If
                    
                    iCount(iFiber) = iFiber
                    
                    GoTo Omit_Fiber
                End If
            End If
            
            strCable(iFiber) = vItem(0)
            If UBound(vItem) < 2 Then
                strSource(iFiber) = ""
            Else
                strSource(iFiber) = vItem(2)
            End If
            
            iCount(iFiber) = j
            
Omit_Fiber:
            
            iFiber = iFiber + 1
        Next j
    Next i
    
    For i = iCFF To iCTF
        strCable(i) = strName
        strSource(i) = strSource(0)
        iCount(i) = iCFC + i - iCFF
    Next i
    
    strResult = ""
    
    strPrevious = strCable(1)
    iStart = iCount(1)
    iEnd = iStart
    
    For i = 2 To UBound(strCable)
        If strCable(i) = strPrevious Then
            iEnd = iCount(i)
        Else
            strLine = strPrevious & ": " & iStart
            
            If Not iEnd = iStart Then strLine = strLine & "-" & iEnd
            If Not strSource(i - 1) = "" Then strLine = strLine & ": " & strSource(i)
            
            If strResult = "" Then
                strResult = strLine
            Else
                strResult = strResult & " + " & strLine
            End If
            
            strPrevious = strCable(i)
            iStart = iCount(i)
            iEnd = iStart
        End If
    Next i
    
    MoveCounts = strResult
End Function
