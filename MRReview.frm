VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MRReview 
   Caption         =   "Make Ready Review"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9480.001
   OleObjectBlob   =   "MRReview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MRReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSS As AcadSelectionSet

Private Sub cbExport_Click()
    If lbAll.ListCount < 1 Then Exit Sub
    
    Dim strPath, strName, strFile As String
    Dim vTemp As Variant
    
    strPath = ThisDrawing.Path
    strName = ThisDrawing.Name
    vTemp = Split(strName, " ")
    
    strFile = strPath & "\" & vTemp(0) & " MR Report.csv"
    
    Open strFile For Output As #1
    
    Print #1, "Structure Number,Owner,New Attachment,MR,Status"
    
    For i = 0 To lbAll.ListCount - 1
        Print #1, lbAll.List(i, 0) & "," & lbAll.List(i, 1) & "," & lbAll.List(i, 2) & "," & lbAll.List(i, 3) & "," & lbAll.List(i, 4)
    Next i
    
    Close #1
End Sub

Private Sub cbFilter_Click()
    If lbAll.ListCount < 1 Then Exit Sub
    If cbOwner.Value = "" And cbStatus.Value = "" Then Exit Sub
    
    Dim strOwner, strStatus As String
    Dim iResult As Integer
    
    strOwner = cbOwner.Value
    strStatus = cbStatus.Value
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        iResult = 0
        
        If strOwner = "" Then
            iResult = 1
        Else
            If strOwner = lbAll.List(i, 1) Then iResult = 1
        End If
        
        If strStatus = "" Then
            iResult = iResult + 1
        Else
            If InStr(lbAll.List(i, 4), strStatus) > 0 Then iResult = iResult + 1
        End If
        
        If iResult < 2 Then lbAll.RemoveItem i
    Next i
End Sub

Private Sub cbFRColumn_Change()
    Dim vLine As Variant
    
    cbFRValue.Clear
    If lbAll.ListCount < 2 Then Exit Sub
    
    Select Case cbFRColumn.Value
        Case "0: Pole Number"
            cbFRValue.AddItem lbAll.List(0, 0)
            
            For i = 1 To lbAll.ListCount - 1
                For j = 0 To cbFRValue.ListCount - 1
                    If cbFRValue.List(j) = lbAll.List(i, 0) Then GoTo Exit_PN
                Next j
        
                cbFRValue.AddItem lbAll.List(i, 0)
Exit_PN:
            Next i
        Case "1: Owner"
            cbFRValue.AddItem lbAll.List(0, 1)
            
            For i = 1 To lbAll.ListCount - 1
                For j = 0 To cbFRValue.ListCount - 1
                    If cbFRValue.List(j) = lbAll.List(i, 1) Then GoTo Exit_J
                Next j
        
                cbFRValue.AddItem lbAll.List(i, 1)
Exit_J:
            Next i
        Case "2: New At"
            cbFRValue.AddItem "none"
            cbFRValue.AddItem "**T"
            cbFRValue.AddItem "**F"
            cbFRValue.AddItem "**O"
            cbFRValue.AddItem "<="
            cbFRValue.AddItem ">="
            cbFRValue.AddItem "="
            
            cbFRValue.Value = "none"
        Case "3: MR"
            cbFRValue.AddItem "none"
            
            For i = 0 To lbAll.ListCount - 1
                If Not lbAll.List(i, 4) = "" Then
                    vLine = Split(lbAll.List(i, 3), ", ")
            
                    For j = 0 To UBound(vLine)
                        If Not cbFRValue.ListCount < 0 Then
                            For k = 0 To cbFRValue.ListCount - 1
                                If cbFRValue.List(k) = vLine(j) Then GoTo Exit_L
                            Next k
                        End If
                        
                        cbFRValue.AddItem vLine(j)
Exit_L:
                    Next j
                End If
            Next i
            
            If cbFRValue.ListCount > 0 Then
                For i = cbFRValue.ListCount - 1 To 0 Step -1
                    If cbFRValue.List(i) = "" Then cbFRValue.RemoveItem i
                    'If cbFRValue.List(i) = " " Then cbFRValue.RemoveItem i
                Next i
            End If
            
            cbFRValue.Value = "none"
        Case "4: Status"
            If Not lbAll.List(0, 4) = "" Then
                vLine = Split(lbAll.List(0, 4), ";")
                For j = 0 To UBound(vLine)
                    cbFRValue.AddItem vLine(j)
                Next j
            End If
            
            For i = 1 To lbAll.ListCount - 1
                If Not lbAll.List(i, 4) = "" Then
                    vLine = Split(lbAll.List(i, 4), ";")
            
                    For j = 0 To UBound(vLine)
                        If Not cbFRValue.ListCount < 0 Then
                            For k = 0 To cbFRValue.ListCount - 1
                                If cbFRValue.List(k) = vLine(j) Then GoTo Exit_K
                            Next k
                        End If
                        
                        cbFRValue.AddItem vLine(j)
Exit_K:
                    Next j
                End If
            Next i
            
            If cbFRValue.ListCount > 0 Then
                For i = cbFRValue.ListCount - 1 To 0 Step -1
                    If cbFRValue.List(i) = "" Then cbFRValue.RemoveItem i
                    'If cbFRValue.List(i) = " " Then cbFRValue.RemoveItem i
                Next i
            End If
        Case Else
    End Select
End Sub

Private Sub cbFROption_Change()
    'If cbFROption.Value = "Filter" Then
        'cbFRRun.Caption = "Filter"
    'Else
        'cbFRRun.Caption = "Remove"
    'End If
End Sub

Private Sub cbFRRun_Click()
    If lbAll.ListCount < 2 Then Exit Sub
    
    Select Case cbFRColumn.Value
        Case "0: Pole Number"
            Call FilterRemovePN
        Case "1: Owner"
            Call FilterRemoveOwner
        Case "2: New At"
            Call FilterRemoveNewAt
        Case "3: MR"
            Call FilterRemoveMR
        Case "4: Status"
            Call FilterRemoveStatus
    End Select
End Sub

Private Sub cbGetPoles_Click()
    Me.Hide
        Call GetAllPoles
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRemove24only_Click()
    Dim iIndex As Integer
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iCOMM, iNew, iTemp As Integer
    Dim vLine, vAttach As Variant
    Dim strTemp As String
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        iIndex = CInt(lbAll.List(i, 5))
        
        Set objBlock = objSS.Item(iIndex)
        vAttList = objBlock.GetAttributes
        
        If vAttList(15).TextString = "" Then GoTo Next_I
        
        vLine = Split(vAttList(15).TextString, " ")
        
        If UBound(vLine) > 0 Then GoTo Next_I
        
        strTemp = Replace(UCase(vLine(0)), "T", "")
            
        vAttach = Split(strTemp, "-")
        iNew = CInt(vAttach(0)) * 12
        If UBound(vAttach) > 0 Then iNew = iNew + CInt(vAttach(1))
        
        iMax = FindMaxHeight(iIndex)
        
        'MsgBox "Pwr Max:  " & iMax
        
        If iMax < 288 Then GoTo Next_I
        
        iCOMM = FindTopCOMM(iIndex)
        
        'MsgBox "Top COMM:  " & iCOMM
        
        If iCOMM > 0 Then GoTo Next_I
        
        lbAll.RemoveItem i
Next_I:
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub cbRemoveCOMM1_Click()
    Dim iIndex As Integer
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iCOMM, iNew, iTemp As Integer
    Dim vLine, vAttach As Variant
    Dim strTemp As String
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        iIndex = CInt(lbAll.List(i, 5))
        
        Set objBlock = objSS.Item(iIndex)
        vAttList = objBlock.GetAttributes
        
        If vAttList(15).TextString = "" Then GoTo Next_I
        
        iMax = FindMaxHeight(iIndex)
        
        iCOMM = FindTopCOMM(iIndex) + 12
        If iCOMM = 1012 Then GoTo Next_I
        
        vLine = Split(vAttList(15).TextString, " ")
        For j = 0 To UBound(vLine)
            strTemp = Replace(UCase(vLine(j)), "T", "")
            vAttach = Split(strTemp, "-")
            
            iNew = CInt(vAttach(0)) * 12
            If UBound(vAttach) > 0 Then iNew = iNew + CInt(vAttach(1))
            
            If iNew > iMax Then GoTo Next_I
            If Not iNew = iCOMM Then GoTo Next_I
        Next j
        
        lbAll.RemoveItem i
Next_I:
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub cbRemoveCOMM2_Click()
    Dim iIndex As Integer
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iCOMM, iNew, iTemp As Integer
    Dim vLine, vAttach As Variant
    Dim strTemp As String
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        iIndex = CInt(lbAll.List(i, 5))
        
        Set objBlock = objSS.Item(iIndex)
        vAttList = objBlock.GetAttributes
        
        If vAttList(15).TextString = "" Then GoTo Next_I
        
        iMax = FindMaxHeight(iIndex)
        
        iCOMM = FindTopCOMM(iIndex) + 24
        If iCOMM = 1024 Then GoTo Next_I
        
        vLine = Split(vAttList(15).TextString, " ")
        For j = 0 To UBound(vLine)
            strTemp = Replace(UCase(vLine(j)), "T", "")
            vAttach = Split(strTemp, "-")
            
            iNew = CInt(vAttach(0)) * 12
            If UBound(vAttach) > 0 Then iNew = iNew + CInt(vAttach(1))
            
            If iNew > iMax Then GoTo Next_I
            If Not iNew = iCOMM Then GoTo Next_I
        Next j
        
        lbAll.RemoveItem i
Next_I:
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub cbRemoveNone_Click()
    For i = lbAll.ListCount - 1 To 0 Step -1
        If lbAll.List(i, 2) = "none" Then lbAll.RemoveItem i
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub cbRemoveOCALC_Click()
    If lbAll.ListCount < 1 Then Exit Sub
    If tbRemoveOcalc.Value = "" Then Exit Sub
    
    Dim objTemp As AcadBlockReference
    Dim vAtt As Variant
    Dim strOwner, strStatus As String
    Dim iIndex As Integer
    
    strOwner = UCase(tbRemoveOcalc.Value)
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        If InStr(lbAll.List(i, 4), "OCALC") > 0 Then
            iIndex = CInt(lbAll.List(i, 5))
            
            lbAll.List(i, 4) = Replace(lbAll.List(i, 4), "OCALC;", "")
            lbAll.List(i, 4) = Replace(lbAll.List(i, 4), "OCALC", "")
            
            Set objTemp = objSS.Item(iIndex)
            vAtt = objTemp.GetAttributes
            vAtt(24).TextString = Replace(vAtt(24).TextString, "OCALC;", "")
            vAtt(24).TextString = Replace(vAtt(24).TextString, "OCALC", "")
            objTemp.Update
        End If
    Next i
End Sub

Private Sub cbRemoveQC_Click()
    For i = lbAll.ListCount - 1 To 0 Step -1
        If InStr(lbAll.List(i, 4), "MR-QC") > 0 Then lbAll.RemoveItem i
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub cbRemoveTag_Click()
    For i = lbAll.ListCount - 1 To 0 Step -1
        If InStr(UCase(lbAll.List(i, 2)), "T") > 0 Then lbAll.RemoveItem i
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub cbReset_Click()
    lbAll.Clear
    
    For i = 0 To lbList.ListCount - 1
        If cbInclude.Value = False Then
            If Not lbList.List(i, 0) = "POLE" Then
                lbAll.AddItem lbList.List(i, 0)
                lbAll.List(lbAll.ListCount - 1, 1) = lbList.List(i, 1)
                lbAll.List(lbAll.ListCount - 1, 2) = lbList.List(i, 2)
                lbAll.List(lbAll.ListCount - 1, 3) = lbList.List(i, 3)
                lbAll.List(lbAll.ListCount - 1, 4) = lbList.List(i, 4)
                lbAll.List(lbAll.ListCount - 1, 5) = lbList.List(i, 5)
            End If
        Else
            lbAll.AddItem lbList.List(i, 0)
            lbAll.List(lbAll.ListCount - 1, 1) = lbList.List(i, 1)
            lbAll.List(lbAll.ListCount - 1, 2) = lbList.List(i, 2)
            lbAll.List(lbAll.ListCount - 1, 3) = lbList.List(i, 3)
            lbAll.List(lbAll.ListCount - 1, 4) = lbList.List(i, 4)
            lbAll.List(lbAll.ListCount - 1, 5) = lbList.List(i, 5)
        End If
    Next i
    
    tbListCount.Value = lbAll.ListCount
End Sub

Private Sub cbSort_Click()
    Dim strTemp, strTotal As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 5) As String
    
    iCount = lbAll.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbAll.List(b, 0) > lbAll.List(b + 1, 0) Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbAll.Selected(b) = True
                    lbAll.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbAll.List(b + 1, 0)
                strAtt(1) = lbAll.List(b + 1, 1)
                strAtt(2) = lbAll.List(b + 1, 2)
                strAtt(3) = lbAll.List(b + 1, 3)
                strAtt(4) = lbAll.List(b + 1, 4)
                strAtt(5) = lbAll.List(b + 1, 5)
                
                lbAll.List(b + 1, 0) = lbAll.List(b, 0)
                lbAll.List(b + 1, 1) = lbAll.List(b, 1)
                lbAll.List(b + 1, 2) = lbAll.List(b, 2)
                lbAll.List(b + 1, 3) = lbAll.List(b, 3)
                lbAll.List(b + 1, 4) = lbAll.List(b, 4)
                lbAll.List(b + 1, 5) = lbAll.List(b, 5)
                
                lbAll.List(b, 0) = strAtt(0)
                lbAll.List(b, 1) = strAtt(1)
                lbAll.List(b, 2) = strAtt(2)
                lbAll.List(b, 3) = strAtt(3)
                lbAll.List(b, 4) = strAtt(4)
                lbAll.List(b, 5) = strAtt(5)
            End If
        Next b
    Next a
End Sub

Private Sub Label24NoCOMM_Click()
    Dim iIndex As Integer
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iCOMM, iNew, iTemp As Integer
    Dim vLine, vAttach As Variant
    Dim strTemp As String
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        iIndex = CInt(lbAll.List(i, 5))
        
        Set objBlock = objSS.Item(iIndex)
        vAttList = objBlock.GetAttributes
        
        If vAttList(15).TextString = "" Then GoTo Next_I
        
        vLine = Split(vAttList(15).TextString, " ")
        
        If UBound(vLine) > 0 Then GoTo Next_I
        
        strTemp = Replace(UCase(vLine(0)), "T", "")
            
        vAttach = Split(strTemp, "-")
        iNew = CInt(vAttach(0)) * 12
        If UBound(vAttach) > 0 Then iNew = iNew + CInt(vAttach(1))
        
        iMax = FindMaxHeight(iIndex)
        
        'MsgBox "Pwr Max:  " & iMax
        
        If iMax < 288 Then GoTo Next_I
        
        iCOMM = FindTopCOMM(iIndex)
        
        'MsgBox "Top COMM:  " & iCOMM
        
        If iCOMM > 0 Then GoTo Next_I
        
        lbAll.RemoveItem i
Next_I:
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub LabelCOMM1_Click()
    Dim iIndex As Integer
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iCOMM, iNew, iTemp As Integer
    Dim vLine, vAttach As Variant
    Dim strTemp As String
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        iIndex = CInt(lbAll.List(i, 5))
        
        Set objBlock = objSS.Item(iIndex)
        vAttList = objBlock.GetAttributes
        
        If vAttList(15).TextString = "" Then GoTo Next_I
        
        iMax = FindMaxHeight(iIndex)
        
        iCOMM = FindTopCOMM(iIndex) + 12
        If iCOMM = 1012 Then GoTo Next_I
        
        vLine = Split(vAttList(15).TextString, " ")
        For j = 0 To UBound(vLine)
            strTemp = Replace(UCase(vLine(j)), "T", "")
            vAttach = Split(strTemp, "-")
            
            iNew = CInt(vAttach(0)) * 12
            If UBound(vAttach) > 0 Then iNew = iNew + CInt(vAttach(1))
            
            If iNew > iMax Then GoTo Next_I
            If Not iNew = iCOMM Then GoTo Next_I
        Next j
        
        lbAll.RemoveItem i
Next_I:
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub LabelCOMM2_Click()
    Dim iIndex As Integer
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iCOMM, iNew, iTemp As Integer
    Dim vLine, vAttach As Variant
    Dim strTemp As String
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        iIndex = CInt(lbAll.List(i, 5))
        
        Set objBlock = objSS.Item(iIndex)
        vAttList = objBlock.GetAttributes
        
        If vAttList(15).TextString = "" Then GoTo Next_I
        
        iMax = FindMaxHeight(iIndex)
        
        iCOMM = FindTopCOMM(iIndex) + 24
        If iCOMM = 1024 Then GoTo Next_I
        
        vLine = Split(vAttList(15).TextString, " ")
        For j = 0 To UBound(vLine)
            strTemp = Replace(UCase(vLine(j)), "T", "")
            vAttach = Split(strTemp, "-")
            
            iNew = CInt(vAttach(0)) * 12
            If UBound(vAttach) > 0 Then iNew = iNew + CInt(vAttach(1))
            
            If iNew > iMax Then GoTo Next_I
            If Not iNew = iCOMM Then GoTo Next_I
        Next j
        
        lbAll.RemoveItem i
Next_I:
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub LabelMRQC_Click()
    For i = lbAll.ListCount - 1 To 0 Step -1
        If InStr(lbAll.List(i, 4), "MR-QC") > 0 Then lbAll.RemoveItem i
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub LabelNoNew_Click()
    For i = lbAll.ListCount - 1 To 0 Step -1
        If lbAll.List(i, 2) = "none" Then lbAll.RemoveItem i
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub LabelTAG_Click()
    For i = lbAll.ListCount - 1 To 0 Step -1
        If InStr(UCase(lbAll.List(i, 2)), "T") > 0 Then lbAll.RemoveItem i
    Next i
    
    tbListCount = lbAll.ListCount
End Sub

Private Sub lbAll_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'Dim str As String
    Dim objTemp As AcadBlockReference
    Dim vCoords, vAttList As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    Dim iIndex, iList As Integer
    
    Me.Hide
    
    iIndex = CInt(lbAll.List(lbAll.ListIndex, 5))
    Set objTemp = objSS.Item(iIndex)
    
    vCoords = objTemp.InsertionPoint
    
    'MsgBox iIndex
    
    'vCoords = Split(lbAll.List(lbAll.ListIndex, 5), ",")
    
    viewCoordsB(0) = vCoords(0) - 300
    viewCoordsB(1) = vCoords(1) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 300
    viewCoordsE(1) = vCoords(1) + 300
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Load MRWorksheet
        Set MRWorksheet.objPole = objSS.Item(iIndex)
        Call MRWorksheet.GetPoleData
        Call MRWorksheet.SortList
        
        MRWorksheet.cbGetAttach.Enabled = False
        MRWorksheet.cbSaveData.SetFocus
        MRWorksheet.show
    
        vAttList = objTemp.GetAttributes
        
        If MRWorksheet.iChanged = 1 Then
            If InStr(vAttList(24).TextString, "MR-QC") = 0 Then
                vAttList(24).TextString = "MR-QC;" & vAttList(24).TextString
            End If
            objTemp.Update
            
            lbAll.List(lbAll.ListIndex, 4) = vAttList(24).TextString
            iList = CInt(lbAll.List(lbAll.ListIndex, 5))
            lbList.List(iList, 4) = vAttList(24).TextString
        End If
    Unload MRWorksheet
    
    If Not lbAll.ListCount = lbAll.ListIndex + 1 Then lbAll.ListIndex = lbAll.ListIndex + 1
    lbAll.Selected(lbAll.ListIndex) = True
    
    Me.show
End Sub

Private Sub lbAll_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Dim objTemp As AcadBlockReference
            Dim vCoords, vAttList As Variant
            Dim viewCoordsB(0 To 2) As Double
            Dim viewCoordsE(0 To 2) As Double
            Dim iIndex As Integer
    
            Me.Hide
    
            iIndex = CInt(lbAll.List(lbAll.ListIndex, 5))
            Set objTemp = objSS.Item(iIndex)
    
            vCoords = objTemp.InsertionPoint
    
            viewCoordsB(0) = vCoords(0) - 300
            viewCoordsB(1) = vCoords(1) - 300
            viewCoordsB(2) = 0#
            viewCoordsE(0) = vCoords(0) + 300
            viewCoordsE(1) = vCoords(1) + 300
            viewCoordsE(2) = 0#
    
            ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
            Load MRWorksheet
                Set MRWorksheet.objPole = objSS.Item(iIndex)
                Call MRWorksheet.GetPoleData
                Call MRWorksheet.SortList
                
                MRWorksheet.cbGetAttach.Enabled = False
                MRWorksheet.cbSaveData.SetFocus
                MRWorksheet.show
    
                vAttList = objTemp.GetAttributes
        
                If MRWorksheet.iChanged = 1 Then
                    If InStr(vAttList(24).TextString, "MR-QC") = 0 Then
                        vAttList(24).TextString = "MR-QC;" & vAttList(24).TextString
                    End If
                    lbAll.List(lbAll.ListIndex, 4) = vAttList(24).TextString
                End If
            Unload MRWorksheet
            
            If Not lbAll.ListCount = lbAll.ListIndex + 1 Then lbAll.ListIndex = lbAll.ListIndex + 1
            lbAll.Selected(lbAll.ListIndex) = True
    
            Me.show
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbAll.ColumnCount = 6
    lbAll.ColumnWidths = "126;48;72;72;96;12"
    
    lbList.ColumnCount = 6
    lbList.ColumnWidths = "30;30;30;30;30;30"
    
    cbSSType.AddItem "All"
    cbSSType.AddItem "Window"
    cbSSType.Value = "Window"
    
    cbFROption.AddItem "Filter"
    cbFROption.AddItem "Remove"
    cbFROption.Value = "Filter"
    
    cbFRColumn.AddItem "0: Pole Number"
    cbFRColumn.AddItem "1: Owner"
    cbFRColumn.AddItem "2: New At"
    cbFRColumn.AddItem "3: MR"
    cbFRColumn.AddItem "4: Status"
    'cbFRColumn.Value = "1: Owner"
    
    Call GetAllPoles
    
    tbListCount.Value = lbAll.ListCount
    
    'Call FindOwnerStatus
End Sub

Private Sub GetAllPoles()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    'Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine, vTemp As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim iCount, iIndex, iTest As Integer
    
    On Error Resume Next
    
    lbAll.Clear
    cbOwner.Clear
    cbStatus.Clear
    
    iCount = 0
    iIndex = 0
    iTest = 0
    
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
    
    'Select Case cbSSType.Value
        'Case "All"
            'objSS.Select acSelectionSetAll, , , filterType, filterValue
        'Case Else
            vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
            vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
            
            objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    'End Select
    
    iCount = objSS.count
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        vAttList = objBlock.GetAttributes
        
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        'If cbInclude.Value = False Then
            'If vAttList(0).TextString = "POLE" Then GoTo Next_objBlock
        'End If
        
        lbAll.AddItem vAttList(0).TextString
        lbAll.List(iIndex, 1) = vAttList(2).TextString
        If vAttList(15).TextString = "" Then
            lbAll.List(iIndex, 2) = "none"
        Else
            lbAll.List(iIndex, 2) = vAttList(15).TextString
        End If
        lbAll.List(iIndex, 3) = "none"
        
        For j = 16 To 23
            If Not vAttList(j).TextString = "" Then
                vLine = Split(vAttList(j).TextString, "=")
                    
                If InStr(vLine(1), ")") > 0 Then
                    If lbAll.List(iIndex, 3) = "none" Then
                        lbAll.List(iIndex, 3) = vLine(0)
                    Else
                        lbAll.List(iIndex, 3) = lbAll.List(iIndex, 3) & ", " & vLine(0)
                    End If
                End If
                
                If InStr(LCase(vLine(1)), "v") > 0 Then
                    If lbAll.List(iIndex, 3) = "none" Then
                        lbAll.List(iIndex, 3) = "LASHED"
                    Else
                        lbAll.List(iIndex, 3) = "LASHED, " & lbAll.List(iIndex, 3)
                    End If
                End If
            End If
        Next j
        
        For j = 9 To 23
            If Not vAttList(j).TextString = "" Then iTest = 1
        Next j
        
        If iTest = 0 Then lbAll.List(iIndex, 3) = "no attach"
        
        'If InStr(vAttList(24).TextString, "MR-QC;") > 0 Then
            'lbAll.List(iIndex, 4) = "MR-QC"
        'Else
            'lbAll.List(iIndex, 4) = ""
        'End If
        lbAll.List(iIndex, 4) = vAttList(24).TextString
        
        lbAll.List(iIndex, 5) = i
        
        iIndex = iIndex + 1
        
Next_objBlock:
    Next i
    
    If lbAll.ListCount > 0 Then
        lbList.Clear
        
        For i = 0 To lbAll.ListCount - 1
            lbList.AddItem lbAll.List(i, 0)
            lbList.List(i, 1) = lbAll.List(i, 1)
            lbList.List(i, 2) = lbAll.List(i, 2)
            lbList.List(i, 3) = lbAll.List(i, 3)
            lbList.List(i, 4) = lbAll.List(i, 4)
            lbList.List(i, 5) = lbAll.List(i, 5)
        Next i
        
        If cbInclude.Value = False Then
            For i = lbAll.ListCount - 1 To 0 Step -1
                If lbAll.List(i, 0) = "POLE" Then lbAll.RemoveItem i
            Next i
        End If
    End If
    
    'objSS.Clear
    'objSS.Delete
    
    tbListCount.Value = lbAll.ListCount
    
    'If Not cbSSType.Value = "All" Then Me.show
End Sub

Private Sub UserForm_Terminate()
    objSS.Clear
    objSS.Delete
End Sub

Private Function FindMaxHeight(iIndex As Integer)
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iTemp As Integer
    Dim vLine, vAttach, vTemp As Variant
    
    iMax = 840  ' 70 feet
        
    Set objBlock = objSS.Item(iIndex)
    vAttList = objBlock.GetAttributes
        
        If Not vAttList(9).TextString = "" Then
            vLine = Split(vAttList(9).TextString, " ")
            
            For j = 0 To UBound(vLine)
                If InStr(vLine(j), ")") > 0 Then
                    vTemp = Split(vLine(j), ")")
                    vAttach = Split(vTemp(1), "-")
                Else
                    vAttach = Split(vLine(j), "-")
                End If
                
                iTemp = CInt(vAttach(0)) * 12
                If UBound(vAttach) > 0 Then iTemp = iTemp + CInt(vAttach(1))
                
                iTemp = iTemp - 40
                
                If iTemp < iMax Then iMax = iTemp
            Next j
        End If
        
        If Not vAttList(10).TextString = "" Then
            vLine = Split(vAttList(10).TextString, " ")
            
            For j = 0 To UBound(vLine)
                If InStr(vLine(j), ")") > 0 Then
                    vTemp = Split(vLine(j), ")")
                    vAttach = Split(vTemp(1), "-")
                Else
                    vAttach = Split(vLine(j), "-")
                End If
                
                iTemp = CInt(vAttach(0)) * 12
                If UBound(vAttach) > 0 Then iTemp = iTemp + CInt(vAttach(1))
                
                iTemp = iTemp - 30
                
                If iTemp < iMax Then iMax = iTemp
            Next j
        End If
        
        If Not vAttList(11).TextString = "" Then
            vLine = Split(vAttList(11).TextString, " ")
            
            For j = 0 To UBound(vLine)
                If InStr(vLine(j), ")") > 0 Then
                    vTemp = Split(vLine(j), ")")
                    vAttach = Split(vTemp(1), "-")
                Else
                    vAttach = Split(vLine(j), "-")
                End If
                
                iTemp = CInt(vAttach(0)) * 12
                If UBound(vAttach) > 0 Then iTemp = iTemp + CInt(vAttach(1))
                
                iTemp = iTemp - 40
                
                If iTemp < iMax Then iMax = iTemp
            Next j
        End If
        
        If Not vAttList(12).TextString = "" Then
            vLine = Split(vAttList(12).TextString, " ")
            
            For j = 0 To UBound(vLine)
                If InStr(vLine(j), ")") > 0 Then
                    vTemp = Split(vLine(j), ")")
                    vAttach = Split(vTemp(1), "-")
                Else
                    vAttach = Split(vLine(j), "-")
                End If
                
                iTemp = CInt(vAttach(0)) * 12
                If UBound(vAttach) > 0 Then iTemp = iTemp + CInt(vAttach(1))
                
                iTemp = iTemp - 40
                
                If iTemp < iMax Then iMax = iTemp
            Next j
        End If
        
        If Not vAttList(13).TextString = "" Then
            vLine = Split(vAttList(13).TextString, " ")
            
            For j = 0 To UBound(vLine)
                If InStr(vLine(j), ")") > 0 Then
                    vTemp = Split(vLine(j), ")")
                    vAttach = Split(vTemp(1), "-")
                Else
                    vAttach = Split(vLine(j), "-")
                End If
                
                iTemp = CInt(vAttach(0)) * 12
                If UBound(vAttach) > 0 Then iTemp = iTemp + CInt(vAttach(1))
                
                iTemp = iTemp - 40
                
                If iTemp < iMax Then iMax = iTemp
            Next j
        End If
        
        If Not vAttList(14).TextString = "" Then
            vLine = Split(vAttList(14).TextString, " ")
            
            For j = 0 To UBound(vLine)
                If InStr(vLine(j), ")") > 0 Then
                    vTemp = Split(vLine(j), ")")
                    vAttach = Split(vTemp(1), "-")
                Else
                    vAttach = Split(vLine(j), "-")
                End If
                
                iTemp = CInt(vAttach(0)) * 12
                If UBound(vAttach) > 0 Then iTemp = iTemp + CInt(vAttach(1))
                
                iTemp = iTemp - 12
                
                If iTemp < iMax Then iMax = iTemp
            Next j
        End If
        
    FindMaxHeight = iMax
End Function

Private Function FindTopCOMM(iIndex As Integer)
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iTemp As Integer
    Dim vTemp, vLine, vAttach As Variant
    Dim strTemp As String
    
    iMax = 0
        
    Set objBlock = objSS.Item(iIndex)
    vAttList = objBlock.GetAttributes
    
    For i = 16 To 23
        If Not vAttList(i).TextString = "" Then
            strTemp = Replace(UCase(vAttList(i).TextString), "C", "")
            'strTemp = Replace(strTemp, "D", "")
            strTemp = Replace(strTemp, "O", "")
            strTemp = Replace(strTemp, "S", "")
            strTemp = Replace(strTemp, "X", "")
            
            vTemp = Split(strTemp, "=")
            vLine = Split(vTemp(1), " ")
            
            For j = 0 To UBound(vLine)
                If InStr(vLine(j), "D") > 0 Then GoTo Next_J
                If InStr(vLine(j), ")") > 0 Then
                    FindTopCOMM = 1000
                    Exit Function
                    'vTemp = Split(vLine(j), ")")
                    'vAttach = Split(vTemp(1), "-")
                Else
                    vAttach = Split(vLine(j), "-")
                End If
                
                iTemp = CInt(vAttach(0)) * 12
                If UBound(vAttach) > 0 Then iTemp = iTemp + CInt(vAttach(1))
                
                If iTemp > iMax Then iMax = iTemp
Next_J:
            Next j
        End If
    Next i
        
    FindTopCOMM = iMax
End Function

Private Function FindNewAttach(iIndex As Integer)
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iMax, iTemp As Integer
    Dim vLine, vAttach As Variant
    Dim strTemp As String
    
    iMax = 0
        
    Set objBlock = objSS.Item(iIndex)
    vAttList = objBlock.GetAttributes
    
    If vAttList(15).TextString = "" Then
        vLine = 0
    Else
        vLine = Split(vAttList(15).TextString, " ")
        
        For i = 0 To UBound(vLine)
            strTemp = Replace(UCase(vLine(i)), "T", "")
            
            vAttach = Split(strTemp, "-")
            iTemp = CInt(vAttach(0)) * 12
            If UBound(vAttach) > 0 Then iTemp = iTemp + CInt(vAttach(1))
                
            vLine(i) = iTemp
        Next i
    End If
    
    FindNewAttach = vLine
End Function

Private Sub FindOwnerStatus()
    If lbAll.ListCount < 2 Then Exit Sub
    
    'MsgBox "Here"
    Dim vLine As Variant
    
    cbOwner.Clear
    cbStatus.Clear
    
    cbOwner.AddItem lbAll.List(0, 1)
    
    If Not lbAll.List(0, 4) = "" Then
        vLine = Split(lbAll.List(0, 4), ";")
        For j = 0 To UBound(vLine)
            cbStatus.AddItem vLine(j)
        Next j
    End If
    
    For i = 1 To lbAll.ListCount - 1
        For j = 0 To cbOwner.ListCount - 1
            If cbOwner.List(j) = lbAll.List(i, 1) Then GoTo Exit_J
        Next j
        
        cbOwner.AddItem lbAll.List(i, 1)
Exit_J:
        
        If Not lbAll.List(i, 4) = "" Then
            vLine = Split(lbAll.List(i, 4), ";")
            
            For j = 0 To UBound(vLine)
                If Not cbStatus.ListCount < 0 Then
                    For k = 0 To cbStatus.ListCount - 1
                        If cbStatus.List(k) = vLine(j) Then GoTo Exit_K
                    Next k
                End If
                
                cbStatus.AddItem vLine(j)
Exit_K:
            Next j
        End If
    Next i
    
    If cbOwner.ListCount > 0 Then
        cbFRValue.Clear
        For i = 0 To cbOwner.ListCount - 1
            cbFRValue.AddItem cbOwner.List(i)
        Next i
    End If
    
    If cbStatus.ListCount > 0 Then
        For i = cbStatus.ListCount - 1 To 0 Step -1
            If cbStatus.List(i) = "" Then cbStatus.RemoveItem i
            'If cbStatus.List(i) = " " Then cbStatus.RemoveItem i
        Next i
    End If
End Sub

Private Sub FilterRemovePN()
    If lbAll.ListCount < 2 Then Exit Sub
    'If cbFRValue.Value = "" Then Exit Sub
    
    Dim strOwner As String
    
    strOwner = UCase(cbFRValue.Value)
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        If cbFROption.Value = "Filter" Then
            If InStr(lbAll.List(i, 0), strOwner) = 0 Then lbAll.RemoveItem i
        Else
            If InStr(lbAll.List(i, 0), strOwner) > 0 Then lbAll.RemoveItem i
        End If
    Next i
    
    tbListCount.Value = lbAll.ListCount
End Sub

Private Sub FilterRemoveOwner()
    If lbAll.ListCount < 2 Then Exit Sub
    If cbFRValue.Value = "" Then Exit Sub
    
    Dim strOwner As String
    
    strOwner = cbFRValue.Value
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        If cbFROption.Value = "Filter" Then
            If Not lbAll.List(i, 1) = strOwner Then lbAll.RemoveItem i
        Else
            If lbAll.List(i, 1) = strOwner Then lbAll.RemoveItem i
        End If
    Next i
    
    tbListCount.Value = lbAll.ListCount
End Sub

Private Sub FilterRemoveStatus()
    If lbAll.ListCount < 2 Then Exit Sub
    'If cbFRValue.Value = "" Then Exit Sub
    
    Dim vLine As Variant
    Dim strStatus, strLine As String
    Dim iCount As Integer
    
    strStatus = cbFRValue.Value
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        iCount = 0
        
        If Not lbAll.List(i, 4) = "" Then
            vLine = Split(lbAll.List(i, 4), ";")
            For j = 0 To UBound(vLine)
                If vLine(j) = strStatus Then iCount = 1
            Next j
        End If
            
        If cbFROption.Value = "Filter" Then
            If iCount = 0 Then lbAll.RemoveItem i
        Else
            If iCount = 1 Then lbAll.RemoveItem i
        End If
    Next i
    
    tbListCount.Value = lbAll.ListCount
End Sub

Private Sub FilterRemoveNewAt()
    If lbAll.ListCount < 2 Then Exit Sub
    If cbFRValue.Value = "" Then Exit Sub
    
    Dim vLine, vTemp As Variant
    Dim strType, strLine As String
    Dim iSHeight, iCHeight As Integer
    Dim iCount As Integer
    
    Select Case cbFRValue.Value
        Case "none"
            For i = lbAll.ListCount - 1 To 0 Step -1
                If cbFROption.Value = "Filter" Then
                    If Not lbAll.List(i, 2) = "none" Then lbAll.RemoveItem i
                Else
                    If lbAll.List(i, 2) = "none" Then lbAll.RemoveItem i
                End If
            Next i
        Case "**T"
            For i = lbAll.ListCount - 1 To 0 Step -1
                If cbFROption.Value = "Filter" Then
                    If InStr(LCase(lbAll.List(i, 2)), "t") = 0 Then lbAll.RemoveItem i
                Else
                    If InStr(LCase(lbAll.List(i, 2)), "t") > 0 Then lbAll.RemoveItem i
                End If
            Next i
        Case "**F"
            For i = lbAll.ListCount - 1 To 0 Step -1
                If cbFROption.Value = "Filter" Then
                    If InStr(LCase(lbAll.List(i, 2)), "f") = 0 Then lbAll.RemoveItem i
                Else
                    If InStr(LCase(lbAll.List(i, 2)), "f") > 0 Then lbAll.RemoveItem i
                End If
            Next i
        Case "**O"
            For i = lbAll.ListCount - 1 To 0 Step -1
                If cbFROption.Value = "Filter" Then
                    If InStr(LCase(lbAll.List(i, 2)), "o") = 0 Then lbAll.RemoveItem i
                Else
                    If InStr(LCase(lbAll.List(i, 2)), "o") > 0 Then lbAll.RemoveItem i
                End If
            Next i
        Case Else
            If InStr(cbFRValue.Value, "=") = 0 Then Exit Sub
            
            vLine = Split(cbFRValue.Value, "=")
            If vLine(1) = "" Then Exit Sub
            
            strType = vLine(0)
            
            vTemp = Split(vLine(1), "-")
            iSHeight = CInt(vTemp(0)) * 12
            If UBound(vTemp) > 0 Then iSHeight = iSHeight + CInt(vTemp(1))
            
            For i = lbAll.ListCount - 1 To 0 Step -1
                'If InStr(lbAll.List(i, 2), " ") > 0 Then GoTo Next_NewAt
                If lbAll.List(i, 2) = "none" Then
                    lbAll.RemoveItem i
                    GoTo Next_NewAt
                End If
                
                strLine = LCase(lbAll.List(i, 2))
                strLine = Replace(strLine, "o", "")
                strLine = Replace(strLine, "f", "")
                strLine = Replace(strLine, "t", "")
                vLine = Split(strLine, " ")
                
                iCount = 0
                For j = 0 To UBound(vLine)
                    vTemp = Split(vLine(j), "-")
                    iCHeight = CInt(vTemp(0)) * 12
                    If UBound(vTemp) > 0 Then iCHeight = iCHeight + CInt(vTemp(1))
                    
                    Select Case strType
                        Case "<"
                            If iCHeight < iSHeight Then iCount = iCount + 1
                        Case ">"
                            If iCHeight > iSHeight Then iCount = iCount + 1
                        Case Else
                            If iCHeight = iSHeight Then iCount = iCount + 1
                    End Select
                Next j
                    
                If cbFROption.Value = "Filter" Then
                    If iCount = 0 Then lbAll.RemoveItem i
                Else
                    If iCount = UBound(vTemp) + 1 Then lbAll.RemoveItem i
                End If
Next_NewAt:
            Next i
    End Select
        
    tbListCount.Value = lbAll.ListCount
    
    
    
    
    
    
    
    
    
    
    
    

            'cbFRValue.AddItem "none"
            'cbFRValue.AddItem "**T"
            'cbFRValue.AddItem "**F"
            'cbFRValue.AddItem "**E"
            'cbFRValue.AddItem "<="
            'cbFRValue.AddItem ">="
            'cbFRValue.AddItem "="
End Sub

Private Sub FilterRemoveMR()
    If lbAll.ListCount < 2 Then Exit Sub
    If cbFRValue.Value = "" Then Exit Sub
    
    'Dim vLine As Variant
    
    If cbFRValue.Value = "none" Then
        For i = lbAll.ListCount - 1 To 0 Step -1
            If cbFROption.Value = "Filter" Then
                If Not lbAll.List(i, 3) = "none" Then lbAll.RemoveItem i
            Else
                If lbAll.List(i, 3) = "none" Then lbAll.RemoveItem i
            End If
        Next i
        
        tbListCount.Value = lbAll.ListCount
        Exit Sub
    End If
    
    Dim vLine As Variant
    Dim strMR As String
    Dim iCount As Integer
    
    strMR = cbFRValue.Value
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        If InStr(lbAll.List(i, 3), ",") > 0 Then
            vLine = Split(lbAll.List(i, 3), ", ")
            iCount = 0
            For j = 0 To UBound(vLine)
                If vLine(j) = strMR Then
                    iCount = 1
                End If
            Next j
            
            If cbFROption.Value = "Filter" And iCount = 0 Then lbAll.RemoveItem i
        Else
            If cbFROption.Value = "Filter" Then
                If Not lbAll.List(i, 3) = strMR Then lbAll.RemoveItem i
            Else
                If lbAll.List(i, 3) = strMR Then lbAll.RemoveItem i
            End If
        End If
'Next_MR:
    Next i
        
    tbListCount.Value = lbAll.ListCount
End Sub
