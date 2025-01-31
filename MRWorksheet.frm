VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MRWorksheet 
   Caption         =   "MR Worksheet"
   ClientHeight    =   8790.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   OleObjectBlob   =   "MRWorksheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MRWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public objPole As AcadBlockReference
Public iChanged As Integer

Dim iListIndex As Integer
Dim iSave As Integer
Dim iMaxFt, iMaxIn As Integer
Dim strAttachmentItem As String
Dim vCoordsAttach As Variant
Dim dPosition As Double
Dim dRotate As Double
Dim dInsertPnt(0 To 2) As Double

Private Sub cbAddCLR_Click()
    Dim iListIndex As Integer
    
    lbMR.AddItem
    iListIndex = lbMR.ListCount - 1
    
    If cbAddCLR.Caption = "Update" Then
        iListIndex = lbMR.ListIndex
        lbMR.RemoveItem iListIndex
        lbMR.AddItem "", iListIndex
        
        cbAddCLR.Caption = "Add CLR"
    End If
    
    lbMR.List(iListIndex, 0) = tbMRCompany.Value
    lbMR.List(iListIndex, 1) = tbMR.Value
    lbMR.List(iListIndex, 2) = cbType.Value
    lbMR.List(iListIndex, 3) = tbSpan.Value
    lbMR.List(iListIndex, 4) = tbDist.Value
    If Not tbEF.Value = "" Then
        lbMR.List(iListIndex, 5) = tbEF.Value & "-" & tbEI.Value
    Else
        lbMR.List(iListIndex, 5) = ""
    End If
    If Not tbPF.Value = "" Then
        lbMR.List(iListIndex, 6) = tbPF.Value & "-" & tbPI.Value
    Else
        lbMR.List(iListIndex, 6) = ""
    End If
    
    tbMRCompany.Value = ""
    tbMR.Value = ""
    cbType.Value = ""
    tbSpan.Value = ""
    tbDist.Value = ""
    tbEF.Value = ""
    tbEI.Value = ""
    tbPF.Value = ""
    tbPI.Value = ""
    
    Call RemoveEmptyRows
End Sub

Private Sub cbAddNote_Click()
    If cbNote.Value = "" Then Exit Sub
    
    Dim strNote As String
    Dim vAttList As Variant
    
    Select Case cbNote.Value
        Case "Pwr replace with 5 Taller Pole"
            strNote = "NOTE-TALL"
        Case "Pwr replace with 10 Taller Pole"
            strNote = "NOTE-TALL10"
        Case "Pwr replace Defective Pole"
            strNote = "NOTE-DEF"
        Case "Pwr Raise Attachment"
            strNote = "NOTE-PRA"
        Case "Replace Note with Other"
            strNote = "NOTE-OTHER"
        Case "Use Existing Holes"
            strNote = "NOTE-UEH"
        Case "Pole Permited Previously"
            vAttList = objPole.GetAttributes
            
            strNote = "NOTE-REF=" & vAttList(1).TextString
        Case "OCALC Required"
            strNote = "OCALC"
        Case "Extra Seperation"
            strNote = "NOTE-EC"
        Case Else
            MsgBox "Empty"
            Exit Sub
    End Select
    
    'MsgBox strNote
    
    If tbNoteList.Value = "" Then
        tbNoteList.Value = strNote
    Else
        tbNoteList.Value = tbNoteList.Value & vbCr & strNote
    End If
End Sub

Private Sub cbDelete_Click()
    If lbWorkspace.ListIndex < 0 Then Exit Sub
    
    lbWorkspace.RemoveItem lbWorkspace.ListIndex
    iSave = 1
End Sub

Private Sub cbDOWN_Click()
    Dim str0, str1, str2 As String
    Dim str3, str4, str5 As String
    Dim i, i2 As Integer
    
    If lbWorkspace.ListIndex < 0 Then Exit Sub
    If lbWorkspace.ListIndex = (lbWorkspace.ListCount - 1) Then Exit Sub
    i = lbWorkspace.ListIndex
    i2 = i + 1
    
    str0 = lbWorkspace.List(i, 0)
    str1 = lbWorkspace.List(i, 1)
    str2 = lbWorkspace.List(i, 2)
    str3 = lbWorkspace.List(i, 3)
    str4 = lbWorkspace.List(i, 4)
    str5 = lbWorkspace.List(i, 5)
    
    lbWorkspace.List(i, 0) = lbWorkspace.List(i2, 0)
    lbWorkspace.List(i, 1) = lbWorkspace.List(i2, 1)
    lbWorkspace.List(i, 2) = lbWorkspace.List(i2, 2)
    lbWorkspace.List(i, 3) = lbWorkspace.List(i2, 3)
    lbWorkspace.List(i, 4) = lbWorkspace.List(i2, 4)
    lbWorkspace.List(i, 5) = lbWorkspace.List(i2, 5)
    
    lbWorkspace.List(i2, 0) = str0
    lbWorkspace.List(i2, 1) = str1
    lbWorkspace.List(i2, 2) = str2
    lbWorkspace.List(i2, 3) = str3
    lbWorkspace.List(i2, 4) = str4
    lbWorkspace.List(i2, 5) = str5
    
    lbWorkspace.ListIndex = i2
End Sub

Private Sub cbFiveFootTaller_Click()
    If lbWorkspace.ListCount < 1 Then Exit Sub
    
    Dim vLine, vItem, vHC As Variant
    Dim strLine, strHC, strType As String
    Dim iH, iIncrease As Integer
    
    If Not tbTallerPole.Value = "" Then
        iIncrease = CInt(tbTallerPole.Value)
    Else
        Exit Sub
    End If
    
    strLine = Replace(tbPoleData.Value, vbLf, "")
    vLine = Split(strLine, vbCr)
    vItem = Split(vLine(1), vbTab)
    strLine = vItem(0) & vbTab & "(" & vItem(1) & ")"
    vHC = Split(UCase(vItem(1)), "-")
    
    On Error Resume Next
    
    strType = ""
    
    If InStr(vHC(0), "C") Then
        strType = "C"
        vHC(0) = Replace(vHC(0), "C", "")
    End If
    
    If InStr(vHC(0), "S") Then
        strType = "S"
        vHC(0) = Replace(vHC(0), "S", "")
    End If
    
    iH = CInt(vHC(0)) + iIncrease
    
    If Not Err = 0 Then Exit Sub
    
    strLine = strLine & strType & iH & "-" '& vHC(1)
    
    Select Case iH
        Case Is = 35
            If vHC(1) = "?" Then
                strLine = strLine & "4"
            Else
                If CInt(vHC(1)) < 4 Then
                    strLine = strLine & vHC(1)
                Else
                    strLine = strLine & "4"
                End If
            End If
        Case Is = 40
            If vHC(1) = "?" Then
                strLine = strLine & "3"
            Else
                If CInt(vHC(1)) < 3 Then
                    strLine = strLine & vHC(1)
                Else
                    strLine = strLine & "3"
                End If
            End If
        Case Is = 45, Is = 50
            If vHC(1) = "?" Then
                strLine = strLine & "2"
            Else
                If CInt(vHC(1)) < 3 Then
                    strLine = strLine & vHC(1)
                Else
                    strLine = strLine & "2"
                End If
            End If
        Case Is > 50
            If vHC(1) = "?" Then
                strLine = strLine & "1"
            Else
                If CInt(vHC(1)) < 2 Then
                    strLine = strLine & vHC(1)
                Else
                    strLine = strLine & "1"
                End If
            End If
    End Select
    vLine(1) = strLine
    
    strLine = vLine(0)
    
    For i = 1 To UBound(vLine)
        strLine = strLine & vbCr & vLine(i)
    Next i
    
    tbPoleData.Value = strLine
    
    iIncrease = CInt(iIncrease * 0.6)
    
    For i = 0 To lbWorkspace.ListCount - 1
        Select Case lbWorkspace.List(i, 0)
            Case "PWR"
                vLine = Split(lbWorkspace.List(i, 2), "-")
                lbWorkspace.List(i, 3) = CInt(vLine(0)) + iIncrease & "-" & vLine(1)
            Case "COMM"
                lbWorkspace.List(i, 3) = lbWorkspace.List(i, 2)
        End Select
        
        'If lbWorkspace.List(i, 0) = "PWR" Then
            'vLine = Split(lbWorkspace.List(i, 2), "-")
            'lbWorkspace.List(i, 3) = CInt(vLine(0)) + iIncrease & "-" & vLine(1)
        'End If
    Next i
    
    Call SortList
    
    Select Case tbTallerPole.Value
        Case "5"
            cbNote.Value = "Pwr replace with 5 Taller Pole"
        Case "10"
            cbNote.Value = "Pwr replace with 10 Taller Pole"
    End Select
End Sub

Private Sub cbFixDown_Click()
    Dim vTemp, vAttach, vCW As Variant
    Dim strLine As String
    Dim iCurrentIn, iAttachIn As Integer
    Dim iTempFt, iTempIn As Integer
    
    iAttachIn = 1000
    
    For i = 0 To lbWorkspace.ListCount - 1
        CheckForViolations (i)
    Next i
    
    For i = 0 To lbWorkspace.ListCount - 1
        Select Case lbWorkspace.List(i, 0)
            Case "PWR"
                vTemp = Split(lbWorkspace.List(i, 4), " ")
                vAttach = Split(vTemp(1), "-")
                iCurrentIn = CInt(vAttach(0)) * 12 + CInt(vAttach(1))
                
                If iCurrentIn < iAttachIn Then iAttachIn = iCurrentIn
            Case "UTC"
                iTempFt = CInt(iAttachIn / 12 - 0.5)
                iTempIn = iAttachIn - (iTempFt * 12)
                If iTempIn > 11 Then
                    iTempIn = iTempIn - 12
                    iTempFt = iTempFt + 1
                End If
                
                lbWorkspace.List(i, 3) = iTempFt & "-" & iTempIn
                iAttachIn = iAttachIn - CInt(tbBC.Value)
            Case "COMM"
                iTempFt = CInt(iAttachIn / 12 - 0.5)
                iTempIn = iAttachIn - (iTempFt * 12)
                If iTempIn > 11 Then
                    iTempIn = iTempIn - 12
                    iTempFt = iTempFt + 1
                End If
                
                lbWorkspace.List(i, 3) = iTempFt & "-" & iTempIn
                
                If Not lbWorkspace.ListCount = i + 1 Then
                    If lbWorkspace.List(i, 1) = lbWorkspace.List(i + 1, 1) Then
                        iAttachIn = iAttachIn - CInt(tbBSC.Value)
                    Else
                        iAttachIn = iAttachIn - CInt(tbBC.Value)
                    End If
                End If
        End Select
    Next i
End Sub

Private Sub cbGetAttach_Click()
    If iSave = 1 Then
        result = MsgBox("Save changes to Pole Attachments?", vbYesNo, "Save Changes")
        If result = vbYes Then
            Call SaveAttachments
            iSave = 0
        End If
    End If
    
    Dim objEntity As AcadEntity
    Dim vReturnPnt As Variant
    
    Me.Hide
    
    On Error Resume Next
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Pole: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objPole = objEntity
    
    If Not objPole.Name = "sPole" Then GoTo Exit_Sub
    
    Call GetPoleData
    
    Call SortList
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbGetData_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim vAttList As Variant
    Dim vLine, vTemp As Variant
    Dim vItem, vMR, vAttach As Variant
    Dim iEInch, iPInch As Integer
    Dim strLine, strCO As String
    Dim strTemp As String
    Dim strPrompt As String
    Dim iSpan, iAttach As Integer
    Dim iTest, iIndex As Integer
    Dim dEF As Double
    Dim lColor As Long
    
    On Error Resume Next
    
    Me.Hide
    
    tbPrePF.Value = ""
    tbPrePI.Value = ""
    
    iSpan = 0
    iAttach = 0
    strPrompt = "Select Cable Span & Bottom Attachment:"
    
Get_Block:
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, strPrompt
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    vAttList = objBlock.GetAttributes
    
    Select Case objBlock.Name
        Case "cable_span"
            If vAttList(0).TextString = "" Then GoTo Bad_Block
            iSpan = 1
            vLine = Split(vAttList(0).TextString, " ")   '<----------------------------
            strTemp = vLine(0)
            strLine = vLine(1)
    
            vLine = Split(strLine, "-")
            tbEF.Value = vLine(0)
            tbEI.Value = vLine(1)
            
            iTest = CInt(vLine(0)) * 12 + CInt(vLine(1))
            
            Select Case LCase(Left(strTemp, 1))
                Case "r"
                    cbType.Value = "RD"
                    If iTest < 216 Then
                        lColor = 255
                    Else
                        lColor = 0
                    End If
                Case "d"
                    cbType.Value = "DW"
                    If iTest < 204 Then
                        lColor = 255
                    Else
                        lColor = 0
                    End If
                Case "m"
                    cbType.Value = "MS"
                    If iTest < 204 Then
                        lColor = 255
                    Else
                        lColor = 0
                    End If
                Case Else
                    cbType.Value = "NA"
                    GoTo Exit_Sub
            End Select
            tbEF.ForeColor = lColor
            tbEI.ForeColor = lColor  '<---------------------------------------------------
    
            strTemp = vAttList(2).TextString
            If Right(strTemp, 1) = "'" Then
                strTemp = Left(strTemp, Len(strTemp) - 1)
                tbSpan.Value = strTemp
            Else
                tbSpan.Value = strTemp
            End If
    
            dEF = CDbl(strTemp)
            tbDist.Value = CStr(dEF / 2)
        Case "sPole"
            iAttach = 1
            
            Load SelectAttach
                
                If Not vAttList(9).TextString = "" Then
                    vItem = Split(vAttList(9).TextString, " ")
                    
                    For i = 0 To UBound(vItem)
                        SelectAttach.lbAttach.AddItem "NEUTRAL"
                        iIndex = SelectAttach.lbAttach.ListCount - 1
                        
                        If InStr(vItem(i), ")") > 0 Then
                            vItem(i) = Replace(vItem(i), "(", "")
                            vMR = Split(vItem(i), ")")
                            
                            SelectAttach.lbAttach.List(iIndex, 1) = vMR(0)
                            SelectAttach.lbAttach.List(iIndex, 2) = vMR(1)
                        Else
                            SelectAttach.lbAttach.List(iIndex, 1) = vItem(i)
                            SelectAttach.lbAttach.List(iIndex, 2) = "x"
                        End If
                    Next i
                End If
                
                If Not vAttList(10).TextString = "" Then
                    vItem = Split(vAttList(10).TextString, " ")
                    
                    For i = 0 To UBound(vItem)
                        SelectAttach.lbAttach.AddItem "TRANSFORMER"
                        iIndex = SelectAttach.lbAttach.ListCount - 1
                        
                        If InStr(vItem(i), ")") > 0 Then
                            vItem(i) = Replace(vItem(i), "(", "")
                            vMR = Split(vItem(i), ")")
                            
                            SelectAttach.lbAttach.List(iIndex, 1) = vMR(0)
                            SelectAttach.lbAttach.List(iIndex, 2) = vMR(1)
                        Else
                            SelectAttach.lbAttach.List(iIndex, 1) = vItem(i)
                            SelectAttach.lbAttach.List(iIndex, 2) = "x"
                        End If
                    Next i
                End If
                
                If Not vAttList(11).TextString = "" Then
                    vItem = Split(vAttList(11).TextString, " ")
                    
                    For i = 0 To UBound(vItem)
                        SelectAttach.lbAttach.AddItem "LOW POWER"
                        iIndex = SelectAttach.lbAttach.ListCount - 1
                        
                        If InStr(vItem(i), ")") > 0 Then
                            vItem(i) = Replace(vItem(i), "(", "")
                            vMR = Split(vItem(i), ")")
                            
                            SelectAttach.lbAttach.List(iIndex, 1) = vMR(0)
                            SelectAttach.lbAttach.List(iIndex, 2) = vMR(1)
                        Else
                            SelectAttach.lbAttach.List(iIndex, 1) = vItem(i)
                            SelectAttach.lbAttach.List(iIndex, 2) = "x"
                        End If
                    Next i
                End If
                
                If Not vAttList(12).TextString = "" Then
                    vItem = Split(vAttList(12).TextString, " ")
                    
                    For i = 0 To UBound(vItem)
                        SelectAttach.lbAttach.AddItem "ANTENNA"
                        iIndex = SelectAttach.lbAttach.ListCount - 1
                        
                        If InStr(vItem(i), ")") > 0 Then
                            vItem(i) = Replace(vItem(i), "(", "")
                            vMR = Split(vItem(i), ")")
                            
                            SelectAttach.lbAttach.List(iIndex, 1) = vMR(0)
                            SelectAttach.lbAttach.List(iIndex, 2) = vMR(1)
                        Else
                            SelectAttach.lbAttach.List(iIndex, 1) = vItem(i)
                            SelectAttach.lbAttach.List(iIndex, 2) = "x"
                        End If
                    Next i
                End If
                
                If Not vAttList(13).TextString = "" Then
                    vItem = Split(vAttList(13).TextString, " ")
                    
                    For i = 0 To UBound(vItem)
                        SelectAttach.lbAttach.AddItem "ST LT CIRCUIT"
                        iIndex = SelectAttach.lbAttach.ListCount - 1
                        
                        If InStr(vItem(i), ")") > 0 Then
                            vItem(i) = Replace(vItem(i), "(", "")
                            vMR = Split(vItem(i), ")")
                            
                            SelectAttach.lbAttach.List(iIndex, 1) = vMR(0)
                            SelectAttach.lbAttach.List(iIndex, 2) = vMR(1)
                        Else
                            SelectAttach.lbAttach.List(iIndex, 1) = vItem(i)
                            SelectAttach.lbAttach.List(iIndex, 2) = "x"
                        End If
                    Next i
                End If
                
                If Not vAttList(14).TextString = "" Then
                    vItem = Split(vAttList(14).TextString, " ")
                    
                    For i = 0 To UBound(vItem)
                        SelectAttach.lbAttach.AddItem "ST LT"
                        iIndex = SelectAttach.lbAttach.ListCount - 1
                        
                        If InStr(vItem(i), ")") > 0 Then
                            vItem(i) = Replace(vItem(i), "(", "")
                            vMR = Split(vItem(i), ")")
                            
                            SelectAttach.lbAttach.List(iIndex, 1) = vMR(0)
                            SelectAttach.lbAttach.List(iIndex, 2) = vMR(1)
                        Else
                            SelectAttach.lbAttach.List(iIndex, 1) = vItem(i)
                            SelectAttach.lbAttach.List(iIndex, 2) = "x"
                        End If
                    Next i
                End If
                
                For j = 16 To 23
                    If Not vAttList(j).TextString = "" Then
                        vMR = Split(vAttList(j).TextString, "=")
                        strCO = vMR(0)
                        vItem = Split(UCase(vMR(1)), " ")
                    
                        For i = 0 To UBound(vItem)
                            SelectAttach.lbAttach.AddItem strCO
                            iIndex = SelectAttach.lbAttach.ListCount - 1
                                                
                            vItem(i) = Replace(vItem(i), "C", "")
                            vItem(i) = Replace(vItem(i), "O", "")
                            vItem(i) = Replace(vItem(i), "X", "")
                        
                            If InStr(vItem(i), ")") > 0 Then
                                vItem(i) = Replace(vItem(i), "(", "")
                                vMR = Split(vItem(i), ")")
                            
                                SelectAttach.lbAttach.List(iIndex, 1) = vMR(0)
                                SelectAttach.lbAttach.List(iIndex, 2) = vMR(1)
                            Else
                                SelectAttach.lbAttach.List(iIndex, 1) = vItem(i)
                                SelectAttach.lbAttach.List(iIndex, 2) = "x"
                            End If
                        Next i
                    End If
                Next j
                
                SelectAttach.lbAttach.Selected(SelectAttach.lbAttach.ListCount - 1) = True
                
                SelectAttach.show
                
                iIndex = SelectAttach.lbAttach.ListIndex
                
                If Not SelectAttach.lbAttach.List(iIndex, 2) = "x" Then
                    If Not SelectAttach.lbAttach.List(iIndex, 1) = "x" Then
                        vAttach = Split(SelectAttach.lbAttach.List(iIndex, 1), "-")
                        iEInch = CInt(vAttach(0)) * 12
                        If UBound(vAttach) > 0 Then iEInch = iEInch + CInt(vAttach(1))
                            
                        vAttach = Split(SelectAttach.lbAttach.List(iIndex, 2), "-")
                        iPInch = CInt(vAttach(0)) * 12
                        If UBound(vAttach) > 0 Then iPInch = iPInch + CInt(vAttach(1))
                        
                        tbMRCompany.Value = SelectAttach.lbAttach.List(iIndex, 0)
                        tbMR.Value = iPInch - iEInch
                    End If
                Else
                    tbMRCompany.Value = SelectAttach.lbAttach.List(iIndex, 0)
                    tbMR.Value = 0
                End If
                
            Unload SelectAttach
            
            
            
        Case "pole_attach"
            iAttach = 1
            'tbPrePole.Value = vAttList(0).TextString
            tbMRCompany.Value = vAttList(2).TextString
            
            If vAttList(4).TextString = "" Then
                tbMR.Value = 0
            Else
                vTemp = Split(vAttList(4).TextString, " ")
                If UBound(vTemp) = 0 Then
                    tbMR.Value = 0
                Else
                    vTemp(1) = Replace(vTemp(1), """", "")
                    Select Case Left(vAttList(4).TextString, 1)
                        Case "L"
                            tbMR.Value = 0 - CInt(vTemp(1))
                        Case "R"
                            tbMR.Value = CInt(vTemp(1))
                    End Select
                End If
            End If
    End Select
Bad_Block:
    If iSpan = 0 Then
        If iAttach = 0 Then
            strPrompt = "Select Cable Span & Bottom Attachment:"
        Else
            strPrompt = "Select Cable Span:"
        End If
    Else
        If iAttach = 0 Then
            strPrompt = "Select Bottom Attachment:"
        Else
            strPrompt = "Nothing Needed:"
        End If
    End If
    
    GoTo Get_Block
    
Exit_Sub:
    If strPrompt = "Select Bottom Attachment:" Then
        Dim iInt As Integer
        
        iInt = lbWorkspace.ListCount - 1
        
        tbMRCompany.Value = lbWorkspace.List(iInt, 1)
        tbMR.Value = "0"
    End If
    
    cbAddCLR.SetFocus
    Me.show
End Sub

Private Sub cbMRReport_Click()
    Me.Hide
        Load MRSheets
            MRSheets.show
        Unload MRSheets
    Me.show
End Sub

Private Sub cbPlaceCLR_Click()
    Dim str1, str2 As String
    Dim dScale As Double
    Dim xDiff, yDiff As Double
    Dim obrBR As AcadBlockReference
    Dim returnPnt, returnPnt2 As Variant
    Dim varItem, vTemp As Variant
    Dim layerObj As AcadLayer
    
    On Error Resume Next
    
    If lbMR.ListIndex < 0 Then Exit Sub
    If lbMR.List(lbMR.ListIndex, 0) = "" Then Exit Sub
    
    dScale = 0.75   'CInt(cbScale.Value) / 100 * 1.333333
    dRotate = 0
    
    Err = 0
    str1 = lbMR.List(lbMR.ListIndex, 0) & " " & lbMR.List(lbMR.ListIndex, 2) & " CLR"
    
    str2 = Replace(lbMR.List(lbMR.ListIndex, 5), "-", "'") & """"
    If Not lbMR.List(lbMR.ListIndex, 6) = "" Then
        str2 = "(" & str2 & ") " & Replace(lbMR.List(lbMR.ListIndex, 6), "-", "'") & """"
    End If
    
    Me.Hide
    
    If lbMR.List(lbMR.ListIndex, 2) = "DW" Then
        Call AddDriveways
        GoTo Place_Block:
    End If
    
    returnPnt = ThisDrawing.Utility.GetPoint(, "Place Clearance: ")
    dInsertPnt(0) = returnPnt(0)
    dInsertPnt(1) = returnPnt(1)
    dInsertPnt(2) = 0#
    
    returnPnt2 = ThisDrawing.Utility.GetPoint(, "Rotation Direction: ")
    If dInsertPnt(0) > returnPnt2(0) Then
        xDiff = dInsertPnt(0) - returnPnt2(0)
        yDiff = dInsertPnt(1) - returnPnt2(1)
    Else
        xDiff = returnPnt2(0) - dInsertPnt(0)
        yDiff = returnPnt2(1) - dInsertPnt(1)
    End If
    If xDiff = 0 Then
        dRotate = 1.570796327
    Else
        dRotate = Atn(yDiff / xDiff)
    End If
    
Place_Block:
    
    Set obrBR = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "clr", dScale, dScale, dScale, dRotate)
    obrBR.Layer = "Integrity Roads-Clearance"
    varItem = obrBR.GetAttributes
    varItem(0).TextString = str1
    varItem(1).TextString = str2
    obrBR.Update
    
    Me.show
End Sub

Private Sub cbPlaceNote_Click()
    If lbWorkspace.ListCount = 0 Then GoTo Exit_Sub
    
    Dim returnPoint As Variant
    Dim insertionPnt(0 To 2) As Double
    Dim dRevCloud(0 To 5) As Double
    Dim dNote(0 To 2) As Double
    Dim dScale As Double
    Dim dPosition As Double
    Dim objBlock As AcadBlockReference
    Dim layerObj As AcadLayer
    Dim vAttList, vELine, vPLine As Variant
    Dim iPI, iEI As Integer
    Dim iMR, iNote As Integer
    Dim strAtt0, strAtt1, strAtt2, strAtt3, strAtt4 As String
    Dim strLayer As String
    
    Dim vStr As Variant
    Dim str, str1, strCommand As String
    Dim lwpPnt(0 To 3) As Double
    Dim lineObj As AcadLWPolyline
    Dim n, counter As Integer
    
  'On Error Resume Next
    
    strAtt0 = tbPoleNumber.Value
    
    iMR = 0
    iNote = 0
    dPosition = 1#
    
    Me.Hide
    dScale = 1#     'CInt(cbScale.Value) / 100
    
    returnPoint = ThisDrawing.Utility.GetPoint(, "Select point:")
    n = 0
    For Each Item In returnPoint
        insertionPnt(n) = Item
        n = n + 1
    Next Item
    
    dRevCloud(0) = insertionPnt(0) - (4 * dScale)
    dRevCloud(1) = insertionPnt(1) + (20 * dScale)
    dRevCloud(2) = 0#
    
    str = "pole_attach_title"
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
    objBlock.Layer = "Integrity Attachments-Existing"
    vAttList = objBlock.GetAttributes
    vAttList(0).TextString = strAtt0
    
    objBlock.Update
    
    insertionPnt(0) = insertionPnt(0) + (78 * dScale)
    insertionPnt(1) = insertionPnt(1) - (4.5 * dScale)
    
    str = "pole_attach"
    For i = 0 To (lbWorkspace.ListCount - 1)
        dPosition = dPosition + 0.01
        strAtt1 = dPosition
        strAtt2 = lbWorkspace.List(i, 1)
        iPI = 0: iEI = 0
        
        If Not lbWorkspace.List(i, 2) = "" Then
            vELine = Split(lbWorkspace.List(i, 2), "-")
            
            If UBound(vELine) = 0 Then
                iEI = CInt(vELine(0)) * 12
            Else
                iEI = CInt(vELine(0)) * 12 + CInt(vELine(1))
            End If
        End If
        
        If Not lbWorkspace.List(i, 3) = "" Then
            vPLine = Split(lbWorkspace.List(i, 3), "-")
            
            If UBound(vPLine) = 0 Then
                iPI = CInt(vPLine(0)) * 12
            Else
                iPI = CInt(vPLine(0)) * 12 + CInt(vPLine(1))
            End If
        End If
        
        
        
        If iPI = 0 Then
            strAtt3 = vELine(0) & "'" & vELine(1) & """"
            strAtt4 = ""
            strLayer = "Integrity Attachments-Existing"
            GoTo Place_Attachment
        End If
        
        If lbWorkspace.List(i, 0) = "NEW" Then
            strAtt3 = vPLine(0) & "'" & vPLine(1) & """"
            strAtt4 = "NEW"
            strLayer = "Integrity Attachments-New"
            If iNote < 1 Then iNote = 1
            GoTo Place_Attachment
        End If
        
        If iEI = 0 Then
            strAtt3 = vPLine(0) & "'" & vPLine(1) & """"
            strAtt4 = "ATTACH"
            strLayer = "Integrity Attachments-MR"
            iNote = 2
            GoTo Place_Attachment
        End If
        
        If cbCalloutType.Value = "(Existing)Proposed" Then
            strAtt3 = "(" & vELine(0) & "'" & vELine(1) & """)"
            strAtt4 = vPLine(0) & "'" & vPLine(1) & """"
            strLayer = "Integrity Attachments-MR"
            iNote = 2
            GoTo Place_Attachment
        End If
        
        strAtt3 = vPLine(0) & "'" & vPLine(1) & """"
        strLayer = "Integrity Attachments-MR"
        iNote = 2
        
        iMR = iPI - iEI
        Select Case iMR
            Case Is < 0
                strAtt4 = "LOWER " & Abs(iMR) & """"
            Case Is = 0
                strAtt4 = "TRANSFER"
            Case Is > 0
                strAtt4 = "RAISE " & Abs(iMR) & """"
        End Select
        
Place_Attachment:
        'If vAttList(2).TextString = "ATT" Or vAttList(2).TextString = "BST" Then
            'If iNote > 0 Then iNote = 2
        'End If
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        objBlock.Layer = strLayer
        
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = strAtt0
        vAttList(1).TextString = strAtt1
        vAttList(2).TextString = strAtt2
        vAttList(3).TextString = strAtt3
        vAttList(4).TextString = strAtt4
        objBlock.Update
        
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    Next i
    
    vAttList = objPole.GetAttributes
        
    Select Case vAttList(2).TextString
        Case "ATT", "BST"
            If iNote = 1 Then iNote = 3
    End Select
               
    Select Case iNote
        Case 0
            Set layerObj = ThisDrawing.Layers.Add("Integrity Attachments-Existing")
            ThisDrawing.ActiveLayer = layerObj
            dRevCloud(3) = insertionPnt(0) + (30 * dScale)
        Case 1
            Set layerObj = ThisDrawing.Layers.Add("Integrity Attachments-New")
            ThisDrawing.ActiveLayer = layerObj
            dRevCloud(3) = insertionPnt(0) + (58 * dScale)
        Case 2
            Set layerObj = ThisDrawing.Layers.Add("Integrity Attachments-MR")
            ThisDrawing.ActiveLayer = layerObj
            dRevCloud(3) = insertionPnt(0) + (84 * dScale)
        Case 3
            Set layerObj = ThisDrawing.Layers.Add("Integrity Attachments-MR")
            ThisDrawing.ActiveLayer = layerObj
            dRevCloud(3) = insertionPnt(0) + (58 * dScale)
    End Select
    dRevCloud(4) = insertionPnt(1) + (3 * dScale)
    dRevCloud(5) = 0#
    
    strCommand = "revcloud r " & dRevCloud(0) & "," & dRevCloud(1)
    strCommand = strCommand & " " & dRevCloud(3) & "," & dRevCloud(4) & " " & vbCr
    ThisDrawing.SendCommand strCommand
    
    Set layerObj = ThisDrawing.Layers.Add("0")
    ThisDrawing.ActiveLayer = layerObj
    
    '<---------------------------------------------------------------------------------------------
Place_Info:
    
    If cbPoleDataOn.Value = False Then GoTo Place_Note
    
    str = "pole_info"
    strLayer = "Integrity Pole-Info"
    
    'counter = 4 + lbCompany.ListCount

    insertionPnt(0) = insertionPnt(0) - (74 * dScale)
    insertionPnt(1) = insertionPnt(1) - (12 * dScale)
    
    dPosition = 0#
    
    Dim vText, vLine As Variant
    Dim strText As String
    
    strText = Replace(tbPoleData.Value, vbLf, "")
    vText = Split(strText, vbCr)
    
    For w = 0 To UBound(vText)
        If vText(w) = "" Then GoTo Next_W
        vLine = Split(vText(w), vbTab)
    
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        objBlock.Layer = strLayer
        
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = strAtt0
        vAttList(1).TextString = dPosition
        vAttList(2).TextString = vLine(0)
        If UBound(vLine) = 1 Then vAttList(3).TextString = vLine(1)
        objBlock.Update
    
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
        dPosition = dPosition + 1
Next_W:
    Next w
    
    lwpPnt(0) = insertionPnt(0) - (4 * dScale)
    lwpPnt(1) = insertionPnt(1) + (7 * dScale)
    lwpPnt(2) = lwpPnt(0) + (100 * dScale)
    lwpPnt(3) = lwpPnt(1)
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpPnt)
    lineObj.Layer = strLayer
    lineObj.Update
    
Place_Note:
    
    If cbPWRNoteOn.Value = False Then GoTo Exit_Sub
    If cbNote.Value = "" Then GoTo Exit_Sub
    
    Select Case cbNote.Value
        Case "Power to replace with Taller Pole"
            str = "Notes-TallerPole5"
        Case "Power to replace Defective Pole"
            str = "Notes-DefectivePole"
        Case Else
            str = ""
    End Select
    
    If str = "" Then GoTo Exit_Sub
    
    dNote(0) = (dRevCloud(0) + dRevCloud(3)) / 2
    dNote(1) = dRevCloud(1) + 10
    dNote(2) = 0#
    
    'vLine = dNote
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dNote, str, dScale, dScale, dScale, 0#)
    objBlock.Layer = strLayer
    objBlock.Update
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbSaveData_Click()
    Call SaveAttachments
    iSave = 0
    
    If cbGetAttach.Enabled = False Then Me.Hide
End Sub

Private Sub cbSort_Click()
    Call SortList
End Sub

Private Sub cbUP_Click()
    Dim str0, str1, str2 As String
    Dim str3, str4, str5 As String
    Dim i, i2 As Integer
    
    If lbWorkspace.ListIndex < 1 Then Exit Sub
    i = lbWorkspace.ListIndex
    i2 = i - 1
    
    str0 = lbWorkspace.List(i, 0)
    str1 = lbWorkspace.List(i, 1)
    str2 = lbWorkspace.List(i, 2)
    str3 = lbWorkspace.List(i, 3)
    str4 = lbWorkspace.List(i, 4)
    str5 = lbWorkspace.List(i, 5)
    
    lbWorkspace.List(i, 0) = lbWorkspace.List(i2, 0)
    lbWorkspace.List(i, 1) = lbWorkspace.List(i2, 1)
    lbWorkspace.List(i, 2) = lbWorkspace.List(i2, 2)
    lbWorkspace.List(i, 3) = lbWorkspace.List(i2, 3)
    lbWorkspace.List(i, 4) = lbWorkspace.List(i2, 4)
    lbWorkspace.List(i, 5) = lbWorkspace.List(i2, 5)
    
    lbWorkspace.List(i2, 0) = str0
    lbWorkspace.List(i2, 1) = str1
    lbWorkspace.List(i2, 2) = str2
    lbWorkspace.List(i2, 3) = str3
    lbWorkspace.List(i2, 4) = str4
    lbWorkspace.List(i2, 5) = str5
    
    lbWorkspace.ListIndex = i2
End Sub

Private Sub cbUpdate_Click()
    If iSave = 1 Then
        result = MsgBox("Save changes to Pole Attachments?", vbYesNo, "Save Changes")
        If result = vbYes Then
            Call SaveAttachments
            iSave = 0
        End If
    End If
    
    Me.Hide
End Sub

Private Sub cbUpdateAttach_Click()
    Dim objSS3 As AcadSelectionSet
    Dim entBlock As AcadEntity
    Dim obrBlock As AcadBlockReference
    Dim obrAttach2 As AcadBlockReference
    Dim attList As Variant
    Dim vArray As Variant
    Dim str, str1 As String
    Dim strCompany, strExist As String
    Dim strProposed, strNote As String
    Dim strLayer  As String
    Dim dTest, dScale As Double
    Dim dCoords(0 To 2) As Double
    Dim insertionPnt(0 To 2) As Double
    Dim iEI, iPI, iDiff As Integer
    
    Me.Hide
    
    dTest = 2#
    dCoords(0) = 0#
    dCoords(1) = 0#
    
  On Error Resume Next
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS3.SelectOnScreen
    For Each entBlock In objSS3
        If TypeOf entBlock Is AcadBlockReference Then
            Set obrBlock = entBlock
            If obrBlock.Name = "pole_attach" Then
                Select Case dCoords(1)
                    Case Is = 0, Is < obrBlock.InsertionPoint(1)
                        dCoords(0) = obrBlock.InsertionPoint(0)
                        dCoords(1) = obrBlock.InsertionPoint(1)
                        dCoords(2) = obrBlock.InsertionPoint(2)
                        
                        dScale = obrBlock.XScaleFactor
                End Select
                
                'entBlock.Delete
            End If
        End If
    Next entBlock
    
'    objSS3.Clear
    objSS3.Erase    '<--------------------------------------------------------------------------------------------------------------
    objSS3.Delete
    
    Me.Hide
    
    insertionPnt(0) = dCoords(0)
    insertionPnt(1) = dCoords(1)
    insertionPnt(2) = dCoords(2)
    
    str = "pole_attach"
    For i = 0 To (lbWorkspace.ListCount - 1)
        strCompany = lbWorkspace.List(i, 1)
        strExist = lbWorkspace.List(i, 2)
        strProposed = lbWorkspace.List(i, 3)
        
        If strExist = "" Then
            iEI = 0
        Else
            vArray = Split(strExist, "-")
            iEI = CInt(vArray(0)) * 12 + CInt(vArray(1))
        End If
        
        If strProposed = "" Then
            iPI = 0
        Else
            vArray = Split(strProposed, "-")
            iPI = CInt(vArray(0)) * 12 + CInt(vArray(1))
        End If
        
        iDiff = iPI - iEI
        
        strExist = Replace(strExist, "-", "'")
        strExist = strExist & """"
        
        strProposed = Replace(strProposed, "-", "'")
        strProposed = strProposed & """"
        
        Select Case iDiff
            Case Is = 0
                strNote = "TRANSFER"
                str1 = strProposed
                strLayer = "Integrity Attachments-MR"
            Case Is > 0
                If iEI > 0 Then
                    strNote = "RAISE " & Abs(iDiff) & """"
                    str1 = strProposed
                    strLayer = "Integrity Attachments-MR"
                Else
                    strNote = "NEW"
                    str1 = strProposed
                    strLayer = "Integrity Attachments-New"
                End If
            Case Is < 0
                If iPI > 0 Then
                    strNote = "LOWER " & Abs(iDiff) & """"
                    str1 = strProposed
                    strLayer = "Integrity Attachments-MR"
                Else
                    strNote = ""
                    str1 = strExist
                    strLayer = "Integrity Attachments-Existing"
                End If
        End Select
        
        'MsgBox tbUnitedPole.Value & vbCr & dPosition + (i / 100) & vbCr & strCompany & vbCr & str1 & vbCr & strNote
        
        Set obrAttach2 = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        obrAttach2.Layer = strLayer
        
        attList = obrAttach2.GetAttributes
        attList(0).TextString = tbUnitedPole.Value
        attList(1).TextString = dPosition + (i / 100)
        attList(2).TextString = strCompany
        attList(3).TextString = str1
        attList(4).TextString = strNote
        obrAttach2.Update
        
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    Next i
    
    cbUpdateAttach.Enabled = False
    
    lbWorkspace.Clear
    
    If Not Err = 0 Then MsgBox "Error: " & Err.Number & vbCr & Err.Description
    
    Me.show
End Sub

Private Sub cbUpdateCLR_Click()
    If lbMR.ListIndex < 0 Then Exit Sub
    
    Dim objEntity As AcadEntity
    Dim objCLR As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    Dim iIndex As Integer
    
    iIndex = lbMR.ListIndex
    
    Me.Hide
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Cable Span: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objCLR = objEntity
    vAttList = objCLR.GetAttributes
    
    If lbMR.List(iIndex, 6) = "" Then
        vAttList(0).TextString = lbMR.List(iIndex, 5)
    Else
        vAttList(0).TextString = "(" & lbMR.List(iIndex, 5) & ")" & lbMR.List(iIndex, 6)
    End If
    
    objCLR.Update
Exit_Sub:
    
End Sub

Private Sub cbUpdateLine_Click()
    Dim vAttach As Variant
    Dim strNote, strTemp As String
    Dim iCurrentIn, iLowIn, iUpperIn As Integer
    Dim iNewFt, iNewIn, iDiff As Integer
    Dim iTempFt, iTempIn As Integer
    
    If Not tbExisting.Value = "" Then
        vAttach = Split(tbExisting.Value, "-")
        If UBound(vAttach) = 0 Then tbExisting.Value = tbExisting.Value & "-0"
    End If
    
    If Not tbProposed.Value = "" Then
        vAttach = Split(tbProposed.Value, "-")
        If UBound(vAttach) = 0 Then tbExisting.Value = tbExisting.Value & "-0"
    End If
    
    lbWorkspace.List(iListIndex, 0) = tbType.Value
    lbWorkspace.List(iListIndex, 1) = tbCompany.Value
    lbWorkspace.List(iListIndex, 2) = tbExisting.Value
    lbWorkspace.List(iListIndex, 3) = tbProposed.Value
    lbWorkspace.List(iListIndex, 4) = tbNotes.Value
    
    CheckForViolations (iListIndex)
    
    tbType.Value = ""
    tbCompany.Value = ""
    tbExisting.Value = ""
    tbProposed.Value = ""
    tbNotes.Value = ""
    
    cbUpdateLine.Enabled = False
    
    If iListIndex = (lbWorkspace.ListCount - 1) Then Call CheckRoadCLR
    iSave = 1
End Sub

Private Sub L24_Click()
    If L24.Caption = "xx-xx" Then Exit Sub
    
    If cbUpdateLine.Enabled = True Then
        If tbType.Value = "NEW" Then tbProposed.Value = L24.Caption
        cbUpdateLine.SetFocus
        Exit Sub
    End If
    
    Dim iIndex As Integer
    
    lbWorkspace.AddItem "NEW"
    iIndex = lbWorkspace.ListCount - 1
    lbWorkspace.List(iIndex, 1) = "NEW 6M"
    lbWorkspace.List(iIndex, 3) = L24.Caption
    
    Call SortList
    iSave = 1
End Sub

Private Sub Label83_Click()
    Dim vFirstPnt, vSecondPnt As Variant
    Dim dX, dY, dZ As Double
    
    Me.Hide
    
    vFirstPnt = ThisDrawing.Utility.GetPoint(, "Select First Point: ")
    vSecondPnt = ThisDrawing.Utility.GetPoint(, "Select Second Point: ")
    
    dX = vSecondPnt(0) - vFirstPnt(0)
    dY = vSecondPnt(1) - vFirstPnt(1)
    
    dZ = Sqr(dX * dX + dY * dY)
    tbDist.Value = CInt(dZ)
    
    Me.show
End Sub

Private Sub lbMR_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo Exit_Sub
    
    Dim i As Integer
    Dim vTemp As Variant
    
    i = lbMR.ListIndex
    
    tbMRCompany.Value = lbMR.List(i, 0)
    tbMR.Value = lbMR.List(i, 1)
    cbType.Value = lbMR.List(i, 2)
    tbSpan.Value = lbMR.List(i, 3)
    tbDist.Value = lbMR.List(i, 4)
    
    If Not lbMR.List(i, 5) = "" Then
        vTemp = Split(lbMR.List(i, 5), "-")
        tbEF.Value = vTemp(0)
        If UBound(vTemp) > 0 Then tbEI.Value = vTemp(1)
    End If
    
    If Not lbMR.List(i, 6) = "" Then
        vTemp = Split(lbMR.List(i, 6), "-")
        tbPF.Value = vTemp(0)
        If UBound(vTemp) > 0 Then tbPI.Value = vTemp(1)
    End If
    cbAddCLR.Caption = "Update"
    
Exit_Sub:
End Sub

Private Sub lbMR_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            lbMR.RemoveItem lbMR.ListIndex
    End Select
End Sub

Private Sub lbWorkspace_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    iListIndex = lbWorkspace.ListIndex
    
    tbType.Value = lbWorkspace.List(iListIndex, 0)
    tbCompany.Value = lbWorkspace.List(iListIndex, 1)
    tbExisting.Value = lbWorkspace.List(iListIndex, 2)
    tbProposed.Value = lbWorkspace.List(iListIndex, 3)
    tbNotes.Value = lbWorkspace.List(iListIndex, 4)
    
    cbUpdateLine.Enabled = True
    tbType.Enabled = False
    tbCompany.Enabled = False
    tbExisting.Enabled = False
    tbNotes.Enabled = False
    
    tbProposed.SetFocus
    tbProposed.SelStart = 0
    tbProposed.SelLength = Len(tbProposed.Value)
End Sub

Private Sub Lcomm1_Click()
    If Lcomm1.Caption = "xx-xx" Then Exit Sub
    
    If cbUpdateLine.Enabled = True Then
        If tbType.Value = "NEW" Then tbProposed.Value = Lcomm1.Caption
        cbUpdateLine.SetFocus
        Exit Sub
    End If
    
    Dim iIndex As Integer
    
    lbWorkspace.AddItem "NEW"
    iIndex = lbWorkspace.ListCount - 1
    lbWorkspace.List(iIndex, 1) = "NEW 6M"
    lbWorkspace.List(iIndex, 3) = Lcomm1.Caption
    
    Call SortList
    iSave = 1
End Sub

Private Sub Lcomm2_Click()
    If Lcomm2.Caption = "xx-xx" Then Exit Sub
    
    If cbUpdateLine.Enabled = True Then
        If tbType.Value = "NEW" Then tbProposed.Value = Lcomm2.Caption
        cbUpdateLine.SetFocus
        Exit Sub
    End If
    
    Dim iIndex As Integer
    
    lbWorkspace.AddItem "NEW"
    iIndex = lbWorkspace.ListCount - 1
    lbWorkspace.List(iIndex, 1) = "NEW 6M"
    lbWorkspace.List(iIndex, 3) = Lcomm2.Caption
    
    Call SortList
    iSave = 1
End Sub

Private Sub LMaxHeight_Click()
    If LMaxHeight.Caption = "xx-xx" Then Exit Sub
    
    If cbUpdateLine.Enabled = True Then
        If tbType.Value = "NEW" Then tbProposed.Value = LMaxHeight.Caption
        cbUpdateLine.SetFocus
        Exit Sub
    End If
    
    Dim iIndex As Integer
    
    lbWorkspace.AddItem "NEW"
    iIndex = lbWorkspace.ListCount - 1
    lbWorkspace.List(iIndex, 1) = "NEW 6M"
    lbWorkspace.List(iIndex, 3) = LMaxHeight.Caption
    
    Call SortList
    iSave = 1
End Sub

Private Sub lUnlock_Click()
    tbType.Enabled = True
    tbCompany.Enabled = True
    tbExisting.Enabled = True
    tbNotes.Enabled = True

End Sub

Private Sub tbBC_Change()
    If tbBC.Value = "" Then Exit Sub
    
    For i = 0 To lbWorkspace.ListCount - 1
        'If lbWorkspace.List(i, 0) = "COMM" Then
            CheckForViolations (i)
        'End If
    Next i
End Sub

Private Sub tbBSC_Change()
    If tbBSC.Value = "" Then Exit Sub
    
    For i = 0 To lbWorkspace.ListCount - 1
        'If lbWorkspace.List(i, 0) = "COMM" Then
            CheckForViolations (i)
        'End If
    Next i
End Sub

Private Sub tbCompany_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbExisting.Enabled = True
End Sub

Private Sub tbLL_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tbLL.SelStart = 0
    tbLL.SelLength = Len(tbLL.Value)
End Sub

Private Sub tbProposed_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbNotes.Enabled = True
End Sub

Private Sub tbSP_Change()
    If tbSP.Value = "" Then Exit Sub
    
    For i = 0 To lbWorkspace.ListCount - 1
        'If lbWorkspace.List(i, 0) = "PWR" Then
            CheckForViolations (i)
        'End If
    Next i
End Sub

Private Sub tbSS_Change()
    If tbSS.Value = "" Then Exit Sub
    
    For i = 0 To lbWorkspace.ListCount - 1
        'If lbWorkspace.List(i, 0) = "PWR" Then
            CheckForViolations (i)
        'End If
    Next i
End Sub

Private Sub tbST_Change()
    If tbST.Value = "" Then Exit Sub
    
    For i = 0 To lbWorkspace.ListCount - 1
        'If lbWorkspace.List(i, 0) = "PWR" Then
            CheckForViolations (i)
        'End If
    Next i
End Sub

Private Sub UserForm_Initialize()
    cbUP.Caption = Chr(225)
    cbDOWN.Caption = Chr(226)
    
    cbNote.AddItem ""
    cbNote.AddItem "Pwr replace with 5 Taller Pole"
    cbNote.AddItem "Pwr replace with 10 Taller Pole"
    cbNote.AddItem "Pwr replace Defective Pole"
    cbNote.AddItem "Pwr Raise Attachment"
    cbNote.AddItem "Extra Seperation"
    cbNote.AddItem "Use Existing Holes"
    cbNote.AddItem "Pole Permited Previously"
    cbNote.AddItem "OCALC Required"
    cbNote.AddItem "Replace Note with Other"
    'cbNote.AddItem "Power to replace with Thicker Pole"
    'cbNote.AddItem "Power to replace with Like Pole"
    'cbNote.AddItem "pole permitted"
    'cbNote.AddItem "replace pole"
    'cbNote.AddItem "assumes completion"
    'cbNote.AddItem "same bolt"
    'cbNote.AddItem "future same bolt"
    'cbNote.AddItem "dist stops"
    'cbNote.AddItem "fdh ground"
    'cbNote.AddItem "fdh ground wire"
    'cbNote.AddItem "additional seperation"
    'cbNote.AddItem "cant lower replace pole"
    'cbNote.AddItem "cant lower determine"
    'cbNote.Value = ""
    
    cbType.AddItem ""
    cbType.AddItem "RD"
    cbType.AddItem "DW"
    cbType.AddItem "MS"
    
    cbAlign.AddItem "Align"
    cbAlign.AddItem "Leader"
    cbAlign.Value = "Align"
    
    cbCalloutType.AddItem "(Existing)Proposed"
    cbCalloutType.AddItem "Proposed  MR-Note"
    cbCalloutType.Value = "Proposed  MR-Note"
    
    tbType.AddItem "PWR"
    tbType.AddItem "NEW"
    tbType.AddItem "COMM"
    
    lbWorkspace.Clear
    lbWorkspace.ColumnCount = 6
    lbWorkspace.ColumnWidths = "48;72;48;48;96;72"
    
    lbMR.Clear
    lbMR.ColumnCount = 7
    lbMR.ColumnWidths = "48;34;48;48;48;48;48"
    
    If Not lbWorkspace.ListCount = 0 Then
        For i = 0 To lbWorkspace.ListCount - 1
            CheckForViolations (i)
        Next i
    End If
    
    iListIndex = 0
    iSave = 0
    iChanged = 0
End Sub

Private Sub CheckForViolations(iIndex As Integer)
    Dim vAttach, vTemp, vLine As Variant
    Dim iCurrentFt, iCurrentIn As Integer
    Dim iLowFt, iLowIn As Integer
    Dim iUpperFt, iUpperIn As Integer
    Dim iNewFt, iNewIn As Integer
    Dim iDiff As Integer
    Dim strType, strCompany, strNote As String
    
    strType = lbWorkspace.List(iIndex, 0)
    strCompany = lbWorkspace.List(iIndex, 1)
    
    If Not lbWorkspace.List(iIndex, 3) = "" Then
        vAttach = Split(lbWorkspace.List(iIndex, 3), "-")
    Else
        vAttach = Split(lbWorkspace.List(iIndex, 2), "-")
    End If
            
    iCurrentIn = CInt(vAttach(0)) * 12 + CInt(vAttach(1))
            
    Select Case strType
        Case "PWR"
            Select Case strCompany
                Case "LOW POWER", "NEUTRAL", "ST LT CIR"
                    iLowIn = iCurrentIn - CInt(tbSP.Value)
                Case "TRANSFORMER"
                    iLowIn = iCurrentIn - CInt(tbST.Value)
                Case "ST LT"
                    iLowIn = iCurrentIn - CInt(tbSS.Value)
            End Select
            
            iNewFt = CInt(iLowIn / 12 - 0.5)
            iNewIn = iLowIn - (iNewFt * 12)
            If iNewIn > 11 Then
                iNewIn = iNewIn - 12
                iNewFt = iNewFt + 1
            End If
            
            strNote = "MAX: " & iNewFt & "-" & iNewIn
            lbWorkspace.List(iIndex, 4) = strNote
            
            For c = iIndex To lbWorkspace.ListCount - 1
                If lbWorkspace.List(c, 0) = "NEW" Then
                    If Not lbWorkspace.List(c, 3) = "" Then
                        vAttach = Split(lbWorkspace.List(c, 3), "-")
                    Else
                        vAttach = Split(lbWorkspace.List(c, 2), "-")
                    End If
                    iUpperIn = CInt(vAttach(0)) * 12 + CInt(vAttach(1))

                    If iLowIn - iUpperIn < 0 Then
                        lbWorkspace.List(c, 5) = "PWR"
                    Else
                        lbWorkspace.List(c, 5) = ""
                    End If
                End If
            Next c
        Case "NEW"
            lbWorkspace.List(iIndex, 5) = ""
            For g = 0 To iIndex + 1
                If g = iIndex Then GoTo Next_G
                If g > lbWorkspace.ListCount - 1 Then GoTo Next_G
                
                lbWorkspace.List(g, 3) = Replace(lbWorkspace.List(g, 3), " ", "")
                If Not lbWorkspace.List(g, 3) = "" Then
                    vAttach = Split(lbWorkspace.List(g, 3), "-")
                Else
                    vAttach = Split(lbWorkspace.List(g, 2), "-")
                End If
                'vAttach = Split(lbWorkspace.List(g, 4), " ")
                'vTemp = Split(vAttach(1), "-")
                'iUpperIn = CInt(vTemp(0)) * 12
                'If UBound(vTemp) > 0 Then iUpperIn = iUpperIn + CInt(vTemp(1))
                iUpperIn = CInt(vAttach(0)) * 12
                If UBound(vAttach) > 0 Then iUpperIn = iUpperIn + CInt(vAttach(1))
                
                If lbWorkspace.List(g, 0) = "PWR" Then
                    If iUpperIn - iCurrentIn < 0 Then
                        lbWorkspace.List(iIndex, 5) = "PWR"
                    End If
                Else
                    If iUpperIn - iCurrentIn > 0 Then
                        lbWorkspace.List(iIndex, 5) = lbWorkspace.List(iIndex, 5) & "  COMM"
                    End If
                End If
Next_G:
            Next g
        Case "COMM"
            If Not lbWorkspace.List(iIndex - 1, 3) = "" Then
                vAttach = Split(lbWorkspace.List(iIndex - 1, 3), "-")
            Else
                vAttach = Split(lbWorkspace.List(iIndex - 1, 2), "-")
            End If
            
            '<---------------------------------------------------------------Upper Attachment
            iUpperIn = CInt(vAttach(0)) * 12 + CInt(vAttach(1))
            iDiff = iUpperIn - iCurrentIn
            
            strNote = "SEP: " & Abs(iDiff)
            Select Case lbWorkspace.List(iIndex - 1, 1)
                Case tbCompany.Value
                    If iDiff < CInt(tbBSC.Value) Then
                        lbWorkspace.List(iIndex, 5) = "SEPERATION"
                    Else
                        lbWorkspace.List(iIndex, 5) = ""
                    End If
                Case "UTC NEW"
                    iTempIn = iCurrentIn + CInt(tbBC.Value)
                    iTempFt = CInt(iTempIn / 12 - 0.5)
                    iTempIn = iTempIn - (iTempFt * 12)
                    If iTempIn > 11 Then
                        iTempFt = iTempFt + 1
                        iTempIn = iTempIn - 12
                    End If

                    If iDiff < CInt(tbBC.Value) Then
                        lbWorkspace.List(iIndex, 5) = "SEPERATION"
                    Else
                        lbWorkspace.List(iIndex, 5) = ""
                    End If

                    strNote = "MIN: " & iTempFt & "-" & iTempIn
                Case Else
                    If iDiff < CInt(tbBC.Value) Then
                        lbWorkspace.List(iIndex, 5) = "SEPERATION"
                    Else
                        lbWorkspace.List(iIndex, 5) = ""
                    End If
            End Select
            
            lbWorkspace.List(iIndex, 5) = strNote
            
            '<---------------------------------------------------------------Lower Attachment
            If iIndex = lbWorkspace.ListCount - 1 Then GoTo Skip_Lower
            If Not lbWorkspace.List(iIndex + 1, 3) = "" Then
                vAttach = Split(lbWorkspace.List(iIndex + 1, 3), "-")
            Else
                vAttach = Split(lbWorkspace.List(iIndex + 1, 2), "-")
            End If
            
            iLowIn = CInt(vAttach(0)) * 12 + CInt(vAttach(1))
            iDiff = iCurrentIn - iLowIn
            
            strNote = "SEP: " & Abs(iDiff)
            
            If tbCompany.Value = lbWorkspace.List(iIndex + 1, 1) Then
                If iDiff < CInt(tbBSC.Value) Then
                    lbWorkspace.List(iIndex + 1, 5) = "SEPERATION"
                Else
                    lbWorkspace.List(iIndex + 1, 5) = ""
                End If
            Else
                If iDiff < CInt(tbBC.Value) Then
                    lbWorkspace.List(iIndex + 1, 5) = "SEPERATION"
                Else
                    lbWorkspace.List(iIndex + 1, 5) = ""
                End If
            End If
            lbWorkspace.List(iIndex + 1, 5) = strNote
            
Skip_Lower:
            'strTemp = lbWorkspace.List(iIndex - 1, 2) '= tbExisting.Value
    End Select
End Sub

Public Sub SortList()
    Dim strArrayList(), strArraySorted(), strData() As String
    Dim strListItem() As String     '<---------------------------------------------Sort
    Dim attArray, attItem, vTemp As Variant
    Dim vNotes As Variant
    Dim str1, str2 As String
    Dim i, iDWGNum, test1, place1, temp1 As Integer
    Dim iOCALC As Integer
    Dim strTemp As String
    Dim iFeet, iInch As Integer
    Dim tempFt, tempIn As Integer
    Dim iMaxIn, iMinIn, iTInch As Integer
    Dim iMaxFt, iMinFt As Integer
    Dim iCD As Integer
    
  'On Error Resume Next
    
    test1 = 0
    iOCALC = 0
    iCD = 0
    'ReDim strArrayList(0 To lbAttachments.ListCount)
    If lbWorkspace.ListCount < 1 Then Exit Sub
    ReDim strData(0 To lbWorkspace.ListCount - 1)
    ReDim strListItem(0 To lbWorkspace.ListCount - 1)
    ReDim strArraySorted(0 To lbWorkspace.ListCount - 1)
    
    For i = 0 To UBound(strListItem)
        strListItem(i) = i
    Next i
    
    For i = 0 To UBound(strData)
        If Not lbWorkspace.List(i, 3) = "" Then
            vTemp = Split(lbWorkspace.List(i, 3), "-")
            iInch = CInt(vTemp(0)) * 12 + CInt(vTemp(1))
        Else
            vTemp = Split(lbWorkspace.List(i, 2), "-")
            iInch = CInt(vTemp(0)) * 12 + CInt(vTemp(1))
        End If
        
        strData(i) = iInch
    Next i
    
    test1 = UBound(strData) - 1

    For i = UBound(strData) To (LBound(strData) + 1) Step -1
        For j = LBound(strData) To (i - 1)
            If CInt(strData(j)) < CInt(strData(j + 1)) Then
                temp1 = strData(j + 1)
                strData(j + 1) = strData(j)
                strData(j) = temp1
                
                strTemp = strListItem(j + 1)
                strListItem(j + 1) = strListItem(j)
                strListItem(j) = strTemp
            End If
        Next j
    Next i
    
    For i = LBound(strListItem) To UBound(strListItem)
        strArraySorted(i) = lbWorkspace.List(strListItem(i), 0) & vbTab & lbWorkspace.List(strListItem(i), 1) & vbTab
        strArraySorted(i) = strArraySorted(i) & lbWorkspace.List(strListItem(i), 2) & vbTab & lbWorkspace.List(strListItem(i), 3) & vbTab
        strArraySorted(i) = strArraySorted(i) & lbWorkspace.List(strListItem(i), 4) & vbTab & lbWorkspace.List(strListItem(i), 5) & vbTab
    Next i
    
    iMaxIn = 840
    iMinIn = 0
    
    lbWorkspace.Clear
    For i = LBound(strListItem) To UBound(strListItem)
        lbWorkspace.AddItem
        attItem = Split(strArraySorted(i), vbTab)
        lbWorkspace.List(i, 0) = attItem(0)
        lbWorkspace.List(i, 1) = attItem(1)
        lbWorkspace.List(i, 2) = attItem(2)
        lbWorkspace.List(i, 3) = attItem(3)
        lbWorkspace.List(i, 4) = attItem(4)
        lbWorkspace.List(i, 5) = attItem(5)
        
        If attItem(3) = "" Then
            vTemp = Split(attItem(2), "-")
            tempFt = CInt(vTemp(0))
            tempIn = CInt(vTemp(1))
        Else
            vTemp = Split(attItem(3), "-")
            tempFt = CInt(vTemp(0))
            tempIn = CInt(vTemp(1))
        End If
        
        Select Case attItem(1)
            Case "NEUTRAL"
                If tempIn < 4 Then
                    tempIn = tempIn + 12
                    tempFt = tempFt - 1
                End If
                tempFt = tempFt - 3
                tempIn = tempIn - 4
                iTInch = tempFt * 12 + tempIn
            
                If iTInch < iMaxIn Then
                    iMaxIn = iTInch
                End If
                
                lbWorkspace.List(i, 5) = "MAX: " & tempFt & "-" & tempIn
            Case "LOW POWER", "ST LT CIRCUIT"
                If tempIn < 4 Then
                    tempIn = tempIn + 12
                    tempFt = tempFt - 1
                End If
                tempFt = tempFt - 3
                tempIn = tempIn - 4
                iTInch = tempFt * 12 + tempIn
            
                If iTInch < iMaxIn Then
                    iMaxIn = iTInch
                End If
                
                lbWorkspace.List(i, 5) = "MAX: " & tempFt & "-" & tempIn
            Case "TRANSFORMER"
                If tempIn < 6 Then
                    tempIn = tempIn + 12
                    tempFt = tempFt - 1
                End If
                tempFt = tempFt - 2
                tempIn = tempIn - 6
                iTInch = tempFt * 12 + tempIn
            
                If iTInch < iMaxIn Then
                    iMaxIn = iTInch
                End If
                
                lbWorkspace.List(i, 5) = "MAX: " & tempFt & "-" & tempIn
            Case "ST LT"
                tempFt = tempFt - 1
                iTInch = tempFt * 12 + tempIn
            
                If iTInch < iMaxIn Then
                    iMaxIn = iTInch
                End If
                
                lbWorkspace.List(i, 5) = "MAX: " & tempFt & "-" & tempIn
            Case "NEW 6M"
                iOCALC = iOCALC + 1
            Case Else
                iTInch = tempFt * 12 + tempIn
                If iTInch > iMinIn Then iMinIn = iTInch
                
                iOCALC = iOCALC + 1
                If InStr(lbWorkspace.List(i, 1), "C-WIRE") > 0 Then
                    iOCALC = iOCALC - 1
                    iCD = 1
                End If
                If InStr(lbWorkspace.List(i, 1), "DROP") > 0 Then
                    iOCALC = iOCALC - 1
                    iCD = 1
                End If
        End Select
    Next i
    
    If iMaxIn < 288 Then L24.Enabled = False
    If iMaxIn < (iMinIn + 12) Then Lcomm1.Enabled = False
    If iMaxIn < (iMinIn + 24) Then Lcomm2.Enabled = False
    
    iMaxFt = Int(iMaxIn / 12)
    iMaxIn = iMaxIn - (iMaxFt * 12)
    
    LMaxHeight.Caption = iMaxFt & "-" & iMaxIn
    
    iMinFt = Int(iMinIn / 12)
    iMinIn = iMinIn - (iMinFt * 12)
    iMinFt = iMinFt + 1
    
    Lcomm1.Caption = iMinFt & "-" & iMinIn
    
    iMinFt = iMinFt + 1
    
    Lcomm2.Caption = iMinFt & "-" & iMinIn
    
    iOCALC = iOCALC + iCD
    
    If iOCALC > 3 Then
        vNotes = Split(tbPoleData.Value, vbCr)
        vTemp = Split(Replace(vNotes(1), vbLf, ""), vbTab)
        
        Select Case vTemp(0)
            Case "MTE", "MTEMC", "DRE", "DREMC", "SWPS"
                If InStr(tbNoteList.Value, "OCALC") < 1 Then
                    If tbNoteList.Value = "" Then
                        tbNoteList.Value = "OCALC"
                    Else
                        tbNoteList.Value = tbNoteList.Value & vbCr & "OCALC"
                    End If
                End If
        End Select
    End If
End Sub

Private Sub CheckRoadCLR()
    Dim vLine, vHeight, vTemp As Variant
    Dim strLine As String
    Dim strPoleAtt, strTestAtt As String
    Dim dSpan, dDist, dMR As Double
    Dim iPoleFeet, iPoleInch, iPoleMR As Integer
    Dim iCLRFeet, iCLRInch, iCLRMR As Integer
    Dim iTestFeet, iTestInch, iTestMR As Integer
    Dim iTest As Integer
    Dim lColor As Long
    Dim strCompany, strExist As String
    Dim strNote As String
    Dim vList As Variant
    Dim iEI, iPI, iDiff As Integer
    
    On Error Resume Next
    
    strCompany = lbWorkspace.List(lbWorkspace.ListCount - 1, 1)
    strExist = lbWorkspace.List(lbWorkspace.ListCount - 1, 2)
    strProposed = lbWorkspace.List(lbWorkspace.ListCount - 1, 3)
    
    If strExist = "" Then Exit Sub
    If strProposed = "" Then
        If lbMR.ListCount > 0 Then
            For i = 0 To lbMR.ListCount - 1
                lbMR.List(i, 6) = ""
            Next i
        End If
        Exit Sub
    End If
    
    vList = Split(strExist, "-")
    iEI = CInt(vList(0)) * 12 + CInt(vList(1))
    
    vList = Split(strProposed, "-")
    iPI = CInt(vList(0)) * 12 + CInt(vList(1))
    
    iPoleMR = iPI - iEI
    
    For i = 0 To lbMR.ListCount - 1
        iTestMR = lbMR.List(i, 1)
        
        vLine = Split(lbMR.List(i, 5), "-")
        iCLRFeet = CInt(vLine(0))
        iCLRInch = iCLRFeet * 12 + CInt(vLine(1))
    
        iTest = iTestMR - iPoleMR
    
        dSpan = CDbl(lbMR.List(i, 3))
        dDist = CDbl(lbMR.List(i, 4))
        dMR = dDist / dSpan
    
        iCLRInch = iCLRInch + iPoleMR + Round(iTest * dMR)
        iCLRFeet = 0
    
        Do While iCLRInch > 12
            iCLRInch = iCLRInch - 12
            iCLRFeet = iCLRFeet + 1
        Loop
    
        Do While iCLRInch < 0
            iCLRInch = iCLRInch + 12
            iCLRFeet = iCLRFeet - 1
        Loop
    
        lbMR.List(i, 6) = iCLRFeet & "-" & iCLRInch
    Next i
End Sub

Private Sub PlaceCLR()
    Dim str1, str2 As String
    Dim dScale As Double
    Dim xDiff, yDiff As Double
    Dim obrBR As AcadBlockReference
    Dim returnPnt, returnPnt2 As Variant
    Dim varItem, vTemp As Variant
    Dim layerObj As AcadLayer
    
    dScale = CInt(cbScale.Value) / 100 * 1.333333
    dRotate = 0
    
    vTemp = Split(tbPreAttach.Value, vbTab)
    
    
    str1 = vTemp(0) & " " & cbPreType.Value & " CLR"
    str2 = tbPreEF.Value & "'"
    If Not tbPreEI.Value = "" Then str2 = str2 & tbPreEI.Value & """"
    If Not tbPrePF.Value = "" Then
        str2 = "(" & str2 & ") " & tbPrePF.Value & "'"
        If Not tbPrePI.Value = "" Then str2 = str2 & tbPrePI.Value & """"
    End If
    
    Me.Hide
    
    If cbType.Value = "DW" Then
        Call AddDriveways
        GoTo Place_Block:
    End If
    
    returnPnt = ThisDrawing.Utility.GetPoint(, "Place Clearance: ")
    dInsertPnt(0) = returnPnt(0)
    dInsertPnt(1) = returnPnt(1)
    dInsertPnt(2) = 0#
    
    returnPnt2 = ThisDrawing.Utility.GetPoint(, "Rotation Direction: ")
    If dInsertPnt(0) > returnPnt2(0) Then
        xDiff = dInsertPnt(0) - returnPnt2(0)
        yDiff = dInsertPnt(1) - returnPnt2(1)
    Else
        xDiff = returnPnt2(0) - dInsertPnt(0)
        yDiff = returnPnt2(1) - dInsertPnt(1)
    End If
    If xDiff = 0 Then
        dRotate = 1.570796327
    Else
        dRotate = Atn(yDiff / xDiff)
    End If
    
Place_Block:
    
    Set obrBR = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "clr", dScale, dScale, dScale, dRotate)
    obrBR.Layer = "Integrity Roads-Clearance"
    varItem = obrBR.GetAttributes
    varItem(0).TextString = str1
    varItem(1).TextString = str2
    obrBR.Update
    
    Me.show
Exit_Sub:
End Sub

Private Sub AddDriveways()
    Dim lineObj As AcadLWPolyline
    Dim insertPnt(0 To 2), leaderPnt(0 To 3) As Double
    Dim dScale As Double
    Dim returnPnt, returnPnt2, returnPnt3 As Variant
    Dim vEOPL As Variant
    Dim xDist, yDist, zDist, dRatio, dRatio2 As Double
    Dim dOffset As Double
    Dim iDirection As Integer 'L=-1 R=1
    
    'Me.hide
   On Error Resume Next
   
    dScale = 0.75
    'If Not Err = 0 Then dScale = 1
    
    returnPnt2 = ThisDrawing.Utility.GetPoint(, vbCr & "Start line: ")
    returnPnt = ThisDrawing.Utility.GetPoint(returnPnt2, vbCr & "End Line: ")
    'returnPnt3 = returnPnt
    
    xDist = returnPnt(0) - returnPnt2(0)
    yDist = returnPnt(1) - returnPnt2(1)
    zDist = Sqr((xDist * xDist) + (yDist * yDist))
    dRatio = (133 * dScale) / zDist
    dRatio2 = (100 * dScale) / zDist
    If xDist = 0 Then
        dRotate = 1.570796327
    Else
        dRotate = Atn(yDist / xDist)
    End If
'MsgBox dRotate
    
    leaderPnt(0) = returnPnt(0) - dRatio * xDist
    leaderPnt(1) = returnPnt(1) - dRatio * yDist
    leaderPnt(2) = returnPnt(0)
    leaderPnt(3) = returnPnt(1)
    dInsertPnt(0) = returnPnt(0) - dRatio2 * xDist
    dInsertPnt(1) = returnPnt(1) - dRatio2 * yDist
    dInsertPnt(2) = 0#
    
    Select Case xDist
        Case Is < 0
            iDirection = -1
            'dRotate = dRotate * (-1)
        Case Else
            iDirection = 1
    End Select
    
        
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(leaderPnt)
    lineObj.Layer = "Integrity Roads-Clearance"
    
    vEOPL = lineObj.Offset(21.3333 * dScale * iDirection)
    lineObj.Update
End Sub

Private Sub RemoveEmptyRows()
    For i = lbMR.ListCount - 1 To 0 Step -1
        If lbMR.List(i, 0) = "" Then lbMR.RemoveItem i
    Next i
End Sub

Private Sub SaveAttachments()
    'If lbWorkspace.ListCount < 1 Then Exit Sub
    
    Dim strComm(8) As String
    Dim strComm1, strComm2, strComm3, strComm4 As String
    Dim strComm5, strComm6, strComm7, strComm8 As String
    Dim strLine, strExtra, strCO, strLetter As String
    Dim vAttList, vLine, vItem As Variant
    
    For i = 0 To 8
        strComm(i) = ""
    Next i
    
    strComm1 = "": strComm2 = "": strComm3 = "": strComm4 = ""
    strComm5 = "": strComm6 = "": strComm7 = "": strComm8 = ""
    
    vAttList = objPole.GetAttributes
    
    For i = 9 To 23
        vAttList(i).TextString = ""
    Next i
    
    For i = 0 To lbWorkspace.ListCount - 1
        Select Case lbWorkspace.List(i, 0)
            Case "PWR"
                Select Case lbWorkspace.List(i, 1)
                    Case "NEUTRAL"
                        strLine = lbWorkspace.List(i, 2)
                        If Not lbWorkspace.List(i, 3) = "" Then strLine = "(" & strLine & ")" & lbWorkspace.List(i, 3)
                        
                        If vAttList(9).TextString = "" Then
                            vAttList(9).TextString = strLine
                        Else
                            vAttList(9).TextString = vAttList(9).TextString & " " & strLine
                        End If
                    Case "TRANSFORMER"
                        strLine = lbWorkspace.List(i, 2)
                        If Not lbWorkspace.List(i, 3) = "" Then strLine = "(" & strLine & ")" & lbWorkspace.List(i, 3)
                        
                        If vAttList(10).TextString = "" Then
                            vAttList(10).TextString = strLine
                        Else
                            vAttList(10).TextString = vAttList(10).TextString & " " & strLine
                        End If
                    Case "LOW POWER"
                        strLine = lbWorkspace.List(i, 2)
                        If Not lbWorkspace.List(i, 3) = "" Then strLine = "(" & strLine & ")" & lbWorkspace.List(i, 3)
                        
                        If vAttList(11).TextString = "" Then
                            vAttList(11).TextString = strLine
                        Else
                            vAttList(11).TextString = vAttList(10).TextString & " " & strLine
                        End If
                    Case "ANTENNA"
                        strLine = lbWorkspace.List(i, 2)
                        If Not lbWorkspace.List(i, 3) = "" Then strLine = "(" & strLine & ")" & lbWorkspace.List(i, 3)
                        
                        If vAttList(12).TextString = "" Then
                            vAttList(12).TextString = strLine
                        Else
                            vAttList(12).TextString = vAttList(10).TextString & " " & strLine
                        End If
                    Case "ST LT CIRCUIT"
                        strLine = lbWorkspace.List(i, 2)
                        If Not lbWorkspace.List(i, 3) = "" Then strLine = "(" & strLine & ")" & lbWorkspace.List(i, 3)
                        
                        If vAttList(13).TextString = "" Then
                            vAttList(13).TextString = strLine
                        Else
                            vAttList(13).TextString = vAttList(10).TextString & " " & strLine
                        End If
                    Case "ST LT"
                        strLine = lbWorkspace.List(i, 2)
                        If Not lbWorkspace.List(i, 3) = "" Then strLine = "(" & strLine & ")" & lbWorkspace.List(i, 3)
                        
                        If vAttList(14).TextString = "" Then
                            vAttList(14).TextString = strLine
                        Else
                            vAttList(14).TextString = vAttList(10).TextString & " " & strLine
                        End If
                End Select
            Case "NEW"
                If vAttList(15).TextString = "" Then
                    vAttList(15).TextString = lbWorkspace.List(i, 3)
                Else
                    vAttList(15).TextString = vAttList(15).TextString & " " & lbWorkspace.List(i, 3)
                End If
                
                If lbWorkspace.List(i, 4) = "MTE TAG" Then vAttList(15).TextString = vAttList(15).TextString & "T"
                If lbWorkspace.List(i, 1) = "NEW OHG" Then vAttList(15).TextString = vAttList(15).TextString & "O"
                If lbWorkspace.List(i, 4) = "FUTURE" Then vAttList(15).TextString = vAttList(15).TextString & "F"
            Case "COMM"
                strCO = lbWorkspace.List(i, 1)
                If InStr(strCO, "C-WIRE") > 0 Then
                    strExtra = "c"
                    strCO = Replace(strCO, " C-WIRE", "")
                End If
                
                If InStr(strCO, "DROP") > 0 Then
                    strExtra = "d"
                    strCO = Replace(strCO, " DROP", "")
                End If
                
                If lbWorkspace.List(i, 1) = "EXTEND" Then
                    If strExtra = "" Then
                        strExtra = "e"
                    Else
                        strExtra = strExtra & "e"
                    End If
                End If
                
                If InStr(strCO, "OHG") > 0 Then
                    If strExtra = "" Then
                        strExtra = "o"
                    Else
                        strExtra = strExtra & "o"
                    End If
                    strCO = Replace(strCO, " OHG", "")
                End If
                
                If InStr(strCO, "LASH TO ") > 0 Then
                    If strExtra = "" Then
                        strExtra = "v"
                    Else
                        strExtra = strExtra & "v"
                    End If
                    strCO = Replace(strCO, "LASH TO ", "")
                End If
                
                If InStr(strCO, " TAP") > 0 Then
                    If strExtra = "" Then
                        strExtra = "p"
                    Else
                        strExtra = strExtra & "p"
                    End If
                    strCO = Replace(strCO, " TAP", "")
                End If
                
                If InStr(strCO, " SS") > 0 Then
                    If strExtra = "" Then
                        strExtra = "s"
                    Else
                        strExtra = strExtra & "s"
                    End If
                    strCO = Replace(strCO, " SS", "")
                End If
                
                If InStr(strCO, " TAG") > 0 Then
                    If strExtra = "" Then
                        strExtra = "t"
                    Else
                        strExtra = strExtra & "t"
                    End If
                    strCO = Replace(strCO, " TAG", "")
                End If
                
                If Not lbWorkspace.List(i, 2) = "" Then
                    strLine = lbWorkspace.List(i, 2)
                    
                    If Not lbWorkspace.List(i, 3) = "" Then
                        strLine = "(" & strLine & ")" & lbWorkspace.List(i, 3) & strExtra
                    Else
                        strLine = strLine & strExtra
                    End If
                    
                    strExtra = ""
                Else
                    strLine = lbWorkspace.List(i, 3) & strExtra & "x"
                    strExtra = ""
                End If
                
                For j = 1 To 8
                    vLine = Split(strComm(j), "=")
                    If strComm(j) = "" Then
                        strComm(j) = strCO & "=" & strLine
                        GoTo Found_strComm
                    End If
                    
                    If vLine(0) = strCO Then
                        strComm(j) = strComm(j) & " " & strLine
                        GoTo Found_strComm
                    End If
                Next j
                
                MsgBox "Comms full"
                Exit Sub
Found_strComm:
        End Select
    Next i
    
    For i = 1 To 8
        vAttList(15 + i).TextString = strComm(i)
    Next i
    
    If cbPoleDataOn.Value = False Then GoTo Add_Notes
    
    vAttList(2).TextString = ""
    vAttList(3).TextString = ""
    vAttList(4).TextString = ""
    vAttList(5).TextString = ""
    
    strLine = Replace(tbPoleData.Value, vbLf, "")
    vLine = Split(strLine, vbCr)
    
    vItem = Split(vLine(1), vbTab)
    strCO = vItem(0)
    vAttList(2).TextString = strCO
    vAttList(5).TextString = vItem(1)
    
    vItem = Split(vLine(2), vbTab)
    vAttList(3).TextString = vItem(1)
    
    For i = 3 To UBound(vLine)
        vItem = Split(vLine(i), vbTab)
        
        If UBound(vItem) = 0 Then GoTo Next_I
        
        If vItem(0) = strCO Then
            If vAttList(4).TextString = "" Then
                vAttList(4).TextString = vItem(1)
            Else
                vAttList(4).TextString = vAttList(4).TextString & " " & vItem(1)
            End If
        Else
            If vAttList(4).TextString = "" Then
                vAttList(4).TextString = vItem(0) & "=" & vItem(1)
            Else
                vAttList(4).TextString = vAttList(4).TextString & " " & vItem(0) & "=" & vItem(1)
            End If
        End If
Next_I:
    Next i
    
Add_Notes:
    
    If tbNoteList.Value = "" Then
        vAttList(24).TextString = ""
    Else
        strLine = Replace(tbNoteList.Value, vbLf, "")
        strLine = Replace(strLine, vbCr, ";")
        
        vAttList(24).TextString = strLine
    End If
    
Exit_Sub:
    objPole.Update
    
    iChanged = 1
End Sub

Public Sub GetPoleData()
    Dim vAttList As Variant
    Dim vAll, vLine, vCount As Variant
    Dim strLine, strCO, strTemp As String
    Dim iIndex As Integer
        
    vAttList = objPole.GetAttributes
    
    tbPoleNumber.Value = vAttList(0).TextString
    
    lbWorkspace.Clear
    lbMR.Clear
    tbPoleData.Value = ""
    
    cbNote.Value = ""
    
    If vAttList(1).TextString = "DEFECTIVE" Then cbNote.Value = "Power to replace Defective Pole"
    If vAttList(1).TextString = "REPLACE" Then cbNote.Value = "Power to replace Defective Pole"
    
    '<----------------------------------------------------------------------Pole Data
    
    strLine = vAttList(0).TextString
    strLine = strLine & vbCr & vAttList(2).TextString & vbTab & vAttList(5).TextString
    strLine = strLine & vbCr & vAttList(2).TextString & vbTab & vAttList(3).TextString
    
    If InStr(vAttList(5).TextString, ")") > 0 Then
        strTemp = Replace(vAttList(5).TextString, "(", "")
        strTemp = Replace(strTemp, "S", "")
        strTemp = Replace(strTemp, "C", "")
        strTemp = Replace(strTemp, "H", "")
        vAll = Split(strTemp, ")")
        
        vLine = Split(vAll(0), "-")
        'If InStr(vLine(0), "S") > 0 Then vLine(0) = Replace(vLine(0), "S", "")
        'If InStr(vLine(0), "C") > 0 Then vLine(0) = Replace(vLine(0), "C", "")
        'If InStr(vLine(0), "H") > 0 Then vLine(0) = Replace(vLine(0), "H", "")
        
        vCount = Split(vAll(1), "-")
        'If InStr(vCount(0), "S") > 0 Then vCount(0) = Replace(vCount(0), "S", "")
        'If InStr(vCount(0), "C") > 0 Then vCount(0) = Replace(vCount(0), "C", "")
        'If InStr(vCount(0), "H") > 0 Then vCount(0) = Replace(vCount(0), "H", "")
        
        'MsgBox vCount(0) & " , " & vLine(0) & vbCr & vCount(1) & " , " & vLine(1)
        
        If CInt(vCount(0)) > CInt(vLine(0)) Then
            cbNote.Value = "Power to replace with Taller Pole"
        Else
            If CInt(vCount(1)) < CInt(vLine(1)) Then
                cbNote.Value = "Power to replace with Thicker Pole"
            Else
                cbNote.Value = "Power to replace with Like Pole"
            End If
        End If
    End If
    
    If Not vAttList(4).TextString = "" Then
        vAll = Split(vAttList(4).TextString, " ")
        For i = 0 To UBound(vAll)
            If InStr(vAll(i), "=") > 0 Then
                vLine = Split(vAll(i), "=")
                
                strLine = strLine & vbCr & vLine(0) & vbTab & vLine(1)
            Else
                strLine = strLine & vbCr & vAttList(2).TextString & vbTab & vAll(i)
            End If
        Next i
    End If
    
    If Not vAttList(8).TextString = "" Then
        Select Case vAttList(8).TextString
            Case "M"
                vAttList(8).TextString = "MGNV"
            Case "T"
                vAttList(8).TextString = "TGB"
            Case "B"
                vAttList(8).TextString = "BROKEN GRD"
        End Select
        
        strLine = strLine & vbCr & vAttList(8).TextString
    End If
    
    tbPoleData.Value = strLine
    
    tbLL.Value = vAttList(7).TextString
    
    '<----------------------------------------------------------------------Power
    
    If Not vAttList(9).TextString = "" Then
        lbWorkspace.AddItem "PWR"
        iIndex = lbWorkspace.ListCount - 1
        lbWorkspace.List(iIndex, 1) = "NEUTRAL"
        
        strLine = vAttList(9).TextString
        If InStr(strLine, ")") > 0 Then
            vLine = Split(strLine, ")")
            
            lbWorkspace.List(iIndex, 2) = Replace(vLine(0), "(", "")
            lbWorkspace.List(iIndex, 3) = vLine(1)
        Else
            If InStr(strLine, ">") > 0 Then
                vLine = Split(strLine, ">")
                vCount = Split(vLine(1), "-")
                vCount(0) = CInt(vCount(0)) + 1
                
                strLine = ""
                If Not vLine(0) = "" Then strLine = vLine(0)
                strLine = strLine & vCount(0)
                If UBound(vCount) > 0 Then strLine = strLine & "-" & vCount(1)
            End If
            
            lbWorkspace.List(iIndex, 2) = strLine
        End If
    End If
    
    If Not vAttList(10).TextString = "" Then
        lbWorkspace.AddItem "PWR"
        iIndex = lbWorkspace.ListCount - 1
        lbWorkspace.List(iIndex, 1) = "TRANSFORMER"
        
        strLine = vAttList(10).TextString
        If InStr(strLine, ")") > 0 Then
            vLine = Split(strLine, ")")
            
            lbWorkspace.List(iIndex, 2) = Replace(vLine(0), "(", "")
            lbWorkspace.List(iIndex, 3) = vLine(1)
        Else
            If InStr(strLine, ">") > 0 Then
                vLine = Split(strLine, ">")
                vCount = Split(vLine(1), "-")
                vCount(0) = CInt(vCount(0)) + 1
                
                strLine = ""
                If Not vLine(0) = "" Then strLine = vLine(0)
                strLine = strLine & vCount(0)
                If UBound(vCount) > 0 Then strLine = strLine & "-" & vCount(1)
            End If
            
            lbWorkspace.List(iIndex, 2) = strLine
        End If
    End If
    
    If Not vAttList(11).TextString = "" Then
        lbWorkspace.AddItem "PWR"
        iIndex = lbWorkspace.ListCount - 1
        lbWorkspace.List(iIndex, 1) = "LOW POWER"
        
        strLine = vAttList(11).TextString
        If InStr(strLine, ")") > 0 Then
            vLine = Split(strLine, ")")
            
            lbWorkspace.List(iIndex, 2) = Replace(vLine(0), "(", "")
            lbWorkspace.List(iIndex, 3) = vLine(1)
        Else
            If InStr(strLine, ">") > 0 Then
                vLine = Split(strLine, ">")
                vCount = Split(vLine(1), "-")
                vCount(0) = CInt(vCount(0)) + 1
                
                strLine = ""
                If Not vLine(0) = "" Then strLine = vLine(0)
                strLine = strLine & vCount(0)
                If UBound(vCount) > 0 Then strLine = strLine & "-" & vCount(1)
            End If
            
            lbWorkspace.List(iIndex, 2) = strLine
        End If
    End If
    
    If Not vAttList(12).TextString = "" Then
        lbWorkspace.AddItem "PWR"
        iIndex = lbWorkspace.ListCount - 1
        lbWorkspace.List(iIndex, 1) = "ANTENNA"
        
        strLine = vAttList(12).TextString
        If InStr(strLine, ")") > 0 Then
            vLine = Split(strLine, ")")
            
            lbWorkspace.List(iIndex, 2) = Replace(vLine(0), "(", "")
            lbWorkspace.List(iIndex, 3) = vLine(1)
        Else
            If InStr(strLine, ">") > 0 Then
                vLine = Split(strLine, ">")
                vCount = Split(vLine(1), "-")
                vCount(0) = CInt(vCount(0)) + 1
                
                strLine = ""
                If Not vLine(0) = "" Then strLine = vLine(0)
                strLine = strLine & vCount(0)
                If UBound(vCount) > 0 Then strLine = strLine & "-" & vCount(1)
            End If
            
            lbWorkspace.List(iIndex, 2) = strLine
        End If
    End If
    
    If Not vAttList(13).TextString = "" Then
        strLine = vAttList(13).TextString
        vAll = Split(strLine, " ")
        For i = 0 To UBound(vAll)
            lbWorkspace.AddItem "PWR"
            iIndex = lbWorkspace.ListCount - 1
            lbWorkspace.List(iIndex, 1) = "ST LT CIRCUIT"
            
            If InStr(vAll(i), ")") > 0 Then
                vLine = Split(vAll(i), ")")
            
                lbWorkspace.List(iIndex, 2) = Replace(vLine(0), "(", "")
                lbWorkspace.List(iIndex, 3) = vLine(1)
            Else
            If InStr(strLine, ">") > 0 Then
                vLine = Split(strLine, ">")
                vCount = Split(vLine(1), "-")
                vCount(0) = CInt(vCount(0)) + 1
                
                strLine = ""
                If Not vLine(0) = "" Then strLine = vLine(0)
                strLine = strLine & vCount(0)
                If UBound(vCount) > 0 Then strLine = strLine & "-" & vCount(1)
            End If
            
                lbWorkspace.List(iIndex, 2) = vAll(i)
            End If
        Next i
    End If
    
    If Not vAttList(14).TextString = "" Then
        strLine = vAttList(14).TextString
        vAll = Split(strLine, " ")
        For i = 0 To UBound(vAll)
            lbWorkspace.AddItem "PWR"
            iIndex = lbWorkspace.ListCount - 1
            lbWorkspace.List(iIndex, 1) = "ST LT"
        
            If InStr(vAll(i), ")") > 0 Then
                vLine = Split(vAll(i), ")")
            
                lbWorkspace.List(iIndex, 2) = Replace(vLine(0), "(", "")
                lbWorkspace.List(iIndex, 3) = vLine(1)
            Else
                If InStr(strLine, ">") > 0 Then
                    vLine = Split(strLine, ">")
                    vCount = Split(vLine(1), "-")
                    vCount(0) = CInt(vCount(0)) + 1
                    
                    strLine = ""
                    If Not vLine(0) = "" Then strLine = vLine(0)
                    strLine = strLine & vCount(0)
                    If UBound(vCount) > 0 Then strLine = strLine & "-" & vCount(1)
                End If
                
                lbWorkspace.List(iIndex, 2) = vAll(i)
            End If
        Next i
    End If
    
    '<----------------------------------------------------------------------New Attachment
    
    If Not vAttList(15).TextString = "" Then
        strLine = UCase(vAttList(15).TextString)
        vAll = Split(strLine, " ")
        For i = 0 To UBound(vAll)
            lbWorkspace.AddItem "NEW"
            iIndex = lbWorkspace.ListCount - 1
            lbWorkspace.List(iIndex, 1) = "NEW 6M"
            
            If InStr(vAll(i), "O") > 0 Then
                vAll(i) = Replace(vAll(i), "O", "")
                lbWorkspace.List(iIndex, 4) = "NEW OHG"
            End If
            
            If InStr(vAll(i), "P") > 0 Then
                lbWorkspace.List(iIndex, 1) = lbWorkspace.List(iIndex, 1) & " TAP"
                vAll(i) = Replace(vAll(i), "P", "")
            End If
            
            If InStr(vAll(i), "T") > 0 Then
                vAll(i) = Replace(vAll(i), "T", "")
                lbWorkspace.List(iIndex, 4) = "MTE TAG"
            End If
            
            If InStr(vAll(i), "F") > 0 Then
                vAll(i) = Replace(vAll(i), "F", "")
                lbWorkspace.List(iIndex, 4) = "FUTURE"
            End If
            
            lbWorkspace.List(iIndex, 3) = vAll(i)
        Next i
    End If
    
    '<----------------------------------------------------------------------Communications
    
    For n = 16 To 23
        If Not vAttList(n).TextString = "" Then
            strLine = UCase(vAttList(n).TextString)
            vLine = Split(strLine, "=")
            strCO = vLine(0)
        
            vAll = Split(vLine(1), " ")
            For i = 0 To UBound(vAll)
                lbWorkspace.AddItem "COMM"
                iIndex = lbWorkspace.ListCount - 1
        
                lbWorkspace.List(iIndex, 1) = strCO
        
                If InStr(vAll(i), "C") > 0 Then
                    lbWorkspace.List(iIndex, 1) = lbWorkspace.List(iIndex, 1) & " C-WIRE"
                    vAll(i) = Replace(vAll(i), "C", "")
                End If
        
                If InStr(vAll(i), "D") > 0 Then
                    lbWorkspace.List(iIndex, 1) = lbWorkspace.List(iIndex, 1) & " DROP"
                    vAll(i) = Replace(vAll(i), "D", "")
                End If
            
                If InStr(vAll(i), "E") > 0 Then
                    lbWorkspace.List(iIndex, 4) = "EXTEND"
                    vAll(i) = Replace(vAll(i), "E", "")
                End If
            
                If InStr(vAll(i), "O") > 0 Then
                    lbWorkspace.List(iIndex, 1) = lbWorkspace.List(iIndex, 1) & " OHG"
                    vAll(i) = Replace(vAll(i), "O", "")
                End If
            
                If InStr(vAll(i), "P") > 0 Then
                    lbWorkspace.List(iIndex, 1) = lbWorkspace.List(iIndex, 1) & " TAP"
                    vAll(i) = Replace(vAll(i), "P", "")
                End If
        
                If InStr(vAll(i), "S") > 0 Then
                    lbWorkspace.List(iIndex, 1) = lbWorkspace.List(iIndex, 1) & " SS"
                    vAll(i) = Replace(vAll(i), "S", "")
                End If
            
                If InStr(vAll(i), "T") > 0 Then
                    lbWorkspace.List(iIndex, 1) = lbWorkspace.List(iIndex, 1) & " TAG"
                    vAll(i) = Replace(vAll(i), "T", "")
                End If
            
                If InStr(vAll(i), "V") > 0 Then
                    lbWorkspace.List(iIndex, 1) = "LASH TO " & lbWorkspace.List(iIndex, 1)
                    vAll(i) = Replace(vAll(i), "V", "")
                End If
            
                If InStr(vAll(i), "X") > 0 Then
                    lbWorkspace.List(iIndex, 4) = "ATTACH"
                    vAll(i) = Replace(vAll(i), "X", "")
                End If
            
                If InStr(vAll(i), ")") > 0 Then
                    vLine = Split(vAll(i), ")")
            
                    lbWorkspace.List(iIndex, 2) = Replace(vLine(0), "(", "")
                    lbWorkspace.List(iIndex, 3) = vLine(1)
                Else
                    If lbWorkspace.List(iIndex, 4) = "ATTACH" Then
                        lbWorkspace.List(iIndex, 3) = vAll(i)
                    Else
                        lbWorkspace.List(iIndex, 2) = vAll(i)
                    End If
                End If
            Next i
        End If
    Next n
    
    tbNoteList.Value = ""
    
    If Not vAttList(24).TextString = "" Then
        vLine = Split(vAttList(24).TextString, ";")
        strLine = vLine(0)
        If UBound(vLine) > 0 Then
            For i = 1 To UBound(vLine)
                If Not vLine(i) = "" Then strLine = strLine & vbCr & vLine(i)
            Next i
        End If
    
        tbNoteList.Value = strLine
    End If
    
    If InStr(tbNoteList.Value, "NOTE-DEF") < 1 Then
        If vAttList(1).TextString = "DEFECTIVE" Then tbNoteList.Value = tbNoteList.Value & vbCr & "NOTE-DEF"
        If vAttList(1).TextString = "REPLACE" Then tbNoteList.Value = tbNoteList.Value & vbCr & "NOTE-DEF"
    End If
    
Exit_Sub:
End Sub
