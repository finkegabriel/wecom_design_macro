VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RearrangeCounts 
   Caption         =   "Rearrange Counts"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   OleObjectBlob   =   "RearrangeCounts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RearrangeCounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objBlock As AcadBlockReference
Dim iListIndex As Integer

Private Sub cbAddRange_Click()
    If cbAddRange.Caption = "Update" Then
        If tbRFibers.Value = "" Then GoTo Done_Updating
        If tbRName.Value = "" Then tbRName.Value = "**"
        If tbRCounts.Value = "" Then tbRCounts.Value = "*-*"
        
        lbRange.List(iListIndex, 0) = tbRFibers.Value
        lbRange.List(iListIndex, 1) = tbRName.Value
        lbRange.List(iListIndex, 2) = tbRCounts.Value
        
Done_Updating:
        cbAddRange.Caption = "Add Range"
        Exit Sub
    End If
    
    If cbCableSize.Value = "" Then Exit Sub
    If tbRFibers.Value = "" Then Exit Sub
    If tbRName.Value = "" Then Exit Sub
    If tbRCounts.Value = "" Then Exit Sub
    
    'lbRange.AddItem tbRFibers.Value
    'lbRange.List(lbRange.ListCount - 1, 1) = tbRName.Value
    'lbRange.List(lbRange.ListCount - 1, 2) = tbRCounts.Value
    
    Dim vCounts As Variant
    Dim strName As String
    Dim iIndex, iSize, iTemp As Integer
    Dim iStart, iEnd As Integer
    Dim iLStart, iLEnd As Integer
    Dim iLCStart, iLCEnd As Integer
    
    iSize = CInt(cbCableSize.Value)
    
    vCounts = Split(tbRFibers.Value, "-")
    iStart = CInt(vCounts(0))
    If UBound(vCounts) = 0 Then
        iEnd = iStart
    Else
        iEnd = CInt(vCounts(1))
    End If
    
    If iStart > iSize Then Exit Sub
    'If iEnd > iSize Then iEnd = iSize
    
    For i = 0 To lbRange.ListCount - 1
        If Not lbRange.List(i, 1) = "**" And Not lbRange.List(i, 2) = "*-*" Then GoTo Not_Present
        
        strName = lbRange.List(i, 1)
        
        vCounts = Split(lbRange.List(i, 0), "-")
        iLStart = CInt(vCounts(0))
        iLEnd = CInt(vCounts(1))
        
        If iStart > iLEnd Then GoTo Not_Present
        
        Select Case (iLStart - iStart)
            Case Is = 0
                Select Case (iLEnd - iEnd)
                    Case Is < 0
                        lbRange.List(i, 1) = tbRName.Value
                        If tbRCounts.Value = "*-*" Then
                            lbRange.List(i, 2) = tbRCounts.Value
                        Else
                            vCounts = Split(tbRCounts.Value, "-")
                            vCounts(1) = CInt(vCounts(0)) + iLEnd - iLStart
                            
                            lbRange.List(i, 2) = vCounts(0) & "-" & vCounts(1)
                        End If
                    Case Is = 0
                        lbRange.List(i, 0) = tbRFibers.Value
                        lbRange.List(i, 1) = tbRName.Value
                        lbRange.List(i, 2) = tbRCounts.Value
                    Case Is > 0
                        lbRange.List(i, 0) = tbRFibers.Value
                        lbRange.List(i, 1) = tbRName.Value
                        lbRange.List(i, 2) = tbRCounts.Value
                        
                        iTemp = iEnd + 1
                        lbRange.AddItem iTemp & "-" & iLEnd, i + 1
                        lbRange.List(i + 1, 1) = strName
                        lbRange.List(i + 1, 2) = "*-*"
                End Select
            Case Is > 0
                Select Case (iLEnd - iEnd)
                    
                End Select
        End Select
        
Not_Present:
    Next i
    
    tbRFibers.Value = ""
    tbRName.Value = ""
    tbRCounts.Value = ""
    tbRFibers.SetFocus
End Sub

Private Sub cbAddToList_Click()
    If tbEndFibers.Value = "" Then Exit Sub
    If tbSource.Value = "" Then Exit Sub
    
    lbList.AddItem tbEndFibers.Value
    If UCase(tbSource.Value) = "XD" Then tbSource.Value = UCase(tbSource.Value) & ": " & tbEndFibers.Value
    lbList.List(lbList.ListCount - 1, 1) = tbSource.Value
    
    tbEndFibers.Value = ""
    tbSource.Value = ""
    tbEndFibers.SetFocus
End Sub

Private Sub cbCableSize_Change()
    If cbCableSize.Value = "" Then Exit Sub
    
    lbRange.Clear
    lbRange.AddItem "1-" & cbCableSize.Value
    lbRange.List(0, 1) = "**"
    lbRange.List(0, 2) = "*-*"
End Sub

Private Sub cbGetBlock_Click()
    Dim objEntity As AcadEntity
    Dim vReturnPnt As Variant
    Dim vAttList As Variant
    Dim strLine, strCable As String
    Dim iStart, iEnd, iSize As Integer
    
    Me.Hide
    
    On Error Resume Next
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    vAttList = objBlock.GetAttributes
    
    Dim vLine, vItem, vCounts, vTemp As Variant
    
    Select Case objBlock.Name
        Case "sPole"
            strLine = vAttList(25).TextString
        Case "sPed", "sHH", "sFP", "sMH", "sPanel"
            strLine = vAttList(5).TextString
        Case "Callout"
            strCable = vAttList(1).TextString
            vLine = Split(strCable, ")")
            vItem = Split(vLine(0), "(")
            iSize = CInt(vItem(1))
            'If Not vItem(1) = cbCableSize.Value Then
                'If cbCableSize.Value = "" Then
                    'cbCableSize.Value = vItem(1)
                'Else
                    'GoTo Exit_Sub
                'End If
            'End If
            
            strLine = Replace(vAttList(2).TextString, "\P", " + ")
            strCable = vAttList(1).TextString
            'MsgBox strLine
            GoTo Fixed_Line
        Case "CableCounts"
            strCable = vAttList(1).TextString
            vLine = Split(strCable, ")")
            vItem = Split(vLine(0), "(")
            iSize = CInt(vItem(1))
            'If Not vItem(1) = cbCableSize.Value Then
                'If cbCableSize.Value = "" Then
                    'cbCableSize.Value = vItem(1)
                'Else
                    'GoTo Exit_Sub
                'End If
            'End If
            
            strLine = Replace(vAttList(0).TextString, "\P", " + ")
            strCable = vAttList(1).TextString
            'MsgBox strLine
            GoTo Fixed_Line
        Case Else
            GoTo Exit_Sub
    End Select
    
    vLine = Split(strLine, " / ")
    strCable = vLine(0)
    
    vItem = Split(strCable, ")")
    vTemp = Split(vItem(0), "(")
    iSize = CInt(vTemp(1))
    
    If lbList.ListCount < 1 Then
        lbList.AddItem "1-" & iSize
        lbList.List(0, 1) = "1-" & iSize
    End If
    'If Not vTemp(1) = cbCableSize.Value Then
        'If cbCableSize.Value = "" Then
            'cbCableSize.Value = vTemp(1)
        'Else
            'GoTo Exit_Sub
        'End If
    'End If
    
    strLine = vLine(1)
    
Fixed_Line:
    
    'If cbCableSize.Value = "" Then GoTo Exit_Sub
    If strLine = "" Then GoTo Exit_Sub
    
    Dim strName, strSource As String
    Dim iFiber As Integer
    
    iFiber = 1
    lbCounts.Clear
    'lbCounts.AddItem ""
    
    vLine = Split(strLine, " + ")
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        strName = vItem(0)
        If UBound(vItem) < 2 Then
            strSource = ""
        Else
            strSource = vItem(2)
        End If
        
        vCounts = Split(vItem(1), "-")
        iStart = CInt(vCounts(0))
        If UBound(vCounts) < 1 Then
            iEnd = iStart
        Else
            iEnd = CInt(vCounts(1))
        End If
        
        strLine = iFiber & "-"
        iFiber = iFiber + iEnd - iStart
        strLine = strLine & iFiber
        
        lbCounts.AddItem strLine
        lbCounts.List(lbCounts.ListCount - 1, 1) = strName
        lbCounts.List(lbCounts.ListCount - 1, 2) = vItem(1)
        lbCounts.List(lbCounts.ListCount - 1, 3) = strSource
        
        iFiber = iFiber + 1
        
        'For j = iStart To iEnd
            'lbCounts.AddItem iFiber
            'lbCounts.List(lbCounts.ListCount - 1, 1) = strName
            'lbCounts.List(lbCounts.ListCount - 1, 2) = j
            'lbCounts.List(lbCounts.ListCount - 1, 3) = strSource
            
            'iFiber = iFiber + 1
        'Next j
    Next i
    
    tbCable.Value = strCable
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbRearrange_Click()
    If lbList.ListCount < 1 Then Exit Sub
    If lbCounts.ListCount < 1 Then Exit Sub
    If cbEndFiber.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vCounts, vTemp As Variant
    Dim strResult, strLine, strName, strSource As String
    Dim iFStart, iFEnd As Integer
    Dim iCStart, iCEnd As Integer
    Dim iTStart, iTEnd As String
    Dim iFiber As Integer
    
    iFiber = 1
    strResult = ""
    
    If InStr(tbCable.Value, cbCableSize.Value) = 0 Then
        Exit Sub
    End If
    
    tbCable.Value = Replace(tbCable.Value, cbCableSize.Value, cbEndFiber.Value)
    
    For i = 0 To lbList.ListCount - 1
        vLine = Split(lbList.List(i, 1), ": ")
        
        Select Case UBound(vLine)
            Case Is = 0
                vCounts = Split(vLine(0), "-")
                iFStart = CInt(vCounts(0))
                If UBound(vCounts) = 0 Then
                    iFEnd = iFStart
                Else
                    iFEnd = CInt(vCounts(1))
                End If
                
                For j = 0 To lbCounts.ListCount - 1
                    strLine = ""
                    vCounts = Split(lbCounts.List(j, 0), "-")
                    iCStart = CInt(vCounts(0))
                    If UBound(vCounts) = 0 Then
                        iCEnd = iCStart
                    Else
                        iCEnd = CInt(vCounts(1))
                    End If
                    
                    'MsgBox "FS: " & iFStart & vbTab & "FE: " & iFEnd & vbCr & "CS: " & iCStart & vbTab & "CE: " & iCEnd
                    
                    Select Case iFStart - iCStart
                        Case Is = 0
                            Select Case iFEnd - iCEnd
                                Case Is < 0
                                    vCounts = Split(lbCounts.List(j, 2), "-")
                                    vCounts(1) = CInt(vCounts(0)) + iFEnd - iFStart
                                    strLine = lbCounts.List(j, 1) & ": " & vCounts(0) & "-" & vCounts(1)
                                    If Not lbCounts.List(j, 3) = "" Then strLine = strLine & ": " & lbCounts.List(j, 3)
                                Case Is = 0
                                    strLine = lbCounts.List(j, 1) & ": " & lbCounts.List(j, 2)
                                    
                                    If Not lbCounts.List(j, 3) = "" Then strLine = strLine & ": " & lbCounts.List(j, 3)
                                Case Is > 0
                                    strLine = lbCounts.List(j, 1) & ": " & lbCounts.List(j, 2)
                                    
                                    If Not lbCounts.List(j, 3) = "" Then strLine = strLine & ": " & lbCounts.List(j, 3)
                                    iFStart = iCEnd + 1
                            End Select
                        Case Is > 0
                            If iFStart > iCEnd Then GoTo Next_line
                            
                            vCounts = Split(lbCounts.List(j, 2), "-")
                            vCounts(0) = CInt(vCounts(0)) + iFStart - iCStart
                            
                            Select Case iFEnd - iCEnd
                                Case Is < 0
                                    vCounts(1) = CInt(vCounts(0)) + iFEnd - iFStart
                                    
                                    strLine = lbCounts.List(j, 1) & ": " & vCounts(0) & "-" & vCounts(1)
                                    If Not lbCounts.List(j, 3) = "" Then strLine = strLine & ": " & lbCounts.List(j, 3)
                                Case Is = 0
                                    vCounts(1) = CInt(vCounts(0)) + iFEnd - iFStart
                                    
                                    strLine = lbCounts.List(j, 1) & ": " & vCounts(0) & "-" & vCounts(1)
                                    
                                    If Not lbCounts.List(j, 3) = "" Then strLine = strLine & ": " & lbCounts.List(j, 3)
                                Case Is > 0
                                    strLine = lbCounts.List(j, 1) & ": " & vCounts(0) & "-" & vCounts(1)
                                    
                                    If Not lbCounts.List(j, 3) = "" Then strLine = strLine & ": " & lbCounts.List(j, 3)
                                    iFStart = iCEnd + 1
                            End Select
                        Case Is < 0
                            'If iCStart > iFEnd Then GoTo Next_Line
                            
                            
                            GoTo Next_line
                    End Select
                    
                    If strResult = "" Then
                        strResult = strLine
                    Else
                        strResult = strResult & vbCr & strLine
                    End If
Next_line:
                    'MsgBox "i: " & i & vbCr & "j: " & j & vbCr & "strLine:  " & strLine
                Next j
            Case Is > 0
                strLine = lbList.List(i, 1)
        
                If strResult = "" Then
                    strResult = strLine
                Else
                    strResult = strResult & vbCr & strLine
                End If
        End Select
    Next i
    
    iFiber = 1
    strSource = ""
    vLine = Split(strResult, vbCr)
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        If UBound(vItem) > 1 Then
            If Not vItem(2) = "" Then strSource = vItem(2)
        End If
        
        vCounts = Split(vItem(1), "-")
        iTStart = CInt(vCounts(0))
        If UBound(vCounts) = 0 Then
            iTEnd = iTStart
        Else
            iTEnd = CInt(vCounts(1))
        End If
        
        If vItem(0) = "XD" Then
            iTEnd = iFiber + iTEnd - iTStart
            iTStart = iFiber
            
            vLine(i) = vItem(0) & ": " & iTStart & "-" & iTEnd
            If UBound(vItem) > 1 Then vLine(i) = vLine(i) & ": " & vItem(2)
            iFiber = iTEnd + 1
        Else
            iFiber = iFiber + iTEnd - iTStart + 1
        End If
    Next i
    
    If Not strSource = "" Then
        vItem = Split(vLine(0), ": ")
                
        If UBound(vItem) < 2 Then
            vLine(0) = vLine(0) & ": " & strSource
        Else
            If vItem(2) = "" Then vLine(0) = vItem(0) & ": " & vItem(1) & ": " & strSource
        End If
    End If
    
    strResult = vLine(0)
    If UBound(vLine) > 0 Then
        For i = 1 To UBound(vLine)
            If Not strSource = "" Then
                vItem = Split(vLine(i), ": ")
                
                If UBound(vItem) < 2 Then
                    vLine(i) = vLine(i) & ": " & strSource
                Else
                    If vItem(2) = "" Then vLine(i) = vItem(0) & ": " & vItem(1) & ": " & strSource
                End If
            End If
            
            strResult = strResult & vbCr & vLine(i)
        Next i
    End If
    
    tbResult.Value = UCase(strResult)
    Call ConsolidateCounts
End Sub

Private Sub cbUpdateBlock_Click()
    If tbCable.Value = "" Then Exit Sub
    If tbResult.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vCounts, vTemp As Variant
    Dim vAttList As Variant
    Dim strLine, strPosition As String
    
    vAttList = objBlock.GetAttributes
    vTemp = Split(tbCable.Value, ": ")
    If UBound(vTemp) = 0 Then
        strPosition = ""
    Else
        strPosition = vTemp(0)
    End If
    
    strLine = Replace(tbResult.Value, vbCr, " + ")
    strLine = Replace(strLine, vbLf, "")
    
    Select Case objBlock.Name
        Case "sPole"
            strLine = tbCable.Value & " / " & strLine
            
            vTemp = Split(vAttList(25).TextString, vbCr)
            For i = 0 To UBound(vTemp)
                vLine = Split(vTemp(i), ": ")
                If vLine(0) = strPosition Then
                    vTemp(i) = strLine
                End If
            Next i
            
            strLine = vTemp(0)
            If UBound(vTemp) > 0 Then
                For i = 1 To UBound(vTemp)
                    strLine = strLine & vbCr & vTemp(i)
                Next i
            End If
            
            vAttList(25).TextString = strLine
        Case "sPed", "sHH", "sFP", "sMH", "sPanel"
            strLine = tbCable.Value & " / " & strLine
            
            vTemp = Split(vAttList(5).TextString, vbCr)
            For i = 0 To UBound(vTemp)
                vLine = Split(vTemp(i), ": ")
                If vLine(0) = strPosition Then
                    vTemp(i) = strLine
                End If
            Next i
            
            strLine = vTemp(0)
            If UBound(vTemp) > 0 Then
                For i = 1 To UBound(vTemp)
                    strLine = strLine & vbCr & vTemp(i)
                Next i
            End If
            
            vAttList(5).TextString = strLine
        Case "Callout"
            vAttList(1).TextString = tbCable.Value
            vAttList(2).TextString = Replace(strLine, " + ", "\P")
        Case "CableCounts"
            vAttList(1).TextString = tbCable.Value
            vAttList(0).TextString = Replace(strLine, " + ", "\P")
    End Select
    
    objBlock.Update
End Sub

Private Sub CommandButton1_Click()
    If lbRange.ListCount < 1 Then Exit Sub
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim vLine As Variant
    Dim strResult, strName As String
    Dim strTemp, strLine As String
    Dim iFiber, iDiff As Integer
    Dim iFTest, iCTest As Integer
    Dim iRStart, iREnd As Integer
    Dim iCStart, iCEnd As Integer
    Dim iType, iLen As Integer
    Dim bFound As Boolean
    'Dim lLen As Long
    
    strResult = ""
    iFiber = 1
    
    For i = 0 To lbRange.ListCount - 1
        vLine = Split(lbRange.List(i, 0), "-")
        iRStart = CInt(vLine(0))
        iREnd = CInt(vLine(1))
        strName = lbRange.List(i, 1)
        
        iType = 0
        If Left(strName, 1) = "*" Then iType = iType + 1
        If Right(strName, 1) = "*" Then iType = iType + 2
        strName = Replace(strName, "*", "")
        
        For j = 0 To lbCounts.ListCount - 1
            strLine = lbCounts.List(j, 1)
            If strLine = "XD" Then GoTo Next_Count_Line
            
            bFound = False
            Select Case iType
                Case Is = 0
                    If strLine = strName Then bFound = True
                Case Is = 1
                    iLen = Len(strName)
                    'lLen = Len(strName)
                    If Right(strLine, iLen) = strName Then bFound = True
                Case Is = 2
                    iLen = Len(strName)
                    'lLen = Len(strName)
                    If Left(strLine, iLen) = strName Then bFound = True
                Case Is = 3
                    If strName = "" Then
                        bFound = True
                    Else
                        If InStr(strLine, strName) > 0 Then bFound = True
                    End If
            End Select
            If bFound = False Then GoTo Next_Count_Line
            
            vLine = Split(lbCounts.List(j, 0), "-")
            iCStart = CInt(vLine(0))
            iCEnd = CInt(vLine(1))
            iDiff = iCEnd - iCStart
            
            Select Case iREnd - iRStart
                Case Is = iDiff
                    
            End Select
            
            
Next_Count_Line:
        Next j
    Next i
End Sub

Private Sub lbList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strAtt(1) As String
    Dim iIndex As Integer
    
    Select Case KeyCode
        'Case vbKeyReturn
            'iListIndex = lbList.ListIndex
    
            'tbEndFibers.Value = lbList.List(iListIndex, 0)
            'tbSource.Value = lbList.List(iListIndex, 1)
    
            'cbAddToList.Caption = "Update"
        Case vbKeyUp
            iIndex = lbList.ListIndex
            If iIndex < 1 Then Exit Sub
            
            strAtt(0) = lbList.List(iIndex, 0)
            strAtt(1) = lbList.List(iIndex, 1)
            
            lbList.List(iIndex, 0) = lbList.List(iIndex - 1, 0)
            lbList.List(iIndex, 1) = lbList.List(iIndex - 1, 1)
            
            lbList.List(iIndex - 1, 0) = strAtt(0)
            lbList.List(iIndex - 1, 1) = strAtt(1)
        Case vbKeyDown
            iIndex = lbList.ListIndex
            If iIndex > lbList.ListCount - 2 Then Exit Sub
            
            strAtt(0) = lbList.List(iIndex, 0)
            strAtt(1) = lbList.List(iIndex, 1)
            
            lbList.List(iIndex, 0) = lbList.List(iIndex + 1, 0)
            lbList.List(iIndex, 1) = lbList.List(iIndex + 1, 1)
            
            lbList.List(iIndex + 1, 0) = strAtt(0)
            lbList.List(iIndex + 1, 1) = strAtt(1)
        Case vbKeyDelete
            lbList.RemoveItem lbList.ListIndex
    End Select
End Sub

Private Sub lbRange_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    iListIndex = lbRange.ListIndex
    
    tbRFibers.Value = lbRange.List(iListIndex, 0)
    tbRName.Value = lbRange.List(iListIndex, 1)
    tbRCounts.Value = lbRange.List(iListIndex, 2)
    
    cbAddRange.Caption = "Update"
End Sub

Private Sub lbRange_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strAtt(2) As String
    Dim iIndex As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            iListIndex = lbRange.ListIndex
    
            tbRFibers.Value = lbRange.List(iListIndex, 0)
            tbRName.Value = lbRange.List(iListIndex, 1)
            tbRCounts.Value = lbRange.List(iListIndex, 2)
    
            cbAddRange.Caption = "Update"
        Case vbKeyUp
            iIndex = lbRange.ListIndex
            If iIndex < 1 Then Exit Sub
            
            strAtt(0) = lbRange.List(iIndex, 0)
            strAtt(1) = lbRange.List(iIndex, 1)
            strAtt(2) = lbRange.List(iIndex, 2)
            
            lbRange.List(iIndex, 0) = lbRange.List(iIndex - 1, 0)
            lbRange.List(iIndex, 1) = lbRange.List(iIndex - 1, 1)
            lbRange.List(iIndex, 2) = lbRange.List(iIndex - 1, 2)
            
            lbRange.List(iIndex - 1, 0) = strAtt(0)
            lbRange.List(iIndex - 1, 1) = strAtt(1)
            lbRange.List(iIndex - 1, 2) = strAtt(2)
        Case vbKeyDown
            iIndex = lbRange.ListIndex
            If iIndex > lbRange.ListCount - 2 Then Exit Sub
            
            strAtt(0) = lbRange.List(iIndex, 0)
            strAtt(1) = lbRange.List(iIndex, 1)
            strAtt(2) = lbRange.List(iIndex, 2)
            
            lbRange.List(iIndex, 0) = lbRange.List(iIndex + 1, 0)
            lbRange.List(iIndex, 1) = lbRange.List(iIndex + 1, 1)
            lbRange.List(iIndex, 2) = lbRange.List(iIndex + 1, 2)
            
            lbRange.List(iIndex + 1, 0) = strAtt(0)
            lbRange.List(iIndex + 1, 1) = strAtt(1)
            lbRange.List(iIndex + 1, 2) = strAtt(2)
    End Select
End Sub

Private Sub tbRCounts_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeySubtract
            If tbRCounts.Value = "*-" Then
                tbRCounts.Value = "*-*"
                cbAddRange.SetFocus
                
                Exit Sub
            End If
            
            Dim vFibers, vCounts As Variant
            Dim iFiber, iCount As Integer
            
            vFibers = Split(tbRFibers.Value, "-")
            iFiber = CInt(vFibers(0))
            While iFiber > 12
                iFiber = iFiber - 12
            Wend
            
            vCounts = Split(tbRCounts.Value, "-")
            iCount = CInt(vCounts(0))
            While iCount > 12
                iCount = iCount - 12
            Wend
            
            If Not iFiber = iCount Then
                MsgBox "Off Color"
                Exit Sub
            End If
            
            iCount = CInt(vCounts(0)) + CInt(vFibers(1)) - CInt(vFibers(0))
            tbRCounts.Value = tbRCounts.Value & "-" & iCount
            
            cbAddRange.SetFocus
    End Select
End Sub

Private Sub tbRFibers_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If tbRFibers.Value = "" Then Exit Sub
    
    If InStr(tbRFibers.Value, "-") < 1 Then tbRFibers.Value = tbRFibers.Value & "-" & tbRFibers.Value
End Sub

Private Sub UserForm_Initialize()
    lbCounts.ColumnCount = 4
    lbCounts.ColumnWidths = "48;132;48;126"
    
    lbRange.ColumnCount = 3
    lbRange.ColumnWidths = "42;96;54"

    lbList.ColumnCount = 2
    lbList.ColumnWidths = "72;72"
    
    cbCableSize.AddItem "12"
    cbCableSize.AddItem "24"
    cbCableSize.AddItem "48"
    cbCableSize.AddItem "72"
    cbCableSize.AddItem "96"
    cbCableSize.AddItem "144"
    cbCableSize.AddItem "216"
    cbCableSize.AddItem "288"
    cbCableSize.AddItem "360"
    cbCableSize.AddItem "432"
    cbCableSize.AddItem "576"
    cbCableSize.AddItem "876"
    
    cbEndFiber.AddItem "12"
    cbEndFiber.AddItem "24"
    cbEndFiber.AddItem "48"
    cbEndFiber.AddItem "72"
    cbEndFiber.AddItem "96"
    cbEndFiber.AddItem "144"
    cbEndFiber.AddItem "216"
    cbEndFiber.AddItem "288"
    cbEndFiber.AddItem "360"
    cbEndFiber.AddItem "432"
    cbEndFiber.AddItem "576"
    cbEndFiber.AddItem "876"
    
    tbRName.ControlTipText = "* is a wildcard.  Ex: F1000-* will match anything that starts with F1000-"
End Sub

Private Sub ConsolidateCounts()
    Dim vLine, vItem, vCounts, vTemp As Variant
    Dim strLine, strPrevious, strSource As String
    Dim iEnd, iStart As Integer
    Dim iAStart, iAEnd As Integer
    
    strLine = Replace(tbResult.Value, vbLf, "")
    vLine = Split(strLine, vbCr)
    
    If UBound(vLine) = 0 Then Exit Sub
    
    vItem = Split(vLine(UBound(vLine)), ": ")
    strPrevious = vItem(0)
    
    vCounts = Split(vItem(1), "-")
    iStart = CInt(vCounts(0))
    iEnd = CInt(vCounts(UBound(vCounts)))
    
    If UBound(vItem) > 1 Then
        strSource = vItem(2)
    Else
        strSource = ""
    End If
    
    For i = UBound(vLine) - 1 To 0 Step -1
        vItem = Split(vLine(i), ": ")
        If Not vItem(0) = strPrevious Then GoTo Next_line
        
        If UBound(vItem) > 1 Then
            If Not vItem(2) = strSource Then GoTo Next_line
        End If
        
        vCounts = Split(vItem(1), "-")
        iAStart = CInt(vCounts(0))
        iAEnd = CInt(vCounts(UBound(vCounts)))
        
        If iStart = iAEnd + 1 Then
            iAEnd = iEnd
            vLine(i) = vItem(0) & ": " & iAStart & "-" & iAEnd
            If UBound(vItem) > 1 Then vLine(i) = vLine(i) & ": " & vItem(2)
            
            vLine(i + 1) = ""
        End If
        
Next_line:
        strPrevious = vItem(0)
        iStart = iAStart
        iEnd = iAEnd
        If UBound(vItem) > 1 Then strSource = vItem(2)
    Next i
    
    strLine = vLine(0)
    For i = 1 To UBound(vLine)
        If Not vLine(i) = "" Then strLine = strLine & vbCr & vLine(i)
    Next i
    
    tbResult.Value = strLine
End Sub
