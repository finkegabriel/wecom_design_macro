VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertToSPole 
   Caption         =   "Convert Pole to 2.0"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "ConvertToSPole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConvertToSPole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objPole As AcadBlockReference

Private Sub cbGetInfo_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strTemp As String
    Dim iCable, iSplice As Integer
    
    On Error Resume Next
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
        Err = 0
    End If
    
    iCable = 1
    iSplice = 1
    
    Me.Hide
    
    objSS.SelectOnScreen
    For Each objEntity In objSS
        If Not TypeOf objEntity Is AcadBlockReference Then GoTo Next_objEntity
        
        Set objBlock = objEntity
        
        Select Case objBlock.Name
            Case "iPole"
                vAttList = objBlock.GetAttributes
    
                tbIPoleNum.Value = vAttList(0).TextString
                tbiStatus.Value = vAttList(1).TextString
                    cbStatus.Value = True
                tbiOwner.Value = vAttList(2).TextString
                    If Not vAttList(2).TextString = "" Then cbOwner.Value = True
                tbiOwnerNum.Value = vAttList(3).TextString
                    If Not vAttList(3).TextString = "" Then cbOwnerNum.Value = True
                tbiOtherNum.Value = vAttList(4).TextString
                    If Not vAttList(4).TextString = "" Then cbOtherNum.Value = True
                tbiHC.Value = vAttList(5).TextString
                    If Not vAttList(5).TextString = "" Then cbHC.Value = True
                tbiYear.Value = vAttList(6).TextString
                tbiLL.Value = vAttList(7).TextString
                tbiGRD.Value = vAttList(8).TextString
                    If Not vAttList(8).TextString = "" Then cbGRD.Value = True
            Case "pole_attach"
                vAttList = objBlock.GetAttributes
            
                lbAttach.AddItem vAttList(2).TextString, 0
            
                strTemp = Replace(vAttList(3).TextString, "'", "-")
                strTemp = Replace(strTemp, """", "")
            
                If vAttList(4).TextString = "" Then
                    lbAttach.List(0, 1) = strTemp
                    lbAttach.List(0, 2) = ""
                Else
                    Select Case Left(vAttList(4).TextString, 1)
                        Case "N", "F"
                            lbAttach.List(0, 1) = ""
                            lbAttach.List(0, 2) = strTemp
                        Case "A"
                            lbAttach.List(0, 1) = "X"
                            lbAttach.List(0, 2) = strTemp
                        Case "R", "L"
                            lbAttach.List(0, 2) = strTemp
                        
                            vLine = Split(strTemp, "-")
                            iFeet = CInt(vLine(0))
                            If UBound(vLine) < 1 Then
                                iInch = 0
                            Else
                                iInch = CInt(vLine(1))
                            End If
                            iTotal = iFeet * 12 + iInch
                        
                            vLine = Split(vAttList(4).TextString, " ")
                            strLine = Replace(vLine(1), """", "")
                        
                            If Left(vAttList(4).TextString, 1) = "R" Then
                                iRL = 0 - CInt(strLine)
                            Else
                                iRL = CInt(strLine)
                            End If
                        
                            iTotal = iTotal + iRL
                            iFeet = Int(iTotal / 12)
                            iInch = iTotal - (iFeet * 12)
                        
                            lbAttach.List(0, 1) = iFeet & "-" & iInch
                        Case "T"
                            lbAttach.List(0, 1) = strTemp
                            lbAttach.List(0, 2) = strTemp
                        Case Else
                            lbAttach.List(0, 1) = strTemp
                            lbAttach.List(0, 2) = ""
                    End Select
            
                    lbAttach.List(0, 3) = vAttList(4).TextString
                End If
            Case "pole_unit"
                vAttList = objBlock.GetAttributes
                
                lbUnits.AddItem vAttList(3).TextString, 0
            Case "terminal"
                vAttList = objBlock.GetAttributes
                
                strTemp = Replace(vAttList(0).TextString, "\P", " + ")
                If tbSplices.Value = "" Then
                    tbSplices.Value = "[A" & iSplice & "] " & strTemp
                Else
                    tbSplices.Value = tbSplices.Value & vbCr & "[A" & iSplice & "] " & strTemp
                End If
                iSplice = iSplice + 1
            Case "CableCounts"
                vAttList = objBlock.GetAttributes
                
                strTemp = vAttList(1).TextString & " / " & Replace(vAttList(0).TextString, "\P", " + ")
                If tbCable.Value = "" Then
                    tbCable.Value = "  A" & iCable & ": " & strTemp
                Else
                    tbCable.Value = tbCable.Value & vbCr & "  A" & iCable & ": " & strTemp
                End If
                iCable = iCable + 1
            Case Else
        End Select
        
Next_objEntity:
    Next objEntity
    
    objSS.Clear
    objSS.Delete
    
    Call SortAttachments
    
    cbUpdatesPole.SetFocus
    
    Me.show
End Sub

Private Sub cbGetsPole_Click()
    Dim objEntity As AcadEntity
    'Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt, vLine As Variant
    Dim strLine As String
    
    Dim vCoords As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    On Error Resume Next
    
    Me.Hide
    
    If Not objPole Is Nothing Then
        vCoords = objPole.InsertionPoint
    
        viewCoordsB(0) = vCoords(0) - 300
        viewCoordsB(1) = vCoords(1) - 300
        viewCoordsB(2) = 0#
        viewCoordsE(0) = vCoords(0) + 300
        viewCoordsE(1) = vCoords(1) + 300
        viewCoordsE(2) = 0#
        
        ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    End If
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select sPole:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objPole = objEntity
    If Not objPole.Name = "sPole" Then GoTo Exit_Sub
    
    Call ClearForm
    
    vAttList = objPole.GetAttributes
    
    tbSPoleNum.Value = vAttList(0).TextString
    tbsStatus.Value = vAttList(1).TextString
    tbsOwner.Value = vAttList(2).TextString
    tbsOwnerNum.Value = vAttList(3).TextString
    tbsOtherNum.Value = vAttList(4).TextString
    tbsHC.Value = vAttList(5).TextString
    tbsYear.Value = vAttList(6).TextString
    tbsLL.Value = vAttList(7).TextString
    tbsGRD.Value = vAttList(8).TextString
    
    If Not vAttList(25).Value = "" Then tbCable.Value = vAttList(25).TextString
    
    If Not vAttList(26).Value = "" Then tbSplices.Value = vAttList(26).TextString
    
    If Not vAttList(27).Value = "" Then
        vLine = Split(vAttList(27).TextString, ";;")
        For i = 0 To UBound(vLine)
            lbUnits.AddItem vLine(i)
        Next i
    End If
    
Exit_Sub:
    cbGetInfo.SetFocus
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbUpdateAtt_Click()
    If lbAttach.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbAttach.ListCount - 1
        If Not lbAttach.List(i, 2) = "" Then
            If InStr(lbAttach.List(i, 0), "NEW") > 0 Then lbAttach.List(i, 0) = "UTC"
            If lbAttach.List(i, 3) = "FUTURE" Then lbAttach.List(i, 0) = lbAttach.List(i, 0) & " FUTURE"
            
            lbAttach.List(i, 1) = lbAttach.List(i, 2)
            lbAttach.List(i, 2) = ""
            lbAttach.List(i, 3) = ""
        End If
    Next i
End Sub

Private Sub cbUpdatesPole_Click()
    Dim strLine, strUnits As String
    Dim vAttList, vLine As Variant
    Dim iAtt As Integer
    
    strLine = ""
    strUnits = ""
    
    Call ClearAttributes
    
    vAttList = objPole.GetAttributes
    
    If cbPoleNum.Value = True Then vAttList(0).TextString = UCase(tbIPoleNum.Value)
    If cbStatus.Value = True Then vAttList(1).TextString = UCase(tbiStatus.Value)
    If cbOwner.Value = True Then vAttList(2).TextString = UCase(tbiOwner.Value)
    If cbOwnerNum.Value = True Then vAttList(3).TextString = UCase(tbiOwnerNum.Value)
    If cbOtherNum.Value = True Then vAttList(4).TextString = UCase(tbiOtherNum.Value)
    If cbHC.Value = True Then vAttList(5).TextString = UCase(tbiHC.Value)
    If cbYear.Value = True Then vAttList(6).TextString = UCase(tbiYear.Value)
    If cbLL.Value = True Then vAttList(7).TextString = UCase(tbiLL.Value)
    If cbGRD.Value = True Then vAttList(8).TextString = UCase(tbiGRD.Value)
    
    If lbAttach.ListCount < 1 Then GoTo Exit_Sub
    
    iAtt = 16
    
    For i = 0 To lbAttach.ListCount - 1
        strLine = lbAttach.List(i, 1)
        If Not lbAttach.List(i, 2) = "" Then
            strLine = "(" & strLine & ")" & lbAttach.List(i, 2)
        End If
        
        If lbAttach.List(i, 0) = "NEUTRAL" Then
            If vAttList(9).TextString = "" Then
                vAttList(9).TextString = strLine
            Else
                vAttList(9).TextString = vAttList(9).TextString & " " & strLine
            End If
            GoTo Next_I
        End If
        
        If lbAttach.List(i, 0) = "TRANSFORMER" Then
            If vAttList(10).TextString = "" Then
                vAttList(10).TextString = strLine
            Else
                vAttList(10).TextString = vAttList(10).TextString & " " & strLine
            End If
            GoTo Next_I
        End If
        
        If lbAttach.List(i, 0) = "LOW POWER" Then
            If vAttList(11).TextString = "" Then
                vAttList(11).TextString = strLine
            Else
                vAttList(11).TextString = vAttList(11).TextString & " " & strLine
            End If
            GoTo Next_I
        End If
        
        If lbAttach.List(i, 0) = "ANTENNA" Then
            If vAttList(12).TextString = "" Then
                vAttList(12).TextString = strLine
            Else
                vAttList(12).TextString = vAttList(12).TextString & " " & strLine
            End If
            GoTo Next_I
        End If
        
        If InStr(lbAttach.List(i, 0), "ST LT C") > 0 Then
            If vAttList(13).TextString = "" Then
                vAttList(13).TextString = strLine
            Else
                vAttList(13).TextString = vAttList(13).TextString & " " & strLine
            End If
            GoTo Next_I
        End If
        
        If lbAttach.List(i, 0) = "ST LT" Then
            If vAttList(14).TextString = "" Then
                vAttList(14).TextString = strLine
            Else
                vAttList(14).TextString = vAttList(14).TextString & " " & strLine
            End If
            GoTo Next_I
        End If
        
        If InStr(strLine, "()") > 0 Then
            strLine = Replace(strLine, "()", "")
            
            If vAttList(15).TextString = "" Then
                vAttList(15).TextString = strLine
            Else
                vAttList(15).TextString = vAttList(15).TextString & " " & strLine
            End If
            GoTo Next_I
        End If
        
        If iAtt = 16 Then
            vAttList(16).TextString = lbAttach.List(i, 0) & "=" & strLine
            iAtt = iAtt + 1
            GoTo Next_I
        End If
        
        For j = 16 To iAtt
            If vAttList(j).TextString = "" Then GoTo New_COMM
            
            vLine = Split(vAttList(j).TextString, "=")
            If vLine(0) = lbAttach.List(i, 0) Then
                vAttList(j).TextString = vAttList(j).TextString & " " & strLine
                iAtt = iAtt + 1
                GoTo Next_I
            End If
        Next j
        
New_COMM:
        vAttList(j).TextString = lbAttach.List(i, 0) & "=" & strLine
        iAtt = iAtt + 1
        
Next_I:
    Next i
    
    If Not tbCable.Value = "" Then
        strUnits = Replace(tbCable.Value, vbLf, "")
        
        vAttList(25).TextString = strUnits
    End If
    
    If Not tbSplices.Value = "" Then
        strUnits = Replace(tbSplices.Value, vbLf, "")
        
        vAttList(26).TextString = strUnits
    End If
    
    If lbUnits.ListCount > 0 Then
        strUnits = lbUnits.List(0)
        If lbUnits.ListCount > 1 Then
            For i = 1 To lbUnits.ListCount - 1
                strUnits = strUnits & ";;" & lbUnits.List(i)
            Next i
        End If
        
        vAttList(27).TextString = strUnits
    End If
    
Exit_Sub:
    lbAttach.Clear
    lbUnits.Clear
    
    cbGetsPole.SetFocus
    
    objPole.Update
End Sub

Private Sub Label10_Click()
    cbUpdateAtt.Enabled = True
End Sub

Private Sub lbUnits_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            lbUnits.RemoveItem lbUnits.ListIndex
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbAttach.ColumnCount = 4
    lbAttach.ColumnWidths = "84;36;36;84"
End Sub

Private Sub SortAttachments()
    Dim strArrayList(), strArraySorted(), strData() As String
    Dim strListItem() As String     '<---------------------------------------------Sort
    'Dim attArray, attItem As Variant
    'Dim str1, str2 As String
    'Dim i, iDWGNum, test1, place1, temp1 As Integer
    Dim vLine As Variant
    Dim strTemp As String
    Dim iFeet, iInch As Integer
    'Dim tempFt, tempIn As Integer
    'Dim iNewFt, iNewIn As Integer
    Dim i As Integer
    
    'test1 = 0
    
    ReDim strData(0 To lbAttach.ListCount - 1)
    ReDim strListItem(0 To lbAttach.ListCount - 1)
    ReDim strArraySorted(0 To lbAttach.ListCount - 1)
    
    For i = 0 To UBound(strListItem)
        strListItem(i) = i
    Next i
    
    For i = 0 To UBound(strData)
        If lbAttach.List(i, 2) = "" Then
            vLine = Split(lbAttach.List(i, 1), "-")
        Else
            vLine = Split(lbAttach.List(i, 2), "-")
        End If
        
        iFeet = CInt(vLine(0))
        iInch = CInt(vLine(1)) + iFeet * 12
        
        strData(i) = iInch
    Next i
    
    'test1 = UBound(strData) - 1

    For i = UBound(strData) To (LBound(strData) + 1) Step -1
        For j = LBound(strData) To (i - 1)
            If CInt(strData(j)) < CInt(strData(j + 1)) Then
                strTemp = strData(j + 1)
                strData(j + 1) = strData(j)
                strData(j) = strTemp
                
                strTemp = strListItem(j + 1)
                strListItem(j + 1) = strListItem(j)
                strListItem(j) = strTemp
            End If
        Next j
    Next i
    
    For i = LBound(strListItem) To UBound(strListItem)
        strArraySorted(i) = lbAttach.List(strListItem(i), 0) & vbTab & lbAttach.List(strListItem(i), 1) & vbTab & lbAttach.List(strListItem(i), 2) & vbTab & lbAttach.List(strListItem(i), 3)
    Next i
        
    lbAttach.Clear
    For i = LBound(strListItem) To UBound(strListItem)
        vLine = Split(strArraySorted(i), vbTab)
        
        lbAttach.AddItem vLine(0)
        lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(1)
        lbAttach.List(lbAttach.ListCount - 1, 2) = vLine(2)
        lbAttach.List(lbAttach.ListCount - 1, 3) = vLine(3)
    Next i
End Sub

Private Sub ClearForm()
    tbSPoleNum.Value = ""
    tbIPoleNum.Value = ""
    cbPoleNum.Value = False
    
    tbsStatus.Value = ""
    tbiStatus.Value = ""
    cbStatus.Value = False
    
    tbsOwner.Value = ""
    tbiOwner.Value = ""
    cbOwner.Value = False
    
    tbsOwnerNum.Value = ""
    tbiOwnerNum.Value = ""
    cbOwnerNum.Value = False
    
    tbsOtherNum.Value = ""
    tbiOtherNum.Value = ""
    cbOtherNum.Value = False
    
    tbsHC.Value = ""
    tbiHC.Value = ""
    cbHC.Value = False
    
    tbsYear.Value = ""
    tbiYear.Value = ""
    cbYear.Value = False
    
    tbsLL.Value = ""
    tbiLL.Value = ""
    cbLL.Value = False
    
    tbsGRD.Value = ""
    tbiGRD.Value = ""
    cbGRD.Value = False
    
    lbAttach.Clear
    lbUnits.Clear
    
    tbCable.Value = ""
    tbSplices.Value = ""
End Sub

Private Sub ClearAttributes()
    Dim vAttList As Variant
    
    vAttList = objPole.GetAttributes
    
    For i = 9 To 23
        vAttList(i).TextString = ""
    Next i
    
    objPole.Update
End Sub
