VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlacePoleData 
   Caption         =   "Place Structure Data"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   OleObjectBlob   =   "PlacePoleData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlacePoleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objPole As AcadBlockReference
Dim iUpdate As Integer

Private Sub cbAddAttach_Click()
    Dim strMR As String
    
    If Not tbEAttach.Value = "" Then
        If InStr(tbEAttach.Value, "-") = 0 Then tbEAttach.Value = tbEAttach.Value & "-0"
    End If
    If Not tbPAttach.Value = "" Then
        If InStr(tbPAttach.Value, "-") = 0 Then tbPAttach.Value = tbPAttach.Value & "-0"
    End If
    
    If cbAddAttach.Caption = "Update" Then
        lbAttach.List(lbAttach.ListIndex, 0) = cbAttachment.Value
        lbAttach.List(lbAttach.ListIndex, 1) = tbEAttach.Value
        lbAttach.List(lbAttach.ListIndex, 2) = tbPAttach.Value
        
        cbAddAttach.Caption = "Add Data"
        lbAttach.Enabled = True
        GoTo Exit_Sub
    End If
    
    lbAttach.AddItem cbAttachment.Value
    lbAttach.List(lbAttach.ListCount - 1, 1) = tbEAttach.Value
    lbAttach.List(lbAttach.ListCount - 1, 2) = tbPAttach.Value
    
    lbAttach.ListIndex = lbAttach.ListCount - 1
    
Exit_Sub:
    If tbPAttach.Value = "" Then
        strMR = ""
    Else
        strMR = GetMR(CStr(tbEAttach.Value), CStr(tbPAttach.Value))
    End If
    
    lbAttach.List(lbAttach.ListIndex, 3) = strMR
    
    cbAttachment.Value = ""
    tbEAttach.Value = ""
    tbPAttach.Value = ""
    
    Call SortAttachments
End Sub

Private Sub cbAddData_Click()
    If cbAddData.Caption = "Update" Then
        lbData.List(lbData.ListIndex, 1) = tbDValue.Value
        
        cbDType.Value = ""
        cbDType.Enabled = True
        tbDValue.Value = ""
        
        cbAddData.Caption = "Add Data"
        iUpdate = 1
        Exit Sub
    End If
    
    Select Case cbDType.Value
        Case "Type"
            lbData.AddItem "1  Type"
        Case "Owner #"
            lbData.AddItem "2  Owner #"
        Case "Other #"
            lbData.AddItem "3  Other #"
        Case "Ground"
            lbData.AddItem "9  Ground"
    End Select
    lbData.List(lbData.ListCount - 1, 1) = tbDValue.Value
    
    cbDType.Value = ""
    tbDValue.Value = ""
    iUpdate = 1
End Sub

Private Sub cbAddUnit_Click()
    If cbAddUnit.Caption = "Update" Then
        lbUnits.List(lbUnits.ListIndex, 0) = tbUnit.Value
        lbUnits.List(lbUnits.ListIndex, 1) = tbQuantity.Value
        lbUnits.List(lbUnits.ListIndex, 2) = tbUNote.Value
        
        lbUnits.Enabled = True
        cbAddUnit.Caption = "Add Unit"
        GoTo Exit_Sub
    End If
    
    lbUnits.AddItem tbUnit.Value
    lbUnits.List(lbUnits.ListCount - 1, 1) = tbQuantity.Value
    lbUnits.List(lbUnits.ListCount - 1, 2) = tbUNote.Value
    
Exit_Sub:
    
    tbUnit.Value = ""
    tbQuantity.Value = ""
    tbUNote.Value = ""
    iUpdate = 1
End Sub

Private Sub cbCombine_Click()
    If lbUnits.ListCount < 2 Then Exit Sub
    
    Dim vCurrent, vOther, vItem As Variant
    Dim strCurrent As String
    
    For i = lbUnits.ListCount - 1 To 0 Step -1
        strCurrent = lbUnits.List(i, 0)
        If lbUnits.List(i, 1) = "" Then lbUnits.List(i, 1) = 1
        
        For j = 0 To i - 1
            If lbUnits.List(j, 0) = strCurrent Then
                If lbUnits.List(j, 1) = "" Then lbUnits.List(j, 1) = 1
                lbUnits.List(j, 1) = CInt(lbUnits.List(j, 1)) + CInt(lbUnits.List(i, 1))
                lbUnits.RemoveItem i
                GoTo Next_I
            End If
        Next j
Next_I:
    Next i
End Sub

Private Sub cbDeleteCallout_Click()
    Dim objDelete As AcadSelectionSet
    
    On Error Resume Next
    Me.Hide
    
    Set objDelete = ThisDrawing.SelectionSets.Add("objDelete")
    objDelete.SelectOnScreen
    
    objDelete.Erase
    objDelete.Delete
    
    Me.show
End Sub

Private Sub cbGetsPole_Click()
    If iUpdate = 1 Then
        Dim result As Integer
        
        result = MsgBox("Save changes to Structure?", vbYesNo, "Save Changes")
        If result = vbYes Then
            If objPole.Name = "sPole" Then
                Call SaveAttachments
            Else
                Call SaveBuried
            End If
    
            MsgBox objPole.Name & "  Updated."
        End If
        
        iUpdate = 0
    End If
    
    Dim objObject As AcadObject
    'Dim objBlock As AcadBlockReference
    Dim vBasePnt, vAttList As Variant
    Dim vLine, vItem, vTemp As Variant
    Dim strTemp As String
    Dim dTemp As Double
    
    lbData.Clear
    lbAttach.Clear
    lbUnits.Clear
    tbNotes.Value = ""
    'lbCables.Clear
    'lbSplices.Clear
    
    'cbUpdatePole.Enabled = True
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objObject, vBasePnt, "Select Pole: "
    If TypeOf objObject Is AcadBlockReference Then
        Set objPole = objObject
    Else
        MsgBox "Not a valid object."
        Me.show
        Exit Sub
    End If
    
    Select Case objPole.Name
        Case "sPole"
        Case "sPed", "sHH", "sMH", "sPanel", "sFP"
            Call GetBuriedPlant
            Me.show
            Exit Sub
        Case Else
            MsgBox "Not a valid pole."
            Me.show
            Exit Sub
    End Select
    
    dTemp = GetScale(objPole.InsertionPoint)
    If dTemp > 0 Then cbScale.Value = dTemp * 100
    
    vAttList = objPole.GetAttributes
    
    tbPoleNumber.Value = vAttList(0).TextString
    tbOwner.Value = vAttList(2).TextString
    
    lbData.AddItem "1  Type"
    If vAttList(5).TextString = "" Then
        lbData.List(lbData.ListCount - 1, 1) = "?-?"
    Else
        lbData.List(lbData.ListCount - 1, 1) = vAttList(5).TextString
    End If
    
    
    
    Select Case vAttList(3).TextString
        Case ""
            lbData.AddItem "2  Owner #"
            lbData.List(lbData.ListCount - 1, 1) = "NA"
        Case "NO TAG"
            lbData.AddItem "2  Owner #"
            lbData.List(lbData.ListCount - 1, 1) = "NO TAG"
        Case Else
            vLine = Split(vAttList(3).TextString, " ")
            For j = 0 To UBound(vLine)
                If j = 0 Then
                    lbData.AddItem "2  Owner #"
                    lbData.List(lbData.ListCount - 1, 1) = vLine(j)
                Else
                    lbData.AddItem "3  Other #"
                    lbData.List(lbData.ListCount - 1, 1) = vLine(j)
                End If
            Next j
    End Select
    
    'If vAttList(3).TextString = "" Then
        'lbData.AddItem "2  Owner #"
        'lbData.List(lbData.ListCount - 1, 1) = "NA"
    'Else
        'vLine = Split(vAttList(3).TextString, " ")
        'For j = 0 To UBound(vLine)
            'If j = 0 Then
                'lbData.AddItem "2  Owner #"
                'lbData.List(lbData.ListCount - 1, 1) = vLine(j)
            'Else
                'lbData.AddItem "3  Other #"
                'lbData.List(lbData.ListCount - 1, 1) = vLine(j)
            'End If
        'Next j
    'End If
    
    
    
    If Not vAttList(4).TextString = "" Then
        vLine = Split(vAttList(4).TextString, " ")
        
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), "=")
            
            If UBound(vItem) = 0 Then
                lbData.AddItem "3  Other #"
                lbData.List(lbData.ListCount - 1, 1) = tbOwner.Value & "=" & vLine(i)
            Else
                For j = 1 To UBound(vItem)
                    Select Case vItem(0)
                        Case "A"
                            vItem(0) = "ATT"
                        Case "CH"
                            vItem(0) = "CHARTER"
                        Case "CO"
                            vItem(0) = "COMCAST"
                        Case "D"
                            vItem(0) = "DREMC"
                        Case "M"
                            vItem(0) = "MTEMC"
                        Case "N"
                            vItem(0) = "NES"
                        Case "T"
                            vItem(0) = "TDS"
                    End Select
                    
                    lbData.AddItem "3  Other #"
                    lbData.List(lbData.ListCount - 1, 1) = vItem(0) & "=" & vItem(j)
                Next j
            End If
        Next i
    End If
    
    
    lbData.AddItem "9  Ground"
    Select Case UCase(vAttList(8).TextString)
        Case "", "N"
            lbData.List(lbData.ListCount - 1, 1) = "NO GRD"
        Case "M"
            lbData.List(lbData.ListCount - 1, 1) = "MGNV"
        Case "T"
            lbData.List(lbData.ListCount - 1, 1) = "TGB"
        Case "B"
            lbData.List(lbData.ListCount - 1, 1) = "BROKEN GRD"
        Case Else
            lbData.List(lbData.ListCount - 1, 1) = vAttList(8).TextString
    End Select
    
    If cbAddCoords.Value = True Then
        If Not vAttList(7).TextString = "" Then
            vLine = Split(vAttList(7).TextString, ",")
            
            dTemp = CLng(CDbl(vLine(0)) * 10000000) / 10000000
            strTemp = dTemp & ","
            
            dTemp = CLng(CDbl(vLine(1)) * 10000000) / 10000000
            strTemp = strTemp & dTemp
            
            lbData.AddItem "99  Lat,Long"
            lbData.List(lbData.ListCount - 1, 1) = strTemp
        End If
    End If
    
    '<---------------------------------------------------------------- Add Attachments
    
    For n = 9 To 23
        If vAttList(n).TextString = "" Then GoTo Next_N
        
        
        Select Case n
            Case Is = 9
                vLine = Split(vAttList(n).TextString, " ")
                For k = 0 To UBound(vLine)
                    lbAttach.AddItem "NEUTRAL"
                    
                    If InStr(vLine(k), ")") > 0 Then
                        vTemp = Split(vLine(k), ")")
                        vTemp(0) = Replace(vTemp(0), "(", "")
                        
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vTemp(0)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = vTemp(1)
                        lbAttach.List(lbAttach.ListCount - 1, 3) = GetMR(CStr(vTemp(0)), CStr(vTemp(1)))
                    Else
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(k)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                        lbAttach.List(lbAttach.ListCount - 1, 3) = ""
                    End If
                Next k
            Case Is = 10
                vLine = Split(vAttList(n).TextString, " ")
                For k = 0 To UBound(vLine)
                    lbAttach.AddItem "TRANSFORMER"
                    
                    If InStr(vLine(k), ")") > 0 Then
                        vTemp = Split(vLine(k), ")")
                        vTemp(0) = Replace(vTemp(0), "(", "")
                        
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vTemp(0)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = vTemp(1)
                        lbAttach.List(lbAttach.ListCount - 1, 3) = GetMR(CStr(vTemp(0)), CStr(vTemp(1)))
                    Else
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(k)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                        lbAttach.List(lbAttach.ListCount - 1, 3) = ""
                    End If
                Next k
            Case Is = 11, Is = 12, Is = 13
                vLine = Split(vAttList(n).TextString, " ")
                For k = 0 To UBound(vLine)
                    
                    Select Case n
                        Case Is = 11
                            lbAttach.AddItem "LOW POWER"
                        Case Is = 12
                            lbAttach.AddItem "ANTENNA"
                        Case Is = 13
                            lbAttach.AddItem "ST LT CIR"
                    End Select
                    
                    If InStr(vLine(k), ")") > 0 Then
                        vTemp = Split(vLine(k), ")")
                        vTemp(0) = Replace(vTemp(0), "(", "")
                        
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vTemp(0)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = vTemp(1)
                        lbAttach.List(lbAttach.ListCount - 1, 3) = GetMR(CStr(vTemp(0)), CStr(vTemp(1)))
                    Else
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(k)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                        lbAttach.List(lbAttach.ListCount - 1, 3) = ""
                    End If
                Next k
            Case Is = 14
                vLine = Split(vAttList(n).TextString, " ")
                For k = 0 To UBound(vLine)
                    lbAttach.AddItem "ST LT"
                    
                    If InStr(vLine(k), ")") > 0 Then
                        vTemp = Split(vLine(k), ")")
                        vTemp(0) = Replace(vTemp(0), "(", "")
                        
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vTemp(0)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = vTemp(1)
                        lbAttach.List(lbAttach.ListCount - 1, 3) = GetMR(CStr(vTemp(0)), CStr(vTemp(1)))
                    Else
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(k)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                        lbAttach.List(lbAttach.ListCount - 1, 3) = ""
                    End If
                Next k
            Case Is = 15
                vLine = Split(UCase(vAttList(n).TextString), " ")
                For k = 0 To UBound(vLine)
                    lbAttach.AddItem "NEW 6M"
                    lbAttach.List(lbAttach.ListCount - 1, 1) = ""
                    lbAttach.List(lbAttach.ListCount - 1, 3) = "NEW"
                    
                    If InStr(vLine(k), "T") > 0 Then
                        vLine(k) = Replace(vLine(k), "T", "")
                        lbAttach.List(lbAttach.ListCount - 1, 3) = "MTE TAG"
                    End If
                    
                    If InStr(vLine(k), "O") > 0 Then
                        vLine(k) = Replace(vLine(k), "O", "")
                        lbAttach.List(lbAttach.ListCount - 1, 0) = "NEW 6M OHG"
                    End If
                    
                    If InStr(vLine(k), "P") > 0 Then
                        vLine(k) = Replace(vLine(k), "P", "")
                        lbAttach.List(lbAttach.ListCount - 1, 0) = "NEW 6M TAP"
                    End If
                    
                    If InStr(vLine(k), "F") > 0 Then
                        vLine(k) = Replace(vLine(k), "F", "")
                        lbAttach.List(lbAttach.ListCount - 1, 3) = "FUTURE"
                    End If
                    
                    lbAttach.List(lbAttach.ListCount - 1, 2) = vLine(k)
                Next k
            Case Is > 15
                vLine = Split(UCase(vAttList(n).TextString), "=")
                vItem = Split(vLine(1), " ")
                For k = 0 To UBound(vItem)
                    strTemp = vLine(0)
                    
                    If InStr(vItem(k), "C") > 0 Then
                        strTemp = strTemp & " C-WIRE"
                        vItem(k) = Replace(vItem(k), "C", "")
                    End If
                    
                    If InStr(vItem(k), "D") > 0 Then
                        strTemp = strTemp & " DROP"
                        vItem(k) = Replace(vItem(k), "D", "")
                    End If
                    
                    If InStr(vItem(k), "O") > 0 Then
                        strTemp = strTemp & " OHG"
                        vItem(k) = Replace(vItem(k), "O", "")
                    End If
                    
                    If InStr(vItem(k), "P") > 0 Then
                        strTemp = strTemp & " TAP"
                        vItem(k) = Replace(vItem(k), "P", "")
                    End If
                    
                    If InStr(vItem(k), "S") > 0 Then
                        strTemp = strTemp & " SS"
                        vItem(k) = Replace(vItem(k), "S", "")
                    End If
                    
                    If InStr(vItem(k), "V") > 0 Then
                        strTemp = "LASH TO " & strTemp
                        vItem(k) = Replace(vItem(k), "V", "")
                    End If
                    
                    lbAttach.AddItem strTemp
                    
                    lbAttach.List(lbAttach.ListCount - 1, 3) = ""
                    
                    If InStr(vItem(k), "X") > 0 Then
                        lbAttach.List(lbAttach.ListCount - 1, 3) = "ATTACH"
                        vItem(k) = Replace(vItem(k), "X", "")
                        lbAttach.List(lbAttach.ListCount - 1, 1) = ""
                        lbAttach.List(lbAttach.ListCount - 1, 2) = vItem(k)
                        
                        GoTo Next_K
                    End If
                    
                    If InStr(vItem(k), "E") > 0 Then
                        lbAttach.List(lbAttach.ListCount - 1, 3) = "EXTEND"
                        vItem(k) = Replace(vItem(k), "E", "")
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vItem(k)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                        
                        GoTo Next_K
                    End If
                    
                    If InStr(vItem(k), ")") > 0 Then
                        vTemp = Split(vItem(k), ")")
                        vTemp(0) = Replace(vTemp(0), "(", "")
                        
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vTemp(0)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = vTemp(1)
                        lbAttach.List(lbAttach.ListCount - 1, 3) = GetMR(CStr(vTemp(0)), CStr(vTemp(1)))
                    Else
                        lbAttach.List(lbAttach.ListCount - 1, 1) = vItem(k)
                        lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                    End If
Next_K:
                Next k
        End Select
Next_N:
    Next n
    
    'att 25
    'If Not vAttList(25).TextString = "" Then
        'vAttList(25).TextString = Replace(vAttList(25).TextString, vbLf, "")
        'vLine = Split(vAttList(25).TextString, vbCr)
        'For i = 0 To UBound(vLine)
            'vItem = Split(vLine(i), " / ")
            'lbCables.AddItem vItem(0)
            'lbCables.List(lbCables.ListCount - 1, 1) = vItem(1)
        'Next i
    'End If
    
    'att 26
    'If Not vAttList(26).TextString = "" Then
        'vAttList(26).TextString = Replace(vAttList(26).TextString, vbLf, "")
        ''If InStr(vAttList(26).TextString, " + ") > 0 Then vAttList(26).TextString = Replace(vAttList(26).TextString, " + ", vbCr)
        
        'vLine = Split(vAttList(26).TextString, vbCr)
        'For i = 0 To UBound(vLine)
            'lbSplices.AddItem vLine(i)
        'Next i
    'End If
    
    'att 24
    If Not vAttList(24).TextString = "" Then
        tbNotes.Value = ""
        
        vAttList(24).TextString = Replace(vAttList(24).TextString, vbLf, "")
        vAttList(24).TextString = Replace(vAttList(24).TextString, vbCr, "")
        vLine = Split(vAttList(24).TextString, ";")
        
        tbNotes.Value = vLine(0)
        If UBound(vLine) > 0 Then
            For i = 1 To UBound(vLine)
                If Not vLine(i) = "" Then tbNotes.Value = tbNotes.Value & vbCr & vLine(i)
            Next i
        End If
    End If
    
    'att 27
    If Not vAttList(27).TextString = "" Then
        vAttList(27).TextString = Replace(vAttList(27).TextString, vbLf, "")
        vLine = Split(vAttList(27).TextString, ";;")
        For i = 0 To UBound(vLine)
            If InStr(vLine(i), "=") = 0 Then vLine(i) = vLine(i) & "=1"
                
            vItem = Split(vLine(i), "=")
            lbUnits.AddItem vItem(0)
            If InStr(vItem(1), "  ") > 0 Then
                vTemp = Split(vItem(1), "  ")
                lbUnits.List(lbUnits.ListCount - 1, 1) = vTemp(0)
                lbUnits.List(lbUnits.ListCount - 1, 2) = vTemp(1)
            Else
                lbUnits.List(lbUnits.ListCount - 1, 1) = vItem(1)
                lbUnits.List(lbUnits.ListCount - 1, 2) = ""
            End If
        Next i
    End If
    
    Call SortAttachments
    
    cbUpdate.Enabled = True
    cbPlaceData.Enabled = True
    
    cbPlaceData.SetFocus
    
    Me.show
End Sub

Private Sub cbPlaceData_Click()
    'If lbWorkspace.ListCount = 0 Then GoTo Exit_Sub
    If tbPoleNumber.Value = "" Then GoTo Exit_Sub
    
    If Not objPole.Name = "sPole" Then
        Call PlaceBuriedData
        Exit Sub
    End If
    
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
    
  On Error Resume Next
    
    strAtt0 = tbPoleNumber.Value
    
    iMR = 0
    iNote = 0
    dPosition = 1#
    
    Me.Hide
    dScale = CInt(cbScale.Value) / 100
    
    Err = 0
    returnPoint = ThisDrawing.Utility.GetPoint(, "Select point:")
    If Not Err = 0 Then GoTo Exit_Sub
    
    n = 0
    For Each Item In returnPoint
        insertionPnt(n) = Item
        n = n + 1
    Next Item
    
    If lbAttach.ListCount < 1 Then GoTo Place_Info
    If cbAttachments.Value = False Then GoTo Place_Info
    
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
    If lbAttach.ListCount < 1 Then GoTo Place_Info
    
    For i = 0 To (lbAttach.ListCount - 1)
        dPosition = dPosition + 0.01
        strAtt1 = dPosition
        strAtt2 = lbAttach.List(i, 0)
        iPI = 0: iEI = 0
        
        If lbAttach.List(i, 2) = "" Then
            strAtt3 = lbAttach.List(i, 1)
            If lbAttach.List(i, 3) = "EXTEND" Then
                strLayer = "Integrity Attachments-MR"
                iNote = 2
            Else
                strLayer = "Integrity Attachments-Existing"
            End If
        Else
            strAtt3 = lbAttach.List(i, 2)
            If InStr(lbAttach.List(i, 0), "NEW") > 0 Then
                strLayer = "Integrity Attachments-New"
                If iNote < 1 Then iNote = 1
                If lbAttach.List(i, 3) = "MTE TAG" Then iNote = 4
            Else
                strLayer = "Integrity Attachments-MR"
                iNote = 2
            End If
        End If
        
        strAtt3 = Replace(UCase(strAtt3), "-", "'") & """"
        If lbAttach.List(i, 3) = "" Then
            strAtt4 = ""
        Else
            If lbAttach.List(i, 3) = Null Then
                strAtt4 = ""
            Else
                strAtt4 = lbAttach.List(i, 3)
            End If
        End If
        
        
        If InStr(lbAttach.List(i, 0), "NEW") > 0 Then
            strLayer = "Integrity Attachments-New"
            If iNote < 1 Then iNote = 1
        End If
        
Place_Attachment:
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        objBlock.Layer = strLayer
        
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = strAtt0
        vAttList(1).TextString = strAtt1
        vAttList(2).TextString = strAtt2
        vAttList(3).TextString = strAtt3
        If strAtt4 = "" Then
            vAttList(4).TextString = ""
        Else
            vAttList(4).TextString = strAtt4
        End If
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
        Case 4
            Set layerObj = ThisDrawing.Layers.Add("Integrity Attachments-New")
            ThisDrawing.ActiveLayer = layerObj
            dRevCloud(3) = insertionPnt(0) + (84 * dScale)
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
    
    str = "pole_info"
    strLayer = "Integrity Pole-Info"

    insertionPnt(0) = insertionPnt(0) - (74 * dScale)
    insertionPnt(1) = insertionPnt(1) - (12 * dScale)
    
    dPosition = 0#
    
    
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
    objBlock.Layer = strLayer
        
    vAttList = objBlock.GetAttributes
    vAttList(0).TextString = strAtt0
    vAttList(1).TextString = "0.0"
    vAttList(2).TextString = tbPoleNumber.Value
    objBlock.Update
    
    insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    
    
    Dim vText, vLine As Variant
    
    For w = 0 To lbData.ListCount - 1
        If lbData.List(w, 0) = "" Then GoTo Next_W
        
        vText = Split(lbData.List(w, 0), " ")
        strAtt1 = dPosition & "." & vText(0)
        
        Select Case vText(0)
            Case "1", "2"
                strAtt2 = tbOwner.Value
                strAtt3 = lbData.List(w, 1)
            Case "3"
                vLine = Split(lbData.List(w, 1), "=")
                
                If UBound(vLine) > 0 Then
                    strAtt2 = vLine(0)
                    strAtt3 = vLine(1)
                Else
                    strAtt2 = tbOwner.Value
                    strAtt3 = lbData.List(w, 1)
                End If
            Case "9", "99"
                strAtt2 = lbData.List(w, 1)
                strAtt3 = ""
            Case Else
                GoTo Next_W
        End Select
    
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        objBlock.Layer = strLayer
        If vText(0) = "1" Then
            If InStr(lbData.List(w, 1), ")") > 0 Then objBlock.Layer = "Integrity Notes"
        End If
        
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = strAtt0
        vAttList(1).TextString = strAtt1
        vAttList(2).TextString = strAtt2
        vAttList(3).TextString = strAtt3
        objBlock.Update
    
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
Next_W:
    Next w
    
    lwpPnt(0) = insertionPnt(0) - (4 * dScale)
    lwpPnt(1) = insertionPnt(1) + (7 * dScale)
    lwpPnt(2) = lwpPnt(0) + (100 * dScale)
    lwpPnt(3) = lwpPnt(1)
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpPnt)
    lineObj.Layer = strLayer
    lineObj.Update
    
    '<---------------------------------------------------------------------------------------------
Place_Units:
    If cbPlaceUnits.Value = False Then GoTo Place_Note
    If lbUnits.ListCount < 1 Then GoTo Place_Note
    
    str = "pole_unit"
    strLayer = "Integrity Pole-Units"
    
    dPosition = 2#
    
    insertionPnt(1) = insertionPnt(1) - (1 * dScale)
    
    For v = 0 To lbUnits.ListCount - 1
        strAtt1 = dPosition
        strAtt2 = "N/A"
        strAtt3 = lbUnits.List(v, 0) & "=" & lbUnits.List(v, 1)
        If Not lbUnits.List(v, 2) = "" Then strAtt3 = strAtt3 & "  " & lbUnits.List(v, 2)
    
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        objBlock.Layer = strLayer
        
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = strAtt0
        vAttList(1).TextString = strAtt1
        vAttList(2).TextString = strAtt2
        vAttList(3).TextString = strAtt3
        objBlock.Update
    
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    Next v
    
    'GoTo Exit_Sub
    
Place_Note:
    If cbPlaceNotes.Value = False Then GoTo Exit_Sub
    If tbNotes.Value = "" Then GoTo Exit_Sub
    
    Dim objCircle As AcadCircle
    Dim dRadius As Double
    Dim iOffset As Integer
    'Dim dNote(0 To 2) As Double
    
    dRadius = 20#
    iOffset = 0
    
    str = Replace(tbNotes.Value, vbLf, "")
    vStr = Split(str, vbCr)
    
    dNote(0) = (dRevCloud(0) + dRevCloud(3)) / 2
    dNote(1) = dRevCloud(1) - 20
    dNote(2) = 0#
    
    For i = 0 To UBound(vStr)
        Select Case vStr(i)
            Case "OCALC"
                Set objCircle = ThisDrawing.ModelSpace.AddCircle(objPole.InsertionPoint, dRadius)
                objCircle.Layer = "Integrity Notes"
                objCircle.Update
                
                GoTo Next_I
            Case "NOTE-TALL"
                str = "Notes-TallerPole5"
                iOffset = 30
            Case "NOTE-TALL10"
                str = "Notes-TALL10"
                iOffset = 30
            Case "NOTE-DEF"
                str = "Notes-DefectivePole"
                iOffset = 30
            Case "NOTE-PRA"
                str = "Notes-PRA"
                iOffset = 30
            Case "NOTE-UEH"
                str = "Notes-UEH"
                iOffset = 30
            Case "NOTE-EC"
                str = "Notes-EC"
                iOffset = 40
            Case "NOTE-OTHER"
                str = "Notes-Other"
                iOffset = 30
            Case Else
                str = ""
        End Select
        
        If str = "" Then GoTo Next_I
        
        dNote(1) = dNote(1) + iOffset
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dNote, str, dScale, dScale, dScale, 0#)
        objBlock.Layer = "Integrity Notes"
        objBlock.Update
        
Next_I:
    Next i
    
Exit_Sub:
    cbGetsPole.SetFocus
    Me.show
End Sub

Private Sub cbQuit_Click()
    If iUpdate = 1 Then
        Dim result As Integer
        
        result = MsgBox("Save changes to Structure?", vbYesNo, "Save Changes")
        If result = vbYes Then
            If objPole.Name = "sPole" Then
                Call SaveAttachments
            Else
                Call SaveBuried
            End If
    
            MsgBox objPole.Name & "  Updated."
        End If
        
        iUpdate = 0
    End If
    
    Me.Hide
End Sub

Private Sub cbUpdate_Click()
    If objPole.Name = "sPole" Then
        Call SaveAttachments
    Else
        Call SaveBuried
    End If
    
    MsgBox objPole.Name & "  Updated."
    iUpdate = 0
End Sub

Private Sub lbAttach_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbAttach.ListCount < 0 Then Exit Sub
    
    Dim iIndex As Integer
    
    iIndex = lbAttach.ListIndex
    If iIndex < 0 Then Exit Sub
    
    cbAttachment.Value = lbAttach.List(iIndex, 0)
    tbEAttach.Value = lbAttach.List(iIndex, 1)
    tbPAttach.Value = lbAttach.List(iIndex, 2)
    
    lbAttach.Enabled = False
    cbAddAttach.Caption = "Update"
    
    iUpdate = 1
End Sub

Private Sub lbAttach_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbAttach.ListIndex < 0 Then Exit Sub
    
    Dim iIndex As Integer
    
    iIndex = lbAttach.ListIndex
    
    Select Case KeyCode
        Case vbKeyDelete
            cbAttachment.Value = lbAttach.List(iIndex, 0)
            tbEAttach.Value = lbAttach.List(iIndex, 1)
            tbPAttach.Value = lbAttach.List(iIndex, 2)
            
            lbAttach.RemoveItem iIndex
            iUpdate = 1
    End Select
End Sub

Private Sub lbData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vLine As Variant
    
    vLine = Split(lbData.List(lbData.ListIndex, 0), "  ")
    cbDType.Value = vLine(1)
    tbDValue.Value = lbData.List(lbData.ListIndex, 1)
    
    cbAddData.Caption = "Update"
End Sub

Private Sub lbData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbData.ListCount < 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            lbData.RemoveItem lbData.ListIndex
        End Select
End Sub

Private Sub lbUnits_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbUnit.Value = lbUnits.List(lbUnits.ListIndex, 0)
    tbQuantity.Value = lbUnits.List(lbUnits.ListIndex, 1)
    tbUNote.Value = lbUnits.List(lbUnits.ListIndex, 2)
    
    lbUnits.Enabled = False
    cbAddUnit.Caption = "Update"
End Sub

Private Sub lbUnits_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbUnits.ListCount < 1 Then Exit Sub
    
    Dim iIndex As Integer
    iIndex = lbUnits.ListIndex
    
    Select Case KeyCode
        Case vbKeyDelete
            tbUnit.Value = lbUnits.List(iIndex, 0)
            tbQuantity.Value = lbUnits.List(iIndex, 1)
            tbUNote.Value = lbUnits.List(iIndex, 2)
            
            lbUnits.RemoveItem lbUnits.ListIndex
            iUpdate = 1
    End Select
End Sub

Private Sub UserForm_Initialize()
    iUpdate = 0
    
    lbData.ColumnCount = 2
    lbData.ColumnWidths = "60;180"
    
    lbAttach.ColumnCount = 4
    lbAttach.ColumnWidths = "84;36;36;84"
    
    lbUnits.ColumnCount = 3
    lbUnits.ColumnWidths = "96;36;48"
    
    cbDType.AddItem "Type"
    cbDType.AddItem "Owner #"
    cbDType.AddItem "Other #"
    cbDType.AddItem "Ground"
    
    cbAttachment.AddItem "NEUTRAL"
    cbAttachment.AddItem "TRANSFORMER"
    cbAttachment.AddItem "LOW POWER"
    cbAttachment.AddItem "ANTENNA"
    cbAttachment.AddItem "ST LT CIR"
    cbAttachment.AddItem "ST LT"
    cbAttachment.AddItem "NEW 6M"
    cbAttachment.AddItem "NEW 10M"
    cbAttachment.AddItem "CLEC"
    cbAttachment.AddItem "XO"
    cbAttachment.AddItem "ZAYO"
    cbAttachment.AddItem "LEVEL3"
    cbAttachment.AddItem "ATT"
    cbAttachment.AddItem "TDS"
    cbAttachment.AddItem "PWR OHG"
    cbAttachment.AddItem "NEW OHG"
    cbAttachment.AddItem "CLEC OHG"
    cbAttachment.AddItem "TELCO OHG"
    cbAttachment.AddItem "ATT OHG"
    cbAttachment.AddItem "TDS OHG"
    cbAttachment.AddItem "UTC OHG"
    cbAttachment.AddItem "OHG"
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    cbGetsPole.SetFocus
    
    If InStr(UCase(ThisDrawing.Path), "\NES") > 0 Then cbAttachments.Value = False
End Sub

Private Function GetMR(strE As String, strP As String)
    Dim vLine, vItem As Variant
    Dim strMR As String
    Dim iEF, iEI, iPF, iPI, iMR As Integer
    
    Select Case strE
        Case ""
            If InStr(cbAttachment.Value, "NEW") > 0 Then
                strMR = "NEW"
            Else
                strMR = "ATTACH"
            End If
        Case strP
            strMR = "TRANSFER"
        Case Else
            If strP = "" Then
                strMR = ""
            Else
                vLine = Split(strE, "-")
                iEF = CInt(vLine(0))
                iEI = CInt(vLine(1))
                iEI = iEF * 12 + iEI
            
                vLine = Split(strP, "-")
                iPF = CInt(vLine(0))
                iPI = CInt(vLine(1))
                iPI = iPF * 12 + iPI
            
                iMR = iPI - iEI
            
                If iMR > 0 Then
                    strMR = "RAISE " & iMR & """"
                Else
                    strMR = "LOWER " & Abs(iMR) & """"
                End If
            End If
    End Select
    
    GetMR = strMR
End Function

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

Private Sub SaveBuried()
    Dim vAttList, vLine, vItem As Variant
    Dim strUnits As String
    
    vAttList = objPole.GetAttributes
    
    If lbData.ListCount > -1 Then
        For i = 0 To lbData.ListCount - 1
            'vLine = Split("", "=")
            Select Case lbData.List(i, 0)
                Case "1  Type"
                    vAttList(2).TextString = lbData.List(i, 1)
                Case "9  Ground"
                    vAttList(4).TextString = lbData.List(i, 1)
            End Select
        Next i
    End If
    
    strUnits = ""
    
    If lbUnits.ListCount > 0 Then
        strLine = "+" & lbUnits.List(0, 0) & "=" & lbUnits.List(0, 1)
        
        For i = 0 To lbUnits.ListCount - 1
            strLine = lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
            If Not lbUnits.List(i, 2) = "" Then strLine = strLine & "  " & lbUnits.List(i, 2)
            
            If strUnits = "" Then
                strUnits = strLine
            Else
                strUnits = strUnits & ";;" & strLine
            End If
        Next i
    End If
    
    vAttList(7).TextString = strUnits
    
Exit_Sub:
    objPole.Update

End Sub

Private Sub SaveAttachments()
    If lbAttach.ListCount < 1 Then Exit Sub
    
    Dim strComm(8) As String
    Dim strLine, strExtra, strCO As String
    Dim strUnits As String
    Dim vAttList, vLine, vItem As Variant
    
    For i = 0 To 8
        strComm(i) = ""
    Next i
    
    vAttList = objPole.GetAttributes
    
    For i = 9 To 23
        vAttList(i).TextString = ""
    Next i
    
    For i = 0 To lbAttach.ListCount - 1
        strCO = lbAttach.List(i, 0)
        strLine = lbAttach.List(i, 1)
        If Not lbAttach.List(i, 2) = "" Then strLine = "(" & strLine & ")" & lbAttach.List(i, 2)
                   
        Select Case strCO
            Case "NEUTRAL"
                If vAttList(9).TextString = "" Then
                    vAttList(9).TextString = strLine
                Else
                    vAttList(9).TextString = vAttList(9).TextString & " " & strLine
                End If
            Case "TRANSFORMER"
                If vAttList(10).TextString = "" Then
                    vAttList(10).TextString = strLine
                Else
                    vAttList(10).TextString = vAttList(10).TextString & " " & strLine
                End If
            Case "LOW POWER"
                If vAttList(11).TextString = "" Then
                    vAttList(11).TextString = strLine
                Else
                    vAttList(11).TextString = vAttList(10).TextString & " " & strLine
                End If
            Case "ANTENNA"
                If vAttList(12).TextString = "" Then
                    vAttList(12).TextString = strLine
                Else
                    vAttList(12).TextString = vAttList(10).TextString & " " & strLine
                End If
            Case "ST LT CIRCUIT"
                If vAttList(13).TextString = "" Then
                    vAttList(13).TextString = strLine
                Else
                    vAttList(13).TextString = vAttList(10).TextString & " " & strLine
                End If
            Case "ST LT"
                If vAttList(14).TextString = "" Then
                    vAttList(14).TextString = strLine
                Else
                    vAttList(14).TextString = vAttList(10).TextString & " " & strLine
                End If
            Case Else
                If InStr(strCO, "NEW ") > 0 Then
                    If vAttList(15).TextString = "" Then
                        vAttList(15).TextString = lbAttach.List(i, 2)
                    Else
                        vAttList(15).TextString = vAttList(15).TextString & " " & lbAttach.List(i, 2)
                    End If
                    
                    If lbAttach.List(i, 3) = "MTE TAG" Then
                        vAttList(15).TextString = vAttList(15).TextString & "T"
                    End If
                    If lbAttach.List(i, 0) = "NEW OHG" Then
                        vAttList(15).TextString = vAttList(15).TextString & "O"
                    End If
                    If lbAttach.List(i, 3) = "FUTURE" Then
                        vAttList(15).TextString = vAttList(15).TextString & "F"
                    End If
                Else
                    If InStr(strCO, "C-WIRE") > 0 Then
                        strExtra = "c"
                        strCO = Replace(strCO, " C-WIRE", "")
                    End If
                
                    If InStr(strCO, "DROP") > 0 Then
                        strExtra = "d"
                        strCO = Replace(strCO, " DROP", "")
                    End If
                
                    If lbAttach.List(i, 3) = "EXTEND" Then
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
                
                    If InStr(strCO, " SS") > 0 Then
                        strExtra = "s"
                        strCO = Replace(strCO, " SS", "")
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
                
                    If Not lbAttach.List(i, 1) = "" Then
                        strLine = lbAttach.List(i, 1)
                    
                        If Not lbAttach.List(i, 2) = "" Then
                            strLine = "(" & strLine & ")" & lbAttach.List(i, 2) & strExtra
                        Else
                            strLine = strLine & strExtra
                        End If
                    
                        strExtra = ""
                    Else
                        strLine = lbAttach.List(i, 2) & strExtra & "x"
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
                End If
Found_strComm:
        End Select
    Next i
    
    For i = 1 To 8
        vAttList(15 + i).TextString = strComm(i)
    Next i
    
    strCO = ""
    
    vAttList(2).TextString = tbOwner.Value
    vAttList(3).TextString = ""
    vAttList(4).TextString = ""
    
    
    If lbData.ListCount > -1 Then
        For i = 0 To lbData.ListCount - 1
            vLine = Split("", "=")
            Select Case lbData.List(i, 0)
                Case "1  Type"
                    vAttList(5).TextString = lbData.List(i, 1)
                Case "2  Owner #"
                    If vAttList(3).TextString = "" Then
                        vAttList(3).TextString = lbData.List(i, 1)
                    Else
                        vAttList(3).TextString = vAttList(3).TextString & " " & lbData.List(i, 1)
                    End If
                Case "3  Other #"
                    vLine = Split(lbData.List(i, 1), "=")
                    
                    If UBound(vLine) < 1 Then
                        vAttList(3).TextString = vAttList(3).TextString & " " & lbData.List(i, 1)
                    Else
                        If tbOwner.Value = vLine(0) Then
                            vAttList(3).TextString = vAttList(3).TextString & " " & lbData.List(i, 1)
                        Else
                            If vAttList(4).TextString = "" Then
                                vAttList(4).TextString = lbData.List(i, 1)
                            Else
                                vAttList(4).TextString = vAttList(4).TextString & " " & lbData.List(i, 1)
                            End If
                        End If
                    End If
                Case "9  Ground"
                    vAttList(8).TextString = lbData.List(i, 1)
            End Select
        Next i
    End If
    
    strUnits = ""
    
    If lbUnits.ListCount > 0 Then
        strLine = "+" & lbUnits.List(0, 0) & "=" & lbUnits.List(0, 1)
        
        For i = 0 To lbUnits.ListCount - 1
            strLine = lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
            If Not lbUnits.List(i, 2) = "" Then strLine = strLine & "  " & lbUnits.List(i, 2)
            
            If strUnits = "" Then
                strUnits = strLine
            Else
                strUnits = strUnits & ";;" & strLine
            End If
        Next i
    End If
    
    vAttList(27).TextString = strUnits
    
Exit_Sub:
    objPole.Update
End Sub

Private Sub GetBuriedPlant()
    Dim vAttList, vLine As Variant
    Dim dTemp As Double
    
    dTemp = GetScale(objPole.InsertionPoint)
    If dTemp > 0 Then cbScale.Value = dTemp * 100
    
    vAttList = objPole.GetAttributes
    
    tbPoleNumber.Value = vAttList(0).TextString
    tbOwner.Value = "UTC"
    
    lbData.AddItem "1  Type"
    If vAttList(2).TextString = "" Then
        lbData.List(lbData.ListCount - 1, 1) = "?"
    Else
        lbData.List(lbData.ListCount - 1, 1) = vAttList(2).TextString
    End If
    
    lbData.AddItem "9  Ground"
    Select Case UCase(vAttList(4).TextString)
        Case "", "N"
            lbData.List(lbData.ListCount - 1, 1) = "NO GRD"
        Case "M"
            lbData.List(lbData.ListCount - 1, 1) = "MGN"
        Case "T"
            lbData.List(lbData.ListCount - 1, 1) = "TGB"
        Case "B"
            lbData.List(lbData.ListCount - 1, 1) = "BROKEN GRD"
        Case Else
            lbData.List(lbData.ListCount - 1, 1) = vAttList(4).TextString
    End Select
    
    If cbAddCoords.Value = True Then
        If Not vAttList(3).TextString = "" Then
            vLine = Split(vAttList(3).TextString, ",")
            
            dTemp = CLng(CDbl(vLine(0)) * 10000000) / 10000000
            strTemp = dTemp & ","
            
            dTemp = CLng(CDbl(vLine(1)) * 10000000) / 10000000
            strTemp = strTemp & dTemp
            
            lbData.AddItem "99  Lat,Long"
            lbData.List(lbData.ListCount - 1, 1) = strTemp
        End If
    End If
    
    If Not vAttList(7).TextString = "" Then
        vAttList(7).TextString = Replace(vAttList(7).TextString, vbLf, "")
        vLine = Split(vAttList(7).TextString, ";;")
        For i = 0 To UBound(vLine)
            If InStr(vLine(i), "=") = 0 Then vLine(i) = vLine(i) & "=1"
                
            vItem = Split(vLine(i), "=")
            lbUnits.AddItem vItem(0)
            If InStr(vItem(1), "  ") > 0 Then
                vTemp = Split(vItem(1), "  ")
                lbUnits.List(lbUnits.ListCount - 1, 1) = vTemp(0)
                lbUnits.List(lbUnits.ListCount - 1, 2) = vTemp(1)
            Else
                lbUnits.List(lbUnits.ListCount - 1, 1) = vItem(1)
                lbUnits.List(lbUnits.ListCount - 1, 2) = ""
            End If
        Next i
    End If
    
    cbUpdate.Enabled = True
    cbPlaceData.Enabled = True
    
    cbPlaceData.SetFocus
End Sub

Private Sub PlaceBuriedData()
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
    Dim vText, vLine As Variant
    
    
    strAtt0 = tbPoleNumber.Value
    'dPosition = 1#
    
    Me.Hide
    dScale = CInt(cbScale.Value) / 100
    
    returnPoint = ThisDrawing.Utility.GetPoint(, "Select point:")
    n = 0
    For Each Item In returnPoint
        insertionPnt(n) = Item
        n = n + 1
    Next Item
    
Place_Info:
    
    str = "pole_info"
    strLayer = "Integrity Pole-Info"

    'insertionPnt(0) = insertionPnt(0) - (74 * dScale)
    'insertionPnt(1) = insertionPnt(1) - (12 * dScale)
    
    dPosition = 0#
    
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
    objBlock.Layer = strLayer
        
    vAttList = objBlock.GetAttributes
    vAttList(0).TextString = strAtt0
    vAttList(1).TextString = "0.0"
    vAttList(2).TextString = tbPoleNumber.Value
    objBlock.Update
    
    insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    
    For w = 0 To lbData.ListCount - 1
        If lbData.List(w, 0) = "" Then GoTo Next_W
        
        vText = Split(lbData.List(w, 0), " ")
        strAtt1 = dPosition & "." & vText(0)
        
        Select Case vText(0)
            Case "1", "2"
                strAtt2 = lbData.List(w, 1)
                strAtt3 = ""
            'Case "3"
                'vLine = Split(lbData.List(w, 1), "=")
                
                'If UBound(vLine) > 0 Then
                    'strAtt2 = vLine(0)
                    'strAtt3 = vLine(1)
                'Else
                    'strAtt2 = tbOwner.Value
                    'strAtt3 = lbData.List(w, 1)
                'End If
            Case "9", "99"
                strAtt2 = lbData.List(w, 1)
                strAtt3 = ""
            Case Else
                GoTo Next_W
        End Select
    
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0)
        objBlock.Layer = strLayer
        
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = strAtt0
        vAttList(1).TextString = strAtt1
        vAttList(2).TextString = strAtt2
        'vAttList(3).TextString = strAtt3
        vAttList(3).TextString = ""
        objBlock.Update
    
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
Next_W:
    Next w
    
    lwpPnt(0) = insertionPnt(0) - (4 * dScale)
    lwpPnt(1) = insertionPnt(1) + (7 * dScale)
    lwpPnt(2) = lwpPnt(0) + (100 * dScale)
    lwpPnt(3) = lwpPnt(1)
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpPnt)
    lineObj.Layer = strLayer
    lineObj.Update
    
Place_Units:
    If cbPlaceUnits.Value = False Then GoTo Exit_Sub
    If lbUnits.ListCount < 1 Then GoTo Exit_Sub
    str = "pole_unit"
    strLayer = "Integrity Pole-Units"
    
    dPosition = 2#
    
    insertionPnt(1) = insertionPnt(1) - (1 * dScale)
    
    For v = 0 To lbUnits.ListCount - 1
        strAtt1 = dPosition
        strAtt2 = "N/A"
        strAtt3 = lbUnits.List(v, 0) & "=" & lbUnits.List(v, 1)
        If Not lbUnits.List(v, 2) = "" Then strAtt3 = strAtt3 & "  " & lbUnits.List(v, 2)
    
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, str, dScale, dScale, dScale, 0#)
        objBlock.Layer = strLayer
        
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = strAtt0
        vAttList(1).TextString = strAtt1
        vAttList(2).TextString = strAtt2
        vAttList(3).TextString = strAtt3
        objBlock.Update
    
        insertionPnt(1) = insertionPnt(1) - (9 * dScale)
    Next v
    
Exit_Sub:
    cbGetsPole.SetFocus
    Me.show
End Sub

Private Function GetScale(vCoords As Variant)
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim dMinX, dMaxX, dMinY, dMaxY As Double
    Dim dScale As Double
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        dScale = objBlock.XScaleFactor
        dMinX = objBlock.InsertionPoint(0)
        dMaxX = dMinX + (1652 * dScale)
        
        dMinY = objBlock.InsertionPoint(1)
        dMaxY = dMinY + (1052 * dScale)
        
        If vCoords(0) > dMinX And vCoords(0) < dMaxX Then
            If vCoords(1) > dMinY And vCoords(1) < dMaxY Then
                dScale = dScale * 1.333333333
                GoTo Exit_Sub
            End If
        End If
    Next objBlock
    
    dScale = 0
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    'MsgBox dScale
    GetScale = dScale
End Function
