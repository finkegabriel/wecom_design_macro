VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArcViewPole 
   Caption         =   "Add Info to Pole"
   ClientHeight    =   9121.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11880
   OleObjectBlob   =   "ArcViewPole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArcViewPole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim objPole As AcadBlockReference

Private Sub cbAddAttach_Click()
    Dim strMR As String
    
    If cbAddAttach.Caption = "Update" Then
        lbAttach.List(lbAttach.ListIndex, 0) = cbAttachment.Value
        lbAttach.List(lbAttach.ListIndex, 1) = tbEAttach.Value
        lbAttach.List(lbAttach.ListIndex, 2) = tbPAttach.Value
        
        cbAddAttach.Caption = "Add Data"
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
End Sub

Private Sub cbAddCable_Click()
    Load ArcViewPoleCable
    
        ArcViewPoleCable.show
        
        If Not ArcViewPoleCable.tbCounts.Value = "" Then
            lbCables.AddItem ArcViewPoleCable.tbCable.Value
            lbCables.List(lbCables.ListCount - 1, 1) = Replace(ArcViewPoleCable.tbCounts.Value, vbCr, " + ")
            lbCables.List(lbCables.ListCount - 1, 1) = Replace(lbCables.List(lbCables.ListCount - 1, 1), vbLf, "")
        End If
        
    Unload ArcViewPoleCable
End Sub

Private Sub cbAddData_Click()
    If cbAddData.Caption = "Update" Then
        lbData.List(lbData.ListIndex, 1) = tbDValue.Value
        
        cbDType.Value = ""
        cbDType.Enabled = True
        tbDValue.Value = ""
        
        cbAddData.Caption = "Add Data"
        Exit Sub
    End If
    
    Select Case cbDType.Value
        Case "H-C"
            lbData.AddItem "1  H-C"
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
End Sub

Private Sub cbAddSplice_Click()
    If cbAddSplice.Caption = "Update" Then
        lbSplices.List(lbSplices.ListIndex) = tbSValue.Value
        
        tbSValue.Value = ""
        
        cbAddSplice.Caption = "Add Splice"
        Exit Sub
    End If
    
    If tbSValue.Value = "" Then Exit Sub
    lbSplices.List(lbSplices.ListCount - 1) = tbSValue.Value
    
    tbSValue.Value = ""
End Sub

Private Sub cbAddUnit_Click()
    If cbAddUnit.Caption = "Update" Then
        lbUnits.List(lbUnits.ListIndex, 0) = tbUnit.Value
        lbUnits.List(lbUnits.ListIndex, 1) = tbQuantity.Value
        lbUnits.List(lbUnits.ListIndex, 2) = tbUNote.Value
        
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
End Sub

Private Sub cbGetAttachemnt_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vLine, vItem As Variant
    Dim strTemp, strLine As String
    Dim iFeet, iInch, iTotal, iRL As Integer
    
    Me.Hide
    
  On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        
    objSS.SelectOnScreen
    For Each objEntity In objSS
        If Not TypeOf objEntity Is AcadBlockReference Then GoTo Next_objEntity
        
        Set objBlock = objEntity
        If objBlock.Name = "pole_attach" Then
            vAttList = objBlock.GetAttributes
            
            lbAttach.AddItem vAttList(2).TextString, 0
            
            strTemp = Replace(vAttList(3).TextString, "'", "-")
            strTemp = Replace(strTemp, """", "")
            
            If vAttList(4).TextString = "" Then
                lbAttach.List(0, 1) = strTemp
                lbAttach.List(0, 2) = ""
            Else
                Select Case Left(vAttList(4).TextString, 1)
                    Case "N"
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
            End If
            
            lbAttach.List(0, 3) = vAttList(4).TextString
        End If
        
        
Next_objEntity:
    Next objEntity
    
    objSS.Clear
    objSS.Delete
    
    Call SortAttachments
    
    Me.show
End Sub

Private Sub cbGetCables_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vBasePnt As Variant
    Dim strLine As String
    
    Me.Hide
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Get Cable Callout: "
    
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox "Invalid selection"
        GoTo Exit_Sub
    End If
    
    Set objBlock = objEntity
    
    If Not objBlock.Name = "CableCounts" Then
        MsgBox "Invalid block"
        Me.show
        Exit Sub
    End If
    
    vAttList = objBlock.GetAttributes
    
    lbCables.AddItem vAttList(1).TextString
    lbCables.List(lbCables.ListCount - 1, 1) = Replace(vAttList(0).TextString, "\P", " + ")
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbGetClosures_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vBasePnt As Variant
    Dim vLine As Variant
    Dim strLine As String
    
    Me.Hide
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Get Closure Callout: "
    
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox "Invalid selection"
        GoTo Exit_Sub
    End If
    
    Set objBlock = objEntity
    
    If Not objBlock.Name = "terminal" Then
        MsgBox "Invalid block"
        GoTo Exit_Sub
    End If
    
    vAttList = objBlock.GetAttributes
    'vLine = Split(vAttList(0).TextString, "\P")
    
    'For j = 0 To UBound(vLine)
        'lbSplices.AddItem vLine(j)
    'Next j
    
    strLine = Replace(vAttList(0).TextString, "\P", " + ")
    lbSplices.AddItem strLine
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbGetPole_Click()
    Dim objObject As AcadObject
    'Dim objBlock As AcadBlockReference
    Dim vBasePnt, vAttList As Variant
    Dim vLine, vItem As Variant
    Dim strTemp As String
    
    lbData.Clear
    lbAttach.Clear
    lbUnits.Clear
    lbCables.Clear
    lbSplices.Clear
    
    cbUpdatePole.Enabled = True
    
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
    
    If Not objPole.Name = "iPole" Then
        MsgBox "Not a valid pole."
        Me.show
        Exit Sub
    End If
    
    vAttList = objPole.GetAttributes
    
    tbPoleNumber.Value = vAttList(0).TextString
    tbOwner.Value = vAttList(2).TextString
    
    lbData.AddItem "1  H-C"
    If vAttList(5).TextString = "" Then
        lbData.List(lbData.ListCount - 1, 1) = "?-?"
    Else
        lbData.List(lbData.ListCount - 1, 1) = vAttList(5).TextString
    End If
    
    
    
    If vAttList(3).TextString = "" Then
        lbData.AddItem "2  Owner #"
        lbData.List(lbData.ListCount - 1, 1) = "NA"
    Else
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
    End If
    
    
    
    If Not vAttList(4).TextString = "" Then
        vLine = Split(vAttList(4).TextString, " ")
        
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), "=")
            
            If UBound(vItem) = 0 Then
                lbData.AddItem "3  Other #"
                lbData.List(lbData.ListCount - 1, 1) = "???  " & vLine(i)
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
                    lbData.List(lbData.ListCount - 1, 1) = vItem(0) & "  " & vItem(j)
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
    
    Me.show
    Exit Sub
    '<---------------------------------------------------------------- not complete - Add Attachments
    Dim iFeet, iInch, iMax, iMin As Integer
    Dim str1 As String
    
    iMax = 1200
    iMin = 0
    
    For n = 9 To 26
        If vAttList(n).TextString = "" Then GoTo Next_N
        
        vLine = Split(vAttList(n).TextString, " ")
        
        Select Case n
            Case Is = 9
                For k = 0 To UBound(vLine)
                    lbAttach.AddItem "NEUTRAL"
                    lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(k)
                    
                    vItem = Split(vLine(k), "-")
                    iFeet = CInt(vItem(0))
                    If UBound(vItem) < 1 Then
                        iInch = 0
                    Else
                        iInch = CInt(vItem(1))
                    End If
                    
                    If iMax > (iFeet * 12 + iInch) Then
                        iMax = iFeet * 12 + iInch - CInt(tbNeu.Value)
                        iFeet = Int(iMax / 12)
                        iInch = iMax - (iFeet * 12)
                    
                        tbMaxAtt.Value = iFeet & "-" & iInch
                    End If
                Next k
            Case Is = 10
                For k = 0 To UBound(vLine)
                    lbAttach.AddItem "TRANSFORMER"
                    lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(k)
                    
                    vItem = Split(vLine(k), "-")
                    iFeet = CInt(vItem(0))
                    If UBound(vItem) < 1 Then
                        iInch = 0
                    Else
                        iInch = CInt(vItem(1))
                    End If
                    
                    If iMax > (iFeet * 12 + iInch - CInt(tbTrans.Value)) Then
                        iMax = iFeet * 12 + iInch - CInt(tbTrans.Value)
                        iFeet = Int(iMax / 12)
                        iInch = iMax - (iFeet * 12)
                    
                        tbMaxAtt.Value = iFeet & "-" & iInch
                    End If
                Next k
            Case Is = 11, Is = 12, Is = 13
                For k = 0 To UBound(vLine)
                    Select Case n
                        Case Is = 11
                            lbAttach.AddItem "LOW POWER"
                        Case Is = 11
                            lbAttach.AddItem "ANTENNA"
                        Case Is = 11
                            lbAttach.AddItem "ST LT CIR"
                    End Select
                    
                    lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(k)
                    
                    vItem = Split(vLine(k), "-")
                    iFeet = CInt(vItem(0))
                    If UBound(vItem) < 1 Then
                        iInch = 0
                    Else
                        iInch = CInt(vItem(1))
                    End If
                    
                    If iMax > (iFeet * 12 + iInch - CInt(tbLP.Value)) Then
                        iMax = iFeet * 12 + iInch - CInt(tbLP.Value)
                        iFeet = Int(iMax / 12)
                        iInch = iMax - (iFeet * 12)
                    
                        tbMaxAtt.Value = iFeet & "-" & iInch
                    End If
                Next k
            Case Is = 14
                For k = 0 To UBound(vLine)
                    lbAttach.AddItem "ST LT"
                    lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(k)
                    
                    vItem = Split(vLine(k), "-")
                    iFeet = CInt(vItem(0))
                    If UBound(vItem) < 1 Then
                        iInch = 0
                    Else
                        iInch = CInt(vItem(1))
                    End If
                    
                    If iMax > (iFeet * 12 + iInch - CInt(tbStLt.Value)) Then
                        iMax = iFeet * 12 + iInch - CInt(tbStLt.Value)
                        iFeet = Int(iMax / 12)
                        iInch = iMax - (iFeet * 12)
                    
                        tbMaxAtt.Value = iFeet & "-" & iInch
                    End If
                Next k
            Case Is = 15
                strTemp = "CATV"
                For k = 0 To UBound(vLine)
                    str1 = ""
                    If CInt(Left(vLine(j), 1)) < 1 Then
                        strTemp = vLine(j)
                        GoTo Next_CATV
                    End If
            
                    If Len(vLine(j)) < 3 Then vLine(j) = vLine(j) & "-0"
                    
                    If UCase(Right(vLine(j), 1)) = "X" Then
                        str1 = "ATTACH"
                        vLine(j) = Replace(vItem(j), "X", "")
                    End If
                    
                    If UCase(Right(vLine(j), 1)) = "C" Then
                        strTemp = strTemp & " C-WIRE"
                        vLine(j) = Replace(vItem(j), "C", "")
                    End If
                    
                    lbAttach.AddItem strTemp
                    lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(j)
                    lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                    lbAttach.List(lbAttach.ListCount - 1, 3) = str1
Next_CATV:
                Next k
            Case Is = 16
                strTemp = "ATT"
                For k = 0 To UBound(vLine)
                    str1 = ""
                    If CInt(Left(vLine(j), 1)) < 1 Then
                        strTemp = vLine(j)
                        GoTo Next_ATT
                    End If
            
                    If Len(vLine(j)) < 3 Then vLine(j) = vLine(j) & "-0"
                    
                    If UCase(Right(vLine(j), 1)) = "X" Then
                        str1 = "ATTACH"
                        vLine(j) = Replace(vItem(j), "X", "")
                    End If
                    
                    If UCase(Right(vLine(j), 1)) = "C" Then
                        strTemp = strTemp & " C-WIRE"
                        vLine(j) = Replace(vItem(j), "C", "")
                    End If
                    
                    lbAttach.AddItem strTemp
                    lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(j)
                    lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                    lbAttach.List(lbAttach.ListCount - 1, 3) = str1
Next_ATT:
                Next k
            Case Is = 17
                strTemp = "TDS"
                For k = 0 To UBound(vLine)
                    str1 = ""
                    If CInt(Left(vLine(j), 1)) < 1 Then
                        strTemp = vLine(j)
                        GoTo Next_TDS
                    End If
            
                    If Len(vLine(j)) < 3 Then vLine(j) = vLine(j) & "-0"
                    
                    If UCase(Right(vLine(j), 1)) = "X" Then
                        str1 = "ATTACH"
                        vLine(j) = Replace(vItem(j), "X", "")
                    End If
                    
                    If UCase(Right(vLine(j), 1)) = "C" Then
                        strTemp = strTemp & " C-WIRE"
                        vLine(j) = Replace(vItem(j), "C", "")
                    End If
                    
                    lbAttach.AddItem strTemp
                    lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(j)
                    lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                    lbAttach.List(lbAttach.ListCount - 1, 3) = str1
Next_TDS:
                Next k
            Case Is = 18
                strTemp = "UTC"
                For k = 0 To UBound(vLine)
                    str1 = ""
                    If CInt(Left(vLine(j), 1)) < 1 Then
                        strTemp = vLine(j)
                        GoTo Next_UTC
                    End If
            
                    If Len(vLine(j)) < 3 Then vLine(j) = vLine(j) & "-0"
                    
                    If UCase(Right(vLine(j), 1)) = "X" Then
                        str1 = "ATTACH"
                        vLine(j) = Replace(vItem(j), "X", "")
                    End If
                    
                    If UCase(Right(vLine(j), 1)) = "C" Then
                        strTemp = strTemp & " C-WIRE"
                        vLine(j) = Replace(vItem(j), "C", "")
                    End If
                    
                    lbAttach.AddItem strTemp
                    lbAttach.List(lbAttach.ListCount - 1, 1) = vLine(j)
                    lbAttach.List(lbAttach.ListCount - 1, 2) = ""
                    lbAttach.List(lbAttach.ListCount - 1, 3) = str1
Next_UTC:
                Next k
        End Select
        '<------------------------------------------------------------- Continue getting Pole Attachments
Next_N:
    Next n
    
    Me.show
End Sub

Private Sub cbGetsPole_Click()
    Dim objObject As AcadObject
    'Dim objBlock As AcadBlockReference
    Dim vBasePnt, vAttList As Variant
    Dim vLine, vItem, vTemp As Variant
    Dim strTemp As String
    
    lbData.Clear
    lbAttach.Clear
    lbUnits.Clear
    lbCables.Clear
    lbSplices.Clear
    
    cbUpdatePole.Enabled = True
    
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
    
    If Not objPole.Name = "sPole" Then
        MsgBox "Not a valid pole."
        Me.show
        Exit Sub
    End If
    
    vAttList = objPole.GetAttributes
    
    tbPoleNumber.Value = vAttList(0).TextString
    tbOwner.Value = vAttList(2).TextString
    
    lbData.AddItem "1  H-C"
    If vAttList(5).TextString = "" Then
        lbData.List(lbData.ListCount - 1, 1) = "?-?"
    Else
        lbData.List(lbData.ListCount - 1, 1) = vAttList(5).TextString
    End If
    
    
    
    If vAttList(3).TextString = "" Then
        lbData.AddItem "2  Owner #"
        lbData.List(lbData.ListCount - 1, 1) = "NA"
    Else
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
    End If
    
    
    
    If Not vAttList(4).TextString = "" Then
        vLine = Split(vAttList(4).TextString, " ")
        
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), "=")
            
            If UBound(vItem) = 0 Then
                lbData.AddItem "3  Other #"
                lbData.List(lbData.ListCount - 1, 1) = "???  " & vLine(i)
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
    
    '<---------------------------------------------------------------- Add Attachments
    
    For n = 9 To 24
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
                vLine = Split(vAttList(n).TextString, " ")
                For k = 0 To UBound(vLine)
                    lbAttach.AddItem "NEW 6M"
                    
                    lbAttach.List(lbAttach.ListCount - 1, 1) = ""
                    lbAttach.List(lbAttach.ListCount - 1, 2) = vLine(k)
                    lbAttach.List(lbAttach.ListCount - 1, 3) = "NEW"
                Next k
            Case Is > 15
                vLine = Split(UCase(vAttList(n).TextString), "=")
                vItem = Split(vLine(1), " ")
                For k = 0 To UBound(vItem)
                    If InStr(vItem(k), "C") > 0 Then
                        vLine(0) = vLine(0) & " C-WIRE"
                        vItem(k) = Replace(vItem(k), "C", "")
                    End If
                    
                    If InStr(vItem(k), "D") > 0 Then
                        vLine(0) = vLine(0) & " DROP"
                        vItem(k) = Replace(vItem(k), "D", "")
                    End If
                    
                    If InStr(vItem(k), "O") > 0 Then
                        vLine(0) = vLine(0) & " OHG"
                        vItem(k) = Replace(vItem(k), "O", "")
                    End If
                    
                    lbAttach.AddItem vLine(0)
                    
                    If InStr(vItem(k), "X") > 0 Then
                        lbAttach.List(lbAttach.ListCount - 1, 3) = "ATTACH"
                        vItem(k) = Replace(vItem(k), "X", "")
                    Else
                        lbAttach.List(lbAttach.ListCount - 1, 3) = ""
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
                Next k
        End Select
Next_N:
    Next n
    
    'att 25
    If Not vAttList(25).TextString = "" Then
        vAttList(25).TextString = Replace(vAttList(25).TextString, vbLf, "")
        vLine = Split(vAttList(25).TextString, vbCr)
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), " / ")
            lbCables.AddItem vItem(0)
            lbCables.List(lbCables.ListCount - 1, 1) = vItem(1)
        Next i
    End If
    
    'att 26
    If Not vAttList(26).TextString = "" Then
        vAttList(26).TextString = Replace(vAttList(26).TextString, vbLf, "")
        'If InStr(vAttList(26).TextString, " + ") > 0 Then vAttList(26).TextString = Replace(vAttList(26).TextString, " + ", vbCr)
        
        vLine = Split(vAttList(26).TextString, vbCr)
        For i = 0 To UBound(vLine)
            lbSplices.AddItem vLine(i)
        Next i
    End If
    
    'att 27
    If Not vAttList(27).TextString = "" Then
        vAttList(27).TextString = Replace(vAttList(27).TextString, vbLf, "")
        vLine = Split(vAttList(27).TextString, ";;")
        For i = 0 To UBound(vLine)
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
    
    Me.show
End Sub

Private Sub cbGetUnits_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vLine, vItem As Variant
    Dim strTemp, strLine As String
    Dim iFeet, iInch, iTotal, iRL As Integer
    
    Me.Hide
    
  On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        
    objSS.SelectOnScreen
    For Each objEntity In objSS
        If Not TypeOf objEntity Is AcadBlockReference Then GoTo Next_objEntity
        
        Set objBlock = objEntity
        If objBlock.Name = "pole_unit" Then
            vAttList = objBlock.GetAttributes
            
            vLine = Split(vAttList(3).TextString, "=")
            lbUnits.AddItem vLine(0), 0
            
            vLine(1) = Replace(vLine(1), "  ", " ")
            vItem = Split(vLine(1), " ")
            lbUnits.List(0, 1) = vItem(0)
            If UBound(vItem) < 1 Then
                lbUnits.List(0, 2) = ""
            Else
                lbUnits.List(0, 2) = vItem(1)
            End If
        End If
Next_objEntity:
    Next objEntity
    
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbPlace_Click()
    If lbCables.ListIndex < 0 Then Exit Sub
    
    Dim objCallout As AcadBlockReference
    'Dim strBlock As String
    
    'strBlock = "Callout"
    
    Dim returnPnt As Variant
    Dim vBlockCoords(0 To 2) As Double
    Dim lwpCoords(0 To 3) As Double
    Dim dPrevious(0 To 2) As Double
    Dim dOrigin(0 To 2) As Double
    Dim dScale As Double
    Dim lineObj As AcadLWPolyline
    Dim objCircle As AcadCircle
    Dim objCircle2 As AcadCircle
    Dim objText As AcadText
    Dim objText2 As AcadText
    Dim str1, str2, str3, strLetter As String
    Dim lowNum, highNum, strArray() As String
    Dim mainStatus, tempStatus As String
    Dim vLine, vCounts As Variant
    Dim iIndex As Integer
    
    Dim strAtt0, strAtt1, strAtt2 As String
    Dim strLayer As String
    
    iIndex = lbCables.ListIndex
    
    vLine = Split(lbCables.List(iIndex, 0), ": ")
    strAtt0 = tbPoleNumber.Value & ": " & vLine(0)
    strAtt1 = vLine(1)
    strAtt2 = Replace(lbCables.List(iIndex, 1), " + ", "\P")
    
    If Left(vLine(1), 2) = "CO" Then
        strLayer = "Integrity Proposed-Aerial"
    Else
        strLayer = "Integrity Proposed-Buried"
    End If
    
  'On Error Resume Next
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
    'If cbSuffix.Value = "" And cbCblType.Value = "CO" Then
        'str1 = "Do you need a cable suffix?"
        'result = MsgBox(str1, vbYesNo)
        'If result = 6 Then Exit Sub
    'End If
    
    Me.Hide
    
    returnPnt = ThisDrawing.Utility.GetPoint(, "Select Point: ")
    lwpCoords(0) = returnPnt(0)
    lwpCoords(1) = returnPnt(1)
    dPrevious(0) = returnPnt(0)
    dPrevious(1) = returnPnt(1)
    dPrevious(2) = 0#
    
    returnPnt = ThisDrawing.Utility.GetPoint(dPrevious, "Select Point: ")
    vBlockCoords(0) = returnPnt(0)
    vBlockCoords(1) = returnPnt(1)
    vBlockCoords(2) = returnPnt(2)
    lwpCoords(2) = vBlockCoords(0)
    lwpCoords(3) = vBlockCoords(1)
    
    If cbPlacement.Value = "Leader" Then
        Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
        lineObj.Layer = strLayer
        lineObj.Update
    Else
        Set objCircle = ThisDrawing.ModelSpace.AddCircle(dPrevious, 8)
        objCircle.Layer = strLayer
        objCircle.Update
        Set objCircle2 = ThisDrawing.ModelSpace.AddCircle(returnPnt, 8)
        objCircle2.Layer = strLayer
        objCircle2.Update
        
        strLetter = UCase(ThisDrawing.Utility.GetString(0, "Enter Callout Letter:"))
        Set objText = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8)
        Set objText2 = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8)
        objText.Layer = strLayer
        objText.Alignment = acAlignmentMiddle
        objText.TextAlignmentPoint = dPrevious
        objText2.Layer = strLayer
        objText2.Alignment = acAlignmentMiddle
        objText2.TextAlignmentPoint = vBlockCoords
        objText.Update
        objText2.Update
        
        vBlockCoords(0) = vBlockCoords(0) + 8
        lwpCoords(0) = 0
    End If
    
    
    'dScale = CDbl(cbScale.Value) / 100
    If dScale = 0 Then dScale = 0.75
        
    Set objCallout = ThisDrawing.ModelSpace.InsertBlock(vBlockCoords, "Callout", dScale, dScale, dScale, 0)
    attItem = objCallout.GetAttributes
    
    attItem(0).TextString = strAtt0
    attItem(1).TextString = strAtt1
    attItem(2).TextString = strAtt2
    
    objCallout.Layer = strLayer
    
    If lwpCoords(2) < lwpCoords(0) Then
        vBlockCoords(0) = vBlockCoords(0) - (75 * dScale)
        objCallout.InsertionPoint = vBlockCoords
    End If
    objCallout.Update
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbSort_Click()
    Call SortAttachments
End Sub

Private Sub cbUpdatePole_Click()
    Dim strLine As String
    Dim vAttList, vLine As Variant
    
    strLine = ""
    lbTemp.Clear
    
    vAttList = objPole.GetAttributes
    
    '<------------------------------------------------------------------- Data
    
    vAttList(0).TextString = tbPoleNumber.Value
    vAttList(2).TextString = tbOwner.Value
    
    If lbData.ListCount > 0 Then
        For i = 0 To lbData.ListCount - 1
            Select Case Left(lbData.List(i, 0), 1)
                Case "1"
                    vAttList(5).TextString = Replace(lbData.List(i, 1), " - ", "-")
                Case "2"
                    vAttList(3).TextString = lbData.List(i, 1)
                Case "3"
                    If strLine = "" Then
                        If InStr(lbData.List(i, 1), "  ") > 0 Then
                            strLine = tbOwner.Value & "*" & lbData.List(i, 1)
                        Else
                            strLine = lbData.List(i, 1)
                        End If
                    Else
                        If InStr(lbData.List(i, 1), "  ") > 0 Then
                            strLine = strLine & " " & tbOwner.Value & "*" & lbData.List(i, 1)
                        Else
                            strLine = strLine & " " & lbData.List(i, 1)
                        End If
                    End If
                Case "9"
                    vAttList(8).TextString = lbData.List(i, 1)
            End Select
        Next i
        
        vAttList(4).TextString = strLine
    End If
    
    '<------------------------------------------------------------------- Attachments
    
    If lbAttach.ListCount > 0 Then
        Dim strN, strT, strLP, strA, strSLC, strSL As String
        Dim strNew, strOHG As String
    
        For i = 0 To lbAttach.ListCount - 1
            strLine = ""
            
            If lbAttach.List(i, 1) = "" Then
                strLine = lbAttach.List(i, 2)
            Else
                If lbAttach.List(i, 2) = "" Then
                    strLine = lbAttach.List(i, 1)
                Else
                    strLine = "(" & lbAttach.List(i, 1) & ")" & lbAttach.List(i, 2)
                End If
            End If
            
            Select Case lbAttach.List(i, 0)
                Case "NEUTRAL"
                    If strN = "" Then
                        strN = strLine
                    Else
                        strN = strN & " " & strLine
                    End If
                Case "TRANSFORMER"
                    If strT = "" Then
                        strT = strLine
                    Else
                        strT = strT & " " & strLine
                    End If
                Case "LOW POWER"
                    If strLP = "" Then
                        strLP = strLine
                    Else
                        strLP = strLP & " " & strLine
                    End If
                Case "ANTENNA"
                    If strA = "" Then
                        strA = strLine
                    Else
                        strA = strA & " " & strLine
                    End If
                Case "ST LT CIR"
                    If strSLC = "" Then
                        strSLC = strLine
                    Else
                        strSLC = strSLC & " " & strLine
                    End If
                Case "ST LT"
                    If strSL = "" Then
                        strSL = strLine
                    Else
                        strSL = strSL & " " & strLine
                    End If
                Case Else
                    If InStr(lbAttach.List(i, 3), "NEW") > 0 Then
                        If strNew = "" Then
                            strNew = strLine
                        Else
                            strNew = strNew & " " & strLine
                        End If
                    Else
                        If lbTemp.ListCount = 0 Then
                            lbTemp.AddItem lbAttach.List(i, 0)
                            lbTemp.List(0, 1) = strLine
                        Else
                            For k = 0 To lbTemp.ListCount - 1
                                If lbTemp.List(k, 0) = lbAttach.List(i, 0) Then
                                    lbTemp.List(k, 1) = lbTemp.List(k, 1) & " " & strLine
                                    GoTo Found_Match
                                End If
                            Next k
                            
                            lbTemp.AddItem lbAttach.List(i, 0)
                            lbTemp.List(lbTemp.ListCount - 1, 1) = strLine
                            
Found_Match:
                        End If
                    End If
            End Select
            
        Next i
        
        If Not strN = "" Then vAttList(9).TextString = strN
        If Not strT = "" Then vAttList(10).TextString = strT
        If Not strLP = "" Then vAttList(11).TextString = strLP
        If Not strA = "" Then vAttList(12).TextString = strA
        If Not strSLC = "" Then vAttList(13).TextString = strSLC
        If Not strSL = "" Then vAttList(14).TextString = strSL
        If Not strNew = "" Then vAttList(15).TextString = strNew
        
        If lbTemp.ListCount > 0 Then
            For k = 0 To lbTemp.ListCount - 1
                Select Case k
                    Case Is = 0
                        vAttList(16).TextString = lbTemp.List(0, 0) & "=" & lbTemp.List(0, 1)
                    Case Is = 1
                        vAttList(17).TextString = lbTemp.List(1, 0) & "=" & lbTemp.List(1, 1)
                    Case Is = 2
                        vAttList(18).TextString = lbTemp.List(2, 0) & "=" & lbTemp.List(2, 1)
                    Case Is = 3
                        vAttList(19).TextString = lbTemp.List(3, 0) & "=" & lbTemp.List(3, 1)
                    Case Is = 4
                        vAttList(20).TextString = lbTemp.List(4, 0) & "=" & lbTemp.List(4, 1)
                    Case Is = 5
                        vAttList(21).TextString = lbTemp.List(5, 0) & "=" & lbTemp.List(5, 1)
                    Case Is = 6
                        vAttList(22).TextString = lbTemp.List(6, 0) & "=" & lbTemp.List(6, 1)
                    Case Is = 7
                        vAttList(23).TextString = lbTemp.List(7, 0) & "=" & lbTemp.List(7, 1)
                    Case Else
                        MsgBox "Maximum Comms Reached"
                End Select
            Next k
        End If
    End If
    
    '<------------------------------------------------------------------- Cables
    strLine = ""
    
    If lbCables.ListCount > 0 Then
        For i = 0 To lbCables.ListCount - 1
            If strLine = "" Then
                strLine = lbCables.List(i, 0) & " / " & lbCables.List(i, 1)
            Else
                strLine = strLine & vbCr & lbCables.List(i, 0) & " / " & lbCables.List(i, 1)
            End If
        Next i
        
        vAttList(25).TextString = strLine
    End If
    
    '<------------------------------------------------------------------- Splices
    strLine = ""
    
    If lbSplices.ListCount > 0 Then
        For i = 0 To lbSplices.ListCount - 1
            If strLine = "" Then
                strLine = lbSplices.List(i, 0)
            Else
                strLine = strLine & " + " & lbSplices.List(i, 0)
            End If
        Next i
        
        vAttList(26).TextString = strLine
    End If
    
    '<------------------------------------------------------------------- Units
    strLine = ""
    
    If lbUnits.ListCount > 0 Then
        For i = 0 To lbUnits.ListCount - 1
            If strLine = "" Then
                strLine = lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
                If Not lbUnits.List(i, 2) = "" Then strLine = strLine & "  " & lbUnits.List(i, 2)
            Else
                strLine = strLine & ";;" & lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
                If Not lbUnits.List(i, 2) = "" Then strLine = strLine & "  " & lbUnits.List(i, 2)
            End If
        Next i
        
        vAttList(27).TextString = strLine
    End If
    
    cbUpdatePole.Enabled = False
End Sub

Private Sub Label6_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub lbAttach_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cbAddAttach.Caption = "Update"
    
    cbAttachment.Value = lbAttach.List(lbAttach.ListIndex, 0)
    tbEAttach.Value = lbAttach.List(lbAttach.ListIndex, 1)
    tbPAttach.Value = lbAttach.List(lbAttach.ListIndex, 2)
End Sub

Private Sub lbAttach_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim str0, str1, str2, str3 As String
    
    Select Case KeyCode
        Case vbKeyDelete
            If lbAttach.ListIndex < 0 Then Exit Sub
            lbAttach.RemoveItem lbAttach.ListIndex
        Case vbKeyUp
            If lbAttach.ListIndex = 0 Then Exit Sub
            
            str0 = lbAttach.List(lbAttach.ListIndex, 0)
            str1 = lbAttach.List(lbAttach.ListIndex, 1)
            str2 = lbAttach.List(lbAttach.ListIndex, 2)
            str3 = lbAttach.List(lbAttach.ListIndex, 3)
            
            lbAttach.List(lbAttach.ListIndex, 0) = lbAttach.List(lbAttach.ListIndex - 1, 0)
            lbAttach.List(lbAttach.ListIndex, 1) = lbAttach.List(lbAttach.ListIndex - 1, 1)
            lbAttach.List(lbAttach.ListIndex, 2) = lbAttach.List(lbAttach.ListIndex - 1, 2)
            lbAttach.List(lbAttach.ListIndex, 3) = lbAttach.List(lbAttach.ListIndex - 1, 3)
            
            lbAttach.List(lbAttach.ListIndex - 1, 0) = str0
            lbAttach.List(lbAttach.ListIndex - 1, 1) = str1
            lbAttach.List(lbAttach.ListIndex - 1, 2) = str2
            lbAttach.List(lbAttach.ListIndex - 1, 3) = str3
            
            'lbAttach.ListIndex = lbAttach.ListIndex - 1
        Case vbKeyDown
            If lbAttach.ListIndex = lbAttach.ListCount - 1 Then Exit Sub
            
            str0 = lbAttach.List(lbAttach.ListIndex, 0)
            str1 = lbAttach.List(lbAttach.ListIndex, 1)
            str2 = lbAttach.List(lbAttach.ListIndex, 2)
            str3 = lbAttach.List(lbAttach.ListIndex, 3)
            
            lbAttach.List(lbAttach.ListIndex, 0) = lbAttach.List(lbAttach.ListIndex + 1, 0)
            lbAttach.List(lbAttach.ListIndex, 1) = lbAttach.List(lbAttach.ListIndex + 1, 1)
            lbAttach.List(lbAttach.ListIndex, 2) = lbAttach.List(lbAttach.ListIndex + 1, 2)
            lbAttach.List(lbAttach.ListIndex, 3) = lbAttach.List(lbAttach.ListIndex + 1, 3)
            
            lbAttach.List(lbAttach.ListIndex + 1, 0) = str0
            lbAttach.List(lbAttach.ListIndex + 1, 1) = str1
            lbAttach.List(lbAttach.ListIndex + 1, 2) = str2
            lbAttach.List(lbAttach.ListIndex + 1, 3) = str3
            
            'lbAttach.ListIndex = lbAttach.ListIndex + 1
    End Select
End Sub

Private Sub lbCables_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
    
    Load ArcViewPoleCable
    
        ArcViewPoleCable.tbCable.Value = lbCables.List(lbCables.ListIndex, 0)
        ArcViewPoleCable.tbCounts.Value = Replace(lbCables.List(lbCables.ListIndex, 1), " + ", vbCr)
    
        ArcViewPoleCable.show
        
        If Not ArcViewPoleCable.tbCounts.Value = "" Then
            lbCables.List(lbCables.ListIndex, 0) = ArcViewPoleCable.tbCable.Value
            lbCables.List(lbCables.ListIndex, 1) = Replace(ArcViewPoleCable.tbCounts.Value, vbCr, " + ")
            lbCables.List(lbCables.ListIndex, 1) = Replace(lbCables.List(lbCables.ListIndex, 1), vbLf, "")
        End If
        
    Unload ArcViewPoleCable
    
    Me.show
End Sub

Private Sub lbCables_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim str0, str1 As String
    
    Select Case KeyCode
        Case vbKeyDelete
            lbCounts.RemoveItem lbCounts.ListIndex
        Case vbKeyUp
            If lbCounts.ListIndex = 0 Then Exit Sub
            
            str0 = lbCounts.List(lbCounts.ListIndex, 0)
            str1 = lbCounts.List(lbCounts.ListIndex, 1)
            
            lbCounts.List(lbCounts.ListIndex, 0) = lbCounts.List(lbCounts.ListIndex - 1, 0)
            lbCounts.List(lbCounts.ListIndex, 1) = lbCounts.List(lbCounts.ListIndex - 1, 1)
            
            lbCounts.List(lbCounts.ListIndex - 1, 0) = str0
            lbCounts.List(lbCounts.ListIndex - 1, 1) = str1
            
            'lbCounts.ListIndex = lbCounts.ListIndex - 1
        Case vbKeyDown
            If lbCounts.ListIndex = lbCounts.ListCount - 1 Then Exit Sub
            
            str0 = lbCounts.List(lbCounts.ListIndex, 0)
            str1 = lbCounts.List(lbCounts.ListIndex, 1)
            
            lbCounts.List(lbCounts.ListIndex, 0) = lbCounts.List(lbCounts.ListIndex + 1, 0)
            lbCounts.List(lbCounts.ListIndex, 1) = lbCounts.List(lbCounts.ListIndex + 1, 1)
            
            lbCounts.List(lbCounts.ListIndex + 1, 0) = str0
            lbCounts.List(lbCounts.ListIndex + 1, 1) = str1
            
            'lbCounts.ListIndex = lbCounts.ListIndex + 1
    End Select
End Sub

Private Sub lbData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Select Case Left(lbData.List(lbData.ListIndex, 0), 1)
        Case "1"
            cbDType.Value = "H-C"
        Case "2"
            cbDType.Value = "Owner #"
        Case "3"
            cbDType.Value = "Other #"
        Case "9"
            cbDType.Value = "Ground"
    End Select
    cbDType.Enabled = False
    cbAddData.Caption = "Update"
    
    tbDValue.Value = lbData.List(lbData.ListIndex, 1)
End Sub

Private Sub lbData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim str0, str1 As String
    
    Select Case KeyCode
        Case vbKeyDelete
            lbData.RemoveItem lbData.ListIndex
        Case vbKeyReturn
            Select Case Left(lbData.List(lbData.ListIndex, 0), 1)
                Case "1"
                    cbDType.Value = "H-C"
                Case "2"
                    cbDType.Value = "Owner #"
                Case "3"
                    cbDType.Value = "Other #"
                Case "9"
                    cbDType.Value = "Ground"
            End Select
            cbDType.Enabled = False
            cbAddData.Caption = "Update"
    
            tbDValue.Value = lbData.List(lbData.ListIndex, 1)
        Case vbKeyUp
            If lbData.ListIndex = 0 Then Exit Sub
            
            str0 = lbData.List(lbData.ListIndex, 0)
            str1 = lbData.List(lbData.ListIndex, 1)
            
            lbData.List(lbData.ListIndex, 0) = lbData.List(lbData.ListIndex - 1, 0)
            lbData.List(lbData.ListIndex, 1) = lbData.List(lbData.ListIndex - 1, 1)
            
            lbData.List(lbData.ListIndex - 1, 0) = str0
            lbData.List(lbData.ListIndex - 1, 1) = str1
            
            'lbData.ListIndex = lbData.ListIndex - 1
        Case vbKeyDown
            If lbData.ListIndex = lbData.ListCount - 1 Then Exit Sub
            
            str0 = lbData.List(lbData.ListIndex, 0)
            str1 = lbData.List(lbData.ListIndex, 1)
            
            lbData.List(lbData.ListIndex, 0) = lbData.List(lbData.ListIndex + 1, 0)
            lbData.List(lbData.ListIndex, 1) = lbData.List(lbData.ListIndex + 1, 1)
            
            lbData.List(lbData.ListIndex + 1, 0) = str0
            lbData.List(lbData.ListIndex + 1, 1) = str1
            
            'lbData.ListIndex = lbData.ListIndex + 1
    End Select
End Sub

Private Sub lbSplices_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cbAddSplice.Caption = "Update"
    
    tbSValue.Value = lbSplices.List(lbSplices.ListIndex)
End Sub

Private Sub lbUnits_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cbAddUnit.Caption = "Update"
    
    tbUnit.Value = lbUnits.List(lbUnits.ListIndex, 0)
    tbQuantity.Value = lbUnits.List(lbUnits.ListIndex, 1)
    tbUNote.Value = lbUnits.List(lbUnits.ListIndex, 2)
End Sub

Private Sub lbUnits_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim str0, str1, str2 As String
    
    Select Case KeyCode
        Case vbKeyDelete
            lbUnits.RemoveItem lbUnits.ListIndex
        Case vbKeyUp
            If lbUnits.ListIndex = 0 Then Exit Sub
            
            str0 = lbUnits.List(lbUnits.ListIndex, 0)
            str1 = lbUnits.List(lbUnits.ListIndex, 1)
            str2 = lbUnits.List(lbUnits.ListIndex, 2)
            
            lbUnits.List(lbUnits.ListIndex, 0) = lbUnits.List(lbUnits.ListIndex - 1, 0)
            lbUnits.List(lbUnits.ListIndex, 1) = lbUnits.List(lbUnits.ListIndex - 1, 1)
            lbUnits.List(lbUnits.ListIndex, 2) = lbUnits.List(lbUnits.ListIndex - 1, 2)
            
            lbUnits.List(lbUnits.ListIndex - 1, 0) = str0
            lbUnits.List(lbUnits.ListIndex - 1, 1) = str1
            lbUnits.List(lbUnits.ListIndex - 1, 2) = str2
            
            'lbUnits.ListIndex = lbUnits.ListIndex - 1
        Case vbKeyDown
            If lbUnits.ListIndex = lbUnits.ListCount - 1 Then Exit Sub
            
            str0 = lbUnits.List(lbUnits.ListIndex, 0)
            str1 = lbUnits.List(lbUnits.ListIndex, 1)
            str2 = lbUnits.List(lbUnits.ListIndex, 2)
            
            lbUnits.List(lbUnits.ListIndex, 0) = lbUnits.List(lbUnits.ListIndex + 1, 0)
            lbUnits.List(lbUnits.ListIndex, 1) = lbUnits.List(lbUnits.ListIndex + 1, 1)
            lbUnits.List(lbUnits.ListIndex, 2) = lbUnits.List(lbUnits.ListIndex + 1, 2)
            
            lbUnits.List(lbUnits.ListIndex + 1, 0) = str0
            lbUnits.List(lbUnits.ListIndex + 1, 1) = str1
            lbUnits.List(lbUnits.ListIndex + 1, 2) = str2
            
            'lbUnits.ListIndex = lbUnits.ListIndex + 1
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbData.ColumnCount = 2
    lbData.ColumnWidths = "60;180"
    
    lbAttach.ColumnCount = 4
    lbAttach.ColumnWidths = "84;36;36;84"
    
    lbUnits.ColumnCount = 3
    lbUnits.ColumnWidths = "96;36;48"
    
    lbCables.ColumnCount = 2
    lbCables.ColumnWidths = "96;285"
    
    lbTemp.ColumnCount = 2
    lbTemp.ColumnWidths = "48;72"
    
    cbDType.AddItem "H-C"
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
    
    cbPlacement.AddItem "Leader"
    cbPlacement.AddItem "Away"
    cbPlacement.Value = "Leader"
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
End Sub

Private Function GetMR(strE As String, strP As String)
    Dim vLine, vItem As Variant
    Dim strMR As String
    Dim iEF, iEI, iPF, iPI, iMR As Integer
    
    Select Case strE
        Case ""
            strMR = "NEW"
        Case "X"
            strMR = "ATTACH"
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
