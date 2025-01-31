VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnitForm 
   Caption         =   "Design Units"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8535.001
   OleObjectBlob   =   "UnitForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objBlock As AcadBlockReference
Dim iListIndex As Integer
Dim dScale As Double
Dim strLastUnit As String

Private Sub cbAddCommon_Click()
    If lbUnits.ListIndex < 0 Then Exit Sub
    
    Dim iIndex, i As Integer
    
    iIndex = lbUnits.ListIndex
    
    lbCommon.AddItem lbUnits.List(iIndex, 0)
    i = lbCommon.ListCount - 1
    
    lbCommon.List(i, 1) = lbUnits.List(iIndex, 1)
    If lbUnits.List(iIndex, 2) = Null Then
        lbCommon.List(i, 2) = lbUnits.List(iIndex, 2)
    Else
        lbCommon.List(i, 2) = lbUnits.List(iIndex, 2)
    End If
End Sub

Private Sub cbAddCommonUnits_Click()
    If lbCommon.ListCount < 1 Then Exit Sub
    
    Dim iIndex, i As Integer
    
    'iIndex = lbCommon.ListIndex
    For i = 0 To lbCommon.ListCount - 1
        lbUnits.AddItem lbCommon.List(i, 0)
        iIndex = lbUnits.ListCount - 1
    
        lbUnits.List(iIndex, 1) = lbCommon.List(i, 1)
        lbUnits.List(iIndex, 2) = lbCommon.List(i, 2)
    Next i
End Sub

Private Sub cbAddtoDWG_Click()
    Dim vAttList As Variant
    Dim strLine As String
    Dim i, iAtt As Integer
    
    strLine = ""
    
    If lbUnits.ListCount < 1 Then GoTo Add_Info
    
    strLine = lbUnits.List(0, 0) & "=" & lbUnits.List(0, 1)
    If Not lbUnits.List(0, 2) = "" Then strLine = strLine & "  " & lbUnits.List(0, 2)
    
    If lbUnits.ListCount > 1 Then
        For i = 1 To lbUnits.ListCount - 1
            strLine = strLine & ";;" & lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
            If Not lbUnits.List(i, 2) = "" Then strLine = strLine & "  " & lbUnits.List(i, 2)
        Next i
    End If
Add_Info:
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 27
        Case Else   '   "sPed", "sHH", "sFP"
            iAtt = 7
    End Select
    
    vAttItem = objBlock.GetAttributes
    
    vAttItem(iAtt).TextString = strLine
    objBlock.Update
End Sub

Private Sub cbAddUnit_Click()
    Dim str, strMsg, str2, str3 As String
    Dim varStr As Variant
    Dim vLine, vItem As Variant
    Dim result, iSize As Integer
    Dim iIndex As Integer
    
    If cbAddUnit.Caption = "Update Unit" Then
        lbUnits.Enabled = True
        cbPrefix.Enabled = True
        'cbAddUnit.Enabled = True
        'cbGetUnit.Enabled = True
        cbAddtoDWG.Enabled = True
        cbAddUnit.Caption = "Add Unit"
    
        lbUnits.RemoveItem iListIndex
        
        vLine = Split(tbUnit.Value, "=")
        lbUnits.AddItem vLine(0), iListIndex
        If InStr(vLine(1), "  ") > 0 Then
            vItem = Split(vLine(1), "  ")
            lbUnits.List(iListIndex, 1) = vItem(0)
            lbUnits.List(iListIndex, 2) = vItem(1)
        Else
            lbUnits.List(iListIndex, 1) = vLine(1)
            lbUnits.List(iListIndex, 2) = ""
        End If
    
        'cbUpdateUnit.Enabled = False
        Exit Sub
    End If
    
    str = tbUnit.Value
    If Right(str, 1) = "=" Then str = str & "1"
    
    Select Case Left(str, 4)
        Case "+HAC", "+HBF"
            If InStr(UCase(ThisDrawing.Path), "UNITED") > 0 Then
                'str = "+" & cbPrefix.Value & cbUnits.Value
                
                strMsg = "Will this serve customers?"
                result = MsgBox(strMsg, vbYesNo, "Closure Type")

                If result = vbYes Then
                    If cbUnits.Value = "(288)" Then
                        str = str & "  YJ"
                    Else
                        str = str & "  #1413 "
                    End If
                Else
                    str = str & "  " & cbSuffix.Value
                End If
            Else
                str = str & "  " & cbSuffix.Value
                
                If cbSuffix.Value = "G6S" Then
                    lbUnits.AddItem str
                    str = "+1X32 SPLITTER=1"
                'Else
                    'str = str & "  " & cbSuffix.Value
                End If
            End If
        Case "+CO(", "+BFO", "+UO("
            'str = str + "'"
            If cbSuffix.Value = "LOOP" Then str = str & "  LOOP"
        Case "+FDH"
            If cbUnits.Value = "432" Then
                str = "+OCFH432CADNABA2AL=1"
            Else
                str = "+OCFH" & cbUnits.Value & "CABNABA2AL=1"
            End If
            
            lbUnits.AddItem str
            str = "+1X32 SPLITTER=" & cbSuffix.Value
        Case "+WCA", "BM70", "BM72", "BM73"
            'str = str + "'"
        Case "+BM6"
            If Left(str, 5) = "+BM60" Then Call RemoveIfromBFO
            'str = str + "'"
        'Case "+UD(", "+RRC", "+PE2", "+R3-"
            'str = str + "'"
    End Select
    
    vLine = Split(str, "=")
    lbUnits.AddItem vLine(0)
    iIndex = lbUnits.ListCount - 1
    
    If InStr(vLine(1), "  ") > 0 Then
        vItem = Split(vLine(1), "  ")
        lbUnits.List(iIndex, 1) = vItem(0)
        lbUnits.List(iIndex, 2) = vItem(1)
    Else
        lbUnits.List(iIndex, 1) = vLine(1)
        lbUnits.List(iIndex, 2) = ""
    End If
    
    If Left(str, 4) = "+UD(" Then
        lbUnits.AddItem "+BM61D"
        iIndex = lbUnits.ListCount - 1
        lbUnits.List(iIndex, 1) = lbUnits.List(iIndex - 1, 1)
    End If
    
    strLastUnit = ""
cbPrefix.SetFocus
End Sub

Private Sub cbCallouts_Click()
    If lbUnits.ListCount < 1 Then Exit Sub
    
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strLayer As String
    Dim dInsertPnt(2) As Double
    Dim dScale As Double
    Dim dPosition As Double
    
    On Error Resume Next
    Me.Hide
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    objSS.SelectOnScreen
    
    If objSS.count < 1 Then GoTo Exit_Sub
    
    dInsertPnt(0) = 0#
    dInsertPnt(1) = 0#
    dInsertPnt(2) = 0#
    
    For Each objBlock In objSS
        If objBlock.Name = "pole_unit" Then
            If objBlock.InsertionPoint(1) > dInsertPnt(1) Then
                dInsertPnt(0) = objBlock.InsertionPoint(0)
                dInsertPnt(1) = objBlock.InsertionPoint(1)
                dScale = objBlock.XScaleFactor
            End If
            strLayer = objBlock.Layer
            vAttList = objBlock.GetAttributes
            strAtt0 = vAttList(0).TextString
            objBlock.Delete
        End If
    Next objBlock
    
    dPosition = 2#
    
    For v = 0 To lbUnits.ListCount - 1
        strAtt1 = dPosition
        strAtt2 = "N/A"
        strAtt3 = lbUnits.List(v, 0) & "=" & lbUnits.List(v, 1)
        If Not lbUnits.List(v, 2) = "" Then strAtt3 = strAtt3 & "  " & lbUnits.List(v, 2)
    
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, "pole_unit", dScale, dScale, dScale, 0#)
        objBlock.Layer = strLayer
        
        vAttList = objBlock.GetAttributes
        vAttList(0).TextString = strAtt0
        vAttList(1).TextString = strAtt1
        vAttList(2).TextString = strAtt2
        vAttList(3).TextString = strAtt3
        objBlock.Update
    
        dInsertPnt(1) = dInsertPnt(1) - (9 * dScale)
    Next v
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbCLR_Click()
    lbUnits.Clear
End Sub

Private Sub cbCombine_Click()
    If lbUnits.ListCount < 2 Then Exit Sub
    
    Dim vCurrent, vOther, vItem As Variant
    Dim strCurrent As String
    
    For i = lbUnits.ListCount - 1 To 0 Step -1
        strCurrent = lbUnits.List(i, 0)
        
        For j = 0 To i - 1
            If lbUnits.List(j, 0) = strCurrent Then
                lbUnits.List(j, 1) = CInt(lbUnits.List(j, 1)) + CInt(lbUnits.List(i, 1))
                lbUnits.RemoveItem i
                GoTo Next_I
            End If
        Next j
Next_I:
    Next i
End Sub

Private Sub cbDelete_Click()
    lbUnits.RemoveItem (lbUnits.ListIndex)
End Sub

Private Sub cbDOWN_Click()
    Dim str1 As String
    Dim i, i2 As Integer
    
    If lbUnits.ListIndex = (lbUnits.ListCount - 1) Then Exit Sub
    i = lbUnits.ListIndex
    i2 = i + 1
    str1 = lbUnits.List(i)
    lbUnits.List(i) = lbUnits.List(i2)
    lbUnits.List(i2) = str1
    
    lbUnits.ListIndex = i2
End Sub

Private Sub cbGetCallouts_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objUnit As AcadBlockReference
    Dim objMText As AcadMText
    Dim vAtt As Variant
    Dim vLine, vItem, vTemp As Variant
    Dim strLine As String
    
    Me.Hide
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
        Err = 0
    End If
    
    objSS.SelectOnScreen
    
    If objSS.count = 0 Then GoTo Exit_Sub
    
    'lbUnits.Clear
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadBlockReference Then
            Set objUnit = objEntity
            
            If Not objUnit.Name = "pole_unit" Then GoTo Next_objEntity
            
            vAtt = objUnit.GetAttributes
            vLine = Split(vAtt(3).TextString, "=")
            
            lbUnits.AddItem vLine(0), 0
            vItem = Split(vLine(1), "  ")
            lbUnits.List(0, 1) = Replace(vItem(0), "'", "")
            If UBound(vItem) > 0 Then lbUnits.List(0, 2) = vItem(1)
        End If
        
        If TypeOf objEntity Is AcadMText Then
            Set objMText = objEntity
            
            vLine = Split(objMText.TextString, "\P")
            
            For i = 0 To UBound(vLine)
                strLine = ""
                vItem = Split(vLine(i), "=")
                
                If vItem(0) = "+UNITS" Then GoTo Next_Unit
                If InStr(vItem(0), "(LOOP)") > 0 Then
                    vItem(0) = Replace(vItem(0), "(LOOP)", "")
                    strLine = "LOOP"
                End If
                
                If InStr(vItem(0), "FOSC") > 0 Then
                    vTemp = Split(vItem(0), "FOSC")
                    vItem(0) = vTemp(0)
                    strLine = "FOSC" & vTemp(1)
                End If
                
                lbUnits.AddItem vItem(0)
                lbUnits.List(lbUnits.ListCount - 1, 1) = Replace(vItem(1), "'", "")
                lbUnits.List(lbUnits.ListCount - 1, 2) = strLine
Next_Unit:
            Next i
        End If
Next_objEntity:
    Next objEntity
    
    If lbUnits.ListCount > 0 Then
        For i = lbUnits.ListCount - 1 To 0 Step -1
            If lbUnits.List(i, 0) = "+UNITS" Then lbUnits.RemoveItem i
        Next i
    End If
    
    If lbMissing.ListCount > 0 And lbUnits.ListCount > 0 Then
        For i = lbMissing.ListCount - 1 To 0 Step -1
            strTemp = lbMissing.List(i, 0)
            
            For j = 0 To lbUnits.ListCount - 1
                If InStr(lbUnits.List(j, 0), strTemp) > 0 Then
                    If lbMissing.List(i, 1) = "" Then
                        lbMissing.RemoveItem i
                        GoTo Next_Missing
                    End If
                    
                    If lbMissing.List(i, 1) = lbUnits.List(j, 1) Then
                        lbMissing.RemoveItem i
                        GoTo Next_Missing
                    Else
                        lbMissing.List(i, 1) = CInt(lbMissing.List(i, 1)) - CInt(lbUnits.List(j, 1))
                        GoTo Next_Missing
                    End If
                End If
            Next j
            
Next_Missing:
        Next i
    End If
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbGetExisting_Click()
    Dim objObject As AcadObject
    Dim vBasePnt, vAttList As Variant
    Dim vLine, vItem, vTemp As Variant
    Dim strTemp, strType, strGround As String
    Dim iAtt As Integer
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objObject, vBasePnt, "Select Pole/Buried Plant: "
    If TypeOf objObject Is AcadBlockReference Then
        Set objBlock = objObject
    Else
        MsgBox "Not a valid object."
        Me.show
        Exit Sub
    End If
    
    Select Case objBlock.Name
        Case "sPed", "sHH", "sFP", "sPanel", "sMH"
            iAtt = 7
        Case "sPole"
            iAtt = 27
        Case Else
            MsgBox "Not a valid block."
            Me.show
            Exit Sub
    End Select
    
    vAttList = objBlock.GetAttributes
    
    lbUnits.Clear
    
    If Not vAttList(iAtt).TextString = "" Then
        vAttList(iAtt).TextString = Replace(vAttList(iAtt).TextString, vbLf, "")
        vLine = Split(vAttList(iAtt).TextString, ";;")
        For i = 0 To UBound(vLine)
            If Not Left(vLine(i), 1) = "+" Then vLine(i) = "+" & vLine(i)
            
            vItem = Split(vLine(i), "=")
            lbUnits.AddItem vItem(0)
            
            vItem(1) = Replace(vItem(1), "'", "")
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
    
    cbAddtoDWG.Enabled = True
    cbCallouts.Enabled = True
    cbGetCallouts.Enabled = True
    
    'Dim strLine, strInfo As String
    Dim strCable As String
    Dim strClient As String
    Dim iIndex As Integer
    
    tbStructure.Value = vAttList(0).TextString
    
    lbMissing.Clear
    lbAttach.Clear
    
    lbInfo.Clear
    lbInfo.AddItem "Block:"
    lbInfo.List(0, 1) = objBlock.Name
    
    Select Case objBlock.Name
        Case "sPole"
            lbInfo.AddItem "Owner:"
            lbInfo.List(1, 1) = vAttList(2).TextString
            
            lbInfo.AddItem "H-C:"
            lbInfo.List(2, 1) = vAttList(5).TextString
            
            lbInfo.AddItem "Ground:"
            lbInfo.List(3, 1) = vAttList(8).TextString
            
            For i = 9 To 14
                If Not vAttList(i).TextString = "" Then
                    vLine = Split(vAttList(i).TextString, " ")
                    
                    For j = 0 To UBound(vLine)
                        lbAttach.AddItem vAttList(i).TagString
                        iIndex = lbAttach.ListCount - 1
                        lbAttach.List(iIndex, 1) = vLine(j)
                    Next j
                End If
            Next i
            
            If Not vAttList(15).TextString = "" Then
                vLine = Split(vAttList(15).TextString, " ")
                    
                For j = 0 To UBound(vLine)
                    lbAttach.AddItem "--> NEW"
                    iIndex = lbAttach.ListCount - 1
                    lbAttach.List(iIndex, 1) = vLine(j)
                Next j
            End If
            
            For i = 16 To 23
                If Not vAttList(i).TextString = "" Then
                    vTemp = Split(vAttList(i).TextString, "=")
                    vLine = Split(vTemp(1), " ")
                    
                    For j = 0 To UBound(vLine)
                        lbAttach.AddItem vTemp(0)
                        iIndex = lbAttach.ListCount - 1
                        lbAttach.List(iIndex, 1) = vLine(j)
                    Next j
                End If
            Next i
            
            If Not vAttList(25).TextString = "" Then
                vLine = Split(vAttList(25).TextString, " / ")
                vItem = Split(vLine(0), ": ")
                lbMissing.AddItem "+" & vItem(1)
                lbMissing.List(lbMissing.ListCount - 1, 1) = ""
    
                Select Case Left(UCase(vAttList(8).TextString), 1)
                    Case "M", "T"
                        strGround = "Y"
                    Case Else
                        strGround = ""
                End Select
                
                vTemp = Split(vItem(1), ")")
                If InStr(vTemp(1), "M") > 0 And strGround = "Y" Then
                    iIndex = 0
                    For i = 0 To lbAttach.ListCount - 1
                        If lbAttach.List(i, 0) = "--> NEW" Then iIndex = iIndex + 1
                    Next i
                    If iIndex = 0 Then iIndex = 1
                    
                    lbMissing.AddItem "+PM2A"
                    lbMissing.List(lbMissing.ListCount - 1, 1) = iIndex
                End If
            End If
            
            If Not vAttList(26).TextString = "" Then
                lbMissing.AddItem "+HACO"
                lbMissing.List(lbMissing.ListCount - 1, 1) = ""
                
                lbMissing.AddItem "+PM52"
                lbMissing.List(lbMissing.ListCount - 1, 1) = "1"
            End If
            
            vLine = Split(UCase(ThisDrawing.Path), "DROPBOX\")
            If UBound(vLine) < 0 Then
                Select Case Left(vLine(1), 4)
                    Case "UNIT"
                        strClient = "UTC"
                    Case "TDS "
                        strClient = "TDS"
                    Case "LORE"
                        strClient = "LOR"
                    'Case "MAST"
                        'strClient = "ZAYO"
                    'Case "ECC "
                        'strClient = "ZAYO"
                    Case Else
                        strClient = "***"
                End Select
                
                If lbAttach.ListCount > 0 Then
                    For i = 0 To lbAttach.ListCount - 1
                        If InStr(lbAttach.List(i, 0), strClient) > 0 Then
                            If InStr(lbAttach.List(i, 1), ")") > 0 Or InStr(UCase(lbAttach.List(i, 1)), "X") > 0 Then
                                lbMissing.AddItem "+WC1"
                                lbMissing.List(lbMissing.ListCount - 1, 1) = "1"
                                
                                lbMissing.AddItem "+WC2"
                                lbMissing.List(lbMissing.ListCount - 1, 1) = "1"
                            End If
                        End If
                    Next i
                End If
            End If
            
        Case Else
            lbInfo.AddItem "Type:"
            lbInfo.List(1, 1) = vAttList(2).TextString
            
            lbInfo.AddItem "Ground:"
            lbInfo.List(2, 1) = vAttList(4).TextString
            
            If Not vAttList(5).TextString = "" Then
                vLine = Split(vAttList(5).TextString, " / ")
                vItem = Split(vLine(0), ": ")
                lbMissing.AddItem "+" & vItem(1)
                lbMissing.List(lbMissing.ListCount - 1, 1) = ""
    
                Select Case Left(UCase(vAttList(8).TextString), 1)
                    Case "M", "T"
                        lbMissing.AddItem "+BM2A"
                        lbMissing.List(lbMissing.ListCount - 1, 1) = "1"
                    Case "T"
                        lbMissing.AddItem "+BM2(5/8X8)"
                        lbMissing.List(lbMissing.ListCount - 1, 1) = "1"
                End Select
                
                vTemp = Split(vItem(1), ")")
                If InStr(vTemp(1), "M") > 0 And strGround = "Y" Then
                    lbMissing.AddItem "+PM2A"
                    lbMissing.List(lbMissing.ListCount - 1, 1) = "1"
                End If
            End If
            
            If Not vAttList(26).TextString = "" Then
                lbMissing.AddItem "+HBFO"
                lbMissing.List(lbMissing.ListCount - 1, 1) = ""
            End If
    End Select
    
    If lbMissing.ListCount > 0 And lbUnits.ListCount > 0 Then
        For i = lbMissing.ListCount - 1 To 0 Step -1
            strTemp = lbMissing.List(i, 0)
            
            For j = 0 To lbUnits.ListCount - 1
                If InStr(lbUnits.List(j, 0), strTemp) > 0 Then
                    If lbMissing.List(i, 1) = "" Then
                        lbMissing.RemoveItem i
                        GoTo Next_Missing
                    End If
                    
                    If lbMissing.List(i, 1) = lbUnits.List(j, 1) Then
                        lbMissing.RemoveItem i
                        GoTo Next_Missing
                    Else
                        lbMissing.List(i, 1) = CInt(lbMissing.List(i, 1)) - CInt(lbUnits.List(j, 1))
                        GoTo Next_Missing
                    End If
                End If
            Next j
            
Next_Missing:
        Next i
    End If
    
    Me.show
End Sub

Private Sub cbGetUnits_Click()
    Dim objSS As AcadSelectionSet
    Dim objItem As AcadObject
    Dim obrUnit As AcadBlockReference
    Dim vAttList, vCblList, vTemp As Variant
    Dim vList1, vList2, vList3 As Variant
    Dim str1, str2, str3, strLine As String
    Dim strUnit, strQty, strNote As String
    Dim strCblSize As String
    
    strCblSize = "24"
    Me.Hide
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
        Err = 0
    End If
    
    objSS.SelectOnScreen
    
    If objSS.count = 0 Then GoTo Exit_Sub
    
    lbUnits.Clear
    
    For Each objItem In objSS
        If TypeOf objItem Is AcadBlockReference Then
            Set obrUnit = objItem
            'MsgBox obrUnit.Layer
            'GoTo exit_sub
            Select Case obrUnit.Name
                Case "cable_span"
                    vAttList = obrUnit.GetAttributes
                    vCblList = Split(vAttList(1).TextString, " ")
                    vTemp = Split(vAttList(2).TextString, "' ")
                    If UBound(vTemp) > 0 Then
                        str2 = Replace(vTemp(0), "+", "")
                    Else
                        str2 = Replace(vTemp(0), "'", "")
                    End If
                    'str2 = vAttList(2).TextString
                    'str2 = Left(str2, Len(str2) - 1)
                    
                    For i = LBound(vCblList) To UBound(vCblList)
                        str1 = vCblList(i)
                        If str1 = "" Then GoTo Next_Item
                        If UCase(Right(str1, 1)) = "F" Then str1 = Left(str1, Len(str1) - 1)
                        
                        If obrUnit.Layer = "Integrity Cable-Aerial Text" Then
                            strUnit = "+CO(" & str1 & ")"
                            If i = 0 Then
                                strUnit = strUnit & "6M-EHS"
                            Else
                                strUnit = strUnit & "E"
                            End If
                        Else
                            strUnit = "+UO(" & str1 & ")"
                            
                            lbUnits.AddItem "+RRCONDUIT"
                            lbUnits.List(lbUnits.ListCount - 1, 1) = str2
                            lbUnits.List(lbUnits.ListCount - 1, 2) = ""
                            'strline = "+RRCONDUIT=" & str2 '& "'"
                        End If
                        
                        lbUnits.AddItem strUnit
                        lbUnits.List(lbUnits.ListCount - 1, 1) = str2
                        lbUnits.List(lbUnits.ListCount - 1, 2) = ""
Next_Item:
                    Next i
                Case "Map coil"
                    vAttList = obrUnit.GetAttributes
                    str1 = Replace(vAttList(0).TextString, "'", "")
                    'str1 = Left(str1, Len(str1) - 1)
                    str2 = vAttList(1).TextString
                    str2 = Replace(str2, "F", "")
                    str2 = Replace(str2, " ", "")
                    'str2 = Left(str2, Len(str2) - 1)
                    
                    If obrUnit.Layer = "Integrity Map Coils-Aerial" Then
                        strUnit = "+CO(" & str2 & ")E" '& str1 & " (LOOP)"
                    Else
                        strUnit = "+UO(" & str2 & ")" '& str1 & " (LOOP)"
                    End If
                    strQty = str1
                    'strNote = "LOOP"
                    
                    If CInt(str2) > CInt(strCblSize) Then strCblSize = str2
                    
                    lbUnits.AddItem strUnit
                    lbUnits.List(lbUnits.ListCount - 1, 1) = strQty
                    lbUnits.List(lbUnits.ListCount - 1, 2) = "LOOP"
                'Case "iClosure"
                Case "Map splice"
                    vAttList = obrUnit.GetAttributes
                    str1 = vAttList(0).TextString
                    
                    If obrUnit.Layer = "Integrity Map Splices-Aerial" Then
                        strLine = "+HACO(" & strCblSize & ")=1  " & str1
                    Else
                        strLine = "+HBFO(" & strCblSize & ")=1  " & str1
                    End If
                    
                    strCblSize = str2
                    lbUnits.AddItem strLine
                    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
                    lbUnits.List(lbUnits.ListCount - 1, 2) = str1
                Case "iClosure"
                    vAttList = obrUnit.GetAttributes
                    str1 = vAttList(0).TextString
                    
                    Select Case str1
                        Case "G5"
                            strUnit = "(288)"  '=1  G5"
                        Case "G4"
                            If strCblSize = "24" Then
                                strUnit = "(24)"   '=1  G4"
                            Else
                                strUnit = "(" & strCblSize & ")"   '=1  G4"
                            End If
                    End Select
                    
                    If obrUnit.Layer = "Integrity Map Splices-Aerial" Then
                        strUnit = "+HACO" & strUnit
                    Else
                        strUnit = "+HBFO" & strUnit
                    End If
                    
                    strCblSize = str2
                    lbUnits.AddItem strUnit
                    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
                    lbUnits.List(lbUnits.ListCount - 1, 2) = str1
                Case "pole_attach"
                    vAttList = obrUnit.GetAttributes
                    If InStr(vAttList(2).TextString, "UTC") > 0 Then
                        If Not vAttList(4).TextString = "" Then
                            'strline = "+WC1=1"
                            lbUnits.AddItem "+WC1"
                            lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
                            lbUnits.List(lbUnits.ListCount - 1, 2) = ""
                        End If
                    End If
                Case "pole_info"
                    vAttList = obrUnit.GetAttributes
                    If vAttList(1).TextString = "0.1" Then
                        If vAttList(2).TextString = "UTC" Then
                            vCblList = Split(vAttList(3).TextString, "-")
                            If CInt(vCblList(1)) > 5 Then
                                lbUnits.AddItem "+A" & vCblList(0) & "-4=1"
                                lbUnits.AddItem "+WC1=1"
                                lbUnits.AddItem "+XXPOLE=1"
                                
                                vAttList(3).TextString = "(" & vAttList(3).TextString & ") " & vCblList(0) & "-4"
                                obrUnit.Update
                            End If
                        End If
                    End If
                Case "grd"
                    If obrUnit.Layer = "Integrity Proposed" Then
                        strLine = "+PM2A=1"
                    End If
                    
                    lbUnits.AddItem strLine
                Case "IS_Pwr_Pole"
                    vAttList = obrUnit.GetAttributes
                    vList1 = Split(vAttList(0).TextString, "/")
                    vList2 = Split(vList1(UBound(vList1)), "L")
                    vList3 = Split(vList2(UBound(vList2)), "R")
                    str1 = vList3(UBound(vList3))
                    If Right(str1, 1) = "5" Or Right(str1, 1) = "0" Then
                        lbUnits.AddItem "+PM52=1"
                    End If
                Case "ExGuyOR", "ExGuyOL"
                    vAttList = obrUnit.GetAttributes
                    str1 = vAttList(2).TextString
                    If obrUnit.Layer = "Integrity Proposed" Then
                        strLine = "+" & str1 & "=1"
                        lbUnits.AddItem strLine
                        lbUnits.AddItem "+PM11=1"
                    End If
                    If obrUnit.Layer = "Integrity Guys" Then
                        strLine = "+" & str1 & "=1"
                        lbUnits.AddItem strLine
                        lbUnits.AddItem "+PM11=2"
                    End If
                Case "ohgL", "ohgR"
                    vAttList = obrUnit.GetAttributes
                    str1 = vAttList(1).TextString
                    str2 = vAttList(0).TextString
                    
                    If obrUnit.Layer = "Integrity Proposed" Then
                        strLine = "+" & str1 & "=" & str2
                        lbUnits.AddItem strLine
                    End If
                Case "ExAncOR", "ExAncOL"
                    vAttList = obrUnit.GetAttributes
                    str1 = vAttList(0).TextString
                    
                    If obrUnit.Layer = "Integrity Proposed" Then
                        strLine = "+" & str1 & "=1"
                        lbUnits.AddItem strLine
                    End If
                    If obrUnit.Layer = "Integrity Guys" Then
                        strLine = "+" & str1 & "=1"
                        lbUnits.AddItem strLine
                    End If
                Case "notes"
                    vAttList = obrUnit.GetAttributes
                    vList1 = vAttList(0).TextString
                    vList2 = Split(vList1, "=")
                    
                    If vList2(0) = "T.T" Then
                        strLine = "+R3-5=" & vList2(1)
                        lbUnits.AddItem strLine
                    End If
                Case "__Trim"
                    vAttList = obrUnit.GetAttributes
                    vList1 = vAttList(0).TextString
                    vList2 = Split(vList1, "=")
                    
                    'If vList2(0) = "T.T" Then
                        strLine = "+R3-5=" & vList2(1)
                        If Not Right(vList2(1), 1) = "'" Then
                            strLine = strLine & "'"
                        End If
                        lbUnits.AddItem strLine
                    'End If
                Case "iLHH", "iLPED"
                    vAttList = obrUnit.GetAttributes
                    str1 = "+" & vAttList(1).TextString & "=1"
                    lbUnits.AddItem str1
                Case "iPorts"
                    vAttList = obrUnit.GetAttributes
                    str1 = "+" & vAttList(0).TextString & " PORT TERM=1"
                    lbUnits.AddItem str1
            End Select
        End If
    Next objItem
    
Exit_Sub:
    ssUnits8.Clear
    ssUnits8.Delete
    Me.show
End Sub

Private Sub cbQuit_Click()
    Dim layerObj As AcadLayer
  On Error Resume Next
    Set layerObj = ThisDrawing.Layers.Add("0")
    
    ThisDrawing.ActiveLayer = layerObj
    Me.Hide
End Sub

Private Sub cbPrefix_Change()
    cbUnits.Clear
    cbSuffix.Clear
    
    tbUnit.Value = "+" & cbPrefix.Value
    
    Select Case cbPrefix.Value
        Case "A"
            cbUnits.AddItem "30-4"
            cbUnits.AddItem "30-5"
            cbUnits.AddItem "30-6"
            cbUnits.AddItem "30-7"
            cbUnits.AddItem "35-4"
            cbUnits.AddItem "40-4"
            cbUnits.AddItem "40-5"
            cbUnits.AddItem "45-4"
            cbUnits.AddItem "45-5"
            cbUnits.AddItem "50-3"
        Case "BA"
            cbUnits.AddItem "2"
            cbUnits.AddItem "3"
            cbUnits.AddItem "4"
            cbUnits.AddItem "5"
        Case "BDO"
            cbUnits.AddItem "3"
            cbUnits.AddItem "4"
            cbUnits.AddItem "5"
            cbUnits.AddItem "7"
        Case "BFO"
            cbUnits.AddItem "(12)"
            cbUnits.AddItem "(24)"
            cbUnits.AddItem "(36)"
            cbUnits.AddItem "(48)"
            cbUnits.AddItem "(72)"
            cbUnits.AddItem "(144)"
            cbUnits.AddItem "(216)"
            cbUnits.AddItem "(288)"
            cbUnits.AddItem "(432)"
                cbSuffix.AddItem ""
                cbSuffix.AddItem "I"
                cbSuffix.AddItem "T"
                cbSuffix.AddItem "LOOP"
        Case "BHF"
            cbUnits.AddItem "(30X48X36)"
                cbSuffix.AddItem ""
                cbSuffix.AddItem "T"
        Case "BM"
            cbUnits.AddItem "2(5/8)(8)"
            cbUnits.AddItem "2A"
            cbUnits.AddItem "6M"
            cbUnits.AddItem "10M"
            cbUnits.AddItem "21"
            cbUnits.AddItem "52"
            cbUnits.AddItem "53"
            cbUnits.AddItem "60"
            cbUnits.AddItem "61"
            cbUnits.AddItem "61D"
            cbUnits.AddItem "71"
            cbUnits.AddItem "72"
            cbUnits.AddItem "73"
            cbUnits.AddItem "80"
            cbUnits.AddItem "81"
            cbUnits.AddItem "82"
            cbUnits.AddItem "83"
        Case "CO"
            cbUnits.AddItem "(12)"
            cbUnits.AddItem "(24)"
            cbUnits.AddItem "(36)"
            cbUnits.AddItem "(48)"
            cbUnits.AddItem "(72)"
            cbUnits.AddItem "(96)"
            cbUnits.AddItem "(144)"
            cbUnits.AddItem "(216)"
            cbUnits.AddItem "(288)"
            cbUnits.AddItem "(432)"
                cbSuffix.AddItem ""
                cbSuffix.AddItem "E"
                cbSuffix.AddItem "6M-EHS"
                cbSuffix.AddItem "6M"
                cbSuffix.AddItem "10M"
                cbSuffix.AddItem "LOOP"
        Case "FDH"
            cbUnits.AddItem "96"
            cbUnits.AddItem "144"
            cbUnits.AddItem "288"
            cbUnits.AddItem "432"
                cbSuffix.AddItem "1"
                cbSuffix.AddItem "2"
                cbSuffix.AddItem "3"
                cbSuffix.AddItem "4"
                cbSuffix.AddItem "5"
                cbSuffix.AddItem "6"
                cbSuffix.AddItem "7"
                cbSuffix.AddItem "8"
                cbSuffix.AddItem "9"
                cbSuffix.AddItem "10"
                cbSuffix.AddItem "11"
                cbSuffix.AddItem "12"
                cbSuffix.AddItem "13"
                cbSuffix.AddItem "14"
        Case "HACO"
            cbUnits.AddItem "(12)"
            cbUnits.AddItem "(24)"
            cbUnits.AddItem "(36)"
            cbUnits.AddItem "(48)"
            cbUnits.AddItem "(72)"
            cbUnits.AddItem "(96)"
            cbUnits.AddItem "(144)"
            cbUnits.AddItem "(216)"
            cbUnits.AddItem "(288)"
            cbUnits.AddItem "(432)"
                cbSuffix.AddItem "#1413"
                cbSuffix.AddItem "FOSC A"
                cbSuffix.AddItem "FOSC B"
                cbSuffix.AddItem "FOSC D"
                cbSuffix.AddItem "G4"
                cbSuffix.AddItem "G5"
                cbSuffix.AddItem "G6"
                cbSuffix.AddItem "G6S"
        Case "HBFO"
            cbUnits.AddItem "(12)"
            cbUnits.AddItem "(24)"
            cbUnits.AddItem "(36)"
            cbUnits.AddItem "(48)"
            cbUnits.AddItem "(72)"
            cbUnits.AddItem "(96)"
            cbUnits.AddItem "(144)"
            cbUnits.AddItem "(216)"
            cbUnits.AddItem "(288)"
            cbUnits.AddItem "(432)"
        Case "PE"
            cbUnits.AddItem "1-2"
            cbUnits.AddItem "1-3"
            cbUnits.AddItem "1-4"
            cbUnits.AddItem "2-2"
            cbUnits.AddItem "2-3"
            cbUnits.AddItem "2-4"
                cbSuffix.AddItem ""
                cbSuffix.AddItem "G"
        Case "PF"
            cbUnits.AddItem "1-5"
            cbUnits.AddItem "1-7"
            cbUnits.AddItem "3-5"
            cbUnits.AddItem "5-3"
            cbUnits.AddItem "5-4"
                cbSuffix.AddItem ""
                cbSuffix.AddItem "A"
        Case "PM"
            cbUnits.AddItem "2"
            cbUnits.AddItem "2A"
            cbUnits.AddItem "11"
            cbUnits.AddItem "21"
            cbUnits.AddItem "52"
        Case "R"
            cbUnits.AddItem "1-5"
            cbUnits.AddItem "2-5"
            cbUnits.AddItem "1-10"
            cbUnits.AddItem "2-10"
            cbUnits.AddItem "3-5"
        Case "SE"
            cbUnits.AddItem "AO2"
            cbUnits.AddItem "BO2"
            cbUnits.AddItem "AO12"
            cbUnits.AddItem "BO12"
        Case "UD"
            cbUnits.AddItem "(1X1-2)"
            cbUnits.AddItem "(1X2-2)"
            cbUnits.AddItem "(1X1-4)"
        Case "UHF"
            cbUnits.AddItem "(17X30X18)"
            cbUnits.AddItem "(24X36X24)"
            cbUnits.AddItem "(30X48X36)"
            cbUnits.AddItem "(36X60X36)"
        Case "UO"
            cbUnits.AddItem "(12)"
            cbUnits.AddItem "(24)"
            cbUnits.AddItem "(36)"
            cbUnits.AddItem "(48)"
            cbUnits.AddItem "(72)"
            cbUnits.AddItem "(96)"
            cbUnits.AddItem "(144)"
            cbUnits.AddItem "(216)"
            cbUnits.AddItem "(288)"
            cbUnits.AddItem "(432)"
                cbSuffix.AddItem ""
                cbSuffix.AddItem "LOOP"
        Case "UHF"
            cbUnits.AddItem "(24X36X24)"
        Case "W"
            cbUnits.AddItem "CA"
            cbUnits.AddItem "C1"
            cbUnits.AddItem "C2"
            cbUnits.AddItem "PE"
            cbUnits.AddItem "PE2"
            cbUnits.AddItem "HACO"
            cbUnits.AddItem "HBFO"
            cbUnits.AddItem "POLE"
            cbUnits.AddItem "SEA"
            cbUnits.AddItem "SEB"
            cbUnits.AddItem "SUD"
            cbUnits.AddItem "UD"
        Case "XX"
            cbUnits.AddItem "PE"
            cbUnits.AddItem "PE2"
            cbUnits.AddItem "PF"
            cbUnits.AddItem "PM11"
            cbUnits.AddItem "PM52"
            cbUnits.AddItem "POLE"
            cbUnits.AddItem "PED"
            cbUnits.AddItem "HH"
            cbUnits.AddItem "SEA"
            cbUnits.AddItem "CABLE"
            cbUnits.AddItem "CABLE & STRAND"
            cbUnits.AddItem "STRAND"
    End Select
End Sub

Private Sub cbSort_Click()
    Dim strCO, strUO, strBDO, strHACO As String
    Dim strBM, strPM, strPE, strPF As String
    Dim strMisc, strTemp As String
    Dim vTemp As Variant
    
    strCO = ""
    strUO = ""
    strBDO = ""
    strHACO = ""
    strBM = ""
    strPM = ""
    strPE = ""
    strPF = ""
    strMisc = ""
    
    For i = 0 To lbUnits.ListCount - 1
        strTemp = lbUnits.List(i)
        Select Case Left(strTemp, 3)
            Case "+CO"
                strCO = strCO & strTemp & "**"
            Case "+UO"
                strUO = strUO & strTemp & "**"
            Case "+BD", "+UH"
                strBDO = strBDO & strTemp & "**"
            Case "+HA", "+HB"
                strHACO = strHACO & strTemp & "**"
            Case "+BM"
                strBM = strBM & strTemp & "**"
            Case "+PM"
                strPM = strPM & strTemp & "**"
            Case "+PE"
                strPE = strPE & strTemp & "**"
            Case "+PF"
                strPF = strPF & strTemp & "**"
            Case Else
                strMisc = strMisc & strTemp & "**"
        End Select
    Next i
    
    lbUnits.Clear
    
    vTemp = Split(strCO, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
    
    vTemp = Split(strUO, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
    
    vTemp = Split(strBDO, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
    
    vTemp = Split(strHACO, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
    
    vTemp = Split(strBM, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
    
    vTemp = Split(strPE, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
    
    vTemp = Split(strPF, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
    
    vTemp = Split(strPM, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
    
    vTemp = Split(strMisc, "**")
    For i = 0 To UBound(vTemp) - 1
        lbUnits.AddItem vTemp(i)
    Next i
End Sub

Private Sub cbSuffix_Change()
    If cbSuffix.Value = "LOOP" Then
        If cbPrefix.Value = "CO" Then
            tbUnit.Value = "+" & cbPrefix.Value & cbUnits.Value & "E="
        Else
            tbUnit.Value = "+" & cbPrefix.Value & cbUnits.Value & "="
        End If
    Else
        If cbPrefix.Value = "HACO" Or cbPrefix.Value = "HBFO" Then
            tbUnit.Value = "+" & cbPrefix.Value & cbUnits.Value & "="
        Else
            tbUnit.Value = "+" & cbPrefix.Value & cbUnits.Value & cbSuffix.Value & "="
        End If
    End If
    
    tbUnit.SetFocus
End Sub

Private Sub cbUnits_Change()
    tbUnit.Value = "+" & cbPrefix.Value & cbUnits.Value
    
    If cbPrefix.Value = "CO" Then cbSuffix.Value = ""
    If cbPrefix.Value = "HACO" Then cbSuffix.Value = ""
    If cbPrefix.Value = "HBFO" Then cbSuffix.Value = ""
    
    Select Case tbUnit.Value
        Case "+BM60"
            cbSuffix.Clear
            cbSuffix.AddItem "(1.25)D"
            cbSuffix.AddItem "(1.5)D"
            cbSuffix.AddItem "(2)"
            cbSuffix.AddItem "(2)D"
            cbSuffix.AddItem "(4)"
            cbSuffix.AddItem "(4)D"
            cbSuffix.AddItem "(4)DS"
            cbSuffix.AddItem "(4)DSR"
            cbSuffix.AddItem "(4)DSRR"
            cbSuffix.AddItem "(4)S"
            cbSuffix.AddItem "(4)SR"
            cbSuffix.AddItem "(4)SRR"
            cbSuffix.AddItem "(4)R"
            cbSuffix.AddItem "(4)RR"
        Case "+BM61"
            cbSuffix.Clear
            cbSuffix.AddItem ""
            cbSuffix.AddItem "D"
        Case Else
            If Left(tbUnit.Value, 5) = "+HACO" Or Left(tbUnit.Value, 5) = "+HBFO" Then
                Dim str As String
                
                str = Replace(cbUnits.Value, "(", "")
                str = Replace(str, ")", "")
                Select Case str
                    Case "12", "24"
                        If InStr(UCase(ThisDrawing.Path), "LORETTO") > 0 Then
                            cbSuffix.Value = "G4"
                        Else
                            cbSuffix.Value = "FOSC A"
                        End If
                    Case "36", "48"
                        If InStr(UCase(ThisDrawing.Path), "LORETTO") > 0 Then
                            cbSuffix.Value = "G5"
                        Else
                            cbSuffix.Value = "FOSC A"
                        End If
                    Case "72", "96"
                        If InStr(UCase(ThisDrawing.Path), "LORETTO") > 0 Then
                            cbSuffix.Value = "G5"
                        Else
                            cbSuffix.Value = "FOSC B"
                        End If
                    Case "144"
                        If InStr(UCase(ThisDrawing.Path), "LORETTO") > 0 Then
                            cbSuffix.Value = "G5"
                        Else
                            cbSuffix.Value = "FOSC D"
                        End If
                    Case Else
                        If InStr(UCase(ThisDrawing.Path), "LORETTO") > 0 Then
                            cbSuffix.Value = "G6"
                        Else
                            cbSuffix.Value = "FOSC D"
                        End If
                End Select
                'tbUnit.Value = tbUnit.Value & "="
                'tbUnit.SetFocus
            Else
                tbUnit.Value = tbUnit.Value & "="
                tbUnit.SetFocus
            End If
    End Select
End Sub

Private Sub l1x32_Click()
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+1X32 SPLITTER"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+1X32 SPLITTER" Then
            lbUnits.List(i, 1) = CInt(lbUnits.List(i, 1)) + 1
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+1X32 SPLITTER"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label10_Click()
    Dim str1, str2 As String
    
    On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+PE1-2G=1"
        Exit Sub
    End If
    
    If strLastUnit = "PE1-2G" Then
        str1 = lbUnits.List(lbUnits.ListCount - 1)
        str2 = Right(str1, Len(str1) - 8)
        str1 = Left(str1, 8)
        str2 = str2 + 1
        lbUnits.RemoveItem (lbUnits.ListCount - 1)
        lbUnits.AddItem str1 & str2
    Else
        lbUnits.AddItem "+PE1-2G=1"
    End If
    strLastUnit = "PE1-2G"
End Sub

Private Sub Label11_Click()
    Dim count As Integer
    Dim strItem As String
    Dim vItem As Variant
    
    count = 0
    
    For i = (lbUnits.ListCount - 1) To 0 Step -1
        If Left(lbUnits.List(i, 0), 4) = "+PE1" Then
            count = count + CInt(lbUnits.List(i, 1))
        End If
    Next i
    
    If count = 0 Then Exit Sub
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+PM11" Then
            lbUnits.List(i, 1) = count
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+PM11"
    lbUnits.List(lbUnits.ListCount - 1, 1) = count
End Sub

Private Sub Label12_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+PE1-3G"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+PE1-3G" Then
            lbUnits.List(i, 1) = CInt(lbUnits.List(i, 1)) + 1
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+PE1-3G"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label13_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+PF1-7A"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+PF1-7A" Then
            lbUnits.List(i, 1) = CInt(lbUnits.List(i, 1)) + 1
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+PF1-7A"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label14_Click()
    Dim str1, str2 As String
    
    On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+PE1-4G=1"
        Exit Sub
    End If
    
    If strLastUnit = "PE1-4G" Then
        str1 = lbUnits.List(lbUnits.ListCount - 1)
        str2 = Right(str1, Len(str1) - 8)
        str1 = Left(str1, 8)
        str2 = str2 + 1
        lbUnits.RemoveItem (lbUnits.ListCount - 1)
        lbUnits.AddItem str1 & str2
    Else
        lbUnits.AddItem "+PE1-4G=1"
    End If
    strLastUnit = "PE1-4G"
End Sub

Private Sub Label19_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+FP"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+FP" Then
            lbUnits.List(i, 1) = CInt(lbUnits.List(i, 1)) + 1
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+FP"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
    lbUnits.List(lbUnits.ListCount - 1, 2) = ""
End Sub

Private Sub Label20_Click()
    tbUnit.Value = "+RRCONDUIT="
    tbUnit.SetFocus
    strLastUnit = "RRUD"
End Sub

Private Sub Label21_Click()
    tbUnit.Value = "+PE2-3G="
    tbUnit.SetFocus
    'strLastUnit = "PE2-3G"
End Sub

Private Sub Label22_Click()
    tbUnit.Value = "+R3-5="
    tbUnit.SetFocus
    'strLastUnit = "R3-5"
End Sub

Private Sub Label23_Click()
    Dim entPole As AcadObject
    Dim obrGP As AcadBlockReference
    Dim attItem, basePnt As Variant
    Dim dScale As Double
    
    Me.Hide
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Block: "
    If TypeOf entPole Is AcadBlockReference Then
        Set obrGP = entPole
    Else
        MsgBox "Not a valid block."
        Exit Sub
    End If
    
    dScale = obrGP.XScaleFactor
    cbScale.Value = (dScale * 100)
    
    Me.show
End Sub

Private Sub Label24_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+BHF(30X48X36)"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+BHF(30X48X36)" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+BHF(30X48X36)"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label25_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+BDO3"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+BDO3" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+BDO3"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label26_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+BDO5"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+BDO5" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+BDO5"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label27_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+BDO(12126)"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+BDO(12126)" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+BDO(12126)"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label35_Click()
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+XXPE"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+XXPE" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+XXPE"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label36_Click()
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+XXPF"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+XXPF" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+XXPF"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label37_Click()
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+XXPOLE"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+XXPOLE" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+XXPOLE"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label38_Click()
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+XXPED"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+XXPED" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+XXPED"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label39_Click()
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+XXHH"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+XXHH" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+XXHH"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label5_Click() '<---------------------------- Get DWG Number
    Dim entPole As AcadObject
    Dim obrGP2 As AcadBlockReference
    Dim attItem, basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Staking Sheet: "
    If TypeOf entPole Is AcadBlockReference Then
        Set obrGP2 = entPole
    Else
        MsgBox "Not a valid pole."
        Exit Sub
    End If
    
    attItem = obrGP2.GetAttributes
    tbDWG.Value = Right(attItem(0).TextString, 3)
    
    Me.show
End Sub

Private Sub Label6_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+PM2A"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+PM2A" Then
            lbUnits.List(i, 1) = CInt(lbUnits.List(i, 1)) + 1
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+PM2A"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label7_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+PM52"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+PM52" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+PM52"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub Label9_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+PF1-5A"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+PF1-5A" Then
            lbUnits.List(i, 1) = CInt(lbUnits.List(i, 1)) + 1
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+PF1-5A"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub LabelBM2_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+BM2A"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+BM2A" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+BM2A"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub LabelBM602D_Click()
    Dim str1, str2 As String
    
    On Error Resume Next
    
    If strLastUnit = "UD(2)" Then
        Select Case tbUnit.Value
            Case "+UD(1X1-2"")="
                tbUnit.Value = "+UD(1X2-2)="
            Case "+UD(1X2-2"")="
                tbUnit.Value = "+UD(1X3-2)="
            Case Else
                tbUnit.Value = "+UD(1X1-2)="
        End Select
    End If
        
    tbUnit.SetFocus
    strLastUnit = "UD(2)"
End Sub

Private Sub LabelBM604D_Click()
    Dim str1, str2 As String
    
    On Error Resume Next
    
    If strLastUnit = "UD(4)" Then
        Select Case tbUnit.Value
            Case "+UD(1X1-4"")="
                tbUnit.Value = "+UD(1X2-4)="
            Case "+UD(1X2-4"")="
                tbUnit.Value = "+UD(1X3-4)="
            Case Else
                tbUnit.Value = "+UD(1X1-4)="
        End Select
    End If
        
    tbUnit.SetFocus
    strLastUnit = "UD(4)"
End Sub

Private Sub LabelBM81_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+BM81"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+BM81" Then
            lbUnits.List(i, 1) = CInt(lbUnits.List(i, 1)) + 1
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+BM81"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub LabelBM82_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+BM82"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+BM82" Then
            lbUnits.List(i, 1) = CInt(lbUnits.List(i, 1)) + 1
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem "+BM82"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub LabelHH_Click()
    'On Error Resume Next
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem "+UHF(24X36X24)"
        lbUnits.List(0, 1) = "1"
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If lbUnits.List(i, 0) = "+UHF(24X36X24)" Then Exit Sub
    Next i
    
    lbUnits.AddItem "+UHF(24X36X24)"
    lbUnits.List(lbUnits.ListCount - 1, 1) = "1"
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub LabelUD1_Click()
    tbUnit.Value = "+UD(1X1-1.25)="
    tbUnit.SetFocus
    strLastUnit = "UD(1.25)"
End Sub

Private Sub lbCommon_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbCommon.ListCount < 1 Then Exit Sub
    
    Dim str1, str2, str3 As String
    Dim i, i2 As Integer
    
    Select Case KeyCode
        Case vbKeyDown
            If lbCommon.ListIndex = (lbCommon.ListCount - 1) Then Exit Sub
            i = lbCommon.ListIndex
            i2 = i + 1
            
            str1 = lbCommon.List(i, 0)
            str2 = lbCommon.List(i, 1)
            If lbCommon.List(i, 2) > 0 Then
                str3 = lbCommon.List(i, 2)
            Else
                str3 = ""
            End If
            
            
            lbCommon.List(i, 0) = lbCommon.List(i2, 0)
            lbCommon.List(i, 1) = lbCommon.List(i2, 1)
            If lbCommon.List(i2, 2) > 0 Then
                lbCommon.List(i, 2) = lbCommon.List(i2, 2)
            Else
                lbCommon.List(i, 2) = ""
            End If
            
            lbCommon.List(i2, 0) = str1
            lbCommon.List(i2, 1) = str2
            lbCommon.List(i2, 2) = str3
            'lbCommon.ListIndex = i2
        Case vbKeyUp
            If lbCommon.ListIndex = 0 Then Exit Sub
            i = lbCommon.ListIndex
            i2 = i - 1
            
            str1 = lbCommon.List(i, 0)
            str2 = lbCommon.List(i, 1)
            If lbCommon.List(i, 2) > 0 Then
                str3 = lbCommon.List(i, 2)
            Else
                str3 = ""
            End If
            
            lbCommon.List(i, 0) = lbCommon.List(i2, 0)
            lbCommon.List(i, 1) = lbCommon.List(i2, 1)
            If lbCommon.List(i2, 2) > 0 Then
                lbCommon.List(i, 2) = lbCommon.List(i2, 2)
            Else
                lbCommon.List(i, 2) = ""
            End If
            
            lbCommon.List(i2, 0) = str1
            lbCommon.List(i2, 1) = str2
            lbCommon.List(i2, 2) = str3
            'lbCommon.ListIndex = i2
        Case vbKeyDelete
            lbCommon.RemoveItem (lbCommon.ListIndex)
    End Select
End Sub

Private Sub lbMissing_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbMissing.ListIndex < 0 Then Exit Sub
    
    Dim iIndex As Integer
    
    iIndex = lbMissing.ListIndex
    
    If lbMissing.List(iIndex, 1) = "" Then
        tbUnit.Value = lbMissing.List(iIndex, 0) & "="
        lbMissing.RemoveItem iIndex
        
        tbUnit.SetFocus
        Exit Sub
    End If
    
    If lbUnits.ListCount < 1 Then
        lbUnits.AddItem lbMissing.List(iIndex, 0)
        lbUnits.List(0, 1) = lbMissing.List(iIndex, 1)
        
        lbMissing.RemoveItem iIndex
        Exit Sub
    End If
    
    For i = 0 To lbUnits.ListCount - 1
        If InStr(lbUnits.List(i, 0), lbMissing.List(iIndex, 0)) > 0 Then
            lbUnits.List(i, 1) = CInt(lbMissing.List(iIndex, 1)) + CInt(lbUnits.List(i, 1))
            lbMissing.RemoveItem iIndex
            Exit Sub
        End If
    Next i
    
    lbUnits.AddItem lbMissing.List(iIndex, 0)
    lbUnits.List(lbUnits.ListCount - 1, 1) = lbMissing.List(iIndex, 1)
    lbMissing.RemoveItem iIndex
End Sub

Private Sub lbUnits_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    iListIndex = lbUnits.ListIndex
    lbUnits.Enabled = False
    cbPrefix.Enabled = False
    'cbAddUnit.Enabled = False
    'cbGetUnit.Enabled = False
    cbAddtoDWG.Enabled = False
    'cbUpdateUnit.Enabled = True
    cbAddUnit.Caption = "Update Unit"
    
    tbUnit.Value = lbUnits.List(iListIndex, 0) & "=" & lbUnits.List(iListIndex, 1)
    If Not lbUnits.List(iListIndex, 2) = "" Then tbUnit.Value = tbUnit.Value & "  " & lbUnits.List(iListIndex, 2)
    tbUnit.SetFocus
    
    Dim iStart, iLength As Integer
    Dim strTemp1 As String
    Dim vItems As Variant
    
    strTemp1 = tbUnit.Value
    iLength = Len(strTemp1)
    vItems = Split(strTemp1, "=")
    iStart = iLength - Len(vItems(1))
    
    If Right(vItems(1), 1) = "'" Then
        iLength = Len(vItems(1)) - 1
    Else
        iLength = Len(vItems(1))
    End If
    
    tbUnit.SelStart = iStart
    tbUnit.SelLength = iLength
    
End Sub

Private Sub lbUnits_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbUnits.ListCount < 1 Then Exit Sub
    
    Dim str1, str2, str3 As String
    Dim i, i2 As Integer
    
    Select Case KeyCode
        Case vbKeyDown
            If lbUnits.ListIndex = (lbUnits.ListCount - 1) Then Exit Sub
            i = lbUnits.ListIndex
            i2 = i + 1
            
            str1 = lbUnits.List(i, 0)
            str2 = lbUnits.List(i, 1)
            If lbUnits.List(i, 2) > 0 Then
                str3 = lbUnits.List(i, 2)
            Else
                str3 = ""
            End If
            
            
            lbUnits.List(i, 0) = lbUnits.List(i2, 0)
            lbUnits.List(i, 1) = lbUnits.List(i2, 1)
            If lbUnits.List(i2, 2) > 0 Then
                lbUnits.List(i, 2) = lbUnits.List(i2, 2)
            Else
                lbUnits.List(i, 2) = ""
            End If
            
            lbUnits.List(i2, 0) = str1
            lbUnits.List(i2, 1) = str2
            lbUnits.List(i2, 2) = str3
            'lbUnits.ListIndex = i2
        Case vbKeyUp
            If lbUnits.ListIndex = 0 Then Exit Sub
            i = lbUnits.ListIndex
            i2 = i - 1
            
            str1 = lbUnits.List(i, 0)
            str2 = lbUnits.List(i, 1)
            If lbUnits.List(i, 2) > 0 Then
                str3 = lbUnits.List(i, 2)
            Else
                str3 = ""
            End If
            
            lbUnits.List(i, 0) = lbUnits.List(i2, 0)
            lbUnits.List(i, 1) = lbUnits.List(i2, 1)
            If lbUnits.List(i2, 2) > 0 Then
                lbUnits.List(i, 2) = lbUnits.List(i2, 2)
            Else
                lbUnits.List(i, 2) = ""
            End If
            
            lbUnits.List(i2, 0) = str1
            lbUnits.List(i2, 1) = str2
            lbUnits.List(i2, 2) = str3
            'lbUnits.ListIndex = i2
        Case vbKeyDelete
            lbUnits.RemoveItem (lbUnits.ListIndex)
    End Select
End Sub

Private Sub UserForm_Initialize()
    cbPrefix.AddItem "A"
    cbPrefix.AddItem "BA"
    cbPrefix.AddItem "BDO"
    cbPrefix.AddItem "BFO"
    cbPrefix.AddItem "BHF"
    cbPrefix.AddItem "BM"
    cbPrefix.AddItem "CO"
    cbPrefix.AddItem "FDH"
    cbPrefix.AddItem "HACO"
    cbPrefix.AddItem "HBFO"
    cbPrefix.AddItem "PE"
    cbPrefix.AddItem "PF"
    cbPrefix.AddItem "PM"
    cbPrefix.AddItem "R"
    cbPrefix.AddItem "SE"
    cbPrefix.AddItem "UD"
    cbPrefix.AddItem "UHF"
    cbPrefix.AddItem "UO"
    cbPrefix.AddItem "W"
    cbPrefix.AddItem "XX"
    cbPrefix.AddItem "FP"

    cbUP.Caption = Chr(225)
    cbDOWN.Caption = Chr(226)
    
    strLastUnit = ""
    
    lbUnits.ColumnCount = 3
    lbUnits.ColumnWidths = "96;36;48"
    
    lbCommon.ColumnCount = 3
    lbCommon.ColumnWidths = "96;36;60"
    
    lbInfo.ColumnCount = 2
    lbInfo.ColumnWidths = "60;80"
    
    lbAttach.ColumnCount = 2
    lbAttach.ColumnWidths = "60;80"
    
    lbMissing.ColumnCount = 2
    lbMissing.ColumnWidths = "84;56"
End Sub

Private Sub ValidateUnits()
    Dim iNeeded, iThere As Integer
    
    iNeeded = 0
    iThere = 0
    
    On Error Resume Next
    
    For i = 0 To (lbUnits.ListCount - 1)
        If Right(lbUnits.List(i), 2) = "P)" And Left(lbUnits.List(i), 2) = "+C" Then iNeeded = 1
        If Left(lbUnits.List(i), 5) = "+PM52" Then iThere = 1
    Next i
    
    If iThere > 0 Then Exit Sub
    If Not (iNeeded - iThere) = 0 Then lbUnits.AddItem "+PM52=1"
    'MsgBox "Needed: " & iNeeded & vbCr & "There: " & iThere
End Sub

Private Sub RemoveIfromBFO()
    Dim strItem, strDuct, strTemp As String
    Dim vItem, vTemp As Variant
    Dim iExisting, iNew, iDuct As Integer
    
    For i = 0 To lbUnits.ListCount - 1
        If InStr(lbUnits.List(i, 0), ")I") > 0 Then Exit Sub
    Next i
    
    For i = 0 To lbUnits.ListCount - 1
        strItem = lbUnits.List(i, 0)
        
        If InStr(strItem, "+BFO") > 0 Then
            If lbUnits.List(i, 2) = "LOOP" Then GoTo Next_Item
            
            iExisting = CInt(lbUnits.List(i, 1))
            strDuct = Replace(lbUnits.List(i, 0), ")", ")I")
            
            vTemp = Split(tbUnit.Value, "=")
            strTemp = Replace(vTemp(1), "'", "")
            iDuct = CInt(strTemp)
            iNew = iExisting - iDuct
            
            lbUnits.List(i, 1) = iNew
            i = i + 1
            lbUnits.AddItem strDuct, i
            lbUnits.List(i, 1) = iDuct
            Exit Sub
        End If
        
Next_Item:
    Next i
End Sub
