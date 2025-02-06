VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateCounts 
   Caption         =   "Validate Counts"
   ClientHeight    =   9000.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12615
   OleObjectBlob   =   "ValidateCounts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateCounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSSPole As AcadSelectionSet
Dim objSS As AcadSelectionSet
Dim vPnt1, vPnt2 As Variant
Dim strStatus As String
Dim iIndex, iPole, iCallout As Integer

Private Sub cbChangeCallout_Click()
    Dim vLine, vItem, vTemp As Variant
    Dim strLine As String
    
    strLine = Replace(tbPole.Value, vbLf, "")
    vLine = Split(strLine, vbCr)
    strLine = ""
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        If strLine = "" Then
            strLine = vItem(0) & ": " & vItem(1)
        Else
            strLine = strLine & vbCr & vItem(0) & ": " & vItem(1)
        End If
    Next i
    
    tbCallout.Value = strLine
End Sub

Private Sub cbChangePole_Click()
    Dim vPrevious, vLine, vItem, vTemp As Variant
    Dim strLine, strList, strSource As String
    
    strList = Replace(tbPole.Value, vbLf, "")
    vPrevious = Split(strList, vbCr)
    strList = ""
    For i = 0 To UBound(vPrevious)
        
    Next i
    
    strSource = "??"
    
    
    
    
    
    strLine = Replace(tbCallout.Value, vbLf, "")
    vLine = Split(strLine, vbCr)
    strLine = ""
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        
        For j = 0 To UBound(vPrevious)
            vTemp = Split(vPrevious(j), ": ")
            If vTemp(0) = vItem(0) Then
                strSource = vTemp(2)
                GoTo Found_Source
            End If
        Next j
        
Found_Source:
        If strLine = "" Then
            strLine = vItem(0) & ": " & vItem(1) & ": " & strSource
        Else
            strLine = strLine & vbCr & vItem(0) & ": " & vItem(1) & ": " & strSource
        End If
    Next i
    
    tbPole.Value = strLine
End Sub

Private Sub cbGetCallouts_Click()
    Dim objBlock As AcadBlockReference
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vAtt As Variant
    Dim vResult, vLine, vTemp As Variant
    Dim strPole, strCounts As String
    Dim iPos As Integer
    
    On Error Resume Next
    
    Me.Hide
        
    Err = 0
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    lbList.Clear
    objSSPole.Clear
    
    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH,sMH"
    filterType = grpCode
    filterValue = grpValue
    objSSPole.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    If objSSPole.count < 1 Then GoTo Skip_This:
    
    For i = 0 To objSSPole.count - 1
        Set objBlock = objSSPole.Item(i)
        
        vAtt = objBlock.GetAttributes
        If vAtt(0).TextString = "POLE" Then GoTo Next_Pole
        If vAtt(0).TextString = "" Then GoTo Next_Pole
        
        strPole = vAtt(0).TextString
        
        Select Case objBlock.Name
            Case "sPole"
                iPos = 25
            Case Else
                iPos = 5
        End Select
        If vAtt(iPos).TextString = "" Then GoTo Next_Pole
        
        vLine = Split(vAtt(iPos).TextString, vbCr)
        For j = 0 To UBound(vLine)
            lbList.AddItem strPole
            lbList.List(lbList.ListCount - 1, 1) = UCase(vAtt(iPos).TextString)
            lbList.List(lbList.ListCount - 1, 2) = "none"
            lbList.List(lbList.ListCount - 1, 3) = ""
            lbList.List(lbList.ListCount - 1, 4) = ""
            lbList.List(lbList.ListCount - 1, 5) = ""
            lbList.List(lbList.ListCount - 1, 6) = ""
            lbList.List(lbList.ListCount - 1, 7) = i
            lbList.List(lbList.ListCount - 1, 8) = vAtt(iPos + 2).TextString
        Next j
        
Next_Pole:
    Next i
    
    Call CheckUnits
    
Skip_This:
    objSS.Clear
    
    grpCode(0) = 2
    grpValue(0) = "Callout"
    filterType = grpCode
    filterValue = grpValue
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        
        vAtt = objBlock.GetAttributes
        If InStr(vAtt(1).TextString, "+HO1") > 0 Then GoTo Next_objBlock
        If vAtt(1).TextString = "CALLOUT" Then GoTo Next_objBlock
        
        vLine = Split(vAtt(0).TextString, ": ")
        For j = 0 To lbList.ListCount - 1
            If lbList.List(j, 0) = vLine(0) Then
                vTemp = Split(lbList(j, 1), ": ")
                If vTemp(0) = vLine(1) Then
                    lbList.List(j, 2) = vLine(1) & ": " & UCase(vAtt(1).TextString) & " / " & UCase(vAtt(2).TextString)
                    strCounts = ConvertToCallout(CStr(lbList.List(j, 1)))
                    If strCounts = lbList.List(j, 2) Then lbList.List(j, 3) = "Y"
                    
                    vResult = ValidateCounts(UCase(vAtt(1).TextString), UCase(vAtt(2).TextString))
                    If vResult(0) = "Y" Then
                        lbList.List(j, 4) = "Y"
                    Else
                        lbList.List(j, 4) = ""
                    End If
                    
                    If vResult(1) = "Y" Then
                        lbList.List(j, 5) = "Y"
                    Else
                        lbList.List(j, 5) = ""
                    End If
                    
                    lbList.List(j, 7) = lbList.List(j, 7) & ";;" & i
                    GoTo Next_objBlock
                End If
            End If
        Next j
Next_objBlock:
    Next i
    
Exit_Sub:
    tbListcount.Value = lbList.ListCount
    Me.show
End Sub

Private Sub cbQuit_Click()
    objSS.Clear
    objSS.Delete
    
    Me.Hide
End Sub

Private Sub cbRemove_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    For i = lbList.ListCount - 1 To 0 Step -1
        If lbList.List(i, 3) = "Y" Then
            If lbList.List(i, 4) = "Y" Then
                If lbList.List(i, 5) = "Y" Then
                    If lbList.List(i, 6) = "Y" Then
                        lbList.RemoveItem i
                        GoTo Next_I
                    End If
                End If
            End If
        End If
        
        If lbList.List(i, 2) = "none" Then
            If lbList.List(i, 6) = "Y" Then lbList.RemoveItem i
        End If
Next_I:
    Next i
    
    tbListcount.Value = lbList.ListCount
End Sub

Private Sub cbUpdate_Click()
    strStatus = ""
    
    Call UpdatePole
    Call UpdateCallout
    tbUnits.Value = ""
    
    For i = 0 To lbList.ListCount - 1
        lbList.List(i, 6) = ""
    Next i
    Call CheckUnits
    
    If strStatus = "" Then strStatus = "No updates have been made."
    
    MsgBox strStatus
End Sub

Private Sub cbUpdateUnits_Click()
    If tbUnits.Value = "" Then Exit Sub
    
    Dim objSSCO As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList, vUnits As Variant
    Dim strLayer, strUnits As String
    Dim dInsertPnt(2) As Double
    Dim dScale As Double
    Dim dPosition As Double
    
    On Error Resume Next
    Me.Hide
    
    Set objSSCO = ThisDrawing.SelectionSets.Add("objSSCO")
    If Not Err = 0 Then
        Set objSSCO = ThisDrawing.SelectionSets.Item("objSSCO")
        Err = 0
    End If
    objSSCO.SelectOnScreen
    
    If objSSCO.count < 1 Then GoTo Exit_Sub
    
    dInsertPnt(0) = 0#
    dInsertPnt(1) = 0#
    dInsertPnt(2) = 0#
    
    For Each objBlock In objSSCO
        If objBlock.Name = "pole_unit" Then
            If objBlock.InsertionPoint(1) > dInsertPnt(1) Then
                dInsertPnt(0) = objBlock.InsertionPoint(0)
                dInsertPnt(1) = objBlock.InsertionPoint(1)
                dScale = objBlock.XScaleFactor
            End If
            strLayer = objBlock.Layer
            objBlock.Delete
        End If
    Next objBlock
    
    dPosition = 2#
    
    strUnits = Replace(tbUnits.Value, vbLf, "")
    vUnits = Split(strUnits, vbCr)
    
    For v = 0 To UBound(vUnits)
        strAtt1 = dPosition
        strAtt2 = "N/A"
        strAtt3 = vUnits(v)
    
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
    objSSCO.Clear
    objSSCO.Delete
    
    Me.show
End Sub

Private Sub CommandButton1_Click()
    Dim vLine, vItem, vTemp As Variant
    Dim vPrevious, vCurrent, vCounts As Variant
    Dim strLine, strName As String
    Dim iPF, iPL, iCF, iCL As Integer
    
    strLine = Replace(tbPole.Value, vbLf, "")
    vLine = Split(strLine, vbCr)
    
    If UBound(vLine) > 0 Then
        For i = UBound(vLine) To 1 Step -1
            vCurrent = Split(vLine(i), ": ")
            vPrevious = Split(vLine(i - 1), ": ")
            
            If vCurrent(0) = vPrevious(0) Then
                If vCurrent(2) = vPrevious(2) Then
                    vCounts = Split(vCurrent(1), "-")
                    iCF = CInt(vCounts(0))
                    If UBound(vCounts) > 0 Then
                        iCL = CInt(vCounts(1))
                    Else
                        iCL = iCF
                    End If
                    
                    vCounts = Split(vPrevious(1), "-")
                    iPF = CInt(vCounts(0))
                    If UBound(vCounts) > 0 Then
                        iPL = CInt(vCounts(1))
                    Else
                        iPL = iPF
                    End If
                    
                    If iCF = iPL + 1 Then
                        iPL = iCL
                        strLine = vPrevious(0) & ": " & iPF & "-" & iPL & ": " & vPrevious(2)
                        
                        vLine(i - 1) = strLine
                        vLine(i) = ""
                    End If
                End If
            End If
        Next i
        
        strLine = ""
        For i = 0 To UBound(vLine)
            If Not vLine(i) = "" Then
                If strLine = "" Then
                    strLine = vLine(i)
                Else
                    strLine = strLine & vbCr & vLine(i)
                End If
            End If
        Next i
        
        tbPole.Value = strLine
    End If
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub lbList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim objBlock As AcadBlockReference
    Dim vLine As Variant
    Dim strText As String
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    iIndex = lbList.ListIndex
    
    vLine = Split(lbList.List(iIndex, 7), ";;")
    iPole = CInt(vLine(0))
    If UBound(vLine) > 0 Then
        iCallout = CInt(vLine(1))
    Else
        iCallout = -1
    End If
    cbUpdate.Enabled = True
    'cbUpdatePole.Enabled = True
    
    Me.Hide
    
    Set objBlock = objSSPole.Item(iPole)
    viewCoordsB(0) = objBlock.InsertionPoint(0) - 300
    viewCoordsB(1) = objBlock.InsertionPoint(1) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = viewCoordsB(0) + 600
    viewCoordsE(1) = viewCoordsB(1) + 600
    viewCoordsE(2) = 0#
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    vLine = Split(UCase(lbList.List(lbList.ListIndex, 1)), " / ")
    
    tbPoleCable.Value = vLine(0)
    strText = Replace(vLine(1), " + ", vbCr)
    tbPole.Value = strText
    
    If lbList.List(iIndex, 2) = "" Then Exit Sub
    If lbList.List(iIndex, 2) = "none" Then GoTo Skip_Callout
    
    vLine = Split(UCase(lbList.List(iIndex, 2)), " / ")
    
    tbCable.Value = vLine(0)
    strText = Replace(vLine(1), "\P", vbCr)
    tbCallout.Value = strText
    
Skip_Callout:
    strText = Replace(lbList.List(iIndex, 8), ";;", vbCr)
    If strText = "" Then
        tbUnits.Value = ""
    Else
        tbUnits.Value = strText
    End If
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbList.ColumnCount = 9
    lbList.ColumnWidths = "120;120;120;36;36;36;36;18;6"
    
    iIndex = -1
    
    On Error Resume Next
    
    Set objSSPole = ThisDrawing.SelectionSets.Add("objSSPole")
    If Not Err = 0 Then
        Set objSSPole = ThisDrawing.SelectionSets.Item("objSSPole")
    End If
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    End If
End Sub

Private Function ValidateCounts(strCable As String, strCallout As String)
    Dim strResult, strColor As String
    Dim vResult As Variant
    Dim vLine, vItem, vCount, vTemp As Variant
    
    Dim iCable, iCount, iTCount As Integer
    Dim iLow, iHigh As Integer
    
    vLine = Split(strCable, "(")
    vItem = Split(vLine(1), ")")
    iCable = CInt(vItem(0))
    iCount = 0
    strColor = "Y"
    
    vLine = Split(strCallout, "\P")
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        vCount = Split(vItem(UBound(vItem)), "-")
        
        iTCount = iCount + 1
        While iTCount > 12
            iTCount = iTCount - 12
        Wend
        
        iLow = CInt(vCount(0))
        While iLow > 12
            iLow = iLow - 12
        Wend
        
        If Not iLow = iTCount Then strColor = "N"
        
        If UBound(vCount) = 0 Then
            iCount = iCount + 1
        Else
            iCount = iCount + CInt(vCount(1)) - CInt(vCount(0)) + 1
        End If
    Next i
    
    If iCount = iCable Then
        strResult = "Y"
    Else
        strResult = "N"
    End If
    
    'MsgBox strCable & vbCr & vbCr & strCallout
    
    strResult = strResult & "," & strColor
    vResult = Split(strResult, ",")
    
    ValidateCounts = vResult
End Function

Private Sub UpdatePole()
    If iPole < 0 Then GoTo Exit_Sub
    
    Dim objBlock As AcadBlockReference
    Dim vAtt, vLine As Variant
    Dim strText, strUnits As String
    Dim i, iPos As Integer
    
    strText = Replace(UCase(tbPole.Value), vbLf, "")
    strText = tbPoleCable.Value & " / " & Replace(strText, vbCr, " + ")
    
    strUnits = Replace(tbUnits.Value, vbLf, "")
    strUnits = Replace(strUnits, vbCr, ";;")
    
    lbList.List(iIndex, 1) = strText
    If Not strUnits = "" Then lbList.List(iIndex, 8) = strUnits
    
    Set objBlock = objSSPole.Item(iPole)
    vAtt = objBlock.GetAttributes
    Select Case objBlock.Name
        Case "sPole"
            iPos = 25
        Case Else
            iPos = 5
    End Select
    
    vAtt(iPos).TextString = strText
    vAtt(iPos + 2).TextString = strUnits
    
    objBlock.Update
    
    strStatus = "Pole"
    
Exit_Sub:
    tbPoleCable.Value = ""
    tbPole.Value = ""
    tbUnits.Value = ""
    
    iPole = -1
End Sub

Private Sub UpdateCallout()
    If iCallout < 0 Then GoTo Exit_Sub
    If tbCallout.Value = "" Then GoTo Exit_Sub
    
    Dim objBlock As AcadBlockReference
    Dim vAtt, vLine As Variant
    Dim strText As String
    Dim i As Integer
    
    strText = Replace(UCase(tbCallout.Value), vbLf, "")
    strText = Replace(strText, vbCr, "\P")
    
    lbList.List(iIndex, 2) = UCase(tbCable.Value) & " / " & strText
    
    vLine = Split(lbList.List(iIndex, 7), ";;")
    i = CInt(vLine(1))
    Set objBlock = objSS.Item(i)
    vAtt = objBlock.GetAttributes
    
    vLine = Split(UCase(tbCable.Value), ": ")
    tbCable.Value = vLine(1)
    vAtt(1).TextString = tbCable.Value
    vAtt(2).TextString = strText
    
    objBlock.Update
        
    vResult = ValidateCounts(tbCable.Value, strText)
    If vResult(0) = "Y" Then lbList.List(iIndex, 3) = "Y"
    If vResult(1) = "Y" Then lbList.List(iIndex, 4) = "Y"
    
    If strStatus = "" Then
        strStatus = "Callout has been updated."
    Else
        strStatus = strStatus & " and Callout have been updated."
    End If
    
Exit_Sub:
    tbCable.Value = ""
    tbCallout.Value = ""
    
    iCallout = -1
    cbUpdate.Enabled = False
End Sub

Private Function ConvertToCallout(strText As String)
    Dim vLine, vItem, vTemp As Variant
    Dim strLine As String
    
    vLine = Split(strText, " / ")
    vItem = Split(vLine(1), " + ")
    
    For i = 0 To UBound(vItem)
        vTemp = Split(vItem(i), ": ")
        vItem(i) = vTemp(0) & ": " & vTemp(1)
    Next i
    
    strLine = ""
    
    For i = 0 To UBound(vItem)
        If strLine = "" Then
            strLine = vItem(i)
        Else
            strLine = strLine & "\P" & vItem(i)
        End If
    Next i
    
    strLine = vLine(0) & " / " & strLine
    ConvertToCallout = strLine
End Function

Private Sub CheckUnits()
    If lbList.ListCount < 1 Then Exit Sub
    
    Dim vUnit, vLine, vItem, vTemp As Variant
    Dim strCable As String
    
    For i = 0 To lbList.ListCount - 1
        If lbList.List(i, 8) = "" Then GoTo Next_I
        
        vLine = Split(lbList.List(i, 1), " / ")
        vItem = Split(vLine(0), ": ")
        strCable = vItem(1)
        
        vUnit = Split(lbList.List(i, 8), ";;")
        For j = 0 To UBound(vUnit)
            If InStr(vUnit(j), strCable) > 0 Then
                If InStr(vUnit(j), "LOOP") = 0 Then
                    lbList.List(i, 6) = "Y"
                End If
            End If
        Next j
        
Next_I:
    Next i
End Sub
