VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} aaPlanningWorksheet 
   Caption         =   "Planning Worksheet"
   ClientHeight    =   8895.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18990
   OleObjectBlob   =   "aaPlanningWorksheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "aaPlanningWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSS As AcadSelectionSet
    
Private Sub cbAddSB_Click()
    If lbStructures.ListCount < 1 Then Exit Sub
    
    Dim strLine, strTemp As String
    Dim objBlock As AcadBlockReference
    Dim vAttList, vLine, vItem As Variant
    
    If tbNewBoundary.Value = "" Then
        strLine = " "
    Else
        strLine = tbNewBoundary.Value
    End If
    
    For i = 0 To lbStructures.ListCount - 1
        If lbStructures.Selected(i) = True Then
            lbStructures.List(i, 8) = strLine
            
            Set objBlock = objSS.Item(CInt(lbStructures.List(i, 0)))
            vAttList = objBlock.GetAttributes
            
            vLine = Split(vAttList(28).TextString, ";;")
            vLine(7) = strLine
            
            strTemp = vLine(0)
            For j = 1 To UBound(vLine)
                strTemp = strTemp & ";;" & vLine(j)
            Next j
            vAttList(28).TextString = strTemp
            objBlock.Update
        End If
    Next i
    
    For i = 0 To cbBoundary.ListCount - 1
        If cbBoundary.List(i) = strLine Then Exit Sub
    Next i
    
    cbBoundary.AddItem strLine
End Sub

Private Sub cbAddSplitter_Click()
    If lbStructures.ListCount < 1 Then Exit Sub
    If Not cbBoundary.Value = "All" Then
        MsgBox "Need to display All boundaries"
        Exit Sub
    End If
    
    Dim objBlock As AcadBlockReference
    Dim vAttList, vLine, vItem As Variant
    Dim strTemp As String
    Dim strLine, strPrevious As String
    
    strPrevious = ""
    
    For i = 0 To lbStructures.ListCount - 1
        If lbStructures.Selected(i) = True Then
            strLine = Replace(lbStructures.List(i, 1), " - ", "")
            lbStructures.List(i, 5) = cbSplitter.Value
            lbStructures.List(i, 8) = strLine
            tbNewBoundary.Value = strLine
            
            strPrevious = lbStructures.List(i, 2)
        
            Set objBlock = objSS.Item(CInt(lbStructures.List(i, 0)))
            vAttList = objBlock.GetAttributes
            
            vLine = Split(vAttList(28).TextString, ";;")
            If vLine(3) = " " Then
                vLine(3) = "0:1:0"
            Else
                vItem = Split(vLine(3), ":")
                Select Case UBound(vItem)
                    Case Is = 0
                        vLine(3) = vLine(3) & ":1:0"
                    Case Is = 1
                        vLine(3) = vItem(0) & ":" & CInt(vItem(1)) + 1 & ":0"
                    Case Is > 1
                        vLine(3) = vItem(0) & ":" & CInt(vItem(1)) + 1 & ":" & vItem(2)
                End Select
            End If
            vLine(4) = cbSplitter.Value
            vLine(7) = strLine
            
            lbStructures.List(i, 4) = vLine(3)
            
            strTemp = vLine(0)
            For j = 1 To UBound(vLine)
                strTemp = strTemp & ";;" & vLine(j)
            Next j
            vAttList(28).TextString = strTemp
            objBlock.Update
            
            GoTo Exit_Sub
        End If
    Next i
    
Exit_Sub:
    
    If strPrevious = "" Then Exit Sub
    
    For i = 0 To lbStructures.ListCount - 1
        If lbStructures.List(i, 1) = strPrevious Then
            Set objBlock = objSS.Item(CInt(lbStructures.List(i, 0)))
            vAttList = objBlock.GetAttributes
            
            vLine = Split(vAttList(28).TextString, ";;")
            If vLine(3) = " " Then
                vLine(3) = "0:1:0"
            Else
                vItem = Split(vLine(3), ":")
                Select Case UBound(vItem)
                    Case Is = 0
                        vLine(3) = vLine(3) & ":1:0"
                    Case Is = 1
                        vLine(3) = vItem(0) & ":" & CInt(vItem(1)) + 1 & ":0"
                    Case Is > 1
                        vLine(3) = vItem(0) & ":" & CInt(vItem(1)) + 1 & ":" & vItem(2)
                End Select
            End If
            
            lbStructures.List(i, 4) = vLine(3)
            
            strTemp = vLine(0)
            For j = 1 To UBound(vLine)
                strTemp = strTemp & ";;" & vLine(j)
            Next j
            vAttList(28).TextString = strTemp
            objBlock.Update
            
            strPrevious = lbStructures.List(i, 2)
            
            i = -1
        End If
    Next i
End Sub

Private Sub cbBoundary_Change()
    Call GetList
End Sub

Private Sub cbGetStructures_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine As Variant
    Dim vPnt1, vPnt2 As Variant
    'Dim strPole, strOwner, strCompany, strNew As String
    'Dim strExisting, strProposed, strExtra, strNote As String
    'Dim strAttachments As String
    'Dim strDWG As String
    Dim iIndex As Integer
    'Dim iExist, iProp, iPole As Integer
    
    On Error Resume Next
    
    Me.Hide
    
    lbStructures.Clear
    'lbCompany.Clear
    'cbCompany.Clear
    'tbOwners.Value = ""
    'strDWG = "Window"
    'iPole = 0
        
    Err = 0
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    objSS.Clear
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    iIndex = 0
    cbBoundary.Clear
    cbBoundary.AddItem "All"
    cbBoundary.AddItem " "
    cbBoundary.Value = "All"
    
'    For Each objBlock In objSS
'        vAttList = objBlock.GetAttributes
'        If vAttList(28).TextString = "" Then GoTo Next_objBlock
'
'        vLine = Split(vAttList(28).TextString, ";;")
'        lbStructures.AddItem iIndex
'        lbStructures.List(iIndex, 1) = vLine(0)
'        lbStructures.List(iIndex, 2) = vLine(1)
'        lbStructures.List(iIndex, 3) = vLine(2)
'        lbStructures.List(iIndex, 4) = vLine(3)
'        lbStructures.List(iIndex, 5) = vLine(4)
'        lbStructures.List(iIndex, 6) = vLine(5)
'        lbStructures.List(iIndex, 7) = vLine(6)
'        lbStructures.List(iIndex, 8) = vLine(7)
'        lbStructures.List(iIndex, 9) = vLine(8)
'
'        iIndex = iIndex + 1
'Next_objBlock:
'    Next objBlock
    
    If lbStructures.ListCount > 0 Then
        For i = 0 To lbStructures.ListCount - 1
            For j = 0 To cbBoundary.ListCount - 1
                If cbBoundary.List(j) = lbStructures.List(i, 8) Then GoTo Next_Structure
            Next j
            
            cbBoundary.AddItem lbStructures.List(i, 8)
            
Next_Structure:
        Next i
    End If
    
Exit_Sub:
    
    tbListCount.Value = lbStructures.ListCount
    
    Call SortList
    
    Me.show
End Sub

Private Sub cbTraceStructure_Click()
    If lbStructures.ListCount < 1 Then Exit Sub
    If Not cbBoundary.Value = "All" Then
        MsgBox "Need to display All boundaries"
        Exit Sub
    End If
    If tbTrace.Value = "" Then
        MsgBox "Need to enter a structure number"
        Exit Sub
    End If
    
    Dim objBlock As AcadBlockReference
    Dim vAttList, vLine, vItem As Variant
    Dim strTemp As String
    Dim strLine, strPrevious As String
    
    strPrevious = tbTrace.Value
    lbStructures.Clear
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        vAttList = objBlock.GetAttributes
        If InStr(vAttList(28).TextString, ";;") > 0 Then
            vLine = Split(vAttList(28).TextString, ";;")
            If vLine(0) = strPrevious Then
                lbStructures.AddItem i, 0
                lbStructures.List(0, 1) = vLine(0)
                lbStructures.List(0, 2) = vLine(1)
                lbStructures.List(0, 3) = vLine(2)
                lbStructures.List(0, 4) = vLine(3)
                lbStructures.List(0, 5) = vLine(4)
                lbStructures.List(0, 6) = vLine(5)
                lbStructures.List(0, 7) = vLine(6)
                lbStructures.List(0, 8) = vLine(7)
                lbStructures.List(0, 9) = vLine(8)
                
                strPrevious = vLine(1)
                i = -1
            End If
        End If
    Next i
    
    Call GetTotals
End Sub

Private Sub Label60_Click()
    If lbStructures.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbStructures.ListCount - 1
        If lbStructures.Selected(i) = True Then
            cbBoundary.Value = lbStructures.List(i, 8)
            Exit Sub
        End If
    Next i
End Sub

Private Sub Label62_Click()
    If lbStructures.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbStructures.ListCount - 1
        If lbStructures.Selected(i) = True Then
            tbNewBoundary.Value = lbStructures.List(i, 8)
            Exit Sub
        End If
    Next i
End Sub

Private Sub Label66_Click()
    If lbStructures.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbStructures.ListCount - 1
        If lbStructures.Selected(i) = True Then
            tbTrace.Value = lbStructures.List(i, 1)
            Exit Sub
        End If
    Next i
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub lbStructures_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbStructures.ListIndex < 0 Then Exit Sub
    
    Dim iIndex As Integer
    Dim vLine As Variant
    Dim strLine As String
    
    iIndex = lbStructures.ListIndex
    
    Me.Hide
    
    Load aaStructureData
        aaStructureData.tbNumber.Value = lbStructures.List(iIndex, 1)
        aaStructureData.tbPrevious.Value = lbStructures.List(iIndex, 2)
        aaStructureData.tbDistance.Value = lbStructures.List(iIndex, 3)
        
        vLine = Split(lbStructures.List(iIndex, 4), ":")
        Select Case UBound(vLine)
            Case Is = 0
                aaStructureData.tbTrunk.Value = vLine(0)
            Case Is = 1
                aaStructureData.tbTrunk.Value = vLine(0)
                aaStructureData.tbF1.Value = vLine(1)
            Case Is > 1
                aaStructureData.tbTrunk.Value = vLine(0)
                aaStructureData.tbF1.Value = vLine(1)
                aaStructureData.tbF15.Value = vLine(2)
        End Select
        aaStructureData.tbSplitter.Value = lbStructures.List(iIndex, 5)
        
        vLine = Split(lbStructures.List(iIndex, 6), ",")
        aaStructureData.tbRES.Value = vLine(0)
        If UBound(vLine) > 0 Then aaStructureData.tbBUS.Value = vLine(1)
        
        aaStructureData.tbTaps.Value = lbStructures.List(iIndex, 7)
        aaStructureData.tbBoundary.Value = lbStructures.List(iIndex, 8)
        aaStructureData.tbCable.Value = lbStructures.List(iIndex, 9)
        
        aaStructureData.show
        
        If aaStructureData.iSave = 0 Then GoTo Exit_Sub
        
        lbStructures.List(iIndex, 1) = aaStructureData.tbNumber.Value
        lbStructures.List(iIndex, 2) = aaStructureData.tbPrevious.Value
        lbStructures.List(iIndex, 3) = aaStructureData.tbDistance.Value
        strLine = aaStructureData.tbTrunk.Value & ":" & aaStructureData.tbF1.Value & ":" & aaStructureData.tbF15.Value
        lbStructures.List(iIndex, 4) = strLine
        lbStructures.List(iIndex, 5) = aaStructureData.tbSplitter.Value
        strLine = aaStructureData.tbRES.Value & "," & aaStructureData.tbBUS.Value
        lbStructures.List(iIndex, 6) = strLine
        lbStructures.List(iIndex, 7) = aaStructureData.tbTaps.Value
        lbStructures.List(iIndex, 8) = aaStructureData.tbBoundary.Value
        lbStructures.List(iIndex, 9) = aaStructureData.tbCable.Value
        
    Unload aaStructureData
    
    strLine = lbStructures.List(iIndex, 1)
    For i = 1 To 9
        strLine = strLine & ";;" & lbStructures.List(iIndex, i)
    Next i
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    Set objBlock = objSS.Item(CInt(lbStructures.List(iIndex, 0)))
    vAttList = objBlock.GetAttributes
    vAttList(28).TextString = strLine
    objBlock.Update
    
Exit_Sub:
    
    Me.show
End Sub

Private Sub UserForm_Deactivate()
    objSS.Clear
    objSS.Delete
End Sub

Private Sub UserForm_Initialize()
    lbStructures.ColumnCount = 10
    lbStructures.ColumnWidths = "36;120;120;24;48;36;72;144;96;234"
    
    cbSplitter.AddItem "2"
    cbSplitter.AddItem "4"
    cbSplitter.AddItem "8"
    cbSplitter.AddItem "16"
    cbSplitter.AddItem "32"
    cbSplitter.AddItem "64"
    'cbSplitter.AddItem "2:16"
    'cbSplitter.AddItem "2:32"
    'cbSplitter.AddItem "0:16"
    'cbSplitter.AddItem "0:32"
    
    'cbSplitter.AddItem "4:8"
    'cbSplitter.AddItem "4:16"
    'cbSplitter.AddItem "8:2"
    'cbSplitter.AddItem "8:4"
    'cbSplitter.AddItem "8:8"
    'cbSplitter.AddItem "16:2"
    'cbSplitter.AddItem "16:4"
    
    cbSplitter.Value = "32"
    
    On Error Resume Next
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    End If
End Sub

Private Sub GetList()
    If cbBoundary.Value = "" Then Exit Sub
    If cbBoundary.ListCount < 1 Then Exit Sub
    If objSS.count < 1 Then Exit Sub
    
    Dim objBlock As AcadBlockReference
    Dim vAttList, vItem As Variant
    
    lbStructures.Clear
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        vAttList = objBlock.GetAttributes
        
        If vAttList(28).TextString = "" Then GoTo Next_Object
        vItem = Split(vAttList(28).TextString, ";;")
        
        If Not UBound(vItem) = 8 Then GoTo Next_Object
        If Not cbBoundary.Value = "All" Then
            If Not vItem(7) = cbBoundary.Value Then GoTo Next_Object
        End If
        
        lbStructures.AddItem i
        lbStructures.List(lbStructures.ListCount - 1, 1) = vItem(0)
        lbStructures.List(lbStructures.ListCount - 1, 2) = vItem(1)
        lbStructures.List(lbStructures.ListCount - 1, 3) = vItem(2)
        lbStructures.List(lbStructures.ListCount - 1, 4) = vItem(3)
        lbStructures.List(lbStructures.ListCount - 1, 5) = vItem(4)
        lbStructures.List(lbStructures.ListCount - 1, 6) = vItem(5)
        lbStructures.List(lbStructures.ListCount - 1, 7) = vItem(6)
        lbStructures.List(lbStructures.ListCount - 1, 8) = vItem(7)
        lbStructures.List(lbStructures.ListCount - 1, 9) = vItem(8)
        
Next_Object:
    Next i
    
    tbListCount = lbStructures.ListCount
    
    Call SortList
    Call GetTotals
End Sub

Private Sub SortList()
    If lbStructures.ListCount < 1 Then Exit Sub
    
    Dim vNumber, vL, vR As Variant
    Dim strRoute, strPole As String
    Dim strRoute1, strPole1 As String
    Dim strAtt(9) As String
    Dim iCount As Integer
    Dim dPole, dPole1 As Double
    
    On Error Resume Next
    
    iCount = lbStructures.ListCount - 1
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            vNumber = Split(lbStructures.List(b, 1), " - ")
            strRoute = vNumber(0)
            strPole = vNumber(1)
            dPole = CDbl(strPole)
            
            vNumber = Split(lbStructures.List(b + 1, 1), " - ")
            strRoute1 = vNumber(0)
            strPole1 = vNumber(1)
            dPole1 = CDbl(strPole1)
            
            If strRoute > strRoute1 Then
                    strAtt(0) = lbStructures.List(b + 1, 0)
                    strAtt(1) = lbStructures.List(b + 1, 1)
                    strAtt(2) = lbStructures.List(b + 1, 2)
                    strAtt(3) = lbStructures.List(b + 1, 3)
                    strAtt(4) = lbStructures.List(b + 1, 4)
                    strAtt(5) = lbStructures.List(b + 1, 5)
                    strAtt(6) = lbStructures.List(b + 1, 6)
                    strAtt(7) = lbStructures.List(b + 1, 7)
                    strAtt(8) = lbStructures.List(b + 1, 8)
                    strAtt(9) = lbStructures.List(b + 1, 9)
                
                    lbStructures.List(b + 1, 0) = lbStructures.List(b, 0)
                    lbStructures.List(b + 1, 1) = lbStructures.List(b, 1)
                    lbStructures.List(b + 1, 2) = lbStructures.List(b, 2)
                    lbStructures.List(b + 1, 3) = lbStructures.List(b, 3)
                    lbStructures.List(b + 1, 4) = lbStructures.List(b, 4)
                    lbStructures.List(b + 1, 5) = lbStructures.List(b, 5)
                    lbStructures.List(b + 1, 6) = lbStructures.List(b, 6)
                    lbStructures.List(b + 1, 7) = lbStructures.List(b, 7)
                    lbStructures.List(b + 1, 8) = lbStructures.List(b, 8)
                    lbStructures.List(b + 1, 9) = lbStructures.List(b, 9)
                
                    lbStructures.List(b, 0) = strAtt(0)
                    lbStructures.List(b, 1) = strAtt(1)
                    lbStructures.List(b, 2) = strAtt(2)
                    lbStructures.List(b, 3) = strAtt(3)
                    lbStructures.List(b, 4) = strAtt(4)
                    lbStructures.List(b, 5) = strAtt(5)
                    lbStructures.List(b, 6) = strAtt(6)
                    lbStructures.List(b, 7) = strAtt(7)
                    lbStructures.List(b, 8) = strAtt(8)
                    lbStructures.List(b, 9) = strAtt(9)
            ElseIf strRoute = strRoute1 Then
                If dPole > dPole1 Then
                'If strPole > strPole1 Then
                        strAtt(0) = lbStructures.List(b + 1, 0)
                        strAtt(1) = lbStructures.List(b + 1, 1)
                        strAtt(2) = lbStructures.List(b + 1, 2)
                        strAtt(3) = lbStructures.List(b + 1, 3)
                        strAtt(4) = lbStructures.List(b + 1, 4)
                        strAtt(5) = lbStructures.List(b + 1, 5)
                        strAtt(6) = lbStructures.List(b + 1, 6)
                        strAtt(7) = lbStructures.List(b + 1, 7)
                        strAtt(8) = lbStructures.List(b + 1, 8)
                        strAtt(9) = lbStructures.List(b + 1, 9)
                
                        lbStructures.List(b + 1, 0) = lbStructures.List(b, 0)
                        lbStructures.List(b + 1, 1) = lbStructures.List(b, 1)
                        lbStructures.List(b + 1, 2) = lbStructures.List(b, 2)
                        lbStructures.List(b + 1, 3) = lbStructures.List(b, 3)
                        lbStructures.List(b + 1, 4) = lbStructures.List(b, 4)
                        lbStructures.List(b + 1, 5) = lbStructures.List(b, 5)
                        lbStructures.List(b + 1, 6) = lbStructures.List(b, 6)
                        lbStructures.List(b + 1, 7) = lbStructures.List(b, 7)
                        lbStructures.List(b + 1, 8) = lbStructures.List(b, 8)
                        lbStructures.List(b + 1, 9) = lbStructures.List(b, 9)
                
                        lbStructures.List(b, 0) = strAtt(0)
                        lbStructures.List(b, 1) = strAtt(1)
                        lbStructures.List(b, 2) = strAtt(2)
                        lbStructures.List(b, 3) = strAtt(3)
                        lbStructures.List(b, 4) = strAtt(4)
                        lbStructures.List(b, 5) = strAtt(5)
                        lbStructures.List(b, 6) = strAtt(6)
                        lbStructures.List(b, 7) = strAtt(7)
                        lbStructures.List(b, 8) = strAtt(8)
                        lbStructures.List(b, 9) = strAtt(9)
                End If
            End If
        Next b
    Next a
End Sub

Private Sub GetTotals()
    If lbStructures.ListCount < 1 Then Exit Sub
    
    Dim vLine, vItem As Variant
    Dim iRes, iBus As Integer
    Dim iCoil, iCoilSize As Integer
    Dim lLength As Long
    
    iRes = 0
    iBus = 0
    iCoil = 1
    lLength = 0
    
    iCoilSize = CInt(tbACoilSize.Value)
    
    For i = 0 To lbStructures.ListCount - 1
        If Not lbStructures.List(i, 3) = "" Then lLength = lLength + CLng(lbStructures.List(i, 3))
        
        If InStr(lbStructures.List(i, 6), ",") > 0 Then
            vItem = Split(lbStructures.List(i, 6), ",")
            iRes = iRes + CInt(vItem(0))
            iBus = iBus + CInt(vItem(1))
        End If
        
        If InStr(lbStructures.List(i, 6), ",") > 0 Or Not lbStructures.List(i, 5) = " " Then iCoil = iCoil + 1
    Next i
    
    tbListCount = lbStructures.ListCount
    tbTotalLength.Value = lLength
    tbResBus.Value = iRes & "," & iBus
    tbNumberACoil.Value = iCoil
    
    tbLenCoil.Value = lLength + (iCoil * iCoilSize)
End Sub
