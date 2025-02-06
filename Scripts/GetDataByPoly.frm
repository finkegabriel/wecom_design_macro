VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetDataByPoly 
   Caption         =   "Get Data by Polygon"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "GetDataByPoly.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GetDataByPoly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dCoords() As Double
    
Private Sub cbClearTab_Click()
    lbTab.Clear
    
    For i = 0 To lbBlocks.ListCount - 1
        lbBlocks.List(i, 2) = Replace(lbBlocks.List(i, 2), "+ ", "")
    Next i
End Sub

Private Sub cbGetPolygon_Click()
    Dim objEntity As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim vReturnPnt As Variant
    Dim vCoords As Variant
    Dim iTemp, iCounter As Integer
    
    Me.Hide
    
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Polygon: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not objEntity.ObjectName = "AcDbPolyline" Then
        MsgBox "Error: Invalid Selection."
        GoTo Exit_Sub
    End If
    
    Set objLWP = objEntity
    vCoords = objLWP.Coordinates
    
    iTemp = (UBound(vCoords) + 1) / 2 * 3 - 1
    ReDim dCoords(iTemp) As Double
    
    iCounter = 0
    For i = 0 To UBound(vCoords) Step 2
        dCoords(iCounter) = vCoords(i)
        iCounter = iCounter + 1
        dCoords(iCounter) = vCoords(i + 1)
        iCounter = iCounter + 1
        dCoords(iCounter) = 0#
        iCounter = iCounter + 1
    Next i
    
    For i = 0 To lbLayers.ListCount - 1
        Call GetLayerData(CStr(lbLayers.List(i)))
    Next i
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbSort_Click()
    Dim iCount As Integer
    Dim iIndex As Integer
    Dim strAtt(1) As String
    
    iCount = lbTab.ListCount - 1
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbTab.List(b, 0) > lbTab.List(b + 1, 0) Then
                strAtt(0) = lbTab.List(b + 1, 0)
                strAtt(1) = lbTab.List(b + 1, 1)
                
                lbTab.List(b + 1, 0) = lbTab.List(b, 0)
                lbTab.List(b + 1, 1) = lbTab.List(b, 1)
                
                lbTab.List(b, 0) = strAtt(0)
                lbTab.List(b, 1) = strAtt(1)
            End If
        Next b
    Next a
End Sub

Private Sub lbBlocks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbBlocks.ListCount < 1 Then Exit Sub
    If InStr(lbBlocks.List(lbBlocks.ListIndex, 2), "+ ") > 0 Then Exit Sub
    
    Dim objSS As AcadSelectionSet
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim filterType, filterValue As Variant
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vList, vItem As Variant
    Dim strBlock, strLayer As String
    Dim strLine As String
    
    strLayer = lbBlocks.List(lbBlocks.ListIndex, 0)
    strBlock = lbBlocks.List(lbBlocks.ListIndex, 1)
    lbBlocks.List(lbBlocks.ListIndex, 2) = "+ " & lbBlocks.List(lbBlocks.ListIndex, 2)
    
    grpCode(0) = 2
    grpValue(0) = strBlock
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Add("objSS")
        objSS.Clear
        Err = 0
    End If
    
    objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords, filterType, filterValue
    
    For Each objBlock In objSS
        If Not objBlock.Layer = strLayer Then GoTo Next_objBlock
        
        vAttList = objBlock.GetAttributes
        Select Case objBlock.Name
            Case "sPole"
                strLine = vAttList(27).TextString
                
                Call GetStructureInfo(CStr(strLine))
            Case "sPed", "sHH", "sMH", "sPanel", "sFP"
                strLine = vAttList(7).TextString
                
                Call GetStructureInfo(CStr(strLine))
            Case "Customer"
                strLine = "<> " & vAttList(0).TextString
                
                Call GetCustomerInfo(CStr(strLine))
            Case "cable_span"
                strLine = vAttList(1).TextString & ";;" & Replace(vAttList(2).TextString, "'", "")
                
                Call GetCableSpanInfo(CStr(strLine))
            Case "Map coil"
                strLine = "CO(" & Replace(vAttList(1).TextString, "F", "") & ")COIL;;" & Replace(vAttList(0).TextString, "'", "")
                
                Call GetCableSpanInfo(CStr(strLine))
            Case "Map splice"
                strLine = "* " & vAttList(0).TextString
                
                Call GetCustomerInfo(CStr(strLine))
            Case "ExGuyOL", "ExGuyOR"
                strLine = "* PE1"
                If InStr(LCase(strLayer), "remove") > 0 Then strLine = "* XXPE1"
                If InStr(LCase(strLayer), "exist") > 0 Then GoTo Next_objBlock
                
                Call GetCustomerInfo(CStr(strLine))
            Case "ExAncOL", "ExAncOR"
                strLine = "* PF1"
                If InStr(LCase(strLayer), "remove") > 0 Then strLine = "* XXPF1"
                If InStr(LCase(strLayer), "exist") > 0 Then GoTo Next_objBlock
                
                Call GetCustomerInfo(CStr(strLine))
        End Select
        
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub

Private Sub UserForm_Initialize()
    lbLines.ColumnCount = 6
    lbLines.ColumnWidths = "120;36;72;36;72;72"
    
    lbBlocks.ColumnCount = 3
    lbBlocks.ColumnWidths = "120;120;48"
    
    lbTab.ColumnCount = 2
    lbTab.ColumnWidths = "132;54"
    
    Dim objLayer As AcadLayer
    For Each objLayer In ThisDrawing.Layers
        lbLayers.AddItem objLayer.Name
    Next objLayer
    
    If lbLayers.ListCount > 1 Then
        Dim iCount As Integer
        Dim iIndex As Integer
        Dim strAtt As String
    
        iCount = lbLayers.ListCount - 1
    
        On Error Resume Next
    
        For a = iCount To 0 Step -1
            For b = 0 To a - 1
                If lbLayers.List(b, 0) > lbLayers.List(b + 1, 0) Then
                    strAtt = lbLayers.List(b + 1, 0)
                    lbLayers.List(b + 1, 0) = lbLayers.List(b, 0)
                    lbLayers.List(b, 0) = strAtt
                End If
            Next b
        Next a
    End If
End Sub

Private Sub GetLayerData(strLayer As String)
    Dim objSS As AcadSelectionSet
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim filterType, filterValue As Variant
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim objLine As AcadLine
    Dim objLWP As AcadLWPolyline
    Dim vAttList, vLine As Variant
    Dim vReturnPnt As Variant
    
    Dim vCoords, vArray As Variant
    Dim strTemp As String
    Dim iTemp, iCounter As Integer
    Dim lLine, lLWP, lBlock, lTotal As Long
    Dim lNLine, lNLWP As Long
    
    lLine = 0
    lLWP = 0
    lBlock = 0
    lNLine = 0
    lNLWP = 0
    
    grpCode(0) = 8
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Add("objSS")
        objSS.Clear
        Err = 0
    End If
    
    objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords, filterType, filterValue
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadLWPolyline Then
            Set objLWP = objEntity
            
            lLWP = lLWP + CLng(objLWP.Length)
            lNLWP = lNLWP + 1
        ElseIf TypeOf objEntity Is AcadLine Then
            Set objLine = objEntity
            
            lLine = lLine + CLng(objLine.Length)
            lNLine = lNLine + 1
        ElseIf TypeOf objEntity Is AcadBlockReference Then
            Set objBlock = objEntity
            
            If lbBlocks.ListCount < 1 Then
                lbBlocks.AddItem strLayer
                lbBlocks.List(0, 1) = objBlock.Name
                lbBlocks.List(0, 2) = "1"
                
                GoTo Next_objEntity
            End If
            
            For i = 0 To lbBlocks.ListCount - 1
                If lbBlocks.List(i, 0) = strLayer Then
                    If lbBlocks.List(i, 1) = objBlock.Name Then
                        lbBlocks.List(i, 2) = CInt(lbBlocks.List(i, 2)) + 1
                        
                        GoTo Next_objEntity
                    End If
                End If
            Next i
            
            lbBlocks.AddItem strLayer
            lbBlocks.List(lbBlocks.ListCount - 1, 1) = objBlock.Name
            lbBlocks.List(lbBlocks.ListCount - 1, 2) = "1"
        End If
Next_objEntity:
    Next objEntity
    
    lTotal = lLine + lLWP
    If lTotal > 0 Then
        lbLines.AddItem strLayer
        lbLines.List(lbLines.ListCount - 1, 1) = lNLine
        lbLines.List(lbLines.ListCount - 1, 2) = lLine
        lbLines.List(lbLines.ListCount - 1, 3) = lNLWP
        lbLines.List(lbLines.ListCount - 1, 4) = lLWP
        lbLines.List(lbLines.ListCount - 1, 5) = lTotal
    End If
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub GetStructureInfo(strLine As String)
    If strLine = "" Then Exit Sub
    
    Dim vList, vItem, vTemp As Variant
    
    vList = Split(strLine, ";;")
    For i = 0 To UBound(vList)
        vItem = Split(vList(i), "=")
        vItem(0) = Replace(vItem(0), "+", "")
        vItem(1) = Replace(vItem(1), "'", "")
        
        If InStr(vItem(1), " ") > 0 Then
            vTemp = Split(vItem(1), "  ")
            vItem(0) = vItem(0) & vbTab & vTemp(UBound(vTemp))
            vItem(1) = vTemp(0)
        Else
            vItem(0) = vItem(0) & vbTab
        End If
        
        If lbTab.ListCount < 1 Then
            lbTab.AddItem vItem(0)
            lbTab.List(0, 1) = vItem(1)
            
            GoTo Next_I
        End If
        
        For j = 0 To lbTab.ListCount - 1
            If vItem(0) = lbTab.List(j, 0) Then
                lbTab.List(j, 1) = CLng(lbTab.List(j, 1)) + CLng(vItem(1))
                GoTo Next_I
            End If
        Next j
        
        lbTab.AddItem vItem(0)
        lbTab.List(lbTab.ListCount - 1, 1) = vItem(1)
        
Next_I:
    Next i
End Sub

Private Sub GetCustomerInfo(strLine As String)
    If strLine = "" Then Exit Sub
    
    If lbTab.ListCount < 1 Then
        lbTab.AddItem strLine
        lbTab.List(0, 1) = 1
        
        Exit Sub
    End If
    
    For i = 0 To lbTab.ListCount - 1
        If strLine = lbTab.List(i, 0) Then
            lbTab.List(i, 1) = CInt(lbTab.List(i, 1)) + 1
            
            Exit Sub
        End If
    Next i
    
    lbTab.AddItem strLine
    lbTab.List(lbTab.ListCount - 1, 1) = 1
End Sub

Private Sub GetCableSpanInfo(strLine As String)
    If strLine = "" Then Exit Sub
    
    Dim vItem, vCables As Variant
    Dim strCable As String
    
    vItem = Split(strLine, ";;")
    vCables = Split(vItem(0), " ")
    
    For i = 0 To UBound(vCables)
        strCable = "* " & vCables(i)
        
        If lbTab.ListCount < 1 Then
            lbTab.AddItem strCable
            lbTab.List(0, 1) = vItem(1)
            
            GoTo Next_I
        End If
        
        For j = 0 To lbTab.ListCount - 1
            If lbTab.List(j, 0) = strCable Then
                lbTab.List(j, 1) = CLng(lbTab.List(j, 1)) + CInt(vItem(1))
                
                GoTo Next_I
            End If
        Next j
        
        lbTab.AddItem strCable
        lbTab.List(lbTab.ListCount - 1, 1) = vItem(1)
Next_I:
    Next i
End Sub
