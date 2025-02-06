VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertBlocks 
   Caption         =   "Convert Blocks"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   OleObjectBlob   =   "ConvertBlocks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConvertBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSS As AcadSelectionSet

Private Sub cbConvert_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    
    Dim objEntity As AcadEntity
    Dim objPoint As AcadPoint
    Dim objBlock As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    Dim vBlockAtt As Variant
    
    Dim objPlace As AcadBlockReference
    Dim vAttList As Variant
    Dim dInsertPnt(0 To 2) As Double
    Dim dScale As Double
    
    Dim amap As AcadMap
    Dim ODRcs As ODRecords
    Dim ODRc As ODRecord
    Dim tbl As ODTable
    Dim tbls As ODTables
    
    Dim strAttList() As String
    Dim strAttTemp() As String
    Dim strLine, strTemp As String
    Dim str1, str2, str3 As String
    Dim iCount, iTest As Integer
    Dim iTemp As Integer
    Dim vLine, vTemp, vForm As Variant
    Dim vData As Variant
    
    Dim vPnt1, vPnt2 As Variant
    Dim dCoords() As Double
    Dim vReturnPnt, vCoords As Variant
    Dim iCounter As Integer
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    
    On Error Resume Next
    
    ReDim strAttList(lbTo.ListCount - 1)
    ReDim strAttTemp(lbTo.ListCount - 1)
    
    iCount = 0
    strLine = ""
    dScale = CDbl(cbScale.Value)
    
    For i = 0 To UBound(strAttList)
        strAttList(i) = lbTo.List(i, 2)
    Next i
    
    For i = 0 To UBound(strAttList)
        If InStr(strAttList(i), "{") > 0 Then
            iCount = iCount + 1
        End If
    Next i
    
    If cbFromType.Value = "Block" Then
        grpCode(0) = 2
    Else
        grpCode(0) = 8
    End If
    
    grpValue(0) = cbFromList.Value
    filterType = grpCode
    filterValue = grpValue
            
    objSS.Clear
    
    Err = 0
    
    Me.Hide
    
    Select Case cbSelection.Value
        Case "All"
            objSS.Select acSelectionSetAll, , , filterType, filterValue
        Case "Window"
            vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
            vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
            
            dPnt1(0) = vPnt1(0)
            dPnt1(1) = vPnt1(1)
            dPnt1(2) = vPnt1(2)
            
            dPnt2(0) = vPnt2(0)
            dPnt2(1) = vPnt2(1)
            dPnt2(2) = vPnt2(2)
            
            objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
        Case "Polygon"
                ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Polygon Border: "
                If Not objEntity.ObjectName = "AcDbPolyline" Then
                    MsgBox "Error: Invalid Selection."
                    Me.show
                    Exit Sub
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
                
                objSS.SelectByPolygon acSelectionSetCrossingPolygon, dCoords, filterType, filterValue
    End Select
    
    MsgBox cbFromType.Value & " found:  " & objSS.count
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadPoint Then       '<------------------------------------------------- From Point w/OD
            Set objPoint = objEntity
            dInsertPnt(0) = objPoint.Coordinates(0)
            dInsertPnt(1) = objPoint.Coordinates(1)
            dInsertPnt(2) = objPoint.Coordinates(2)
            
            If iCount = 0 Then
                For i = 0 To UBound(strAttList)
                    strAttTemp(i) = strAttList(i)
                Next i
                GoTo Place_Block
            End If
    
            Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
            Set tbls = amap.Projects(ThisDrawing).ODTables
            If tbls.count > 0 Then
                For Each tbl In tbls
                    If tbl.Name = cbFromList.Value Then GoTo Exit_For
                Next
            End If
Exit_For:

            Set ODRcs = tbl.GetODRecords
            
            boolVal = ODRcs.Init(objEntity, True, False)
            Set ODRc = ODRcs.Record
            
            For i = 0 To UBound(strAttList)
                If InStr(strAttList(i), "}") Then
                    vLine = Split(strAttList(i), "}")
                    iTest = UBound(vLine)
                    
                    If vLine(iTest) = "" Then iTest = iTest - 1
                    
                    strLine = ""
                    
                    For j = 0 To iTest
                        If InStr(vLine(j), "{") Then
                            vTemp = Split(vLine(j), "{")
                            strLine = strLine & vTemp(0)
                            strTemp = vTemp(1)
                            
                            vTemp = Split(strTemp, "~")
                            iTemp = CInt(vTemp(0))
                            
                            Select Case UBound(vTemp)
                                Case Is = 0
                                    If strLine = "" Then
                                        strLine = ODRc.Item(iTemp).Value
                                    Else
                                        strLine = strLine & ODRc.Item(iTemp).Value
                                    End If
                                    
                                Case Is = 1
                                    str1 = ODRc.Item(iTemp).Value
                                    vData = Split(str1, " ")
                                    
                                    Select Case UCase(vTemp(1))
                                        Case "L"
                                            If strLine = "" Then
                                                strLine = vData(UBound(vData))
                                            Else
                                                strLine = strLine & " " & vData(UBound(vData))
                                            End If
                                        Case "-L"
                                            For h = 0 To UBound(vData) - 1
                                                If strLine = "" Then
                                                    strLine = vData(h)
                                                Else
                                                    strLine = strLine & " " & vData(h)
                                                End If
                                            Next h
                                    End Select
                                    
                                Case Is = 2
                                    str1 = ODRc.Item(iTemp).Value
                                    str2 = vTemp(1)
                                    str3 = vTemp(2)
                                    
                                    strTemp = GetTextSegment((str1), (str2), (str3))
                                    If strLine = "" Then
                                        strLine = strTemp
                                    Else
                                        strLine = strLine & " " & strTemp
                                    End If
                            End Select
                        Else
                            If strLine = "" Then
                                strLine = vLine(j)
                            Else
                                strLine = strLine & " " & vLine(j)
                            End If
                        End If
                    Next j
                    strAttTemp(i) = strLine
                Else
                    strAttTemp(i) = strAttList(i)
                End If
            Next i
            
            
            
            
        Else                                        '<------------------------------------------------- From Block
            Set objBlock = objEntity
            dInsertPnt(0) = objBlock.InsertionPoint(0)
            dInsertPnt(1) = objBlock.InsertionPoint(1)
            dInsertPnt(2) = objBlock.InsertionPoint(2)
            
            If iCount = 0 Then
                For i = 0 To UBound(strAttList)
                    strAttTemp(i) = strAttList(i)
                Next i
                GoTo Place_Block
            End If
            
            vBlockAtt = objBlock.GetAttributes
            
            For i = 0 To UBound(strAttList)
                If InStr(strAttList(i), "}") Then
                    vLine = Split(strAttList(i), "}")
                    iTest = UBound(vLine)
                    
                    If vLine(iTest) = "" Then iTest = iTest - 1
                    
                    strLine = ""
                    
                    For j = 0 To iTest
                        If InStr(vLine(j), "{") Then 'Left(vLine(j), 1) = "{" Then
                            vTemp = Split(vLine(j), "{")
                            strLine = strLine & vTemp(0)
                            strTemp = vTemp(1)
                            'vLine(j) = Replace(vLine(j), "{", "")
                            
                            vTemp = Split(strTemp, "~")
                            iTemp = CInt(vTemp(0))
                            
                            Select Case UBound(vTemp)
                                Case Is = 0
                                    If strLine = "" Then
                                        strLine = vBlockAtt(iTemp).TextString
                                    Else
                                        strLine = strLine & " " & vBlockAtt(iTemp).TextString
                                    End If
                                    
                                Case Is = 2
                                    str1 = vBlockAtt(iTemp).TextString
                                    str2 = vTemp(1)
                                    str3 = vTemp(2)
                                    
                                    strTemp = GetTextSegment((str1), (str2), (str3))
                                    If strLine = "" Then
                                        strLine = strTemp
                                    Else
                                        strLine = strLine & " " & strTemp
                                    End If
                            End Select
                        Else
                            If strLine = "" Then
                                strLine = vLine(j)
                            Else
                                strLine = strLine & " " & vLine(j)
                            End If
                        End If
                    Next j
                    strAttTemp(i) = strLine
                Else
                    strAttTemp(i) = strAttList(i)
                End If
            Next i
        End If                                        '<------------------------------------------------- Finish Getting Data
        
Place_Block:
            
        Set objPlace = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, cbToList.Value, dScale, dScale, dScale, 0#)
        objPlace.Layer = cbToLayer.Value
        vAttList = objPlace.GetAttributes
            
        For i = 0 To UBound(strAttList)
            vAttList(i).TextString = strAttTemp(i)
        Next i
        objPlace.Update
            
        If cbDelete.Value = True Then objEntity.Delete
    Next objEntity
    
            '<------------------------------------------------------------------------------
    Select Case cbFromType.Value
        Case "Blocks"
        Case Else
    End Select
    
            '<------------------------------------------------------------------------------
    
Exit_Sub:
    
    cbConvert.Enabled = False
    
    Me.show
End Sub

Private Sub cbFromList_Change()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS2 As AcadSelectionSet
    
    On Error Resume Next
    
    If cbFromType.Value = "Block" Then
        Dim objBlock As AcadBlockReference
        Dim vAttList As Variant
    
        grpCode(0) = 2
        grpValue(0) = cbFromList.Value
        filterType = grpCode
        filterValue = grpValue
    
        Err = 0
        Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
        If Not Err = 0 Then
            Set objSS2 = ThisDrawing.SelectionSets.Item("objSS2")
            objSS2.Clear
        End If
    
        objSS2.Select acSelectionSetAll, , , filterType, filterValue
        
        For Each objBlock In objSS2
            vAttList = objBlock.GetAttributes
            lbFrom.Clear
            For i = 0 To UBound(vAttList)
                lbFrom.AddItem
                lbFrom.List(i, 0) = i
                lbFrom.List(i, 1) = vAttList(i).TagString
            Next i
            GoTo Exit_Next
        Next objBlock
Exit_Next:
    Else
        Dim amap As AcadMap
        Dim ODRcs As ODRecords
        Dim ODRc As ODRecord
        Dim tbl As ODTable
        Dim tbls As ODTables
        Dim refColumn As ODFieldDef
        Dim objEntity As AcadEntity
        Dim boolVal As Boolean
        Dim iCount As Integer
        Dim strTest As String
    
        Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
        Set tbls = amap.Projects(ThisDrawing).ODTables
        If tbls.count > 0 Then
            For Each tbl In tbls
                If tbl.Name = cbFromList.Value Then GoTo Exit_For
            Next
        End If
Exit_For:

        Set ODRcs = tbl.GetODRecords
        
        iCount = 0
        Err = 0
        
        While Err = 0
            Set refColumn = tbl.ODFieldDefs(iCount)
            strTest = refColumn.Name
            If Not Err = 0 Then GoTo Exit_objEntity
            
            lbFrom.AddItem
            lbFrom.List(iCount, 0) = iCount
            lbFrom.List(iCount, 1) = strTest
            iCount = iCount + 1
        Wend
        
Exit_objEntity:
    End If
    
    objSS2.Clear
    objSS2.Delete
End Sub

Private Sub cbFromType_Change()
    Select Case cbFromType.Value
        Case "Block"
            Dim objBlocks As AcadBlocks
            Dim strLine As String
            
            Set objBlocks = ThisDrawing.Blocks
            For i = 0 To objBlocks.count - 1
                strLine = objBlocks(i).Name
                If Not Left(strLine, 1) = "*" Then cbFromList.AddItem objBlocks(i).Name
            Next i
        Case "Point w/OD"
            Dim amap As AcadMap
            Dim tbl As ODTable
            Dim tbls As ODTables
    
            Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
            Set tbls = amap.Projects(ThisDrawing).ODTables
            If tbls.count > 0 Then
                cbFromList.Clear
                For Each tbl In tbls
                    cbFromList.AddItem tbl.Name
                Next
            End If
    End Select
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbToList_Change()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS2 As AcadSelectionSet
    
    On Error Resume Next
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    grpCode(0) = 2
    grpValue(0) = cbToList.Value
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
    If Not Err = 0 Then
        Set objSS2 = ThisDrawing.SelectionSets.Item("objSS2")
        objSS2.Clear
    End If
    
    objSS2.Select acSelectionSetAll, , , filterType, filterValue
        
    Set objBlock = objSS2.Item(0)
    
    vAttList = objBlock.GetAttributes
    lbTo.Clear
    For i = 0 To UBound(vAttList)
        lbTo.AddItem
        lbTo.List(i, 0) = i
        lbTo.List(i, 1) = vAttList(i).TagString
        lbTo.List(i, 2) = ""
    Next i
    
End Sub

Private Sub cbToList_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    cbFromType.SetFocus
End Sub

Private Sub cbTransferAll_Click()
    'If lbTo.ListIndex < 0 Then Exit Sub
    'If lbFrom.ListIndex < 0 Then Exit Sub
    If Not lbTo.ListCount = lbFrom.ListCount Then Exit Sub
    
    For i = 0 To lbTo.ListCount - 1
        lbTo.List(i, 2) = "{" & lbFrom.List(i, 0) & "}"
    Next i
End Sub

Private Sub cbTransferOne_Click()
    If lbTo.ListIndex < 0 Then Exit Sub
    If lbFrom.ListIndex < 0 Then Exit Sub
    
    lbTo.List(lbTo.ListIndex, 2) = "{" & lbFrom.List(lbFrom.ListIndex, 0) & "}"
End Sub

Private Sub cbUpdate_Click()
    lbTo.List(lbTo.ListIndex, 1) = lAttTag.Caption
    lbTo.List(lbTo.ListIndex, 2) = tbValue.Value
    cbUpdate.Enabled = False
    
    lAttTag.Caption = "Attribute Tag"
    tbValue.Value = ""
End Sub

Private Sub lbFrom_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If cbUpdate.Enabled = True Then
        tbValue = tbValue & "{" & lbFrom.List(lbFrom.ListIndex, 0) & "}"
        tbValue.SetFocus
    End If
End Sub

Private Sub lbTo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    lAttTag.Caption = lbTo.List(lbTo.ListIndex, 1)
    tbValue.Value = lbTo.List(lbTo.ListIndex, 2)
    cbUpdate.Enabled = True
    tbValue.SetFocus
End Sub

Private Sub lbTo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            lbTo.List(lbTo.ListIndex, 2) = ""
    End Select
End Sub

Private Sub UserForm_Deactivate()
    objSS.Clear
    objSS.Delete
End Sub

Private Sub UserForm_Initialize()
    cbScale.AddItem ""
    cbScale.AddItem "0.5"
    cbScale.AddItem "0.75"
    cbScale.AddItem "1.0"
    cbScale.AddItem "2.0"
    cbScale.AddItem "10"
    cbScale.AddItem "12"
    cbScale.Value = "1.0"
    
    cbSelection.AddItem "All"
    cbSelection.AddItem "Window"
    cbSelection.AddItem "Polygon"
    cbSelection.Value = "Window"
    
    cbFromType.AddItem "Block"
    cbFromType.AddItem "Point w/OD"
    cbFromType.Value = "Block"
    
    lbFrom.Clear
    lbFrom.ColumnCount = 2
    lbFrom.ColumnWidths = "20;90"
    
    lbTo.Clear
    lbTo.ColumnCount = 3
    lbTo.ColumnWidths = "20;80;70"
    
    cbToList.SetFocus
    
    On Error Resume Next
    
    Dim objBlocks As AcadBlocks
    Dim strLine As String
            
    Set objBlocks = ThisDrawing.Blocks
    For i = 0 To objBlocks.count - 1
        strLine = objBlocks(i).Name
        If Not Left(strLine, 1) = "*" Then cbToList.AddItem objBlocks(i).Name
    Next i
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        cbToLayer.AddItem objLayer.Name
    Next objLayer
    cbToLayer.Value = "0"
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
End Sub

Private Function GetTextSegment(strText As String, strSearch As String, strSegment As String)
    Dim vText As Variant
    Dim iSegment As Integer
    Dim strResult As String

    iSegment = CInt(strSegment)
    vText = Split(strText, strSearch)

    Select Case iSegment
        Case Is < 0
            iSegment = Abs(iSegment) - 1
            strResult = ""

            For i = 0 To UBound(vText)
                If Not i = iSegment Then
                    If strResult = "" Then
                        strResult = vText(i)
                    Else
                        strResult = strResult & strSearch & vText(i)
                    End If
                End If
            Next i
        Case Is > 0
            iSegment = iSegment - 1
            If Not iSegment > UBound(vText) Then
                strResult = vText(iSegment)
            End If
    End Select

    GetTextSegment = strResult
End Function
