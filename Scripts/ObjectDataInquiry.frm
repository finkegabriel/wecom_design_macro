VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ObjectDataInquiry 
   Caption         =   "Object Data Inquiry"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   OleObjectBlob   =   "ObjectDataInquiry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ObjectDataInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAmount_Change()
    If cbAmount.Value = "Sum" Then
        Label7.Visible = True
        tbIndex.Visible = True
    Else
        Label7.Visible = False
        tbIndex.Visible = False
    End If
End Sub

Private Sub cbDeselectAll_Click()
    If lbValues.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbValues.ListCount - 1
        lbValues.Selected(i) = False
    Next i
End Sub

Private Sub cbFromList_Change()
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
    
    On Error Resume Next
    
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
    
    lbFrom.Clear
        
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
    
    'For i = 0 To lbFrom.ListCount - 1
        'If lbFrom.List(i, 1) = "FID_NUMBER" Then
            'lbFrom.Selected(i) = True
            'Exit Sub
        'End If
    'Next i
End Sub

Private Sub cbGetPolygon_Click()
    If cbAmount.Value = "Sum" And tbIndex.Value = "" Then Exit Sub
    
    Dim amap As AcadMap
    Dim tbl As ODTable
    Dim tbls As ODTables
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objPoint As AcadPoint
    Dim objLine As AcadLine
    Dim objLWP As AcadLWPolyline
    Dim vReturnPnt As Variant
    Dim vLine As Variant
    Dim strLine, strTemp As String
    Dim iIndex, iLIndex, iAmount As Integer
    Dim dCoords() As Double
    Dim iCounter As Integer
    'Dim iNPG, iGN As Integer
    
    On Error Resume Next
    
    'iNPG = CInt(tbNPG.Value)
    'iGN = 0
    
    If lbFrom.ListIndex < 0 Then Exit Sub
    iIndex = CInt(lbFrom.List(lbFrom.ListIndex, 0))
    
    lbValues.Clear
    
    If tbIndex.Visible = True And Not tbIndex.Value = "" Then iLIndex = CInt(tbIndex.Value)
    
    'strLine = ""
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
    Set tbls = amap.Projects(ThisDrawing).ODTables
    If tbls.count > 0 Then
        For Each tbl In tbls
            If tbl.Name = cbFromList.Value Then GoTo Exit_For
        Next
    End If
    
    Exit Sub
    
Exit_For:
    
    Set ODRcs = tbl.GetODRecords
    
    Me.Hide
    
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
    
    grpCode(0) = 8
    grpValue(0) = cbFromList.Value
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    
    objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords, filterType, filterValue
    
        'MsgBox objSS.Item(0).ObjectName & vbCr & objSS.Item(0).ObjectID
        'GoTo Exit_Sub
    
    
    For Each objEntity In objSS
        'Select Case objEntity.ObjectName
            'Case "AcDbPoint"
                'Set objPoint = objEntity
            
                boolVal = ODRcs.Init(objEntity, True, False)
                Set ODRc = ODRcs.Record
                
                strTemp = ODRc.Item(iIndex).Value
                If strTemp = "" Then GoTo Next_Object
                
                If lbValues.ListCount > 0 Then
                    For i = 0 To lbValues.ListCount - 1
                        If strTemp = lbValues.List(i, 1) Then
                            If cbAmount.Value = "Sum" Then
                                iAmount = CInt(ODRc.Item(iLIndex).Value)
                                lbValues.List(i, 0) = CLng(lbValues.List(i, 0)) + iAmount
                            Else
                                lbValues.List(i, 0) = CInt(lbValues.List(i, 0)) + 1
                            End If
                            
                            GoTo Next_Object
                        End If
                    Next i
                End If
                
                If cbAmount.Value = "Sum" Then
                    iAmount = CInt(ODRc.Item(iLIndex).Value)
                    lbValues.AddItem iAmount
                    lbValues.List(lbValues.ListCount - 1, 1) = strTemp
                Else
                    lbValues.AddItem "1"
                    lbValues.List(lbValues.ListCount - 1, 1) = strTemp
                End If
            'Case "AcDbLine"
            'Case "AcDbPolyline"
            'Case "AcDbBlockReference"
        'End Select
Next_Object:
    Next objEntity
    
    'tbResult.Value = strLine
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbSearchText_Click()
    If lbValues.ListCount < 1 Then Exit Sub
    If tbFind.Value = "" Then Exit Sub
    
    Dim strText, strFind As String
    
    strFind = UCase(tbFind.Value)
    
    For i = 0 To lbValues.ListCount - 1
        strText = UCase(lbValues.List(i, 1))
        
        If InStr(strText, strFind) > 0 Then lbValues.Selected(i) = True
    Next i
End Sub

Private Sub cbSelectAll_Click()
    If lbValues.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbValues.ListCount - 1
        lbValues.Selected(i) = True
    Next i
End Sub

Private Sub cbTotalSelected_Click()
    If lbValues.ListCount < 1 Then Exit Sub
    
    Dim lTotal, lTemp As Long
    
    lTotal = 0#
    
    For i = 0 To lbValues.ListCount - 1
        If lbValues.Selected(i) = True Then
            lTemp = CLng(lbValues.List(i, 0))
            lTotal = lTotal + lTemp
        End If
    Next i
    
    tbTotal.Value = lTotal
End Sub

Private Sub lbValues_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbValues.ListCount < 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            If lbValues.ListIndex < 0 Then Exit Sub
            
            lbValues.RemoveItem lbValues.ListIndex
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbFrom.Clear
    lbFrom.ColumnCount = 2
    lbFrom.ColumnWidths = "20;90"
    
    lbValues.Clear
    lbValues.ColumnCount = 2
    lbValues.ColumnWidths = "48;154"
    
    cbAmount.AddItem "Count"
    cbAmount.AddItem "Sum"
    cbAmount.Value = "Count"
    
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
End Sub
