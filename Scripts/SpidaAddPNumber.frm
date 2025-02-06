VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SpidaAddPNumber 
   Caption         =   "Find Poles Numbers by FID"
   ClientHeight    =   9975.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120.001
   OleObjectBlob   =   "SpidaAddPNumber.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SpidaAddPNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbFindFID_Click()
    If tbFID.Value = "" Then Exit Sub
    If lbPoles.ListCount < 1 Then Exit Sub
    
    Dim strFID As String
    
    strFID = tbFID.Value
    For i = 0 To lbPoles.ListCount - 1
        If strFID = lbPoles.List(i, 2) Then
            lbPoles.Selected(i) = True
            lbPoles.ListIndex = i
            tbPoleNumber.Value = lbPoles.List(i, 3)
            
            GoTo Found_FID
        End If
    Next i
    
Found_FID:
    
End Sub

Private Sub cbGetData_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    
    Dim objEntity As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objPoint As AcadPoint
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    Dim amap As AcadMap
    Dim ODRcs As ODRecords
    Dim ODRc As ODRecord
    Dim tbl As ODTable
    Dim tbls As ODTables
    
    Dim vReturnPnt, vCoords As Variant
    Dim dCoords() As Double
    Dim iTemp, iCounter As Integer
    
    Dim strFID, strCoords, strTest As String
    
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
        Err = 0
    End If
    
    grpCode(0) = 8
    grpValue(0) = "OHStructure"
    filterType = grpCode
    filterValue = grpValue
    
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
    
    objSS.SelectByPolygon acSelectionSetCrossingPolygon, dCoords, filterType, filterValue
    
    'MsgBox objSS.count
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
    Set tbls = amap.Projects(ThisDrawing).ODTables
    If tbls.count > 0 Then
        For Each tbl In tbls
            If tbl.Name = "OHStructure" Then GoTo Exit_For
        Next
    End If
Exit_For:
    
    'Set ODRcs = tbl.GetODRecords
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadPoint Then
            Set objPoint = objEntity
            strCoords = objPoint.Coordinates(1) & "," & objPoint.Coordinates(0)
            
            Set ODRcs = tbl.GetODRecords
            boolVal = ODRcs.Init(objEntity, True, False)
            Set ODRc = ODRcs.Record
            
            strFID = ODRc.Item(10).Value
            
            'MsgBox strFID
            'GoTo Exit_Sub
            
            For i = 0 To lbPoles.ListCount - 1
                If lbPoles.List(i, 2) = strFID Then
                    lbPoles.List(i, 0) = strCoords
                    lbPoles.List(i, 1) = ODRc.Item(11).Value
                    GoTo Found_FID
                End If
            Next i
Found_FID:
        End If
Next_Object:
    Next objEntity
    
    objSS.Clear
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    objSS.SelectByPolygon acSelectionSetCrossingPolygon, dCoords, filterType, filterValue
    
    For Each objBlock In objSS
        strTest = objBlock.InsertionPoint(1) & "," & objBlock.InsertionPoint(0)
        
        For i = 0 To lbPoles.ListCount - 1
            If lbPoles.List(i, 0) = strTest Then
                vAttList = objBlock.GetAttributes
                lbPoles.List(i, 3) = vAttList(0).TextString
                
                GoTo Found_Block
            End If
        Next i
Found_Block:
    Next objBlock
    
Exit_Sub:
    
    objSS.Clear
    objSS.Delete
    Me.show
End Sub

Private Sub cbNextLine_Click()
    If lbPoles.ListIndex = lbPoles.ListCount - 1 Then Exit Sub
    
    Dim iIndex As Integer
    
    iIndex = lbPoles.ListIndex + 1
    lbPoles.ListIndex = iIndex
    lbPoles.Selected(iIndex) = True
    
    tbFID.Value = lbPoles.List(iIndex, 2)
    tbPoleNumber.Value = lbPoles.List(iIndex, 3)
    
    tbPoleNumber.SelStart = 0
    tbPoleNumber.SelLength = Len(tbPoleNumber.Value)
    tbPoleNumber.SetFocus
End Sub

Private Sub lbPoles_Click()
    If lbPoles.ListIndex < 0 Then Exit Sub
    
    tbFID.Value = lbPoles.List(lbPoles.ListIndex, 2)
    tbPoleNumber.Value = lbPoles.List(lbPoles.ListIndex, 3)
    
    tbPoleNumber.SelStart = 0
    tbPoleNumber.SelLength = Len(tbPoleNumber.Value)
    tbPoleNumber.SetFocus
End Sub

Private Sub tbPasted_Change()
    If tbPasted.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vTemp As Variant
    Dim strLine As String
    Dim iIndex As Integer
    
    lbPoles.Clear
    
    strLine = Replace(tbPasted.Value, vbLf, "")
    vLine = Split(strLine, vbCr)
    
    For i = 0 To UBound(vLine)
        If InStr(vLine(i), ",") < 1 Then GoTo Next_line
        
        vItem = Split(vLine(i), ",")
        
        'For j = 0 To UBound(vItem)
            'vItem(j) = Replace(vItem(j), ";;", ",")
        'Next j
        
        lbPoles.AddItem "0,0"
        iIndex = lbPoles.ListCount - 1
        lbPoles.List(iIndex, 1) = ""
        lbPoles.List(iIndex, 2) = vItem(1)
        lbPoles.List(iIndex, 3) = vItem(2)
        
Next_line:
    Next i
    
    tbListcount.Value = lbPoles.ListCount
End Sub

Private Sub UserForm_Initialize()
    lbPoles.ColumnCount = 4
    lbPoles.ColumnWidths = "144;60;120;114"
End Sub
