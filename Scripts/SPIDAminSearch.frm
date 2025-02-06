VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SPIDAminSearch 
   Caption         =   "SPIDAmin Search"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   OleObjectBlob   =   "SPIDAminSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SPIDAminSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbCreateReports_Click()
    Dim vLine As Variant
    Dim strHeader, strMiddle, strEnd As String
    Dim strLine, strTemp, strSearch As String
    Dim strFileName, strFileBase, strJob As String
    Dim iCounter As Integer
    
    iCounter = 1
    
    strJob = tbJobNumber.Value & " SEARCH "
    strFileBase = ThisDrawing.Path & "\"
    
    strHeader = "{""id"":699858,""userID"":" & tbUserID.Value & ",""name"":"""
    strMiddle = """,""service"":""OH Structure"",""groups"":[{""id"":699859,""entries"":[{""field"":""FID Number"",""id"":699860,""unit"":null,""value"":"""
    strEnd = """,""operator"":""is in the list""}]}]}"
    
    strTemp = Replace(tbResult.Value, vbLf, "")
    vLine = Split(strTemp, vbCr)
    
    For i = 0 To UBound(vLine)
        If vLine(i) = "" Then GoTo Next_line
        
        If iCounter < 10 Then
            strSearch = strJob & "0" & iCounter
        Else
            strSearch = strJob & iCounter
        End If
        
        strLine = strHeader
        strLine = strLine & strSearch & strMiddle & vLine(i) & strEnd
        
        strFileName = strFileBase & strSearch & ".json"
        Open strFileName For Output As #1
        Print #1, strLine
        Close #1
        
        'MsgBox strFilename & vbCr & vbCr & strLine
        iCounter = iCounter + 1
Next_line:
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
    
    For i = 0 To lbFrom.ListCount - 1
        If lbFrom.List(i, 1) = "FID_NUMBER" Then
            lbFrom.Selected(i) = True
            Exit Sub
        End If
    Next i
End Sub

Private Sub cbGetPolygon_Click()
    Dim amap As AcadMap
    Dim tbl As ODTable
    Dim tbls As ODTables
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objPoint As AcadPoint
    Dim vReturnPnt As Variant
    Dim vLine As Variant
    Dim strLine, strTemp As String
    Dim iIndex As Integer
    Dim dCoords() As Double
    Dim iCounter, iTotal As Integer
    Dim iNPG, iGN As Integer
    
    On Error Resume Next
    
    iNPG = CInt(tbNPG.Value)
    iGN = 0
    iTotal = 0
    
    If lbFrom.ListIndex < 0 Then Exit Sub
    iIndex = CInt(lbFrom.List(lbFrom.ListIndex, 0))
    
    strLine = ""
    
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
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadPoint Then
            Set objPoint = objEntity
            
            boolVal = ODRcs.Init(objEntity, True, False)
            Set ODRc = ODRcs.Record
            
            strTemp = ODRc.Item(iIndex).Value
            If strTemp = "" Then GoTo Next_Object
            
            iGN = iGN + 1
            iTotal = iTotal + 1
            
            If strLine = "" Then
                strLine = strTemp
            Else
                If iGN > iNPG Then
                    strLine = strLine & vbCr & vbCr & strTemp
                    iGN = 1
                Else
                    strLine = strLine & "," & strTemp
                End If
            End If
        End If
Next_Object:
    Next objEntity
    
    tbResult.Value = strLine
    
    cbCreateReports.Enabled = True
    tbListcount.Value = iTotal
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub tbResult_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbResult.SelStart = 0
    tbResult.SelLength = Len(tbResult.Value)
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbResult.SelStart = 0
    tbResult.SelLength = Len(tbResult.Value)
End Sub

Private Sub UserForm_Initialize()
    lbFrom.Clear
    lbFrom.ColumnCount = 2
    lbFrom.ColumnWidths = "20;90"
    
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
    
    Dim vLine As Variant
    Dim strLine, strName, strTemp As String
    
    strName = ""
    strTemp = LCase(ThisDrawing.Path)
    If InStr(strTemp, "united") > 0 Then strName = "UTC"
    'If InStr(strTemp, "mastec") > 0 Then strName = "MAS"
    'If InStr(strTemp, "ecc ") > 0 Then strName = "ECC"
    
    vLine = Split(ThisDrawing.Name, " ")
    strName = strName & vLine(0)
    
    tbJobNumber.Value = strName
    
    If cbFromList.ListCount > 0 Then
        For i = 0 To cbFromList.ListCount - 1
            strName = LCase(cbFromList.List(i))
            If InStr(strName, "structure") > 0 Then
                If InStr(strName, "oh") > 0 Then
                    cbFromList.Value = cbFromList.List(i)
                    GoTo Past_Structure
                End If
            End If
        Next i
    End If
Past_Structure:
    
End Sub
