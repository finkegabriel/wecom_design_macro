VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} yDEMeters 
   Caption         =   "Dickson Electric Meters"
   ClientHeight    =   9000.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
   OleObjectBlob   =   "yDEMeters.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "yDEMeters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbChange_Click()
    If tbText.Value = "" Then Exit Sub
    If cbType.Value = "" Then Exit Sub
    If lbAddresses.ListCount < 1 Then Exit Sub
    
    Dim strText, strType As String
    Dim iColumn As Integer
    
    strText = UCase(tbText.Value)
    Select Case Left(cbType.Value, 1)
        Case "B"
            strType = "B"
        Case "C"
            strType = "C"
        Case "M"
            strType = "M"
        Case "R"
            strType = "R"
        Case "S"
            strType = "S"
        Case "T"
            strType = "T"
        Case Else
            strType = "X"
    End Select
    iColumn = cbColumn.ListIndex
    
    For i = 0 To lbAddresses.ListCount - 1
        If InStr(UCase(lbAddresses.List(i, iColumn)), strText) > 0 Then
            lbAddresses.List(i, 3) = strType
        End If
    Next i
    
    Call GetTotals
    
    If tbHistory.Value = "" Then
        tbHistory.Value = strType & ": " & strText
    Else
        tbHistory.Value = tbHistory.Value & vbCr & strType & ": " & strText
    End If
    
    tbText.Value = ""
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
    
End Sub

Private Sub cbGetPolygon_Click()
    If lbFrom.ListCount < 1 Then Exit Sub
    
    Dim iTest As Integer
    iTest = 0
    For i = 0 To lbFrom.ListCount - 1
        If lbFrom.Selected(i) = True Then iTest = iTest + 1
    Next i
    If iTest = 0 Then Exit Sub
    
    Dim amap As AcadMap
    Dim tbl As ODTable
    Dim tbls As ODTables
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objPoint As AcadPoint
    Dim vReturnPnt As Variant
    Dim vLine As Variant
    Dim strLine, strTemp As String
    Dim iIndex, iLIndex, iAmount As Integer
    Dim dCoords() As Double
    Dim dInsert(2) As Double
    Dim iCounter, iTotla As Integer
    
    On Error Resume Next
    
    lbAddresses.Clear
    
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
    
    iTotal = 0
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadPoint Then
            Set objPoint = objEntity
            
            dInsert(0) = objPoint.Coordinates(0)
            dInsert(1) = objPoint.Coordinates(1)
            dInsert(2) = objPoint.Coordinates(2)
            
            boolVal = ODRcs.Init(objEntity, True, False)
            Set ODRc = ODRcs.Record
            
            iTotal = iTotal + 1
            iCounter = 0
            
            For i = 0 To lbFrom.ListCount - 1
                If lbFrom.Selected(i) = True Then
                    strTemp = ODRc.Item(i).Value
                    If strTemp = "" Then GoTo Next_Field
                    
                    If iCounter > 0 Then lbAddresses.List(lbAddresses.ListCount - 1, 3) = "M"
                    
                    If Left(strTemp, 1) = " " Then strTemp = Right(strTemp, Len(strTemp) - 1)
                    vLine = Split(strTemp, " ")
                    
                    strTemp = ""
                    If UBound(vLine) > 0 Then
                        strTemp = vLine(1)
                        If UBound(vLine) > 1 Then
                            For j = 2 To UBound(vLine)
                                strTemp = strTemp & " " & vLine(j)
                            Next j
                        End If
                    End If
                    
                    lbAddresses.AddItem lbFrom.List(i, 1)
                    lbAddresses.List(lbAddresses.ListCount - 1, 1) = vLine(0)
                    lbAddresses.List(lbAddresses.ListCount - 1, 2) = strTemp
                    If iCounter = 0 Then
                        lbAddresses.List(lbAddresses.ListCount - 1, 3) = "R"
                    Else
                        lbAddresses.List(lbAddresses.ListCount - 1, 3) = "M"
                        Select Case (iCounter Mod 2)
                            Case Is = 0
                                dInsert(0) = dInsert(0) - 20
                                dInsert(1) = dInsert(1) - 20
                            Case Else
                                dInsert(0) = dInsert(0) + 20
                        End Select
                    End If
                    lbAddresses.List(lbAddresses.ListCount - 1, 4) = dInsert(0) & "," & dInsert(1)
    
                    iCounter = iCounter + 1
                End If
Next_Field:
            Next i
        End If
        
Next_Object:
    Next objEntity
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Call GetTotals
    tbMeters.Value = iTotal
    
    Me.show
End Sub

Private Sub lbAddresses_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MsgBox lbAddresses.List(lbAddresses.ListIndex, 4)
End Sub

Private Sub UserForm_Initialize()
    lbAddresses.Clear
    lbAddresses.ColumnCount = 5
    lbAddresses.ColumnWidths = "96;60;156;48;90"
    
    lbFrom.Clear
    lbFrom.ColumnCount = 2
    lbFrom.ColumnWidths = "20;90"
    
    cbType.AddItem "BUSINESS"
    cbType.AddItem "CHURCH"
    cbType.AddItem "MDU"
    cbType.AddItem "RESIDENCE"
    cbType.AddItem "SCHOOL"
    cbType.AddItem "TRAILER"
    cbType.AddItem "EXTENSION"
    
    cbColumn.AddItem "Field"
    cbColumn.AddItem "Number"
    cbColumn.AddItem "Street Name"
    cbColumn.Value = "Street Name"
    
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

Private Sub GetTotals()
    If lbAddresses.ListCount < 1 Then Exit Sub
    
    Dim iB, iC, iM, iR, iSc, iT, iX As Integer
    
    iB = 0
    iC = 0
    iM = 0
    iR = 0
    iSc = 0
    iT = 0
    iX = 0
    
    For i = 0 To lbAddresses.ListCount - 1
        Select Case Left(lbAddresses.List(i, 3), 1)
             Case "B"
                iB = iB + 1
             Case "C"
                iC = iC + 1
             Case "M"
                iM = iM + 1
             Case "R"
                iR = iR + 1
             Case "S"
                iSc = iSc + 1
             Case "T"
                iT = iT + 1
             Case Else
                iX = iX + 1
        End Select
    Next i
    
    tbListCount.Value = lbAddresses.ListCount
    
    tbBUS.Value = iB
    tbChu.Value = iC
    tbMDU.Value = iM
    tbRES.Value = iR
    tbSch.Value = iSc
    tbTrlr.Value = iT
    tbExt.Value = iX
End Sub
