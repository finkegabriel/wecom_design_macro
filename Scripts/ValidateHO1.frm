VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateHO1 
   Caption         =   "Validate HO1s"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "ValidateHO1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateHO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSS As AcadSelectionSet


Private Sub cbExport_Click()
    If lbPoles.ListCount < 1 Then Exit Sub
    
    Dim strExistingName, strFileName As String
    Dim strTemp, strFormName, strOriginal As String
    Dim strAttach, strBorder, strLine As String
    Dim vTemp, vItem, vLine As Variant
    Dim fName As String
    Dim objExcel As Workbook
    Dim objSheet As Worksheet
    Dim objDoc As Object
    Dim iRow As Integer
    
    iRow = 3
    'C:\Users\integrity.2\Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms
    
    vTemp = Split(ThisDrawing.Name, " ")
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " SPLICING SHEET.xlsx"
    
    vLine = Split(LCase(ThisDrawing.Path), "dropbox")
    strOriginal = vLine(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\SPLICING SHEET.xlsx"
    
    fName = Dir(strFileName)
    If fName = "" Then
        Exit Sub
    End If
    
    'Set objExcel = CreateObject("Excel.Application")
    'objExcel.Visible = False
    
    fName = Dir(strExistingName)
    If fName = "" Then
        Set objExcel = Workbooks.Open(strOriginal)
        objExcel.SaveAs (strExistingName)
        MsgBox "Created New File."
    Else
        Set objExcel = Workbooks.Open(strExistingName)
        'MsgBox "Opened Existing File"
    End If
    
    Set objSheet = objExcel.Sheets("SPLICES")
    
    For i = 0 To lbPoles.ListCount - 1
        objSheet.Cells(iRow, 1).Value = lbPoles.List(i, 0)
        objSheet.Cells(iRow, 2).Value = lbPoles.List(i, 2)
        objSheet.Cells(iRow, 3).Value = lbPoles.List(i, 1)
        
        strLine = ""
        vTemp = Split(lbPoles.List(i, 4), "] ")
        vLine = Split(vTemp(1), " + ")
        For j = 0 To UBound(vLine)
            If vLine(j) = "" Then GoTo Next_J
            
            vItem = Split(vLine(j), ": ")
            If strLine = "" Then
                strLine = vItem(0) & ": " & vItem(1)
            Else
                strLine = strLine & " + " & vItem(0) & ": " & vItem(1)
            End If
Next_J:
        Next j
        objSheet.Cells(iRow, 4).Value = strLine
            
        iRow = iRow + 1
    Next i
    
    iRow = iRow - 1
    strBorder = "A3:D" & iRow
    objSheet.Range(strBorder).Borders.Weight = 2
    
    objSheet.Range("A2").AutoFilter
    
    objExcel.Save
    objExcel.Close
    
    '<---------------------------------------- Save data to csv file
    'Dim strPath, strName, strFile As String
    'Dim vTemp As Variant
    
    'strPath = ThisDrawing.Path
    'strName = ThisDrawing.Name
    'vTemp = Split(strName, " ")
    
    'strFile = strPath & "\" & vTemp(0) & " Closure Locations.csv"
    
    'Open strFile For Output As #1
    
    'Print #1, "Location,Closure Type,HO1,Counts Spliced"
    
    'For i = 0 To lbPoles.ListCount - 1
        'Print #1, lbPoles.List(i, 0) & "," & lbPoles.List(i, 2) & "," & lbPoles.List(i, 1) & "," & lbPoles.List(i, 4)
    'Next i
    
    'Close #1
End Sub

Private Sub cbGetData_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    Dim vAttList, vUnits, vItem, vLine As Variant
    Dim vTemp, vReturnPnt, vCoords As Variant
    Dim vCounts As Variant
    Dim filterType, filterValue As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim strLine, strSpliced, strClosure As String
    Dim iTemp, iCounter, iHO1 As Integer
    Dim iStart, iEnd As Integer
    Dim dCoords() As Double

    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH,sPanel,Callout,sMH"
    filterType = grpCode
    filterValue = grpValue
    
    lbPoles.Clear
    lbCallouts.Clear
    
    cbRemoveSame.Enabled = True
    
  On Error Resume Next
  
    Me.Hide
    
    If cbType.Value = "Polygon" Then
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
    
        objSS.SelectByPolygon acSelectionSetCrossingPolygon, dCoords
        'objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords
    Else
        vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
        vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
        dPnt1(0) = vPnt1(0)
        dPnt1(1) = vPnt1(1)
        dPnt1(2) = vPnt1(2)
    
        dPnt2(0) = vPnt2(0)
        dPnt2(1) = vPnt2(1)
        dPnt2(2) = vPnt2(2)
  
        objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    End If
    
    ReDim dCoords(0 To 2) As Double
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
            
        Select Case objBlock.Name
            Case "sPole"
                vAttList = objBlock.GetAttributes
                    
                If vAttList(0).TextString = "" Then GoTo Next_Object
                If vAttList(0).TextString = "POLE" Then GoTo Next_Object
                If vAttList(26).TextString = "" Then GoTo Next_Object
                
                If vAttList(26).TextString = "[A1] " Then
                    vAttList(26).TextString = ""
                    objBlock.Update
                    GoTo Next_Object
                End If
                
                strClosure = "none"
                vUnits = Split(vAttList(27).TextString, ";;")
                For j = 0 To UBound(vUnits)
                    Select Case Left(vUnits(j), 5)
                        Case "+HACO", "+HBFO", "+WHAC", "+WHBF"
                            vItem = Split(vUnits(j), "=")
                            strClosure = Replace(vItem(0), "+", "")
                            GoTo Found_Closure
                        End Select
                Next j
Found_Closure:
                iHO1 = 0
                
                strSpliced = vAttList(26).TextString
                vUnits = Split(strSpliced, " + ")
                For j = 0 To UBound(vUnits)
                    If vUnits(j) = "" Then GoTo Next_line
                    
                    vItem = Split(vUnits(j), ": ")
                    If InStr(vItem(1), "FUTURE") > 0 Then GoTo Next_line
                    vCounts = Split(vItem(1), "-")
                    If UBound(vCounts) = 0 Then
                        iHO1 = iHO1 + 1
                    Else
                        iStart = CInt(vCounts(0))
                        iEnd = CInt(vCounts(1))
                        
                        If iStart = iEnd Then
                            iHO1 = iHO1 + 1
                        Else
                            iHO1 = iHO1 + iEnd - iStart + 1
                        End If
                    End If
Next_line:
                Next j
                
                lbPoles.AddItem vAttList(0).TextString
                lbPoles.List(lbPoles.ListCount - 1, 1) = iHO1
                lbPoles.List(lbPoles.ListCount - 1, 2) = strClosure
                lbPoles.List(lbPoles.ListCount - 1, 3) = i
                lbPoles.List(lbPoles.ListCount - 1, 4) = strSpliced
                
            Case "sPed", "sHH", "sPanel"
                vAttList = objBlock.GetAttributes
                    
                If vAttList(0).TextString = "xx" Then GoTo Next_Object
                'If vAttList(3).TextString = "" Then GoTo Next_Object
                'If vAttList(3).TextString = "POLE" Then GoTo Next_Object
                If vAttList(6).TextString = "" Then GoTo Next_Object
                
                
                If vAttList(6).TextString = "[A1] " Then
                    vAttList(6).TextString = ""
                    objBlock.Update
                    GoTo Next_Object
                End If
                
                strClosure = "none"
                vUnits = Split(vAttList(7).TextString, ";;")
                For j = 0 To UBound(vUnits)
                    Select Case Left(vUnits(j), 5)
                        Case "+HACO", "+HBFO", "+WHAC", "+WHBF"
                            vItem = Split(vUnits(j), "=")
                            strClosure = Replace(vItem(0), "+", "")
                            GoTo Found_Closure2
                        End Select
                Next j
Found_Closure2:
                
                iHO1 = 0
                
                strSpliced = vAttList(6).TextString
                vUnits = Split(strSpliced, " + ")
                For j = 0 To UBound(vUnits)
                    If vUnits(j) = "" Then GoTo Next_Line2
                    
                    vItem = Split(vUnits(j), ": ")
                    If InStr(vItem(1), "FUTURE") > 0 Then GoTo Next_line
                    vCounts = Split(vItem(1), "-")
                    If UBound(vCounts) = 0 Then
                        iHO1 = iHO1 + 1
                    Else
                        iStart = CInt(vCounts(0))
                        iEnd = CInt(vCounts(1))
                        
                        If iStart = iEnd Then
                            iHO1 = iHO1 + 1
                        Else
                            iHO1 = iHO1 + iEnd - iStart + 1
                        End If
                    End If
Next_Line2:
                Next j
                
                lbPoles.AddItem vAttList(0).TextString
                lbPoles.List(lbPoles.ListCount - 1, 1) = iHO1
                lbPoles.List(lbPoles.ListCount - 1, 2) = strClosure
                lbPoles.List(lbPoles.ListCount - 1, 3) = i
                lbPoles.List(lbPoles.ListCount - 1, 4) = strSpliced
                
            Case "Callout"
                vAttList = objBlock.GetAttributes
                
                If objBlock.Layer = "Integrity Existing" Then GoTo Next_Object
                If Not Left(vAttList(1).TextString, 1) = "+" Then GoTo Next_Object
                
                vUnits = Split(vAttList(1).TextString, "=")
                iHO1 = CInt(vUnits(1))
                
                vUnits = Split(vAttList(0).TextString, ":")
                
                lbCallouts.AddItem vUnits(0)
                lbCallouts.List(lbCallouts.ListCount - 1, 1) = iHO1
                lbCallouts.List(lbCallouts.ListCount - 1, 3) = i
                
        End Select
                   
Next_Object:
    'Next objEntity
    Next i
    
Exit_Sub:
    Call GetTotals
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRemoveSame_Click()
    If lbPoles.ListCount < 1 Then Exit Sub
    If lbCallouts.ListCount < 1 Then Exit Sub
    
    For i = lbPoles.ListCount - 1 To 0 Step -1
        For j = lbCallouts.ListCount - 1 To 0 Step -1
            If lbPoles.List(i, 0) = lbCallouts.List(j, 0) Then
                If lbPoles.List(i, 1) = lbCallouts.List(j, 1) Then
                    lbPoles.RemoveItem i
                    lbCallouts.RemoveItem j
                
                    GoTo Next_I
                End If
            End If
        Next j
Next_I:
    Next i
    
    Call GetTotals
End Sub

Private Sub cbUpdate_Click()
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine, vItem, vCounts As Variant
    Dim strLine As String
    Dim iIndex, iHO1 As Integer
    Dim iStart, iEnd As Integer
    
    iHO1 = 0
    
    If tbSingle.Enabled = True Then   ' lbCallouts Edit
        'Do something
    Else                              ' lbPoles Edit
        If tbMultiple.Value = "" Then
            strLine = ""
            lbPoles.List(lbPoles.ListIndex, 1) = 0
        Else
            strLine = Replace(tbMultiple.Value, vbCr, " + ")
            strLine = Replace(strLine, vbLf, "")
        
            vLine = Split(strLine, " + ")
            For i = 0 To UBound(vLine)
                vItem = Split(vLine(i), ": ")
                If UBound(vItem) < 1 Then GoTo Next_line
                
                vCounts = Split(vItem(1), "-")
                If UBound(vCounts) = 0 Then
                    iHO1 = iHO1 + 1
                Else
                    iStart = CInt(vCounts(0))
                    iEnd = CInt(vCounts(1))
                    
                    If iStart = iEnd Then
                        iHO1 = iHO1 + 1
                    Else
                        iHO1 = iHO1 + 1 + iEnd - iStart
                    End If
                End If
Next_line:
            Next i
            lbPoles.List(lbPoles.ListIndex, 1) = iHO1
        End If
        
        iIndex = CInt(lbPoles.List(lbPoles.ListIndex, 3))
        Set objBlock = objSS.Item(iIndex)
        vAttList = objBlock.GetAttributes
        
        Select Case objBlock.Name
            Case "sPole"
                vAttList(26).TextString = strLine
            Case Else
                vAttList(6).TextString = strLine
        End Select
        objBlock.Update
    End If
    
    lbPoles.Enabled = True
    lbCallouts.Enabled = True
    cbGetData.Enabled = True
    cbRemoveSame.Enabled = True
    
    cbUpdate.Enabled = False
    
    tbSingle.Value = ""
    tbSingle.Enabled = False
    
    tbMultiple.Value = ""
    
    Call GetTotals
End Sub

Private Sub lbCallouts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim objBlock As AcadBlockReference
    Dim vCoords, vAttList As Variant
    'Dim vLine, vItem As Variant
    'Dim strLine As String
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    Dim iIndex As Integer
    
    Me.Hide
    
    iIndex = CInt(lbCallouts.List(lbCallouts.ListIndex, 2))
    Set objBlock = objSS.Item(iIndex)
    
    vCoords = objBlock.InsertionPoint
    
    viewCoordsB(0) = vCoords(0) - 200
    viewCoordsB(1) = vCoords(1) - 200
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 200
    viewCoordsE(1) = vCoords(1) + 200
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Me.show
End Sub

Private Sub lbPoles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim objBlock As AcadBlockReference
    Dim vCoords, vAttList As Variant
    Dim vLine, vItem As Variant
    Dim strLine As String
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    Dim iIndex As Integer
    
    Me.Hide
    
    lbPoles.Enabled = False
    lbCallouts.Enabled = False
    cbGetData.Enabled = False
    cbRemoveSame.Enabled = False
    
    cbUpdate.Enabled = True
    
    iIndex = CInt(lbPoles.List(lbPoles.ListIndex, 3))
    Set objBlock = objSS.Item(iIndex)
    
    'vCoords = Split(lbPoles.List(lbPoles.ListIndex, 2), ",")
    vCoords = objBlock.InsertionPoint
    
    viewCoordsB(0) = vCoords(0) - 200
    viewCoordsB(1) = vCoords(1) - 200
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 200
    viewCoordsE(1) = vCoords(1) + 200
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    vAttList = objBlock.GetAttributes
    
    Select Case objBlock.Name
        Case "sPole"
            tbMultiple.Value = Replace(vAttList(26).TextString, " + ", vbCr)
        Case Else
            tbMultiple.Value = Replace(vAttList(6).TextString, " + ", vbCr)
    End Select
    
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbPoles.ColumnCount = 5
    lbPoles.ColumnWidths = "120;48;60;28;4"
    
    lbCallouts.ColumnCount = 3
    lbCallouts.ColumnWidths = "120;48;24"
    
    cbType.AddItem "Window"
    cbType.AddItem "Polygon"
    cbType.Value = "Window"
    
  On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
End Sub

Private Sub GetTotals()
    Dim lTotal As Long
    
    lTotal = 0
    
    tbUnitCount.Value = lbPoles.ListCount
    tbSpanCount.Value = lbCallouts.ListCount
    
    If Not lbPoles.ListCount < 1 Then
        For i = 0 To lbPoles.ListCount - 1
            lTotal = lTotal + CInt(lbPoles.List(i, 1))
        Next i
        
        tbUnits.Value = lTotal
    Else
        tbUnits.Value = "0"
    End If
    
    lTotal = 0
    
    If Not lbCallouts.ListCount < 1 Then
        For i = 0 To lbCallouts.ListCount - 1
            lTotal = lTotal + CInt(lbCallouts.List(i, 1))
        Next i
        
        tbSpans.Value = lTotal
    Else
        tbSpans.Value = "0"
    End If
    
    tbDiffCount.Value = lbPoles.ListCount - lbCallouts.ListCount
    tbDiffTotal.Value = CLng(tbUnits.Value) - CLng(tbSpans.Value)
End Sub

Private Sub UserForm_Terminate()
    objSS.Clear
    objSS.Delete
End Sub
