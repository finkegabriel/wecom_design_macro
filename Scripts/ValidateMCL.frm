VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateMCL 
   Caption         =   "Validate MCL Files"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   OleObjectBlob   =   "ValidateMCL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateMCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbAddAdresses_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vLine, vItem, vCount As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim strText, strExistingName, strLine As String
    Dim iCount As Integer
    
    grpCode(0) = 2
    grpValue(0) = "Customer,SG"
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    Me.Hide
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    dPnt1(0) = vPnt1(0)
    dPnt1(1) = vPnt1(1)
    dPnt1(2) = vPnt1(2)
    
    dPnt2(0) = vPnt2(0)
    dPnt2(1) = vPnt2(1)
    dPnt2(2) = vPnt2(2)
    
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    vTemp = Split(ThisDrawing.Name, " ")
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " SERVED ADDRESSES.csv"
    
    Open strExistingName For Output As #1
    
    Print #1, "House #,Street Name,TN83F N,TN83F E,Type"
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        Select Case objBlock.Name
            Case "SG"
                If vAttList(2).TextString = "" Then GoTo Next_objBlock
                
                strLine = "SmartGrid," & vAttList(1).TextString & "," & objBlock.InsertionPoint(1)
                strLine = strLine & "," & objBlock.InsertionPoint(0) & ",SmartGrid"
                Print #1, strLine
            Case Else
                If vAttList(4).TextString = "" Then GoTo Next_objBlock
                
                strLine = vAttList(1).TextString & "," & vAttList(2).TextString & "," & objBlock.InsertionPoint(1)
                strLine = strLine & "," & objBlock.InsertionPoint(0) & "," & vAttList(0).TextString
                Print #1, strLine
        End Select
        
Next_objBlock:
    Next objBlock
    
Clear_objSS:
    objSS.Clear
    objSS.Delete
    
    Close #1
    MsgBox "Done"
    Me.show
    
    Exit Sub
    
    
    
    
    
    
    
    If lbList.ListCount < 1 Then Exit Sub
    
    'Dim strExistingName, strFileName As String
    'Dim strTemp, strFormName, strOriginal As String
    'Dim strAttach, strBorder As String
    'Dim vTemp, vTemp2, vLine As Variant
    'Dim fName As String
    'Dim objExcel As Workbook
    'Dim objSheet As Worksheet
    'Dim objDoc As Object
    'Dim iRow As Integer
    
    vTemp = Split(ThisDrawing.Name, " ")
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " Served Addresses.csv"
    
    Open strExistingName For Output As #1
    
    Print #1, "House #,Street Name"
    
    For i = 0 To lbList.ListCount - 1
        If Not lbList.List(i, 3) = "  " Then
            Print #1, Replace(lbList.List(i, 3), "  ", ",")
            
            'vTemp = Split(lbList.List(i, 3), "  ")
            'objSheet.Cells(iRow, 1).Value = vTemp(0)
            'objSheet.Cells(iRow, 2).Value = vTemp(1)
            
            'iRow = iRow + 1
        End If
    Next i
    
    Close #1
    MsgBox "Done"
    Exit Sub
    'iRow = 2
    
    vTemp = Split(ThisDrawing.Name, " ")
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " FIBER COUNT SHEET.xlsx"
    
    vLine = Split(LCase(ThisDrawing.Path), "dropbox")
    strOriginal = vLine(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\FIBER COUNT SHEET.xlsx"
    
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
    
    strTemp = "Served Address"
    For Each objSheet In objExcel.Sheets
        If objSheet.Name = strTemp Then GoTo Found_Sheet
    Next objSheet
    
    objExcel.Sheets.Add.Name = strTemp
    Set objSheet = objExcel.Sheets(strTemp)
Found_Sheet:
    
    objSheet.Rows(1).Font.Bold = True
    
    objSheet.Cells(1, 1).Value = "House #"
    objSheet.Cells(1, 2).Value = "Street Name"
    
    For i = 0 To lbList.ListCount - 1
        If Not lbList.List(i, 3) = "  " Then
            vTemp = Split(lbList.List(i, 3), "  ")
            objSheet.Cells(iRow, 1).Value = vTemp(0)
            objSheet.Cells(iRow, 2).Value = vTemp(1)
            
            iRow = iRow + 1
        End If
    Next i
'Exit_Sub:
    
    objSheet.Columns(1).ColumnWidth = 15
    objSheet.Columns(2).ColumnWidth = 20
    
    objSheet.Columns(1).HorizontalAlignment = 3
    objSheet.Columns(2).HorizontalAlignment = 3
    
    iRow = iRow - 1
    strBorder = "A1:B" & iRow
    objSheet.Range(strBorder).Borders.Weight = 2
    
    objSheet.Range("A1").AutoFilter
    
    objExcel.Save
    objExcel.Close
    
    MsgBox "Done."
End Sub

Private Sub cbAddSplitters_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    Dim strExistingName, strFileName As String
    Dim strTemp, strFormName, strOriginal As String
    Dim strAttach, strBorder As String
    Dim vTemp, vTemp2, vLine As Variant
    Dim fName As String
    Dim objExcel As Workbook
    Dim objSheet As Worksheet
    Dim objDoc As Object
    Dim iRow As Integer
    
    iRow = 2
    
    vTemp = Split(ThisDrawing.Name, " ")
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " FIBER COUNT SHEET.xlsx"
    vLine = Split(LCase(ThisDrawing.Path), "dropbax")
    strOriginal = vLine(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\FIBER COUNT SHEET.xlsx"
    
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
    
    strTemp = TabStrip1.SelectedItem.Caption
    For Each objSheet In objExcel.Sheets
        If objSheet.Name = strTemp Then GoTo Found_Sheet
    Next objSheet
    
    objExcel.Sheets.Add.Name = strTemp
    Set objSheet = objExcel.Sheets(strTemp)
Found_Sheet:
    
    objSheet.Rows(1).Font.Bold = True
    
    objSheet.Cells(1, 1).Value = strTemp
    objSheet.Cells(1, 2).Value = "Pole/Ped #"
    objSheet.Cells(1, 3).Value = "House #"
    objSheet.Cells(1, 4).Value = "Street Name"
    objSheet.Cells(1, 5).Value = "Type"
    objSheet.Cells(1, 6).Value = "Subscriber Info"
    objSheet.Cells(1, 7).Value = "Notes"
    objSheet.Cells(1, 8).Value = "Splitter Location"
    
    For i = 0 To lbList.ListCount - 1
        If lbList.List(i, 0) = strTemp Then
            If Not lbList.List(i, 4) = "SPLITTER" Then GoTo Next_line
            
            objSheet.Cells(iRow, 1).Value = lbList.List(i, 1)
            If Left(lbList.List(i, 2), 1) = "<" Then
                objSheet.Cells(iRow, 2).Value = ""
            Else
                objSheet.Cells(iRow, 2).Value = lbList.List(i, 2)
            End If
            vTemp = Split(lbList.List(i, 3), "  ")
            objSheet.Cells(iRow, 3).Value = vTemp(0)
            objSheet.Cells(iRow, 4).Value = vTemp(1)
            objSheet.Cells(iRow, 5).Value = lbList.List(i, 4)
            objSheet.Cells(iRow, 8).Value = lbList.List(i, 5)
            
            objSheet.Cells(iRow, 1).Interior.color = RGB(0, 176, 240)
            
            iRow = iRow + 1
        End If
Next_line:
    Next i
    
    objSheet.Columns(1).ColumnWidth = 15
    objSheet.Columns(2).ColumnWidth = 20
    objSheet.Columns(3).ColumnWidth = 15
    objSheet.Columns(4).ColumnWidth = 20
    objSheet.Columns(5).ColumnWidth = 10
    objSheet.Columns(6).ColumnWidth = 20
    objSheet.Columns(7).ColumnWidth = 20
    objSheet.Columns(8).ColumnWidth = 20
    
    objSheet.Columns(1).HorizontalAlignment = 3
    objSheet.Columns(2).HorizontalAlignment = 3
    objSheet.Columns(3).HorizontalAlignment = 3
    objSheet.Columns(4).HorizontalAlignment = 3
    objSheet.Columns(5).HorizontalAlignment = 3
    objSheet.Columns(6).HorizontalAlignment = 3
    objSheet.Columns(7).HorizontalAlignment = 3
    objSheet.Columns(8).HorizontalAlignment = 3
    
    iRow = iRow - 1
    strBorder = "A1:H" & iRow
    objSheet.Range(strBorder).Borders.Weight = 2
    
    objSheet.Range("A1").AutoFilter
    
    objExcel.Save
    objExcel.Close
End Sub

Private Sub cbCheckData_Click()
    If lbTab.ListCount < 1 Then Exit Sub
    
    For i = lbTab.ListCount - 1 To 0 Step -1
            If lbTab.List(i, 8) = "Y" Then
                lbTab.RemoveItem i
            Else
                If lbTab.List(i, 3) = "  " Then lbTab.RemoveItem i
            End If
    Next i
End Sub

Private Sub cbExport_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    Dim strExistingName, strFileName As String
    Dim strTemp, strFormName, strOriginal As String
    Dim strAttach, strBorder As String
    Dim vTemp, vTemp2, vLine As Variant
    Dim fName As String
    Dim objExcel As Workbook
    Dim objSheet As Worksheet
    Dim objDoc As Object
    Dim iRow As Integer
    
    iRow = 2
    
    vTemp = Split(ThisDrawing.Name, " ")
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " FIBER COUNT SHEET.xlsx"
    
    vLine = Split(LCase(ThisDrawing.Path), "dropbox")
    strOriginal = vLine(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\FIBER COUNT SHEET.xlsx"
    
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
    
    strTemp = TabStrip1.SelectedItem.Caption
    For Each objSheet In objExcel.Sheets
        If objSheet.Name = strTemp Then GoTo Found_Sheet
    Next objSheet
    
    objExcel.Sheets.Add.Name = strTemp
    Set objSheet = objExcel.Sheets(strTemp)
Found_Sheet:
    
    objSheet.Rows(1).Font.Bold = True
    
    objSheet.Cells(1, 1).Value = strTemp
    objSheet.Cells(1, 2).Value = "Pole/Ped #"
    objSheet.Cells(1, 3).Value = "House #"
    objSheet.Cells(1, 4).Value = "Street Name"
    objSheet.Cells(1, 5).Value = "Type"
    objSheet.Cells(1, 6).Value = "Subscriber Info"
    objSheet.Cells(1, 7).Value = "Notes"
    
    For i = 0 To lbList.ListCount - 1
        If lbList.List(i, 0) = strTemp Then
            objSheet.Cells(iRow, 1).Value = lbList.List(i, 1)
            If Left(lbList.List(i, 2), 1) = "<" Then
                objSheet.Cells(iRow, 2).Value = ""
            Else
                objSheet.Cells(iRow, 2).Value = lbList.List(i, 2)
            End If
            vTemp = Split(lbList.List(i, 3), "  ")
            objSheet.Cells(iRow, 3).Value = vTemp(0)
            objSheet.Cells(iRow, 4).Value = vTemp(1)
            objSheet.Cells(iRow, 5).Value = lbList.List(i, 4)
            objSheet.Cells(iRow, 6).Value = lbList.List(i, 5)
            
            If lbList.List(i, 8) = "Y" Then
                If Not lbList.List(i, 4) = "  " Then objSheet.Cells(iRow, 1).Interior.color = RGB(146, 208, 80)
            End If
            
            iRow = iRow + 1
        End If
    Next i
    
    objSheet.Columns(1).ColumnWidth = 15
    objSheet.Columns(2).ColumnWidth = 20
    objSheet.Columns(3).ColumnWidth = 15
    objSheet.Columns(4).ColumnWidth = 20
    objSheet.Columns(5).ColumnWidth = 10
    objSheet.Columns(6).ColumnWidth = 20
    objSheet.Columns(7).ColumnWidth = 20
    
    objSheet.Columns(1).HorizontalAlignment = 3
    objSheet.Columns(2).HorizontalAlignment = 3
    objSheet.Columns(3).HorizontalAlignment = 3
    objSheet.Columns(4).HorizontalAlignment = 3
    objSheet.Columns(5).HorizontalAlignment = 3
    objSheet.Columns(6).HorizontalAlignment = 3
    objSheet.Columns(7).HorizontalAlignment = 3
    
    iRow = iRow - 1
    strBorder = "A1:G" & iRow
    objSheet.Range(strBorder).Borders.Weight = 2
    
    objSheet.Range("A1").AutoFilter
    
    objExcel.Save
    objExcel.Close
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub TabStrip1_Change()
    Dim strTab As String
    
    strTab = TabStrip1.SelectedItem.Caption
    
    If strTab = "All" Then
        Call GetAll
    Else
        Call GetCable(CStr(strTab))
    End If
End Sub

Private Sub UserForm_Initialize()
    lbList.ColumnCount = 9
    lbList.ColumnWidths = "48;48;48;48;48;48;48;48;48;48"
    
    lbTab.ColumnCount = 9
    lbTab.ColumnWidths = "72;36;120;144;24;144;144;24;30"
    
    On Error Resume Next
    
    Dim strFile, strFolder, strTemp As String
    Dim strFileName As String
    Dim vName, vLine, vItem As Variant
    Dim strLine, strTab, strCable As String
    Dim fName As String
    Dim iIndex, iCount As Integer
    
    strFolder = ThisDrawing.Path & "\*.*"
    
    strFile = Dir$(strFolder)
    
    Do While strFile <> ""
        If InStr(strFile, ".mcl") Then
            strTab = Replace(strFile, ".mcl", "")
            vLine = Split(strTab, " -")
            TabStrip1.Tabs.Add vLine(1), vLine(1)
            
            strFileName = ThisDrawing.Path & "\" & strFile
            
            Open strFileName For Input As #2
            
            Line Input #2, strLine
            strCable = vLine(1)
    
            While Not EOF(2)
                Line Input #2, strLine
                If strLine = "" Then GoTo Next_line
                
                vLine = Split(strLine, vbTab)
        
                lbList.AddItem strCable
                iIndex = lbList.ListCount - 1
        
                lbList.List(iIndex, 1) = vLine(0)
                
                If Left(vLine(1), 1) = "<" Then
                    lbList.List(iIndex, 2) = " "
                Else
                    lbList.List(iIndex, 2) = vLine(1)
                End If
                
                If Left(vLine(2), 1) = "<" Then
                    lbList.List(iIndex, 3) = "  "
                Else
                    lbList.List(iIndex, 3) = vLine(2) & "  " & vLine(3)
                End If
                
                If Left(vLine(4), 1) = "<" Then
                    lbList.List(iIndex, 4) = " "
                Else
                    lbList.List(iIndex, 4) = vLine(4)
                End If
                
                If Left(vLine(5), 1) = "<" Then
                    lbList.List(iIndex, 5) = " "
                Else
                    lbList.List(iIndex, 5) = vLine(5)
                End If
                
                lbList.List(iIndex, 6) = "n/a"
                lbList.List(iIndex, 7) = " "
                lbList.List(iIndex, 8) = " "
Next_line:
            Wend
    
            Close #2
        End If
        strFile = Dir$
    Loop
    
    Call GetCustomers
    
    Call GetAll
    
    TextBox1.Value = lbTab.ListCount
End Sub

Private Sub GetAll()
    If lbList.ListCount < 1 Then Exit Sub
    
    lbTab.Clear
    
    For i = 0 To lbList.ListCount - 1
        lbTab.AddItem lbList.List(i, 0)
        lbTab.List(i, 1) = lbList.List(i, 1)
        lbTab.List(i, 2) = lbList.List(i, 2)
        lbTab.List(i, 3) = lbList.List(i, 3)
        lbTab.List(i, 4) = lbList.List(i, 4)
        lbTab.List(i, 5) = lbList.List(i, 5)
        lbTab.List(i, 6) = lbList.List(i, 6)
        lbTab.List(i, 7) = lbList.List(i, 7)
        lbTab.List(i, 8) = lbList.List(i, 8)
    Next i
End Sub

Private Sub GetCable(strCable As String)
    If lbList.ListCount < 1 Then Exit Sub
    
    Dim iIndex As Integer
    
    lbTab.Clear
    iIndex = 0
    
    For i = 0 To lbList.ListCount - 1
        If lbList.List(i, 0) = strCable Then
            lbTab.AddItem lbList.List(i, 0)
            lbTab.List(iIndex, 1) = lbList.List(i, 1)
            lbTab.List(iIndex, 2) = lbList.List(i, 2)
            lbTab.List(iIndex, 3) = lbList.List(i, 3)
            lbTab.List(iIndex, 4) = lbList.List(i, 4)
            lbTab.List(iIndex, 5) = lbList.List(i, 5)
            lbTab.List(iIndex, 6) = lbList.List(i, 6)
            lbTab.List(iIndex, 7) = lbList.List(i, 7)
            lbTab.List(iIndex, 8) = lbList.List(i, 8)
            
            iIndex = iIndex + 1
        End If
    Next i
End Sub

Private Sub GetCustomers()
    If lbList.ListCount < 1 Then Exit Sub
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vLine, vItem, vCount As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim strText, strExistingName, strLine As String
    Dim iCount As Integer
    
    grpCode(0) = 2
    grpValue(0) = "Customer,SG"
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    dPnt1(0) = vPnt1(0)
    dPnt1(1) = vPnt1(1)
    dPnt1(2) = vPnt1(2)
    
    dPnt2(0) = vPnt2(0)
    dPnt2(1) = vPnt2(1)
    dPnt2(2) = vPnt2(2)
    
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    vTemp = Split(ThisDrawing.Name, " ")
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " Served Addresses.csv"
    
    Open strExistingName For Output As #1
    
    Print #1, "House #,Street Name,TN83F N,TN83F E,Type"
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        Select Case objBlock.Name
            Case "SG"
                If vAttList(2).TextString = "" Then GoTo Next_objBlock
                
                strLine = "SmartGrid," & vAttList(1).TextString & "," & objBlock.InsertionPoint(1)
                strLine = strLine & "," & objBlock.InsertionPoint(0) & ",SmartGrid"
                Print #1, strLine
                
                vLine = Split(vAttList(2).TextString, " - ")
                If UBound(vLine) < 1 Then GoTo Next_objBlock
                
                vCount = Split(vLine(1), ": ")
                If InStr(vCount(0), "(") > 0 Then
                    vCount(0) = Replace(vCount(0), "(", "")
                    vCount(1) = Replace(vCount(1), ")", "")
                End If
                
                For i = 0 To lbList.ListCount - 1
                    If vCount(0) = lbList.List(i, 0) Then
                        If vCount(1) = lbList.List(i, 1) Then
                            lbList.List(i, 6) = "SG  " & vAttList(1).TextString
                            lbList.List(i, 7) = "SMARTGRID"
                            GoTo Next_objBlock
                        End If
                    End If
                Next i
            Case Else
                If vAttList(4).TextString = "" Then GoTo Next_objBlock
                
                strLine = vAttList(1).TextString & "," & vAttList(2).TextString & "," & objBlock.InsertionPoint(1)
                strLine = strLine & "," & objBlock.InsertionPoint(0) & "," & vAttList(0).TextString
                Print #1, strLine
                
                vLine = Split(vAttList(4).TextString, " - ")
                If UBound(vLine) < 1 Then GoTo Next_objBlock
                
                vCount = Split(vLine(1), ": ")
                If InStr(vCount(0), "(") > 0 Then
                    vCount(0) = Replace(vCount(0), "(", "")
                    vCount(1) = Replace(vCount(1), ")", "")
                End If
                
                For i = 0 To lbList.ListCount - 1
                    If vCount(0) = lbList.List(i, 0) Then
                        If vCount(1) = lbList.List(i, 1) Then
                            lbList.List(i, 6) = vAttList(1).TextString & "  " & vAttList(2).TextString
                            lbList.List(i, 7) = vAttList(0).TextString
                            GoTo Next_objBlock
                        End If
                    End If
                Next i
        End Select
        
Next_objBlock:
    Next objBlock
    
Clear_objSS:
    objSS.Clear
    objSS.Delete
    
    Close #1
    
    Call ValidateCustomers
End Sub

Private Sub ValidateCustomers()
    If lbList.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbList.ListCount - 1
        If lbList.List(i, 3) = lbList.List(i, 6) Then
            If lbList.List(i, 4) = lbList.List(i, 7) Then lbList.List(i, 8) = "Y"
        End If
        
        If lbList.List(i, 2) = "<>" Then
            If lbList.List(i, 6) = "n/a" Then lbList.List(i, 8) = "Y"
        End If
        
        If lbList.List(i, 2) = " " Then
            If lbList.List(i, 6) = "n/a" Then lbList.List(i, 8) = "Y"
        End If
    Next i
End Sub
