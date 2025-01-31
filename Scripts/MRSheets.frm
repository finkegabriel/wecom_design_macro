VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MRSheets 
   Caption         =   "MR Sheets"
   ClientHeight    =   10695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   OleObjectBlob   =   "MRSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MRSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAddReport_Click()
    If cbReports.Value = "" Then Exit Sub
    
    lbReports.AddItem cbReports.Value
End Sub

Private Sub cbATTForm_Click()
    Dim strExistingName, strFileName As String
    Dim strTemp, strFormName As String
    Dim strAttach As String
    Dim vTemp, vTemp2, vLine As Variant
    Dim fName As String
    Dim objExcel As Workbook
    Dim objDoc As Object
    Dim iRow As Integer
    
    iRow = 4
    
    vTemp = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vTemp(0) & " MR Report - ATT-Owner Lower.txt"
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " ATT - Pole-Data-003_Pole_Data_Request-Part (A).xlsx"
    
    'fName = Dir(strFileName)
    'If fName = "" Then
        'Exit Sub
    'End If
    
    Open strFileName For Input As #1
    
    'Set objExcel = CreateObject("Excel.Application")
    'objExcel.Visible = False
    
    fName = Dir(strExistingName)
    If fName = "" Then
        strTemp = ThisDrawing.Path
        vTemp2 = Split(LCase(strTemp), "dropbox")
        strFormName = vTemp2(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\5 - FORMS\ATT\ATT - Pole-Data-003_Pole_Data_Request-Part (A).xlsx"
        
        Set objExcel = Workbooks.Open(strFormName)
        objExcel.SaveAs (strExistingName)
        MsgBox "Created New File."
    Else
        Set objExcel = Workbooks.Open(strExistingName)
        MsgBox "Opened Existing File"
    End If
    
    objExcel.Sheets("Sheet1").Cells(1, 17).Value = Date
    objExcel.Sheets("Sheet1").Cells(2, 9).Value = vTemp(0)
    
    While Not EOF(1)
        Line Input #1, strTemp
        If strTemp = "" Then GoTo Next_line
        If Left(strTemp, 1) = "-" Then GoTo Next_line
        
        vTemp2 = Split(strTemp, vbTab)
        
        If Not vTemp2(0) = "ATT" Then
            While Not strTemp = ""
                Line Input #1, strTemp
            Wend
            GoTo Next_line
        End If
        
        objExcel.Sheets("Sheet1").Cells(iRow, 2).Value = vTemp2(UBound(vTemp2))
        
        Line Input #1, strTemp
        While Not Left(strTemp, 1) = "-"
            vTemp2 = Split(strTemp, vbTab)
            Select Case vTemp2(0)
                Case "UNITED#"
                    objExcel.Sheets("Sheet1").Cells(iRow, 3).Value = vTemp2(UBound(vTemp2))
                Case "SIZE-CLASS"
                    vLine = Split(vTemp2(1), "-")
                    If UBound(vLine) > 0 Then objExcel.Sheets("Sheet1").Cells(iRow, 5).Value = vLine(0)
                    If UBound(vLine) > 0 Then objExcel.Sheets("Sheet1").Cells(iRow, 6).Value = vLine(1)
                Case "LOCATION:"
                    vLine = Split(vTemp2(1), " ")
                    If UBound(vLine) > 0 Then objExcel.Sheets("Sheet1").Cells(iRow, 4).Value = vLine(1)
                Case Else
                    If UBound(vTemp2) = 0 Then objExcel.Sheets("Sheet1").Cells(iRow, 7).Value = strTemp
            End Select
            
            Line Input #1, strTemp
        Wend
        
        strAttach = ""
        
        While Not strTemp = ""
            vLine = Split(strTemp, vbTab)
            
            If vLine(UBound(vLine)) = "NEW" Or vLine(UBound(vLine)) = "FUTURE" Then
                If strAttach = "" Then
                    strAttach = vLine(UBound(vLine) - 1)
                Else
                    strAttach = strAttach & vbCrLf & vLine(UBound(vLine) - 1)
                End If
                
                'If objExcel.Sheets("Sheet1").Cells(iRow, 10).Value = "" Then
                    'objExcel.Sheets("Sheet1").Cells(iRow, 10).Value = vLine(UBound(vLine) - 1)
                'Else
                    'objExcel.Sheets("Sheet1").Cells(iRow, 10).Value = objExcel.Sheets("Sheet1").Cells(iRow, 10).Value & vbCrLf & vLine(UBound(vLine) - 1)
                'End If
                
                If vLine(UBound(vLine)) = "FUTURE" Then
                    strAttach = strAttach & " FUTURE"
                End If
            
                objExcel.Sheets("Sheet1").Cells(iRow, 9).Value = "C"
            End If
            
            Line Input #1, strTemp
        Wend
        objExcel.Sheets("Sheet1").Cells(iRow, 10).Value = strAttach
        
        iRow = iRow + 1
        
Next_line:
    Wend
    
    'MsgBox objExcel.Sheets("Sheet1").Cells(4, 3).Value
    
    objExcel.Save
    objExcel.Close
    Close #1
    
    MsgBox "Form Completed."
End Sub

Private Sub cbCompany_Change()
    If cbCompany.Value = "" Then Exit Sub
    'If cbCompany.Value = "MTEMC" Then
        'lbCompany.Clear
        'Exit Sub
    'End If
    
    Dim vLine, vItem As Variant
    Dim vAttach, vValue, vHeight As Variant
    Dim strCompany As String
    Dim iIndex, iTest As Integer
    Dim iExist, iProp As Integer
    
    lbCompany.Clear
    strCompany = cbCompany.Value
    
    For i = 0 To lbAll.ListCount - 1
        iTest = 0
        
        If strCompany = "LASH" Then
            If InStr(UCase(lbAll.List(i, 5)), "V ") > 0 Or InStr(UCase(lbAll.List(i, 5)), "V;;") > 0 Then
                lbCompany.AddItem lbAll.List(i, 0)
                iIndex = lbCompany.ListCount - 1
                lbCompany.List(iIndex, 1) = lbAll.List(i, 2)
                lbCompany.List(iIndex, 2) = ""
                lbCompany.List(iIndex, 3) = ""
                lbCompany.List(iIndex, 4) = lbAll.List(i, 6)
                lbCompany.List(iIndex, 5) = lbAll.List(i, 8)
            End If
            
            If Right(UCase(lbAll.List(i, 5)), 1) = "V" Then
                lbCompany.AddItem lbAll.List(i, 0)
                iIndex = lbCompany.ListCount - 1
                lbCompany.List(iIndex, 1) = lbAll.List(i, 2)
                lbCompany.List(iIndex, 2) = ""
                lbCompany.List(iIndex, 3) = ""
                lbCompany.List(iIndex, 4) = lbAll.List(i, 6)
                lbCompany.List(iIndex, 5) = lbAll.List(i, 8)
            End If
            
            GoTo Next_I
        End If
        
        If lbAll.List(i, 2) = strCompany Then
            lbCompany.AddItem lbAll.List(i, 0)
            iIndex = lbCompany.ListCount - 1
            lbCompany.List(iIndex, 1) = lbAll.List(i, 2)
            lbCompany.List(iIndex, 2) = ""
            lbCompany.List(iIndex, 3) = ""
            lbCompany.List(iIndex, 4) = lbAll.List(i, 6)
            lbCompany.List(iIndex, 5) = lbAll.List(i, 8)
            iTest = 1
        End If
        
        Select Case strCompany
            Case "MTE", "MTEMC", "DRE", "DREMC" 'May need more PWR CO added
            Case Else
                If InStr(lbAll.List(i, 5), strCompany) = 0 Then GoTo Next_I
        End Select
        
        If iTest = 0 Then
            lbCompany.AddItem lbAll.List(i, 0)
            iIndex = lbCompany.ListCount - 1
            lbCompany.List(iIndex, 1) = lbAll.List(i, 2)
            lbCompany.List(iIndex, 2) = ""
            lbCompany.List(iIndex, 3) = ""
            lbCompany.List(iIndex, 4) = lbAll.List(i, 6)
            lbCompany.List(iIndex, 5) = lbAll.List(i, 8)
        End If
        
        vLine = Split(UCase(lbAll.List(i, 5)), ";;")
        For j = 0 To UBound(vLine)
            vItem = Split(vLine(j), "=")
            
            Select Case strCompany
                Case vItem(0)
                    vAttach = Split(vItem(1), " ")
                    For k = 0 To UBound(vAttach)
                        vAttach(k) = Replace(vAttach(k), "C", "")
                        vAttach(k) = Replace(vAttach(k), "D", "")
                        vAttach(k) = Replace(vAttach(k), "F", "")
                        vAttach(k) = Replace(vAttach(k), "O", "")
                        vAttach(k) = Replace(vAttach(k), "T", "")
                        vAttach(k) = Replace(vAttach(k), "V", "")
                        If InStr(vAttach(k), "X") > 0 Then
                            vAttach(k) = Replace(vAttach(k), "X", "")
                            If lbCompany.List(iIndex, 3) = "" Then
                                lbCompany.List(iIndex, 3) = "1"
                            Else
                                lbCompany.List(iIndex, 3) = CInt(lbCompany.List(iIndex, 3)) + 1
                            End If
                            GoTo Next_K
                        End If
                    
                        If InStr(vAttach(k), ")") > 0 Then
                            vAttach(k) = Replace(vAttach(k), "(", "")
                            vValue = Split(vAttach(k), ")")
                        
                            vHeight = Split(vValue(0), "-")
                            iExist = CInt(vHeight(0)) * 12
                            If UBound(vHeight) > 0 Then iExist = iExist + CInt(vHeight(1))
                        
                            vHeight = Split(vValue(1), "-")
                            iProp = CInt(vHeight(0)) * 12
                            If UBound(vHeight) > 0 Then iProp = iProp + CInt(vHeight(1))
                        
                            Select Case iProp - iExist
                                Case Is < 0
                                    If lbCompany.List(iIndex, 2) = "" Then
                                        lbCompany.List(iIndex, 2) = "1"
                                    Else
                                        lbCompany.List(iIndex, 2) = CInt(lbCompany.List(iIndex, 2)) + 1
                                    End If
                                Case Else
                                    If lbCompany.List(iIndex, 3) = "" Then
                                        lbCompany.List(iIndex, 3) = "1"
                                    Else
                                        lbCompany.List(iIndex, 3) = CInt(lbCompany.List(iIndex, 3)) + 1
                                    End If
                            End Select
                        End If
Next_K:
                Next k
            Case "MTE", "MTEMC", "DRE", "DREMC"     'May need more PWR CO added
                Select Case vItem(0)
                    Case "NEUTRAL", "TRANSFORMER", "LOW POWER", "ANTENNA", "ST LT CIRCUIT", "ST LT", "PWR"
                        If InStr(vItem(1), ")") > 0 Then
                            If lbCompany.List(iIndex, 3) = "" Then
                                lbCompany.List(iIndex, 3) = "1"
                            Else
                                lbCompany.List(iIndex, 3) = CInt(lbCompany.List(iIndex, 3)) + 1
                            End If
                        End If
                End Select
            End Select
        Next j
        
Next_I:
    Next i
    
    Dim iPoles, iVisits, iOwner, iLower, iRaise As Integer
    Dim iLTotal, iRTotal, iTemp As Integer
    
    iPoles = lbCompany.ListCount
    iVisits = 0
    iOwner = 0
    iLower = 0
    iRaise = 0
    iLTotal = 0
    iRTotal = 0
    
    For i = 0 To iPoles - 1
        iTemp = 0
        
        If lbCompany.List(i, 1) = cbCompany.Value Then
            iTemp = iTemp + 1
            iOwner = iOwner + 1
        End If
        
        If Not lbCompany.List(i, 2) = "" Then
            iTemp = iTemp + 1
            iLower = iLower + 1
            iLTotal = iLTotal + CInt(lbCompany.List(i, 2))
        End If
        
        If Not lbCompany.List(i, 3) = "" Then
            iTemp = iTemp + 1
            iRaise = iRaise + 1
            iRTotal = iRTotal + CInt(lbCompany.List(i, 3))
        End If
        
        If iTemp > 0 Then iVisits = iVisits + 1
    Next i
    
    tbAttPole.Value = iPoles
    tbAttVisit.Value = iVisits
    tbAtt.Value = iOwner
    tbATTL.Value = iLower
    tbATTR.Value = iRaise
    tbATTLTotal.Value = iLTotal
    tbATTRTotal.Value = iRTotal
    
    Dim strLine As String
    
    strLine = ""
    
    lbReports.Clear
    
    If InStr(cbCompany.Value, "TDS") > 0 Then
        strLine = "TDS Reports Needed:" & vbCr & vbCr & "Owner" & vbCr & "Non Owner/MR"
        lbReports.AddItem "Owner"
        lbReports.AddItem "Non Owner/MR"
    End If
    
    If InStr(cbCompany.Value, "ATT") > 0 Then
        strLine = "ATT Reports Needed:" & vbCr & vbCr & "Owner/Lower" & vbCr & "Non Owner/R/T/A"
        lbReports.AddItem "Owner/Lower"
        lbReports.AddItem "Non Owner/R/T/A"
    End If
    
    If lbReports.ListCount < 1 Then lbReports.AddItem "All MR"
    
    'If cbCompany.Value = "NEW 6M" And lbCompany.ListCount > 0 Then
    If lbCompany.ListCount > 0 Then
        Select Case cbCompany.Value
            Case "NEW 6M", "LASH"
                strLine = lbCompany.List(0, 1) & "=1"
        
                If lbCompany.ListCount > 1 Then
                    For i = 1 To lbCompany.ListCount - 1
                        vLine = Split(strLine, ";;")
                        For j = 0 To UBound(vLine)
                            vItem = Split(vLine(j), "=")
                            If lbCompany.List(i, 1) = vItem(0) Then
                                vItem(1) = CInt(vItem(1)) + 1
                                vLine(j) = vItem(0) & "=" & vItem(1)
                                
                                strLine = vLine(0)
                                If UBound(vLine) > 0 Then
                                    For k = 1 To UBound(vLine)
                                        strLine = strLine & ";;" & vLine(k)
                                    Next k
                                End If
                                
                                GoTo Next_lbCompany
                            End If
                        Next j
                
                        strLine = strLine & ";;" & lbCompany.List(i, 1) & "=1"
Next_lbCompany:
                    Next i
                End If
        
                strLine = Replace(strLine, ";;", vbCr)
            Case Else
        End Select
    End If
    
    LHelp.Caption = strLine
    tbResult.Value = ""
End Sub

Private Sub cbCreateReports_Click()
    If lbReports.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbReports.ListCount - 1
        Call GetReports(lbReports.List(i))
    Next i
End Sub

Private Sub cbFillForm_Click()
    Dim strFileName, fName As String
    Dim vTemp As Variant
    
    vTemp = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vTemp(0)
    
    Select Case cbFormList.Value
        Case ""
            Exit Sub
        Case "ATT Form"
            strFileName = strFileName & " MR Report - ATT-Owner Lower.txt"
            
            fName = Dir(strFileName)
            If fName = "" Then
                MsgBox "MR Report - ATT-Owner Lower  form not found"
                Exit Sub
            End If
            
            Call AttForm
        Case "TDS Owner Form"
            strFileName = strFileName & " MR Report - TDS-Owner.txt"
            
            fName = Dir(strFileName)
            If fName = "" Then
                MsgBox "MR Report - TDS-Owner  form not found"
                Exit Sub
            End If
            
            Call TDSForm(CStr(strFileName))
        Case "TDS MR Form"
            strFileName = strFileName & " MR Report - TDS-NonOwner MR.txt"
            
            fName = Dir(strFileName)
            If fName = "" Then
                MsgBox "MR Report - TDS-NonOwner MR  form not found"
                Exit Sub
            End If
            
            Call TDSForm(CStr(strFileName))
    End Select
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

'Private Sub cbTDS_Click()
'    Dim objDocument
'    Dim objField As Word.FormField
'    Dim strFile, strFileName As String
'    Dim strLine As String
'    Dim vTemp, vLine As Variant
'    Dim fName As String
'    Dim iCount As Integer
'
'    iCount = 0
'
'    strFile = "C:\Integrity\Temp\Delete This\CustomerTest\TDS Pole Form with info.docx"
'    Set objDocument = CreateObject("word.application")
'    objDocument.Documents.Open strFile
'    objDocument.Visible = False
'
'    'strFileName = ThisDrawing.Path & "\" & vTemp(0) & " MR Report - ATT-Owner Lower.txt"
'
'    'fName = Dir(strFileName)
'    'If fName = "" Then
'        'Exit Sub
'    'End If
'
'    'Open strFileName For Input As #1
'
'
'
'
'
'
'    'objDocument.ActiveDocument.FormFields.Item(29).result = "TEST"
'    objDocument.ActiveDocument.FormFields(28).result = Date
'    objDocument.ActiveDocument.FormFields(29).result = "Changed"
'
'    objDocument.Documents.Close
'    'Close #1
'
'    MsgBox "Form Completed."
'End Sub

Private Sub cbWindow_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vItem, vLine, vTemp As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim strPole, strOwner, strCompany, strNew As String
    Dim strExisting, strProposed, strExtra, strNote As String
    Dim strAttachments As String
    Dim strDWG As String
    Dim iIndex As Integer
    Dim iExist, iProp, iPole As Integer
    
    On Error Resume Next
    
    Me.Hide
    
    lbAll.Clear
    lbCompany.Clear
    cbCompany.Clear
    tbOwners.Value = ""
    strDWG = "Window"
    iPole = 0
        
    Err = 0
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    End If
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = "POLE" Then GoTo Next_objBlock
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        
        iPole = iPole + 1
        strAttachments = ""
    
        strOwner = Replace(tbOwners.Value, vbLf, "")
        
        If vAttList(2).TextString = "" Then GoTo Next_Z
        
        If strOwner = "" Then
            strOwner = vAttList(2).TextString & " : 1"
        Else
            vTemp = Split(strOwner, vbCr)
            For i = 0 To UBound(vTemp)
                vItem = Split(vTemp(i), " : ")
                If vItem(0) = vAttList(2).TextString Then
                    vTemp(i) = vItem(0) & " : " & CInt(vItem(1)) + 1
                    GoTo Found_Item
                End If
            Next i
            
            strOwner = strOwner & vbCr & vAttList(2).TextString & " : 1"
            GoTo Next_Z
Found_Item:
            strOwner = vTemp(0)
            If UBound(vTemp) > 0 Then
                For i = 1 To UBound(vTemp)
                    strOwner = strOwner & vbCr & vTemp(i)
                Next i
            End If
        End If
Next_Z:
        tbOwners.Value = strOwner
        
        lbAll.AddItem vAttList(0).TextString, 0
        iIndex = lbAll.ListCount - 1
        If vAttList(5).TextString = "" Then
            lbAll.List(0, 1) = ""
        Else
            lbAll.List(0, 1) = vAttList(5).TextString
        End If
        If vAttList(2).TextString = "" Then
            lbAll.List(0, 2) = ""
        Else
            lbAll.List(0, 2) = vAttList(2).TextString
        End If
        If vAttList(3).TextString = "" Then
            lbAll.List(0, 3) = ""
        Else
            lbAll.List(0, 3) = Replace(vAttList(3).TextString, " ", " & ")
        End If
        If vAttList(4).TextString = "" Then
            lbAll.List(0, 4) = ""
        Else
            If UCase(vAttList(4).TextString) = "NA" Then
                lbAll.List(0, 4) = ""
            ElseIf InStr(vAttList(4).TextString, "=") < 1 Then
                lbAll.List(0, 4) = "???=" & vAttList(4).TextString
            Else
                lbAll.List(0, 4) = vAttList(4).TextString
            End If
        End If
        
        vPnt1 = objBlock.InsertionPoint
        lbAll.List(0, 6) = FindDWGNumber(vPnt1)
        lbAll.List(0, 7) = vAttList(7).TextString
        lbAll.List(0, 8) = vPnt1(0) & "," & vPnt1(1)
        
        For i = 9 To 23
            If Not vAttList(i).TextString = "" Then
                Select Case i
                    Case Is = 9
                        If strAttachments = "" Then
                            strAttachments = "NEUTRAL=" & vAttList(i).TextString
                        Else
                            strAttachments = strAttachments & ";;" & "NEUTRAL=" & vAttList(i).TextString
                        End If
                    Case Is = 10
                        If strAttachments = "" Then
                            strAttachments = "TRANSFORMER=" & vAttList(i).TextString
                        Else
                            strAttachments = strAttachments & ";;" & "TRANSFORMER=" & vAttList(i).TextString
                        End If
                    Case Is = 11
                        If strAttachments = "" Then
                            strAttachments = "LOW POWER=" & vAttList(i).TextString
                        Else
                            strAttachments = strAttachments & ";;" & "LOW POWER=" & vAttList(i).TextString
                        End If
                    Case Is = 12
                        If strAttachments = "" Then
                            strAttachments = "ANTENNA=" & vAttList(i).TextString
                        Else
                            strAttachments = strAttachments & ";;" & "ANTENNA=" & vAttList(i).TextString
                        End If
                    Case Is = 13
                        If strAttachments = "" Then
                            strAttachments = "ST LT CIRCUIT=" & vAttList(i).TextString
                        Else
                            strAttachments = strAttachments & ";;" & "ST LT CIRCUIT=" & vAttList(i).TextString
                        End If
                    Case Is = 14
                        If strAttachments = "" Then
                            strAttachments = "ST LT=" & vAttList(i).TextString
                        Else
                            strAttachments = strAttachments & ";;" & "ST LT=" & vAttList(i).TextString
                        End If
                    Case Is = 15
                        If strAttachments = "" Then
                            strAttachments = "NEW 6M=" & vAttList(i).TextString
                        Else
                            strAttachments = strAttachments & ";;" & "NEW 6M=" & vAttList(i).TextString
                        End If
                    Case Else
                        If strAttachments = "" Then
                            strAttachments = vAttList(i).TextString
                        Else
                            strAttachments = strAttachments & ";;" & vAttList(i).TextString
                        End If
                        
                        vTemp = Split(vAttList(i).TextString, "=")
                        
                        If cbCompany.ListCount < 1 Then
                            cbCompany.AddItem vTemp(0)
                        Else
                            For n = 0 To cbCompany.ListCount - 1
                                If cbCompany.List(n) = vTemp(0) Then GoTo Found_Company
                            Next n
            
                            cbCompany.AddItem vTemp(0)
                        End If
Found_Company:
                End Select
            End If
        Next i
        
        If strAttachments = "" Then
            lbAll.List(0, 5) = ""
        Else
            lbAll.List(0, 5) = strAttachments
        End If
        
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    'Call SortAllList
    
    Call CreateCompanyList
    tbNumberPoles.Value = iPole
    tbResult.Value = ""
    
    Me.show
End Sub

Private Sub Label38_Click()
    'If lbAll.ListCount < 1 Then Exit Sub
    
    'Dim strAtt(8) As String
    'Dim iCount As Integer
    
    'iCount = lbAll.ListCount - 1
    
    'For a = iCount To 0 Step -1
        'For b = 0 To a - 1
            'If lbAll.List(b, 6) > lbAll.List(b + 1, 6) Then
                'strAtt(0) = lbAll.List(b + 1, 0)
                'strAtt(1) = lbAll.List(b + 1, 1)
                'strAtt(2) = lbAll.List(b + 1, 2)
                'strAtt(3) = lbAll.List(b + 1, 3)
                'strAtt(4) = lbAll.List(b + 1, 4)
                'strAtt(5) = lbAll.List(b + 1, 5)
                'strAtt(6) = lbAll.List(b + 1, 6)
                'strAtt(7) = lbAll.List(b + 1, 7)
                'strAtt(8) = lbAll.List(b + 1, 8)

                'lbAll.List(b + 1, 0) = lbAll.List(b, 0)
                'lbAll.List(b + 1, 1) = lbAll.List(b, 1)
                'lbAll.List(b + 1, 2) = lbAll.List(b, 2)
                'lbAll.List(b + 1, 3) = lbAll.List(b, 3)
                'lbAll.List(b + 1, 4) = lbAll.List(b, 4)
                'lbAll.List(b + 1, 5) = lbAll.List(b, 5)
                'lbAll.List(b + 1, 6) = lbAll.List(b, 6)
                'lbAll.List(b + 1, 7) = lbAll.List(b, 7)
                'lbAll.List(b + 1, 8) = lbAll.List(b, 8)

                'lbAll.List(b, 0) = strAtt(0)
                'lbAll.List(b, 1) = strAtt(1)
                'lbAll.List(b, 2) = strAtt(2)
                'lbAll.List(b, 3) = strAtt(3)
                'lbAll.List(b, 4) = strAtt(4)
                'lbAll.List(b, 5) = strAtt(5)
                'lbAll.List(b, 6) = strAtt(6)
                'lbAll.List(b, 7) = strAtt(7)
                'lbAll.List(b, 8) = strAtt(8)
            'End If
        'Next b
    'Next a
End Sub

Private Sub lbAll_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    strLine = GetPoleData(CStr(lbAll.List(lbAll.ListIndex, 0)))
    tbResult.Value = strLine
End Sub

Private Sub lbAll_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strPole As String
    Dim i, i2 As Integer
    Dim strAtt(8) As String
    
    Dim strLine As String
    Dim vLine As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Select Case KeyCode
        Case vbKeyReturn
            Me.Hide
    
            strLine = GetPoleData(CStr(lbAll.List(lbAll.ListIndex, 0)))
            tbResult.Value = strLine
    
            vLine = Split(lbAll.List(lbAll.ListIndex, 8), ",")
    
            viewCoordsB(0) = CDbl(vLine(0)) - 300
            viewCoordsB(1) = CDbl(vLine(1)) - 300
            viewCoordsB(2) = 0#
            viewCoordsE(0) = viewCoordsB(0) + 600
            viewCoordsE(1) = viewCoordsB(1) + 600
            viewCoordsE(2) = 0#
    
            ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
            MsgBox strLine
    
            Me.show
        Case vbKeyUp
            If lbAll.ListIndex = 0 Then Exit Sub
            i = lbAll.ListIndex
            i2 = i - 1
            
            strAtt(0) = lbAll.List(i, 0)
            strAtt(1) = lbAll.List(i, 1)
            strAtt(2) = lbAll.List(i, 2)
            strAtt(3) = lbAll.List(i, 3)
            strAtt(4) = lbAll.List(i, 4)
            strAtt(5) = lbAll.List(i, 5)
            strAtt(6) = lbAll.List(i, 6)
            strAtt(7) = lbAll.List(i, 7)
            strAtt(8) = lbAll.List(i, 8)
            
            lbAll.List(i, 0) = lbAll.List(i2, 0)
            lbAll.List(i, 1) = lbAll.List(i2, 1)
            lbAll.List(i, 2) = lbAll.List(i2, 2)
            lbAll.List(i, 3) = lbAll.List(i2, 3)
            lbAll.List(i, 4) = lbAll.List(i2, 4)
            lbAll.List(i, 5) = lbAll.List(i2, 5)
            lbAll.List(i, 6) = lbAll.List(i2, 6)
            lbAll.List(i, 7) = lbAll.List(i2, 7)
            lbAll.List(i, 8) = lbAll.List(i2, 8)
            
            lbAll.List(i2, 0) = strAtt(0)
            lbAll.List(i2, 1) = strAtt(1)
            lbAll.List(i2, 2) = strAtt(2)
            lbAll.List(i2, 3) = strAtt(3)
            lbAll.List(i2, 4) = strAtt(4)
            lbAll.List(i2, 5) = strAtt(5)
            lbAll.List(i2, 6) = strAtt(6)
            lbAll.List(i2, 7) = strAtt(7)
            lbAll.List(i2, 8) = strAtt(8)
        Case vbKeyDown
            If lbAll.ListIndex = (lbAll.ListCount - 1) Then Exit Sub
            i = lbAll.ListIndex
            i2 = i + 1
            
            strAtt(0) = lbAll.List(i, 0)
            strAtt(1) = lbAll.List(i, 1)
            strAtt(2) = lbAll.List(i, 2)
            strAtt(3) = lbAll.List(i, 3)
            strAtt(4) = lbAll.List(i, 4)
            strAtt(5) = lbAll.List(i, 5)
            strAtt(6) = lbAll.List(i, 6)
            strAtt(7) = lbAll.List(i, 7)
            strAtt(8) = lbAll.List(i, 8)
            
            lbAll.List(i, 0) = lbAll.List(i2, 0)
            lbAll.List(i, 1) = lbAll.List(i2, 1)
            lbAll.List(i, 2) = lbAll.List(i2, 2)
            lbAll.List(i, 3) = lbAll.List(i2, 3)
            lbAll.List(i, 4) = lbAll.List(i2, 4)
            lbAll.List(i, 5) = lbAll.List(i2, 5)
            lbAll.List(i, 6) = lbAll.List(i2, 6)
            lbAll.List(i, 7) = lbAll.List(i2, 7)
            lbAll.List(i, 8) = lbAll.List(i2, 8)
            
            lbAll.List(i2, 0) = strAtt(0)
            lbAll.List(i2, 1) = strAtt(1)
            lbAll.List(i2, 2) = strAtt(2)
            lbAll.List(i2, 3) = strAtt(3)
            lbAll.List(i2, 4) = strAtt(4)
            lbAll.List(i2, 5) = strAtt(5)
            lbAll.List(i2, 6) = strAtt(6)
            lbAll.List(i2, 7) = strAtt(7)
            lbAll.List(i2, 8) = strAtt(8)
    End Select
End Sub

Private Sub lbCompany_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strPole As String
    
    strPole = GetPoleData(CStr(lbCompany.List(lbCompany.ListIndex, 0)))
    tbResult.Value = strPole
    'MsgBox strPole
End Sub

Private Sub lbCompany_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strLine As String
    Dim vLine As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Select Case KeyCode
        Case vbKeyReturn
            Me.Hide
    
            strLine = GetPoleData(CStr(lbCompany.List(lbCompany.ListIndex, 0)))
            tbResult.Value = strLine
    
            vLine = Split(lbCompany.List(lbCompany.ListIndex, 5), ",")
    
            viewCoordsB(0) = CDbl(vLine(0)) - 300
            viewCoordsB(1) = CDbl(vLine(1)) - 300
            viewCoordsB(2) = 0#
            viewCoordsE(0) = viewCoordsB(0) + 600
            viewCoordsE(1) = viewCoordsB(1) + 600
            viewCoordsE(2) = 0#
    
            ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
            MsgBox strLine
    
            Me.show
    End Select
End Sub

Private Sub lbReports_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbReports.ListIndex < 0 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            lbReports.RemoveItem lbReports.ListIndex
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbAll.ColumnCount = 9
    lbAll.ColumnWidths = "102;36;48;60;102;144;48;42;0"
    
    lbCompany.ColumnCount = 6
    lbCompany.ColumnWidths = "102;48;48;48;48;0"
    
    cbReports.AddItem "All"
    cbReports.AddItem "All MR"
    cbReports.AddItem "Visits"
    cbReports.AddItem "Owner"
    cbReports.AddItem "Owner/MR"
    cbReports.AddItem "Non Owner/MR"
    cbReports.AddItem "Non Owner/R/T/A"
    cbReports.AddItem "Owner/Lower"
    cbReports.AddItem "Lower"
    cbReports.AddItem "R/T/A"
    cbReports.Value = "Visits"
    
    cbFormList.AddItem "ATT Form"
    cbFormList.AddItem "TDS Owner Form"
    cbFormList.AddItem "TDS MR Form"
End Sub

Private Sub CreateCompanyList()
    If lbAll.ListCount < 1 Then Exit Sub
    
    cbCompany.AddItem "NEW 6M"
    cbCompany.AddItem "LASH"
    
    For i = 0 To lbAll.ListCount - 1
        For j = 0 To cbCompany.ListCount - 1
            If cbCompany.List(j) = lbAll.List(i, 2) Then GoTo Next_I
        Next j
            
        cbCompany.AddItem lbAll.List(i, 2)
Next_I:
    Next i
End Sub

Private Sub SortAllList()
    If lbAll.ListCount < 1 Then Exit Sub
    
    Dim vNumber, vL, vR As Variant
    Dim strRoute, strPole As String
    Dim strRoute1, strPole1 As String
    Dim strAtt(8) As String
    Dim iCount As Integer
    
    'Dim strTemp As String
    
    On Error Resume Next
    
    iCount = lbAll.ListCount - 1
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            vL = Split(UCase(lbAll.List(b, 0)), "L")
            vR = Split(vL(UBound(vL)), "R")
            strPole = vR(UBound(vR))
            strRoute = Left(lbAll.List(b, 0), (Len(lbAll.List(b, 0)) - Len(strPole)))
            strPole = Replace(strPole, "X", "")
           
            vL = Split(UCase(lbAll.List(b + 1, 0)), "L")
            vR = Split(vL(UBound(vL)), "R")
            strPole1 = vR(UBound(vR))
            strRoute1 = Left(lbAll.List(b, 0), (Len(lbAll.List(b, 0)) - Len(strPole1)))
            strPole1 = Replace(strPole1, "X", "")
            
            'strTemp = strRoute & vbTab & strPole & vbCr & strRoute1 & vbTab & strPole1
            'strTemp = strTemp & vbCr & (strRoute > strRoute1) & vbCr
            'strTemp = strTemp & (strRoute = strRoute1) & vbTab & (strPole > strPole1)
            'MsgBox strTemp
            
            If strRoute > strRoute1 Then
                    strAtt(0) = lbAll.List(b + 1, 0)
                    strAtt(1) = lbAll.List(b + 1, 1)
                    strAtt(2) = lbAll.List(b + 1, 2)
                    strAtt(3) = lbAll.List(b + 1, 3)
                    strAtt(4) = lbAll.List(b + 1, 4)
                    strAtt(5) = lbAll.List(b + 1, 5)
                    strAtt(6) = lbAll.List(b + 1, 6)
                    strAtt(7) = lbAll.List(b + 1, 7)
                    strAtt(8) = lbAll.List(b + 1, 8)
                
                    lbAll.List(b + 1, 0) = lbAll.List(b, 0)
                    lbAll.List(b + 1, 1) = lbAll.List(b, 1)
                    lbAll.List(b + 1, 2) = lbAll.List(b, 2)
                    lbAll.List(b + 1, 3) = lbAll.List(b, 3)
                    lbAll.List(b + 1, 4) = lbAll.List(b, 4)
                    lbAll.List(b + 1, 5) = lbAll.List(b, 5)
                    lbAll.List(b + 1, 6) = lbAll.List(b, 6)
                    lbAll.List(b + 1, 7) = lbAll.List(b, 7)
                    lbAll.List(b + 1, 8) = lbAll.List(b, 8)
                
                    lbAll.List(b, 0) = strAtt(0)
                    lbAll.List(b, 1) = strAtt(1)
                    lbAll.List(b, 2) = strAtt(2)
                    lbAll.List(b, 3) = strAtt(3)
                    lbAll.List(b, 4) = strAtt(4)
                    lbAll.List(b, 5) = strAtt(5)
                    lbAll.List(b, 6) = strAtt(6)
                    lbAll.List(b, 7) = strAtt(7)
                    lbAll.List(b, 8) = strAtt(8)
            ElseIf strRoute = strRoute1 Then
                If strPole > strPole1 Then
                        strAtt(0) = lbAll.List(b + 1, 0)
                        strAtt(1) = lbAll.List(b + 1, 1)
                        strAtt(2) = lbAll.List(b + 1, 2)
                        strAtt(3) = lbAll.List(b + 1, 3)
                        strAtt(4) = lbAll.List(b + 1, 4)
                        strAtt(5) = lbAll.List(b + 1, 5)
                        strAtt(6) = lbAll.List(b + 1, 6)
                        strAtt(7) = lbAll.List(b + 1, 7)
                        strAtt(8) = lbAll.List(b + 1, 8)
                
                        lbAll.List(b + 1, 0) = lbAll.List(b, 0)
                        lbAll.List(b + 1, 1) = lbAll.List(b, 1)
                        lbAll.List(b + 1, 2) = lbAll.List(b, 2)
                        lbAll.List(b + 1, 3) = lbAll.List(b, 3)
                        lbAll.List(b + 1, 4) = lbAll.List(b, 4)
                        lbAll.List(b + 1, 5) = lbAll.List(b, 5)
                        lbAll.List(b + 1, 6) = lbAll.List(b, 6)
                        lbAll.List(b + 1, 7) = lbAll.List(b, 7)
                        lbAll.List(b + 1, 8) = lbAll.List(b, 8)
                
                        lbAll.List(b, 0) = strAtt(0)
                        lbAll.List(b, 1) = strAtt(1)
                        lbAll.List(b, 2) = strAtt(2)
                        lbAll.List(b, 3) = strAtt(3)
                        lbAll.List(b, 4) = strAtt(4)
                        lbAll.List(b, 5) = strAtt(5)
                        lbAll.List(b, 6) = strAtt(6)
                        lbAll.List(b, 7) = strAtt(7)
                        lbAll.List(b, 8) = strAtt(8)
                End If
            End If
        Next b
    Next a
    
End Sub

Private Function GetAttachmentData(strLine As String)
    Dim vResult As Variant
    Dim strResult(3) As String
    Dim vLine, vItem, vTemp As Variant
    Dim iExist, iProp As Integer
    
    strLine = UCase(strLine)
    
    strResult(0) = ""
    strResult(1) = ""
    strResult(2) = ""
    strResult(3) = ""
    
    If InStr(strLine, "C") > 0 Then
        strResult(3) = " C-WIRE"
        strLine = Replace(strLine, "C", "")
    End If
    
    If InStr(strLine, "D") > 0 Then
        strResult(3) = " DROP"
        strLine = Replace(strLine, "D", "")
    End If
    
    If InStr(strLine, "F") > 0 Then
        strResult(2) = "FUTURE"
        strLine = Replace(strLine, "F", "")
    End If
    
    If InStr(strLine, "O") > 0 Then
        strResult(3) = " OHG"
        strLine = Replace(strLine, "O", "")
    End If
    
    If InStr(strLine, "S") > 0 Then
        strResult(3) = " SS"
        strLine = Replace(strLine, "S", "")
    End If
    
    If InStr(strLine, "T") > 0 Then
        strResult(2) = "MTE TAG"
        strLine = Replace(strLine, "T", "")
    End If
    
    If InStr(strLine, "V") > 0 Then
        strResult(3) = "LASH TO "
        strLine = Replace(strLine, "V", "")
    End If
    
    If InStr(strLine, "X") > 0 Then
        strResult(2) = "ATTACH"
        strLine = Replace(strLine, "X", "")
    End If
    
    If InStr(strLine, ")") > 0 Then
        vLine = Split(strLine, ")")
        strResult(0) = Replace(vLine(0), "(", "")
        strResult(1) = vLine(1)
    Else
        strResult(0) = strLine
        GoTo Skip_Next
    End If
    
    vItem = Split(strResult(0), "-")
    iExist = CInt(vItem(0)) * 12 + CInt(vItem(1))
    
    vItem = Split(strResult(1), "-")
    iProp = CInt(vItem(0)) * 12 + CInt(vItem(1))
    
    Select Case (iProp - iExist)
        Case Is > 0
            strResult(2) = "RAISE"
        Case Is = 0
            strResult(2) = "TRANSFER"
        Case Else
            strResult(2) = "LOWER"
    End Select
    
Skip_Next:
    vResult = strResult
    GetAttachmentData = vResult
End Function

Private Function FindDWGNumber(vPnt As Variant)
    Dim objSSTemp As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vAttList As Variant
    Dim dLL(2), dUR(2) As Double
    Dim dPnt(1) As Double
    Dim dScale As Double
    Dim strDWG As String
    
    On Error Resume Next
    
    strDWG = "??"
    dPnt(0) = vPnt(0)
    dPnt(1) = vPnt(1)
    
    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSSTemp = ThisDrawing.SelectionSets.Add("objSSTemp")
    If Not Err = 0 Then
        Set objSSTemp = ThisDrawing.SelectionSets.Item("objSSTemp")
        Err = 0
    End If
    
    objSSTemp.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSSTemp
        vAttList = objBlock.GetAttributes
        dScale = objBlock.XScaleFactor
        
        dLL(0) = objBlock.InsertionPoint(0)
        dLL(1) = objBlock.InsertionPoint(1)
        
        dUR(0) = dLL(0) + (1650 * dScale)
        dUR(1) = dLL(1) + (1050 * dScale)
        
        If dPnt(0) > dLL(0) And dPnt(0) < dUR(0) Then
            If dPnt(1) > dLL(1) And dPnt(1) < dUR(1) Then
                strDWG = vAttList(0).TextString
                'MsgBox "Found:" & strDWG
                GoTo Exit_Sub
            End If
        End If
    Next objBlock
    
Exit_Sub:
    objSSTemp.Clear
    objSSTemp.Delete
    
    FindDWGNumber = strDWG
End Function

Private Function GetPoleData(strPoleNumber As String)
    If lbAll.ListCount < 1 Then Exit Function
    
    Dim strData, strAttach, strLine As String
    Dim vLine, vItem, vTemp As Variant
    
    strData = ""
    strAttach = ""
    
    For i = 0 To lbAll.ListCount - 1
        If lbAll.List(i, 0) = strPoleNumber Then
            If Len(lbAll.List(i, 2)) < 8 Then
                strData = lbAll.List(i, 2) & vbTab & vbTab & lbAll.List(i, 3)
            Else
                strData = lbAll.List(i, 2) & vbTab & lbAll.List(i, 3)
            End If
                
            strData = strData & vbCr & "UNITED#" & vbTab & vbTab & lbAll.List(i, 0)
                
            If Not lbAll.List(i, 4) = "" Then
                vTemp = Split(lbAll.List(i, 4), " ")
                For j = 0 To UBound(vTemp)
                    vItem = Split(vTemp(j), "=")
                    If Len(vItem(0)) < 8 Then
                        strData = strData & vbCr & vItem(0) & vbTab & vbTab & vItem(1)
                    Else
                        strData = strData & vbCr & vItem(0) & vbTab & vItem(1)
                    End If
                Next j
            End If
            
            strData = strData & vbCr & "SIZE-CLASS" & vbTab & lbAll.List(i, 1)  'Could move ahead of SIZE-CLASS
            
            If Not lbAll.List(i, 6) = "" Then strData = strData & vbCr & "LOCATION" & vbTab & lbAll.List(i, 6)
                
            If Not lbAll.List(i, 7) = "" Then strData = strData & vbCr & lbAll.List(i, 7)
                
            strData = strData & vbCr & "----------  ATTACHMENTS  ----------"
            
            strAttach = GetAttachments(CStr(lbAll.List(i, 5)))
            
            GoTo Exit_Next
        End If
    Next i
Exit_Next:
    
    strData = strData & vbCr & strAttach
    
    GetPoleData = strData
End Function

Private Function GetAttachments(strLine As String)
    Dim vLine, vItem, vTemp, vHeight As Variant
    Dim vTotal As Variant
    Dim strAttach, strAll, strCO As String
    Dim strTemp, strExtra As String
    Dim iExist, iCurrent, iNext, iCount As Integer
    
    strAll = ""
    
    vLine = Split(strLine, ";;")
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), "=")
        vTemp = Split(UCase(vItem(1)), " ")
        For j = 0 To UBound(vTemp)
            If vTemp(j) = "" Then GoTo Next_J
            
            strCO = vItem(0)
            strExtra = ""
            
            If InStr(vItem(0), "NEW ") > 0 Then strExtra = "NEW"
            
            If InStr(vTemp(j), "C") > 0 Then
                strCO = vItem(0) & " C_WIRE"
                vTemp(j) = Replace(vTemp(j), "C", "")
            End If
        
            If InStr(vTemp(j), "D") > 0 Then
                strCO = vItem(0) & " DROP"
                vTemp(j) = Replace(vTemp(j), "D", "")
            End If
        
            If InStr(vTemp(j), "F") > 0 Then
                strExtra = "FUTURE"
                vTemp(j) = Replace(vTemp(j), "F", "")
            End If
        
            If InStr(vTemp(j), "O") > 0 Then
                strCO = vItem(0) & " OHG"
                vTemp(j) = Replace(vTemp(j), "O", "")
            End If
        
            If InStr(vTemp(j), "S") > 0 Then
                strCO = vItem(0) & " SS"
                vTemp(j) = Replace(vTemp(j), "S", "")
            End If
        
            If InStr(vTemp(j), "T") > 0 Then
                strExtra = "MTE TAG"
                vTemp(j) = Replace(vTemp(j), "T", "")
            End If
        
            If InStr(vTemp(j), "V") > 0 Then
                strCO = "LASH TO " & vItem(0)
                vTemp(j) = Replace(vTemp(j), "V", "")
            End If
        
            If InStr(vTemp(j), "X") > 0 Then
                strExtra = "ATTACH"
                vTemp(j) = Replace(vTemp(j), "X", "")
            End If
            
            vHeight = Split(vTemp(j), ")")
            vTotal = Split(vHeight(UBound(vHeight)), "-")
            iExist = CInt(vTotal(0)) * 12
            If UBound(vTotal) > 0 Then iExist = iExist + CInt(vTotal(1))
            
            If strAll = "" Then
                strAll = iExist & "=" & strCO & "=" & Replace(vTemp(j), ")", ")=")
            Else
                strAll = strAll & ";;" & iExist & "=" & strCO & "=" & Replace(vTemp(j), ")", ")=")
            End If
            
            If Not strExtra = "" Then strAll = strAll & "=" & strExtra
Next_J:
            
        Next j
    Next i
    
    vLine = Split(strAll, ";;")
    iCount = UBound(vLine)
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            vItem = Split(vLine(b), "=")
            iCurrent = CInt(vItem(0))
            vItem = Split(vLine(b + 1), "=")
            iNext = CInt(vItem(0))
            
            If iCurrent < iNext Then
                strTemp = vLine(b + 1)
                vLine(b + 1) = vLine(b)
                vLine(b) = strTemp
            End If
        Next b
    Next a
    
    strAttach = "NONE"
    
    For i = 0 To iCount
        vItem = Split(vLine(i), "=")
        strTemp = vItem(1)
        If Len(strTemp) < 8 Then
            strTemp = strTemp & vbTab & vbTab & vItem(2)
        Else
            strTemp = strTemp & vbTab & vItem(2)
        End If
        If UBound(vItem) > 2 Then strTemp = strTemp & vbTab & vItem(3)
        
        If strAttach = "NONE" Then
            strAttach = strTemp
        Else
            strAttach = strAttach & vbCr & strTemp
        End If
    Next i
    
    GetAttachments = strAttach
End Function

Private Sub GetReports(strReport As String)
    If lbCompany.ListCount < 1 Then Exit Sub
    
    Dim vTemp As Variant
    Dim strDWGName, strFileName As String
    Dim strPole As String
    Dim iTest, iCount As Integer
    
    iCount = 0
    
    vTemp = Split(ThisDrawing.Name, " ")
    strDWGName = ThisDrawing.Path & "\" & vTemp(0)
    
    Select Case strReport
        Case "All"
            If tbAttPole.Value = "" Or tbAttPole.Value = "0" Then
                MsgBox "No poles to add to a report."
                Exit Sub
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-All Poles.txt"
        Case "All MR"
            If tbATTR.Value = "" Or tbATTR.Value = "0" Then
                If tbATTL.Value = "" Or tbATTL.Value = "0" Then
                    MsgBox "No poles to add to a report."
                    Exit Sub
                End If
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-All MR.txt"
        Case "Visits"
            If tbAttVisit.Value = "" Or tbAttVisit.Value = "0" Then
                MsgBox "No poles to add to a report."
                Exit Sub
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-Visits.txt"
        Case "Owner"
            If tbAtt.Value = "" Or tbAtt.Value = "0" Then
                MsgBox "No poles to add to a report."
                Exit Sub
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-Owner.txt"
        Case "Owner/MR"
            If tbATTR.Value = "" Or tbATTR.Value = "0" Then
                If tbATTL.Value = "" Or tbATTL.Value = "0" Then
                    MsgBox "No poles to add to a report."
                    Exit Sub
                End If
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-Owner MR.txt"
        Case "Non Owner/MR"
            If tbATTR.Value = "" Or tbATTR.Value = "0" Then
                If tbATTL.Value = "" Or tbATTL.Value = "0" Then
                    MsgBox "No poles to add to a report."
                    Exit Sub
                End If
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-NonOwner MR.txt"
        Case "Non Owner/R/T/A"
            If tbATTR.Value = "" Or tbATTR.Value = "0" Then
                If Not tbATTL.Value = "" Or Not tbATTL.Value = "0" Then
                    MsgBox "No poles to add to a report."
                    Exit Sub
                End If
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-NonOwner RTA.txt"
        Case "Lower"
            If tbATTL.Value = "" Or tbATTL.Value = "0" Then
                MsgBox "No poles to add to a report."
                Exit Sub
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-Lower.txt"
        Case "R/T/A"
            If tbATTR.Value = "" Or tbATTR.Value = "0" Then
                MsgBox "No poles to add to a report."
                Exit Sub
            End If
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-RTA.txt"
        Case Else
            iTest = CInt(tbAtt.Value) + CInt(tbATTL.Value)
            If iTest = 0 Then Exit Sub
            
            strDWGName = strDWGName & " MR Report - " & cbCompany.Value & "-Owner Lower.txt"
    End Select
    
    Open strDWGName For Output As #1
    
    For i = 0 To lbCompany.ListCount - 1
        Select Case strReport
            Case "All"
                strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                Print #1, strPole
                Print #1, ""
                Print #1, ""
                
                iCount = iCount + 1
            Case "All MR"
                If Not lbCompany.List(i, 2) = "" Or Not lbCompany.List(i, 3) = "" Then
                    strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                    Print #1, strPole
                    Print #1, ""
                    Print #1, ""
                
                    iCount = iCount + 1
                End If
            Case "Visits"
                If lbCompany.List(i, 1) = cbCompany.Value Or Not lbCompany.List(i, 2) = "" Or Not lbCompany.List(i, 3) = "" Then
                    strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                    Print #1, strPole
                    Print #1, ""
                    Print #1, ""
                
                    iCount = iCount + 1
                End If
            Case "Owner"
                If lbCompany.List(i, 1) = cbCompany.Value Then
                    strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                    Print #1, strPole
                    Print #1, ""
                    Print #1, ""
                
                    iCount = iCount + 1
                End If
            Case "Owner/MR"
                If Not lbCompany.List(i, 2) = "" Or Not lbCompany.List(i, 3) = "" Then
                    If lbCompany.List(i, 1) = cbCompany.Value Then
                        strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                        Print #1, strPole
                        Print #1, ""
                        Print #1, ""
                
                        iCount = iCount + 1
                    End If
                End If
            Case "Non Owner/MR"
                If Not lbCompany.List(i, 2) = "" Or Not lbCompany.List(i, 3) = "" Then
                    If Not lbCompany.List(i, 1) = cbCompany.Value Then
                        strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                        Print #1, strPole
                        Print #1, ""
                        Print #1, ""
                
                        iCount = iCount + 1
                    End If
                End If
            Case "Non Owner/R/T/A"
                If lbCompany.List(i, 2) = "" And Not lbCompany.List(i, 3) = "" Then
                    If Not lbCompany.List(i, 1) = cbCompany.Value Then
                        strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                        Print #1, strPole
                        Print #1, ""
                        Print #1, ""
                
                        iCount = iCount + 1
                    End If
                End If
            Case "Owner/Lower"
                If lbCompany.List(i, 1) = cbCompany.Value Or Not lbCompany.List(i, 2) = "" Then
                    strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                    Print #1, strPole
                    Print #1, ""
                    Print #1, ""
                
                    iCount = iCount + 1
                End If
            Case "Lower"
                If Not lbCompany.List(i, 2) = "" Then
                    strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                    Print #1, strPole
                    Print #1, ""
                    Print #1, ""
                
                    iCount = iCount + 1
                End If
            Case Else
                If Not lbCompany.List(i, 3) = "" Then
                    strPole = GetPoleData(CStr(lbCompany.List(i, 0)))
                
                    Print #1, strPole
                    Print #1, ""
                    Print #1, ""
                
                    iCount = iCount + 1
                End If
        End Select
Next_I:
    Next i
    
    Close #1
    
    MsgBox iCount & " poles in report:" & vbCr & strDWGName
End Sub

Private Sub AttForm()
    Dim strExistingName, strFileName As String
    Dim strTemp, strFormName As String
    Dim strAttach As String
    Dim vTemp, vTemp2, vLine As Variant
    Dim fName, strLocal As String
    Dim objExcel As Workbook
    Dim objDoc As Object
    Dim iRow As Integer
    
    iRow = 4
    
    vTemp = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vTemp(0) & " MR Report - ATT-Owner Lower.txt"
    strExistingName = ThisDrawing.Path & "\" & vTemp(0) & " ATT - Pole-Data-003_Pole_Data_Request-Part (A).xlsx"
    strLocal = "C:\Integrity\VBA\Forms\"
    
    'fName = Dir(strFileName)
    'If fName = "" Then
        'Exit Sub
    'End If
    
    Open strFileName For Input As #1
    
    'Set objExcel = CreateObject("Excel.Application")
    'objExcel.Visible = False
    
    fName = Dir(strExistingName)
    If fName = "" Then
        strTemp = ThisDrawing.Path
        If InStr(LCase(strTemp), "dropbox") > 0 Then
            vTemp2 = Split(LCase(strTemp), "dropbox")
            strFormName = vTemp2(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\5 - FORMS\ATT\ATT - Pole-Data-003_Pole_Data_Request-Part (A).xlsx"
        Else
            strFormName = strLocal & "ATT - Pole-Data-003_Pole_Data_Request-Part (A).xlsx"
        End If
        
        Set objExcel = Workbooks.Open(strFormName)
        objExcel.SaveAs (strExistingName)
        MsgBox "Created New File."
    Else
        Set objExcel = Workbooks.Open(strExistingName)
        MsgBox "Opened Existing File"
    End If
    
    objExcel.Sheets("Sheet1").Cells(1, 17).Value = Date
    objExcel.Sheets("Sheet1").Cells(2, 9).Value = vTemp(0)
    
    While Not EOF(1)
        Line Input #1, strTemp
        If strTemp = "" Then GoTo Next_line
        If Left(strTemp, 1) = "-" Then GoTo Next_line
        
        vTemp2 = Split(strTemp, vbTab)
        
        If Not vTemp2(0) = "ATT" Then
            While Not strTemp = ""
                Line Input #1, strTemp
            Wend
            GoTo Next_line
        End If
        
        objExcel.Sheets("Sheet1").Cells(iRow, 2).Value = vTemp2(UBound(vTemp2))
        
        Line Input #1, strTemp
        While Not Left(strTemp, 1) = "-"
            vTemp2 = Split(strTemp, vbTab)
            Select Case vTemp2(0)
                Case "UNITED#"
                    objExcel.Sheets("Sheet1").Cells(iRow, 3).Value = vTemp2(UBound(vTemp2))
                Case "SIZE-CLASS"
                    vLine = Split(vTemp2(1), "-")
                    If UBound(vLine) > 0 Then objExcel.Sheets("Sheet1").Cells(iRow, 5).Value = vLine(0)
                    If UBound(vLine) > 0 Then objExcel.Sheets("Sheet1").Cells(iRow, 6).Value = vLine(1)
                Case "LOCATION"
                    vLine = Split(vTemp2(1), " ")
                    If UBound(vLine) > 0 Then objExcel.Sheets("Sheet1").Cells(iRow, 4).Value = vLine(1)
                Case Else
                    If UBound(vTemp2) = 0 Then objExcel.Sheets("Sheet1").Cells(iRow, 7).Value = strTemp
            End Select
            
            Line Input #1, strTemp
        Wend
        
        strAttach = ""
        
        While Not strTemp = ""
            vLine = Split(strTemp, vbTab)
            
            If vLine(UBound(vLine)) = "NEW" Or vLine(UBound(vLine)) = "FUTURE" Then
                If strAttach = "" Then
                    strAttach = vLine(UBound(vLine) - 1)
                Else
                    strAttach = strAttach & vbCrLf & vLine(UBound(vLine) - 1)
                End If
                
                'If objExcel.Sheets("Sheet1").Cells(iRow, 10).Value = "" Then
                    'objExcel.Sheets("Sheet1").Cells(iRow, 10).Value = vLine(UBound(vLine) - 1)
                'Else
                    'objExcel.Sheets("Sheet1").Cells(iRow, 10).Value = objExcel.Sheets("Sheet1").Cells(iRow, 10).Value & vbCrLf & vLine(UBound(vLine) - 1)
                'End If
                
                If vLine(UBound(vLine)) = "FUTURE" Then
                    strAttach = strAttach & " FUTURE"
                End If
            
                objExcel.Sheets("Sheet1").Cells(iRow, 9).Value = "C"
            End If
            
            Line Input #1, strTemp
        Wend
        objExcel.Sheets("Sheet1").Cells(iRow, 10).Value = strAttach
        
        iRow = iRow + 1
        
Next_line:
    Wend
    
    'MsgBox objExcel.Sheets("Sheet1").Cells(4, 3).Value
    
    objExcel.Save
    objExcel.Close
    Close #1
    
    MsgBox "Form Completed."
End Sub

Private Sub TDSForm(strFileName As String)
    'Dim objDocument As Document
    Dim objDocument
    Dim objField As Word.FormField
    Dim strFormFile, strLine As String
    Dim strFormName As String
    Dim strTDS, strPWR, strLocation As String
    Dim strLL, strOwner, strMR, strType As String
    Dim vTemp, vLine As Variant
    Dim fName, strLocal As String
    Dim iCount, iField, iCost, iItem As Integer
    
    iCount = 0
    
    Set objDocument = CreateObject("word.application")
    strLocal = "C:\Integrity\VBA\Forms\"
    
    vTemp = Split(ThisDrawing.Name, " ")
    'strFile = "C:\Integrity\Temp\Delete This\CustomerTest\TDS Pole Form with info.docx"
    'strFormFile = "C:\Integrity\Temp\Delete This\CustomerTest\TDS Pole Form.docx"
    
    Select Case cbFormList.Value
        Case "TDS Owner Form"
            strFormFile = ThisDrawing.Path & "\" & vTemp(0) & " TDS Application for Pole Attachments (Form600P).docx"
            
            fName = Dir(strFormFile)
            If fName = "" Then
                strTemp = ThisDrawing.Path
                If InStr(LCase(strTemp), "dropbox") > 0 Then
                    vLine = Split(LCase(strTemp), "dropbox")
                    strFormName = vLine(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\TDS Application for Pole Attachments (Form600P).docx"
                Else
                    strFormName = strLocal & "TDS Application for Pole Attachments (Form600P).docx"
                End If
                
                'MsgBox strFormName
                objDocument.Documents.Open strFormName
                objDocument.Visible = True
                objDocument.ActiveDocument.SaveAs strFormFile
            Else
                objDocument.Documents.Open strFormFile
                objDocument.Visible = True
            End If
            
            objDocument.ActiveDocument.FormFields(3).result = Date
            objDocument.ActiveDocument.FormFields(4).result = vTemp(0)
            objDocument.ActiveDocument.FormFields(18).result = "125.00"
            objDocument.ActiveDocument.FormFields(19).result = "250.00"
            objDocument.ActiveDocument.FormFields(25).result = "Tim Thompson / Director of Construction"
            objDocument.ActiveDocument.FormFields(26).result = "UNITED COMMUNICATIONS"
            objDocument.ActiveDocument.FormFields(27).result = vTemp(0)
            objDocument.ActiveDocument.FormFields(28).result = Date
            
            iField = 29
            iCost = 375
        Case "TDS MR Form"
            strFormFile = ThisDrawing.Path & "\" & vTemp(0) & " TDS Third Party MRW Request Form.docx"
            
            fName = Dir(strFormFile)
            If fName = "" Then
                strTemp = ThisDrawing.Path
                If InStr(LCase(strTemp), "dropbox") > 0 Then
                    vLine = Split(LCase(strTemp), "dropbox")
                    strFormName = vLine(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\TDS Third Party MRW Request Form.docx"
                Else
                    strFormName = strLocal & "TDS Third Party MRW Request Form.docx"
                End If
                
                'MsgBox strFormFile
                objDocument.Documents.Open strFormName
                objDocument.Visible = True
                objDocument.ActiveDocument.SaveAs strFormFile
            Else
                objDocument.Documents.Open strFormFile
                objDocument.Visible = True
            End If
            
            objDocument.ActiveDocument.FormFields(3).result = Date
            objDocument.ActiveDocument.FormFields(4).result = vTemp(0)
            'objDocument.ActiveDocument.FormFields(10).result = "125.00"
            objDocument.ActiveDocument.FormFields(10).result = "250.00"
            'objDocument.ActiveDocument.FormFields(26).result = "UNITED COMMUNICATIONS"
            'objDocument.ActiveDocument.FormFields(27).result = vTemp(0)
            'objDocument.ActiveDocument.FormFields(28).result = Date
            
            iField = 14
            iCost = 250
    End Select
    
    'Set objDocument = Documents.Open(strFile)
    
    Open strFileName For Input As #1
    
    While Not EOF(1)
        Line Input #1, strTemp
        If strTemp = "" Then GoTo Next_line
        If Left(strTemp, 1) = "-" Then GoTo Next_line
        
        'strTDS = "UNK"
        vTemp = Split(strTemp, vbTab)
        
        If vTemp(0) = "TDS" Then
            strOwner = "TDS"
            strTDS = vTemp(UBound(vTemp))
            strPWR = "N/A"
        Else
            strOwner = vTemp(0)
            strTDS = "UNK"
            strPWR = vTemp(UBound(vTemp))
        End If
        
        strType = "OL"
        strLocation = ""
        strLL = ""
        strMR = "NO"
        
        Line Input #1, strTemp
        While Not Left(strTemp, 1) = "-"
            vTemp = Split(strTemp, vbTab)
            Select Case vTemp(0)
                Case "TDS"
                    strTDS = vTemp(UBound(vTemp))
                Case "MTE", "MTEMC", "DREMC", "DRE", "NES"
                    strPWR = vTemp(0) & ": " & vTemp(UBound(vTemp))
                Case "LOCATION"
                    strLocation = vTemp(UBound(vTemp))
                Case Else
                    If UBound(vTemp) = 0 Then
                        If InStr(strTemp, ",") > 0 Then
                            strLL = strTemp
                        End If
                    End If
            End Select
            
            Line Input #1, strTemp
        Wend
        
        While Not strTemp = ""
            If Not Left(strTemp, 1) = "-" Then
                vTemp = Split(strTemp, vbTab)
                If vTemp(UBound(vTemp)) = "NEW" Then strType = "SW"
                
                If InStr(vTemp(0), "TDS") > 0 Then
                    If InStr(vTemp(UBound(vTemp) - 1), "(") > 0 Then strMR = "YES-TDS"
                End If
            End If
            Line Input #1, strTemp
        Wend
        
        objDocument.ActiveDocument.FormFields(iField).result = strType
        iField = iField + 1
        objDocument.ActiveDocument.FormFields(iField).result = strTDS
        iField = iField + 1
        objDocument.ActiveDocument.FormFields(iField).result = strPWR
        iField = iField + 1
        objDocument.ActiveDocument.FormFields(iField).result = strLocation
        iField = iField + 1
        objDocument.ActiveDocument.FormFields(iField).result = strLL
        iField = iField + 1
        objDocument.ActiveDocument.FormFields(iField).result = strOwner
        iField = iField + 1
        objDocument.ActiveDocument.FormFields(iField).result = "UNK"
        iField = iField + 1
        objDocument.ActiveDocument.FormFields(iField).result = strMR
        iField = iField + 1
        If strType = "SW" Then
            objDocument.ActiveDocument.FormFields(iField).result = "FIBER WITH STRAND"
        Else
            objDocument.ActiveDocument.FormFields(iField).result = "FIBER OVERLASHED TO EXISTING"
        End If
        iField = iField + 1
        iCount = iCount + 1
Next_line:
    Wend
    
    Select Case cbFormList.Value
        Case "TDS Owner Form"
            If iCount > 10 Then
                iItem = (iCount - 10) * 25
                iCost = iCost + iItem
                objDocument.ActiveDocument.FormFields(20).result = iItem & ".00"
        
                If iCount > 50 Then
                    iItem = 200
                    iCost = iCost + iItem
                    objDocument.ActiveDocument.FormFields(21).result = iItem & ".00"
                End If
            End If
    
            objDocument.ActiveDocument.FormFields(22).result = iCost & ".00"
            objDocument.ActiveDocument.FormFields(10).result = iCount
        Case Else
            If iCount > 10 Then
                iItem = (iCount - 10) * 25
                iCost = iCost + iItem
                objDocument.ActiveDocument.FormFields(11).result = iItem & ".00"
        
                If iCount > 50 Then
                    iItem = 200
                    iCost = iCost + iItem
                    objDocument.ActiveDocument.FormFields(12).result = iItem & ".00"
                End If
            End If
    
            objDocument.ActiveDocument.FormFields(13).result = iCost & ".00"
    End Select
    
    objDocument.Documents.Save
    'objDocument.Save
    'objDocument.Close
    Close #1
    
    MsgBox "Form Completed."
End Sub

Private Function GreaterThan(str1 As String, str2 As String)
    Dim i1Start, i1End, i1Len As Integer
    Dim i2Start, i2End, i2Len As Integer
    Dim iShort As Integer
    
    i1Len = Len(str1)
    i2Len = Len(str2)
    
    If i1Len < i2Len Then
        iShort = i1Len
    Else
        iShort = i2Len
    End If
    
    For i = 1 To iShort
        If Not Left(str1, i) = Left(str2, i) Then GoTo Exit_Next
    Next i
Exit_Next:
    
End Function
