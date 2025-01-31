VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TabGeneric 
   Caption         =   "Tabulation"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14310
   OleObjectBlob   =   "TabGeneric.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TabGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSS As AcadSelectionSet

Private Sub cbCopyAll_Click()
    Dim strFileName As String
    Dim strDWGName As String
    Dim vLine As Variant
    Dim objData As New DataObject
    Dim strCopy, strTemp As String
    Dim vUnit As Variant
    
    If cbCopyToFile.Value = True Then
        strFileName = ThisDrawing.Path & "\"
        vLine = Split(ThisDrawing.Name, " ")
        strFileName = strFileName & vLine(0) & "-TAB Totals.txt"

        Open strFileName For Output As #1
        
        For a = 0 To lbFinal.ListCount - 1
            Print #1, lbFinal.List(a, 0) & vbTab & lbFinal.List(a, 1)
        Next a
        
        Close #1
    
        MsgBox "Saved to File"
    Else
        strCopy = lbFinal.List(0, 0) & vbTab & lbFinal.List(0, 1)
        For a = 1 To lbFinal.ListCount - 1
            strCopy = strCopy & vbCr & lbFinal.List(a, 0) & vbTab & lbFinal.List(a, 1)
        Next a
        
        objData.SetText strCopy
        objData.PutInClipboard
    
        MsgBox "Copied to Clipboard"
    End If
End Sub

Private Sub cbCopyTab_Click()
    Dim objData As New DataObject
    Dim strCopy As String
    Dim vLine As Variant
    
    If InStr(lbFinal.List(0, 0), " ") > 0 Then
       vLine = Split(lbFinal.List(0, 0), " ")
       strCopy = vLine(0) & vbTab & vLine(1) & vbTab & lbFinal.List(0, 1)
    Else
        strCopy = lbFinal.List(0, 0) & vbTab & vbTab & lbFinal.List(0, 1)
    End If
    
    For a = 1 To lbFinal.ListCount - 1
        If InStr(lbFinal.List(a, 0), " ") > 0 Then
            vLine = Split(lbFinal.List(a, 0), " ")
            strCopy = strCopy & vbCr & vLine(0) & vbTab & vLine(1) & vbTab & lbFinal.List(a, 1)
        Else
            strCopy = strCopy & vbCr & lbFinal.List(a, 0) & vbTab & vbTab & lbFinal.List(a, 1)
        End If
    Next a
        
    objData.SetText strCopy
    objData.PutInClipboard
    
    MsgBox "Copied to Clipboard"
End Sub

Private Sub cbExport_Click()
    Dim strFileName As String
    Dim strDWGName As String
    Dim str1, str2, str3, strLine As String
    Dim strItem2, strSize As String
    Dim vLine, vHC As Variant
    Dim iTest As Integer
    
    Load FileNameForm
    FileNameForm.show
    strFileName = FileNameForm.tbFileName.Value
    Unload FileNameForm
    
    If Right(strFileName, 4) = ".txt" Then
        strFileName = Left(strFileName, Len(strFileName) - 4)
    End If
    strFileName = strFileName & "-TAB.txt"

    Open strFileName For Output As #1
    
    iTest = 0
    
    For i = 0 To (lbUnits.ListCount - 1)
        strLine = lbUnits.List(i)
        
        Print #1, strLine
    Next i
    
    If lbNotFound.ListCount > 0 Then
        For i = 0 To (lbNotFound.ListCount - 1)
            strLine = lbNotFound.List(i)
        
            Print #1, strLine
        Next i
    End If
    
    MsgBox "Export Complete"
    Close #1
End Sub

Private Sub cbExportTab_Click()
    Dim strFileName As String
    'Dim strDWGName As String
    Dim vLine As Variant
    'Dim objData As New DataObject
    'Dim strCopy, strTemp As String
    'Dim vUnit As Variant
    
    strFileName = ThisDrawing.Path & "\"
    vLine = Split(ThisDrawing.Name, " ")
    strFileName = strFileName & vLine(0) & "-TAB Totals.txt"

    Open strFileName For Output As #1
        
    For a = 0 To lbFinal.ListCount - 1
        Print #1, Replace(lbFinal.List(a, 0), "  ", vbTab) & vbTab & lbFinal.List(a, 1)
    Next a
        
    Close #1
    
    MsgBox "Saved to File"
End Sub

Private Sub cbGetFromBlocks_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    Dim vAttList, vUnits, vClosure As Variant
    Dim vItem, vTemp, vRich As Variant
    Dim vReturnPnt, vCoords As Variant
    Dim filterType, filterValue As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim grpValue(0) As Variant
    Dim grpCode(0) As Integer
    Dim str1, str2, strUnit As String
    Dim strExchange As String
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim iRes, iBus As Integer
    Dim iAF, iBF, iUF, iUNK As Integer
    Dim iHO1A, iHO1B As Integer
    Dim iTemp, iCounter As Integer
    Dim dCoords() As Double
    
    iRes = 0: iBus = 0
    iAF = 0: iBF = 0: iUF = 0: iUNK = 0
    strExchange = ""

    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH,sPanel,sMH,Customer,cable_span,SS Info"
    filterType = grpCode
    filterValue = grpValue
    
    lbUnits.Clear
    lbNotFound.Clear
    
  On Error Resume Next
    Me.Hide
    
    objSS.Clear
    
    Select Case cbSelection.Value
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
        Case Else
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
    
    
  
    'Err = 0
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadBlockReference Then
            Set objBlock = objEntity
            
            If objBlock.Layer = "Integrity Future" Then GoTo Next_Object
            
            Select Case objBlock.Name
                Case "SS Info"
                    vAttList = objBlock.GetAttributes
                    If Not vAttList(2).TextString = "" Then strExchange = vAttList(2).TextString
                Case "sPole"
                    iHO1A = 0
                    iHO1B = 0
                    vAttList = objBlock.GetAttributes
                    
                    If vAttList(0).TextString = "POLE" Then GoTo Next_Object
                    If vAttList(27).TextString = "" Then GoTo Next_Object
                    
                    str1 = vAttList(0).TextString & vbTab & "sPole" & vbTab
                    vUnits = Split(vAttList(27).TextString, ";;")
                    
                    'MsgBox vAttList(0).TextString & vbCr & UBound(vUnits)
                    
                    For i = 0 To UBound(vUnits)
                        vItem = Split(vUnits(i), "=")
                        If UBound(vItem) = 0 Then
                            str2 = str1 & vItem(0) & vbTab & "1" & vbTab
                        Else
                            vTemp = Split(vItem(1), "  ")
                            If UBound(vTemp) = 0 Then
                                str2 = str1 & vItem(0) & vbTab & vItem(1) & vbTab
                            Else
                                str2 = str1 & vItem(0) & vbTab & vTemp(0) & vbTab & vTemp(1)
                            End If
                            str2 = Replace(str2, "'", "")
                        End If
                        str2 = Replace(str2, "+", "")
                        'MsgBox str2
                        'lbUnits.AddItem str2
                        
                        vRich = Split(str2, vbTab)
                        
                        lbUnits.AddItem vRich(0)
                        lbUnits.List(lbUnits.ListCount - 1, 1) = vRich(1)
                        lbUnits.List(lbUnits.ListCount - 1, 2) = vRich(2)
                        lbUnits.List(lbUnits.ListCount - 1, 3) = vRich(3)
                        lbUnits.List(lbUnits.ListCount - 1, 4) = vRich(4)
                    Next i
                    
                    vUnits = Split(vAttList(26).TextString, " + ")
                    For i = 0 To UBound(vUnits)
                        vItem = Split(vUnits(i), ": ")
                        If UBound(vItem) = 0 Then GoTo Next_sPole_HO1
                        vTemp = Split(vItem(1), "-")
                        If UBound(vTemp) = 0 Then
                            iHO1A = iHO1A + 1
                        Else
                            iHO1A = iHO1A + CInt(vTemp(1)) - CInt(vTemp(0)) + 1
                        End If
Next_sPole_HO1:
                    Next i
                    
                    If iHO1A > 0 Then
                        str2 = str1 & "HO1A" & vbTab & iHO1A & vbTab
                        'lbUnits.AddItem str2
                        
                        vRich = Split(str2, vbTab)
                        
                        lbUnits.AddItem vRich(0)
                        lbUnits.List(lbUnits.ListCount - 1, 1) = vRich(1)
                        lbUnits.List(lbUnits.ListCount - 1, 2) = vRich(2)
                        lbUnits.List(lbUnits.ListCount - 1, 3) = vRich(3)
                        lbUnits.List(lbUnits.ListCount - 1, 4) = vRich(4)
                    End If
                Case "sPed", "sHH", "sPanel", "sMH"
                    iHO1A = 0
                    iHO1B = 0
                    vAttList = objBlock.GetAttributes
                    
                    Select Case vAttList(0).TextString
                        Case "PED", "HH", "PANEL", "MH", ""
                            GoTo Next_Object
                    End Select
                    
                    If vAttList(7).TextString = "" Then GoTo Next_Object
                    
                    str1 = vAttList(0).TextString & vbTab & objBlock.Name & vbTab
                    vUnits = Split(vAttList(7).TextString, ";;")
                    
                    For i = 0 To UBound(vUnits)
                        vItem = Split(vUnits(i), "=")
                        If UBound(vItem) = 0 Then GoTo Next_sPed_HO1
                        If UBound(vItem) = 0 Then
                            str2 = str1 & vItem(0) & vbTab & "1" & vbTab
                        Else
                            vTemp = Split(vItem(1), "  ")
                            If UBound(vTemp) = 0 Then
                                str2 = str1 & vItem(0) & vbTab & vItem(1) & vbTab
                            Else
                                str2 = str1 & vItem(0) & vbTab & vTemp(0) & vbTab & vTemp(1)
                            End If
                            str2 = Replace(str2, "'", "")
                        End If
                        str2 = Replace(str2, "+", "")
                        'lbUnits.AddItem str2
                        
                        vRich = Split(str2, vbTab)
                        
                        lbUnits.AddItem vRich(0)
                        lbUnits.List(lbUnits.ListCount - 1, 1) = vRich(1)
                        lbUnits.List(lbUnits.ListCount - 1, 2) = vRich(2)
                        lbUnits.List(lbUnits.ListCount - 1, 3) = vRich(3)
                        lbUnits.List(lbUnits.ListCount - 1, 4) = vRich(4)
                    Next i
                    
                    vUnits = Split(vAttList(6).TextString, " + ")
                    For i = 0 To UBound(vUnits)
                        vItem = Split(vUnits(i), ": ")
                        vTemp = Split(vItem(1), "-")
                        If UBound(vTemp) = 0 Then
                            iHO1B = iHO1B + 1
                        Else
                            iHO1B = iHO1B + CInt(vTemp(1)) - CInt(vTemp(0)) + 1
                        End If
Next_sPed_HO1:
                    Next i
                    
                    If iHO1B > 0 Then
                        str2 = str1 & "HO1B" & vbTab & iHO1B & vbTab
                        'lbUnits.AddItem str2
                        
                        vRich = Split(str2, vbTab)
                        
                        lbUnits.AddItem vRich(0)
                        lbUnits.List(lbUnits.ListCount - 1, 1) = vRich(1)
                        lbUnits.List(lbUnits.ListCount - 1, 2) = vRich(2)
                        lbUnits.List(lbUnits.ListCount - 1, 3) = vRich(3)
                        lbUnits.List(lbUnits.ListCount - 1, 4) = vRich(4)
                    End If
                Case "Customer"
                    vAttList = objBlock.GetAttributes
                    
                    Select Case vAttList(5).TextString
                        Case "R", "C", "M", "T", ""
                            iRes = iRes + 1
                            
                            GoTo Next_Object
                        Case "B", "S"
                            iBus = iBus + 1
                            
                            GoTo Next_Object
                    End Select
                Case "cable_span"
                    vAttList = objBlock.GetAttributes
                    str1 = Replace(vAttList(2).TextString, "'", "")
                    
                    Select Case objBlock.Layer
                        Case "Integrity Cable-Aerial Text"
                            iAF = iAF + CInt(str1)
                        Case "Integrity Cable-Buried Text"
                            iBF = iBF + CInt(str1)
                        Case Else
                            iUNK = iUNK + CInt(str1)
                    End Select
                    
                    GoTo Next_Object
                Case "SS Info"
                    vAttList = objBlock.GetAttributes
                    If Not vAttList(2).TextString = "" Then strExchange = vAttList(2).TextString
                    
                    GoTo Next_Object
                Case Else
                    GoTo Next_Object
            End Select
        End If
Next_Object:
    Next objEntity
    
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "RU" & vbTab & iRes & vbTab
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "BU" & vbTab & iBus & vbTab
    
    'If iAF > 0 Then lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Aerial" & vbTab & iAF & vbTab
    'If iBF > 0 Then lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Buried" & vbTab & iBF & vbTab
    'If iUF > 0 Then lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "UG" & vbTab & iUF & vbTab
    'If iUNK > 0 Then lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "UNKNOWN" & vbTab & iUNK & vbTab
                        
    lbUnits.AddItem "Data"
    lbUnits.List(lbUnits.ListCount - 1, 1) = ""
    lbUnits.List(lbUnits.ListCount - 1, 2) = "RU"
    lbUnits.List(lbUnits.ListCount - 1, 3) = iRes
    lbUnits.List(lbUnits.ListCount - 1, 4) = ""
                        
    lbUnits.AddItem "Data"
    lbUnits.List(lbUnits.ListCount - 1, 1) = ""
    lbUnits.List(lbUnits.ListCount - 1, 2) = "BU"
    lbUnits.List(lbUnits.ListCount - 1, 3) = iBus
    lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    
    If iAF > 0 Then
        lbUnits.AddItem "Data"
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = "Aerial"
        lbUnits.List(lbUnits.ListCount - 1, 3) = iAF
        lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    End If
    
    If iBF > 0 Then
        lbUnits.AddItem "Data"
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = "Buried"
        lbUnits.List(lbUnits.ListCount - 1, 3) = iBF
        lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    End If
    
    If iUF > 0 Then
        lbUnits.AddItem "Data"
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = "UG"
        lbUnits.List(lbUnits.ListCount - 1, 3) = iUF
        lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    End If
    
    If iUNK > 0 Then
        lbUnits.AddItem "Data"
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = "Unknown"
        lbUnits.List(lbUnits.ListCount - 1, 3) = iUNK
        lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    End If
    
    str1 = Replace(UCase(ThisDrawing.Name), ".DWG", "")
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Project" & vbTab & str1 & vbTab
    lbUnits.AddItem "Data"
    lbUnits.List(lbUnits.ListCount - 1, 1) = ""
    lbUnits.List(lbUnits.ListCount - 1, 2) = "Project"
    lbUnits.List(lbUnits.ListCount - 1, 3) = str1
    lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Exchange" & vbTab & strExchange & vbTab
    lbUnits.AddItem "Data"
    lbUnits.List(lbUnits.ListCount - 1, 1) = ""
    lbUnits.List(lbUnits.ListCount - 1, 2) = "Exchange"
    lbUnits.List(lbUnits.ListCount - 1, 3) = strExchange
    lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    
Exit_Sub:

    Me.show
End Sub

Private Sub cbGetSpans_Click()
    If lbSpans.ListCount = 0 Then Exit Sub
    
    Dim strFileName As String
    Dim strDWGName As String
    Dim vLine As Variant
    
    Dim objData As New DataObject
    Dim strCopy As String
    Dim vUnit As Variant
    
    strCopy = ""
    
    For i = 0 To lbSpans.ListCount - 1
        If i = 0 Then
            strCopy = lbSpans.List(i, 0) & "," & lbSpans.List(i, 2)
        Else
            strCopy = strCopy & vbCr & lbSpans.List(i, 0) & "," & lbSpans.List(i, 2)
        End If
    Next i
    
    If cbCopyToFile.Value = True Then
        strFileName = ThisDrawing.Path & "\"
        vLine = Split(ThisDrawing.Name, " ")
        strFileName = strFileName & vLine(0) & "-TAB Spans.csv"

        Open strFileName For Output As #1
        
        Print #1, strCopy
        Close #1
    
        MsgBox "Copied to File"
    Else
        objData.SetText strCopy
        objData.PutInClipboard
    
        MsgBox "Copied to Clipboard"
    End If
End Sub

Private Sub cbGetTotals_Click()
    Dim vUnit, vTemp As Variant
    Dim strTemp As String
    
    'On Error Resume Next
    
    lbSpans.Clear
    lbFinal.Clear
    
    For i = 0 To lbUnits.ListCount - 1
        'vUnit = Split(lbUnits.List(k), vbTab)
        'If vUnit(0) = "xx" Then GoTo Next_K
        'If vUnit(0) = "Data" Then GoTo Next_K
        'If UBound(vUnit) < 3 Then GoTo Next_K
        
        If lbUnits.List(i, 0) = "xx" Then GoTo Next_K
        If lbUnits.List(i, 0) = "Data" Then GoTo Next_K
        If lbUnits.List(i, 1) = "terminal" Then GoTo Next_K
        
        'If vUnit(1) = "terminal" Then GoTo Next_K
        
        'vTemp = Split(vUnit(3), "(")
        vTemp = Split(lbUnits.List(i, 2), "(")
        
        Select Case vTemp(0)
            Case "CO", "BFO", "UO"
                lbSpans.AddItem lbUnits.List(i, 2)
                lbSpans.List(lbSpans.ListCount - 1, 1) = lbUnits.List(i, 4)
                lbSpans.List(lbSpans.ListCount - 1, 2) = lbUnits.List(i, 3)
        End Select
        
        If lbFinal.ListCount = 0 Then
            strTemp = lbUnits.List(i, 2)
            If Not lbUnits.List(i, 4) = "" Then strTemp = strTemp & "  " & lbUnits.List(i, 4)
            
            lbFinal.AddItem strTemp
            lbFinal.List(0, 1) = lbUnits.List(i, 3)
            'lbFinal.List(lbFinal.ListCount - 1, 1) = vUnit(4)
            GoTo Next_K
        End If
        
        For m = 0 To lbFinal.ListCount - 1
            strTemp = lbUnits.List(i, 2)
            'If UBound(vUnit) > 4 Then
            'If Not lbUnits.List(i, 4) = "" Then lbFinal.List(0, 0) = lbFinal.List(0, 0) & "  " & lbUnits.List(i, 4)
            If Not lbUnits.List(i, 4) = "" Then strTemp = strTemp & "  " & lbUnits.List(i, 4)
            'End If
            'If Not Err = 0 Then
            
            If lbFinal.List(m, 0) = strTemp Then
                lbFinal.List(m, 1) = CLng(lbFinal.List(m, 1)) + CLng(lbUnits.List(i, 3))
                GoTo Next_K
            End If
        Next m
        
        lbFinal.AddItem strTemp
        lbFinal.List(lbFinal.ListCount - 1, 1) = lbUnits.List(i, 3)
        'If UBound(vUnit) > 4 Then
            'If Not vUnit(5) = "" Then lbFinal.List(lbFinal.ListCount - 1, 0) = lbFinal.List(lbFinal.ListCount - 1, 0) & "  " & vUnit(5)
        'End If
Next_K:
    Next i
End Sub

Private Sub cbGetUnits_Click()
    Dim objEntity As AcadEntity
    Dim obrTemp As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    Dim attList, vUnit, vClosure As Variant
    Dim vReturnPnt, vCoords As Variant
    Dim vTemp, vRich As Variant
    Dim str1, str2, strUnit As String
    Dim strExchange As String
    Dim filterType, filterValue As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim iRes, iBus As Integer
    Dim iAF, iBF, iUF, iUNK As Integer
    Dim iTemp, iCounter As Integer
    Dim dCoords() As Double
    
    iRes = 0: iBus = 0
    iAF = 0: iBF = 0: iUF = 0: iUNK = 0
    strExchange = ""

    grpCode(0) = 2
    grpValue(0) = "pole_unit,cable_span,SS Info,Customer,Callout"
    'grpValue(0) = "pole_unit,terminal,RES,BUSINESS,TRLR,MDU,CHURCH,SCHOOL,cable_span,SS Info,Customer,Callout"
    filterType = grpCode
    filterValue = grpValue
    
    lbUnits.Clear
    lbNotFound.Clear
    
  On Error Resume Next
    Me.Hide
    
    objSS.Clear
    
    Select Case cbSelection.Value
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
        Case Else
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
    
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadBlockReference Then
            Set obrTemp = objEntity
            
            If obrTemp.Layer = "Integrity Future" Then GoTo Next_Object
            
            Select Case obrTemp.Name
                Case "pole_unit"
                    attList = obrTemp.GetAttributes
                    
                    If attList(0).TextString = "xx" Then GoTo Next_Object
                    If attList(3).TextString = "" Then GoTo Next_Object
                    
                    str1 = attList(0).TextString & vbTab & obrTemp.Name
                    str2 = attList(3).TextString
                    
                    strUnit = Replace(str2, "+", "")
                    vUnit = Split(strUnit, "=")
                    If UBound(vUnit) > 0 Then
                        vUnit(1) = Replace(vUnit(1), "'", "")
                    Else
                        lbNotFound.AddItem str1 & vbTab & strUnit
                        GoTo Next_Object
                    End If
                    
                    If Left(vUnit(0), 4) = "HACO" Or Left(vUnit(0), 4) = "HBFO" Then
                        vClosure = Split(vUnit(1), "  ")
                        str1 = str1 & vbTab & vUnit(0) & vbTab & vClosure(0)
                        If UBound(vClosure) > 0 Then
                            str1 = str1 & vbTab & vClosure(1)
                        Else
                            str1 = str1 & vbTab & vbTab
                        End If
                    Else
                        Select Case Right(vUnit(1), 1)
                            Case "'"
                                vUnit(1) = Left(vUnit(1), Len(vUnit(1)) - 1)
                                str1 = str1 & vbTab & vUnit(0) & vbTab & vUnit(1) & vbTab
                            Case ")"
                                vTemp = Split(vUnit(1), " (")
                                If Right(vTemp(0), 1) = "'" Then vTemp(0) = Left(vTemp(0), Len(vTemp(0)) - 1)
                                vUnit(1) = vTemp(0)
                                str1 = str1 & vbTab & vUnit(0) & vbTab & vTemp(0) & vbTab & Left(vTemp(1), Len(vTemp(1)) - 1)
                            Case Else
                                If InStr(vUnit(1), "  ") > 0 Then
                                    vUnit(1) = Replace(vUnit(1), "  ", vbTab)
                                    str1 = str1 & vbTab & vUnit(0) & vbTab & vUnit(1)
                                Else
                                    str1 = str1 & vbTab & vUnit(0) & vbTab & vUnit(1) & vbTab
                                End If
                        End Select
                    End If
                Case "Callout"
                    attList = obrTemp.GetAttributes
                    
                    If obrTemp.Layer = "Integrity Existing" Then GoTo Next_Object
                    
                    If Not Left(attList(1).TextString, 1) = "+" Then GoTo Next_Object
                    'If InStr(attList(1).TextString, "+HO1") < 1 Then GoTo Next_Object
                    
                    str1 = vbTab & vbTab
                    str2 = Replace(attList(1).TextString, "+", "")
                    vUnit = Split(str2, "=")
                    
                    If UBound(vUnit) > 0 Then
                        str1 = str1 & vUnit(0) & vbTab & vUnit(1) & vbTab
                    Else
                        lbNotFound.AddItem str1 & vbTab & strUnit
                        GoTo Next_Object
                    End If
                Case "Customer"
                    attList = obrTemp.GetAttributes
                    
                    Select Case attList(5).TextString
                        Case "R", "C", "M", "T", "L", ""
                            iRes = iRes + 1
                        Case "B", "S"
                            iBus = iBus + 1
                    End Select
                    
                    GoTo Next_Object
                Case "cable_span"
                    attList = obrTemp.GetAttributes
                    str1 = Replace(attList(2).TextString, "'", "")
                    
                    Select Case obrTemp.Layer
                        Case "Integrity Cable-Aerial Text"
                            iAF = iAF + CInt(str1)
                        Case "Integrity Cable-Buried Text"
                            iBF = iBF + CInt(str1)
                        Case Else
                            iUNK = iUNK + CInt(str1)
                    End Select
                    
                    GoTo Next_Object
                Case "SS Info"
                    attList = obrTemp.GetAttributes
                    If Not attList(2).TextString = "" Then strExchange = attList(2).TextString
                    
                    GoTo Next_Object
                Case Else
                    GoTo Next_Object
            End Select
            
            'lbUnits.AddItem str1
                        
            vRich = Split(str1, vbTab)
                        
            lbUnits.AddItem vRich(0)
            lbUnits.List(lbUnits.ListCount - 1, 1) = vRich(1)
            lbUnits.List(lbUnits.ListCount - 1, 2) = vRich(2)
            lbUnits.List(lbUnits.ListCount - 1, 3) = vRich(3)
            lbUnits.List(lbUnits.ListCount - 1, 4) = vRich(4)
            
        End If
Next_Object:
    Next objEntity
    
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "RU" & vbTab & iRes & vbTab
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "BU" & vbTab & iBus & vbTab
    
    'If iAF > 0 Then lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Aerial" & vbTab & iAF & vbTab
    'If iBF > 0 Then lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Buried" & vbTab & iBF & vbTab
    'If iUF > 0 Then lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "UG" & vbTab & iUF & vbTab
    'If iUNK > 0 Then lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "UNKNOWN" & vbTab & iUNK & vbTab
                        
    lbUnits.AddItem "Data"
    lbUnits.List(lbUnits.ListCount - 1, 1) = ""
    lbUnits.List(lbUnits.ListCount - 1, 2) = "RU"
    lbUnits.List(lbUnits.ListCount - 1, 3) = iRes
    lbUnits.List(lbUnits.ListCount - 1, 4) = ""
                        
    lbUnits.AddItem "Data"
    lbUnits.List(lbUnits.ListCount - 1, 1) = ""
    lbUnits.List(lbUnits.ListCount - 1, 2) = "BU"
    lbUnits.List(lbUnits.ListCount - 1, 3) = iBus
    lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    
    If iAF > 0 Then
        lbUnits.AddItem "Data"
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = "Aerial"
        lbUnits.List(lbUnits.ListCount - 1, 3) = iAF
        lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    End If
    
    If iBF > 0 Then
        lbUnits.AddItem "Data"
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = "Buried"
        lbUnits.List(lbUnits.ListCount - 1, 3) = iBF
        lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    End If
    
    If iUF > 0 Then
        lbUnits.AddItem "Data"
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = "UG"
        lbUnits.List(lbUnits.ListCount - 1, 3) = iUF
        lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    End If
    
    If iUNK > 0 Then
        lbUnits.AddItem "Data"
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = "Unknown"
        lbUnits.List(lbUnits.ListCount - 1, 3) = iUNK
        lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    End If
    
    str1 = Replace(UCase(ThisDrawing.Name), ".DWG", "")
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Project" & vbTab & str1 & vbTab
    lbUnits.AddItem "Data"
    lbUnits.List(lbUnits.ListCount - 1, 1) = ""
    lbUnits.List(lbUnits.ListCount - 1, 2) = "Project"
    lbUnits.List(lbUnits.ListCount - 1, 3) = str1
    lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Exchange" & vbTab & strExchange & vbTab
    lbUnits.AddItem "Data"
    lbUnits.List(lbUnits.ListCount - 1, 1) = ""
    lbUnits.List(lbUnits.ListCount - 1, 2) = "Exchange"
    lbUnits.List(lbUnits.ListCount - 1, 3) = strExchange
    lbUnits.List(lbUnits.ListCount - 1, 4) = ""
    
    'str1 = Replace(UCase(ThisDrawing.Name), ".DWG", "")
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Project" & vbTab & str1 & vbTab
    
    'lbUnits.AddItem "Data" & vbTab & vbTab & vbTab & "Exchange" & vbTab & strExchange & vbTab
    
Exit_Sub:
    
    Me.show
    
    'If Not Err = 0 Then MsgBox Err.Number & vbCr & Err.Description
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbSort_Click()
    Dim strTemp, strTotal As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    
    iCount = lbFinal.ListCount - 1
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbFinal.List(b, 0) > lbFinal.List(b + 1, 0) Then
                strTemp = lbFinal.List(b + 1, 0)
                strTotal = lbFinal.List(b + 1, 1)
                
                lbFinal.List(b + 1, 0) = lbFinal.List(b, 0)
                lbFinal.List(b + 1, 1) = lbFinal.List(b, 1)
                
                lbFinal.List(b, 0) = strTemp
                lbFinal.List(b, 1) = strTotal
            End If
        Next b
    Next a
    
    iOffset = 0
    For c = 1 To lbFinal.ListCount - 1
        iIndex = c - iOffset
        If lbFinal.List(iIndex, 0) = lbFinal.List(iIndex - 1, 0) Then
            lbFinal.List(iIndex - 1, 1) = CInt(lbFinal.List(iIndex - 1, 1)) + CInt(lbFinal.List(iIndex, 1))
            lbFinal.RemoveItem iIndex
            iOffset = iOffset + 1
        End If
    Next c
End Sub

Private Sub cbUpdate_Click()
    Dim iListIndex As Integer
    Dim strLine As String
    
    'strLine = tbPole.Value & vbTab & tbPosition.Value & vbTab & tbDWG.Value & vbTab & tbLine.Value & vbTab & tbAmount.Value & vbTab & tbNote.Value
    
    iListIndex = lbFinal.ListIndex
    
    lbFinal.List(iListIndex, 0) = tbLine.Value
    lbFinal.List(iListIndex, 1) = tbAmount.Value
    
    'lbNotFound.RemoveItem iListIndex
    'lbUnits.AddItem strLine 'tbUnit.Value
    'lbNotFound.ListIndex = iListIndex
    
    tbLine.Value = ""
    tbAmount.Value = ""
End Sub

Private Sub lbFinal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'Dim vUnit, vTemp As Variant
    
    'vUnit = Split(lbNotFound.Value, vbTab)
    
    tbLine.Value = lbFinal.List(lbFinal.ListIndex, 0)
    tbAmount.Value = lbFinal.List(lbFinal.ListIndex, 1)
End Sub

Private Sub lbUnits_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'tbNewValue.Value = ""
    'tbNewValue.Enabled = False
    
    Select Case KeyCode
        Case vbKeyDelete
            lbUnits.RemoveItem lbUnits.ListIndex
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbSpans.Clear
    lbSpans.ColumnCount = 3
    lbSpans.ColumnWidths = "80;30;50"
    
    lbFinal.Clear
    lbFinal.ColumnCount = 2
    lbFinal.ColumnWidths = "140;40"
    
    lbUnits.Clear
    lbUnits.ColumnCount = 5
    lbUnits.ColumnWidths = "108;30;96;48;42"
    
    cbSelection.AddItem "Window"
    cbSelection.AddItem "Polygon"
    cbSelection.Value = "Window"
    
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
End Sub

Private Sub UserForm_Terminate()
    objSS.Clear
    objSS.Delete
End Sub
