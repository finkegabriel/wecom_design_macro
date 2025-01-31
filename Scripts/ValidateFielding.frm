VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateFielding 
   Caption         =   "Validate Fielding Data"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4095
   OleObjectBlob   =   "ValidateFielding.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateFielding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iConverted As Integer

Private Sub cbCheck_Click()
    cbPole.Value = True
    cbLL.Value = True
    cbGuys.Value = True
    cbBldgs.Value = True
    cbTrim.Value = True
    cbMM.Value = True
    cbNotes.Value = True
    cbBores.Value = True
End Sub

Private Sub cbLL_Click()
    If cbLL.Value = True Then cbPole.Value = True
End Sub

Private Sub cbUncheck_Click()
    cbPole.Value = False
    cbLL.Value = False
    cbGuys.Value = False
    cbBldgs.Value = False
    cbTrim.Value = False
    cbMM.Value = False
    cbNotes.Value = False
    cbBores.Value = False
End Sub

Private Sub cbValidate_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vAttList As Variant
    Dim strLine, strTemp As String
    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem
    Dim strTo, strSubject, strBody As String
        
    If cbPole.Value = True Then Call ValidatePoles
    
    If cbGuys.Value = True Then Call ValidateGuys
    
    If cbBldgs.Value = True Then Call ValidateBldgs
    
    If cbTrim.Value = True Then Call ValidateTrim
    
    If cbNotes.Value = True Then Call ValidateNotes
    
    If cbBores.Value = True Then Call CircleBores

    grpCode(0) = 2
    grpValue(0) = "Development"
    filterType = grpCode
    filterValue = grpValue
        
    On Error Resume Next
    Err = 0
        
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    End If
    
GoTo Skip_This
        
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    strLine = ""
        
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If Not vAttList(0).TextString = "" Then GoTo Next_objBlock
        If vAttList(1).TextString = "" Then GoTo Next_objBlock
        
        If strLine = "" Then
            strLine = vAttList(1).TextString & vbCr
        Else
            strLine = strLine & vbCr & vbCr & vAttList(1).TextString & vbCr
        End If
        
        If Not vAttList(2).TextString = "" Then strLine = strLine & "  Phase: " & vAttList(2).TextString
        If Not vAttList(3).TextString = "" Then strLine = strLine & "  Sec: " & vAttList(3).TextString
        If Not Right(strLine, 1) = vbCr Then strLine = strLine & vbCr
        
        Select Case UCase(vAttList(4).TextString)
            Case "AB"
                strLine = strLine & "APARTMENT BUILDING"
            Case "AC"
                strLine = strLine & "APARTMENT COMPLEX"
            Case "BP"
                strLine = strLine & "BUSINESS PARK"
            Case "C"
                strLine = strLine & "CONDOS"
            Case "IP"
                strLine = strLine & "INDUSTIAL PARK"
            Case "SD"
                strLine = strLine & "SUBDIVISION"
            Case "SM"
                strLine = strLine & "STRIP MALL"
            Case "TH"
                strLine = strLine & "TOWNHOMES"
            Case "TP"
                strLine = strLine & "TRAILER PARK"
            Case Else
                strLine = strLine & UCase(vAttList(4).TextString)
        End Select
        
        If Not vAttList(5).TextString = "" Then strLine = strLine & " Location:" & vbTab & vAttList(5).TextString
Next_objBlock:
    Next objBlock
    
    If Not strLine = "" Then
        strSubject = "New Development(s) found while fielding " & Replace(UCase(ThisDrawing.Name), ".DWG", "")
        'strTo = "jon.wilburn@integrity-us.com"
        strTo = "rich.taylor@integrity-us.com"
        
        strBody = Replace(UCase(strLine), vbCr, "<br>")
        strBody = Replace(UCase(strBody), vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
        strBody = Replace(strBody, vbCr, "<br>")
        
        Set objOutlook = New Outlook.Application
        Set objMail = objOutlook.CreateItem(olMailItem)
    
        objMail.To = strTo
        objMail.Subject = strSubject
        objMail.HTMLBody = strBody
    
        objMail.Send
        
        strLine = "Email sent to Jon about new developments:" & vbCr & vbCr & strLine
        MsgBox strLine
    End If
    
Skip_This:
    
    grpValue(0) = "Planning Note"
    filterType = grpCode
    filterValue = grpValue
    Err = 0
    
    objSS.Clear
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    strLine = ""
        
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = "" Then GoTo Next_Note
        If vAttList(0).TextString = "None" Then GoTo Next_Note
        
        If strLine = "" Then
            strLine = UCase(vAttList(0).TextString)
        Else
            strLine = strLine & vbCr & vbCr & UCase(vAttList(0).TextString)
        End If
        
        If Not vAttList(1).TextString = "" Then strLine = strLine & "  AT  " & UCase(vAttList(1).TextString)
        If Not vAttList(2).TextString = "" Then strLine = strLine & vbCr & " Note(s):  " & UCase(vAttList(2).TextString)
        strLine = strLine & vbCr & "TN83F:" & vbTab & objBlock.InsertionPoint(1) & "," & objBlock.InsertionPoint(0)
Next_Note:
    Next objBlock
    
    If Not strLine = "" Then
        strSubject = "Fielding Notes for Planning " & Replace(UCase(ThisDrawing.Name), ".DWG", "")
        'strTo = "jon.wilburn@integrity-us.com;adam.kemper@integrity-us.com"
        strTo = "rich.taylor@integrity-us.com"
        
        strBody = Replace(UCase(strLine), vbCr, "<br>")
        strBody = Replace(UCase(strBody), vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
        strBody = Replace(strBody, vbCr, "<br>")
        
        Set objOutlook = New Outlook.Application
        Set objMail = objOutlook.CreateItem(olMailItem)
    
        objMail.To = strTo
        objMail.Subject = strSubject
        objMail.HTMLBody = strBody
    
        objMail.Send
        
        strLine = "Email sent to Planning with Fielders' Notes:" & vbCr & vbCr & strLine
        MsgBox strLine
    End If
    
    If cbMM.Value = True Then
        grpCode(0) = 2
        grpValue(0) = "mile marker,mile interstate"
        filterType = grpCode
        filterValue = grpValue
        
        objSS.Select acSelectionSetAll, , , filterType, filterValue
        
        For Each objEntity In objSS
            objEntity.Layer = "Integrity Notes"
        Next objEntity
    End If
    
    objSS.Clear
    objSS.Delete
    
    Me.Hide
    
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub ValidatePoles()
    Dim objSS As AcadSelectionSet
    Dim objPointBlock As AcadBlockReference
    Dim objNewBlock As AcadBlockReference
    Dim objEntity As AcadEntity
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vAttOld, vTemp, vLine As Variant
    Dim dScale As Double
    Dim iCount As Integer
    Dim strLine, strTemp As String
    Dim dN, dE As Double
    Dim vLL As Variant
    
    On Error Resume Next
    
    'dScale = 0.75
    dScale = CDbl(cbScale.Value) / 100
    iCount = 0

    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    If Not Err = 0 Then GoTo Exit_Sub
        
    For Each objEntity In objSS
        Set objPointBlock = objEntity
        vAttOld = objPointBlock.GetAttributes
        
        dE = objPointBlock.InsertionPoint(0)
        dN = objPointBlock.InsertionPoint(1)
        vLL = TN83FtoLL(CDbl(dN), CDbl(dE))
          
        For i = 0 To UBound(vAttOld)
            Select Case i
                Case 2
                    vAttOld(2).TextString = UCase(vAttOld(2).TextString)
                    Select Case vAttOld(2).TextString
                        Case "MTEMC", "LUS", "DREMC", "NES", "SPWS", "DRE", "MTE", "MED"
                            objPointBlock.Layer = "Integrity Poles-Power"
                        Case Else
                            objPointBlock.Layer = "Integrity Poles-Other"
                    End Select
                Case 3
                    If InStr(vAttOld(3).TextString, "=") > 0 Then
                        vTemp = Split(vAttOld(3).TextString, "=")
                        vAttOld(3).TextString = vTemp(1)
                    End If
                Case 4
                    If Not vAttOld(4).TextString = "" Then
                        strTemp = Replace(vAttOld(4).TextString, " 1/2", ".5")
                        strTemp = Replace(strTemp, " 1/4", ".25")
                        
                        vLine = Split(strTemp, " ")
                        For n = 0 To UBound(vLine)
                            If vLine(n) = "NA" Then
                                vLine(n) = ""
                            Else
                                If InStr(vLine(n), "=") = 0 Then vLine(n) = "??=" & vLine(n)
                            End If
                        Next n
                        
                        strLine = ""
                        
                        For n = 0 To UBound(vLine)
                            If Not vLine(n) = "" Then
                                If strLine = "" Then
                                    strLine = vLine(n)
                                Else
                                    strLine = strLine & " " & vLine(n)
                                End If
                            End If
                        Next n
                        
                        vAttOld(4).TextString = strLine
                    End If
                Case 5
                    vAttOld(5).TextString = Replace(vAttOld(5).TextString, " ", "")
                Case 7
                    If cbLL.Value = True Then vAttOld(7).TextString = vLL(0) & "," & vLL(1)
                Case 8
                    Select Case UCase(vAttOld(8).TextString)
                        Case "M"
                            vAttOld(8).TextString = "MGNV"
                        Case "T"
                            vAttOld(8).TextString = "TGB"
                        Case "B"
                            vAttOld(8).TextString = "BROKEN GRD"
                        Case Else
                            vAttOld(8).TextString = "NO GRD"
                    End Select
                Case 9 To 26
                    strLine = Replace(vAttOld(i).TextString, "  ", " ")
                    strLine = Replace(strLine, "= ", "=")
                    
                    If Right(strLine, 1) = " " Then
                        strLine = Left(strLine, Len(strLine) - 1)
                    End If
                    If Left(strLine, 1) = " " Then
                        strLine = Right(strLine, Len(strLine) - 1)
                    End If
                    'If InStr(strLine, "LEVEL 3") > 0 Then strLine = Replace(strLine, "LEVEL 3", "LEVEL_3")
                    vAttOld(i).TextString = strLine
            End Select
        Next i
        
        If Not vAttOld(15).TextString = "" Then
            strLine = Replace(vAttOld(15).TextString, "UTC ", "")
            vTemp = Split(strLine, " ")
            If Not Left(vTemp(0), 1) = "T" Then strLine = vTemp(0) & "T"
            If UBound(vTemp) > 0 Then
                For i = 1 To UBound(vTemp)
                    strLine = strLine & " " & vTemp(i)
                    If Not Left(vTemp(i), 1) = "T" Then strTemp = strTemp & "T"
                Next i
            End If
            vAttOld(15).TextString = strLine
        End If
        
        objPointBlock.Update
Next_objEntity:
    Next objEntity
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub

Private Sub ValidateGuys()
    ValidateGuyType ("ExGuyOL")
    ValidateGuyType ("ExGuyOR")
    ValidateGuyType ("ExAncOL")
    ValidateGuyType ("ExAncOR")
End Sub

Private Sub ValidateGuyType(strName As String)
    Dim objSS5 As AcadSelectionSet
    Dim objPointBlock As AcadBlockReference
    Dim objNewBlock As AcadBlockReference
    Dim objEntity As AcadEntity
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vAttOld As Variant
    Dim dScale As Double
    Dim iCount As Integer
    
    On Error Resume Next
    
    'dScale = 0.75
    dScale = CDbl(cbScale.Value) / 100
    iCount = 0

    grpCode(0) = 2
    grpValue(0) = strName
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS5 = ThisDrawing.SelectionSets.Add("objSS5")
    objSS5.Select acSelectionSetAll, , , filterType, filterValue
    If Not Err = 0 Then GoTo Exit_Sub
        
    For Each objEntity In objSS5
        Set objPointBlock = objEntity
        vAttOld = objPointBlock.GetAttributes
                
        Select Case strName
            Case "ExGuyOL", "ExGuyOR"
                If UBound(vAttOld) = 2 Then
                    vAttOld(2).TextString = "PE1-3G"
                Else
                    vAttOld(3).TextString = "PE1-3G"
                End If
            Case "ExAncOL", "ExAncOR"
                vAttOld(0).TextString = "PF1-5A"
        End Select
        objPointBlock.Update
Next_objEntity:
    Next objEntity
    
Exit_Sub:
    objSS5.Clear
    objSS5.Delete
End Sub

Private Sub ValidateBldgs()
    'ValidateBldgType ("Integrity Building-BUS")
    'ValidateBldgType ("Integrity Building-RES")
    'ValidateBldgType ("Integrity Building-TRL")
    'ValidateBldgType ("Integrity Building-MDU")
    'ValidateBldgType ("Integrity Building-SCH")
    'ValidateBldgType ("Integrity Building-CHU")
    'ValidateBldgType ("Integrity Building Misc")
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim iCount As Integer
    
    On Error Resume Next
    
    iCount = 0
    
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    MsgBox "Found:  " & objSS.count
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        vAttList(5).TextString = UCase(vAttList(5).TextString)
        
        Select Case vAttList(5).TextString
            Case "", "R"
                If Not vAttList(0).TextString = "RESIDENCE" Then
                    vAttList(0).TextString = "RESIDENCE"
                    iCount = iCount + 1
                End If
            Case "B"
                If Not vAttList(0).TextString = "BUSINESS" Then
                    vAttList(0).TextString = "BUSINESS"
                    iCount = iCount + 1
                End If
            Case "C"
                If Not vAttList(0).TextString = "CHURCH" Then
                    vAttList(0).TextString = "CHURCH"
                    iCount = iCount + 1
                End If
            Case "M"
                If Not vAttList(0).TextString = "MDU" Then
                    vAttList(0).TextString = "MDU"
                    iCount = iCount + 1
                End If
            Case "S"
                If Not vAttList(0).TextString = "SCHOOL" Then
                    vAttList(0).TextString = "SCHOOL"
                    iCount = iCount + 1
                End If
            Case "T"
                If Not vAttList(0).TextString = "TRAILER" Then
                    vAttList(0).TextString = "TRAILER"
                    iCount = iCount + 1
                End If
            Case "X"
                If Not vAttList(0).TextString = "EXTENSION" Then
                    vAttList(0).TextString = "EXTENSION"
                    iCount = iCount + 1
                End If
        End Select
        
        objBlock.Update
        
Next_objBlock:
    Next objBlock
    
    MsgBox "Converted  " & iCount & "  customers"
End Sub

Private Sub ValidateTrim()
    Dim objSS3 As AcadSelectionSet
    Dim objTrim As AcadBlockReference
    Dim attList As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim entBlock As AcadEntity
    Dim iCount As Integer
    
    'Me.Hide
    iCount = 0
    
    grpCode(0) = 2
    grpValue(0) = "__Trim"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS3.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objTrim In objSS3
        'If Not entBlock.ObjectName = "AcDbBlockReference" Then GoTo Next_entBlock
        'Set objTrim = entBlock
        'If Not objTrim.Name = "__Trim" Then GoTo Next_entBlock
        
        attList = objTrim.GetAttributes
        attList(0).TextString = "T.T=" & attList(0).TextString & "'"
        objTrim.Layer = "Integrity Notes"
        objTrim.Update
        
        iCount = iCount + 1
Next_entBlock:
    Next objTrim
    
    objSS3.Clear
    objSS3.Delete
    
    MsgBox "Trim Blocks: " & iCount
    'Me.show
End Sub

Private Sub ValidateNotes()
    Dim objSS3 As AcadSelectionSet
    Dim objTrim As AcadBlockReference
    Dim attList As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim entBlock As AcadEntity
    Dim txtLine As AcadText
    Dim mtxtLine As AcadMText
    
    'Me.Hide
    
    grpCode(0) = 8
    grpValue(0) = "Annotations"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    objSS3.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each entBlock In objSS3
        Select Case entBlock.ObjectName
            Case "AcDbText"
                Set txtLine = entBlock
                txtLine.TextString = UCase(txtLine.TextString)
                txtLine.Height = 6
                txtLine.Update
            Case "AcDbMText"
                Set mtxtLine = entBlock
                mtxtLine.TextString = UCase(mtxtLine.TextString)
                mtxtLine.Height = 6
                mtxtLine.Update
        End Select
Next_entBlock:
    Next entBlock
    
    objSS3.Clear
    objSS3.Delete
    
    'Me.show
End Sub

Private Sub CircleBores()
    Dim objCircle As AcadCircle
    Dim objText As AcadText
    Dim objMText As AcadMText
    Dim objBlock As AcadBlockReference
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim vAttList, vTemp, vCoords As Variant
    Dim strLine, strLength As String
    Dim dRadius As Double
    
    On Error Resume Next
    
    Me.Hide
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    grpCode(0) = 2
    grpValue(0) = "Drop"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        strLine = UCase(vAttList(0).TextString)
        
        If InStr(strLine, "BORE") > 0 Then
            vTemp = Split(strLine, "=")
            strLength = Replace(vTemp(1), " ", "")
            strLength = Replace(strLength, "'", "")
            
            dRadius = CDbl(strLength) / 2
            If dRadius < 1 Then GoTo Next_objBlock
            
            vCoords = objBlock.InsertionPoint
            
            Set objCircle = ThisDrawing.ModelSpace.AddCircle(vCoords, dRadius)
            objCircle.Layer = objBlock.Layer
            objCircle.Update
        End If
Next_objBlock:
    Next objBlock
    
    objSS.Clear
    
    grpCode(0) = 0
    grpValue(0) = "TEXT"
    filterType = grpCode
    filterValue = grpValue
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For Each objText In objSS
        strLine = UCase(objText.TextString)
        
        If InStr(strLine, "BORE") > 0 Then
            vTemp = Split(strLine, "=")
            strLength = Replace(vTemp(1), " ", "")
            strLength = Replace(strLength, "'", "")
            
            dRadius = CDbl(strLength) / 2
            If dRadius < 1 Then GoTo Next_ObjText
            
            vCoords = objText.InsertionPoint
            
            Set objCircle = ThisDrawing.ModelSpace.AddCircle(vCoords, dRadius)
            objCircle.Layer = objText.Layer
            objCircle.Update
        End If
Next_ObjText:
    Next objText
    
    objSS.Clear
    
    grpCode(0) = 0
    grpValue(0) = "MTEXT"
    filterType = grpCode
    filterValue = grpValue
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For Each objMText In objSS
        strLine = UCase(objMText.TextString)
        
        If InStr(strLine, "BORE") > 0 Then
            vTemp = Split(strLine, "=")
            strLength = Replace(vTemp(1), " ", "")
            strLength = Replace(strLength, "'", "")
            
            dRadius = CDbl(strLength) / 2
            If dRadius < 1 Then GoTo Next_ObjMText
            
            vCoords = objMText.InsertionPoint
            
            Set objCircle = ThisDrawing.ModelSpace.AddCircle(vCoords, dRadius)
            objCircle.Layer = objMText.Layer
            objCircle.Update
        End If
Next_ObjMText:
    Next objMText
    
    objSS.Clear
    objSS.Delete
Exit_Sub:
    Me.show
End Sub

Private Sub UserForm_Initialize()
    cbScale.AddItem ""
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "75"
    
    iConverted = 0
End Sub

Private Function TN83FtoLL(dNorth As Double, dEast As Double)
    Dim dLat, dLong, dDLat As Double
    Dim dDiffE, dEast0 As Double
    Dim dDiffN, dNorth0 As Double
    Dim dR, dCAR, dCA, dU, dK As Double
    Dim LL(2) As Double
    
    dEast = dEast * 0.3048006096
    dNorth = dNorth * 0.3048006096
    
    dDiffE = dEast - 600000
    dDiffN = dNorth - 166504.1691
    
    dR = 8842127.1422 - dDiffN
    dCAR = Atn(dDiffE / dR)
    dCA = dCAR * 180 / 3.14159265359
    dLong = -86 + dCA / 0.585439726459
    
    dU = dDiffN - dDiffE * Tan(dCAR / 2)
    dDLat = dU * (0.00000901305249 + dU * (-6.77268E-15 + dU * (-3.72351E-20 + dU * -9.2828E-28)))
    dLat = 35.8340607459 + dDLat
    
    dK = 0.999948401424 + (1.23188E-14 * dU * dU) + (4.54E-22 * dU * dU * dU)
    
    LL(0) = dLat
    LL(1) = dLong
    LL(2) = dK
    
    TN83FtoLL = LL
End Function
