VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VerifyAttachmentCallouts 
   Caption         =   "Verify Attachment Callouts"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10335
   OleObjectBlob   =   "VerifyAttachmentCallouts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VerifyAttachmentCallouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSS1 As AcadSelectionSet
Dim objSS2 As AcadSelectionSet
    
Private Sub cbFindDiffer_Click()
    If lbAll.ListCount < 1 Then Exit Sub
    
    Dim vBlock, vCallout As Variant
    Dim strBlock, strCallout As String
    
    For i = 0 To lbAll.ListCount - 1
        If lbAll.List(i, 1) = "" Then GoTo Next_line
        If lbAll.List(i, 2) = "" Then GoTo Next_line
        
        vBlock = Split(lbAll.List(i, 1), ";;")
        vCallout = Split(lbAll.List(i, 2), ";;")
        
        For j = 0 To UBound(vBlock)
            For k = 0 To UBound(vCallout)
                If vBlock(j) = vCallout(k) Then
                    vBlock(j) = ""
                    vCallout(k) = ""
                    
                    GoTo Next_Item
                End If
            Next k
Next_Item:
        Next j
        
        strBlock = ""
        For j = 0 To UBound(vBlock)
            If Not vBlock(j) = "" Then
                If strBlock = "" Then
                    strBlock = vBlock(j)
                Else
                    strBlock = strBlock & ";;" & vBlock(j)
                End If
            End If
        Next j
        
        strCallout = ""
        For j = 0 To UBound(vCallout)
            If Not vCallout(j) = "" Then
                If strCallout = "" Then
                    strCallout = vCallout(j)
                Else
                    strCallout = strCallout & ";;" & vCallout(j)
                End If
            End If
        Next j
        
        lbAll.List(i, 1) = strBlock
        lbAll.List(i, 2) = strCallout
Next_line:
    Next i
    
    For i = lbAll.ListCount - 1 To 0 Step -1
        If lbAll.List(i, 1) = "" Then
            If lbAll.List(i, 2) = "" Then lbAll.RemoveItem i
        End If
    Next i
    
    tbListCount.Value = lbAll.ListCount
End Sub

Private Sub cbGetBlocks_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vItem, vLine, vTemp As Variant
    Dim vPnt1, vPnt2, vCoords As Variant
    Dim strPole, strCompany, strAttach As String
    Dim strExist, strProp, strTemp As String
    Dim strAttachments, strExtra As String
    Dim iFeet, iInch, iDiff As Integer
    
    On Error Resume Next
    
    Me.Hide
    
    lbAll.Clear
        
    Err = 0
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    objSS1.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For n = 0 To objSS1.count - 1
    'For Each objBlock In objSS
        Set objBlock = objSS1.Item(n)
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = "POLE" Then GoTo Next_objBlock
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        
        strAttachments = ""
        
        lbAll.AddItem vAttList(0).TextString, 0
        vCoords = objBlock.InsertionPoint
        
        lbAll.List(0, 1) = "x"
        lbAll.List(0, 2) = ""
        lbAll.List(0, 3) = vCoords(0) & "," & vCoords(1)
        lbAll.List(0, 4) = n
        
        For i = 9 To 23
            If Not vAttList(i).TextString = "" Then
                vLine = Split(UCase(vAttList(i).TextString), " ")
                
                Select Case i
                    Case Is = 9
                        For j = 0 To UBound(vLine)
                            If strAttachments = "" Then
                                strAttachments = "NEUTRAL=" & vLine(j)
                            Else
                                strAttachments = strAttachments & ";;" & "NEUTRAL=" & vLine(j)
                            End If
                        Next j
                    Case Is = 10
                        For j = 0 To UBound(vLine)
                            If strAttachments = "" Then
                                strAttachments = "TRANSFORMER=" & vLine(j)
                            Else
                                strAttachments = strAttachments & ";;" & "TRANSFORMER=" & vLine(j)
                            End If
                        Next j
                    Case Is = 11
                        For j = 0 To UBound(vLine)
                            If strAttachments = "" Then
                                strAttachments = "LOW POWER=" & vLine(j)
                            Else
                                strAttachments = strAttachments & ";;" & "LOW POWER=" & vLine(j)
                            End If
                        Next j
                    Case Is = 12
                        For j = 0 To UBound(vLine)
                            If strAttachments = "" Then
                                strAttachments = "ANTENNA=" & vLine(j)
                            Else
                                strAttachments = strAttachments & ";;" & "ANTENNA=" & vLine(j)
                            End If
                        Next j
                    Case Is = 13
                        For j = 0 To UBound(vLine)
                            If strAttachments = "" Then
                                strAttachments = "ST LT CIR=" & vLine(j)
                            Else
                                strAttachments = strAttachments & ";;" & "ST LT CIR=" & vLine(j)
                            End If
                        Next j
                    Case Is = 14
                        For j = 0 To UBound(vLine)
                            If strAttachments = "" Then
                                strAttachments = "ST LT=" & vLine(j)
                            Else
                                strAttachments = strAttachments & ";;" & "ST LT=" & vLine(j)
                            End If
                        Next j
                    Case Is = 15
                        For j = 0 To UBound(vLine)
                            If strAttachments = "" Then
                                strAttachments = "NEW 6M=" & vLine(j)
                            Else
                                strAttachments = strAttachments & ";;" & "NEW 6M=" & vLine(j)
                            End If
                        Next j
                    Case Else
                        vTemp = Split(UCase(vAttList(i).TextString), "=")
                        vLine = Split(vTemp(1), " ")
                        
                        If strAttachments = "" Then
                            strAttachments = vTemp(0) & "=" & vLine(0)
                        Else
                            strAttachments = strAttachments & ";;" & vTemp(0) & "=" & vLine(0)
                        End If
                        
                        If UBound(vLine) > 0 Then
                            For j = 1 To UBound(vLine)
                                strAttachments = strAttachments & ";;" & vTemp(0) & "=" & vLine(j)
                            Next j
                        End If
                End Select
            End If
        Next i
        
        If strAttachments = "" Then
            lbAll.List(0, 1) = ""
        Else
            lbAll.List(0, 1) = strAttachments
        End If
        
Next_objBlock:
    'Next objBlock
    Next n
    
    'objSS.Clear
    
    grpCode(0) = 2
    grpValue(0) = "pole_attach"
    filterType = grpCode
    filterValue = grpValue
    
    objSS2.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For n = 0 To objSS2.count - 1
    'For Each objBlock In objSS
        Set objBlock = objSS2.Item(n)
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = "" Then GoTo Next_Callout
        
        strExtra = ""
        strPole = vAttList(0).TextString
        strCompany = vAttList(2).TextString
        If InStr(strCompany, " C-WIRE") > 0 Then
            strExtra = "C"
            strCompany = Replace(strCompany, " C-WIRE", "")
        End If
        
        If InStr(strCompany, " DROP") > 0 Then
            If strExtra = "" Then
                strExtra = "D"
            Else
                strExtra = strExtra & "D"
            End If
            strCompany = Replace(strCompany, " DROP", "")
        End If
        
        If InStr(strCompany, " OHG") > 0 Then
            If strExtra = "" Then
                strExtra = "O"
            Else
                strExtra = strExtra & "O"
            End If
            strCompany = Replace(strCompany, " OHG", "")
        End If
        
        If InStr(strCompany, "LASH TO ") > 0 Then
            If strExtra = "" Then
                strExtra = "V"
            Else
                strExtra = strExtra & "V"
            End If
            strCompany = Replace(strCompany, "LASH TO ", "")
        End If
        
        strExist = Replace(vAttList(3).TextString, "'", "-")
        strExist = Replace(strExist, """", "")
        
        If Not vAttList(4).TextString = "" Then
            strTemp = UCase(vAttList(4).TextString)
            
            Select Case Left(strTemp, 1)
                Case "A"
                    strExist = strExist & "X"
                Case "F"
                    strCompany = "NEW 6M"
                    strExist = strExist & "F"
                Case "L", "R"
                    vLine = Split(strExist, "-")
                    iProp = CInt(vLine(0)) * 12
                    If UBound(vLine) > 0 Then iProp = iProp + CInt(vLine(1))
                    
                    vLine = Split(strTemp, " ")
                    strAttach = Replace(vLine(1), """", "")
                    iDiff = CInt(strAttach)
                    
                    If Left(strTemp, 1) = "L" Then
                        iExist = iProp + iDiff
                    Else
                        iExist = iProp - iDiff
                    End If
                    
                    iFeet = Int(iExist / 12)
                    iInch = iExist - (iFeet * 12)
                    
                    If iInch < 0 Then
                        While iInch < 0
                            iInch = iInch + 12
                            iFeet = iFeet - 1
                        Wend
                    End If
                    
                    If iInch > 11 Then
                        While iInch > 11
                            iInch = iInch - 12
                            iFeet = iFeet + 1
                        Wend
                    End If
                    
                    strExist = "(" & iFeet & "-" & iInch & ")" & strExist
                Case "M"
                    strCompany = "NEW 6M"
                    strExist = strExist & "T"
                Case "N"
                    strCompany = "NEW 6M"
                    strExist = strExist
                Case "T"
                    strExist = "(" & strExist & ")" & strExist
            End Select
        End If
        
        strAttachments = strCompany & "=" & strExist & strExtra
        
        If lbAll.ListCount > 0 Then
            For i = 0 To lbAll.ListCount - 1
                If lbAll.List(i, 0) = strPole Then
                    lbAll.List(i, 4) = lbAll.List(i, 4) & ";;" & n
                    
                    If lbAll.List(i, 2) = "" Then
                        lbAll.List(i, 2) = strAttachments
                    Else
                        lbAll.List(i, 2) = strAttachments & ";;" & lbAll.List(i, 2)
                    End If
                End If
            Next i
        End If
        
Next_Callout:
    'Next objBlock
    Next n
    
Exit_Sub:
    'objSS.Clear
    'objSS.Delete
    
    tbListCount.Value = lbAll.ListCount
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub lbAll_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vLine As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    vLine = Split(lbAll.List(lbAll.ListIndex, 3), ",")
    
    viewCoordsB(0) = CDbl(vLine(0)) - 300
    viewCoordsB(1) = CDbl(vLine(1)) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = viewCoordsB(0) + 600
    viewCoordsE(1) = viewCoordsB(1) + 600
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Dim objBlock As AcadBlockReference
    Dim vAtt As Variant
    Dim strLine As String
    'Dim iIndex As Integer
    
    Load VerifyACAtributes
        VerifyACAtributes.Caption = lbAll.List(lbAll.ListIndex, 0) & " " & VerifyACAtributes.Caption
        vLine = Split(lbAll.List(lbAll.ListIndex, 4), ";;")
        
        Set objBlock = objSS1.Item(CInt(vLine(0)))
        vAtt = objBlock.GetAttributes
        
        For i = 9 To 23
            VerifyACAtributes.lbPole.AddItem vAtt(i).TagString
            VerifyACAtributes.lbPole.List(i - 9, 1) = vAtt(i).TextString
            VerifyACAtributes.lbPole.List(i - 9, 2) = ""
        Next i
        
        If UBound(vLine) > 0 Then
            For i = 1 To UBound(vLine)
                Set objBlock = objSS2.Item(CInt(vLine(i)))
                vAtt = objBlock.GetAttributes
                
                VerifyACAtributes.lbCallout.AddItem vAtt(2).TextString, 0
                'iIndex = VerifyACAtributes.lbCallout.ListCount - 1
                VerifyACAtributes.lbCallout.List(0, 1) = vAtt(3).TextString
                VerifyACAtributes.lbCallout.List(0, 2) = vAtt(4).TextString
                VerifyACAtributes.lbCallout.List(0, 3) = vLine(i)
                VerifyACAtributes.lbCallout.List(0, 4) = ""
            Next i
        End If
        
        VerifyACAtributes.show
        
        If VerifyACAtributes.lbPole.ListCount = 0 Then GoTo No_Updates
        
        Set objBlock = objSS1.Item(CInt(vLine(0)))
        vAtt = objBlock.GetAttributes
            
        For i = 0 To VerifyACAtributes.lbPole.ListCount - 1
            If VerifyACAtributes.lbPole.List(i, 2) = "Y" Then
                vAtt(i + 9).TextString = VerifyACAtributes.lbPole.List(i, 1)
            End If
        Next i
            
        objBlock.Update
        
        For i = 0 To VerifyACAtributes.lbCallout.ListCount - 1
            If VerifyACAtributes.lbCallout.List(i, 4) = "Y" Then
                Set objBlock = objSS2.Item(CInt(VerifyACAtributes.lbCallout.List(i, 3)))
                vAtt = objBlock.GetAttributes
                
                vAtt(2).TextString = VerifyACAtributes.lbCallout.List(i, 0)
                vAtt(3).TextString = VerifyACAtributes.lbCallout.List(i, 1)
                vAtt(4).TextString = VerifyACAtributes.lbCallout.List(i, 2)
                
                objBlock.Update
            End If
        Next i
        
No_Updates:
    Unload VerifyACAtributes
    
    Me.show
End Sub

Private Sub lbAll_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            MsgBox lbAll.List(lbAll.ListIndex, 4)
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbAll.ColumnCount = 5
    lbAll.ColumnWidths = "120;192;180;6;6"
    
    On Error Resume Next
    
    Err = 0
    Set objSS1 = ThisDrawing.SelectionSets.Add("objSS1")
    If Not Err = 0 Then
        Set objSS1 = ThisDrawing.SelectionSets.Item("objSS1")
        objSS1.Clear
    End If
    
    Err = 0
    Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
    If Not Err = 0 Then
        Set objSS2 = ThisDrawing.SelectionSets.Item("objSS2")
        objSS2.Clear
    End If
End Sub

Private Sub UserForm_Terminate()
    objSS1.Clear
    objSS1.Delete
    
    objSS2.Clear
    objSS2.Delete
End Sub
