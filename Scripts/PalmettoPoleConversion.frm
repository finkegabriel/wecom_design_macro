VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PalmettoPoleConversion 
   Caption         =   "Palmetto Pole Conversion"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8430.001
   OleObjectBlob   =   "PalmettoPoleConversion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PalmettoPoleConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbConvertWindow_Click()
    Dim vPnt1, vPnt2 As Variant
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objEntity As AcadEntity
    Dim objMText As AcadMText
    
    'Dim strLine, strTemp As String
    'Dim strAttach As String
    'Dim vLine, vAttach As Variant
    
    On Error Resume Next
    
    Me.Hide
    
    'lbAttachments.Clear
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    grpCode(0) = 0
    grpValue(0) = "MTEXT"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For Each objMText In objSS
        Call PlaceBlock(objMText)
    Next objMText
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbCreatePoint_Click()
    Dim objEntity As AcadEntity
    Dim objMText As AcadMText
    Dim vReturnPnt As Variant
    
    Me.Hide
    
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt
    If TypeOf objEntity Is AcadMText Then
        Set objMText = objEntity
    Else
        MsgBox "Not a valid object."
        Me.show
        Exit Sub
    End If
    
    Call PlaceBlock(objMText)
    
    Me.show
End Sub

Private Sub cbGetAttachments_Click()
    Dim vPnt1, vPnt2 As Variant
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objEntity As AcadEntity
    Dim objMText As AcadMText
    
    Dim strLine, strTemp As String
    Dim strAttach As String
    Dim vLine, vAttach As Variant
    
    On Error Resume Next
    
    Me.Hide
    
    lbAttachments.Clear
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    grpCode(0) = 0
    grpValue(0) = "MTEXT"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For Each objMText In objSS
        strLine = UCase(objMText.TextString)
        vLine = Split(strLine, "ATTACHMENT HEIGHTS")
        If UBound(vLine) < 1 Then GoTo Next_ObjMText
        
        strAttach = vLine(1)
        
        strLine = Replace(strAttach, "{", "")
        strLine = Replace(strLine, "}", "")
        strLine = Replace(strLine, "\L", "")
        strLine = Replace(strLine, "\C7;", "")
        strLine = Replace(strLine, "\C6;", "")
        strLine = Replace(strLine, "\C1;", "")
        strLine = Replace(strLine, "\C0;", "")
        strLine = Replace(strLine, "\I;", "")
        
        vAttach = Split(strLine, "\P")
        
        For i = 0 To UBound(vAttach)
            vLine = Split(vAttach(i), "=")
            
            If lbAttachments.ListCount < 1 Then
                lbAttachments.AddItem vLine(0)
                lbAttachments.List(lbAttachments.ListCount - 1, 1) = ""
            Else
                For j = 0 To lbAttachments.ListCount - 1
                    If lbAttachments.List(j, 0) = vLine(0) Then GoTo Next_Attachment
                Next j
                
                lbAttachments.AddItem vLine(0)
                lbAttachments.List(lbAttachments.ListCount - 1, 1) = ""
            End If
Next_Attachment:
        Next i
        
Next_ObjMText:
    Next objMText
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbToAttribute_Click()
    If lbAttachments.ListIndex < 0 Then Exit Sub
    
    lbAttachments.List(lbAttachments.ListIndex, 1) = cbAttribute.Value
End Sub

Private Sub UserForm_Initialize()
    lbAttachments.ColumnCount = 2
    lbAttachments.ColumnWidths = "216;108"
    
    cbAttribute.AddItem ""
    cbAttribute.AddItem "9 NEUTRAL"
    cbAttribute.AddItem "10 TRANSFORMER"
    cbAttribute.AddItem "11 LOW POWER"
    cbAttribute.AddItem "12 ANTENNA"
    cbAttribute.AddItem "13 ST LT CIRCUIT"
    cbAttribute.AddItem "14 ST LT"
    cbAttribute.AddItem "15 NEW"
    cbAttribute.AddItem "15 OHG"
    cbAttribute.AddItem "15 FUTURE"
    cbAttribute.AddItem " COMM"
    cbAttribute.AddItem " COMM DROP"
    cbAttribute.AddItem " COMM C-WIRE"
    cbAttribute.AddItem " COMM OHG"
    cbAttribute.AddItem " COMM SS"
End Sub

Private Sub PlaceBlock(objMText As AcadMText)
    Dim vNE, vAttList As Variant
    Dim objBlock As AcadBlockReference
    Dim vLine, vItem, vTemp As Variant
    Dim strLine, strText As String
    
    Dim dCoords(2) As Double
    Dim strAtt(27) As String
    Dim strCompany, strAttach As String
    Dim iCOMM As Integer
    
    iCOMM = 16
    
    For i = 0 To 27
        strAtt(i) = ""
    Next i
    
    On Error Resume Next
    
    strLine = UCase(objMText.TextString)
    If InStr(strLine, "ATTACHMENT HEIGHTS") < 1 Then Exit Sub
                
    strLine = Replace(strLine, "{", "")
    strLine = Replace(strLine, "}", "")
    strLine = Replace(strLine, "\L", "")
    strLine = Replace(strLine, "\C7;", "")
    strLine = Replace(strLine, "\C6;", "")
    strLine = Replace(strLine, "\C1;", "")
    strLine = Replace(strLine, "\C0;", "")
    strLine = Replace(strLine, "\I;", "")
    
    vLine = Split(strLine, "\P")
    
    strAtt(0) = vLine(0)
    strAtt(2) = "MTE"
    
    vItem = Split(vLine(1), ":")
    If UBound(vItem) > 0 Then
        strAtt(3) = "FID: " & vItem(1)
    Else
        strAtt(3) = "NA"
    End If
    
    vItem = Split(vLine(2), ":")
    vItem(1) = Replace(vItem(1), " ", "")
    vItem(1) = Replace(vItem(1), vbTab, "")
    dCoords(0) = CDbl(vItem(1))
    dCoords(1) = CDbl(vLine(3))
    strAtt(7) = vItem(1) & "," & vLine(3)
    
    vNE = LLtoTN83F(CDbl(dCoords(0)), CDbl(dCoords(1)))
    dCoords(0) = CDbl(vNE(1))
    dCoords(1) = CDbl(vNE(0))
    dCoords(2) = 0#
    'strText = vNE(1) & "," & vNE(0) & vbCr & strNumber & vbCr & "FID: " & strFID
    'strText = strText & vbCr & vItem(1) & "," & vLine(3)
    
    ''MsgBox strAtt(0) & vbCr & strAtt(3) & vbCr & strAtt(7)
    
    If UBound(vLine) > 4 Then
        For i = 5 To UBound(vLine)
            vItem = Split(vLine(i), "=")
            strCompany = vItem(0)
            'vTemp = Split(vItem(0), " ")
            'strCompany = vTemp(0)
            
            vItem(1) = Replace(vItem(1), "'", "-")
            vItem(1) = Replace(vItem(1), """", "")
            vTemp = Split(vItem(1), " MOVE TO ")
            
            If UBound(vTemp) = 0 Then
                strAttach = vTemp(0)
            Else
                strAttach = "(" & vTemp(0) & ")" & vTemp(1)
            End If
            
            If lbAttachments.ListCount > 0 Then
                For j = 0 To lbAttachments.ListCount - 1
                    If lbAttachments.List(j, 0) = strCompany Then
                        Select Case lbAttachments.List(j, 1)
                            Case ""
                                If iCOMM > 16 Then
                                    For k = 16 To 23
                                        If strAtt(k) = "" Then GoTo Not_Found
                                        vTemp = Split(strAtt(k), "=")
                                        If vTemp(0) = strCompany Then
                                            strAtt(k) = strAtt(k) & " " & strAttach
                                            GoTo Next_J
                                        End If
                                    Next k
                                End If
Not_Found:
                                strAtt(iCOMM) = strCompany & "=" & strAttach
                                iCOMM = iCOMM + 1
                            Case "9 NEUTRAL"
                                If strAtt(9) = "" Then
                                    strAtt(9) = strAttach
                                Else
                                    strAtt(9) = strAtt(9) & " " & strAttach
                                End If
                            Case "10 TRANSFORMER"
                                If strAtt(10) = "" Then
                                    strAtt(10) = strAttach
                                Else
                                    strAtt(10) = strAtt(10) & " " & strAttach
                                End If
                            Case "11 LOW POWER"
                                If strAtt(11) = "" Then
                                    strAtt(11) = strAttach
                                Else
                                    strAtt(11) = strAtt(11) & " " & strAttach
                                End If
                            Case "12 ANTENNA"
                                If strAtt(12) = "" Then
                                    strAtt(12) = strAttach
                                Else
                                    strAtt(12) = strAtt(12) & " " & strAttach
                                End If
                            Case "13 ST LT CIRCUIT"
                                If strAtt(13) = "" Then
                                    strAtt(13) = strAttach
                                Else
                                    strAtt(13) = strAtt(13) & " " & strAttach
                                End If
                            Case "14 ST LT"
                                If strAtt(14) = "" Then
                                    strAtt(14) = strAttach
                                Else
                                    strAtt(14) = strAtt(14) & " " & strAttach
                                End If
                            Case "15 NEW"
                                If strAtt(15) = "" Then
                                    strAtt(15) = strAttach
                                Else
                                    strAtt(15) = strAtt(15) & " " & strAttach
                                End If
                            Case "15 OHG"
                                If strAtt(15) = "" Then
                                    strAtt(15) = strAttach & "o"
                                Else
                                    strAtt(15) = strAtt(15) & " " & strAttach & "o"
                                End If
                            Case "15 FUTURE"
                                If strAtt(15) = "" Then
                                    strAtt(15) = strAttach & "f"
                                Else
                                    strAtt(15) = strAtt(15) & " " & strAttach & "f"
                                End If
                            Case Else
                                vItem = Split(strCompany, " ")
                                strCompany = vItem(0)
                                
                                If iCOMM > 16 Then
                                    For k = 16 To 23
                                        If strAtt(k) = "" Then GoTo Not_Found2
                                        
                                        vItem = Split(strAtt(k), "=")
                                        
                                        If strCompany = vItem(0) Then
                                            strAtt(k) = strAtt(k) & " " & strAttach
                                            GoTo Add_Suffix
                                        End If
                                    Next k
                                End If
Not_Found2:
                                strAtt(iCOMM) = vItem(0) & "=" & strAttach
                                k = iCOMM
                                iCOMM = iCOMM + 1
Add_Suffix:
                                Select Case lbAttachments.List(j, 1)
                                    Case " COMM DROP"
                                        strAtt(k) = strAtt(k) & "d"
                                    Case " COMM C-WIRE"
                                        strAtt(k) = strAtt(k) & "c"
                                    Case " COMM OHG"
                                        strAtt(k) = strAtt(k) & "o"
                                    Case " COMM SS"
                                        strAtt(k) = strAtt(k) & "s"
                                End Select
                                
                        End Select
                    End If
Next_J:
                Next j
            Else
                'vItem = Split(vLine(i), "=")
                
                If iCOMM > 16 Then
                    For k = 16 To 23
                        If strAtt(k) = "" Then GoTo Not_Found3
                        vTemp = Split(strAtt(k), "=")
                        If vTemp(0) = strCompany Then
                            strAtt(k) = strAtt(k) & " " & strAttach
                            GoTo Next_J
                        End If
                    Next k
                End If
Not_Found3:
                strAtt(iCOMM) = strCompany & "=" & strAttach
                iCOMM = iCOMM + 1
            End If
        Next i
        
        'strText = dCoords(0) & "," & dCoords(1) & vbCr & strAtt(0)
        'For m = 1 To 27
            'If Not strAtt(m) = "" Then strText = strText & vbCr & strAtt(m)
        'Next m
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dCoords, "sPole", 1#, 1#, 1#, 0#)
        objBlock.Layer = "Integrity Poles-Power"
        vAttList = objBlock.GetAttributes
        For i = 0 To 27
            vAttList(i).TextString = strAtt(i)
        Next i
        
        objBlock.Update
    End If
    'tbResult.Value = strText
End Sub

Private Function LLtoTN83F(dLat As Double, dLong As Double)
    Dim dDLat As Double
    Dim dEast, dDiffE, dEast0 As Double
    Dim dNorth, dDiffN, dNorth0 As Double
    Dim dU, dR, dCA, dK As Double
    Dim NE(2) As Double
    
    dDLat = dLat - 35.8340607459
    dU = dDLat * (110950.2019 + dDLat * (9.25072 + dDLat * (5.64572 + dDLat * 0.017374)))
    dR = 8842127.1422 - dU
    dCA = ((86 + dLong) * 0.585439726459) * 3.14159265359 / 180
    
    dDiffE = dR * Sin(dCA)
    dDiffN = dU + dDiffE * Tan(dCA / 2)
    
    dEast = (dDiffE + 600000) / 0.3048006096
    dNorth = (dDiffN + 166504.1691) / 0.3048006096
    
    dK = 0.999948401424 + (1.23188E-14 * dU * dU) + (4.54E-22 * dU * dU * dU)
    
    NE(0) = dNorth
    NE(1) = dEast
    NE(2) = dK
    
    LLtoTN83F = NE
End Function
