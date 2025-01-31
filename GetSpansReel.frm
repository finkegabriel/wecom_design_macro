VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetSpansReel 
   Caption         =   "Get Spans for Reel Sheet"
   ClientHeight    =   9975.001
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   5160
   OleObjectBlob   =   "GetSpansReel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GetSpansReel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGetData_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    Dim vAttList, vUnits, vItem As Variant
    Dim vTemp, vReturnPnt, vCoords As Variant
    Dim filterType, filterValue As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim strLine As String
    Dim iTemp, iCounter As Integer
    Dim dCoords() As Double

    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH,cable_span,Map coil"
    filterType = grpCode
    filterValue = grpValue
    
    lbUnits.Clear
    lbSpans.Clear
    
  On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
  
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
    
    For Each objEntity In objSS
        If Not TypeOf objEntity Is AcadBlockReference Then GoTo Next_Object
            Set objBlock = objEntity
            
            dCoords = objBlock.InsertionPoint
            
            Select Case objBlock.Name
                Case "sPole"
                    vAttList = objBlock.GetAttributes
                    
                    'If vAttList(0).TextString = "xx" Then GoTo Next_Object
                    If vAttList(0).TextString = "" Then GoTo Next_Object
                    If vAttList(0).TextString = "POLE" Then GoTo Next_Object
                    If vAttList(27).TextString = "" Then GoTo Next_Object
                    
                    vUnits = Split(vAttList(27).TextString, ";;")
                    
                    For i = 0 To UBound(vUnits)
                        If Left(vUnits(i), 4) = "+CO(" Then
                            vItem = Split(vUnits(i), "=")
                            vTemp = Split(vItem(0), ")")
                            'If UCase(Right(vItem(0), 1)) = "E" Then
                                'lbUnits.AddItem Replace(vItem(0), "+", "")
                            'Else
                                lbUnits.AddItem Replace(vTemp(0), "+", "") & ")"
                            'End If
                            
                            strLine = Replace(vItem(1), "'", "")
                            strLine = Replace(strLine, "LOOP", "")
                            strLine = Replace(strLine, " ", "")
                                
                            lbUnits.List(lbUnits.ListCount - 1, 1) = strLine
                            lbUnits.List(lbUnits.ListCount - 1, 2) = dCoords(0) & "," & dCoords(1)
                        End If
                    Next i
                Case "sPed", "sHH"
                    vAttList = objBlock.GetAttributes
                    
                    If vAttList(0).TextString = "xx" Then GoTo Next_Object
                    If vAttList(3).TextString = "" Then GoTo Next_Object
                    If vAttList(3).TextString = "POLE" Then GoTo Next_Object
                    If vAttList(7).TextString = "" Then GoTo Next_Object
                    
                    vUnits = Split(vAttList(7).TextString, ";;")
                    
                    For i = 0 To UBound(vUnits)
                        Select Case Left(vUnits(i), 4)
                            Case "+BFO", "+UO("
                                vItem = Split(vUnits(i), "=")
                                vTemp = Split(vItem(0), ")")
                                
                                lbUnits.AddItem Replace(vTemp(0), "+", "") & ")"
                                
                                strLine = Replace(vItem(1), "'", "")
                                strLine = Replace(strLine, "LOOP", "")
                                strLine = Replace(strLine, " ", "")
                                
                                lbUnits.List(lbUnits.ListCount - 1, 1) = strLine
                                lbUnits.List(lbUnits.ListCount - 1, 2) = dCoords(0) & "," & dCoords(1)
                        End Select
                    Next i
                Case "cable_span"
                    If InStr(LCase(objBlock.Layer), "existing") > 0 Then GoTo Next_Object
                    vAttList = objBlock.GetAttributes
                    
                    If InStr(vAttList(2).TextString, "=") > 0 Then GoTo Next_Object
                    
                    If objBlock.Layer = "Integrity Proposed-Buried" Then
                        strLine = cbBuried.Value & "("
                    Else
                        strLine = "CO("
                    End If
                    
                    If Right(vAttList(1).TextString, 1) = " " Then
                        vAttList(1).TextString = Left(vAttList(1).TextString, Len(vAttList(1).TextString) - 1)
                    End If
                    
                    If vAttList(1).TextString = "" Then
                        strLine = strLine & "?)"
                        
                        lbSpans.AddItem strLine
                        lbSpans.List(lbSpans.ListCount - 1, 1) = Replace(vAttList(2).TextString, "'", "")
                        lbSpans.List(lbSpans.ListCount - 1, 2) = dCoords(0) & "," & dCoords(1)
                    Else
                        If InStr(vAttList(1).TextString, " ") > 0 Then
                            vTemp = Split(vAttList(1).TextString, " ")
                            
                            For i = 0 To UBound(vTemp)
                                If InStr(vTemp(i), "(") > 0 Then
                                    lbSpans.AddItem Replace(UCase(vTemp(i)), "E", "")
                                Else
                                    lbSpans.AddItem strLine & Replace(UCase(vTemp(i)), "F", "") & ")"
                                End If
                                lbSpans.List(lbSpans.ListCount - 1, 1) = Replace(vAttList(2).TextString, "'", "")
                                lbSpans.List(lbSpans.ListCount - 1, 2) = dCoords(0) & "," & dCoords(1)
                            Next i
                        Else
                            'strLine = strLine & Replace(UCase(vAttList(2).TextString), "F", "") & ")"
                            If InStr(vAttList(1).TextString, "(") > 0 Then
                                lbSpans.AddItem Replace(UCase(vAttList(1).TextString), "E", "")
                            Else
                                lbSpans.AddItem strLine & Replace(UCase(vAttList(1).TextString), "F", "") & ")"
                            End If
                            lbSpans.List(lbSpans.ListCount - 1, 1) = Replace(vAttList(2).TextString, "'", "")
                            lbSpans.List(lbSpans.ListCount - 1, 2) = dCoords(0) & "," & dCoords(1)
                        End If
                    End If
                Case "Map coil"
                    If InStr(LCase(objBlock.Layer), "existing") > 0 Then GoTo Next_Object
                    vAttList = objBlock.GetAttributes
                    
                    If InStr(LCase(objBlock.Layer), "buried") > 0 Then
                        lbSpans.AddItem cbBuried.Value & "(" & Replace(UCase(vAttList(1).TextString), "F", "") & ")"
                    Else
                        lbSpans.AddItem "CO(" & Replace(UCase(vAttList(1).TextString), "F", "") & ")"
                    End If
                    lbSpans.List(lbSpans.ListCount - 1, 1) = Replace(vAttList(0).TextString, "'", "")
                    lbSpans.List(lbSpans.ListCount - 1, 2) = dCoords(0) & "," & dCoords(1)
            End Select
            
                    
Next_Object:
    Next objEntity
    
Exit_Sub:
    Call GetTotals
    
    objSS.Clear
    objSS.Delete
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRemoveSame_Click()
    If lbUnits.ListCount < 1 Then Exit Sub
    If lbSpans.ListCount < 1 Then Exit Sub
    
    For i = lbUnits.ListCount - 1 To 0 Step -1
        For j = lbSpans.ListCount - 1 To 0 Step -1
            If lbUnits.List(i, 0) = lbSpans.List(j, 0) Then
                If lbUnits.List(i, 1) = lbSpans.List(j, 1) Then
                    lbUnits.RemoveItem i
                    lbSpans.RemoveItem j
                
                    GoTo Next_I
                End If
            End If
        Next j
Next_I:
    Next i
    
    Call GetTotals
End Sub

Private Sub cbSave_Click()
    If lbSpans.ListCount = 0 Then Exit Sub
    
    Dim strFileName As String
    Dim strDWGName As String
    Dim vLine As Variant
    
    'Dim objData As New DataObject
    Dim strCopy As String
    Dim vUnit As Variant
    
    strCopy = ""
    
    For i = 0 To lbUnits.ListCount - 1
        If i = 0 Then
            strCopy = lbUnits.List(i, 0) & "," & lbUnits.List(i, 1)
        Else
            strCopy = strCopy & vbCr & lbUnits.List(i, 0) & "," & lbUnits.List(i, 1)
        End If
    Next i
    
    'If cbCopyToFile.Value = True Then
        strFileName = ThisDrawing.Path & "\"
        vLine = Split(ThisDrawing.Name, " ")
        strFileName = strFileName & vLine(0) & "-Spans and Coils.csv"

        Open strFileName For Output As #1
        
        Print #1, "UNIT,LENGTH"
        Print #1, strCopy
        Close #1
    
        MsgBox "Copied to File"
    'Else
        'objData.SetText strCopy
        'objData.PutInClipboard
    
        'MsgBox "Copied to Clipboard"
    'End If
End Sub

Private Sub lbSpans_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vCoords, vAttList As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    vCoords = Split(lbSpans.List(lbSpans.ListIndex, 2), ",")
    
    viewCoordsB(0) = vCoords(0) - 200
    viewCoordsB(1) = vCoords(1) - 200
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 200
    viewCoordsE(1) = vCoords(1) + 200
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Me.show
End Sub

Private Sub lbUnits_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vCoords, vAttList As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    vCoords = Split(lbUnits.List(lbUnits.ListIndex, 2), ",")
    
    viewCoordsB(0) = vCoords(0) - 200
    viewCoordsB(1) = vCoords(1) - 200
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 200
    viewCoordsE(1) = vCoords(1) + 200
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbUnits.ColumnCount = 3
    lbUnits.ColumnWidths = "72;30;8"
    
    lbSpans.ColumnCount = 3
    lbSpans.ColumnWidths = "72;30;8"
    
    cbType.AddItem "Window"
    cbType.AddItem "Polygon"
    cbType.Value = "Window"
    
    cbBuried.AddItem "BFO"
    cbBuried.AddItem "UO"
    cbBuried.Value = "UO"
End Sub

Private Sub GetTotals()
    Dim lTotal As Long
    
    lTotal = 0
    
    tbUnitCount.Value = lbUnits.ListCount
    tbSpanCount.Value = lbSpans.ListCount
    
    If Not lbUnits.ListCount < 1 Then
        For i = 0 To lbUnits.ListCount - 1
            lTotal = lTotal + CInt(lbUnits.List(i, 1))
        Next i
        
        tbUnits.Value = lTotal
    Else
        tbUnits.Value = "0"
    End If
    
    lTotal = 0
    
    If Not lbSpans.ListCount < 1 Then
        For i = 0 To lbSpans.ListCount - 1
            lTotal = lTotal + CInt(lbSpans.List(i, 1))
        Next i
        
        tbSpans.Value = lTotal
    Else
        tbSpans.Value = "0"
    End If
    
    tbDiffCount.Value = lbUnits.ListCount - lbSpans.ListCount
    tbDiffTotal.Value = CLng(tbUnits.Value) - CLng(tbSpans.Value)
End Sub
