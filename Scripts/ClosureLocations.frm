VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClosureLocations 
   Caption         =   "Closure Locations"
   ClientHeight    =   7189
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "ClosureLocations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClosureLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAddPoles_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vLine, vItem As Variant
    Dim vReturnPnt As Variant
    Dim iIndex As Integer
    Dim strClosure, strHO1, strCounts As String
    
    On Error Resume Next
    
    Me.Hide
    
Get_Unit:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Get Closure Unit:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    
    If Not objBlock.Name = "pole_unit" Then GoTo Exit_Sub
    
    vAttList = objBlock.GetAttributes
                    
    If InStr(vAttList(3).TextString, "HACO") > 0 Then GoTo Found_Closure
    If Not InStr(vAttList(3).TextString, "HBFO") > 0 Then GoTo Exit_Sub
                    
Found_Closure:
    strClosure = Replace(vAttList(3).TextString, "+", "")
    vLine = Split(strClosure, "=")
    strClosure = vLine(0)
    vItem = Split(vLine(1), "  ")
    strClosure = strClosure & "  " & vItem(1)
    If InStr(vAttList(3).TextString, "+WH") > 0 Then strClosure = "Existing"
    
    If lbClosures.ListCount < 1 Then
        lbClosures.AddItem vAttList(0).TextString
        lbClosures.List(0, 1) = strClosure
        
        iIndex = 0
        
        GoTo Get_Closure
    Else
        'For i = 0 To lbClosures.ListCount - 1
            'If lbClosures.List(i, 0) = vAttList(0).TextString Then
                'lbClosures.List(i, 1) = strClosure
                
                'iIndex = i
                
                'GoTo Get_Closure
            'End If
        'Next i
        
        lbClosures.AddItem vAttList(0).TextString
        iIndex = lbClosures.ListCount - 1
        lbClosures.List(iIndex, 1) = strClosure
    End If
    
Get_Closure:
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Get Closure Callout:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    
    If Not objBlock.Name = "terminal" Then GoTo Exit_Sub
    
    vAttList = objBlock.GetAttributes
    
    strCounts = Replace(vAttList(0).TextString, " ", "")
    strCounts = Replace(strCounts, "\P", " + ")
    'strCounts = Replace(strCounts, vbLf, "")
    
    vLine = Split(vAttList(1).TextString, "=")
    strHO1 = vLine(1)
    
    lbClosures.List(iIndex, 2) = strHO1
    lbClosures.List(iIndex, 3) = strCounts
    
    GoTo Get_Unit
    
Exit_Sub:
    Call GetTotals
    
    Me.show
End Sub

Private Sub cbExport_Click()
    If lbClosures.ListCount < 1 Then Exit Sub
    
    Dim strPath, strName, strFile As String
    Dim vTemp As Variant
    
    strPath = ThisDrawing.Path
    strName = ThisDrawing.Name
    vTemp = Split(strName, " ")
    
    strFile = strPath & "\" & vTemp(0) & " Closure Locations.csv"
    
    Open strFile For Output As #1
    
    Print #1, "Location,Closure Type,HO1,Counts Spliced"
    
    For i = 0 To lbClosures.ListCount - 1
        Print #1, lbClosures.List(i, 0) & "," & lbClosures.List(i, 1) & "," & lbClosures.List(i, 2) & "," & lbClosures.List(i, 3)
    Next i
    
    Close #1
End Sub

Private Sub cbGetPoles_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim filterType, filterValue As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim iAtt As Integer
    Dim vUnits, vItem, vCounts As Variant
    Dim strSpliced, strClosure As String
    Dim strCoords As String
    Dim iHO1 As Integer
    
    
    'Dim vTemp As Variant
    'Dim str1, str2, strUnit As String
    'Dim iRes, iBus As Integer
    'Dim iAF, iBF, iUF, iUNK As Integer
    'Dim strExchange As String

    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH"
    filterType = grpCode
    filterValue = grpValue
    
    lbClosures.Clear
    
    iHO1 = 0
    
  On Error Resume Next
    Me.Hide
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    dPnt1(0) = vPnt1(0)
    dPnt1(1) = vPnt1(1)
    dPnt1(2) = vPnt1(2)
    
    dPnt2(0) = vPnt2(0)
    dPnt2(1) = vPnt2(1)
    dPnt2(2) = vPnt2(2)
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        iHO1 = 0
        
        vAttList = objBlock.GetAttributes
        
        If objBlock.Name = "sPole" Then
            iAtt = 27
        Else
            iAtt = 7
        End If
        
        If vAttList(iAtt).TextString = "" Then GoTo Next_Object
        If Len(vAttList(iAtt).TextString) < 6 Then GoTo Next_Object
        
        vUnits = Split(vAttList(iAtt).TextString, ";;")
        
        For i = 0 To UBound(vUnits)
            Select Case Left(vUnits(i), 5)
                Case "+HACO", "+HBFO", "+WHAC", "+WHBF"
                    vItem = Split(vUnits(i), "=")
                    strClosure = Replace(vItem(0), "+", "")
                    
                    strCoords = objBlock.InsertionPoint(0) & "," & objBlock.InsertionPoint(1)
                    
                    strSpliced = Replace(vAttList(iAtt - 1).TextString, "[A1] ", "")
                    
                    If strSpliced = "" Then
                        iHO1 = 0
                    Else
                        vItem = Split(strSpliced, " + ")
                        
                        For j = 0 To UBound(vItem)
                            If Not vItem(j) = "" Then
                                vTemp = Split(vItem(j), ": ")
                                vCounts = Split(vTemp(1), "-")
                            
                                If UBound(vCounts) = 0 Then
                                    iHO1 = iHO1 + 1
                                Else
                                    iHO1 = iHO1 + CInt(Replace(vCounts(1), ":", "")) - CInt(vCounts(0)) + 1
                                End If
                            End If
                        Next j
                    End If
                    
                    lbClosures.AddItem vAttList(0).TextString
                    lbClosures.List(lbClosures.ListCount - 1, 1) = strClosure
                    lbClosures.List(lbClosures.ListCount - 1, 2) = iHO1
                    lbClosures.List(lbClosures.ListCount - 1, 3) = strSpliced
                    lbClosures.List(lbClosures.ListCount - 1, 4) = strCoords
                    
                    GoTo Next_Object
            End Select
        Next i
        
Next_Object:
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Call GetTotals
    
    Me.show
End Sub

Private Sub cbImport_Click()
    Dim strPath, strName, strFile As String
    Dim strLine As String
    Dim vTemp As Variant
    Dim iIndex As Integer
    
    strPath = ThisDrawing.Path
    strName = ThisDrawing.Name
    vTemp = Split(strName, " ")
    
    strFile = strPath & "\" & vTemp(0) & " Closure Locations.csv"
    
    On Error Resume Next
    
    Open strFile For Input As #1
    If Not Err = 0 Then
        MsgBox "No file found"
        Exit Sub
    End If
    
    lbClosures.Clear
    Line Input #1, strLine
    
    While Not EOF(1)
        Line Input #1, strLine
        
        vTemp = Split(strLine, ",")
        
        lbClosures.AddItem vTemp(0)
        iIndex = lbClosures.ListCount - 1
        lbClosures.List(iIndex, 1) = vTemp(1)
        lbClosures.List(iIndex, 2) = vTemp(2)
        lbClosures.List(iIndex, 3) = vTemp(3)
    Wend
    
    Close #1
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub lbClosures_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vCoords As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    vCoords = Split(lbClosures.List(lbClosures.ListIndex, 4), ",")
    
    viewCoordsB(0) = CDbl(vCoords(0)) - 300
    viewCoordsB(1) = CDbl(vCoords(1)) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = CDbl(vCoords(0)) + 300
    viewCoordsE(1) = CDbl(vCoords(1)) + 300
    viewCoordsE(2) = 0#
    
    Me.Hide
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Me.show
End Sub

Private Sub lbClosures_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbClosures.ListIndex < 0 Then Exit Sub
    
    Dim iIndex As Integer
    
    iIndex = lbClosures.ListIndex
    
    Select Case KeyCode
        Case vbKeyDelete
            lbClosures.RemoveItem iIndex
            
            iIndex = iIndex - 1
            If iIndex < 0 Then iIndex = 0
            
            lbClosures.ListIndex = iIndex
    End Select
    
    Call GetTotals
End Sub

Private Sub UserForm_Initialize()
    lbClosures.ColumnCount = 5
    lbClosures.ColumnWidths = "120;108;36;190;10"
End Sub

Private Sub GetTotals()
    If lbClosures.ListCount < 1 Then Exit Sub
    
    Dim iCount, iNew, iHO1 As Integer
    
    iCount = lbClosures.ListCount
    iHO1 = 0
    iNew = 0
    
    For i = 0 To iCount - 1
        If Not lbClosures.List(i, 1) = "Existing" Then iNew = iNew + 1
        
        If lbClosures.List(i, 2) = "" Then GoTo Next_I
        
        iHO1 = iHO1 + CInt(lbClosures.List(i, 2))
Next_I:
    Next i
    
    tbTotalCount = iCount
    tbTotalNew = iNew
    tbTotalHO1 = iHO1
End Sub
