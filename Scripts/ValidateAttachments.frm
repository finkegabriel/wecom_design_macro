VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateAttachments 
   Caption         =   "Validate Attachments"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5745
   OleObjectBlob   =   "ValidateAttachments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbClearSame_Click()
    If lbPoles.ListCount < 1 Then Exit Sub
    
    For i = lbPoles.ListCount - 1 To 0 Step -1
        If lbPoles.List(i, 1) = lbPoles.List(i, 3) Then
            If lbPoles.List(i, 2) = lbPoles.List(i, 4) Then lbPoles.RemoveItem i
        End If
    Next i
    
    tbDifferent.Value = lbPoles.ListCount
End Sub

'Dim vMapLL, vMapUR As Variant

Private Sub cbGetPoles_Click()
    Dim vDwgLL, vDwgUR As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSDwg As AcadSelectionSet
    Dim objPoleDwg As AcadBlockReference
    Dim vAttDwg, vAttMap As Variant
    Dim vLine, vTemp As Variant
    Dim strPoleNum As String
    Dim strDwgLatLong, strMapLatLong As String
    Dim iPwr, iCOMM As Integer
    
    iPwr = 0
    iCOMM = 0
    
    On Error Resume Next
    
    Me.Hide
    Err = 0
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSDwg = ThisDrawing.SelectionSets.Add("objSSDwg")
    objSSDwg.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    If objSSDwg.count < 1 Then GoTo Exit_Sub
    
    For Each objPoleDwg In objSSDwg
        vAttDwg = objPoleDwg.GetAttributes
        If vAttDwg(7).TextString = "" Then GoTo Next_objDwgPole
        
        strPoleNum = vAttDwg(0).TextString
        strDwgLatLong = vAttDwg(7).TextString
        'strDwgLatLong = objPoleDwg.InsertionPoint(0) & "," & objPoleDwg.InsertionPoint(1)
        
        If strPoleNum = "POLE" Then GoTo Next_objDwgPole
        If strPoleNum = "" Then GoTo Next_objDwgPole
        
        
        If Not vAttDwg(9).TextString = "" Then
            vLine = Split(vAttDwg(9).TextString, " ")
            iPwr = iPwr + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(10).TextString = "" Then
            vLine = Split(vAttDwg(10).TextString, " ")
            iPwr = iPwr + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(11).TextString = "" Then
            vLine = Split(vAttDwg(11).TextString, " ")
            iPwr = iPwr + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(12).TextString = "" Then
            vLine = Split(vAttDwg(12).TextString, " ")
            iPwr = iPwr + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(13).TextString = "" Then
            vLine = Split(vAttDwg(13).TextString, " ")
            iPwr = iPwr + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(14).TextString = "" Then
            vLine = Split(vAttDwg(14).TextString, " ")
            iPwr = iPwr + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(16).TextString = "" Then
            vTemp = Split(vAttDwg(16).TextString, "=")
            vLine = Split(vTemp(UBound(vTemp)), " ")
            iCOMM = iCOMM + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(17).TextString = "" Then
            vTemp = Split(vAttDwg(17).TextString, "=")
            vLine = Split(vTemp(UBound(vTemp)), " ")
            iCOMM = iCOMM + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(18).TextString = "" Then
            vTemp = Split(vAttDwg(18).TextString, "=")
            vLine = Split(vTemp(UBound(vTemp)), " ")
            iCOMM = iCOMM + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(19).TextString = "" Then
            vTemp = Split(vAttDwg(19).TextString, "=")
            vLine = Split(vTemp(UBound(vTemp)), " ")
            iCOMM = iCOMM + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(20).TextString = "" Then
            vTemp = Split(vAttDwg(20).TextString, "=")
            vLine = Split(vTemp(UBound(vTemp)), " ")
            iCOMM = iCOMM + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(21).TextString = "" Then
            vTemp = Split(vAttDwg(21).TextString, "=")
            vLine = Split(vTemp(UBound(vTemp)), " ")
            iCOMM = iCOMM + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(22).TextString = "" Then
            vTemp = Split(vAttDwg(22).TextString, "=")
            vLine = Split(vTemp(UBound(vTemp)), " ")
            iCOMM = iCOMM + UBound(vLine) + 1
        End If
        
        If Not vAttDwg(23).TextString = "" Then
            vTemp = Split(vAttDwg(23).TextString, "=")
            vLine = Split(vTemp(UBound(vTemp)), " ")
            iCOMM = iCOMM + UBound(vLine) + 1
        End If
        
        lbPoles.AddItem strPoleNum
        lbPoles.List(lbPoles.ListCount - 1, 1) = iPwr
        lbPoles.List(lbPoles.ListCount - 1, 2) = iCOMM
        lbPoles.List(lbPoles.ListCount - 1, 3) = ""
        lbPoles.List(lbPoles.ListCount - 1, 4) = ""
        lbPoles.List(lbPoles.ListCount - 1, 5) = strDwgLatLong
        lbPoles.List(lbPoles.ListCount - 1, 6) = objPoleDwg.InsertionPoint(0) & "," & objPoleDwg.InsertionPoint(1)
        lbPoles.List(lbPoles.ListCount - 1, 7) = ""
        
        iPwr = 0
        iCOMM = 0
Next_objDwgPole:
    Next objPoleDwg
    
Exit_Sub:
    objSSDwg.Clear
    objSSDwg.Delete
    
    tbTotal.Value = lbPoles.ListCount
    
    If lbPoles.ListCount > 0 Then Call SortList
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    Err = 0
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub lbDWG_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbPoles.ListCount < 1 Then Exit Sub
    If Not lbPoles.List(0, 3) = "" Then Exit Sub
    If Not lbPoles.List(0, 4) = "" Then Exit Sub
    
    Dim objMainDwg As AcadDocument
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSMap As AcadSelectionSet
    Dim objPole As AcadBlockReference
    Dim vAttMap, vLine, vTemp As Variant
    Dim strFile As String
    Dim strAttach As String
    Dim iPwr, iCOMM As Integer
    
    Set objMainDwg = ThisDrawing
    
    strFile = ThisDrawing.Path & "\MISC\" & lbDWG.List(lbDWG.ListIndex) & ".dwg"
    'MsgBox strFile
    
    If Dir(strFile) = "" Then
        MsgBox "Error opening file."
        Exit Sub
    End If
    
    ThisDrawing.Application.Documents.Open strFile
    ZoomExtents
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSMap = ThisDrawing.SelectionSets.Add("objSSMap")
    objSSMap.Select acSelectionSetAll, , , filterType, filterValue
    If objSSMap.count < 1 Then
        MsgBox "None found"
        GoTo Exit_Sub
    End If
    
    For Each objPole In objSSMap
        iPwr = 0
        iCOMM = 0
        strAttach = ""
        vAttMap = objPole.GetAttributes
        If vAttMap(7).TextString = "" Then GoTo Next_objPole
        
        For i = 0 To lbPoles.ListCount - 1
            If vAttMap(7).TextString = lbPoles.List(i, 5) Then
                If Not vAttMap(9).TextString = "" Then
                    vLine = Split(vAttMap(9).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = "N = " & vAttMap(9).TextString
                End If
        
                If Not vAttMap(10).TextString = "" Then
                    vLine = Split(vAttMap(10).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "TF = " & vAttMap(10).TextString
                End If
        
                If Not vAttMap(11).TextString = "" Then
                    vLine = Split(vAttMap(11).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "LP = " & vAttMap(11).TextString
                End If
        
                If Not vAttMap(12).TextString = "" Then
                    vLine = Split(vAttMap(12).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "ANT = " & vAttMap(12).TextString
                End If
        
                If Not vAttMap(13).TextString = "" Then
                    vLine = Split(vAttMap(13).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "SLC = " & vAttMap(13).TextString
                End If
        
                If Not vAttMap(14).TextString = "" Then
                    vLine = Split(vAttMap(14).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "SL = " & vAttMap(14).TextString
                End If
        
                If Not vAttMap(15).TextString = "" Then
                    strAttach = strAttach & vbCr & "NEW = " & vAttMap(15).TextString
                End If
        
                If Not vAttMap(16).TextString = "" Then
                    vTemp = Split(vAttMap(16).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(17).TextString = "" Then
                    vTemp = Split(vAttMap(17).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(18).TextString = "" Then
                    vTemp = Split(vAttMap(18).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(19).TextString = "" Then
                    vTemp = Split(vAttMap(19).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(20).TextString = "" Then
                    vTemp = Split(vAttMap(20).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(21).TextString = "" Then
                    vTemp = Split(vAttMap(21).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(22).TextString = "" Then
                    vTemp = Split(vAttMap(22).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(23).TextString = "" Then
                    vTemp = Split(vAttMap(23).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(24).TextString = "" Then
                    strAttach = strAttach & vbCr & "OHG = " & vAttMap(24).TextString
                End If
                
                lbPoles.List(i, 3) = iPwr
                lbPoles.List(i, 4) = iCOMM
                lbPoles.List(i, 7) = strAttach
            End If
        Next i
        
Next_objPole:
    Next objPole
    
    ThisDrawing.Close False, strFile
Exit_Sub:
    'objSSMap.Clear
    'objSSMap.Delete
    
    objMainDwg.Activate
End Sub

Private Sub lbOpen_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lbPoles.ListCount < 1 Then Exit Sub
    If Not lbPoles.List(0, 3) = "" Then Exit Sub
    If Not lbPoles.List(0, 4) = "" Then Exit Sub
    
    Dim objMainDwg As AcadDocument
    Dim objDoc As AcadDocument
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSSMap As AcadSelectionSet
    Dim objPole As AcadBlockReference
    Dim vAttMap, vLine, vTemp As Variant
    Dim strFile As String
    Dim strName, strAttach As String
    Dim iPwr, iCOMM As Integer
    
    strName = ThisDrawing.Name
    
    For Each objDoc In AcadApplication.Documents
        If objDoc.Name = lbOpen.List(lbOpen.ListIndex, 0) Then GoTo Found_Doc
    Next objDoc
    
    MsgBox "Not Found" & vbCr & lbOpen.List(lbOpen.ListIndex, 0)
    
    Exit Sub
    
Found_Doc:
    
    objDoc.Activate
    
    ZoomExtents
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSSMap = ThisDrawing.SelectionSets.Add("objSSMap")
    objSSMap.Select acSelectionSetAll, , , filterType, filterValue
    If objSSMap.count < 1 Then
        MsgBox "None found"
        GoTo Exit_Sub
    End If
    
    For Each objPole In objSSMap
        iPwr = 0
        iCOMM = 0
        strAttach = ""
        vAttMap = objPole.GetAttributes
        If vAttMap(7).TextString = "" Then GoTo Next_objPole
        
        For i = 0 To lbPoles.ListCount - 1
            If vAttMap(7).TextString = lbPoles.List(i, 5) Then
                If Not vAttMap(9).TextString = "" Then
                    vLine = Split(vAttMap(9).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = "N = " & vAttMap(9).TextString
                End If
        
                If Not vAttMap(10).TextString = "" Then
                    vLine = Split(vAttMap(10).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "TF = " & vAttMap(10).TextString
                End If
        
                If Not vAttMap(11).TextString = "" Then
                    vLine = Split(vAttMap(11).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "LP = " & vAttMap(11).TextString
                End If
        
                If Not vAttMap(12).TextString = "" Then
                    vLine = Split(vAttMap(12).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "ANT = " & vAttMap(12).TextString
                End If
        
                If Not vAttMap(13).TextString = "" Then
                    vLine = Split(vAttMap(13).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "SLC = " & vAttMap(13).TextString
                End If
        
                If Not vAttMap(14).TextString = "" Then
                    vLine = Split(vAttMap(14).TextString, " ")
                    iPwr = iPwr + UBound(vLine) + 1
                    strAttach = strAttach & vbCr & "SL = " & vAttMap(14).TextString
                End If
        
                If Not vAttMap(15).TextString = "" Then
                    strAttach = strAttach & vbCr & "NEW = " & vAttMap(15).TextString
                End If
        
                If Not vAttMap(16).TextString = "" Then
                    vTemp = Split(vAttMap(16).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(17).TextString = "" Then
                    vTemp = Split(vAttMap(17).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(18).TextString = "" Then
                    vTemp = Split(vAttMap(18).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(19).TextString = "" Then
                    vTemp = Split(vAttMap(19).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(20).TextString = "" Then
                    vTemp = Split(vAttMap(20).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(21).TextString = "" Then
                    vTemp = Split(vAttMap(21).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(22).TextString = "" Then
                    vTemp = Split(vAttMap(22).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(23).TextString = "" Then
                    vTemp = Split(vAttMap(23).TextString, "=")
                    vLine = Split(vTemp(UBound(vTemp)), " ")
                    For j = 0 To UBound(vLine)
                        If Not vLine(j) = "" Then
                            iCOMM = iCOMM + 1
                            strAttach = strAttach & vbCr & vTemp(0) & "= " & vLine(j)
                        End If
                    Next j
                End If
        
                If Not vAttMap(24).TextString = "" Then
                    strAttach = strAttach & vbCr & "OHG = " & vAttMap(24).TextString
                End If
                
                lbPoles.List(i, 3) = iPwr
                lbPoles.List(i, 4) = iCOMM
                lbPoles.List(i, 7) = strAttach
            End If
        Next i
        
Next_objPole:
    Next objPole
    
    'ThisDrawing.Close , strFile
Exit_Sub:
    objSSMap.Clear
    objSSMap.Delete
    
    For Each objDoc In AcadApplication.Documents
        If objDoc.Name = strName Then GoTo Found_Doc2
    Next objDoc
    
    Exit Sub
    
Found_Doc2:
    
    objDoc.Activate
End Sub

Private Sub lbPoles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vCoords As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    vCoords = Split(lbPoles.List(lbPoles.ListIndex, 6), ",")
    'MsgBox vCoords(0) & vbCr & vCoords(1)
    
    viewCoordsB(0) = CDbl(vCoords(0)) - 300
    viewCoordsB(1) = CDbl(vCoords(1)) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = CDbl(vCoords(0)) + 300
    viewCoordsE(1) = CDbl(vCoords(1)) + 300
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Load ValidateAttachPole
        ValidateAttachPole.tbAttach.Value = lbPoles.List(lbPoles.ListIndex, 7)
        ValidateAttachPole.tbPole.Value = lbPoles.List(lbPoles.ListIndex, 0)
        ValidateAttachPole.tbCoords.Value = lbPoles.List(lbPoles.ListIndex, 6)
        
        ValidateAttachPole.show
    Unload ValidateAttachPole
    
    'MsgBox lbPoles.List(lbPoles.ListIndex, 7), , lbPoles.List(lbPoles.ListIndex, 0)
    
    Me.show
End Sub

Private Sub lbPoles_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbPoles.ListCount < 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyReturn
            Dim result, iIndex As Integer
            Dim objCircle As AcadCircle
            Dim dInsert(2) As Double
            Dim vTemp As Variant
            
            iIndex = lbPoles.ListIndex
            
            Me.Hide
                result = MsgBox(lbPoles.List(iIndex, 7), vbYesNo, "Mark " & lbPoles.List(iIndex, 0))
                If result = vbYes Then
                    vTemp = Split(lbPoles.List(iIndex, 6), ",")
                    dInsert(0) = CDbl(vTemp(0))
                    dInsert(1) = CDbl(vTemp(1))
                    dInsert(2) = 0#
                    
                    Set objCircle = ThisDrawing.ModelSpace.AddCircle(dInsert, 80)
                    objCircle.Layer = "Integrity Notes"
                    objCircle.Update
                End If
            Me.show
        Case vbKeyDelete
            lbPoles.RemoveItem lbPoles.ListIndex
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbPoles.ColumnCount = 8
    lbPoles.ColumnWidths = "120;36;36;36;24;6;6;6"
    
    lbOpen.ColumnCount = 2
    lbOpen.ColumnWidths = "264;6"
    
    Dim strFolder, strFile As String
    Dim vTemp As Variant
    
    strFolder = ""
    
    strFolder = ThisDrawing.Path & "\MISC\*.*"
    
    strFile = Dir$(strFolder)
    
    Do While strFile <> ""
        If InStr(strFile, ".dwg") Then
            lbDWG.AddItem Replace(strFile, ".dwg", "")
        End If
        strFile = Dir$
    Loop
    
    Dim objDoc As AcadDocument
    
    For Each objDoc In AcadApplication.Documents
        lbOpen.AddItem objDoc.Name
        lbOpen.List(lbOpen.ListCount - 1, 1) = objDoc.FullName
    Next objDoc
End Sub

Private Sub SortList()
    Dim strTemp, strTotal As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 6) As String
    
    iCount = lbPoles.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbPoles.List(b, 0) > lbPoles.List(b + 1, 0) Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbPoles.Selected(b) = True
                    lbPoles.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbPoles.List(b + 1, 0)
                strAtt(1) = lbPoles.List(b + 1, 1)
                strAtt(2) = lbPoles.List(b + 1, 2)
                strAtt(3) = lbPoles.List(b + 1, 3)
                strAtt(4) = lbPoles.List(b + 1, 4)
                strAtt(5) = lbPoles.List(b + 1, 5)
                strAtt(6) = lbPoles.List(b + 1, 6)
                
                lbPoles.List(b + 1, 0) = lbPoles.List(b, 0)
                lbPoles.List(b + 1, 1) = lbPoles.List(b, 1)
                lbPoles.List(b + 1, 2) = lbPoles.List(b, 2)
                lbPoles.List(b + 1, 3) = lbPoles.List(b, 3)
                lbPoles.List(b + 1, 4) = lbPoles.List(b, 4)
                lbPoles.List(b + 1, 5) = lbPoles.List(b, 5)
                lbPoles.List(b + 1, 6) = lbPoles.List(b, 6)
                
                lbPoles.List(b, 0) = strAtt(0)
                lbPoles.List(b, 1) = strAtt(1)
                lbPoles.List(b, 2) = strAtt(2)
                lbPoles.List(b, 3) = strAtt(3)
                lbPoles.List(b, 4) = strAtt(4)
                lbPoles.List(b, 5) = strAtt(5)
                lbPoles.List(b, 6) = strAtt(6)
            End If
        Next b
    Next a
End Sub
