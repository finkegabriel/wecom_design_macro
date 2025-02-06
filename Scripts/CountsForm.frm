VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CountsForm 
   Caption         =   "Cable Counts Form"
   ClientHeight    =   9735.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   OleObjectBlob   =   "CountsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CountsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objBlock As AcadBlockReference
Dim strPreviousBlock As String
Dim strCableCallout As String
Dim strSplitterFilePath As String
Dim bSave, bUpdatePole As Boolean

Private Sub cbAddClosure_Click()
    If tbCallout.Value = "" Then Exit Sub
    
    Dim objCallout As AcadBlockReference
    Dim returnPnt As Variant
    Dim vBlockCoords(0 To 2) As Double
    Dim lwpCoords(0 To 3) As Double
    Dim dPrevious(0 To 2) As Double
    Dim dOrigin(0 To 2) As Double
    Dim dScale As Double
    Dim lineObj As AcadLWPolyline
    Dim objCircle As AcadCircle
    Dim objCircle2 As AcadCircle
    Dim objText As AcadText
    Dim objText2 As AcadText
    Dim vLine, vCounts, vItem As Variant
    Dim iIndex As Integer
    
    Dim strAtt0, strAtt1, strAtt2 As String
    Dim strLayer, strCounts As String
    Dim iHO1 As Integer
    
    strAtt0 = tbActivePole.Value & ": " & tbPosition.Value
    
    iHO1 = 0
    strCounts = Replace(tbCallout.Value, vbLf, "")
    vCounts = Split(tbCallout.Value, vbCr)
    
    For i = 0 To UBound(vCounts)
        vLine = Split(vCounts(i), ": ")
        vItem = Split(vLine(1), "-")
        
        If UBound(vItem) = 0 Then
            iHO1 = iHO1 + 1
        Else
            iHO1 = iHO1 + CInt(vItem(1)) - CInt(vItem(0)) + 1
        End If
    Next i
    
    If cbCblType.Value = "CO" Then
        strAtt1 = "+HO1A=" & iHO1
        strLayer = "Integrity Proposed-Aerial"
    Else
        strAtt1 = "+HO1B=" & iHO1
        strLayer = "Integrity Proposed-Buried"
    End If
    
    strAtt2 = Replace(tbCallout.Value, vbCr, "\P")
    strAtt2 = Replace(strAtt2, vbLf, "")
    
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
    Me.Hide
    
    returnPnt = ThisDrawing.Utility.GetPoint(, "Select Point: ")
    lwpCoords(0) = returnPnt(0)
    lwpCoords(1) = returnPnt(1)
    dPrevious(0) = returnPnt(0)
    dPrevious(1) = returnPnt(1)
    dPrevious(2) = 0#
    
    returnPnt = ThisDrawing.Utility.GetPoint(dPrevious, "Select Point: ")
    vBlockCoords(0) = returnPnt(0)
    vBlockCoords(1) = returnPnt(1)
    vBlockCoords(2) = returnPnt(2)
    lwpCoords(2) = vBlockCoords(0)
    lwpCoords(3) = vBlockCoords(1)
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
    lineObj.Layer = strLayer
    lineObj.Update
    
    dScale = CDbl(cbScale.Value) / 100
    If dScale = 0 Then dScale = 0.75
        
    Set objCallout = ThisDrawing.ModelSpace.InsertBlock(vBlockCoords, "Callout", dScale, dScale, dScale, 0)
    attItem = objCallout.GetAttributes
    
    attItem(0).TextString = strAtt0
    attItem(1).TextString = strAtt1
    attItem(2).TextString = strAtt2
    
    objCallout.Layer = strLayer
    
    If lwpCoords(2) < lwpCoords(0) Then
        vBlockCoords(0) = vBlockCoords(0) - (75 * dScale)
        objCallout.InsertionPoint = vBlockCoords
    End If
    objCallout.Update
    
    Me.show
End Sub

Private Sub cbAddToPole_Click()
    Me.Hide
    Call UpdateTerminal
    
    'Call UpdateCable
    Call UpdateCableCallout
    Me.show
End Sub

Private Sub cbCallout_Click()
    If tbCblCallout.Value = "" Then Exit Sub
    
    Dim objCallout As AcadBlockReference
    Dim returnPnt As Variant
    Dim vBlockCoords(0 To 2) As Double
    Dim lwpCoords(0 To 3) As Double
    Dim dPrevious(0 To 2) As Double
    Dim dOrigin(0 To 2) As Double
    Dim dScale As Double
    Dim lineObj As AcadLWPolyline
    Dim objCircle As AcadCircle
    Dim objCircle2 As AcadCircle
    Dim objText As AcadText
    Dim objText2 As AcadText
    Dim vLine, vCounts As Variant
    Dim iIndex As Integer
    
    Dim strAtt0, strAtt1, strAtt2 As String
    Dim strLayer As String
    
    strAtt0 = tbActivePole.Value & ": " & tbPosition.Value
    strAtt1 = cbCblType.Value & "(" & cbCableSize.Value & ")"
    If Not cbSuffix.Value = "" Then strAtt1 = strAtt1 & cbSuffix.Value
    strAtt2 = Replace(tbCblCallout.Value, vbCr, "\P")
    strAtt2 = Replace(strAtt2, vbLf, "")
    
    If Left(strAtt1, 2) = "CO" Then
        strLayer = "Integrity Proposed-Aerial"
    Else
        strLayer = "Integrity Proposed-Buried"
    End If
    
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
    Me.Hide
    
    returnPnt = ThisDrawing.Utility.GetPoint(, "Select Point: ")
    lwpCoords(0) = returnPnt(0)
    lwpCoords(1) = returnPnt(1)
    dPrevious(0) = returnPnt(0)
    dPrevious(1) = returnPnt(1)
    dPrevious(2) = 0#
    
    returnPnt = ThisDrawing.Utility.GetPoint(dPrevious, "Select Point: ")
    vBlockCoords(0) = returnPnt(0)
    vBlockCoords(1) = returnPnt(1)
    vBlockCoords(2) = returnPnt(2)
    lwpCoords(2) = vBlockCoords(0)
    lwpCoords(3) = vBlockCoords(1)
    
    If cbPlacement.Value = "Leader" Then
        Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
        lineObj.Layer = strLayer
        lineObj.Update
    Else
        Set objCircle = ThisDrawing.ModelSpace.AddCircle(dPrevious, 8)
        objCircle.Layer = strLayer
        objCircle.Update
        Set objCircle2 = ThisDrawing.ModelSpace.AddCircle(returnPnt, 8)
        objCircle2.Layer = strLayer
        objCircle2.Update
        
        strLetter = UCase(ThisDrawing.Utility.GetString(0, "Enter Callout Letter:"))
        Set objText = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8)
        Set objText2 = ThisDrawing.ModelSpace.AddText(strLetter, dOrigin, 8)
        objText.Layer = strLayer
        objText.Alignment = acAlignmentMiddle
        objText.TextAlignmentPoint = dPrevious
        objText2.Layer = strLayer
        objText2.Alignment = acAlignmentMiddle
        objText2.TextAlignmentPoint = vBlockCoords
        objText.Update
        objText2.Update
        
        vBlockCoords(0) = vBlockCoords(0) + 8
        lwpCoords(0) = 0
    End If
    
    
    dScale = CDbl(cbScale.Value) / 100
    If dScale = 0 Then dScale = 0.75
        
    Set objCallout = ThisDrawing.ModelSpace.InsertBlock(vBlockCoords, "Callout", dScale, dScale, dScale, 0)
    attItem = objCallout.GetAttributes
    
    attItem(0).TextString = strAtt0
    attItem(1).TextString = strAtt1
    attItem(2).TextString = strAtt2
    
    objCallout.Layer = strLayer
    
    If lwpCoords(2) < lwpCoords(0) Then
        vBlockCoords(0) = vBlockCoords(0) - (75 * dScale)
        objCallout.InsertionPoint = vBlockCoords
    End If
    objCallout.Update
    
    Me.show
End Sub

Private Sub cbCreateMCLfile_Click()
    Me.Hide
    
    Load CreateMCL
        CreateMCL.show
    Unload CreateMCL
    
    Me.show
End Sub

Private Sub cbCutCounts_Click()
    If tbTapNeeded.Value = "" Then tbTapNeeded.Value = "1"
    If Not lbCounts.List(i, 3) = "<>" Then Exit Sub
    If tbActivePole.Value = "" Then
        MsgBox "Need to select Active Pole."
        Call GetActivePole
    End If
    
    Dim iSelected As Integer
    
    For i = 0 To lbCounts.ListCount - 1
        If lbCounts.Selected(i) = True Then
            iSelected = i
            GoTo Found_Selected
        End If
    Next i
    
    MsgBox "No Count Selected."
    Exit Sub
    
Found_Selected:
    
    If Not lbCounts.List(i, 3) = "<>" Then
        Dim result As Integer
        
        result = MsgBox("Overwrite previous assignment?", vbYesNo, "Fiber Assigned!")
        If result = vbNo Then Exit Sub
    End If
    
    Dim iFirst, iNeeded, iIndex As Integer
    Dim strName, strLine As String
    Dim iBCount, iECount As Integer
    
    'If tbTapNeeded.Value = "" Then
        'iNeeded = 1
    'Else
        iNeeded = CInt(tbTapNeeded.Value)
    'End If
    
    iIndex = lbCounts.ListIndex
    iFirst = iIndex - iNeeded + 1
    
    If iFirst < 1 Then
        MsgBox "Out of range." & vbCr & "Fibers needed in cable don't exist."
        Exit Sub
    End If
    
    strName = lbCounts.List(iIndex, 1)
    iBCount = 0
    iECount = CInt(lbCounts.List(iIndex, 2))
    
    For i = iIndex To iFirst Step -1
        If strName = lbCounts.List(i, 1) And lbCounts.List(i, 3) = "<>" Then
            iBCount = CInt(lbCounts.List(i, 2))
            
            lbCounts.List(i, 3) = tbActivePole.Value
            lbCounts.List(i, 6) = "FUTURE"
        Else
            strLine = strName & ": " & i & " not available."
            MsgBox strLine
            GoTo Exit_Next
        End If
    Next i
Exit_Next:
    
    If iBCount = 0 Then
        strLine = strName & ": " & iECount
    Else
        If iBCount = iECount Then
            strLine = strName & ": " & iECount
        Else
            strLine = strName & ": " & iBCount & "-" & iECount
        End If
        
        
    End If
    
    If tbCallout.Value = "" Then
        tbCallout.Value = strLine & " FUTURE"
    Else
        tbCallout.Value = tbCallout.Value & vbCr & strLine & " FUTURE"
    End If
    
    lbCounts.ListIndex = i
    
    bUpdatePole = True
    
    Call CreateCallout
End Sub

Private Sub cbDrop_Click()
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim strCable, strCounts, strSpliced, strCallout As String
    Dim vLine, vItem As Variant
    Dim iExist As Integer
    
    iExist = 0
    Me.Hide
    Load CountsTap
        'CountsTap.lbMain.ColumnCount = 4
        'CountsTap.lbMain.ColumnWidths = "24;72;36;30"
        
        CountsTap.tbStructure.Value = tbActivePole.Value
        
        For i = 1 To lbCounts.ListCount - 1
            CountsTap.lbMain.AddItem i
            
            If lbCounts.List(i, 3) = "<>" Then
                CountsTap.lbMain.List(i - 1, 1) = lbCounts.List(i, 1)
                CountsTap.lbMain.List(i - 1, 2) = lbCounts.List(i, 2)
            Else
                CountsTap.lbMain.List(i - 1, 1) = "XD"
                CountsTap.lbMain.List(i - 1, 2) = i
            End If
            CountsTap.lbMain.List(i - 1, 3) = ""
        Next i
        
        CountsTap.show
        
        If CountsTap.cbChanged.Value = False Then GoTo Exit_Sub
        If CountsTap.cbExisting.Value = True Then iExist = 1
        
        strLine = CountsTap.tbPosition.Value & ": " & CountsTap.cbCblType.Value & "(" & CountsTap.cbCableSize.Value & ")"
        If Not CountsTap.cbSuffix.Value = "" Then strLine = strLine & CountsTap.cbSuffix.Value
        strLine = strLine & " / " & Replace(CountsTap.tbResult.Value, vbCr, " + ")
        strLine = Replace(strLine, vbLf, "")
        
        'strLine = cable callout
        
        strSpliced = Replace(CountsTap.tbResult.Value, vbLf, "")
        vLine = Split(strSpliced, vbCr)
        strSpliced = ""
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), ": ")
            If Not vItem(0) = "XD" Then
                If strSpliced = "" Then
                    strSpliced = vLine(i)
                Else
                    strSpliced = strSpliced & " + " & vLine(i)
                End If
            End If
        Next i
        
        strCallout = Replace(strSpliced, " + ", vbCr)
        
        For i = 0 To CountsTap.lbMain.ListCount - 1
            If Left(CountsTap.lbMain.List(i, 3), 1) = "Y" Then
                lbCounts.List(i + 1, 1) = "XD"
                lbCounts.List(i + 1, 2) = lbCounts.List(i + 1, 0)
                
                'lbCounts.List(i + 1, 3) = tbActivePole.Value
                'lbCounts.List(i + 1, 6) = "TAP"
            End If
        Next i
    
        If tbCallout.Value = "" Then
            tbCallout.Value = strCallout
        Else
            tbCallout.Value = tbCallout.Value & vbCr & strCallout
        End If
        
        Dim objEntity As AcadEntity
        Dim objBlock As AcadBlockReference
        Dim vReturnPnt, vAttList As Variant
        Dim iAtt As Integer
        
        On Error Resume Next
        
        Err = 0
        ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block:"
        If Not Err = 0 Then Exit Sub
    
        If Not TypeOf objEntity Is AcadBlockReference Then Exit Sub
        Set objBlock = objEntity
        
        Select Case objBlock.Name
            Case "sPole"
                iAtt = 25
            Case "sPed", "sHH", "sPanel"
                iAtt = 5
            Case Else
                Exit Sub
        End Select
        
        vAttList = objBlock.GetAttributes
        vAttList(iAtt).TextString = strLine
           
        objBlock.Update
        
        bUpdatePole = True
        
        'lbCounts.ListIndex = i
        Call CreateCallout
        
Exit_Sub:
    
    Unload CountsTap
    Me.show
End Sub

Private Sub cbExistingCable_Click()
    Dim objEntity As AcadEntity
    Dim vAttList, vBasePnt As Variant
    Dim vCable, vCounts, vLine, vItem As Variant
    Dim strLine, strNames As String
    Dim strCables As String
    Dim iFiber As Integer
    Dim iAtt As Integer
    
    strNames = ""
    
    Me.Hide
    
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Select Pole: "
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
    Else
        MsgBox "Not a valid object."
        Me.show
        Exit Sub
    End If
    
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 25
        Case "sPed", "sHH", "sPanel", "sMH"
            iAtt = 5
        Case Else
            MsgBox "Not a valid block."
            Me.show
            Exit Sub
    End Select
    
    vAttList = objBlock.GetAttributes
    
    If vAttList(iAtt).TextString = "" Then
        MsgBox "Structure has no Cables."
        Me.show
        Exit Sub
    End If
    
    strCables = Replace(vAttList(iAtt).TextString, vbLf, "")
    vCable = Split(strCables, vbCr)
        
    If UBound(vCable) > 0 Then
        Load SelectCable
            SelectCable.cbNewCable.Enabled = False
                
            For i = 0 To UBound(vCable)
                SelectCable.lbCable.AddItem vCable(i)
            Next i
                
            SelectCable.show
                
            strCables = SelectCable.lbCable.List(SelectCable.lbCable.ListIndex)
        Unload SelectCable
    End If
        
    strCableCallout = strCables
        
    vCable = Split(strCables, " / ")
    strLine = Replace(vCable(1), " + ", vbCr)
        
    tbCblCallout.Value = strLine
    
    tbActivePole.Value = vAttList(0).TextString
    
    vCounts = Split(vCable(0), ")")
    vLine = Split(vCounts(0), "(")
    
    vItem = Split(vLine(0), ": ")
    
    cbCblType.Value = vItem(1)
    cbCableSize.Value = vLine(1)
    cbSuffix.Value = vCounts(1)
    tbPosition.Value = vItem(0)
    
    lbCounts.AddItem ""
    For i = 1 To CInt(vLine(1))
        lbCounts.AddItem i
    Next i
    
    iFiber = 1
    vCounts = Split(tbCblCallout.Value, vbCrLf)
    
    'MsgBox vCounts(0) & " + " & vCounts(1)
    
    For i = 0 To UBound(vCounts)
        If iFiber > CInt(cbCableSize.Value) Then
            MsgBox "Number of fibers in Counts > the cable size"
            GoTo Cable_Filled
        End If
        
        vLine = Split(vCounts(i), ": ")
        vItem = Split(vLine(1), "-")
        
        'If Not vLine(0) = "XD" Then
        If InStr(vLine(0), "XD") < 1 Then
            If strNames = "" Then
                strNames = vLine(0)
            Else
                strNames = strNames & ";" & vLine(0)
            End If
            strPreviousBlock = vLine(2)
            tbF2Name.Value = vLine(0)
        End If
        
        If UBound(vItem) = 0 Then
            lbCounts.List(iFiber, 1) = vLine(0)
            lbCounts.List(iFiber, 2) = vItem(0)
            lbCounts.List(iFiber, 3) = "<>"
            lbCounts.List(iFiber, 4) = "<>"
            lbCounts.List(iFiber, 5) = "<>"
            lbCounts.List(iFiber, 6) = "<>"
            lbCounts.List(iFiber, 7) = "<>"
            iFiber = iFiber + 1
        Else
            For j = CInt(vItem(0)) To CInt(vItem(1))
                lbCounts.List(iFiber, 1) = vLine(0)
                lbCounts.List(iFiber, 2) = j
                lbCounts.List(iFiber, 3) = "<>"
                lbCounts.List(iFiber, 4) = "<>"
                lbCounts.List(iFiber, 5) = "<>"
                lbCounts.List(iFiber, 6) = "<>"
                lbCounts.List(iFiber, 7) = "<>"
                iFiber = iFiber + 1
            Next j
        End If
    Next i
    
Cable_Filled:
    If iFiber < CInt(cbCableSize.Value) + 1 Then
        For i = iFiber To CInt(cbCableSize.Value) + 1
            lbCounts.List(i, 1) = vLine(0)
            lbCounts.List(i, 2) = i
            lbCounts.List(i, 3) = "<>"
            lbCounts.List(i, 4) = "<>"
            lbCounts.List(i, 5) = "<>"
            lbCounts.List(i, 6) = "<>"
            lbCounts.List(i, 7) = "<>"
            Call CreateCallout
        Next i
        
        MsgBox "Number of fibers in Counts < the cable size"
    End If
    
    Call EnableAll
    
    vLine = Split(strNames, ";")
    For i = 0 To UBound(vLine)
        GetSpliced CStr(vLine(i))
    Next i
    
    'Call CreateCallout
    
    Call AutoFiber
    
    iAtt = iAtt + 1
    
    If Not vAttList(iAtt).TextString = "" Then
        Dim vClosures As Variant
        Dim strPosition, strClosure As String
        
        strPosition = "[" & tbPosition.Value & "]"
        
        strClosure = Replace(vAttList(iAtt).TextString, vbLf, "")
        vClosures = Split(strClosure, vbCr)
        
        For i = 0 To UBound(vClosures)
            If InStr(vClosures(i), strPosition) > 0 Then
                vItem = Split(vClosures(i), "] ")
                vLine = Split(vItem(1), " + ")
                
                strLine = vLine(0)
                If UBound(vLine) > 0 Then
                    For j = 1 To UBound(vLine)
                        strLine = strLine & vbCr & vLine(j)
                    Next j
                End If
                
                tbCallout.Value = strLine
                
                GoTo Exit_Sub
            End If
        Next i
    End If
    
Exit_Sub:
    
    cbGetPole.Enabled = True
    cbAddToPole.Enabled = True
    cbNewCable.Enabled = False
    cbExistingCable.Enabled = False
    
    Me.show
End Sub

Private Sub cbGetPole_Click()
    Me.Hide
    Call CreateCallout
    
    Call UpdateTerminal
    
    Call UpdateTerminalCallout
    
    Dim strTest As String
    strTest = tbActivePole.Value
    
    Call GetActivePole
    
    If strTest = tbActivePole.Value Then
        Me.show
        Exit Sub
    End If
    
    Call UpdateCable
    
    Call UpdateCableCallout
    
    Dim vAttList, vList, vItem, vSplice As Variant
    Dim strLine, strPosition As String
    Dim iAtt As Integer
    
    vAttList = objBlock.GetAttributes
    
    Select Case objBlock.Name
        Case "sPole"
            If Not vAttList(8).TextString = "" Then
                Select Case UCase(vAttList(8).TextString)
                    Case "M"
                        vAttList(8).TextString = "MGNV"
                    Case "T"
                        vAttList(8).TextString = "TGB"
                    Case "B"
                        vAttList(8).TextString = "BROKEN"
                    Case "", "NA", "N/A"
                        vAttList(8).TextString = "NO GRD"
                End Select
                tbGround.Value = vAttList(8).TextString
            End If
            
            iAtt = 26
        Case Else
            iAtt = 6
    End Select
    
    If vAttList(iAtt).TextString = "" Then
        tbCallout.Value = ""
    Else
        strLine = Replace(vAttList(iAtt).TextString, vbLf, "")
        
        vList = Split(strLine, vbCr)
        
        For i = 0 To UBound(vList)
            vItem = Split(vList(i), "] ")
            strPosition = Replace(vItem(0), "[", "")
            
            If strPosition = tbPosition.Value Then
                If Not vItem(1) = "" Then
                    vSplice = Split(vItem(1), " + ")
                    
                    strLine = vSplice(0)
                
                    If UBound(vSplice) > 0 Then
                        For j = 1 To UBound(vSplice)
                            strLine = strLine & vbCr & vSplice(j)
                        Next j
                    
                        tbCallout.Value = strLine
                    
                        GoTo Found_Splices
                    End If
                End If
            End If
        Next i
        
        tbCallout.Value = ""
Found_Splices:
    End If
    
    iAtt = iAtt + 1
    
    tbUnits.Value = ""
    If Not vAttList(iAtt).TextString = "" Then
        vList = Split(vAttList(iAtt).TextString, ";;")
        
        strTest = vList(0)
        If UBound(vList) > 0 Then
            For i = 0 To UBound(vList)
                strTest = strTest & vbCr & vList(i)
            Next i
        End If
        
        tbUnits.Value = strTest
    End If
    
    Me.show
End Sub

Private Sub cbInOutTap_Click()
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim strCable, strCounts, strSpliced, strCallout As String
    Dim vLine, vItem As Variant
    'Dim iExist As Integer
    
    'iExist = 0
    Me.Hide
    Load CountsTap
        CountsTap.tbStructure.Value = tbActivePole.Value
        
        For i = 1 To lbCounts.ListCount - 1
            CountsTap.lbMain.AddItem i
            
            If lbCounts.List(i, 3) = "<>" Then
                CountsTap.lbMain.List(i - 1, 1) = lbCounts.List(i, 1)
                CountsTap.lbMain.List(i - 1, 2) = lbCounts.List(i, 2)
            Else
                CountsTap.lbMain.List(i - 1, 1) = "XD"
                CountsTap.lbMain.List(i - 1, 2) = i
            End If
            CountsTap.lbMain.List(i - 1, 3) = ""
        Next i
        
        CountsTap.show
        
        'GoTo Exit_Sub
        
        If CountsTap.cbChanged.Value = False Then GoTo Exit_Sub
        'If CountsTap.cbExisting.Value = True Then iExist = 1
        
        strLine = CountsTap.tbPosition.Value & ": " & CountsTap.cbCblType.Value & "(" & CountsTap.cbCableSize.Value & ")"
        If Not CountsTap.cbSuffix.Value = "" Then strLine = strLine & CountsTap.cbSuffix.Value
        strLine = strLine & " / " & Replace(CountsTap.tbResult.Value, vbCr, " + ")
        strLine = Replace(strLine, vbLf, "")
        
        'strLine = cable callout
        
        strSpliced = Replace(CountsTap.tbResult.Value, vbLf, "")
        vLine = Split(strSpliced, vbCr)
        strSpliced = ""
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), ": ")
            If Not vItem(0) = "XD" Then
                If strSpliced = "" Then
                    strSpliced = vLine(i)
                Else
                    strSpliced = strSpliced & " + " & vLine(i)
                End If
            End If
        Next i
        
        strCallout = Replace(strSpliced, " + ", vbCr)
    
    Unload CountsTap
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt, vAttList As Variant
    Dim strReplace As String
    Dim iAtt As Integer
    
    On Error Resume Next
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 25
        Case "sPed", "sHH", "sPanel"
            iAtt = 5
        Case Else
            GoTo Exit_Sub
    End Select
    
    vAttList = objBlock.GetAttributes
    strReplace = vAttList(0).TextString
    
    strLine = Replace(strLine, "<<???>>", strReplace)
    strCallout = Replace(strCallout, "<<???>>", strReplace)
        
    If tbCallout.Value = "" Then
        tbCallout.Value = strCallout
    Else
        tbCallout.Value = tbCallout.Value & vbCr & strCallout
    End If
    
    vAttList(iAtt).TextString = strLine
    objBlock.Update
        
    bUpdatePole = True
    
    'MsgBox strLine
    'MsgBox strCallout
    
    'Call CreateCallout
    
Exit_Sub:
    
    Me.show
End Sub

Private Sub cbNewCable_Click()
    'If cbCblType.Value = "" Then
        'MsgBox "Need Cable Type"
        'Exit Sub
    'End If
    'If cbCableSize.Value = "" Then
        'MsgBox "Need Cable Size"
        'Exit Sub
    'End If
    If tbF1Name.Value = "" Then
        MsgBox "Need F1 Name:" & vbCr & "Enter ""NA"" if no F1 counts."
        Exit Sub
    End If
    'If tbPosition.Value = "" Then
        'MsgBox "Need Cable Position"
        'Exit Sub
    'End If
    
    Dim strFileName As String
    Dim vName, vLine, vItem As Variant
    Dim fName As String
    
    Me.Hide
    Load NewCableForm
        NewCableForm.show
        
        If Not UCase(tbF1Name.Value) = "NA" Then
            vName = Split(ThisDrawing.Name, " ")
            strFileName = strSplitterFilePath & "\" & vName(0) & " Counts -" & tbF1Name.Value & ".mcl"
    
            fName = Dir(strFileName)
            If fName = "" Then
                If NewCableForm.lbCounts.ListCount < 1 Then GoTo Unload_Form
                
                For i = 0 To NewCableForm.lbCounts.ListCount - 1
                    vLine = Split(NewCableForm.lbCounts.List(i, 1), ": ")
                    
                    If vLine(0) = tbF1Name.Value Then GoTo Found_F1
                Next i
                
                MsgBox "No F1 counts found in Callout."
                GoTo Unload_Form
                
Found_F1:
                    
                vItem = Split(vLine(1), "-")
                
                Open strFileName For Output As #1
        
                Print #1, tbF1Name.Value & " F1 COUNTS"
        
                For i = 1 To CInt(vItem(1))
                    Print #1, i & vbTab & "<>" & vbTab & "<>" & vbTab & "<>" & vbTab & "<>" & vbTab & "<>"
                Next i
        
                Close #1
            End If
        End If
Unload_Form:
    Unload NewCableForm
    
    Call GetActivePole
    
    Dim objEntity As AcadEntity
    Dim objTemp As AcadBlockReference
    Dim vReturnPnt, vList As Variant
    
    On Error Resume Next
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Source Block: "
    If Not Err = 0 Then
        MsgBox "Data not added to block!"
        Me.show
        Exit Sub
    End If
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objTemp = objEntity
    Else
        MsgBox "Data not added to block!"
        Me.show
        Exit Sub
    End If
    
    vList = objTemp.GetAttributes
    strPreviousBlock = vList(0).TextString
    
    Call EnableAll
    
    Call AutoFiber
    
    Call CreateCallout
    
    Dim vText As Variant
    Dim strLine As String
    
    vText = Split(tbCblCallout.Value, vbCr)
    For i = 0 To UBound(vText)
        vLine = Split(vText(i), ": ")
        If InStr(vLine(0), "XD") > 0 Then
            vText(i) = vLine(0) & ": " & vLine(1) & ": END"
        End If
    Next i
    
    strLine = vText(0)
    If UBound(vText) > 0 Then
        For i = 1 To UBound(vText)
            strLine = strLine & vbCr & vText(i)
        Next i
    End If
    
    tbCblCallout.Value = strLine
    
    Call UpdateCable
    
    cbGetPole.Enabled = True
    cbAddToPole.Enabled = True
    cbNewCable.Enabled = False
    cbExistingCable.Enabled = False
    
    Me.show
    
End Sub

Private Sub cbQuit_Click()
    Me.Hide
    
    Dim result As Integer
    
    If bSave = True Then
        result = MsgBox("Save changes to Master List?", vbYesNo, "Save Changes")
        If result = vbYes Then
            Call UpdateTerminal
            Call SaveAll
            bSave = False
        End If
    End If
    
    
    If bUpdatePole = True Then
        result = MsgBox("Save changes to Current Pole?", vbYesNo, "Save Changes")
        If result = vbYes Then
            Call UpdateCableCallout
            Call UpdateTerminal
        End If
    End If
    
End Sub

Private Sub cbRefresh_Click()
    Me.Repaint
End Sub

Private Sub cbRemoveWL_Click()
    If lbCounts.ListIndex < 0 Then Exit Sub
    
    Dim iIndex As Integer
    
    iIndex = lbCounts.ListIndex
    
    lbCounts.List(iIndex, 3) = "<>"
    lbCounts.List(iIndex, 4) = "<>"
    lbCounts.List(iIndex, 5) = "<>"
    lbCounts.List(iIndex, 6) = "<>"
    lbCounts.List(iIndex, 7) = "<>"
    
    If Not iIndex = 0 Then lbCounts.ListIndex = iIndex - 1
End Sub

Private Sub cbSaveSplitter_Click()
    Call SaveAll
End Sub

Private Sub cbSplitter_Click()
    'GoTo Skip_This
    
    Dim vLine As Variant
    Dim iIndex, iCIndex As Integer
    Dim strF1 As String
    Dim vAttList As Variant
    Dim vItem, vCounts, vTemp As Variant
    'Dim strCount(2) As String
    Dim strPrevious, strCurrent As String
    Dim strLine, strItem, strCallout As String
    Dim strMessage As String
    Dim iAtt, iCount, iFiber As Integer
    Dim iStart, iEnd As Integer
    
    strF1 = ""
    For i = 0 To lbCounts.ListCount - 1
        If lbCounts.Selected(i) = True Then
            strF1 = lbCounts.List(i, 1) & ": " & lbCounts.List(i, 2)
            iCIndex = i
            GoTo Found_F1
        End If
    Next i
    
    MsgBox "No fiber selected."
    Exit Sub
    
Found_F1:
    
    Me.Hide
    
    Load CountsSplitter
        CountsSplitter.tbActivePole.Value = tbActivePole.Value
        CountsSplitter.tbSplice.Value = strF1
        CountsSplitter.tbF2Name.Value = Replace(strF1, ": ", "-")
        'If Not strF1 = "" Then CountsSplitter.tbF2Name.Value = Replace(strF1, ": ", "-")
        
        CountsSplitter.show
        
        If CountsSplitter.cbChanged.Value = False Then GoTo Exit_Sub
        
        tbSplitterName.Value = CountsSplitter.tbSplitterName.Value
        
        '<----------------------------------------Add Splices to callout
        If tbCallout.Value = "" Then
            tbCallout.Value = CountsSplitter.tbResult.Value
        Else
            tbCallout.Value = tbCallout.Value & vbCr & CountsSplitter.tbResult.Value
        End If
        
        '<----------------------------------------Add Spliced Fibers to cable
        'If CountsSplitter.iExisting = 0 Then
            For i = 0 To CountsSplitter.lbMain.ListCount - 1
                If Left(CountsSplitter.lbMain.List(i, 3), 2) = "A " Then
                    vLine = Split(CountsSplitter.lbMain.List(i, 3), "A ")
                    iIndex = CInt(vLine(1))
                    
                    lbCounts.List(iIndex, 1) = CountsSplitter.lbMain.List(i, 1)
                    lbCounts.List(iIndex, 2) = CountsSplitter.lbMain.List(i, 2)
                End If
            Next i
        'End If
        
        '<----------------------------------------Assign Splitter location to cable
        'vLine = Split(CountsSplitter.tbSplice.Value, ": ")
        'For i = 0 To CountsSplitter.lbMain.ListCount - 1
            'If lbCounts.List(i, 1) = vLine(0) Then
                'If lbCounts.List(i, 2) = vLine(1) Then
                    lbCounts.List(iCIndex, 3) = CountsSplitter.tbActivePole.Value
                    lbCounts.List(iCIndex, 6) = "SPLITTER"
                    lbCounts.List(iCIndex, 7) = CountsSplitter.tbSplitterName.Value
                    
                    'GoTo Exit_Sub2
                'End If
            'End If
        'Next i
        
'Exit_Sub2:
    
        '<----------------------------------------Add counts to Backspan
        
        Select Case objBlock.Name
            Case "sPole"
                iAtt = 25
            Case "sPed", "sHH", "sFP", "sPanel"
                iAtt = 5
            Case Else
                GoTo Exit_Sub
        End Select
        
        vAttList = objBlock.GetAttributes
        vTemp = Split(vAttList(iAtt).TextString, " / ")
        strCallout = vTemp(0) & " / "
        vLine = Split(vTemp(0), ")")
        vCount = Split(vLine(0), "(")
        iCount = CInt(vCount(1))
        
        strMessage = "iCount: " & iCount & vbCr
        
        ReDim strCount(iCount) As String  '<------------ Break counts out individually into a list
        vLine = Split(vTemp(1), " + ")
        iFiber = 1
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), ": ")
            vCount = Split(vItem(1), "-")
            iStart = CInt(vCount(0))
            If UBound(vCount) > 0 Then
                iEnd = CInt(vCount(1))
            Else
                iEnd = iStart
            End If
            
            For j = iStart To iEnd
                strCount(iFiber) = vItem(0) & ": " & j & ": " & vItem(2)
                iFiber = iFiber + 1
            Next j
        Next i
        
        For i = 0 To CountsSplitter.lbMain.ListCount - 1  '<------------ add Backspan counts
            If InStr(CountsSplitter.lbMain.List(i, 3), "B ") > 0 Then
                vLine = Split(CountsSplitter.lbMain.List(i, 3), "B ")
                iIndex = CInt(vLine(1))
                
                strLine = CountsSplitter.lbMain.List(i, 1) & ": " & CountsSplitter.lbMain.List(i, 2)
                strLine = strLine & ": " & tbActivePole.Value
                strCount(iIndex) = strLine
            End If
        Next i
        
        '<---------------------------- Need to combine them back into an attribute and save current block
        
        strLine = ""
        
        vLine = Split(strCount(1), ": ")
        strPrevious = vLine(0)
        'strCurrent = strPrevious
        iStart = CInt(vLine(1))
        iEnd = iStart
    
        For i = 2 To iCount
            vLine = Split(strCount(i), ": ")
            strCurrent = vLine(0)
        
            If strCurrent = strPrevious Then
                iEnd = CInt(vLine(1))
            Else
                strItem = strPrevious & ": " & iStart
                If iEnd > iStart Then strItem = strItem & "-" & iEnd
                strItem = strItem & ": " & tbActivePole.Value
                
                If strLine = "" Then
                    strLine = strItem
                Else
                    strLine = strLine & " + " & strItem
                End If
                
                strPrevious = strCurrent
                iStart = CInt(vLine(1))
                iEnd = iStart
            End If
        Next i
             
        strItem = strPrevious & ": " & iStart
        If iEnd > iStart Then strItem = strItem & "-" & iEnd
        
        strCallout = strCallout & strLine & " + " & strItem & ": " & tbActivePole.Value
    'MsgBox vAttList(iAtt).TextString & vbCr & strCallout
        vAttList(iAtt).TextString = strCallout
        objBlock.Update
        'tbResult.Value = strLine
    
        bUpdatePole = True
        bSave = True
    
        Call CreateCallout
        Call AddSplitter
        
Exit_Sub:
    Unload CountsSplitter
    
    Me.show
    Exit Sub
    
    
    
    
    
Skip_This:
    
    'If tbF1Name.Value = "" Then
        'MsgBox "You need to add a F1 Name"
        'tbF1Name.SetFocus
        'Exit Sub
    'End If
    
    If tbF2Name.Value = "" Then
        MsgBox "You need to add a F2 Name"
        tbF2Name.SetFocus
        Exit Sub
    End If
    
    If tbSplitterName.Value = "" Then
        MsgBox "You need to add a Splitter Location"
        tbSplitterName.SetFocus
        Exit Sub
    End If
    
    If cbSplitterSize.Value = "" Then
        MsgBox "You need to add a Splitter size"
        cbSplitterSize.SetFocus
        Exit Sub
    End If
    
    If tbSplice.Value = "" Then
        MsgBox "You need to add a Splitter splice"
        tbSplice.SetFocus
        Exit Sub
    End If
    
    If tbFFiber.Value = "" Then
        MsgBox "You need to add the From Fiber"
        tbFFiber.SetFocus
        Exit Sub
    End If
    
    If tbTFiber.Value = "" Then
        MsgBox "You need to add the To Fiber"
        tbTFiber.SetFocus
        Exit Sub
    End If
    
    Dim iSelected As Integer
    'Dim i, j As Integer
    
    For i = 0 To lbCounts.ListCount - 1
        If lbCounts.Selected(i) = True Then
            iSelected = i
            GoTo Found_Selected
        End If
    Next i
    
    MsgBox "No Count Selected."
    Exit Sub
    
Found_Selected:
    
    lbCounts.List(i, 3) = tbActivePole.Value
    lbCounts.List(i, 6) = "SPLITTER"
    lbCounts.List(i, 7) = tbSplitterName.Value
    
    If tbCallout = "" Then
        tbCallout.Value = lbCounts.List(i, 1) & ": " & lbCounts.List(i, 2)
    Else
        tbCallout.Value = tbCallout.Value & vbCr & lbCounts.List(i, 1) & ": " & lbCounts.List(i, 2)
    End If
    
    tbCallout.Value = tbCallout.Value & vbCr & tbF2Name.Value & ": 1-" & cbSplitterSize.Value
    
    '<------------------------------------------------ Update F1.mcl
    
    j = 1
    
    For i = CInt(tbFFiber.Value) To CInt(tbTFiber.Value)
        lbCounts.List(i, 1) = tbF2Name.Value
        lbCounts.List(i, 2) = j
        lbCounts.List(i, 3) = "<>"
        lbCounts.List(i, 4) = "<>"
        lbCounts.List(i, 5) = "<>"
        lbCounts.List(i, 6) = "<>"
        lbCounts.List(i, 7) = "<>"
        
        j = j + 1
    Next i
    
    lbCounts.ListIndex = CInt(tbTFiber.Value)
    lbCounts.Selected(CInt(tbTFiber.Value)) = True
    
    Call CreateCallout
    
    Call AddSplitter
    
    Call SaveAll
    'Call SaveF1
    
    bUpdatePole = True
    bSave = True
End Sub

Private Sub cbSplitterSize_Change()
    Call AutoFiber
End Sub

Private Sub cbTerminal_Click()
    If objBlock Is Nothing Then
        MsgBox "No Pole Selected"
        Exit Sub
    End If
    
    Dim iSelected As Integer
    
    For i = 0 To lbCounts.ListCount - 1
        If lbCounts.Selected(i) = True Then
            iSelected = i
            GoTo Found_Selected
        End If
    Next i
    
    MsgBox "No Count Selected."
    Exit Sub
    
Found_Selected:
    Dim result As Integer
    
    If Not lbCounts.List(i, 3) = "<>" Then
        
        result = MsgBox("Overwrite previous assignment?", vbYesNo, "Fiber Assigned!")
        If result = vbNo Then Exit Sub
    End If
    
    Dim objEntity As AcadEntity
    Dim objRES As AcadBlockReference
    Dim vAttList, vBasePnt As Variant
    Dim strLine As String
    Dim strName As String
    Dim iBCount, iECount As Integer
    
    iBCount = 0: iECount = 0
    
    Me.Hide
    On Error Resume Next
    
Next_One:
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Select Building: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not lbCounts.List(iSelected, 3) = "<>" Then
        
        result = MsgBox("Overwrite previous assignment?", vbYesNo, "Fiber Assigned!")
        If result = vbNo Then Exit Sub
    End If
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objRES = objEntity
        vAttList = objRES.GetAttributes
            
        Select Case objRES.Name
            'Case "RES", "BUSINESS", "MDU", "TRLR", "SCHOOL", "CHURCH", "LOT"
            Case "SG"
                lbCounts.List(iSelected, 3) = tbActivePole.Value
                lbCounts.List(iSelected, 4) = "SG"
                lbCounts.List(iSelected, 5) = vAttList(1).TextString
                lbCounts.List(iSelected, 6) = "SMARTGRID"
                lbCounts.List(iSelected, 7) = vAttList(0).TextString
                
                strLine = lbCounts.List(iSelected, 1) & ": " & lbCounts.List(iSelected, 2)
                'If cbDrop.Value = False Then strline = "(" & strline & ")"
                vAttList(2).TextString = tbActivePole.Value & " - " & strLine
                objRES.Update
                
                strName = lbCounts.List(iSelected, 1)
                If iECount = 0 Then
                    iECount = CInt(lbCounts.List(iSelected, 2))
                Else
                    iBCount = CInt(lbCounts.List(iSelected, 2))
                End If
                
                
            Case "Customer"
                lbCounts.List(iSelected, 3) = tbActivePole.Value
                lbCounts.List(iSelected, 4) = vAttList(1).TextString
                lbCounts.List(iSelected, 5) = vAttList(2).TextString
                lbCounts.List(iSelected, 6) = vAttList(0).TextString
                If vAttList(3).TextString = "" Then
                    lbCounts.List(iSelected, 7) = "<>"
                Else
                    lbCounts.List(iSelected, 7) = vAttList(3).TextString
                End If
                
                strLine = lbCounts.List(iSelected, 1) & ": " & lbCounts.List(iSelected, 2)
                If cbDrop.Value = False Then strLine = "(" & strLine & ")"
                vAttList(4).TextString = tbActivePole.Value & " - " & strLine
                objRES.Update
                
                strName = lbCounts.List(iSelected, 1)
                If iECount = 0 Then
                    iECount = CInt(lbCounts.List(iSelected, 2))
                Else
                    iBCount = CInt(lbCounts.List(iSelected, 2))
                End If
        End Select
    End If
    
    iSelected = iSelected - 1
    
    GoTo Next_One
    
Exit_Sub:
    lbCounts.ListIndex = iSelected
    
    If iECount > 0 Then
        If iBCount = 0 Then
            strName = strName & ": " & iECount
        Else
            strName = strName & ": " & iBCount & "-" & iECount
        End If
        
        strName = strName & ": " & strPreviousBlock
        
        If tbCallout.Value = "" Then
            tbCallout.Value = strName
        Else
            tbCallout.Value = tbCallout.Value & vbCr & strName
        End If
    End If
    
    bUpdatePole = True
    bSave = True
    
    Call CreateCallout
    
    Call UpdateTerminalCallout
    
    Me.show
End Sub

Private Sub cbUpdateCblCO_Click()
    Dim vAttList As Variant
    
    vAttList = objBlock.GetAttributes
    
End Sub

Private Sub cbUpdateCO_Click()
    Call CreateCallout
End Sub

Private Sub cbUpdateHO1_Click()
    Call UpdateTerminalCallout
End Sub

Private Sub Label24_Click()
    If lbCounts.ListIndex < 0 Then Exit Sub
    
    tbF2Name.Value = tbF2Name.Value & "-" & lbCounts.List(lbCounts.ListIndex, 2)
End Sub

Private Sub Label26_Click()
    Dim objEntity As AcadEntity
    Dim objRES As AcadBlockReference
    Dim vAttList, vBasePnt As Variant
    Dim strName As String
    
    Me.Hide
    On Error Resume Next
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Select Building: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objRES = objEntity
        vAttList = objRES.GetAttributes
            
        Select Case objRES.Name
            'Case "RES", "BUSINESS", "MDU", "TRLR", "SCHOOL", "CHURCH", "LOT"
                'tbSplitterName.Value = vAttList(0).TextString & " " & vAttList(1).TextString
            Case "Customer"
                strName = vAttList(1).TextString & " " & vAttList(2).TextString
                tbSplitterName.Value = Replace(strName, "  ", " ")
        End Select
        'Set objRES = Nothing
    End If
    
Exit_Sub:
    Me.show
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub lbCounts_Click()
    tbSplice.Value = lbCounts.List(lbCounts.ListIndex, 1) & ": " & lbCounts.List(lbCounts.ListIndex, 2)
End Sub

Private Sub lbCounts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strLine As String
    Dim iIndex As Integer
    
    iIndex = lbCounts.ListIndex
    
    strLine = "Fiber: " & vbTab & lbCounts.List(iIndex, 0) & vbCr
    strLine = strLine & lbCounts.List(iIndex, 1) & ": " & lbCounts.List(iIndex, 2) & vbCr & vbCr
    strLine = strLine & "Pole: " & vbTab & lbCounts.List(iIndex, 3) & vbCr & vbCr
    strLine = strLine & "Address:  " & lbCounts.List(iIndex, 4) & " " & lbCounts.List(iIndex, 5) & vbCr
    strLine = strLine & "Type: " & vbTab & lbCounts.List(iIndex, 6) & vbCr & vbCr
    strLine = strLine & "Note: " & vbTab & lbCounts.List(iIndex, 7)
    
    MsgBox strLine
End Sub

Private Sub lbCounts_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            tbF1Name.Value = lbCounts.List(lbCounts.ListIndex, 1)
        Case vbKeyF2
            tbF2Name.Value = lbCounts.List(lbCounts.ListIndex, 1) & "-" & lbCounts.List(lbCounts.ListIndex, 2)
        Case vbKeyDelete
            lbCounts.List(lbCounts.ListIndex, 3) = "<>"
            lbCounts.List(lbCounts.ListIndex, 4) = "<>"
            lbCounts.List(lbCounts.ListIndex, 5) = "<>"
            lbCounts.List(lbCounts.ListIndex, 6) = "<>"
            lbCounts.List(lbCounts.ListIndex, 7) = "<>"
    End Select
End Sub

Private Sub tbActivePole_Change()
    tbPoleNumber.Value = tbActivePole.Value
End Sub

Private Sub tbF1Name_Change()
    'tbF2Name.Value = tbF1Name.Value & "-"
End Sub

Private Sub tbSplice_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    tbF2Name.Value = Replace(tbSplice.Value, ": ", "-")
End Sub

Private Sub UserForm_Initialize()
    cbCblType.AddItem "CO"
    cbCblType.AddItem "BFO"
    cbCblType.AddItem "UO"
    cbCblType.Value = "CO"
    
    cbSuffix.AddItem ""
    cbSuffix.AddItem "E"
    cbSuffix.AddItem "6M-EHS"
    cbSuffix.AddItem "6M"
    cbSuffix.AddItem "10M"
    'cbSuffix.Value = "6M"
    
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    cbCableSize.AddItem "12"
    cbCableSize.AddItem "24"
    cbCableSize.AddItem "36"
    cbCableSize.AddItem "48"
    cbCableSize.AddItem "72"
    cbCableSize.AddItem "96"
    cbCableSize.AddItem "144"
    cbCableSize.AddItem "216"
    cbCableSize.AddItem "288"
    cbCableSize.AddItem "360"
    cbCableSize.AddItem "432"
    
    cbPlacement.AddItem "Leader"
    cbPlacement.AddItem "Away"
    cbPlacement.Value = "Leader"
    
    cbSplitterSize.AddItem "2"
    cbSplitterSize.AddItem "4"
    cbSplitterSize.AddItem "8"
    cbSplitterSize.AddItem "16"
    cbSplitterSize.AddItem "32"
    cbSplitterSize.AddItem "64"
    
    lbCounts.Clear
    lbCounts.ColumnCount = 8
    lbCounts.ColumnWidths = "24;72;30;66;36;136;48;96"
    
    bSave = 0
    bUpdatePole = 0
    
    strSplitterFilePath = LCase(ThisDrawing.Path)
End Sub

Private Sub EnableAll()
    cbCutCounts.Enabled = True
    cbTerminal.Enabled = True
    cbDrop.Enabled = True
    cbAddClosure.Enabled = True
    cbCallout.Enabled = True
    cbSplitter.Enabled = True
    cbUpdateHO1.Enabled = True
End Sub

Private Sub DisableAll()
    cbCutCounts.Enabled = False
    cbTerminal.Enabled = False
    cbDrop.Enabled = False
    cbAddClosure.Enabled = False
    cbCallout.Enabled = False
    cbSplitter.Enabled = False
    cbUpdateHO1.Enabled = False
End Sub

Private Sub CreateCallout()
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim vList As Variant
    Dim strCurrent, strPrevious As String
    Dim strList, strLine As String
    Dim iFiber, iStart, iEnd As Integer
    
    tbCblCallout.Value = ""
    strLine = ""
    'strPrevious = ""
    
    strPrevious = lbCounts.List(1, 1)
    If Not lbCounts.List(1, 3) = "<>" Then
        strPrevious = tbXD.Value
        iStart = CInt(lbCounts.List(1, 0))
    Else
        iStart = CInt(lbCounts.List(1, 2))
    End If
    
    For i = 2 To lbCounts.ListCount - 1
        strCurrent = lbCounts.List(i, 1)
        
        If Not lbCounts.List(i, 3) = "<>" Then strCurrent = tbXD.Value
        
        If strCurrent = strPrevious Then
            If Not CInt(lbCounts.List(i, 2)) = CInt(lbCounts.List(i - 1, 2)) + 1 Then
                If Not lbCounts.List(i - 1, 3) = "<>" Then
                    iEnd = CInt(lbCounts.List(i - 1, 0))
                Else
                    iEnd = CInt(lbCounts.List(i - 1, 2))
                End If
        
                If iStart = iEnd Then
                    strPrevious = strPrevious & ": " & iEnd & ": " & tbActivePole.Value
                Else
                    strPrevious = strPrevious & ": " & iStart & "-" & iEnd & ": " & tbActivePole.Value
                End If
            
                If strLine = "" Then
                    strLine = strPrevious
                Else
                    strLine = strLine & vbCr & strPrevious
                End If
            
                If strCurrent = tbXD.Value Then
                    iStart = CInt(lbCounts.List(i, 0))
                Else
                    iStart = CInt(lbCounts.List(i, 2))
                End If
            
                strPrevious = strCurrent
            End If
        Else
            'MsgBox strCurrent & vbCr & strPrevious
            
            'iEnd = CInt(lbCounts.List(i - 1, 2))
        
            If Not lbCounts.List(i - 1, 3) = "<>" Then
                iEnd = CInt(lbCounts.List(i - 1, 0))
            Else
                iEnd = CInt(lbCounts.List(i - 1, 2))
            End If
        
            If iStart = iEnd Then
                strPrevious = strPrevious & ": " & iEnd & ": " & tbActivePole.Value
            Else
                strPrevious = strPrevious & ": " & iStart & "-" & iEnd & ": " & tbActivePole.Value
            End If
            
            If strLine = "" Then
                strLine = strPrevious
            Else
                strLine = strLine & vbCr & strPrevious
            End If
            
            If strCurrent = tbXD.Value Then
                iStart = CInt(lbCounts.List(i, 0))
            Else
                iStart = CInt(lbCounts.List(i, 2))
            End If
            
            strPrevious = strCurrent
        End If
    Next i
    
    iEnd = CInt(lbCounts.List(lbCounts.ListCount - 1, 2))
    strLine = strLine & vbCr & strCurrent & ": " & iStart & "-" & iEnd & ": " & tbActivePole.Value
    
    tbCblCallout = strLine
    'Me.Repaint
End Sub

Private Sub SaveAll()
    Dim strFileName, strText As String
    Dim vName, vLine, vItem, vList As Variant
    Dim vText As Variant
    Dim strLine, strList As String
    Dim fName As String
    Dim iIndex As Integer
    
    strList = ""
    For i = 0 To lbCounts.ListCount - 1
        If Not lbCounts.List(i, 1) = "XD" Then
            If lbCounts.List(i, 6) = "" Then GoTo Next_I
            If lbCounts.List(i, 6) = "TAP" Then GoTo Next_I
            
            If strList = "" Then
                strList = lbCounts.List(i, 1)
            Else
                vLine = Split(strList, ";;")
                For j = 0 To UBound(vLine)
                    If vLine(j) = lbCounts.List(i, 1) Then GoTo Found_Name
                Next j
                
                strList = strList & ";;" & lbCounts.List(i, 1)
Found_Name:
            End If
        End If
Next_I:
    Next i
    
    If strList = "" Then Exit Sub
    
    vList = Split(strList, ";;")
    For n = 0 To UBound(vList)
        vName = Split(ThisDrawing.Name, " ")
        strFileName = strSplitterFilePath & "\" & vName(0) & " Counts -" & vList(n) & ".mcl"
    
        fName = Dir(strFileName)
        If fName = "" Then
            MsgBox "No MCL file for  " & vList(n) & "."
            GoTo Next_N
        End If
        
        Open strFileName For Input As #1
        strText = Input(LOF(1), 1)
        Close #1
        
        strText = Replace(strText, vbLf, "")
        vText = Split(strText, vbCr)
        
        For i = 0 To lbCounts.ListCount - 1
            If lbCounts.List(i, 1) = vList(n) Then
                strLine = lbCounts.List(i, 2)
                If lbCounts.List(i, 6) = "TAP" Then
                    strLine = strLine & vbTab & "<>" & vbTab & "<>" & vbTab & "<>" & vbTab & "<>" & vbTab & "<>" & vbTab & "<>"
                Else
                    If lbCounts.List(i, 3) = "" Or lbCounts.List(i, 3) = "<>" Then
                        strLine = strLine & vbTab & tbActivePole.Value
                    Else
                        strLine = strLine & vbTab & lbCounts.List(i, 3)
                    End If
                    If lbCounts.List(i, 4) = "" Then
                        strLine = strLine & vbTab & "<>"
                    Else
                        strLine = strLine & vbTab & lbCounts.List(i, 4)
                    End If
                    If lbCounts.List(i, 5) = "" Then
                        strLine = strLine & vbTab & "<>"
                    Else
                        strLine = strLine & vbTab & lbCounts.List(i, 5)
                    End If
                    If lbCounts.List(i, 6) = "" Then
                        strLine = strLine & vbTab & "<>"
                    Else
                        strLine = strLine & vbTab & lbCounts.List(i, 6)
                    End If
                    If lbCounts.List(i, 7) = "" Then
                        strLine = strLine & vbTab & "<>"
                    Else
                        strLine = strLine & vbTab & lbCounts.List(i, 7)
                    End If
                End If
                
                For j = 0 To UBound(vText)
                    vItem = Split(vText(j), vbTab)
                    If vItem(0) = lbCounts.List(i, 2) Then
                        vText(j) = strLine
                        GoTo Found_Line
                    End If
                Next j
Found_Line:
                
            End If
        Next i
        
        strText = vText(0)
        If UBound(vText) > 0 Then
            For j = 1 To UBound(vText)
                If Not vText(j) = "" Then strText = strText & vbCr & vText(j)
            Next j
        End If
        
    
        Open strFileName For Output As #2
    
        'Print #2, vList(n) '& " " & tbSplitterName.Value
        Print #2, strText
    
        Close #2
Next_N:
    Next n
    
    bSave = False
End Sub

Private Sub SaveF1()
    Dim strFileName As String
    Dim vName, vLine, vItem As Variant
    Dim strLine As String
    Dim fName As String
    Dim iIndex As Integer
    
    vName = Split(ThisDrawing.Name, " ")
    strFileName = strSplitterFilePath & "\" & vName(0) & " Counts -" & tbF1Name.Value & ".mcl"
    
    fName = Dir(strFileName)
    If fName = "" Then
        MsgBox "No F1 File found."
        Exit Sub
    End If
    
    Load CountFileView
    
    Open strFileName For Input As #2
    
    Line Input #2, strLine
    
    While Not EOF(2)
        Line Input #2, strLine
        vLine = Split(strLine, vbTab)
        
        CountFileView.lbCounts.AddItem vLine(0)
        iIndex = CountFileView.lbCounts.ListCount - 1
        
        CountFileView.lbCounts.List(iIndex, 1) = vLine(1)
        CountFileView.lbCounts.List(iIndex, 2) = vLine(2)
        CountFileView.lbCounts.List(iIndex, 3) = vLine(3)
        CountFileView.lbCounts.List(iIndex, 4) = vLine(4)
        CountFileView.lbCounts.List(iIndex, 5) = vLine(5)
    Wend
    
    Close #2
    
    For i = 0 To lbCounts.ListCount - 1
        If lbCounts.List(i, 1) = tbF1Name.Value Then
            iIndex = CInt(lbCounts.List(i, 2)) - 1
            
            CountFileView.lbCounts.List(iIndex, 1) = lbCounts.List(i, 3)
            'CountFileView.show
            CountFileView.lbCounts.List(iIndex, 2) = lbCounts.List(i, 4)
            CountFileView.lbCounts.List(iIndex, 3) = lbCounts.List(i, 5)
            CountFileView.lbCounts.List(iIndex, 4) = lbCounts.List(i, 6)
            CountFileView.lbCounts.List(iIndex, 5) = lbCounts.List(i, 7)
        End If
    Next i
    
    'Me.hide
    CountFileView.show
    'Me.show
    
    Open strFileName For Output As #3
    
    Print #3, tbF1Name.Value & " F1 COUNTS"
    
    For i = 0 To CountFileView.lbCounts.ListCount - 1
        strLine = ""
        
        strLine = CountFileView.lbCounts.List(i, 0)
        If CountFileView.lbCounts.List(i, 1) = " " Then
            strLine = strLine & vbTab & vbTab & vbTab & vbTab & vbTab
            GoTo Next_I
        End If
        
        strLine = strLine & vbTab & CountFileView.lbCounts.List(i, 1)
        strLine = strLine & vbTab & CountFileView.lbCounts.List(i, 2)
        strLine = strLine & vbTab & CountFileView.lbCounts.List(i, 3)
        strLine = strLine & vbTab & CountFileView.lbCounts.List(i, 4)
        strLine = strLine & vbTab & CountFileView.lbCounts.List(i, 5)
        
Next_I:
        Print #3, strLine
    Next i
    
    Close #3
    
    Unload CountFileView
    
    'bSave = False
End Sub

Private Sub GetSpliced(strName As String)
    Dim vName, vFile As Variant
    Dim vLine As Variant
    Dim iIndex As Integer
    
    vName = Split(ThisDrawing.Name, " ")
    strFileName = strSplitterFilePath & "\" & vName(0) & " Counts -" & strName & ".mcl"

    fName = Dir(strFileName)
    If fName = "" Then
        MsgBox strFileName & vbCr & "not found."
        Exit Sub
    End If

    Open strFileName For Input As #1

    vFile = Split(Input$(LOF(1), 1), vbCrLf)
    Close #1
    
    For i = 1 To lbCounts.ListCount - 1
        If lbCounts.List(i, 1) = strName Then
            iIndex = CInt(lbCounts.List(i, 2))
            
            vLine = Split(vFile(iIndex), vbTab)
            lbCounts.List(i, 3) = vLine(1)
            If vLine(4) = "<>" Then lbCounts.List(i, 3) = "<>"
            lbCounts.List(i, 4) = vLine(2)
            lbCounts.List(i, 5) = vLine(3)
            lbCounts.List(i, 6) = vLine(4)
            lbCounts.List(i, 7) = vLine(5)
        End If
    Next i
End Sub

Private Sub GetActivePole()
    Dim objEntity As AcadEntity
    Dim vBasePnt As Variant
    Dim vAttList As Variant
    
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Select Pole: "
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
    Else
        MsgBox "Not a valid object."
        Exit Sub
    End If
    
    Select Case objBlock.Name
        Case "sPed", "sHH", "sPole", "sPanel", "sMH"
        Case Else
            MsgBox "Not a valid block."
            Exit Sub
    End Select
    
    vAttList = objBlock.GetAttributes
    
    If Not tbActivePole.Value = vAttList(0).TextString Then
        strPreviousBlock = tbActivePole.Value
        tbActivePole.Value = vAttList(0).TextString
    End If
End Sub

Private Sub AutoFiber()
    If cbSplitterSize.Value = "" Then Exit Sub
    If cbCableSize.Value = "" Then Exit Sub
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim iSize, iSPL, iDiff As Integer
    
    iSize = CInt(cbCableSize.Value)
    iSPL = CInt(cbSplitterSize.Value)
    iDiff = iSPL
    
    While iDiff > 12
        iDiff = iDiff - 12
    Wend
    
    iDiff = 12 - iDiff
    
    tbFFiber.Value = iSize - iSPL - iDiff + 1
    tbTFiber.Value = iSize - iDiff
End Sub

Private Sub UpdateCable()
    If tbActivePole.Value = "" Then
        MsgBox "Need to select Active Pole."
        Call GetActivePole
    End If
    
    Dim vAttList As Variant
    Dim vLine, vCable, vItem As Variant
    Dim strAttList As String
    Dim strLine, strFind As String
    Dim strCallout, strTemp As String
    Dim strReplace As String
    Dim result, iPosition As Integer
    Dim iAtt As Integer
    
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 25
        Case Else
            iAtt = 5
    End Select
    
    iPosition = -1
    'strReplace = ": " & strPreviousBlock & " + "
    
    'strCallout = Replace(tbCblCallout.Value, vbLf, strReplace)
    strCallout = Replace(tbCblCallout.Value, vbLf, " + ")
    strCallout = Replace(strCallout, vbCr, "")
    
    strLine = tbPosition.Value & ": " & cbCblType.Value & "(" & cbCableSize.Value & ")"
    If Not cbSuffix.Value = "" Then strLine = strLine & cbSuffix.Value
    strLine = strLine & " / " & strCallout
    
    strCableCallout = strLine
    
    strFind = ""
    vAttList = objBlock.GetAttributes
    
    
    If vAttList(iAtt).TextString = "" Then
        vAttList(iAtt).TextString = strLine
    Else
        strTemp = Replace(vAttList(iAtt).TextString, vbLf, "")
        vLine = Split(strTemp, vbCr)
        For i = 0 To UBound(vLine)
            vCable = Split(vLine(0), ": ")
            If vCable(0) = tbPosition.Value Then
                iPosition = i
            End If
        Next i
        
        Select Case iPosition
            Case Is < 0
                vAttList(iAtt).TextString = vAttList(iAtt).TextString & vbCr & strLine
            Case Else
                strFind = vLine(iPosition)
                vAttList(iAtt).TextString = Replace(vAttList(iAtt).TextString, strFind, strLine)
        End Select
        
    End If
    
    objBlock.Update
    
    bUpdatePole = False
End Sub

Private Sub UpdateTerminal()
    If tbCallout.Value = "" Then Exit Sub
    
    Dim strLine As String
    
    strLine = Replace(tbCallout.Value, vbTab, "")
    strLine = Replace(strLine, vbCr, "")
    strLine = Replace(strLine, vbLf, "")
    strLine = Replace(strLine, vbCrLf, "")
    strLine = Replace(strLine, " ", "")
    
    If Len(strLine) < 3 Then Exit Sub
    
    strLine = tbCallout.Value
    
    If tbActivePole.Value = "" Then
        MsgBox "Need to select Active Pole."
        Call GetActivePole
    End If
    
    Dim vAttList, vLine, vItem As Variant
    Dim strPosition, strSpliced As String
    Dim iAtt As Integer
    
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 26
        Case Else
            iAtt = 6
    End Select
    
    vAttList = objBlock.GetAttributes
    strSpliced = Replace(vAttList(iAtt).TextString, vbLf, "")
    
    strPosition = "[" & tbPosition.Value & "]"
    
    strLine = Replace(tbCallout.Value, vbCr, "")
    vLine = Split(strLine, vbLf)
    
    Dim strText As String
    Dim iTest As Integer
    iTest = 0
    For i = 0 To UBound(vLine)
        If Not vLine(i) = "" Then
            iTest = 1
            If strText = "" Then
                strText = vLine(i)
            Else
                strText = strText & " + " & vLine(i)
            End If
        End If
    Next i
    If iTest = 0 Then Exit Sub
    
    strLine = strPosition & " " & strText
    
    'If UBound(vLine) > 0 Then
        'For i = 1 To UBound(vLine)
            'strLine = strLine & " + " & vLine(i)
        'Next i
    'End If
    
    If strSpliced = "" Then
        vAttList(iAtt).TextString = strLine
    Else
        vLine = Split(strSpliced, vbCr)
        
        For i = 0 To UBound(vLine)
            If InStr(vLine(i), strPosition) > 0 Then
                vLine(i) = strLine
                GoTo Replaced_Position
            End If
        Next i
        
        vLine(UBound(vLine)) = vLine(UBound(vLine)) & vbCr & strLine
        
Replaced_Position:
        
        strSpliced = vLine(0)
        
        If UBound(vLine) > 0 Then
            For i = 1 To UBound(vLine)
                strSpliced = strSpliced & vbCr & vLine(i)
            Next i
        End If
        
        vAttList(iAtt).TextString = strSpliced
    End If
    
    objBlock.Update
    
    bUpdatePole = False
End Sub

Private Sub UpdateTerminalCallout()
    Dim objSSTemp4 As AcadSelectionSet
    Dim entItem As AcadEntity
    Dim objTemp As AcadBlockReference
    Dim vAttList As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    
    Dim strSearch As String
    Dim strAtt0, strAtt1, strAtt2 As String
    Dim strLayer, strCounts, strTemp As String
    Dim iHO1 As Integer
    
    strSearch = tbActivePole.Value & ": " & tbPosition.Value

    grpCode(0) = 2
    grpValue(0) = "Callout"
    filterType = grpCode
    filterValue = grpValue
    
  On Error Resume Next
    
    Set objSSTemp4 = ThisDrawing.SelectionSets.Add("objSSTemp4")
    If Not Err = 0 Then
        Set objSSTemp4 = ThisDrawing.SelectionSets.Item("objSSTemp4")
        Err = 0
    End If
    
    objSSTemp4.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objTemp In objSSTemp4
        vAttList = objTemp.GetAttributes
        
        If vAttList(0).TextString = strSearch Then
            If InStr(vAttList(1).TextString, "+HO1") > 0 Then
                
                'strAtt0 = tbActivePole.Value & ": " & tbPosition.Value
    
                iHO1 = 0
                strCounts = Replace(tbCallout.Value, vbLf, "")
                vCounts = Split(strCounts, vbCr)
                strCounts = ""
    
                For i = 0 To UBound(vCounts)
                    vLine = Split(vCounts(i), ": ")
                    vItem = Split(vLine(1), "-")
                    
                    strTemp = vLine(0) & ": " & vLine(1)
                    If strCounts = "" Then
                        strCounts = strTemp
                    Else
                        strCounts = strCounts & "\P" & strTemp
                    End If
        
                    If UBound(vItem) = 0 Then
                        iHO1 = iHO1 + 1
                    Else
                        iHO1 = iHO1 + CInt(vItem(1)) - CInt(vItem(0)) + 1
                    End If
                Next i
    
                If cbCblType.Value = "CO" Then
                    strAtt1 = "+HO1A=" & iHO1
                    strLayer = "Integrity Proposed-Aerial"
                Else
                    strAtt1 = "+HO1B=" & iHO1
                    strLayer = "Integrity Proposed-Buried"
                End If
    
                'strAtt2 = Replace(tbCallout.Value, vbCr, "\P")
                'strAtt2 = Replace(strAtt2, vbLf, "")
                
                
                'vAttList(0).TextString = strAtt0
                vAttList(1).TextString = strAtt1
                vAttList(2).TextString = strCounts
                vAttList(2).TextString = strAtt2
            
                objTemp.Update
                
                Call UpdateTerminal
                
                GoTo Exit_Sub
            End If
        End If
    Next objTemp
    
    'MsgBox "Callout not found."
    
Exit_Sub:
    objSSTemp4.Clear
    objSSTemp4.Delete
End Sub

Private Sub UpdateCableCallout()
    Dim objSSTemp4 As AcadSelectionSet
    Dim entItem As AcadEntity
    Dim objTemp As AcadBlockReference
    Dim vAttList, vPoleAtt As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    
    Dim vCounts, vLine As Variant
    
    Dim strSearch As String
    Dim strAtt0, strAtt1, strAtt2 As String
    Dim strLayer, strCounts, strTemp As String
    Dim iHO1 As Integer
    
    strSearch = tbActivePole.Value & ": " & tbPosition.Value

    grpCode(0) = 2
    grpValue(0) = "Callout"
    filterType = grpCode
    filterValue = grpValue
    
  On Error Resume Next
    
    Set objSSTemp4 = ThisDrawing.SelectionSets.Add("objSSTemp4")
    If Not Err = 0 Then
        Set objSSTemp4 = ThisDrawing.SelectionSets.Item("objSSTemp4")
        Err = 0
    End If
    
    objSSTemp4.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objTemp In objSSTemp4
        vAttList = objTemp.GetAttributes
        
        If vAttList(0).TextString = strSearch Then
            If InStr(vAttList(1).TextString, "+HO1") < 1 Then
                'strAtt1 = cbCblType.Value & "(" & cbCableSize.Value & ")"
                'If Not cbSuffix.Value = "" Then strAtt1 = strAtt1 & cbSuffix.Value
                'strAtt2 = Replace(tbCblCallout.Value, vbCr, "\P")
                'strAtt2 = Replace(strAtt2, vbLf, "")
                
                'MsgBox strCableCallout
                'GoTo Exit_Sub
                
                vPoleAtt = Split(strCableCallout, " / ")
                
                vCounts = Split(vPoleAtt(0), ": ")
                strAtt1 = vCounts(1)
                
    
                strCounts = ""
                vCounts = Split(vPoleAtt(1), " + ")
                For i = 0 To UBound(vCounts)
                    If vCounts(i) = "" Then GoTo Next_line
                    
                    vLine = Split(vCounts(i), ": ")
                    strTemp = vLine(0) & ": " & vLine(1)
                    If strCounts = "" Then
                        strCounts = strTemp
                    Else
                        strCounts = strCounts & "\P" & strTemp
                    End If
Next_line:
                Next i
                
                vAttList(1).TextString = strAtt1
                vAttList(2).TextString = strCounts
            
                objTemp.Update
                
                'GoTo Exit_Sub
            End If
        End If
    Next objTemp
    
    'MsgBox "Callout not found."
    
Exit_Sub:
    objSSTemp4.Clear
    objSSTemp4.Delete
End Sub

Private Sub AddSplitter()
    Dim objSplitter, objTemp As AcadBlockReference
    Dim objEntity As AcadEntity
    Dim vAttList, vReturnPnt As Variant
    
    Dim lwpCoords(0 To 3) As Double
    Dim vBlockCoords(0 To 2) As Double
    Dim dPrevious(0 To 2) As Double
    Dim lineObj As AcadLWPolyline
    Dim strLayer As String
    Dim dScale As Double
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Closure: "
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objTemp = objEntity
        
        lwpCoords(0) = objTemp.InsertionPoint(0)
        lwpCoords(1) = objTemp.InsertionPoint(1)
    Else
        lwpCoords(0) = vReturnPnt(0)
        lwpCoords(1) = vReturnPnt(1)
    End If
    
    dPrevious(0) = lwpCoords(0)
    dPrevious(1) = lwpCoords(1)
    dPrevious(2) = 0#
    
    vReturnPnt = ThisDrawing.Utility.GetPoint(dPrevious, "Select Point: ")
    vBlockCoords(0) = vReturnPnt(0)
    vBlockCoords(1) = vReturnPnt(1)
    vBlockCoords(2) = vReturnPnt(2)
    lwpCoords(2) = vBlockCoords(0)
    lwpCoords(3) = vBlockCoords(1)
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(lwpCoords)
    
    If cbCblType.Value = "CO" Then
        strLayer = "Integrity Proposed-Aerial"
    Else
        strLayer = "Integrity Proposed-Buried"
    End If
    lineObj.Layer = strLayer
    lineObj.Update
    
    dScale = CDbl(cbScale.Value) / 100
    If dScale = 0 Then dScale = 0.75
        
    Set objSplitter = ThisDrawing.ModelSpace.InsertBlock(vBlockCoords, "iSplitter", dScale, dScale, dScale, 0)
    vAttItem = objSplitter.GetAttributes
    
    vAttItem(0).TextString = CountsSplitter.cbSplitterSize.Value
    vAttItem(1).TextString = CountsSplitter.tbF2Name.Value & " " & CountsSplitter.tbSplitterName.Value
    
    objSplitter.Layer = strLayer
    
    'If lwpCoords(2) < lwpCoords(0) Then
        'vBlockCoords(0) = vBlockCoords(0) - (75 * dScale)
        'objSplitter.InsertionPoint = vBlockCoords
    'End If
    objSplitter.Update
    
    Me.show
End Sub

Private Sub AddTapToPole(strLine As String)
    Dim iBF, iEF, iBC, iEC, iTemp, iCable, iLine As Integer
    Dim dTemp As Double
    Dim vLine, vCount As Variant
    Dim strCounts, strFibers, strXD As String
    
    Me.Hide
    
    vLine = Split(strLine, ": ")
    vCount = Split(vLine(1), "-")
    
    iBC = CInt(vCount(0))
    If UBound(vCount) = 0 Then
        iEC = iBC
    Else
        iEC = CInt(vCount(1))
    End If
    
    iBF = iBC
    iEF = iEC
    
    While iBF > 12
        iBF = iBF - 12
        iEF = iEF - 12
    Wend
    
    'iEF = iEC - (iBF - iBC)
    
    strFibers = iBF & "-" & iEF
    strCounts = vLine(0) & ": " & iBC & "-" & iEC
    
    Load NewCableForm
        NewCableForm.Caption = "Tap Cable"
        NewCableForm.cbCblType.Value = "CO"
        NewCableForm.cbSuffix.Value = "6M-EHS"
        NewCableForm.tbPosition.Value = "A1"
        
        Select Case iEF
            Case Is < 13
                iCable = 12
            Case Is < 25
                iCable = 24
            Case Is < 49
                iCable = 48
            Case Is < 73
                iCable = 72
            Case Is < 97
                iCable = 96
            Case Is < 145
                iCable = 144
            Case Else
                iCable = 0
        End Select
        
        NewCableForm.cbCableSize.Value = iCable
        'NewCableForm.cbCblName.AddItem vLine(0)
            
        iLine = 0
        
        If iBF > 1 Then
            NewCableForm.lbCounts.AddItem "1-" & (iBF - 1)
            strXD = "XD: 1-" & (iBF - 1)
            NewCableForm.lbCounts.List(iLine, 1) = strXD
            iLine = iLine + 1
        End If
        
        NewCableForm.lbCounts.AddItem strFibers
        NewCableForm.lbCounts.List(iLine, 1) = strCounts
        iLine = iLine + 1
        
        If iEF < iCable Then
            NewCableForm.lbCounts.AddItem (iEF + 1) & "-" & iCable
            strXD = "XD: " & (iEF + 1) & "-" & iCable
            NewCableForm.lbCounts.List(iLine, 1) = strXD
            iLine = iLine + 1
        End If
        
        NewCableForm.show
        
    Unload NewCableForm
    
    Me.show
End Sub

