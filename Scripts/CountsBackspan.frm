VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CountsBackspan 
   Caption         =   "Backspan Counts"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "CountsBackspan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CountsBackspan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objCurrent As AcadBlockReference
Dim iSave, iSavePole As Integer

Private Sub cbAddToPole_Click()
    Call UpdatePole
End Sub

Private Sub cbDrop_Click()
    If lbBackspan.ListCount < 1 Then Exit Sub
    
    Dim strCable, strCounts, strSpliced, strCallout As String
    Dim vLine, vItem As Variant
    Dim iIndex As Integer
    
    Load CountsTap
        'CountsTap.lbMain.ColumnCount = 4
        'CountsTap.lbMain.ColumnWidths = "24;72;36;30"
        
        CountsTap.tbStructure.Value = tbStructure.Value
        
        For i = 0 To lbBackspan.ListCount - 1
            CountsTap.lbMain.AddItem lbBackspan.List(i, 0)
            
            If lbBackspan.List(i, 3) = "<>" Then
                CountsTap.lbMain.List(i, 1) = lbBackspan.List(i, 1)
                CountsTap.lbMain.List(i, 2) = lbBackspan.List(i, 2)
            Else
                CountsTap.lbMain.List(i, 1) = "XD"
                CountsTap.lbMain.List(i, 2) = i
            End If
            CountsTap.lbMain.List(i, 3) = ""
        Next i
        
        CountsTap.show
        
        If CountsTap.cbChanged.Value = False Then GoTo Exit_Sub
        
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
        
        strCallout = Replace(vLine(1), " + ", vbCr)
        
        For i = 0 To CountsTap.lbMain.ListCount - 1
            If Left(CountsTap.lbMain.List(i, 3), 1) = "Y" Then
                lbBackspan.List(i, 3) = tbStructure.Value
                lbBackspan.List(i, 6) = "TAP"
                
                iIndex = CInt(lbBackspan.List(i, 0))
                lbCounts.List(iIndex, 3) = tbStructure.Value
                lbCounts.List(iIndex, 6) = "TAP"
            End If
        Next i
        
        If tbTerminal.Value = "" Then
            tbTerminal.Value = strCallout
        Else
           tbTerminal.Value = tbTerminal.Value & vbCr & strCallout
        End If
        
        Dim objEntity As AcadEntity
        Dim objBlock As AcadBlockReference
        Dim vReturnPnt, vAttList As Variant
        Dim iAtt As Integer
        
        Me.Hide
        
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
        
        'Call CreateCallout
        
        iSave = 1
        iSavePole = 1
        
        Me.show
Exit_Sub:
    Unload CountsTap
End Sub

Private Sub cbGetPole_Click()
    If cbGetPole.Caption = "Get Start Pole" Then
        Call GetInfo
        Exit Sub
    End If
    
    Dim objEntity As AcadEntity
    Dim vBasePnt As Variant
    Dim vAttList As Variant
    Dim vLine, vItem, vTemp, vCount As Variant
    Dim strLine, strTerm, strTemp As String
    Dim result As Integer
    Dim iSFiber, iEFiber As Integer
    Dim iStart, iEnd, iIndex As Integer
    
    If lbBackspan.ListCount < 1 Then
        strLine = "No backspan fibers added to list." & vbCr & "They are needed to know which fibers to move back."
        strLine = strLine & vbCr & "Double click the fibers in the Current Structure Counts list to add them."
        MsgBox strLine
        Exit Sub
    End If
    
    If iSavePole > 0 Then
        result = MsgBox("Update changes to Structure attributes?", vbYesNo, "Update Changes")
        If result = vbYes Then
            Call UpdatePole
        End If
    End If
    
  On Error Resume Next
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Select Pole: "
    If TypeOf objEntity Is AcadBlockReference Then
        Set objCurrent = objEntity
    Else
        MsgBox "Not a valid object."
        Exit Sub
    End If
    
    lbCurrentCounts.Clear
    
    Select Case objCurrent.Name
        Case "sPole"
            vAttList = objCurrent.GetAttributes
            strLine = vAttList(25).TextString
            strTerm = vAttList(26).TextString
        Case "sPed", "sHH", "sPanel", "sMH"
            vAttList = objCurrent.GetAttributes
            strLine = vAttList(5).TextString
            strTerm = vAttList(6).TextString
        Case Else
            MsgBox "Not a valid block."
            Exit Sub
    End Select
    
    tbStructure.Value = vAttList(0).TextString
    
    strTerm = Replace(strTerm, " + ", vbCr)
    strTerm = Replace(strTerm, "] ", "]" & vbCr)
    tbTerminal.Value = strTerm
    
    vLine = Split(strLine, " / ")
    tbCableType.Value = vLine(0)
    If UBound(vLine) > 0 Then
        vItem = Split(vLine(1), " + ")
        iSFiber = 1
        iEFiber = 1
        vTemp = Split(vItem(0), ": ")
        vCount = Split(vTemp(1), "-")
        If UBound(vCount) = 0 Then
            iEFiber = 1
        Else
            iEFiber = iSFiber + CInt(vCount(1)) - CInt(vCount(0))
        End If
        
        lbCurrentCounts.AddItem iSFiber & "-" & iEFiber
        lbCurrentCounts.List(0, 1) = vTemp(0)
        lbCurrentCounts.List(0, 2) = vTemp(1)
        lbCurrentCounts.List(0, 3) = vTemp(2)
        
        iSFiber = iEFiber + 1
        If UBound(vItem) > 0 Then
            For i = 1 To UBound(vItem)
                vTemp = Split(vItem(i), ": ")
                vCount = Split(vTemp(1), "-")
                If UBound(vCount) > 0 Then
                    iEFiber = iSFiber + CInt(vCount(1)) - CInt(vCount(0))
                End If
        
                lbCurrentCounts.AddItem iSFiber & "-" & iEFiber
                lbCurrentCounts.List(lbCurrentCounts.ListCount - 1, 1) = vTemp(0)
                lbCurrentCounts.List(lbCurrentCounts.ListCount - 1, 2) = vTemp(1)
                lbCurrentCounts.List(lbCurrentCounts.ListCount - 1, 3) = vTemp(2)
                
                iSFiber = iEFiber + 1
            Next i
        End If
    End If
    
    lbCounts.Clear
    lbCounts.AddItem ""
    
    If lbCurrentCounts.ListCount > 0 Then
        For i = 0 To lbCurrentCounts.ListCount - 1
            vLine = Split(lbCurrentCounts.List(i, 0), "-")
            iFStart = CInt(vLine(0))
            If UBound(vLine) > 0 Then
                iFEnd = CInt(vLine(1))
            Else
                iFEnd = iFStart
            End If
            
            vLine = Split(lbCurrentCounts.List(i, 2), "-")
            iStart = CInt(vLine(0)) - iFStart
            
            For j = iFStart To iFEnd
                lbCounts.AddItem j
                lbCounts.List(j, 1) = lbCurrentCounts.List(i, 1)
                lbCounts.List(j, 2) = iStart + j
                lbCounts.List(j, 3) = "<>"
                lbCounts.List(j, 4) = "<>"
                lbCounts.List(j, 5) = "<>"
                lbCounts.List(j, 6) = "<>"
                lbCounts.List(j, 7) = "<>"
            Next j
            
        Next i
    End If
    
    If lbBackspan.ListCount > 0 Then
        For i = 0 To lbBackspan.ListCount - 1
            If lbBackspan.List(i, 3) = "<>" Then
                iStart = CInt(lbBackspan.List(i, 0))
                
                lbCounts.List(iStart, 1) = lbBackspan.List(i, 1)
                lbCounts.List(iStart, 2) = lbBackspan.List(i, 2)
            End If
        Next i
    
        Call GetTapCallout
    End If
    
    'If lbCurrentCounts.ListCount > 0 Then
        'For i = 0 To lbCurrentCounts.ListCount - 1
            'lbProposedCounts.AddItem lbCurrentCounts.List(i, 0)
            'lbProposedCounts.List(i, 1) = lbCurrentCounts.List(i, 1)
            'lbProposedCounts.List(i, 2) = lbCurrentCounts.List(i, 2)
            'lbProposedCounts.List(i, 3) = lbCurrentCounts.List(i, 3)
        'Next i
    'End If
    
Exit_Sub:
    If cbCutCounts.Enabled = False Then
        cbCutCounts.Enabled = True
        cbDrop.Enabled = True
        cbTerminal.Enabled = True
        cbRemoveWL.Enabled = True
        cbSaveSplitter.Enabled = True
        cbAddToPole.Enabled = True
    End If
    
    Me.show
End Sub

Private Sub cbRemoveWL_Click()
    
    iSave = 1
    iSavePole = 1
End Sub

Private Sub cbSaveSplitter_Click()
    Call SaveMCL
End Sub

Private Sub cbTerminal_Click()
    'If objCurrent Is Nothing Then
        'MsgBox "No Pole Selected"
        'Exit Sub
    'End If
    
    Dim iSelected As Integer
    
    For i = 0 To lbBackspan.ListCount - 1
        If lbBackspan.Selected(i) = True Then
            iSelected = i
            GoTo Found_Selected
        End If
    Next i
    
    MsgBox "No Count Selected."
    Exit Sub
    
Found_Selected:
    Dim result As Integer
    
    If Not lbBackspan.List(i, 3) = "<>" Then
        
        result = MsgBox("Overwrite previous assignment?", vbYesNo, "Fiber Assigned!")
        If result = vbNo Then Exit Sub
    End If
    
    Dim objEntity As AcadEntity
    Dim objRES As AcadBlockReference
    Dim vAttList, vBasePnt As Variant
    Dim strLine As String
    Dim strName As String
    Dim iBCount, iECount As Integer
    Dim iCountIndex As Integer
    
    iBCount = 0: iECount = 0
    
    Me.Hide
    On Error Resume Next
    
Next_One:
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Select Building: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not lbBackspan.List(iSelected, 3) = "<>" Then
        
        result = MsgBox("Overwrite previous assignment?", vbYesNo, "Fiber Assigned!")
        If result = vbNo Then Exit Sub
    End If
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objRES = objEntity
        vAttList = objRES.GetAttributes
            
        Select Case objRES.Name
            'Case "RES", "BUSINESS", "MDU", "TRLR", "SCHOOL", "CHURCH", "LOT"
            Case "SG"
                lbBackspan.List(iSelected, 3) = tbStructure.Value
                lbBackspan.List(iSelected, 4) = "SG"
                lbBackspan.List(iSelected, 5) = vAttList(1).TextString
                lbBackspan.List(iSelected, 6) = "SMARTGRID"
                lbBackspan.List(iSelected, 7) = vAttList(0).TextString
                
                strLine = lbBackspan.List(iSelected, 1) & ": " & lbBackspan.List(iSelected, 2)
                'If cbDrop.Value = False Then strline = "(" & strline & ")"
                vAttList(2).TextString = tbStructure.Value & " - " & strLine
                objRES.Update
                
                strName = lbBackspan.List(iSelected, 1)
                If iECount = 0 Then
                    iECount = CInt(lbBackspan.List(iSelected, 2))
                Else
                    iBCount = CInt(lbBackspan.List(iSelected, 2))
                End If
                
                
            Case "Customer"
                lbBackspan.List(iSelected, 3) = tbStructure.Value
                lbBackspan.List(iSelected, 4) = vAttList(1).TextString
                lbBackspan.List(iSelected, 5) = vAttList(2).TextString
                lbBackspan.List(iSelected, 6) = vAttList(0).TextString
                If vAttList(3).TextString = "" Then
                    lbBackspan.List(iSelected, 7) = "<>"
                Else
                    lbBackspan.List(iSelected, 7) = vAttList(3).TextString
                End If
                
                strLine = lbBackspan.List(iSelected, 1) & ": " & lbBackspan.List(iSelected, 2)
                If cbDrop.Value = False Then strLine = "(" & strLine & ")"
                vAttList(4).TextString = tbStructure.Value & " - " & strLine
                objRES.Update
                
                strName = lbBackspan.List(iSelected, 1)
                If iECount = 0 Then
                    iECount = CInt(lbBackspan.List(iSelected, 2))
                Else
                    iBCount = CInt(lbBackspan.List(iSelected, 2))
                End If
        End Select
    End If
    
    iCountIndex = CInt(lbBackspan.List(iSelected, 0))
    
    lbCounts.List(iCountIndex, 3) = lbBackspan.List(iSelected, 3)
    lbCounts.List(iCountIndex, 4) = lbBackspan.List(iSelected, 4)
    lbCounts.List(iCountIndex, 5) = lbBackspan.List(iSelected, 5)
    lbCounts.List(iCountIndex, 6) = lbBackspan.List(iSelected, 6)
    lbCounts.List(iCountIndex, 7) = lbBackspan.List(iSelected, 7)
    
    iSelected = iSelected - 1
    
    GoTo Next_One
    
Exit_Sub:
    lbBackspan.ListIndex = iSelected
    
    If iECount > 0 Then
        If iBCount = 0 Then
            strName = strName & ": " & iECount
        Else
            strName = strName & ": " & iBCount & "-" & iECount
        End If
        
        strName = strName & ": " & tbStructure.Value
        
        If tbTerminal.Value = "" Then
            tbTerminal.Value = strName
        Else
            tbTerminal.Value = tbTerminal.Value & vbCr & strName
        End If
    End If
    
    'Call UpdateCountsList
    Call GetTapCallout
    
    iSave = 1
    iSavePole = 1
    
    Me.show
End Sub

Private Sub lbCurrentCounts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vLine As Variant
    Dim iStart, iEnd As Integer
    Dim iIndex As Integer
    
    iIndex = lbCurrentCounts.ListIndex
    
    vLine = Split(lbCurrentCounts.List(iIndex, 0), "-")
    iStart = CInt(vLine(0))
    If UBound(vLine) > 0 Then
        iEnd = CInt(vLine(1))
    Else
        iEnd = iStart
    End If
    
    For i = iStart To iEnd
        lbBackspan.AddItem i
        lbBackspan.List(lbBackspan.ListCount - 1, 1) = lbCounts.List(i, 1)
        lbBackspan.List(lbBackspan.ListCount - 1, 2) = lbCounts.List(i, 2)
        lbBackspan.List(lbBackspan.ListCount - 1, 3) = lbCounts.List(i, 3)
        lbBackspan.List(lbBackspan.ListCount - 1, 4) = lbCounts.List(i, 4)
        lbBackspan.List(lbBackspan.ListCount - 1, 5) = lbCounts.List(i, 5)
        lbBackspan.List(lbBackspan.ListCount - 1, 6) = lbCounts.List(i, 6)
        lbBackspan.List(lbBackspan.ListCount - 1, 7) = lbCounts.List(i, 7)
    Next i
    
    'tbFibers.Enabled = True
End Sub

Private Sub tbF2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strFileName As String
    Dim vName, vLine, vItem As Variant
    Dim strLine As String
    Dim fName As String
    Dim iIndex As Integer
    
    vName = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vName(0) & " Counts -" & tbF2.Value & ".mcl"
    
    fName = Dir(strFileName)
    If fName = "" Then
        MsgBox "No MCL file."
        Exit Sub
    End If
    
    'Load CountFileView
    
    Open strFileName For Input As #2
    
    Line Input #2, strLine
    
    While Not EOF(2)
        Line Input #2, strLine
        vLine = Split(strLine, vbTab)
        
        lbCounts.AddItem vLine(0)
        iIndex = lbCounts.ListCount - 1
        
        lbCounts.List(iIndex, 1) = tbF2.Value
        lbCounts.List(iIndex, 2) = vLine(0)
        
        lbCounts.List(iIndex, 3) = vLine(1)
        lbCounts.List(iIndex, 4) = vLine(2)
        
        If vLine(4) = "TAP" Then
            lbCounts.List(iIndex, 5) = "<>"
            lbCounts.List(iIndex, 6) = "<>"
            lbCounts.List(iIndex, 7) = "<>"
        Else
            lbCounts.List(iIndex, 5) = vLine(3)
            lbCounts.List(iIndex, 6) = vLine(4)
            lbCounts.List(iIndex, 7) = vLine(5)
        End If
    Wend
    
    Close #2
    
    'For i = 0 To lbCounts.ListCount - 1
        'If lbCounts.List(i, 1) = tbF2Name.Value Then
            'iIndex = CInt(lbCounts.List(i, 2)) - 1
            
            'lbCounts.List(iIndex, 1) = lbCounts.List(i, 3)
            ''CountFileView.show
            'lbCounts.List(iIndex, 2) = lbCounts.List(i, 4)
            'lbCounts.List(iIndex, 3) = lbCounts.List(i, 5)
            'lbCounts.List(iIndex, 4) = lbCounts.List(i, 6)
            'lbCounts.List(iIndex, 5) = lbCounts.List(i, 7)
        'End If
    'Next i
    
    
    
    
    'Open strFileName For Output As #3
    
    'Print #3, tbF2Name.Value & " " & tbSplitterName.Value
    
    'For i = 0 To lbCounts.ListCount - 1
        'strLine = ""
        
        'strLine = lbCounts.List(i, 0)
        'If lbCounts.List(i, 1) = " " Then
            'strLine = strLine & vbTab & vbTab & vbTab & vbTab & vbTab
            'GoTo Next_I
        'End If
        
        'strLine = strLine & vbTab & lbCounts.List(i, 1)
        'strLine = strLine & vbTab & lbCounts.List(i, 2)
        'strLine = strLine & vbTab & lbCounts.List(i, 3)
        'strLine = strLine & vbTab & lbCounts.List(i, 4)
        'strLine = strLine & vbTab & lbCounts.List(i, 5)
        
'Next_I:
        'Print #3, strLine
    'Next i
    
    'Close #3
End Sub

Private Sub UserForm_Initialize()
    lbCounts.Clear
    lbCounts.ColumnCount = 8
    lbCounts.ColumnWidths = "24;72;30;66;36;136;48;96"
    
    lbBackspan.Clear
    lbBackspan.ColumnCount = 8
    lbBackspan.ColumnWidths = "24;72;30;66;36;136;48;96"
    
    lbCurrentCounts.Clear
    lbCurrentCounts.ColumnCount = 4
    lbCurrentCounts.ColumnWidths = "48;72;42;6"
    
    lbProposedCounts.Clear
    lbProposedCounts.ColumnCount = 4
    lbProposedCounts.ColumnWidths = "48;72;42;6"
    
    iSave = 0
    iSavePole = 0
End Sub

Private Sub GetTapCallout()
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim strLine, strItem As String
    Dim strCurrent, strPrevious As String
    Dim iStart, iEnd As Integer
    Dim iFStart, iFEnd As Integer
    Dim iIndex As Integer
    
    lbCurrentCounts.Clear
    
    strPrevious = lbCounts.List(1, 1)
    If Not lbCounts.List(1, 3) = "<>" Then strPrevious = "XD"
    strCurrent = strPrevious
    iStart = CInt(lbCounts.List(1, 2))
    iEnd = iStart
    iFStart = 1
    iFEnd = 1
    
    For i = 2 To lbCounts.ListCount - 1
        strCurrent = lbCounts.List(i, 1)
        If Not lbCounts.List(i, 3) = "<>" Then strCurrent = "XD"
        
        If strCurrent = strPrevious Then
            iEnd = CInt(lbCounts.List(i, 2))
            iFEnd = iFEnd + 1
        Else
            strLine = iFStart
            If iFEnd > iFStart Then strLine = strLine & "-" & iFEnd
            strItem = iStart
            If iEnd > iStart Then strItem = strItem & "-" & iEnd
            
            lbCurrentCounts.AddItem strLine
            iIndex = lbCurrentCounts.ListCount - 1
            lbCurrentCounts.List(iIndex, 1) = strPrevious
            lbCurrentCounts.List(iIndex, 2) = strItem
            lbCurrentCounts.List(iIndex, 3) = tbStructure.Value
             
            strPrevious = strCurrent
            iStart = CInt(lbCounts.List(i, 2))
            iEnd = iStart
            iFStart = iFEnd + 1
            iFEnd = iFStart
        End If
    Next i
    
    strLine = iFStart
    If iFEnd > iFStart Then strLine = strLine & "-" & iFEnd
    strItem = iStart
    If iEnd > iStart Then strItem = strItem & "-" & iEnd
            
    lbCurrentCounts.AddItem strLine
    iIndex = lbCurrentCounts.ListCount - 1
    lbCurrentCounts.List(iIndex, 1) = strPrevious
    lbCurrentCounts.List(iIndex, 2) = strItem
    lbCurrentCounts.List(iIndex, 3) = tbStructure.Value
End Sub

Private Sub GetInfo()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vBasePnt As Variant
    Dim vAttList As Variant
    Dim strLine, strTerm, strTemp As String
    Dim vLine, vItem, vTemp, vCount As Variant
    Dim iSFiber, iEFiber As Integer
    Dim iStart, iEnd, iIndex As Integer
    
  On Error Resume Next
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Select Pole: "
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
    Else
        MsgBox "Not a valid object."
        Exit Sub
    End If
    
    lbCurrentCounts.Clear
    
    Select Case objBlock.Name
        Case "sPole"
            vAttList = objBlock.GetAttributes
            strLine = vAttList(25).TextString
            strTerm = vAttList(26).TextString
        Case "sPed", "sHH", "sPanel", "sMH"
            vAttList = objBlock.GetAttributes
            strLine = vAttList(5).TextString
            strTerm = vAttList(6).TextString
        Case Else
            MsgBox "Not a valid block."
            Exit Sub
    End Select
    
    'tbStructure.Value = vAttList(0).TextString
    
    'strTerm = Replace(strTerm, " + ", vbCr)
    'strTerm = Replace(strTerm, "] ", "]" & vbCr)
    'tbTerminal.Value = strTerm
    
    vLine = Split(strLine, " / ")
    'tbCableType.Value = vLine(0)
    If UBound(vLine) > 0 Then
        vItem = Split(vLine(1), " + ")
        iSFiber = 1
        iEFiber = 1
        vTemp = Split(vItem(0), ": ")
        vCount = Split(vTemp(1), "-")
        If UBound(vCount) = 0 Then
            iEFiber = 1
        Else
            iEFiber = iSFiber + CInt(vCount(1)) - CInt(vCount(0))
        End If
        
        lbCurrentCounts.AddItem iSFiber & "-" & iEFiber
        lbCurrentCounts.List(0, 1) = vTemp(0)
        lbCurrentCounts.List(0, 2) = vTemp(1)
        lbCurrentCounts.List(0, 3) = vTemp(2)
        
        iSFiber = iEFiber + 1
        If UBound(vItem) > 0 Then
            For i = 1 To UBound(vItem)
                vTemp = Split(vItem(i), ": ")
                vCount = Split(vTemp(1), "-")
                If UBound(vCount) > 0 Then
                    iEFiber = iSFiber + CInt(vCount(1)) - CInt(vCount(0))
                End If
        
                lbCurrentCounts.AddItem iSFiber & "-" & iEFiber
                lbCurrentCounts.List(lbCurrentCounts.ListCount - 1, 1) = vTemp(0)
                lbCurrentCounts.List(lbCurrentCounts.ListCount - 1, 2) = vTemp(1)
                lbCurrentCounts.List(lbCurrentCounts.ListCount - 1, 3) = vTemp(2)
                
                iSFiber = iEFiber + 1
            Next i
        End If
    End If
    
    lbCounts.Clear
    lbCounts.AddItem ""
    
    If lbCurrentCounts.ListCount > 0 Then
        For i = 0 To lbCurrentCounts.ListCount - 1
            vLine = Split(lbCurrentCounts.List(i, 0), "-")
            iFStart = CInt(vLine(0))
            If UBound(vLine) > 0 Then
                iFEnd = CInt(vLine(1))
            Else
                iFEnd = iFStart
            End If
            
            vLine = Split(lbCurrentCounts.List(i, 2), "-")
            iStart = CInt(vLine(0)) - iFStart
            
            For j = iFStart To iFEnd
                lbCounts.AddItem j
                lbCounts.List(j, 1) = lbCurrentCounts.List(i, 1)
                lbCounts.List(j, 2) = iStart + j
                lbCounts.List(j, 3) = "<>"
                lbCounts.List(j, 4) = "<>"
                lbCounts.List(j, 5) = "<>"
                lbCounts.List(j, 6) = "<>"
                lbCounts.List(j, 7) = "<>"
            Next j
            
        Next i
    End If
    
    cbGetPole.Caption = "Get Backspan Pole"
    
    Me.show
End Sub

Private Sub UserForm_Terminate()
    Dim result As Integer
    
    If iSavePole > 0 Then
        result = MsgBox("Update changes to Structure attributes?", vbYesNo, "Update Changes")
        If result = vbYes Then
            Call UpdatePole
        End If
    End If
    
    If iSave > 0 Then
        result = MsgBox("Save changes to Master List?", vbYesNo, "Save Changes")
        If result = vbYes Then
            Call SaveMCL
        End If
    End If
End Sub

Private Sub SaveMCL()
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
                    If lbCounts.List(i, 3) = "" Then
                        strLine = strLine & vbTab & "<>"
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
                strText = strText & vbCr & vText(j)
            Next j
        End If
        
    
        Open strFileName For Output As #2
    
        'Print #2, vList(n) '& " " & tbSplitterName.Value
        Print #2, strText
    
        Close #2
Next_N:
    Next n
    
    iSave = 0
End Sub

Private Sub UpdatePole()
    If lbCurrentCounts.ListCount < 1 Then Exit Sub
    
    Dim vAttList As Variant
    Dim strCable, strTerm, strTemp As String
    
    strTemp = lbCurrentCounts.List(0, 1) & ": " & lbCurrentCounts.List(0, 2) & ": " & lbCurrentCounts.List(0, 3)
    strCable = tbCableType.Value & " / " & strTemp
    
    If lbCurrentCounts.ListCount > 1 Then
        For i = 1 To lbCurrentCounts.ListCount - 1
            strTemp = lbCurrentCounts.List(i, 1) & ": " & lbCurrentCounts.List(i, 2) & ": " & lbCurrentCounts.List(i, 3)
            strCable = strCable & " + " & strTemp
        Next i
    End If
    
    If tbTerminal.Value = "" Then
        strTerm = ""
    Else
        strTerm = Replace(tbTerminal.Value, "]" & vbCr, "] ")
        strTerm = Replace(strTerm, vbCr, " + ")
    End If
    
    Select Case objCurrent.Name
        Case "sPole"
            vAttList = objCurrent.GetAttributes
            vAttList(25).TextString = strCable
            vAttList(26).TextString = strTerm
        Case Else
            vAttList = objCurrent.GetAttributes
            vAttList(5).TextString = strCable
            vAttList(6).TextString = strTerm
    End Select
    
    objCurrent.Update
    iSavePole = 0
End Sub

Private Sub UpdateCountsList()
    Dim iIndex As Integer
    
    For i = 0 To lbBackspan.ListCount - 1
        iIndex = CInt(lbBackspan.List(i, 0))
        
        lbCounts.List(iIndex, 0) = lbBackspan.List(1, 0)
        lbCounts.List(iIndex, 1) = lbBackspan.List(1, 1)
        lbCounts.List(iIndex, 2) = lbBackspan.List(1, 2)
        lbCounts.List(iIndex, 3) = lbBackspan.List(1, 3)
        lbCounts.List(iIndex, 4) = lbBackspan.List(1, 4)
        lbCounts.List(iIndex, 5) = lbBackspan.List(1, 5)
        lbCounts.List(iIndex, 6) = lbBackspan.List(1, 6)
        lbCounts.List(iIndex, 7) = lbBackspan.List(1, 7)
    Next i
End Sub
