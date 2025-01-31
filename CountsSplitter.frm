VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CountsSplitter 
   Caption         =   "Splitter Form"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   OleObjectBlob   =   "CountsSplitter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CountsSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iExisting As Integer
Dim strPreviousSplitter As String
Dim strPreviousSplices As String
Dim strPreviousF1 As String

Private Sub cbCancel_Click()
    cbChanged.Value = False
    Me.Hide
End Sub

Private Sub cbCreateMCL_Click()
    Call SaveMCL
End Sub

Private Sub cbDone_Click()
    Call SaveMCL
    cbChanged.Value = True
    
    Me.Hide
    
    'Dim strTemp As String
    
    'tbSplice.Value = strPreviousF1
    'strPreviousF1 = ""
    
    'tbF2Name.Value = strPreviousSplitter
    'strPreviousSplitter = ""
    
    'strPreviousSplices = strPreviousSplices & vbCr & tbresults.Value
    'tbResult.Value = ""
    
    'Call OpenMCL
End Sub

Private Sub cbDrop_Click()
    If lbMain.ListCount < 1 Then Exit Sub
    
    Dim strCable, strCounts, strSpliced, strCallout As String
    Dim vLine, vItem As Variant
    
    Load CountsTap
        CountsTap.tbStructure.Value = tbActivePole.Value
        
        For i = 0 To lbMain.ListCount - 1
            CountsTap.lbMain.AddItem i + 1
            
            If lbMain.List(i, 3) = "" Then
                CountsTap.lbMain.List(i, 1) = lbMain.List(i, 1)
                CountsTap.lbMain.List(i, 2) = lbMain.List(i, 2)
            Else
                CountsTap.lbMain.List(i, 1) = "XD"
                CountsTap.lbMain.List(i, 2) = i + 1
            End If
            CountsTap.lbMain.List(i, 3) = ""
        Next i
        
        CountsTap.show
        
        If CountsTap.cbChanged.Value = False Then GoTo Exit_Sub
        
        strLine = CountsTap.tbPosition.Value & ": " & CountsTap.cbCblType.Value & "(" & CountsTap.cbCableSize.Value & ")"
        If Not CountsTap.cbSuffix.Value = "" Then strLine = strLine & CountsTap.cbSuffix.Value
        strLine = strLine & " / " & Replace(CountsTap.tbResult.Value, vbCr, " + ")
        strLine = Replace(strLine, vbLf, "")
        
        For i = 0 To CountsTap.lbMain.ListCount - 1
            If Left(CountsTap.lbMain.List(i, 3), 2) = "Y " Then
                lbMain.List(i, 3) = "T"
            End If
        Next i
        
Exit_Sub:
    Unload CountsTap
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt, vAttList As Variant
    Dim strTap As String
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
    strTap = "T " & vAttList(0).TextString
        
    objBlock.Update
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.List(i, 3) = "T" Then lbMain.List(i, 3) = strTap
    Next i

    bUpdatePole = True
    
    Call GetTapCallout
    
    Me.show
End Sub

Private Sub cbSendAhead_Click()
    If lbMain.ListCount < 1 Then Exit Sub
    
    Dim strCable, strCounts, strSpliced, strCallout As String
    Dim vLine, vItem As Variant
    
    Load CountsTap
        CountsTap.tbStructure.Value = tbActivePole.Value
        
        For i = 0 To lbMain.ListCount - 1
            CountsTap.lbMain.AddItem i + 1
            
            If lbMain.List(i, 3) = "" Then
                CountsTap.lbMain.List(i, 1) = lbMain.List(i, 1)
                CountsTap.lbMain.List(i, 2) = lbMain.List(i, 2)
            Else
                CountsTap.lbMain.List(i, 1) = "XD"
                CountsTap.lbMain.List(i, 2) = i + 1
            End If
            CountsTap.lbMain.List(i, 3) = ""
        Next i
        
        For i = 1 To CountsForm.lbCounts.ListCount - 1
            CountsTap.lbTap.AddItem i
            
            If CountsForm.lbCounts.List(i, 3) = "<>" Then
                CountsTap.lbTap.List(i - 1, 1) = CountsForm.lbCounts.List(i, 1)
                CountsTap.lbTap.List(i - 1, 2) = CountsForm.lbCounts.List(i, 2)
            Else
                CountsTap.lbTap.List(i - 1, 1) = "XD"
                CountsTap.lbTap.List(i - 1, 2) = i
            End If
        Next i
        
        CountsForm.tbPosition.Enabled = False
        CountsForm.cbCblType.Enabled = False
        CountsForm.cbCableSize.Enabled = False
        CountsForm.cbSuffix.Enabled = False
        
        CountsTap.show
        
        If CountsTap.cbChanged.Value = False Then GoTo Exit_Sub
        
        For i = 0 To CountsTap.lbMain.ListCount - 1
            If Left(CountsTap.lbMain.List(i, 3), 2) = "Y " Then
                lbMain.List(i, 3) = Replace(CountsTap.lbMain.List(i, 3), "Y ", "A ")
            End If
        Next i
        
Exit_Sub:
    Unload CountsTap

    bUpdatePole = True
    
    Call GetTapCallout
End Sub

Private Sub cbSendToBack_Click()
    If lbMain.ListCount < 1 Then Exit Sub
    
    Dim strCable, strCounts, strSpliced, strCallout As String
    Dim vLine, vItem As Variant
    
    Load CountsTap
        CountsTap.tbStructure.Value = tbActivePole.Value
        
        For i = 0 To lbMain.ListCount - 1
            CountsTap.lbMain.AddItem i + 1
            
            If lbMain.List(i, 3) = "" Then
                CountsTap.lbMain.List(i, 1) = lbMain.List(i, 1)
                CountsTap.lbMain.List(i, 2) = lbMain.List(i, 2)
            Else
                CountsTap.lbMain.List(i, 1) = "XD"
                CountsTap.lbMain.List(i, 2) = i + 1
            End If
            CountsTap.lbMain.List(i, 3) = ""
        Next i
        
        For i = 1 To CountsForm.lbCounts.ListCount - 1
            CountsTap.lbTap.AddItem i
            
            If CountsForm.lbCounts.List(i, 3) = "<>" Then
                CountsTap.lbTap.List(i - 1, 1) = CountsForm.lbCounts.List(i, 1)
                CountsTap.lbTap.List(i - 1, 2) = CountsForm.lbCounts.List(i, 2)
            Else
                CountsTap.lbTap.List(i - 1, 1) = "XD"
                CountsTap.lbTap.List(i - 1, 2) = i
            End If
        Next i
        
        CountsTap.tbPosition.Enabled = False
        CountsTap.cbCblType.Enabled = False
        CountsTap.cbCableSize.Enabled = False
        CountsTap.cbSuffix.Enabled = False
        
        CountsTap.show
        
        If CountsTap.cbChanged.Value = False Then GoTo Exit_Sub
        
        For i = 0 To CountsTap.lbMain.ListCount - 1
            If Left(CountsTap.lbMain.List(i, 3), 2) = "Y " Then
                lbMain.List(i, 3) = Replace(CountsTap.lbMain.List(i, 3), "Y ", "B ")
            End If
        Next i
        
Exit_Sub:
    Unload CountsTap

    bUpdatePole = True
    
    Call GetTapCallout
End Sub

Private Sub cbSplitterSize_Change()
    If cbSplitterSize.Value = "" Then Exit Sub
    If iExisting = 1 Then Exit Sub
    
    lbMain.Clear
    
    Dim iCounter As Integer
    
    iCounter = CInt(tbStartCount.Value)
    
    For i = 0 To CInt(cbSplitterSize.Value) - 1
        lbMain.AddItem iCounter
        lbMain.List(i, 1) = tbF2Name.Value
        lbMain.List(i, 2) = iCounter
        lbMain.List(i, 3) = ""
        
        iCounter = iCounter + 1
    Next i
    
    Call GetTapCallout
End Sub

Private Sub cbSplitterSplice_Click()
    Dim iSelected As Integer
    Dim strLine As String
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.Selected(i) = True Then
            iSelected = i
            GoTo Found_Selected
        End If
    Next i
    
    MsgBox "No Count Selected."
    Exit Sub
    
Found_Selected:
    
    'strPreviousSplitter = tbF2Name.Value
    'strPreviousSplices = tbResult.Value
    'strPreviousF1 = tbSplice.Value
    Call SaveMCL
    
    strLine = lbMain.List(iSelected, 1) & Asc(93 + CInt(lbMain.List(iSelected, 2)))
    tbSplice.Value = lbMain.List(iSelected, 1) & ": " & lbMain.List(iSelected, 2)
    tbF2Name.Value = strLine
    tbResult.Value = ""
    'iExisting = 0
    
    lbMain.Clear
    
End Sub

Private Sub cbTerminal_Click()
    Dim iSelected As Integer
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.Selected(i) = True Then
            iSelected = i
            GoTo Found_Selected
        End If
    Next i
    
    MsgBox "No Count Selected."
    Exit Sub
    
Found_Selected:
    
    Dim result As Integer
    Dim objEntity As AcadEntity
    Dim objRES As AcadBlockReference
    Dim vAttList, vBasePnt As Variant
    Dim strLine, strResult As String
    Dim strName As String
    
    Me.Hide
    On Error Resume Next
    
Next_One:
    
    If Not lbMain.List(iSelected, 3) = "" Then
        result = MsgBox("Overwrite previous assignment?", vbYesNo, "Fiber Assigned!")
        If result = vbNo Then Exit Sub
    End If
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vBasePnt, "Select Building: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objRES = objEntity
        vAttList = objRES.GetAttributes
            
        Select Case objRES.Name
            Case "SG"
                strResult = "C " & tbActivePole.Value & ";;SG;;" & vAttList(1).TextString & ";;SMARTGRID;;" & vAttList(0).TextString
                
                strLine = lbMain.List(iSelected, 1) & ": " & lbMain.List(iSelected, 2)
                
                vAttList(2).TextString = tbActivePole.Value & " - " & strLine
                objRES.Update
                
            Case "Customer"
                strResult = "C " & tbActivePole.Value & ";;" & vAttList(1).TextString & ";;" & vAttList(2).TextString & ";;" & vAttList(0).TextString
                If vAttList(3).TextString = "" Then
                    strResult = strResult & ";;<>"
                Else
                    strResult = strResult & ";;" & vAttList(3).TextString
                End If
                
                strLine = lbMain.List(iSelected, 1) & ": " & lbMain.List(iSelected, 2)
                If cbDrop.Value = False Then strLine = "(" & strLine & ")"
                vAttList(4).TextString = tbActivePole.Value & " - " & strLine
                objRES.Update
            Case Else
                GoTo Exit_Sub
        End Select
        
        lbMain.List(iSelected, 3) = strResult
    End If
    
    iSelected = iSelected - 1
    
    GoTo Next_One
    
Exit_Sub:
    lbMain.ListIndex = iSelected
    
    Call GetTapCallout
    
    Me.show
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

Private Sub tbF2Name_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If tbF2Name.Value = "" Then Exit Sub
    
    lbMain.Clear
    tbResult.Value = ""
    
    Dim vLine, vItem, vTemp As Variant
    Dim strFileName, strName As String
    
    strName = tbF2Name.Value
    vLine = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vLine(0) & " Counts -" & strName & ".mcl"
    
    For i = 0 To tbF2Name.ListCount - 1
        If tbF2Name.List(i) = strName Then GoTo Found_Name
    Next i
    
    iExisting = 0
    Exit Sub
    
Found_Name:
    
    iExisting = 1
    cbSplitterSize.Enabled = False
    
    Call OpenMCL
    
    cbSplitterSize.Value = lbMain.ListCount
End Sub

Private Sub UserForm_Initialize()
    lbMain.ColumnCount = 4
    lbMain.ColumnWidths = "24;72;36;30"
    
    cbSplitterSize.AddItem "2"
    cbSplitterSize.AddItem "4"
    cbSplitterSize.AddItem "8"
    cbSplitterSize.AddItem "16"
    cbSplitterSize.AddItem "32"
    cbSplitterSize.AddItem "64"
    
    iExisting = 0
    'strPreviousSplitter = ""
    'strPreviousSplices = ""
    'strPreviousF1 = ""
    
    Dim strFile, strFolder, strTemp As String
    Dim strFileName As String
    Dim vName, vLine, vItem As Variant
    Dim strLine, strTab, strCable As String
    Dim fName As String
    Dim iIndex, iCount As Integer
    
    strFolder = ThisDrawing.Path & "\*.*"
    
    strFile = Dir$(strFolder)
    
    Do While strFile <> ""
        If InStr(LCase(strFile), ".mcl") Then
            strTab = Replace(UCase(strFile), ".MCL", "")
            vLine = Split(strTab, " -")
            
            tbF2Name.AddItem vLine(1)
        End If
        strFile = Dir$
    Loop
End Sub

Private Sub GetTapCallout()
    Dim strLine, strItem As String
    Dim strCurrent, strPrevious As String
    Dim iStart, iEnd As Integer
    
    tbResult.Value = ""
    
    If strPreviousSplices = "" Then
        strLine = ""
    Else
        strLine = strPreviousSplices
    End If
    
    If iExisting = 0 Then
        If strLine = "" Then
            strLine = tbSplice.Value & ": " & tbActivePole.Value
        Else
            strLine = strLine & vbCr & tbSplice.Value & ": " & tbActivePole.Value
        End If
    End If
    
    strPrevious = lbMain.List(0, 3)
    If Left(strPrevious, 2) = "A " Then strPrevious = "A"
    If Left(strPrevious, 2) = "B " Then strPrevious = "B"
    If Left(strPrevious, 2) = "C " Then strPrevious = "C"
    If strPrevious = "X" Then strPrevious = ""
    
    strCurrent = strPrevious
    iStart = CInt(lbMain.List(0, 2))
    iEnd = iStart
    
    For i = 1 To lbMain.ListCount - 1
        strCurrent = lbMain.List(i, 3)
        If Left(strCurrent, 2) = "A " Then strCurrent = "A"
        If Left(strCurrent, 2) = "B " Then strCurrent = "B"
        If Left(strCurrent, 2) = "C " Then strCurrent = "C"
        If strCurrent = "X" Then strCurrent = ""
        
        If strCurrent = strPrevious Then
            iEnd = CInt(lbMain.List(i, 2))
        Else
            If Not strPrevious = "" Then
                strItem = lbMain.List(i - 1, 1) & ": " & iStart
                If iEnd > iStart Then strItem = strItem & "-" & iEnd
                
                strItem = strItem & ": " & tbActivePole.Value
                
                If strLine = "" Then
                    strLine = strItem
                Else
                    strLine = strLine & vbCr & strItem
                End If
            End If
             
            strPrevious = strCurrent
            iStart = CInt(lbMain.List(i, 2))
            iEnd = iStart
        End If
        
    Next i
             
    If Not strPrevious = "" Then
        strItem = lbMain.List(i - 1, 1) & ": " & iStart
        If iEnd > iStart Then strItem = strItem & "-" & iEnd
        
        strItem = strItem & ": " & tbActivePole.Value
             
        If strLine = "" Then
            strLine = strItem
        Else
            strLine = strLine & vbCr & strItem
        End If
    End If
    
    tbResult.Value = strLine
End Sub

Private Sub SaveMCL()
    Dim strFileName As String
    Dim vName, vLine, vItem As Variant
    Dim strLine, strTemp As String
    'Dim fName As String
    Dim iIndex As Integer
    
    vName = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vName(0) & " Counts -" & tbF2Name.Value & ".mcl"
    
    'fName = Dir(strFilename)
    'If Not fName = "" Then
        'MsgBox "File already exist."
        'Exit Sub
    'End If
    
    Open strFileName For Output As #1
    
    Print #1, tbF2Name.Value & " " & tbSplitterName.Value
    
    For i = 0 To lbMain.ListCount - 1
        If Left(lbMain.List(i, 3), 2) = "C " Then
            'vItem = Split(lbMain.List(i, 3), "C ")
            'vLine = Split(vItem(1), ";;")
            strTemp = Right(lbMain.List(i, 3), Len(lbMain.List(i, 3)) - 2)
            vLine = Split(strTemp, ";;")
                
            strLine = lbMain.List(i, 2) & vbTab & vLine(0) & vbTab & vLine(1) & vbTab & vLine(2) & vbTab & vLine(3) & vbTab & vLine(4)
        Else
            strLine = lbMain.List(i, 2) & vbTab & "<>" & vbTab & "<>" & vbTab & "<>" & vbTab & "<>" & vbTab & "<>"
        End If
        
        Print #1, strLine
    Next i
    
    Close #1
End Sub

Private Sub OpenMCL()
    Dim vLine As Variant
    Dim strFileName, strName As String
    Dim strLine As String
    Dim iIndex As Integer
    
    strName = tbF2Name.Value
    vLine = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vLine(0) & " Counts -" & strName & ".mcl"
    
    Open strFileName For Input As #1
    
    Line Input #1, strLine
    While Not EOF(1)
        Line Input #1, strLine
        If strLine = "" Then GoTo Next_line
        
        vLine = Split(strLine, vbTab)
        
        iIndex = lbMain.ListCount + 1
        lbMain.AddItem iIndex
        iIndex = iIndex - 1
        
        lbMain.List(iIndex, 1) = strName
        lbMain.List(iIndex, 2) = vLine(0)
        lbMain.List(iIndex, 3) = ""
        If Not vLine(4) = "<>" Then lbMain.List(iIndex, 3) = "X"
        
Next_line:
    Wend
    
    Close #1
End Sub

