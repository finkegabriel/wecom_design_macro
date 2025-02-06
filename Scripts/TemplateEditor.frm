VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateEditor 
   Caption         =   "Template Editor"
   ClientHeight    =   12195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "TemplateEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TemplateEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim iIsVar As Integer
Dim strFLetter, strLLetter As String

Private Sub cbAdd_Click()
    Call AllOn
    
    lbCond.List(lbCond.ListIndex, 2) = tbValue.Value
    tbValue.Value = ""
    cbAdd.Enabled = False
    tbValue.Enabled = False
End Sub

Private Sub cbAddNoteColor_Click()
    Dim strLine, strText, strFront, strEnd As String
    Dim iStart, iEnd As Integer
    
    strLine = Replace(tbSpecial.Value, vbLf, "")
    strText = Replace(tbSpecial.SelText, vbLf, "")
    If Right(strText, 1) = vbCr Then strText = Left(strText, Len(strText) - 1)
    
    iStart = InStr(strLine, strText) - 1
    iEnd = iStart + Len(strText)
    
    strFront = Left(strLine, iStart)
    strEnd = Right(strLine, Len(strLine) - iEnd)
    
    'MsgBox "Total: " & Len(strLine) & vbCr & "Start: " & iStart & vbCr & "End: " & iEnd
    
    tbSpecial.Value = strFront & "<p style=""color:" & cbNoteColor.Value & """>" & strText & "</p>" & strEnd
    
    
    
    
    'Dim strText As String
    'Dim iStart, iLength, iTemp As Integer
    'Dim vTemp As Variant
    
    'vTemp = Split(tbSpecial.Value, vbCrLf)
    'iTemp = UBound(vTemp)
    
    'tbSpecial.Value = tbSpecial.Value & "<p style=""color:" & cbNoteColor.Value & """>**add text here**</p>"
    'iStart = Len(tbSpecial.Value) - 21 - iTemp
    'iLength = 17
    'tbSpecial.SelStart = iStart
    'tbSpecial.SelLength = iLength
End Sub

Private Sub cbAddNotesTable_Click()
    Dim strText, strLine As String
    Dim iSelStart As Integer
    
    iSelStart = tbSpecial.SelStart
    strText = Replace(tbSpecial.Value, vbCr, "")
    
    strLine = "<table>" & vbCr & "<tr>Item = Value" & vbCr & "<tr>Item = Value" & vbCr & "<tr>Item = Value" & vbCr & "<tr>Item = Value" & vbCr & "</table>"
    tbSpecial.Value = Left(strText, iSelStart) & strLine & Mid(strText, iSelStart + 1)
End Sub

Private Sub cbAddSOWColor_Click()
    Dim strLine, strText, strFront, strEnd As String
    Dim iStart, iEnd As Integer
    
    strLine = Replace(tbSOW.Value, vbLf, "")
    strText = Replace(tbSOW.SelText, vbLf, "")
    If Right(strText, 1) = vbCr Then strText = Left(strText, Len(strText) - 1)
    
    iStart = InStr(strLine, strText) - 1
    iEnd = iStart + Len(strText)
    
    strFront = Left(strLine, iStart)
    strEnd = Right(strLine, Len(strLine) - iEnd)
    
    'MsgBox "Total: " & Len(strLine) & vbCr & "Start: " & iStart & vbCr & "End: " & iEnd
    
    tbSOW.Value = strFront & "<p style=""color:" & cbSOWColor.Value & """>" & strText & "</p>" & strEnd
End Sub

Private Sub cbAddTable_Click()
    Dim strText, strLine As String
    Dim iSelStart As Integer
    
    iSelStart = tbSOW.SelStart
    strText = Replace(tbSOW.Value, vbCr, "")
    
    strLine = "<table>" & vbCr & "<tr>Item = Value" & vbCr & "<tr>Item = Value" & vbCr & "<tr>Item = Value" & vbCr & "<tr>Item = Value" & vbCr & "</table>"
    tbSOW.Value = Left(strText, iSelStart) & strLine & Mid(strText, iSelStart + 1)
End Sub

Private Sub cbCreate_Click()
    If cbName.Value = "" Then Exit Sub
    
    Dim strFileName As String
    Dim strPath As String
    Dim strLine As String
    Dim vTemp As Variant
    
    strPath = ThisDrawing.Path & "\"
    vTemp = Split(LCase(strPath), "dropbox")
    strFileName = vTemp(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\1-JOBS\PLANNING\SCOPE TEMPLATES\"
    strFileName = strFileName & cbName.Value & ".swt"
    
    'MsgBox strFileName
    'Exit Sub

    Open strFileName For Output As #1
    
    Print #1, "[[INSTRUCTIONS]]"
    Print #1, tbSpecial.Value
    Print #1, "[[SCOPE]]"
    Print #1, tbSOW.Value
    Print #1, "[[CONDITIONS]]"
    If lbCond.ListCount > 0 Then
        For i = 0 To lbCond.ListCount - 1
            strLine = lbCond.List(i, 0) & "=" & lbCond.List(i, 1) & "=" & lbCond.List(i, 2)
            Print #1, strLine
        Next i
    End If
    
    Close #1
    
    MsgBox "Saved"
End Sub

Private Sub cbGetExisting_Click()
    If cbName.Value = "" Then Exit Sub
    
    lbVar.Clear
    lbCond.Clear
    tbSpecial.Value = ""
    tbSOW.Value = ""
    
    Dim strFileName As String
    Dim strPath As String
    Dim strLine As String
    Dim vLine, vTemp, vVar As Variant
    Dim fName As String
    Dim iStatus As Integer
    
    iStatus = 0
    
    strPath = ThisDrawing.Path & "\"
    vTemp = Split(LCase(strPath), "dropbox")
    strFileName = vTemp(0) & "Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\1-JOBS\PLANNING\SCOPE TEMPLATES\"
    strFileName = strFileName & cbName.Value & ".swt"
    
    fName = Dir(strFileName)
    If fName = "" Then
        Exit Sub
    End If
    
    Open strFileName For Input As #2
    
    'MsgBox strFileName
    'Exit Sub
    
    While Not EOF(2)
        Line Input #2, strLine
        
        If InStr(strLine, ">>") > 0 And Not iStatus = 3 Then
            vLine = Split(strLine, ">>")
            For j = 0 To UBound(vLine) - 1
                vTemp = Split(vLine(j), "<<")
                vVar = Split(vTemp(1), "/")
            
                'If lbVariables.ListCount < 0 Then
                    For i = 0 To lbVar.ListCount - 1
                        If lbVar.List(i, 0) = vVar(0) Then GoTo Skip_Adding_Var
                    Next i
                'End If
            
                lbVar.AddItem vVar(0)
                If UBound(vVar) > 0 Then
                    lbVar.List(lbVar.ListCount - 1, 1) = Replace(vVar(1), ">>", "")
                Else
                    lbVar.List(lbVar.ListCount - 1, 1) = "<Empty>"
                End If
        
Skip_Adding_Var:
            Next j
            
            If UBound(vLine) > 2 Then
                For i = 1 To UBound(vLine) - 1
                    vTemp = Split(vLine(i), "<<")
                    vVar = Split(vTemp(1), "/")
            
                    'For j = 0 To lbVariables.ListCount - 1
                        'If lbVariables.List(j, 0) = vVar(0) Then GoTo Skip_Adding_This
                    'Next j
            
                    lbVar.AddItem vVar(0)
                    If UBound(vVar) > 0 Then
                        lbVar.List(lbVar.ListCount - 1, 1) = Replace(vVar(1), ">>", "")
                    Else
                        lbVar.List(lbVar.ListCount - 1, 1) = "<Empty>"
                    End If
Skip_Adding_This:
                    'cbFind.AddItem "<<" & vTemp(1) & ">>"
                Next i
            End If
        End If
        
        If Left(strLine, 2) = "[[" Then
            Select Case Left(strLine, 3)
                Case "[[I"
                    iStatus = 1
                Case "[[S"
                    iStatus = 2
                Case "[[C"
                    iStatus = 3
            End Select
        Else
            Select Case iStatus
                Case Is = 1
                    If tbSpecial.Value = "" Then
                        tbSpecial.Value = strLine
                    Else
                        tbSpecial.Value = tbSpecial.Value & vbCr & strLine
                    End If
                Case Is = 2
                    If tbSOW.Value = "" Then
                        tbSOW.Value = strLine
                    Else
                        tbSOW.Value = tbSOW.Value & vbCr & strLine
                    End If
                Case Is = 3
                    '<---------------------------------------------------------------------
                    vCond = Split(strLine, "=")
                    lbCond.AddItem vCond(0)
                    lbCond.List(lbCond.ListCount - 1, 1) = vCond(1)
                    lbCond.List(lbCond.ListCount - 1, 2) = vCond(2)
            End Select
        End If
    Wend
    
    Close #2
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub Label24_Click()
    tbValue.Value = "<p style=""color:orange"">" & tbValue.Value & "</P>"
End Sub

Private Sub Label25_Click()
    tbValue.Value = "<p style=""color:green"">" & tbValue.Value & "</P>"
End Sub

Private Sub Label26_Click()
    tbValue.Value = "<p style=""color:blue"">" & tbValue.Value & "</P>"
End Sub

Private Sub Label27_Click()
    tbValue.Value = "<p style=""color:purple"">" & tbValue.Value & "</P>"
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub lbCond_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call AllOff
    
    tbValue.Enabled = True
    
    If Not lbCond.List(lbCond.ListIndex, 2) = "" Then tbValue.Value = lbCond.List(lbCond.ListIndex, 2)
    
    cbAdd.Enabled = True
    tbValue.SetFocus
End Sub

Private Sub lbCond_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            lbCond.RemoveItem lbCond.ListIndex
            'MsgBox "Deleted"
        Case Else
    End Select
End Sub

Private Sub lbVar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strPre, strPost As String
    Dim strPVar, strPList As String
    Dim strSVar, strSList As String
    Dim strTemp As String
    Dim vList As Variant
    Dim iIndex As Integer
    
    iIndex = lbVar.ListIndex
    
    If lbVar.List(iIndex, 1) = "<Empty>" Then Exit Sub
    
    strPVar = lbVar.List(iIndex, 0)
    strPList = lbVar.List(iIndex, 1)
    strPre = "<<" & strPVar & "/" & strPList & ">>"
    
    Load TemplateVarEdit
        TemplateVarEdit.tbVar.Value = strPVar
        TemplateVarEdit.tbList.Value = Replace(strPList, ",", vbCr)
        TemplateVarEdit.show
        
        strSVar = TemplateVarEdit.tbVar.Value
        strSList = Replace(TemplateVarEdit.tbList.Value, vbCr, ",")
        strSList = Replace(strSList, vbLf, "")
    Unload TemplateVarEdit
    
    If strSVar = strPVar Then
        If strSList = strPList Then Exit Sub
    Else
        If lbCond.ListCount > 0 Then
            For i = 0 To lbCond.ListCount - 1
                If lbCond.List(i, 0) = strPVar Then lbCond.List(i, 0) = strSVar
            Next i
            
            lbVar.List(iIndex, 0) = strSVar
        
            strTemp = Replace(tbSpecial.Value, "{{" & strPVar & "}}", "{{" & strSVar & "}}")
            tbSpecial.Value = strTemp
        
            strTemp = Replace(tbSOW.Value, "{{" & strPVar & "}}", "{{" & strSVar & "}}")
            tbSOW.Value = strTemp
        End If
    End If
    
    lbVar.List(iIndex, 1) = strSList
    
    strPost = "<<" & lbVar.List(iIndex, 0) & "/" & strSList & ">>"
        
    strTemp = Replace(tbSpecial.Value, strPre, strPost)
    tbSpecial.Value = strTemp
        
    strTemp = Replace(tbSOW.Value, strPre, strPost)
    tbSOW.Value = strTemp
        
    vList = Split(strSList, ",")
        
    For i = 0 To lbCond.ListCount - 1
        If lbCond.List(i, 0) = strPVar Then GoTo Found_lbCond
    Next i
    Exit Sub
        
Found_lbCond:
    
    If lbCond.ListCount < 1 Then Exit Sub
    For j = 0 To UBound(vList)
        If i > lbCond.ListCount - 1 Then
            lbCond.AddItem strPVar
            lbCond.List(lbCond.ListCount - 1, 1) = vList(j)
        Else
            If strPVar = lbCond.List(i, 0) Then
                lbCond.List(i, 1) = vList(j)
            Else
                lbCond.AddItem strPVar, i
                lbCond.List(i, 1) = vList(j)
            End If
        End If
            
        i = i + 1
    Next j
    
End Sub

Private Sub lbVar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            lbVar.RemoveItem lbVar.ListIndex
            'MsgBox "Deleted"
        Case Else
    End Select
End Sub

Private Sub LRedNote_Click()
    Dim strText As String
    Dim iStart, iLength, iTemp As Integer
    Dim vTemp As Variant
    
    vTemp = Split(tbSpecial.Value, vbCrLf)
    iTemp = UBound(vTemp)
    
    tbSpecial.Value = tbSpecial.Value & "<p style=""color:red"">**add text here**</p>"
    iStart = Len(tbSpecial.Value) - 21 - iTemp
    iLength = 17
    tbSpecial.SelStart = iStart
    tbSpecial.SelLength = iLength
End Sub

Private Sub LRedSOW_Click()
    Dim strLine, strText, strFront, strEnd As String
    Dim iStart, iEnd As Integer
    
    strLine = Replace(tbSOW.Value, vbLf, "")
    strText = Replace(tbSOW.SelText, vbLf, "")
    If Right(strText, 1) = vbCr Then strText = Left(strText, Len(strText) - 1)
    
    iStart = InStr(strLine, strText) - 1
    iEnd = iStart + Len(strText)
    
    strFront = Left(strLine, iStart)
    strEnd = Right(strLine, Len(strLine) - iEnd)
    
    'MsgBox "Total: " & Len(strLine) & vbCr & "Start: " & iStart & vbCr & "End: " & iEnd
    
    tbSOW.Value = strFront & "<p style=""color:red"">" & strText & "</p>" & strEnd
    
    'Dim strFind, strReplace As String
    
    'strFind = tbSOW.SelText
    'strReplace = "<p style=""color:red"">" & strFind & "</p>"
    
    'tbSOW.Value = Replace(tbSOW.Value, strFind, strReplace)
End Sub

Private Sub LRedText_Click()
    tbValue.Value = "<p style=""color:red"">" & tbValue.Value & "</P>"
End Sub

Private Sub tbSOW_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strKey As String
    
    'MsgBox KeyCode
    
    Select Case KeyCode
        Case Is = 188
            If strFLetter = "<" Then
                iIsVar = 2
                Call AllOff
                tbVar.Enabled = True
                tbVar.SetFocus
            Else
                strFLetter = "<"
            End If
        Case Is = 219
            If strFLetter = "{" Then
                
                Load TemplateList
                    For i = 0 To lbVar.ListCount - 1
                        If Not lbVar.List(i, 1) = "<Empty>" Then
                            If lbCond.ListCount < 0 Then
                                TemplateList.cbList.AddItem lbVar.List(i, 0)
                            Else
                                For j = 0 To lbCond.ListCount - 1
                                    If lbCond.List(j, 0) = lbVar.List(i, 0) Then GoTo Found_Cond
                                Next j
                                
                                TemplateList.cbList.AddItem lbVar.List(i, 0)
                            End If
Found_Cond:
                        End If
                    Next i
                    
                    If TemplateList.cbList.ListCount < 1 Then
                        MsgBox "No Variables with a List exist."
                        Unload TemplateList
                        GoTo Exit_Sub
                    'Else
                        'If Not lbCond.ListCount < 0 Then
                            'For j = 0 To TemplateList.cbList.ListCount - 1
                                'For k = 0 To lbCond.ListCount - 1
                                    'If TemplateList.cbList.List(j, 0) = lbCond.List(k, 0) Then
                                    'End If
                                'Next k
                            'Next j
                        'End If
                    End If
                    
                    TemplateList.show
                    strTemp = TemplateList.cbList.Value
                    
                    iSelStart = tbSOW.SelStart
                    strText = Replace(tbSOW.Value, vbCr, "")
                    tbSOW.Value = Left(strText, iSelStart) & strTemp & "}}" & Mid(strText, iSelStart + 1)
                    tbSOW.SelStart = iSelStart + Len(strTemp) + 2
                    tbSOW.SetFocus
                    
                    For i = 0 To lbVar.ListCount - 1
                        If lbVar.List(i, 0) = strTemp Then
                            vLine = Split(lbVar.List(i, 1), ",")
                            
                            For n = 0 To UBound(vLine)
                                lbCond.AddItem lbVar.List(i, 0)
                                lbCond.List(lbCond.ListCount - 1, 1) = vLine(n)
                            Next n
                        End If
                    Next i
                    
                Unload TemplateList
                iIsVar = 0
            Else
                strFLetter = "{"
            End If
        Case Else
            strFLetter = ""
    End Select
Exit_Sub:
    
End Sub

Private Sub tbSpecial_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim vLine As Variant
    Dim strKey As String
    Dim strText, strTemp As String
    Dim iSelStart As Integer
    
    'MsgBox KeyCode
    
    Select Case KeyCode
        Case Is = 188
            If strFLetter = "<" Then
                iIsVar = 1
                Call AllOff
                tbVar.Enabled = True
                tbVar.SetFocus
            Else
                strFLetter = "<"
            End If
        Case Is = 219
            If strFLetter = "{" Then
                
                Load TemplateList
                    For i = 0 To lbVar.ListCount - 1
                        If Not lbVar.List(i, 1) = "<Empty>" Then
                            TemplateList.cbList.AddItem lbVar.List(i, 0)
                        End If
                    Next i
                    
                    If TemplateList.cbList.ListCount < 1 Then
                        MsgBox "No Variables with a List exist."
                        Unload TemplateList
                        GoTo Exit_Sub
                    End If
                    
                    TemplateList.show
                    strTemp = TemplateList.cbList.Value
                    
                    iSelStart = tbSpecial.SelStart
                    strText = Replace(tbSpecial.Value, vbCr, "")
                    tbSpecial.Value = Left(strText, iSelStart) & strTemp & "}}" & Mid(strText, iSelStart + 1)
                    tbSpecial.SelStart = iSelStart + Len(strTemp) + 2
                    tbSpecial.SetFocus
                    
                    For i = 0 To lbVar.ListCount - 1
                        If lbVar.List(i, 0) = strTemp Then
                            vLine = Split(lbVar.List(i, 1), ",")
                            
                            For n = 0 To UBound(vLine)
                                lbCond.AddItem lbVar.List(i, 0)
                                lbCond.List(lbCond.ListCount - 1, 1) = vLine(n)
                            Next n
                        End If
                    Next i
                    
                Unload TemplateList
                iIsVar = 0
            Else
                strFLetter = "{"
            End If
        Case Else
            strFLetter = ""
    End Select
Exit_Sub:
    
End Sub

Private Sub tbVar_Change()
    Dim vLine As Variant
    Dim strText As String
    Dim iSelStart As Integer
    
    If Right(tbVar.Value, 1) = "/" Then
        Call AllOn
        strText = Left(tbVar.Value, Len(tbVar.Value) - 1)
        For i = 0 To lbVar.ListCount - 1
            If lbVar.List(i, 0) = strText Then
                tbVar.Value = tbVar.Value & lbVar.List(i, 1)
        
                Select Case iIsVar
                    Case Is = 1
                        iSelStart = tbSpecial.SelStart
                        strText = Replace(tbSpecial.Value, vbCr, "")
                        tbSpecial.Value = Left(strText, iSelStart) & tbVar.Value & ">>" & Mid(strText, iSelStart + 1)
                        tbSpecial.SelStart = iSelStart + Len(tbVar.Value) + 2
                        tbSpecial.SetFocus
                    Case Is = 2
                        iSelStart = tbSOW.SelStart
                        strText = Replace(tbSOW.Value, vbCr, "")
                        tbSOW.Value = Left(strText, iSelStart) & tbVar.Value & ">>" & Mid(strText, iSelStart + 1)
                        tbSOW.SelStart = iSelStart + Len(tbVar.Value) + 2
                        tbSOW.SetFocus
                End Select
        
                tbVar.Value = ""
                iIsVar = 0
                strFLetter = ""
                tbVar.Enabled = False
                
                GoTo Found_Item
            End If
        Next i
    End If
    
    If Right(tbVar.Value, 2) = ">>" Then
        Call AllOn
        
        Select Case iIsVar
            Case Is = 1
                iSelStart = tbSpecial.SelStart
                strText = Replace(tbSpecial.Value, vbCr, "")
                tbSpecial.Value = Left(strText, iSelStart) & tbVar.Value & Mid(strText, iSelStart + 1)
                tbSpecial.SelStart = iSelStart + Len(tbVar.Value)
                tbSpecial.SetFocus
            Case Is = 2
                iSelStart = tbSOW.SelStart
                strText = Replace(tbSOW.Value, vbCr, "")
                tbSOW.Value = Left(strText, iSelStart) & tbVar.Value & Mid(strText, iSelStart + 1)
                tbSOW.SelStart = iSelStart + Len(tbVar.Value)
                tbSOW.SetFocus
        End Select
        
        
        tbVar.Value = Left(tbVar.Value, Len(tbVar.Value) - 2)
        
        For i = 0 To lbVar.ListCount - 1
            If lbVar.List(i, 0) = tbVar.Value Then GoTo Skip_This
        Next i
        
        If InStr(tbVar.Value, "/") < 1 Then
            lbVar.AddItem tbVar.Value
            lbVar.List(lbVar.ListCount - 1, 1) = "<Empty>"
        Else
            vLine = Split(tbVar.Value, "/")

            lbVar.AddItem vLine(0)

            lbVar.List(lbVar.ListCount - 1, 1) = vLine(1)
        End If
        
Skip_This:
        
        tbVar.Value = ""
        iIsVar = 0
        strFLetter = ""
        tbVar.Enabled = False
    End If
Found_Item:

End Sub

Private Sub UserForm_Initialize()
    lbVar.ColumnCount = 2
    lbVar.ColumnWidths = "120;360"
    
    lbCond.ColumnCount = 3
    lbCond.ColumnWidths = "120;120;240"
    
    cbNoteColor.AddItem "RED"
    cbNoteColor.AddItem "ORANGE"
    cbNoteColor.AddItem "GREEN"
    cbNoteColor.AddItem "BLUE"
    cbNoteColor.AddItem "PURPLE"
    'cbNoteColor.AddItem "GRAY"
    cbNoteColor.Value = "RED"
    
    cbSOWColor.AddItem "RED"
    cbSOWColor.AddItem "ORANGE"
    cbSOWColor.AddItem "GREEN"
    cbSOWColor.AddItem "BLUE"
    cbSOWColor.AddItem "PURPLE"
    'cbSOWColor.AddItem "GRAY"
    cbSOWColor.Value = "RED"
    
    iIsVar = 0
    strFLetter = ""
    
    Dim strPath, strFile As String
    Dim strTFolder As String
    Dim vTemp As Variant
    
    strTFolder = ""
    
    strPath = LCase(ThisDrawing.Path)
    'MsgBox strPath
    vTemp = Split(strPath, "dropbox")
    If UBound(vTemp) < 1 Then
        Exit Sub
    End If
    
    strTFolder = vTemp(0) & "\Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\1-JOBS\PLANNING\SCOPE TEMPLATES\*.*"
    
    strFile = Dir$(strTFolder)
    
    Do While strFile <> ""
        If InStr(strFile, ".swt") Then
            cbName.AddItem Replace(strFile, ".swt", "")
        End If
        strFile = Dir$
    Loop
End Sub

Private Sub AllOn()
    tbSpecial.Enabled = True
    tbSOW.Enabled = True
    cbAddTable.Enabled = True
    tbVar.Enabled = True
    cbScan.Enabled = True
    lbVar.Enabled = True
    lbCond.Enabled = True
    'tbValue.Enabled = True
    'cbAdd.Enabled = True
    cbCreate.Enabled = True
End Sub

Private Sub AllOff()
    tbSpecial.Enabled = False
    tbSOW.Enabled = False
    cbAddTable.Enabled = False
    tbVar.Enabled = False
    cbScan.Enabled = False
    lbVar.Enabled = False
    lbCond.Enabled = False
    'tbValue.Enabled = False
    'cbAdd.Enabled = False
    cbCreate.Enabled = False
End Sub
