VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddRemoveUnits 
   Caption         =   "Add / Remove Units"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   OleObjectBlob   =   "AddRemoveUnits.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddRemoveUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objBlock As AcadBlockReference

Private Sub cbAddUnit_Click()
    Dim iIndex As Integer
    
    If cbAddUnit.Caption = "Update" Then
        iIndex = lbUnits.ListIndex
        
        lbUnits.List(iIndex, 0) = tbUnit.Value
        lbUnits.List(iIndex, 1) = tbQuantity.Value
        If Not tbUNote.Value = "" Then
            lbUnits.List(iIndex, 2) = ""
        Else
            lbUnits.List(iIndex, 2) = tbUNote.Value
        End If
        
        tbUnit.Value = ""
        tbQuantity.Value = ""
        tbUNote.Value = ""
        
        cbAddUnit.Caption = "Add Unit"
    Else
        lbUnits.AddItem tbUnit.Value
        
        iIndex = lbUnits.ListCount - 1
        lbUnits.List(iIndex, 1) = tbQuantity.Value
        
        If Not tbUNote.Value = "" Then
            lbUnits.List(iIndex, 2) = ""
        Else
            lbUnits.List(iIndex, 2) = tbUNote.Value
        End If
        
        tbUnit.Value = ""
        tbQuantity.Value = ""
        tbUNote.Value = ""
        
    End If
End Sub

Private Sub cbCombine_Click()
    If lbUnits.ListCount < 2 Then Exit Sub
    
    Dim vCurrent, vOther, vItem As Variant
    Dim strCurrent As String
    
    For i = lbUnits.ListCount - 1 To 0 Step -1
        strCurrent = lbUnits.List(i, 0)
        
        For j = 0 To i - 1
            If lbUnits.List(j, 0) = strCurrent Then
                lbUnits.List(j, 1) = CInt(lbUnits.List(j, 1)) + CInt(lbUnits.List(i, 1))
                lbUnits.RemoveItem i
                GoTo Next_I
            End If
        Next j
Next_I:
    Next i
End Sub

Private Sub cbERemovetUnits_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objItem As AcadBlockReference
    Dim vAttList As Variant
    Dim strAddUnit, strTemp As String
    Dim iRemoveIndex, iAddIndex As Integer
    Dim iAmount As Integer
    
    iRemoveIndex = -1
    iAddIndex = -1
    
    On Error Resume Next
    Me.Hide
    
    Set objSS = ThisDrawing.SelectionSets.Add("objss")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objss")
        Err = 0
    End If
    
    objSS.SelectOnScreen
    
    If objSS.count = 0 Then GoTo Exit_Sub
    
    For Each objEntity In objSS
        If Not TypeOf objEntity Is AcadBlockReference Then GoTo Next_objEntity
        
        Set objItem = objEntity
        vAttList = objItem.GetAttributes
        
        If Not Err = 0 Then GoTo Next_objEntity
            
        Select Case objItem.Name
            Case "ExGuyOL", "ExGuyOR"
                If InStr(LCase(objItem.Layer), "existing") > 0 Then
                    iRemoveIndex = GetExistingIndex("PM11")
                    iAmount = CInt(lbEUnits.List(iRemoveIndex, 2))
                
                    If iAmount > 1 Then
                        lbEUnits.List(iRemoveIndex) = iAmount - 1
                    Else
                        lbEUnits.RemoveItem iRemoveIndex
                    End If
                    
                    iRemoveIndex = GetExistingIndex(CStr(vAttList(2).TextString))
                    iAddIndex = 1001
                    strAddUnit = "+XXPE"
                    
                    objItem.Layer = "Integrity Removed"
                    objItem.Update
                Else
                    iAddIndex = GetProposedIndex("PM11")
                    iAmount = CInt(lbUnits.List(iAddIndex, 1))
                
                    If iAmount > 1 Then
                        lbUnits.List(iAddIndex) = iAmount - 1
                    Else
                        lbUnits.RemoveItem iAddIndex
                    End If
                    
                    iRemoveIndex = -1
                    iAddIndex = GetProposedIndex(CStr(vAttList(2).TextString))
                    
                    objItem.Delete
                End If
            Case "ExAncOL", "ExAncOR"
                If InStr(LCase(objItem.Layer), "existing") > 0 Then
                    iRemoveIndex = GetExistingIndex(CStr(vAttList(0).TextString))
                    iAddIndex = 1001
                    strAddUnit = "+XXPF"
                    
                    objItem.Layer = "Integrity Removed"
                    objItem.Update
                Else
                    iRemoveIndex = -1
                    iAddIndex = GetProposedIndex(CStr(vAttList(0).TextString))
                    
                    objItem.Delete
                End If
            Case "Map splice"
                If InStr(LCase(objItem.Layer), "aerial") > 0 Then iAddIndex = GetProposedIndex("+HACO")
                If InStr(LCase(objItem.Layer), "buried") > 0 Then iAddIndex = GetProposedIndex("+HBFO")
                
                objItem.Delete
            Case "Map coil"
                strTemp = Replace(vAttList(1).TextString, "F", "")
                If InStr(LCase(objItem.Layer), "aerial") > 0 Then strTemp = "+CO(" & strTemp & ")E"
                If InStr(LCase(objItem.Layer), "buried") > 0 Then strTemp = "+UO(" & strTemp & ")"
                
                iAddIndex = GetCoilIndex(strTemp)
                If iAddIndex < 0 Then
                    strTemp = Replace(strTemp, "UO", "BFO")
                    iAddIndex = GetCoilIndex(strTemp)
                End If
                
                objItem.Delete
        End Select
        
        'MsgBox objItem.Name & vbCr & "Remove: " & iRemoveIndex & vbCr & "Add: " & iAddIndex & vbCr & "String: " & strAddUnit
        
        If iRemoveIndex < 0 Then
            If iAddIndex < 0 Then
                GoTo Next_objEntity
            Else
                iAmount = CInt(lbUnits.List(iAddIndex, 1))
                
                If iAmount > 1 Then
                    lbUnits.List(iAddIndex) = iAmount - 1
                Else
                    lbUnits.RemoveItem iAddIndex
                End If
            End If
        Else
            iAmount = CInt(lbEUnits.List(iRemoveIndex, 1))
                
            If iAmount > 1 Then
                lbEUnits.List(iRemoveIndex) = iAmount - 1
            Else
                lbEUnits.RemoveItem iRemoveIndex
            End If
            
            If Not strAddUnit = "" Then
                If lbUnits.ListCount < 0 Then
                    lbUnits.AddItem strAddUnit
                    lbUnits.List(lbUnits.ListCount - 1, 1) = 1
                Else
                    For i = 0 To lbUnits.ListCount - 1
                        If lbUnits.List(i, 0) = strAddUnit Then
                            iAmount = CInt(lbUnits.List(i, 1)) + 1
                            lbUnits.List(i, 1) = iAmount
                            GoTo Next_objEntity
                        End If
                    Next i
                        
                    lbUnits.AddItem strAddUnit
                    lbUnits.List(lbUnits.ListCount - 1, 1) = 1
                End If
                GoTo Next_objEntity
            End If
            
            If iAddIndex > -1 Then
                iAmount = CInt(lbUnits.List(iAddIndex, 1))
                
                If iAmount > 1 Then
                    lbUnits.List(iAddIndex) = iAmount - 1
                Else
                    lbUnits.RemoveItem iAddIndex
                End If
            End If
        End If
        
Next_objEntity:
        Err = 0
        iRemoveIndex = -1
        iAddIndex = -1
        strAddUnit = ""
    Next objEntity
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    ThisDrawing.Regen acActiveViewport
    
    Me.show
End Sub

Private Sub cbGetBlock_Click()
    Dim objObject As AcadObject
    Dim vBasePnt, vAttList As Variant
    Dim vUnits, vLine, vItem, vTemp As Variant
    Dim strTemp As String
    Dim iAtt As Integer
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objObject, vBasePnt, "Select Pole: "
    If TypeOf objObject Is AcadBlockReference Then
        Set objBlock = objObject
    Else
        MsgBox "Not a valid object."
        Me.show
        Exit Sub
    End If
    
    If Not objBlock.Name = "sPole" Then
        MsgBox "Not a valid Block."
        Me.show
        Exit Sub
    End If
    
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 27
        Case "sPed", "sHH", "sFP"
            iAtt = 7
        Case Else
            MsgBox "Not a valid Block."
            Me.show
            Exit Sub
    End Select
    
    lbEUnits.Clear
    lbUnits.Clear
    
    vAttList = objBlock.GetAttributes
    
    If Not vAttList(iAtt).TextString = "" Then
        vAttList(iAtt).TextString = Replace(vAttList(iAtt).TextString, vbLf, "")
        vUnits = Split(vAttList(iAtt).TextString, " <-- ")
        
        If UBound(vUnits) > 0 Then
            vLine = Split(vUnits(0), ";;")
            For i = 0 To UBound(vLine)
                vItem = Split(vLine(i), "=")
                lbEUnits.AddItem vItem(0)
                If InStr(vItem(1), "  ") > 0 Then
                    vTemp = Split(vItem(1), "  ")
                    lbEUnits.List(lbEUnits.ListCount - 1, 1) = vTemp(0)
                    lbEUnits.List(lbEUnits.ListCount - 1, 2) = vTemp(1)
                Else
                    lbEUnits.List(lbEUnits.ListCount - 1, 1) = vItem(1)
                    lbEUnits.List(lbEUnits.ListCount - 1, 2) = ""
                End If
            Next i
        End If
        
        vLine = Split(vUnits(UBound(vUnits)), ";;")
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), "=")
            lbUnits.AddItem vItem(0)
            If InStr(vItem(1), "  ") > 0 Then
                vTemp = Split(vItem(1), "  ")
                lbUnits.List(lbUnits.ListCount - 1, 1) = vTemp(0)
                lbUnits.List(lbUnits.ListCount - 1, 2) = vTemp(1)
            Else
                lbUnits.List(lbUnits.ListCount - 1, 1) = vItem(1)
                lbUnits.List(lbUnits.ListCount - 1, 2) = ""
            End If
        Next i
    End If
    
    cbAddUnit.Enabled = True
    cbERemovetUnits.Enabled = True
    cbGetUnits.Enabled = True
    cbUpdateBlock.Enabled = True
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbUpdateBlock_Click()
    Dim strExisting, strProposed As String
    Dim vAttList As Variant
    Dim iAtt As Integer
    
    strExisting = ""
    strProposed = ""
    
    If lbEUnits.ListCount < 1 Then GoTo Add_Proposed
    
    strExisting = lbEUnits.List(0, 0) & "=" & lbEUnits.List(0, 1)
    If Not lbEUnits.List(0, 2) = "" Then strExisting = strExisting & "  " & lbEUnits.List(0, 2)
    
    If lbEUnits.ListCount > 1 Then
        For i = 1 To lbEUnits.ListCount - 1
            strExisting = strExisting & ";;" & lbEUnits.List(i, 0) & "=" & lbEUnits.List(i, 1)
            If Not lbEUnits.List(i, 2) = "" Then strExisting = strExisting & "  " & lbEUnits.List(i, 2)
        Next i
    End If
    
Add_Proposed:
    
    If lbUnits.ListCount < 1 Then GoTo Save_Block
    
    strProposed = lbUnits.List(0, 0) & "=" & lbUnits.List(0, 1)
    If Not lbUnits.List(0, 2) = "" Then strProposed = strProposed & "  " & lbUnits.List(0, 2)
    
    If lbUnits.ListCount > 1 Then
        For i = 1 To lbUnits.ListCount - 1
            strProposed = strProposed & ";;" & lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
            If Not lbUnits.List(i, 2) = "" Then strProposed = strProposed & "  " & lbUnits.List(i, 2)
        Next i
    End If
    
Save_Block:
    
    If strExisting = "" Then
        strExisting = strProposed
    Else
        strExisting = strExisting & " <-- " & strProposed
    End If
    
    vAttList = objBlock.GetAttributes
    
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 27
        Case Else
            iAtt = 7
    End Select
    
    vAttList(iAtt).TextString = strExisting
    
    objBlock.Update
End Sub

Private Sub lbEUnits_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbEUnits.ListCount < 1 Then Exit Sub
    
    Dim strUnit, strQ, strNote As String
    Dim i, i2 As Integer
    
    Select Case KeyCode
        Case vbKeyDelete
            lbEUnits.RemoveItem lbEUnits.ListIndex
        Case vbKeyDown
            If lbEUnits.ListIndex = (lbEUnits.ListCount - 1) Then Exit Sub
            i = lbEUnits.ListIndex
            i2 = i + 1
            
            strUnit = lbEUnits.List(i, 0)
            strQ = lbEUnits.List(i, 1)
            strNote = lbEUnits.List(i, 2)
            
            lbEUnits.List(i, 0) = lbEUnits.List(i2, 0)
            lbEUnits.List(i, 1) = lbEUnits.List(i2, 1)
            lbEUnits.List(i, 2) = lbEUnits.List(i2, 2)
            
            lbEUnits.List(i2, 0) = strUnit
            lbEUnits.List(i2, 1) = strQ
            lbEUnits.List(i2, 2) = strNote
        Case vbKeyUp
            If lbEUnits.ListIndex = 0 Then Exit Sub
            i = lbEUnits.ListIndex
            i2 = i - 1
            
            strUnit = lbEUnits.List(i, 0)
            strQ = lbEUnits.List(i, 1)
            strNote = lbEUnits.List(i, 2)
            
            lbEUnits.List(i, 0) = lbEUnits.List(i2, 0)
            lbEUnits.List(i, 1) = lbEUnits.List(i2, 1)
            lbEUnits.List(i, 2) = lbEUnits.List(i2, 2)
            
            lbEUnits.List(i2, 0) = strUnit
            lbEUnits.List(i2, 1) = strQ
            lbEUnits.List(i2, 2) = strNote
    End Select
End Sub

Private Sub lbUnits_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbUnit.Value = lbUnits.List(lbUnits.ListIndex, 0)
    tbQuantity.Value = lbUnits.List(lbUnits.ListIndex, 1)
    tbUNote.Value = lbUnits.List(lbUnits.ListIndex, 2)
    
    cbAddUnit.Caption = "Update"
End Sub

Private Sub lbUnits_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbUnits.ListCount < 1 Then Exit Sub
    
    Dim strUnit, strQ, strNote As String
    Dim i, i2 As Integer
    
    Select Case KeyCode
        Case vbKeyDelete
            lbUnits.RemoveItem lbUnits.ListIndex
        Case vbKeyDown
            If lbUnits.ListIndex = (lbUnits.ListCount - 1) Then Exit Sub
            i = lbUnits.ListIndex
            i2 = i + 1
            
            strUnit = lbUnits.List(i, 0)
            strQ = lbUnits.List(i, 1)
            strNote = lbUnits.List(i, 2)
            
            lbUnits.List(i, 0) = lbUnits.List(i2, 0)
            lbUnits.List(i, 1) = lbUnits.List(i2, 1)
            lbUnits.List(i, 2) = lbUnits.List(i2, 2)
            
            lbUnits.List(i2, 0) = strUnit
            lbUnits.List(i2, 1) = strQ
            lbUnits.List(i2, 2) = strNote
        Case vbKeyUp
            If lbUnits.ListIndex = 0 Then Exit Sub
            i = lbUnits.ListIndex
            i2 = i - 1
            
            strUnit = lbUnits.List(i, 0)
            strQ = lbUnits.List(i, 1)
            strNote = lbUnits.List(i, 2)
            
            lbUnits.List(i, 0) = lbUnits.List(i2, 0)
            lbUnits.List(i, 1) = lbUnits.List(i2, 1)
            lbUnits.List(i, 2) = lbUnits.List(i2, 2)
            
            lbUnits.List(i2, 0) = strUnit
            lbUnits.List(i2, 1) = strQ
            lbUnits.List(i2, 2) = strNote
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbUnits.ColumnCount = 3
    lbUnits.ColumnWidths = "96;36;44"
    
    lbEUnits.ColumnCount = 3
    lbEUnits.ColumnWidths = "96;36;44"

End Sub

Private Function GetExistingIndex(strItem As String)
    If lbEUnits.ListCount < 1 Then Exit Function
    
    Dim iIndex As Integer
    
    For iIndex = 0 To lbEUnits.ListCount - 1
        If InStr(lbEUnits.List(iIndex, 0), strItem) > 0 Then GoTo Found_Index
    Next iIndex
    
    iIndex = -1
    
Found_Index:
    
    GetExistingIndex = iIndex
End Function

Private Function GetProposedIndex(strItem As String)
    If lbUnits.ListCount < 1 Then Exit Function
    
    Dim iIndex As Integer
    
    For iIndex = 0 To lbUnits.ListCount - 1
        If InStr(lbUnits.List(iIndex, 0), strItem) > 0 Then GoTo Found_Index
    Next iIndex
    
    iIndex = -1
    
Found_Index:
    
    GetProposedIndex = iIndex
End Function

Private Function GetCoilIndex(strItem As String)
    If lbUnits.ListCount < 1 Then Exit Function
    
    Dim iIndex As Integer
    
    For iIndex = 0 To lbUnits.ListCount - 1
        If InStr(lbUnits.List(iIndex, 0), strItem) > 0 Then
            If InStr(lbUnits.List(iIndex, 2), "LOOP") > 0 Then GoTo Found_Index
        End If
    Next iIndex
    
    iIndex = -1
    
Found_Index:
    
    GetCoilIndex = iIndex
End Function
