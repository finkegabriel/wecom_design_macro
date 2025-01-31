VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCableToBuried 
   Caption         =   "Add Buried Cable Units"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "AddCableToBuried.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCableToBuried"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objPlant As AcadBlockReference

Private Sub cbAddUnit_Click()
    If cbAddUnit.Caption = "Update" Then
        lbUnits.List(lbUnits.ListIndex, 0) = tbUnit.Value
        lbUnits.List(lbUnits.ListIndex, 1) = tbQuantity.Value
        lbUnits.List(lbUnits.ListIndex, 2) = tbUNote.Value
        
        cbAddUnit.Caption = "Add Unit"
        GoTo Exit_Sub
    End If
    
    lbUnits.AddItem tbUnit.Value
    lbUnits.List(lbUnits.ListCount - 1, 1) = tbQuantity.Value
    lbUnits.List(lbUnits.ListCount - 1, 2) = tbUNote.Value
    
Exit_Sub:
    
    tbUnit.Value = ""
    tbQuantity.Value = ""
    tbUNote.Value = ""
End Sub

Private Sub cbCblType_Change()
    If cbCblType.Value = "BFO" Then
        cbSuffix.AddItem ""
        cbSuffix.AddItem "I"
        cbSuffix.AddItem "E(36)"
    Else
        cbSuffix.Clear
    End If
End Sub

Private Sub cbGetBlock_Click()
    Dim objEntity As AcadEntity
    Dim vReturnPnt, vAttList As Variant
    
    Me.Hide
    
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Buried Plant: "
    
    If Not Err = 0 Then
        MsgBox "Not a valid selection"
        Me.show
        Exit Sub
    End If
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox "Not a valid selection"
        Me.show
        Exit Sub
    End If
    
    Set objPlant = objEntity
    Select Case objPlant.Name
        Case "sPed", "sHH"
        Case Else
            MsgBox "Not a valid selection"
            Me.show
            Exit Sub
    End Select
    
    vAttList = objPlant.GetAttributes
    
    tbNumber.Value = vAttList(0).TextString
    
    lbUnits.Clear
    
    If Not vAttList(7).TextString = "" Then
        vAttList(7).TextString = Replace(vAttList(7).TextString, vbLf, "")
        vLine = Split(vAttList(7).TextString, ";;")
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
    
    cbUpdateBlock.Enabled = True
    
    Me.show
End Sub

Private Sub cbGetCableSpan_Click()
    If cbCblType.Value = "" Then
        MsgBox "Need a Cable Type."
        Exit Sub
    End If
    
    If cbCableSize.Value = "" Then
        MsgBox "Need a Cable Size."
        Exit Sub
    End If
    
    cbGetUnits.Enabled = True
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt, vAttList As Variant
    Dim strTemp, strLine As String
    'Dim iQuantity As Integer
    
    Me.Hide
    
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Cable Span Block: "
    
    If Not Err = 0 Then
        MsgBox "Not a valid selection"
        Me.show
        Exit Sub
    End If
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox "Not a valid selection"
        Me.show
        Exit Sub
    End If
    
    Set objBlock = objEntity
    Select Case objBlock.Name
        Case "cable_span"
        Case Else
            MsgBox "Not a valid selection"
            Me.show
            Exit Sub
    End Select
    
    vAttList = objBlock.GetAttributes
    
    strTemp = Replace(vAttList(2).TextString, "'", "")
    tbSpan.Value = strTemp
    'iQuantity = CInt(strTemp)
    
    strLine = "+" & cbCblType.Value & "(" & cbCableSize.Value & ")"
    lbUnits.AddItem strLine, 0
    lbUnits.List(0, 1) = strTemp
    
    Me.show
End Sub

Private Sub cbGetUnits_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strLine As String
    
    On Error Resume Next
    
    Me.Hide
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.SelectOnScreen
    If objSS.count < 1 Then GoTo Exit_Sub
    
    For Each objEntity In objSS
        If Not TypeOf objEntity Is AcadBlockReference Then GoTo Next_objEntity
        
        Set objBlock = objEntity
        
        Select Case objBlock.Name
            Case "Map coil"
                vAttList = objBlock.GetAttributes
                
                strLine = "+" & cbCblType.Value & "(" & Replace(vAttList(1).TextString, "F", "") & ")"
                lbUnits.AddItem strLine, 1
                
                strLine = Replace(vAttList(0).TextString, "'", "")
                lbUnits.List(1, 1) = strLine
                lbUnits.List(1, 2) = "LOOP"
            Case "Map splice"
                vAttList = objBlock.GetAttributes
                
                strLine = "+HBFO(" & cbCableSize.Value & ")"
                lbUnits.AddItem strLine, 2
                lbUnits.List(2, 1) = "1"
                lbUnits.List(2, 2) = vAttList(0).TextString
        End Select
        
Next_objEntity:
    Next objEntity
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbUpdateBlock_Click()
    If lbUnits.ListCount < 1 Then Exit Sub
    
    Dim vAttList As Variant
    Dim strLine As String
    
    vAttList = objPlant.GetAttributes
    
    strLine = lbUnits.List(0, 0) & "=" & lbUnits.List(0, 1)
    If Not lbUnits.List(0, 2) = "" Then strLine = strLine & "  " & lbUnits.List(0, 2)
    
    If lbUnits.ListCount > 1 Then
        For i = 1 To lbUnits.ListCount - 1
            strLine = strLine & ";;" & lbUnits.List(i, 0) & "=" & lbUnits.List(i, 1)
            If Not lbUnits.List(i, 2) = "" Then strLine = strLine & "  " & lbUnits.List(i, 2)
        Next i
    End If
    
    vAttList(7).TextString = strLine
    
    objPlant.Update
End Sub

Private Sub LblRRUD_Click()
    tbUnit.Value = "+RRCONDUIT"
    tbQuantity.SetFocus
    
    If tbSpan.Value = "" Then Exit Sub
    
    If cbCblType.Value = "UO" Then
        tbQuantity.Value = tbSpan.Value
        cbAddUnit.SetFocus
    End If
End Sub

Private Sub lbUnits_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cbAddUnit.Caption = "Update"
    
    tbUnit.Value = lbUnits.List(lbUnits.ListIndex, 0)
    tbQuantity.Value = lbUnits.List(lbUnits.ListIndex, 1)
    tbUNote.Value = lbUnits.List(lbUnits.ListIndex, 2)
End Sub

Private Sub lbUnits_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim str0, str1, str2 As String
    
    Select Case KeyCode
        Case vbKeyDelete
            Dim strUnit, strQuantity, strNote As String
            Dim iIndex, iQuantity As Integer
            
            iIndex = lbUnits.ListIndex
            strUnit = lbUnits.List(iIndex, 0)
            strQuantity = lbUnits.List(iIndex, 1)
            iQuantity = CInt(strQuantity)
            If Not lbUnits.List(iIndex, 2) = "" Then strNote = lbUnits.List(iIndex, 2)
            
            lbUnits.RemoveItem lbUnits.ListIndex
            
            If InStr(strUnit, "BM60") > 0 Then
                
            End If
        Case vbKeyUp
            If lbUnits.ListIndex = 0 Then Exit Sub
            
            str0 = lbUnits.List(lbUnits.ListIndex, 0)
            str1 = lbUnits.List(lbUnits.ListIndex, 1)
            str2 = lbUnits.List(lbUnits.ListIndex, 2)
            
            lbUnits.List(lbUnits.ListIndex, 0) = lbUnits.List(lbUnits.ListIndex - 1, 0)
            lbUnits.List(lbUnits.ListIndex, 1) = lbUnits.List(lbUnits.ListIndex - 1, 1)
            lbUnits.List(lbUnits.ListIndex, 2) = lbUnits.List(lbUnits.ListIndex - 1, 2)
            
            lbUnits.List(lbUnits.ListIndex - 1, 0) = str0
            lbUnits.List(lbUnits.ListIndex - 1, 1) = str1
            lbUnits.List(lbUnits.ListIndex - 1, 2) = str2
            
            'lbUnits.ListIndex = lbUnits.ListIndex - 1
        Case vbKeyDown
            If lbUnits.ListIndex = lbUnits.ListCount - 1 Then Exit Sub
            
            str0 = lbUnits.List(lbUnits.ListIndex, 0)
            str1 = lbUnits.List(lbUnits.ListIndex, 1)
            str2 = lbUnits.List(lbUnits.ListIndex, 2)
            
            lbUnits.List(lbUnits.ListIndex, 0) = lbUnits.List(lbUnits.ListIndex + 1, 0)
            lbUnits.List(lbUnits.ListIndex, 1) = lbUnits.List(lbUnits.ListIndex + 1, 1)
            lbUnits.List(lbUnits.ListIndex, 2) = lbUnits.List(lbUnits.ListIndex + 1, 2)
            
            lbUnits.List(lbUnits.ListIndex + 1, 0) = str0
            lbUnits.List(lbUnits.ListIndex + 1, 1) = str1
            lbUnits.List(lbUnits.ListIndex + 1, 2) = str2
            
            'lbUnits.ListIndex = lbUnits.ListIndex + 1
    End Select
End Sub

Private Sub UserForm_Initialize()
    cbCblType.AddItem ""
    cbCblType.AddItem "BFO"
    cbCblType.AddItem "UO"
    
    cbCableSize.AddItem ""
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
    
    lbUnits.ColumnCount = 3
    lbUnits.ColumnWidths = "96;36;44"
End Sub
