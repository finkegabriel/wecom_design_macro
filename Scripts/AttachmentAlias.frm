VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AttachmentAlias 
   Caption         =   "Attachment Aliases"
   ClientHeight    =   8625.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "AttachmentAlias.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AttachmentAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAdd_Click()
    If tbAlias.Value = "" Then Exit Sub
    If tbCompany.Value = "" Then Exit Sub
    
    If cbAdd.Caption = "Update" Then
        lbAlias.Enabled = True
        cbAdd.Caption = "Add Alias"
        
        lbAlias.List(lbAlias.ListIndex, 0) = UCase(tbAlias.Value)
        lbAlias.List(lbAlias.ListIndex, 1) = UCase(tbCompany.Value)
    Else
        lbAlias.AddItem UCase(tbAlias.Value)
        lbAlias.List(lbAlias.ListCount - 1, 1) = UCase(tbCompany.Value)
    End If
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbReplace_Click()
    If lbAlias.ListCount < 1 Then Exit Sub
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine, vItem As Variant
    Dim strTemp As String
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        vAttList = objBlock.GetAttributes
        
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        If vAttList(1).TextString = "INCOMPLETE" Then GoTo Next_objBlock
        
        If Not vAttList(4).TextString = "" Then
            If InStr(vAttList(4).TextString, "=") > 0 Then
                vLine = Split(UCase(vAttList(4).TextString), " ")
                
                For j = 0 To UBound(vLine)
                    vItem = Split(vLine(j), "=")
                    
                    If UBound(vItem) = 0 Then
                        vAttList(4).TextString = vAttList(2).TextString & "=" & vAttList(4).TextString
                        GoTo Next_Other
                    End If
            
                    For k = 0 To lbAlias.ListCount - 1
                        If vItem(0) = lbAlias.List(k, 0) Then
                            vLine(j) = lbAlias.List(k, 1) & "=" & vItem(1)
                            objBlock.Update
                    
                            GoTo Next_Other
                        End If
                    Next k
Next_Other:
                Next j
                
                strTemp = vLine(0)
                If UBound(vLine) > 0 Then
                    For j = 1 To UBound(vLine)
                        strTemp = strTemp & " " & vLine(j)
                    Next j
                End If
                
                vAttList(4).TextString = strTemp
            Else
                vAttList(4).TextString = vAttList(2).TextString & "=" & vAttList(4).TextString
            End If
        End If
        
        For j = 16 To 23
            If vAttList(j).TextString = "" Then GoTo Next_J
            If InStr(vAttList(j).TextString, "=") < 1 Then GoTo Next_J
            
            vLine = Split(UCase(vAttList(j).TextString), "=")
            If UBound(vLine) < 1 Then GoTo Next_J
            
            For k = 0 To lbAlias.ListCount - 1
                If vLine(0) = lbAlias.List(k, 0) Then
                    vAttList(j).TextString = lbAlias.List(k, 1) & "=" & vLine(1)
                    objBlock.Update
                    
                    GoTo Next_J
                End If
            Next k
            
Next_J:
        Next j
Next_objBlock:
    Next i
    
    grpValue(0) = "Existing_Guys"
    filterType = grpCode
    filterValue = grpValue
    
    objSS.Clear
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        vAttList = objBlock.GetAttributes
        
        For j = 0 To 7
            If vAttList(j).TextString = "" Then GoTo Next_Other2
            If InStr(vAttList(j).TextString, "=") > 0 Then
                vItem = Split(UCase(vAttList(j).TextString), "=")
                
                For k = 0 To lbAlias.ListCount - 1
                    If vItem(0) = lbAlias.List(k, 0) Then
                        vAttList(j).TextString = lbAlias.List(k, 1) & "=" & vItem(1)
                        objBlock.Update
                    
                        GoTo Next_Other2
                    End If
                Next k
            End If
Next_Other2:
        Next j
    Next i
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub lbAlias_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbAlias.Value = lbAlias.List(lbAlias.ListIndex, 0)
    tbCompany.Value = lbAlias.List(lbAlias.ListIndex, 1)
    
    cbAdd.Caption = "Update"
    lbAlias.Enabled = False
End Sub

Private Sub lbAlias_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbAlias.ListCount < 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyDelete
            lbAlias.RemoveItem lbAlias.ListIndex
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbAlias.ColumnCount = 2
    lbAlias.ColumnWidths = "72;162"
    
    Call GetAliases
End Sub

Private Sub GetAliases()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    On Error Resume Next
    
    lbAlias.Clear
    
    grpCode(0) = 2
    grpValue(0) = "Alias"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        vAttList = objBlock.GetAttributes
        
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        If vAttList(0).TextString = "alias" Then GoTo Next_objBlock
        
        lbAlias.AddItem UCase(vAttList(0).TextString), 0
        lbAlias.List(0, 1) = UCase(vAttList(1).TextString)
        
Next_objBlock:
    Next i
    
    objSS.Clear
    objSS.Delete
End Sub
