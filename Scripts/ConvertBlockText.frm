VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertBlockText 
   Caption         =   "Change Block Text"
   ClientHeight    =   5985
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   7110
   OleObjectBlob   =   "ConvertBlockText.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConvertBlockText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbClear_Click()
    lbConditions.Clear
    cbValue.Clear
    cbValue.Value = ""
    lAttTag.Caption = "Attribute Tag"
    cbUpdate.Enabled = False
    lbTo.Enabled = True
    cbToList.Enabled = True
End Sub

Private Sub cbConvert_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS2 As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vBlockAtt As Variant
    Dim vLine, vItem As String
    Dim strType, strFind As String
    Dim iAttList, iCopy As Integer
    Dim strValue, strLayer As String
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = cbToList.Value
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
    If Not Err = 0 Then
        Set objSS2 = ThisDrawing.SelectionSets.Item("objSS2")
        objSS2.Clear
    End If
    
    objSS2.Select acSelectionSetAll, , , filterType, filterValue
    MsgBox cbToList.Value & " found:  " & objSS2.count
    
    For Each objBlock In objSS2
        vBlockAtt = objBlock.GetAttributes
        
        For i = 0 To lbConditions.ListCount - 1
            iAttList = CInt(lbConditions.List(i, 0))
            
            strValue = lbConditions.List(i, 1)
            If Not strValue = vBlockAtt(iAttList).TextString Then
                If Not strValue = "<<all>>" Then GoTo Next_objBlock
            End If
            If strValue = "<<blank>>" Then strValue = ""
            
            strLayer = lbConditions.List(i, 2)
            If strLayer = "<<blank>>" Then strLayer = ""
            
            strType = "R"
            
            If InStr(strLayer, "++") > 0 Then
                If Left(strLayer, 2) = "++" Then
                    strType = "AE"
                Else
                    strType = "AB"
                End If
                
                strLayer = Replace(strLayer, "++", "")
            End If
            
            If InStr(strLayer, "--") > 0 Then
                strType = "MM"
                strLayer = Replace(strLayer, "--", "")
            End If
            
            'If InStr(strLayer, "}") > 0 Then
                'strType = "RA"
                'vLine = Split(strLayer, "}")
                'vItem = Split(vLine(0), "{")
                'iCopy = CInt(vItem(1))
                'strFind = "{" & iCopy & "}"
                
                'strLayer = Replace(strLayer, strFind, vBlockAtt(iCopy).TextString)
            'End If
            
            Select Case strType
                Case "AB"
                    vBlockAtt(iAttList).TextString = strLayer & vBlockAtt(iAttList).TextString
                Case "AE"
                    vBlockAtt(iAttList).TextString = vBlockAtt(iAttList).TextString & strLayer
                Case "MM"
                    vBlockAtt(iAttList).TextString = Replace(vBlockAtt(iAttList).TextString, strLayer, "")
                'Case "RA"
                    'vBlockAtt(iAttList).TextString = strLayer
                Case Else
                    If strValue = "" Then
                        vBlockAtt(iAttList).TextString = strLayer
                    Else
                        If InStr(vBlockAtt(iAttList).TextString, strValue) > 0 Then vBlockAtt(iAttList).TextString = Replace(vBlockAtt(iAttList).TextString, strValue, strLayer)
                    End If
            End Select
            
        Next i
Next_objBlock:
    Next objBlock
    lbConditions.Clear
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbToList_Change()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS2 As AcadSelectionSet
    
    On Error Resume Next
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    grpCode(0) = 2
    grpValue(0) = cbToList.Value
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
    If Not Err = 0 Then
        Set objSS2 = ThisDrawing.SelectionSets.Item("objSS2")
        objSS2.Clear
    End If
    
    objSS2.Select acSelectionSetAll, , , filterType, filterValue
        
    For Each objBlock In objSS2
        vAttList = objBlock.GetAttributes
        lbTo.Clear
        For i = 0 To UBound(vAttList)
            lbTo.AddItem
            lbTo.List(i, 0) = i
            lbTo.List(i, 1) = vAttList(i).TagString
            'lbTo.List(i, 2) = ""
        Next i
        GoTo Exit_Next
    Next objBlock
Exit_Next:
End Sub

Private Sub cbUpdate_Click()
    lbConditions.AddItem
    lbConditions.List(lbConditions.ListCount - 1, 0) = lbTo.List(lbTo.ListIndex, 0)
    If tbChange.Value = "" Then
        lbConditions.List(lbConditions.ListCount - 1, 1) = "<<blank>>"
    Else
        lbConditions.List(lbConditions.ListCount - 1, 1) = cbValue.Value
    End If
    If tbChange.Value = "" Then
        lbConditions.List(lbConditions.ListCount - 1, 2) = "<<blank>>"
    Else
        lbConditions.List(lbConditions.ListCount - 1, 2) = tbChange.Value
    End If
End Sub

Private Sub lbConditions_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            lbConditions.RemoveItem lbConditions.ListIndex
    End Select
End Sub

Private Sub lbTo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    lAttTag.Caption = lbTo.List(lbTo.ListIndex, 1)
    'cbValue.Value = lbTo.List(lbTo.ListIndex, 2)
    cbUpdate.Enabled = True
    lbTo.Enabled = False
    cbToList.Enabled = False
    cbValue.SetFocus
    
    Call GetValues
End Sub

Private Sub UserForm_Initialize()
    lbConditions.Clear
    lbConditions.ColumnCount = 3
    lbConditions.ColumnWidths = "20;80;100"
    
    lbTo.Clear
    lbTo.ColumnCount = 2
    lbTo.ColumnWidths = "20;80"
    
    Dim objBlocks As AcadBlocks
    Dim strLine As String
            
    Set objBlocks = ThisDrawing.Blocks
    For i = 0 To objBlocks.count - 1
        strLine = objBlocks(i).Name
        If Not Left(strLine, 1) = "*" Then cbToList.AddItem objBlocks(i).Name
    Next i
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    'Set objLayers = ThisDrawing.Layers
    'For Each objLayer In objLayers
        'cbToLayer.AddItem objLayer.Name
    'Next objLayer
    'cbToLayer.Value = "0"
End Sub

Private Sub GetValues()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS2 As AcadSelectionSet
    
    Dim iCount As Integer
    Dim strTest As String
    
    On Error Resume Next
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    grpCode(0) = 2
    grpValue(0) = cbToList.Value
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
    If Not Err = 0 Then
        Set objSS2 = ThisDrawing.SelectionSets.Item("objSS2")
        objSS2.Clear
    End If
    
    objSS2.Select acSelectionSetAll, , , filterType, filterValue
    
    cbValue.Clear
    cbValue.AddItem "<<all>>"
        
    For Each objBlock In objSS2
        vAttList = objBlock.GetAttributes
        strTest = vAttList(lbTo.ListIndex).TextString
        iCount = 0
        
        For i = 0 To cbValue.ListCount - 1
            If strTest = cbValue.List(i) Then GoTo Exit_Next_I
        Next i
        
        cbValue.AddItem strTest
Exit_Next_I:
    Next objBlock
End Sub
