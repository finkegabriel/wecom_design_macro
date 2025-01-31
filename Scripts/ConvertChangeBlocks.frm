VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertChangeBlocks 
   Caption         =   "Bulk Change Blocks and Layers"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9090.001
   OleObjectBlob   =   "ConvertChangeBlocks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConvertChangeBlocks"
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
    
    Dim objNewBlock As AcadBlockReference
    Dim vNewBlockAtt As Variant
    Dim dCoords(0 To 2) As Double
    Dim dScale As Double
    
    Dim iAttList As Integer
    Dim strValue, strBlock As String
    Dim strLayer As String
    
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
        dScale = objBlock.XScaleFactor
        dCoords(0) = objBlock.InsertionPoint(0)
        dCoords(1) = objBlock.InsertionPoint(1)
        dCoords(2) = objBlock.InsertionPoint(2)
        
        For i = 0 To lbConditions.ListCount - 1
            iAttList = CInt(lbConditions.List(i, 0))
            strValue = lbConditions.List(i, 1)
            strBlock = lbConditions.List(i, 2)
            strLayer = lbConditions.List(i, 3)
            
            If Left(strValue, 1) = "{" Then
                strValue = Replace(strValue, "{", "")
                strValue = Replace(strValue, "}", "")
            End If
            
            If InStr(UCase(vBlockAtt(iAttList).TextString), UCase(strValue)) Then
            'If vBlockAtt(iAttList).TextString = strValue Then
                Set objNewBlock = ThisDrawing.ModelSpace.InsertBlock(dCoords, strBlock, dScale, dScale, dScale, 0#)
                vNewBlockAtt = objNewBlock.GetAttributes
                
                If UBound(vBlockAtt) < UBound(vNewBlockAtt) Then
                    For m = 0 To UBound(vBlockAtt)
                        vNewBlockAtt(m).TextString = vBlockAtt(m).TextString
                    Next m
                Else
                    For m = 0 To UBound(vNewBlockAtt)
                        vNewBlockAtt(m).TextString = vBlockAtt(m).TextString
                    Next m
                End If
                
                objNewBlock.Layer = strLayer
                objNewBlock.Update
                
                objBlock.Delete
                GoTo Exit_Next
            End If
        Next i
Exit_Next:
    Next objBlock
    MsgBox "Done"
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbToBlock_Change()
    Select Case cbToBlock.Value
        Case "RES"
            cbToLayer.Value = "Integrity Building-RES"
        Case "BUSINESS"
            cbToLayer.Value = "Integrity Building-BUS"
        Case "TRLR"
            cbToLayer.Value = "Integrity Building-TRL"
        Case "MDU"
            cbToLayer.Value = "Integrity Building-MDU"
        Case "SCHOOL"
            cbToLayer.Value = "Integrity Building-SCH"
        Case "CHURCH"
            cbToLayer.Value = "Integrity Building-CHU"
        Case "EXTENTION", "NONRES"
            cbToLayer.Value = "Integrity Building Misc"
    End Select
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
    If tbInStr.Value = "" Then
        lbConditions.List(lbConditions.ListCount - 1, 1) = cbValue.Value
    Else
        lbConditions.List(lbConditions.ListCount - 1, 1) = "{" & tbInStr.Value & "}"
    End If
    'lbConditions.List(lbConditions.ListCount - 1, 1) = cbValue.Value
    lbConditions.List(lbConditions.ListCount - 1, 2) = cbToBlock.Value
    lbConditions.List(lbConditions.ListCount - 1, 3) = cbToLayer.Value
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
    lbConditions.ColumnCount = 4
    lbConditions.ColumnWidths = "20;80;100;100"
    
    lbTo.Clear
    lbTo.ColumnCount = 2
    lbTo.ColumnWidths = "20;80"
    
    Dim objBlocks As AcadBlocks
    Dim strLine As String
            
    Set objBlocks = ThisDrawing.Blocks
    For i = 0 To objBlocks.count - 1
        strLine = objBlocks(i).Name
        If Not Left(strLine, 1) = "*" Then
            cbToList.AddItem objBlocks(i).Name
            cbToBlock.AddItem objBlocks(i).Name
        End If
    Next i
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        cbToLayer.AddItem objLayer.Name
    Next objLayer
    cbToLayer.Value = "0"
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
