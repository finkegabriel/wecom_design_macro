VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertCSV 
   Caption         =   "Convert CSV to Blocks"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "ConvertCSV.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConvertCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbPlace_Click()
    If tbPasted.Value = "" Then Exit Sub
    If cbLat.ListIndex < 0 Then Exit Sub
    If cbLong.ListIndex < 0 Then Exit Sub
    
    Dim vLine, vItem, vTemp As Variant
    Dim vList As Variant
    Dim vCoords, vNE As Variant
    Dim strLine As String
    Dim strFind, strReplace As String
    Dim dCoords(2) As Double
    Dim iIndex, iLat, iLong As Integer
    
    tbPasted.Value = Replace(tbPasted.Value, """", "")
    strLine = Replace(tbPasted.Value, vbLf, "")
    vItem = Split(strLine, vbCr)
    
    For i = 1 To UBound(vItem)
        If vItem(i) = "" Then GoTo Exit_Loop
        
        vLine = Split(vItem(i), ",")
        
        iLat = cbLat.ListIndex
        iLong = cbLong.ListIndex
        
        vNE = LLtoTN83F(CDbl(vLine(iLat)), CDbl(vLine(iLong)))
        dCoords(0) = CDbl(vNE(1))
        dCoords(1) = CDbl(vNE(0))
        dCoords(2) = 0#
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dCoords, cbToList.Value, 1#, 1#, 1#, 0#)
        If lbTo.ListCount > 0 Then
            vAttList = objBlock.GetAttributes
            
            For j = 0 To lbTo.ListCount - 1
                strLine = lbTo.List(j, 2)
                
                If InStr(strLine, "}") > 0 Then
                    vList = Split(strLine, "}")
                    
                    For k = 0 To UBound(vList)
                        vTemp = Split(vList(k), "{")
                        If UBound(vTemp) < 1 Then GoTo Next_K
                        
                        iIndex = CInt(vTemp(1))
                        strFind = "{" & iIndex & "}"
                        strLine = Replace(strLine, strFind, vLine(iIndex))
Next_K:
                    Next k
                End If
                If Not strLine = "" Then vAttList(j).TextString = strLine
            Next j
        End If
        
        objBlock.Layer = cbToLayer.Value
        objBlock.Update
    Next i
    
Exit_Loop:
    
    MsgBox "Done."
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
            lbTo.List(i, 2) = ""
        Next i
        GoTo Exit_Next
    Next objBlock
Exit_Next:
    
    objSS2.Clear
    objSS2.Delete
End Sub

Private Sub lbTo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If cbColumns.Value = "" Then Exit Sub
    If lbTo.ListCount < 1 Then Exit Sub
    
    Dim vLine As Variant
    
    vLine = Split(cbColumns.Value, " ")
    vLine(0) = "{" & vLine(0) & "}"
    
    If lbTo.List(lbTo.ListIndex, 2) = "" Then
        lbTo.List(lbTo.ListIndex, 2) = vLine(0)
    Else
        lbTo.List(lbTo.ListIndex, 2) = lbTo.List(lbTo.ListIndex, 2) & cbSeperator.Value & vLine(0)
    End If
    
    If lbTo.ListIndex < lbTo.ListCount - 1 Then
        lbTo.ListIndex = lbTo.ListIndex + 1
        lbTo.Selected(lbTo.ListIndex) = True
    End If
End Sub

Private Sub tbPasted_Change()
    If tbPasted.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vTemp As Variant
    Dim strLine As String
    Dim iIndex As Integer
    
    tbPasted.Value = Replace(tbPasted.Value, """", "")
    strLine = Replace(tbPasted.Value, vbLf, "")
    vItem = Split(strLine, vbCr)
    vLine = Split(vItem(0), ",")
    
    For i = 0 To UBound(vLine)
        strLine = i & " - " & vLine(i)
        
        cbLat.AddItem strLine
        cbLong.AddItem strLine
        cbColumns.AddItem strLine
    Next i
End Sub

Private Sub UserForm_Initialize()
    lbTo.Clear
    lbTo.ColumnCount = 3
    lbTo.ColumnWidths = "20;80;70"
    
    cbSeperator.AddItem ","
    cbSeperator.AddItem "-"
    cbSeperator.AddItem " "
    cbSeperator.AddItem ";"
    cbSeperator.Value = ","
    
    On Error Resume Next
    
    Dim objBlocks As AcadBlocks
    Dim strLine As String
            
    Set objBlocks = ThisDrawing.Blocks
    For i = 0 To objBlocks.count - 1
        strLine = objBlocks(i).Name
        If Not Left(strLine, 1) = "*" Then cbToList.AddItem objBlocks(i).Name
    Next i
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        cbToLayer.AddItem objLayer.Name
    Next objLayer
    cbToLayer.Value = "0"
End Sub

Private Function LLtoTN83F(dLat As Double, dLong As Double)
    Dim dDLat As Double
    Dim dEast, dDiffE, dEast0 As Double
    Dim dNorth, dDiffN, dNorth0 As Double
    Dim dU, dR, dCA, dK As Double
    Dim NE(2) As Double
    
    dDLat = dLat - 35.8340607459
    dU = dDLat * (110950.2019 + dDLat * (9.25072 + dDLat * (5.64572 + dDLat * 0.017374)))
    dR = 8842127.1422 - dU
    dCA = ((86 + dLong) * 0.585439726459) * 3.14159265359 / 180
    
    dDiffE = dR * Sin(dCA)
    dDiffN = dU + dDiffE * Tan(dCA / 2)
    
    dEast = (dDiffE + 600000) / 0.3048006096
    dNorth = (dDiffN + 166504.1691) / 0.3048006096
    
    dK = 0.999948401424 + (1.23188E-14 * dU * dU) + (4.54E-22 * dU * dU * dU)
    
    NE(0) = dNorth
    NE(1) = dEast
    NE(2) = dK
    
    LLtoTN83F = NE
End Function
