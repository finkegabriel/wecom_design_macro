VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetCableLengths 
   Caption         =   "Get Total Cable Lengths"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4830
   OleObjectBlob   =   "GetCableLengths.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GetCableLengths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbBlock_Click()
    If cbBlock.Value = True Then
        cbLines.Value = False
        cbPolylines.Value = False
    End If
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRun_Click()
    tbTotalFeet.Value = "0"
    tbLines.Value = "0"
    tbPolylines.Value = "0"
    tbBlock.Value = "0"
    
    Me.Hide
    
    For i = 0 To lbLayers.ListCount - 1
        If lbLayers.Selected(i) Then Call GetData(CStr(lbLayers.List(i)))
    Next i
    
    Me.show
End Sub

Private Sub Label1_Click()
    Dim objEntity As AcadEntity
    Dim vReturnPnt As Variant
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Object on Layer:"
    
    For i = 0 To lbLayers.ListCount - 1
        lbLayers.Selected(i) = False
        If lbLayers.List(i) = objEntity.Layer Then lbLayers.Selected(i) = True
    Next i
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        lbLayers.AddItem objLayer.Name
    Next objLayer
    
    Call SortList
    
    For i = 0 To lbLayers.ListCount - 1
        If lbLayers.List(i) = "Integrity Cable-Aerial" Then lbLayers.Selected(i) = True
    Next i
    
    cbSelect.AddItem "Select All"
    cbSelect.AddItem "Window"
    cbSelect.AddItem "Existing Polygon"
    cbSelect.Value = "Existing Polygon"
End Sub

Private Sub GetData(strLayer As String)
    Dim objSS As AcadSelectionSet
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim filterType, filterValue As Variant
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim objLine As AcadLine
    Dim objLWP As AcadLWPolyline
    Dim vAttList, vLine As Variant
    Dim vReturnPnt As Variant
    Dim vCoords, vArray As Variant
    Dim strTemp As String
    Dim dCoords() As Double
    Dim iTemp, iCounter As Integer
    Dim lLine, lLWP, lBlock, lTotal As Long

    lLine = CLng(tbLines.Value)
    lLWP = CLng(tbPolylines.Value)
    lBlock = CLng(tbBlock.Value)
    lTotal = CLng(tbTotalFeet.Value) * 1000
    
    grpCode(0) = 8
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Add("objSS")
        Err = 0
    End If
    
    Select Case cbSelect.Value
        Case "Select All"
            objSS.Select acSelectionSetAll, , , filterType, filterValue
        Case "Window"
            objSS.SelectOnScreen filterType, filterValue
        Case "Existing Polygon"
            ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Polygon: "
            If Not objEntity.ObjectName = "AcDbPolyline" Then
                MsgBox "Error: Invalid Selection."
                GoTo Exit_Sub
            End If
    
            Set objLWP = objEntity
            vCoords = objLWP.Coordinates
    
            iTemp = (UBound(vCoords) + 1) / 2 * 3 - 1
            ReDim dCoords(iTemp) As Double
    
            iCounter = 0
            For i = 0 To UBound(vCoords) Step 2
                dCoords(iCounter) = vCoords(i)
                iCounter = iCounter + 1
                dCoords(iCounter) = vCoords(i + 1)
                iCounter = iCounter + 1
                dCoords(iCounter) = 0#
                iCounter = iCounter + 1
            Next i
    
            objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords, filterType, filterValue
    End Select
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadLWPolyline Then
            Set objLWP = objEntity
            
            lLWP = lLWP + CLng(objLWP.Length)
        ElseIf TypeOf objEntity Is AcadLine Then
            Set objLine = objEntity
            
            lLine = lLine + CLng(objLine.Length)
        ElseIf TypeOf objEntity Is AcadBlockReference Then
            If cbBlock.Value = False Then GoTo Next_objEntity
            
            Set objBlock = objEntity
            If Not objBlock.Name = "cable_span" Then GoTo Next_objEntity
            
            vAttList = objBlock.GetAttributes
            vLine = Split(vAttList(2).TextString, " ")
            strTemp = Replace(vLine(UBound(vLine)), "'", "")
                
            lBlock = lBlock + CLng(strTemp)
        End If
        
Next_objEntity:
    Next objEntity
    
    If cbLines.Value = True Then
        tbLines.Value = lLine
        lTotal = lTotal + lLine
    Else
        tbLines.Value = "0"
    End If
    
    If cbPolylines.Value = True Then
        tbPolylines.Value = lLWP
        lTotal = lTotal + lLWP
    Else
        tbPolylines.Value = "0"
    End If
    
    If cbBlock.Value = True Then
        tbBlock.Value = lBlock
        lTotal = lTotal + lBlock
    Else
        tbBlock.Value = "0"
    End If
    
    tbTotalFeet.Value = lTotal / 1000
    tbTotalMiles.Value = CInt(lTotal / 5.28 + 0.5) / 1000
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
End Sub

Private Sub SortList()
    Dim strLayer As String
    Dim iCount As Integer
    
    iCount = lbLayers.ListCount - 1
    
    For a = iCount To 1 Step -1
        For b = 0 To a - 1
            If lbLayers.List(b) > lbLayers.List(b + 1) Then
                strLayer = lbLayers.List(b + 1)
                
                lbLayers.List(b + 1) = lbLayers.List(b)
                
                lbLayers.List(b) = strLayer
            End If
        Next b
    Next a
End Sub
