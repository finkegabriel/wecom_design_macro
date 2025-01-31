VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateTDOT 
   Caption         =   "Create TDOT Permits"
   ClientHeight    =   9495.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6210
   OleObjectBlob   =   "CreateTDOT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateTDOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCopyLayers_Click()
    Call CopyLayers
    
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Call CurrentLayers
    
    Me.show
End Sub

Private Sub cbCurrent_Click()
    'Call CurrentLayers
End Sub

Private Sub cbExport_Click()
    If tbFile.Value = "" Then Exit Sub
    
    Dim strText, strFileName As String
    
    strFileName = "C:\Integrity\VBA\References\"
    strFileName = strFileName & tbFile.Value & ".plt"
    
    strText = "<<LAYERS>>" & vbCr
    
    If lbLayers.ListCount > 0 Then
        For i = 0 To lbLayers.ListCount - 1
            If lbLayers.List(i, 2) = "YES" Then
                strText = strText & lbLayers.List(i, 0) & vbCr
            End If
        Next i
    End If
    
    strText = strText & vbCr & "<<BLOCKS>>" & vbCr
    
    If lbBlocks.ListCount > 0 Then
        For i = 0 To lbBlocks.ListCount - 1
            If lbBlocks.List(i, 1) = "YES" Then
                strText = strText & lbBlocks.List(i, 0) & vbCr
            End If
        Next i
    End If
    
    Open strFileName For Output As #1
    
    Print #1, strText
    
    Close #1
End Sub

Private Sub cbImport_Click()
    If lbFiles.ListCount < 1 Then Exit Sub
    If lbFiles.ListIndex < 0 Then Exit Sub
    
    tbFile.Value = lbFiles.List(lbFiles.ListIndex)
    
    Dim strLine, strFileName As String
    Dim strLayers, strBlocks As String
    Dim vLayers, vBlocks As Variant
    Dim fName As String
    Dim iStatus As Integer
    
    strFileName = "C:\Integrity\VBA\References\"
    strFileName = strFileName & lbFiles.List(lbFiles.ListIndex) & ".plt"
    
    iStatus = 0
    strLayers = ""
    strBlocks = ""
    
    fName = Dir(strFileName)
    If fName = "" Then
        Exit Sub
    End If
    
    Open strFileName For Input As #2
    
    While Not EOF(2)
        Input #2, strLine
        
        Select Case strLine
            Case "<<LAYERS>>"
                iStatus = 1
            Case "<<BLOCKS>>"
                iStatus = 2
            Case Else
                If iStatus = 1 Then
                    If strLayers = "" Then
                        strLayers = strLine
                    Else
                        strLayers = strLayers & vbTab & strLine
                    End If
                Else
                    If strBlocks = "" Then
                        strBlocks = strLine
                    Else
                        strBlocks = strBlocks & vbTab & strLine
                    End If
                End If
        End Select
Next_line:
    Wend
    
    Close #2
    
    If Not strLayers = "" Then
        vLayers = Split(strLayers, vbTab)
        
        If lbLayers.ListCount > 0 Then
            For i = 0 To lbLayers.ListCount - 1
                For j = 0 To UBound(vLayers)
                    If lbLayers.List(i, 0) = vLayers(j) Then
                        lbLayers.List(i, 2) = "YES"
                        GoTo Next_Layer
                    End If
                Next j
Next_Layer:
            Next i
        End If
    End If
    
    If Not strBlocks = "" Then
        vBlocks = Split(strBlocks, vbTab)
        
        If lbBlocks.ListCount > 0 Then
            For i = 0 To lbBlocks.ListCount - 1
                For j = 0 To UBound(vBlocks)
                    If lbBlocks.List(i, 0) = vBlocks(j) Then
                        lbBlocks.List(i, 1) = "YES"
                        GoTo Next_Block
                    End If
                Next j
Next_Block:
            Next i
        End If
    End If
End Sub

Private Sub cbMaptrim_Click()
    Dim objSSM1 As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objEntTemp As AcadEntity
    Dim objRemove(0) As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objLWPBorder As AcadLWPolyline
    Dim objSSBlock As AcadBlockReference
    Dim objBlock As AcadBlockReference
    Dim objLayer As AcadLayer
    Dim strCommand As String
    Dim strHandle As String
    Dim strRemove As String
    Dim strBlocks As String
    Dim vBlocks As Variant
    Dim vReturnPnt As Variant
    Dim vCoords As Variant
    Dim dLL(0 To 2) As Double
    Dim dUR(0 To 2) As Double
    Dim dLL1(0 To 2) As Double
    Dim dUR1(0 To 2) As Double
    Dim dFrom(0 To 2) As Double
    Dim dTo(0 To 2) As Double
    Dim dDiff(0 To 1) As Double
    Dim dScale As Double
    
  On Error Resume Next
    'Select boundary & find drawing number
    
    strBlocks = ""
    For i = 0 To lbBlocks.ListCount - 1
        If lbBlocks.List(i, 1) = "YES" Then
            If strBlocks = "" Then
                strBlocks = lbBlocks.List(i, 0)
            Else
                strBlocks = strBlocks & vbTab & lbBlocks.List(i, 0)
            End If
        End If
    Next i
    
    If Not strBlocks = "" Then vBlocks = Split(strBlocks, vbTab)
    
    Me.Hide
    
    Err = 0
    Set objSSM1 = ThisDrawing.SelectionSets.Add("objSSM1")
    If Not Err = 0 Then
        Set objSSM1 = ThisDrawing.SelectionSets.Item("objSSM1")
        Err = 0
    End If
    
    Call CopyLayers
    
  While Err = 0
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Trim Border: "
    If Not objEntity.ObjectName = "AcDbPolyline" Then GoTo Exit_Sub
    
    Set objLWP = objEntity
    vCoords = objLWP.Coordinates
    
    dLL(0) = vCoords(0)
    dLL(1) = vCoords(1)
    dUR(0) = vCoords(0)
    dUR(1) = vCoords(1)
    
    For i = 2 To UBound(vCoords)
        If vCoords(i) < dLL(0) Then
            dLL(0) = vCoords(i)
        Else
            If vCoords(i) > dUR(0) Then
                dUR(0) = vCoords(i)
            End If
        End If
        
        i = i + 1
        
        If vCoords(i) < dLL(1) Then
            dLL(1) = vCoords(i)
        Else
            If vCoords(i) > dUR(1) Then
                dUR(1) = vCoords(i)
            End If
        End If
    Next i
    
    dFrom(0) = (dLL(0) + dUR(0)) / 2
    dFrom(1) = (dLL(1) + dUR(1)) / 2
    
    'Get coordinates to DWG Sheet and draw LWP (Save strHandle)
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select DWG Border: "
    If TypeOf objEntity Is AcadBlockReference Then
        Set objSSBlock = objEntity
    Else
        MsgBox "Not a valid entity."
        GoTo Exit_Sub
    End If
    
    dScale = objSSBlock.XScaleFactor
    strRemove = objSSBlock.Layer
    
    dTo(0) = objSSBlock.InsertionPoint(0) + (825 * dScale)
    dTo(1) = objSSBlock.InsertionPoint(1) + (525 * dScale)
    dTo(2) = 0#
    
    dDiff(0) = dTo(0) - dFrom(0)
    dDiff(1) = dTo(1) - dFrom(1)
    
    Set objLWPBorder = ThisDrawing.ModelSpace.AddLightWeightPolyline(vCoords)
    objLWPBorder.Closed = True
    objLWPBorder.Layer = objLWP.Layer
    objLWPBorder.Move dFrom, dTo
    objLWPBorder.Update
    strHandle = objLWPBorder.Handle
    
    'Get objects crossing within dimensions of DWG and copy to DWG
    
    objSSM1.Select acSelectionSetCrossing, dLL, dUR
    
    If objSSM1.count = 0 Then GoTo Exit_Sub
    
    If Not strBlocks = "" Then
        For Each objEntTemp In objSSM1
            If TypeOf objEntTemp Is AcadBlockReference Then
                Set objBlock = objEntTemp
                
                For i = 0 To UBound(vBlocks)
                    If objBlock.Name = vBlocks(i) Then GoTo Next_objEntTemp
                Next i
            
            'If objEntTemp.Layer = strRemove Then
                Set objRemove(0) = objEntTemp
                objSSM1.RemoveItems objRemove
            End If
Next_objEntTemp:
        Next objEntTemp
    End If
    
    strCommand = "COPY" & vbCr & "P" & vbCr & vbCr
    strCommand = strCommand & dFrom(0) & "," & dFrom(1) & ",0" & vbCr
    strCommand = strCommand & dTo(0) & "," & dTo(1) & ",0" & vbCr & vbCr
    
    ThisDrawing.SetVariable "CMDDIA", 0
    ThisDrawing.SendCommand strCommand
    
    objSSM1.Clear
    
    Set objLayer = ThisDrawing.Layers(strRemove)
    objLayer.Lock = True
    
    dLL1(0) = dLL(0) + dDiff(0)
    dLL1(1) = dLL(1) + dDiff(1)
    dUR1(0) = dUR(0) + dDiff(0)
    dUR1(1) = dUR(1) + dDiff(1)
    
    objSSM1.Select acSelectionSetCrossing, dLL1, dUR1
        
    For Each objEntTemp In objSSM1
        If objEntTemp.Layer = strRemove Then
            Set objRemove(0) = objEntTemp
            objSSM1.RemoveItems objRemove
        End If
    Next objEntTemp
    
    strCommand = "_MAPTRIM" & vbCr & "S" & vbCr & "(handent """ & strHandle & """)" & vbCr
    strCommand = strCommand & "N" & vbCr & "Y" & vbCr & "P" & vbCr & vbCr & "O" & vbCr
    strCommand = strCommand & "Y" & vbCr & "Y" & vbCr & "R" & vbCr & "Y" & vbCr

    ThisDrawing.SendCommand strCommand
    
    objSSM1.Clear
    
    objLayer.Lock = False
  Wend
    
Exit_Sub:
    Call CurrentLayers
    
    objSSM1.Delete
    
    ThisDrawing.SetVariable "CMDDIA", 1
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbSelectBlocks_Click()
    If lbBlocks.ListCount < 1 Then Exit Sub
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim strName As String
    
    On Error Resume Next
    
    Me.Hide
    
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    
    strName = objBlock.Name
    
    For i = 0 To lbBlocks.ListCount - 1
        If lbBlocks.List(i, 0) = strName Then
            lbBlocks.List(i, 1) = "YES"
            GoTo Get_Another
        End If
    Next i
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbSelectLayer_Click()
    If lbLayers.ListCount < 1 Then Exit Sub
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim strLayer As String
    Dim strName As String
    
    On Error Resume Next
    
    Me.Hide
    
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select object on Layer:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    strLayer = objEntity.Layer
    
    For i = 0 To lbLayers.ListCount - 1
        If lbLayers.List(i, 0) = strLayer Then
            lbLayers.List(i, 2) = "YES"
            GoTo Exit_Layers
        End If
    Next i
    
Exit_Layers:
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
    
        strName = objBlock.Name
    
        For i = 0 To lbBlocks.ListCount - 1
            If lbBlocks.List(i, 0) = strName Then
                lbBlocks.List(i, 1) = "YES"
                GoTo Get_Another
            End If
        Next i
    End If
    
    GoTo Get_Another
    
Exit_Sub:
    Me.show
End Sub

Private Sub Label3_Click()
    If lbLayers.ListCount < 1 Then Exit Sub
    
    Dim strName, strStatus, strCopy As String
    Dim iIndex, iCount As Integer
    
    iIndex = 0
    
    For iCount = 0 To lbLayers.ListCount - 1
        If lbLayers.List(iCount, 2) = "YES" Then
            strName = lbLayers.List(iCount, 0)
            strStatus = lbLayers.List(iCount, 1)
            strCopy = lbLayers.List(iCount, 2)
            
            lbLayers.RemoveItem iCount
            lbLayers.AddItem strName, iIndex
            lbLayers.List(iIndex, 1) = strStatus
            lbLayers.List(iIndex, 2) = strCopy
            iIndex = iIndex + 1
        End If
    Next iCount
    
    lbLayers.ListIndex = 0
End Sub

Private Sub Label5_Click()
    If lbBlocks.ListCount < 1 Then Exit Sub
    
    Dim strName, strStatus As String
    Dim iIndex, iCount As Integer
    
    iIndex = 0
    
    For iCount = 0 To lbBlocks.ListCount - 1
        If lbBlocks.List(iCount, 1) = "YES" Then
            strName = lbBlocks.List(iCount, 0)
            strStatus = lbBlocks.List(iCount, 1)
            
            lbBlocks.RemoveItem iCount
            lbBlocks.AddItem strName, iIndex
            lbBlocks.List(iIndex, 1) = strStatus
            iIndex = iIndex + 1
        End If
    Next iCount
    
    lbBlocks.ListIndex = 0
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub lbBlocks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Select Case lbBlocks.List(lbBlocks.ListIndex, 1)
        Case "YES"
            lbBlocks.List(lbBlocks.ListIndex, 1) = ""
        Case Else
            lbBlocks.List(lbBlocks.ListIndex, 1) = "YES"
    End Select
End Sub

Private Sub lbLayers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Select Case lbLayers.List(lbLayers.ListIndex, 2)
        Case "YES"
            lbLayers.List(lbLayers.ListIndex, 2) = ""
        Case Else
            lbLayers.List(lbLayers.ListIndex, 2) = "YES"
    End Select
End Sub

Private Sub UserForm_Initialize()
    Dim strTFolder, strFile As String
    Dim vTemp As Variant
    strTFolder = "C:\Integrity\VBA\References\*.*"
    
    strFile = Dir$(strTFolder)
    
    Do While strFile <> ""
        If InStr(strFile, ".plt") Then
            lbFiles.AddItem Replace(strFile, ".plt", "")
        End If
        strFile = Dir$
    Loop

    lbLayers.ColumnCount = 3
    lbLayers.ColumnWidths = "180;60;48"
    
    lbBlocks.ColumnCount = 2
    lbBlocks.ColumnWidths = "126;48"
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        lbLayers.AddItem objLayer.Name
        If objLayer.LayerOn = True Then
            lbLayers.List(lbLayers.ListCount - 1, 1) = "ON"
        Else
            lbLayers.List(lbLayers.ListCount - 1, 1) = "x"
        End If
        lbLayers.List(lbLayers.ListCount - 1, 2) = ""
    Next objLayer
    
    Call SortList
    
    Dim objBlocks As AcadBlocks
    Dim strLine As String
            
    Set objBlocks = ThisDrawing.Blocks
    For i = 0 To objBlocks.count - 1
        strLine = objBlocks(i).Name
        
        If Left(strLine, 1) = "*" Then GoTo Next_Block
        If Left(strLine, 2) = "A$" Then GoTo Next_Block
        
        lbBlocks.AddItem objBlocks(i).Name
        lbBlocks.List(lbBlocks.ListCount - 1, 1) = ""
        
Next_Block:
    Next i
    
    Call SortListBlocks
End Sub

Private Sub SortList()
    Dim a, b As Integer
    Dim iCount As Integer
    Dim strAtt(0 To 2) As String
    
    iCount = lbLayers.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbLayers.List(b, 0) > lbLayers.List(b + 1, 0) Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbLayers.Selected(b) = True
                    lbLayers.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbLayers.List(b + 1, 0)
                strAtt(1) = lbLayers.List(b + 1, 1)
                strAtt(2) = lbLayers.List(b + 1, 2)
                
                lbLayers.List(b + 1, 0) = lbLayers.List(b, 0)
                lbLayers.List(b + 1, 1) = lbLayers.List(b, 1)
                lbLayers.List(b + 1, 2) = lbLayers.List(b, 2)
                
                lbLayers.List(b, 0) = strAtt(0)
                lbLayers.List(b, 1) = strAtt(1)
                lbLayers.List(b, 2) = strAtt(2)
            End If
        Next b
    Next a
End Sub

Private Sub SortListBlocks()
    Dim a, b As Integer
    Dim iCount As Integer
    Dim strAtt(0 To 1) As String
    
    iCount = lbBlocks.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbBlocks.List(b, 0) > lbBlocks.List(b + 1, 0) Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbBlocks.Selected(b) = True
                    lbBlocks.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbBlocks.List(b + 1, 0)
                strAtt(1) = lbBlocks.List(b + 1, 1)
                
                lbBlocks.List(b + 1, 0) = lbBlocks.List(b, 0)
                lbBlocks.List(b + 1, 1) = lbBlocks.List(b, 1)
                
                lbBlocks.List(b, 0) = strAtt(0)
                lbBlocks.List(b, 1) = strAtt(1)
            End If
        Next b
    Next a
End Sub

Private Sub CopyLayers()
    If lbLayers.ListCount < 1 Then Exit Sub
    
    Dim objLayer As AcadLayer
    
    On Error Resume Next
    
    For i = 0 To lbLayers.ListCount - 1
        Set objLayer = ThisDrawing.Layers(lbLayers.List(i, 0))
        If objLayer Is Nothing Then
            Err = 0
            GoTo Next_I
        End If
        
        If InStr(lbLayers.List(i, 0), "Integrity Permits") > 0 Then
            objLayer.LayerOn = True
            objLayer.Lock = True
            GoTo Next_I
        End If
                
        Select Case lbLayers.List(i, 2)
            Case "YES"
                objLayer.LayerOn = True
            Case Else
                objLayer.LayerOn = False
        End Select
Next_I:
    Next i
    
    ThisDrawing.Regen acAllViewports
End Sub

Private Sub CurrentLayers()
    If lbLayers.ListCount < 1 Then Exit Sub
    
    Dim objLayer As AcadLayer
    
    On Error Resume Next
    
    For i = 0 To lbLayers.ListCount - 1
        Set objLayer = ThisDrawing.Layers(lbLayers.List(i, 0))
        If objLayer Is Nothing Then
            Err = 0
            MsgBox "Error"
            GoTo Next_I
        End If
                
        Select Case lbLayers.List(i, 1)
            Case "ON"
                objLayer.LayerOn = True
            Case Else
                objLayer.LayerOn = False
        End Select
        
        If InStr(lbLayers.List(i, 0), "Integrity Permits") > 0 Then objLayer.Lock = False
Next_I:
    Next i
    
    ThisDrawing.Regen acAllViewports
End Sub
