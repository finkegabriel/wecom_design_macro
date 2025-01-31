VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResizeBlocks 
   Caption         =   "Resize Blocks in Titleblocks"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "ResizeBlocks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResizeBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbExport_Click()
    If tbFile.Value = "" Then Exit Sub
    
    Dim strText, strFileName As String
    
    strFileName = "C:\Integrity\VBA\References\"
    strFileName = strFileName & tbFile.Value & ".rbl"
    
    strText = ""
    
    If lbBlocks.ListCount > 0 Then
        For i = 0 To lbBlocks.ListCount - 1
            If lbBlocks.List(i, 1) = "YES" Then
                If strText = "" Then
                    strText = lbBlocks.List(i, 0) & vbCr
                Else
                    strText = strText & lbBlocks.List(i, 0) & vbCr
                End If
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
    Dim strBlocks As String
    Dim vBlocks As Variant
    Dim fName As String
    
    strFileName = "C:\Integrity\VBA\References\"
    strFileName = strFileName & lbFiles.List(lbFiles.ListIndex) & ".rbl"
    
    strBlocks = ""
    
    fName = Dir(strFileName)
    If fName = "" Then
        Exit Sub
    End If
    
    Open strFileName For Input As #2
    
    While Not EOF(2)
        Input #2, strLine
        
        If strBlocks = "" Then
            strBlocks = strLine
        Else
            strBlocks = strBlocks & vbTab & strLine
        End If
    Wend
    
    Close #2
    
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

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbResize_Click()
    Dim vDwgLL, vDwgUR As Variant
    Dim objSSSS, objSSBlocks As AcadSelectionSet
    Dim objBlock, objBlockSS As AcadBlockReference
    Dim dScale, dPnt(2) As Double
    Dim dLL(2), dUR(2) As Double
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vPnt As Variant
    Dim strText As String
    
    On Error Resume Next
    
    strText = ""
    
    For i = 0 To lbBlocks.ListCount - 1
        If lbBlocks.List(i, 1) = "YES" Then
            If strText = "" Then
                strText = lbBlocks.List(i, 0)
            Else
                strText = strText & "," & lbBlocks.List(i, 0)
            End If
        End If
    Next i
    
    If strText = "" Then Exit Sub
    
    Me.Hide
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get UR Corner: ")
    
    If Not Err = 0 Then GoTo Exit_Sub
    
    Set objSSSS = ThisDrawing.SelectionSets.Add("objSSSS")
    If Not Err = 0 Then
        Set objSSSS = ThisDrawing.SelectionSets.Item("objSSSS")
        Err = 0
    End If
    
    Set objSSBlocks = ThisDrawing.SelectionSets.Add("objSSBlocks")
    If Not Err = 0 Then
        Set objSSBlocks = ThisDrawing.SelectionSets.Item("objSSBlocks")
        Err = 0
    End If
    
    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    objSSSS.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    
    grpValue(0) = strText
    filterType = grpCode
    filterValue = grpValue
    
    objSSBlocks.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    
    For Each objBlock In objSSSS
        dScale = objBlock.XScaleFactor * 1.3333333333
        
        vPnt = objBlock.InsertionPoint
        dLL(0) = vPnt(0)
        dLL(1) = vPnt(1)
        
        dUR(0) = dLL(0) + (1650 * dScale)
        dUR(1) = dLL(1) + (1050 * dScale)
        
        For Each objBlockSS In objSSBlocks
            dPnt(0) = objBlockSS.InsertionPoint(0)
            dPnt(1) = objBlockSS.InsertionPoint(1)
            
            If dPnt(0) > dLL(0) And dPnt(0) < dUR(0) Then
                If dPnt(1) > dLL(1) And dPnt(1) < dUR(1) Then
                    objBlockSS.XScaleFactor = dScale
                    objBlockSS.YScaleFactor = dScale
                    objBlockSS.ZScaleFactor = dScale
                    
                    objBlockSS.Update
                End If
            End If
        Next objBlockSS
    Next objBlock
    
Exit_Sub:
    objSSSS.Clear
    objSSSS.Delete
    
    objSSBlocks.Clear
    objSSBlocks.Delete
    
    Me.show
End Sub

Private Sub cbSelectLayer_Click()
    If lbBlocks.ListCount < 1 Then Exit Sub
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim strName As String
    
    On Error Resume Next
    
    Me.Hide
    
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Blocks:"
    If Not Err = 0 Then GoTo Exit_Sub
    
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

Private Sub lbBlocks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Select Case lbBlocks.List(lbBlocks.ListIndex, 1)
        Case "YES"
            lbBlocks.List(lbBlocks.ListIndex, 1) = ""
        Case Else
            lbBlocks.List(lbBlocks.ListIndex, 1) = "YES"
    End Select
End Sub

Private Sub UserForm_Initialize()
    Dim strTFolder, strFile As String
    Dim vTemp As Variant
    strTFolder = "C:\Integrity\VBA\References\*.*"
    
    strFile = Dir$(strTFolder)
    
    Do While strFile <> ""
        If InStr(strFile, ".rbl") Then
            lbFiles.AddItem Replace(strFile, ".rbl", "")
        End If
        strFile = Dir$
    Loop
    
    lbBlocks.ColumnCount = 2
    lbBlocks.ColumnWidths = "126;48"
    
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

Private Function FindDWGNumber(vPnt As Variant)
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vAttList As Variant
    Dim dLL(2), dUR(2) As Double
    Dim dPnt(1) As Double
    Dim dScale As Double
    Dim strDWG As String
    
    On Error Resume Next
    
    strDWG = "??"
    dPnt(0) = vPnt(0)
    dPnt(1) = vPnt(1)
    
    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        dScale = objBlock.XScaleFactor
        
        dLL(0) = objBlock.InsertionPoint(0)
        dLL(1) = objBlock.InsertionPoint(1)
        
        dUR(0) = dLL(0) + (1650 * dScale)
        dUR(1) = dLL(1) + (1050 * dScale)
        
        If dPnt(0) > dLL(0) And dPnt(0) < dUR(0) Then
            If dPnt(1) > dLL(1) And dPnt(1) < dUR(1) Then
                strDWG = vAttList(0).TextString
                'MsgBox "Found:" & strDWG
                GoTo Exit_Sub
            End If
        End If
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    FindDWGNumber = strDWG
End Function
