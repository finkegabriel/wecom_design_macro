VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} yCrystalClearUpdateSS 
   Caption         =   "UserForm3"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12720
   OleObjectBlob   =   "yCrystalClearUpdateSS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "yCrystalClearUpdateSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGetCustomers_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim objCircle As AcadCircle
    Dim vPnt1, vPnt2 As Variant
    Dim vAttList, vLine As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    
    lbPanel.Clear
    cbPanels.Clear
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    Me.Hide
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    dPnt1(0) = vPnt1(0)
    dPnt1(1) = vPnt1(1)
    dPnt1(2) = vPnt1(2)
    
    dPnt2(0) = vPnt2(0)
    dPnt2(1) = vPnt2(1)
    dPnt2(2) = vPnt2(2)
    
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(4).TextString = "" Then GoTo Next_objBlock
        
        vLine = Split(vAttList(4).TextString, ": ")
        lbFibers.AddItem vLine(0)
        lbFibers.List(lbFibers.ListCount - 1, 1) = vLine(1)
        lbFibers.List(lbFibers.ListCount - 1, 2) = vAttList(1).TextString
        lbFibers.List(lbFibers.ListCount - 1, 3) = vAttList(2).TextString
        lbFibers.List(lbFibers.ListCount - 1, 4) = vAttList(3).TextString
        
Next_objBlock:
    Next objBlock
    
    cbPanels.AddItem lbFibers.List(0, 0)
    If lbFibers.ListCount > 1 Then
        For i = 1 To lbFibers.ListCount - 1
            For j = 0 To cbPanels.ListCount - 1
                If lbFibers.List(i, 0) = cbPanels.List(j) Then GoTo Found_Panel
            Next j
            
            cbPanels.AddItem lbFibers.List(i, 0)
Found_Panel:
        Next i
    End If
    
    Call SortPanels
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    'Call SortFList
    
    tbListcount.Value = lbFibers.ListCount
    
    Me.show
End Sub

Private Sub cbPanels_Change()
    If cbPanels.Value = "" Then Exit Sub
    If lbFibers.ListCount < 1 Then Exit Sub
    
    lbPanel.Clear
    
    For i = 0 To lbFibers.ListCount - 1
        If lbFibers.List(i, 0) = cbPanels.Value Then
            lbPanel.AddItem lbFibers.List(i, 0)
            lbPanel.List(lbPanel.ListCount - 1, 1) = lbFibers.List(i, 1)
            lbPanel.List(lbPanel.ListCount - 1, 2) = lbFibers.List(i, 2)
            lbPanel.List(lbPanel.ListCount - 1, 3) = lbFibers.List(i, 3)
            lbPanel.List(lbPanel.ListCount - 1, 4) = lbFibers.List(i, 4)
        End If
    Next i
    
    Call SortList
    
    tbPListcount.Value = lbPanel.ListCount
End Sub

Private Sub cbUpdateXLSX_Click()
    If lbFibers.ListCount < 1 Then Exit Sub
    
    Dim fName As String
    Dim objExcel As Workbook
    Dim objSheet As Worksheet
    Dim objDoc As Object
    Dim strFileName As String
    Dim iRow, iIndex As Integer
    
    Dim vLine As Variant
    Dim strPanel, strPort As String
    
    strFileName = ThisDrawing.Path & "\Panel-Port-Lot.xlsx" 'Panel-Port-Lot.xlsx
    'MsgBox strFileName
    
    fName = Dir(strFileName)
    If fName = "" Then
        MsgBox "File not found."
        Exit Sub
    Else
        Set objExcel = Workbooks.Open(strFileName)
    End If
    
    Set objSheet = objExcel.Sheets("All")
    iRow = 1
    
    While Not objSheet.Cells(iRow, 1) = ""
        strPanel = objSheet.Cells(iRow, 1)
        strPort = objSheet.Cells(iRow, 2)
        
        For i = 0 To lbFibers.ListCount - 1
            If lbFibers.List(i, 0) = strPanel And lbFibers.List(i, 1) = strPort Then
                objSheet.Cells(iRow, 4) = lbFibers.List(i, 2)
                objSheet.Cells(iRow, 5) = lbFibers.List(i, 3)
                
                lbFibers.RemoveItem i
                GoTo Next_iRow
            End If
        Next i
        
        
Next_iRow:
        If lbFibers.ListCount < 1 Then GoTo Exit_Sub
        
        iRow = iRow + 1
    Wend
    
    For i = lbFibers.ListCount - 1 To 0 Step -1
        If lbFibers.List(i, 0) = "LOT" Then
            objSheet.Cells(iRow, 1) = "UNK"
            objSheet.Cells(iRow, 2) = "UNK"
            objSheet.Cells(iRow, 3) = lbFibers.List(i, 1)
        Else
            objSheet.Cells(iRow, 1) = lbFibers.List(i, 0)
            objSheet.Cells(iRow, 2) = lbFibers.List(i, 1)
            objSheet.Cells(iRow, 3) = "UNK"
        End If
        
        objSheet.Cells(iRow, 4) = lbFibers.List(i, 2)
        objSheet.Cells(iRow, 5) = lbFibers.List(i, 3)
        objSheet.Cells(iRow, 6) = lbFibers.List(i, 4)
        
        lbFibers.RemoveItem i
        
        iRow = iRow + 1
    Next i
    
Exit_Sub:
    objExcel.Save
    objExcel.Close
End Sub

Private Sub UserForm_Initialize()
    lbFibers.ColumnCount = 5
    lbFibers.ColumnWidths = "48;48;48;114;6"
    
    lbPanel.ColumnCount = 5
    lbPanel.ColumnWidths = "48;48;48;114;6"
End Sub

Private Sub SortFList()
    If lbFibers.ListCount < 3 Then Exit Sub
    
    Dim strTemp, strTotal As String
    Dim strCurrent, strNext As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 5) As String
    
    iCount = lbFibers.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            strCurrent = lbFibers.List(b, 0)
            Select Case Len(lbFibers.List(b, 1))
                Case Is = 1
                    strCurrent = strCurrent & "00" & lbFibers.List(b, 1)
                Case Is = 2
                    strCurrent = strCurrent & "0" & lbFibers.List(b, 1)
                Case Else
                    strCurrent = strCurrent & lbFibers.List(b, 1)
            End Select
            
            strNew = lbFibers.List(b + 1, 0)
            Select Case Len(lbFibers.List(b + 1, 1))
                Case Is = 1
                    strNew = strNew & "00" & lbFibers.List(b + 1, 1)
                Case Is = 2
                    strNew = strNew & "0" & lbFibers.List(b + 1, 1)
                Case Else
                    strNew = strNew & lbFibers.List(b + 1, 1)
            End Select
            
            
            If strCurrent > strNew Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbFibers.Selected(b) = True
                    lbFibers.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbFibers.List(b + 1, 0)
                strAtt(1) = lbFibers.List(b + 1, 1)
                strAtt(2) = lbFibers.List(b + 1, 2)
                strAtt(3) = lbFibers.List(b + 1, 3)
                
                lbFibers.List(b + 1, 0) = lbFibers.List(b, 0)
                lbFibers.List(b + 1, 1) = lbFibers.List(b, 1)
                lbFibers.List(b + 1, 2) = lbFibers.List(b, 2)
                lbFibers.List(b + 1, 3) = lbFibers.List(b, 3)
                
                lbFibers.List(b, 0) = strAtt(0)
                lbFibers.List(b, 1) = strAtt(1)
                lbFibers.List(b, 2) = strAtt(2)
                lbFibers.List(b, 3) = strAtt(3)
            End If
        Next b
    Next a
End Sub

Private Sub SortList()
    If lbPanel.ListCount < 3 Then Exit Sub
    
    Dim strTemp, strTotal As String
    Dim strCurrent, strNext As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 5) As String
    
    iCount = lbPanel.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            strCurrent = lbPanel.List(b, 0)
            Select Case Len(lbPanel.List(b, 1))
                Case Is = 1
                    strCurrent = strCurrent & "000" & lbPanel.List(b, 1)
                Case Is = 2
                    strCurrent = strCurrent & "00" & lbPanel.List(b, 1)
                Case Is = 3
                    strCurrent = strCurrent & "0" & lbPanel.List(b, 1)
                Case Else
                    strCurrent = strCurrent & lbPanel.List(b, 1)
            End Select
            
            strNew = lbPanel.List(b + 1, 0)
            Select Case Len(lbPanel.List(b + 1, 1))
                Case Is = 1
                    strNew = strNew & "000" & lbPanel.List(b + 1, 1)
                Case Is = 2
                    strNew = strNew & "00" & lbPanel.List(b + 1, 1)
                Case Is = 3
                    strNew = strNew & "0" & lbPanel.List(b + 1, 1)
                Case Else
                    strNew = strNew & lbPanel.List(b + 1, 1)
            End Select
            
            
            If strCurrent > strNew Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbPanel.Selected(b) = True
                    lbPanel.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbPanel.List(b + 1, 0)
                strAtt(1) = lbPanel.List(b + 1, 1)
                strAtt(2) = lbPanel.List(b + 1, 2)
                strAtt(3) = lbPanel.List(b + 1, 3)
                
                lbPanel.List(b + 1, 0) = lbPanel.List(b, 0)
                lbPanel.List(b + 1, 1) = lbPanel.List(b, 1)
                lbPanel.List(b + 1, 2) = lbPanel.List(b, 2)
                lbPanel.List(b + 1, 3) = lbPanel.List(b, 3)
                
                lbPanel.List(b, 0) = strAtt(0)
                lbPanel.List(b, 1) = strAtt(1)
                lbPanel.List(b, 2) = strAtt(2)
                lbPanel.List(b, 3) = strAtt(3)
            End If
        Next b
    Next a
End Sub

Private Sub SortPanels()
    If cbPanels.ListCount < 3 Then Exit Sub
    
    Dim iCount As Integer
    Dim strAtt As String
    
    iCount = cbPanels.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            strCurrent = cbPanels.List(b)
            strNew = cbPanels.List(b + 1)
            
            If strCurrent > strNew Then
                strAtt = cbPanels.List(b + 1)
                cbPanels.List(b + 1) = cbPanels.List(b)
                cbPanels.List(b) = strAtt
            End If
        Next b
    Next a
End Sub
