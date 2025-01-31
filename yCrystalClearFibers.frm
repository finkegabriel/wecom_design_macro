VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} yCrystalClearFibers 
   Caption         =   "Crystal Clear Fibers"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390.001
   OleObjectBlob   =   "yCrystalClearFibers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "yCrystalClearFibers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGetBlocks_Click()
    If lbLots.ListCount < 1 Then Exit Sub
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim objCircle As AcadCircle
    Dim vPnt1, vPnt2 As Variant
    Dim vAttList As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim strSearch As String
    
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    
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
        If vAttList(4).TextString = "" Then
            strSearch = UCase(vAttList(1).TextString & " " & vAttList(2).TextString)
            
            strSearch = Replace(strSearch, "STE", "SUITE")
            strSearch = Replace(strSearch, " PL", ",")
            If InStr(strSearch, " BLVD S") > 0 Then strSearch = Replace(strSearch, " BLVD", ",")
        Else
            strSearch = vAttList(4).TextString
        End If
        
        For i = lbLots.ListCount - 1 To 0 Step -1
            If strSearch = lbLots.List(i, 2) Then
                vAttList(4).TextString = lbLots.List(i, 0) & ": " & lbLots.List(i, 1)
                objBlock.Update
                
                lbLots.RemoveItem i
                
                GoTo Next_objBlock
            End If
        Next i
        
        lbMissing.AddItem vAttList(1).TextString & " " & vAttList(2).TextString
        lbMissing.List(lbMissing.ListCount - 1, 1) = strSearch
        
        vAttList(4).TextString = "LOT: " & strSearch
        objBlock.Update
        
        Set objCircle = ThisDrawing.ModelSpace.AddCircle(objBlock.InsertionPoint, 30#)
        objCircle.Layer = "Integrity Notes"
        objCircle.Update
        
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    tbLListcount.Value = lbLots.ListCount
    tbCListcount.Value = lbMissing.ListCount
    
    Me.show
End Sub

Private Sub cbRemove_Click()
    If lbLots.ListCount < 1 Then Exit Sub
    
    For i = lbLots.ListCount - 1 To 0 Step -1
        If lbLots.List(i, 3) = "Y" Then lbLots.RemoveItem i
    Next i
    
    tbLListcount.Value = lbLots.ListCount
End Sub

Private Sub cbRemoveLot_Click()
    If lbLots.ListCount < 1 Then Exit Sub
    
    For i = lbLots.ListCount - 1 To 0 Step -1
        If lbLots.List(i, 2) = tbLotText.Value Then lbLots.RemoveItem i
    Next i
    
    tbLListcount.Value = lbLots.ListCount
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbLots.ColumnCount = 4
    lbLots.ColumnWidths = "60;60;60;54"
    
    lbMissing.ColumnCount = 2
    lbMissing.ColumnWidths = "150;54"
    
    Dim fName As String
    Dim objExcel As Workbook
    Dim objSheet As Worksheet
    Dim objDoc As Object
    Dim strFileName As String
    Dim iRow, iIndex As Integer
    
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
        lbLots.AddItem objSheet.Cells(iRow, 1)
        iIndex = lbLots.ListCount - 1
        lbLots.List(iIndex, 1) = objSheet.Cells(iRow, 2)
        lbLots.List(iIndex, 2) = UCase(objSheet.Cells(iRow, 3))
        
        iRow = iRow + 1
    Wend
    
    objExcel.Close
    
    tbLListcount.Value = lbLots.ListCount
End Sub
