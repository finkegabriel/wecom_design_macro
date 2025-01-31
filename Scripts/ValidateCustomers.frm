VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateCustomers 
   Caption         =   "Validate Customers"
   ClientHeight    =   8850.001
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   11040
   OleObjectBlob   =   "ValidateCustomers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbFilter_Click()
    If cbCounts.Value = "" Then Exit Sub
    
    For i = lbCustomers.ListCount - 1 To 0 Step -1
        If Not lbCustomers.List(i, 0) = cbCounts.Value Then lbCustomers.RemoveItem i
    Next i
    
    tbCount.Value = lbCustomers.ListCount
    cbFilter.Enabled = False
    cbCounts.Enabled = False
End Sub

Private Sub cbGetCustomers_Click()
    Dim vDwgLL, vDwgUR As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine, vTemp As Variant
    Dim iCount, iIndex, iTest As Integer
    Dim vCoords As Variant
    Dim strCoords As String
    
    On Error Resume Next
    
    lbCustomers.Clear
    cbCounts.Clear
    
    iCount = 0
    iIndex = 0
    iTest = 0
    
    Err = 0
    Me.Hide
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    
    If Not Err = 0 Then
        Me.show
        Exit Sub
    End If
    
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
        Err = 0
    End If
    
    objSS.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    
    iCount = objSS.count
    
    For Each objBlock In objSS
        'Set objBlock = objSS.Item(i)
        vAttList = objBlock.GetAttributes
        
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        If vAttList(1).TextString = "Customer" Then GoTo Next_objBlock
        
        If vAttList(4).TextString = "" Then
            vLine = Split("none - none: none", " - ")
        Else
            vLine = Split(vAttList(4).TextString, " - ")
        End If
        
        vLine(1) = Replace(vLine(1), "(", "")
        vLine(1) = Replace(vLine(1), ")", "")
        vTemp = Split(vLine(1), ": ")
        
        vCoords = objBlock.InsertionPoint
        strCoords = vCoords(0) & "," & vCoords(1)
        
        lbCustomers.AddItem vTemp(0)
        lbCustomers.List(iIndex, 1) = vTemp(1)
        lbCustomers.List(iIndex, 2) = vLine(0)
        lbCustomers.List(iIndex, 3) = vAttList(1).TextString
        lbCustomers.List(iIndex, 4) = vAttList(2).TextString
        lbCustomers.List(iIndex, 5) = vAttList(0).TextString
        lbCustomers.List(iIndex, 6) = vAttList(3).TextString
        lbCustomers.List(iIndex, 7) = ""
        lbCustomers.List(iIndex, 8) = strCoords
        
        If cbCounts.ListCount < 1 Then
            cbCounts.AddItem vTemp(0)
        Else
            For i = 0 To cbCounts.ListCount - 1
                If cbCounts.List(i) = vTemp(0) Then GoTo Found_Count
            Next i
        
            cbCounts.AddItem vTemp(0)
        End If
        
Found_Count:
        
        iIndex = iIndex + 1
        
Next_objBlock:
    Next objBlock
    
    objSS.Clear
    objSS.Delete
    
    tbCount.Value = lbCustomers.ListCount
    cbFilter.Enabled = True
    cbCounts.Enabled = True
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRemoveExtensions_Click()
    For i = lbCustomers.ListCount - 1 To 0 Step -1
        If lbCustomers.List(i, 5) = "EXTENSION" Then lbCustomers.RemoveItem i
    Next i
    
    tbCount.Value = lbCustomers.ListCount
End Sub

Private Sub cbRemoveRef_Click()
    For i = lbCustomers.ListCount - 1 To 0 Step -1
        If InStr(lbCustomers.List(i, 5), "REF") > 0 Then lbCustomers.RemoveItem i
    Next i
    
    tbCount.Value = lbCustomers.ListCount
End Sub

Private Sub cbSave_Click()
    If lbCustomers.ListCount = 0 Then Exit Sub
    
    Dim strFileName As String
    Dim strDWGName As String
    Dim vLine As Variant
    Dim strCopy, strLine As String
    
    
    strCopy = "CABLE NAME,COUNT,POLE NUMBER,HSE #,STREET NAME,TYPE,NOTE"
    
    For i = 0 To lbCustomers.ListCount - 1
        strLine = lbCustomers.List(i, 0) & "," & lbCustomers.List(i, 1) & "," & lbCustomers.List(i, 2)
        strLine = strLine & "," & lbCustomers.List(i, 3) & "," & lbCustomers.List(i, 4) & "," & lbCustomers.List(i, 5)
        strLine = strLine & "," & lbCustomers.List(i, 6)
        
        strCopy = strCopy & vbCr & strLine
    Next i
    
    strFileName = ThisDrawing.Path & "\"
    vLine = Split(ThisDrawing.Name, " ")
    strFileName = strFileName & vLine(0) & "-Customer List "
    
    Select Case cbCounts.Value
        Case ""
            strFileName = strFileName & "ALL"
        Case Else
            strFileName = strFileName & cbCounts.Value
    End Select
    
    strFileName = strFileName & ".csv"

    Open strFileName For Output As #1
        
    Print #1, strCopy
    Close #1
    
    MsgBox "Saved to File"
End Sub

Private Sub cbSort_Click()
    Dim strTemp, strTotal As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 8) As String
    
    iCount = lbCustomers.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If CInt(lbCustomers.List(b, 1)) > CInt(lbCustomers.List(b + 1, 1)) Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbCustomers.Selected(b) = True
                    lbCustomers.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbCustomers.List(b + 1, 0)
                strAtt(1) = lbCustomers.List(b + 1, 1)
                strAtt(2) = lbCustomers.List(b + 1, 2)
                strAtt(3) = lbCustomers.List(b + 1, 3)
                strAtt(4) = lbCustomers.List(b + 1, 4)
                strAtt(5) = lbCustomers.List(b + 1, 5)
                strAtt(6) = lbCustomers.List(b + 1, 6)
                strAtt(7) = lbCustomers.List(b + 1, 7)
                strAtt(8) = lbCustomers.List(b + 1, 8)
                
                lbCustomers.List(b + 1, 0) = lbCustomers.List(b, 0)
                lbCustomers.List(b + 1, 1) = lbCustomers.List(b, 1)
                lbCustomers.List(b + 1, 2) = lbCustomers.List(b, 2)
                lbCustomers.List(b + 1, 3) = lbCustomers.List(b, 3)
                lbCustomers.List(b + 1, 4) = lbCustomers.List(b, 4)
                lbCustomers.List(b + 1, 5) = lbCustomers.List(b, 5)
                lbCustomers.List(b + 1, 6) = lbCustomers.List(b, 6)
                lbCustomers.List(b + 1, 7) = lbCustomers.List(b, 7)
                lbCustomers.List(b + 1, 8) = lbCustomers.List(b, 8)
                
                lbCustomers.List(b, 0) = strAtt(0)
                lbCustomers.List(b, 1) = strAtt(1)
                lbCustomers.List(b, 2) = strAtt(2)
                lbCustomers.List(b, 3) = strAtt(3)
                lbCustomers.List(b, 4) = strAtt(4)
                lbCustomers.List(b, 5) = strAtt(5)
                lbCustomers.List(b, 6) = strAtt(6)
                lbCustomers.List(b, 7) = strAtt(7)
                lbCustomers.List(b, 8) = strAtt(8)
            End If
        Next b
    Next a
End Sub

Private Sub lbCustomers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim objEntity As AcadEntity
    Dim vCoords, vReturnPnt As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    Dim iIndex As Integer
    
    Me.Hide
    
    On Error Resume Next
    
    iIndex = lbCustomers.ListIndex
    
    vCoords = Split(lbCustomers.List(iIndex, 8), ",")
    
    viewCoordsB(0) = vCoords(0) - 300
    viewCoordsB(1) = vCoords(1) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 300
    viewCoordsE(1) = vCoords(1) + 300
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Any Key or Right Click to Exit:"
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbCustomers.Clear
    lbCustomers.ColumnCount = 9
    lbCustomers.ColumnWidths = "72;30;66;36;136;48;114;24;12"
    
    MsgBox "Please let me lnow if you are using this." & vbCr & "Trying to eliminate Forms we may not need."
End Sub
