VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} xxValidateMCL 
   Caption         =   "Validate MCL Files"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "xxValidateMCL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "xxValidateMCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim objSS As AcadSelectionSet

Private Sub cbCheckData_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    For i = lbList.ListCount - 1 To 0 Step -1
            If lbList.List(i, 6) = "Y" Then lbList.RemoveItem i
    Next i
End Sub

Private Sub cbGetCustomers_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vLine, vItem, vCount As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim strText As String
    Dim iCount As Integer
    
    grpCode(0) = 2
    grpValue(0) = "Customer,SG"
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
    
    'objSS.Clear
    'objSS.Select acSelectionSetAll, , , filterType, filterValue
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        Select Case objBlock.Name
            Case "SG"
                If vAttList(2).TextString = "" Then GoTo Next_objBlock
                
                vLine = Split(vAttList(2).TextString, " - ")
                If UBound(vLine) < 1 Then GoTo Next_objBlock
                
                vCount = Split(vLine(1), ": ")
                If InStr(vCount(0), "(") > 0 Then
                    vCount(0) = Replace(vCount(0), "(", "")
                    vCount(1) = Replace(vCount(1), ")", "")
                End If
                
                If Not vCount(0) = tbCableName.Value Then GoTo Next_objBlock
                
                iCount = CInt(vCount(1))
                
                
                
                
                
                strText = vLine(1) & vbTab & "SG - " & vAttList(1).TextString
            Case Else
                If vAttList(4).TextString = "" Then GoTo Next_objBlock
                
                vLine = Split(vAttList(4).TextString, " - ")
                If UBound(vLine) < 1 Then GoTo Next_objBlock
                
                vCount = Split(vLine(1), ": ")
                If InStr(vCount(0), "(") > 0 Then
                    vCount(0) = Replace(vCount(0), "(", "")
                    vCount(1) = Replace(vCount(1), ")", "")
                End If
                
                'MsgBox vCount(0) & "*" & vbCr & tbCableName.Value & "*"
                If Not vCount(0) = tbCableName.Value Then GoTo Next_objBlock
                
                iCount = CInt(vCount(1)) - 1
                lbList.List(iCount, 4) = vAttList(1).TextString & "  " & vAttList(2).TextString
                lbList.List(iCount, 5) = Left(vAttList(0).TextString, 1)
        End Select
        
Next_objBlock:
    Next objBlock
    
Clear_objSS:
    objSS.Clear
    objSS.Delete
    
    Call ValidateCustomers
    
    Me.show
End Sub

Private Sub cbList_Change()
    Dim strFileName As String
    Dim vLine As Variant
    Dim strLine As String
    Dim fName As String
    Dim iIndex As Integer
    
    strFileName = LCase(ThisDrawing.Path) & "\" & cbList.Value & ".mcl"
    
    fName = Dir(strFileName)
    
    Open strFileName For Input As #2
    
    Line Input #2, strLine
    tbCableName.Value = Replace(strLine, " ", "")
    
    lbList.Clear
    
    While Not EOF(2)
        Line Input #2, strLine
        vLine = Split(strLine, vbTab)
        
        lbList.AddItem vLine(0)
        iIndex = lbList.ListCount - 1
        
        lbList.List(iIndex, 1) = vLine(1)
        lbList.List(iIndex, 2) = vLine(2) & "  " & vLine(3)
        If Left(vLine(4), 1) = "<" Then
            lbList.List(iIndex, 3) = " "
        Else
            lbList.List(iIndex, 3) = Left(vLine(4), 1)
        End If
        lbList.List(iIndex, 4) = "n/a"
        lbList.List(iIndex, 5) = " "
        lbList.List(iIndex, 6) = " "
    Wend
    
    Close #2
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    lbList.ColumnCount = 7
    lbList.ColumnWidths = "36;120;144,24,144,24,30"
    
    On Error Resume Next
    
    'Err = 0
    'Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    'If Not Err = 0 Then
        'Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        'objSS.Clear
    'End If
    
    Dim strFile, strFolder, strTemp As String
    Dim vName, vLine, vItem As Variant
    Dim strLine As String
    Dim fName As String
    Dim iIndex, iCount As Integer
    
    strFolder = ThisDrawing.Path & "\*.*"
    
    strFile = Dir$(strFolder)
    
    Do While strFile <> ""
        If InStr(strFile, ".mcl") Then
            cbList.AddItem Replace(strFile, ".mcl", "")
        End If
        strFile = Dir$
    Loop
    
    If cbList.ListCount < 1 Then Exit Sub
    iCount = cbList.ListCount - 1
    
    'Sort list
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If cbList.List(b) > cbList.List(b + 1) Then
                strTemp = cbList.List(b + 1)
                
                cbList.List(b + 1) = cbList.List(b)
                
                cbList.List(b) = strTemp
            End If
        Next b
    Next a
    
End Sub

Private Sub UserForm_Terminate()
    'objSS.Clear
    'objSS.Delete
End Sub

Private Sub ValidateCustomers()
    If lbList.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbList.ListCount - 1
        If lbList.List(i, 2) = lbList.List(i, 4) Then
            If lbList.List(i, 3) = lbList.List(i, 5) Then lbList.List(i, 6) = "Y"
        End If
        
        If lbList.List(i, 1) = "<>" Then
            If lbList.List(i, 4) = "n/a" Then lbList.List(i, 6) = "Y"
        End If
    Next i
End Sub
