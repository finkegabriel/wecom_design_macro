VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15030
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iListIndex As Integer

Private Sub cbAddToList_Click()
    If tbEndFibers.Value = "" Then Exit Sub
    If tbSource.Value = "" Then Exit Sub
    
    lbList.AddItem tbEndFibers.Value
    lbList.List(lbList.ListCount - 1, 1) = tbSource.Value
End Sub

Private Sub cbGet_Click()
    Dim objEntity As AcadEntity
    Dim vReturnPnt As Variant
    Dim vAttList As Variant
    Dim strLine As String
    Dim iStart, iEnd, iSize As Integer
    
    Me.Hide
    
    On Error Resume Next
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    vAttList = objBlock.GetAttributes
    
    Dim vLine, vItem, vCounts, vTemp As Variant
    
    Select Case objBlock.Name
        Case "sPole"
            strLine = vAttList(25).TextString
        Case "sPed", "sHH", "sFP", "sMH", "sPanel"
            strLine = vAttList(5).TextString
        Case "Callout"
            vLine = Split(vAttList(1).TextString, ")")
            vItem = Split(vLine(0), "(")
            iSize = CInt(vItem(1))
            'If Not vItem(1) = cbCableSize.Value Then
                'If cbCableSize.Value = "" Then
                    'cbCableSize.Value = vItem(1)
                'Else
                    'GoTo Exit_Sub
                'End If
            'End If
            
            strLine = Replace(vAttList(2).TextString, "\P", " + ")
            'MsgBox strLine
            GoTo Fixed_Line
        Case Else
            GoTo Exit_Sub
    End Select
    
    vLine = Split(strLine, " / ")
    vItem = Split(vLine(0), ")")
    vTemp = Split(vItem(0), "(")
    iSize = CInt(vTemp(1))
    'If Not vTemp(1) = cbCableSize.Value Then
        'If cbCableSize.Value = "" Then
            'cbCableSize.Value = vTemp(1)
        'Else
            'GoTo Exit_Sub
        'End If
    'End If
    
    strLine = vLine(1)
    
Fixed_Line:
    
    'If cbCableSize.Value = "" Then GoTo Exit_Sub
    If strLine = "" Then GoTo Exit_Sub
    
    Dim strName, strSource As String
    Dim iFiber As Integer
    
    iFiber = 1
    lbCounts.Clear
    lbCounts.AddItem ""
    
    vLine = Split(strLine, " + ")
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        strName = vItem(0)
        If UBound(vItem) < 2 Then
            strSource = ""
        Else
            strSource = vItem(2)
        End If
        
        vCounts = Split(vItem(1), "-")
        iStart = CInt(vCounts(0))
        If UBound(vCounts) < 1 Then
            iEnd = iStart
        Else
            iEnd = CInt(vCounts(1))
        End If
        
        'strLine = iFiber & "-"
        'iFiber = iFiber + iEnd - iStart
        'strLine = strLine & iFiber
        
        'lbCounts.AddItem strLine
        'lbCounts.List(lbCounts.ListCount - 1, 1) = strName
        'lbCounts.List(lbCounts.ListCount - 1, 2) = vItem(1)
        'lbCounts.List(lbCounts.ListCount - 1, 3) = strSource
        
        'iFiber = iFiber + 1
        
        For j = iStart To iEnd
            lbCounts.AddItem iFiber
            lbCounts.List(lbCounts.ListCount - 1, 1) = strName
            lbCounts.List(lbCounts.ListCount - 1, 2) = j
            lbCounts.List(lbCounts.ListCount - 1, 3) = strSource
            
            iFiber = iFiber + 1
        Next j
    Next i
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbGetBlock_Click()
    Dim objEntity As AcadEntity
    Dim vReturnPnt As Variant
    Dim vAttList As Variant
    Dim strLine As String
    Dim iStart, iEnd, iSize As Integer
    
    Me.Hide
    
    On Error Resume Next
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    vAttList = objBlock.GetAttributes
    
    Dim vLine, vItem, vCounts, vTemp As Variant
    
    Select Case objBlock.Name
        Case "sPole"
            strLine = vAttList(25).TextString
        Case "sPed", "sHH", "sFP", "sMH", "sPanel"
            strLine = vAttList(5).TextString
        Case "Callout"
            vLine = Split(vAttList(1).TextString, ")")
            vItem = Split(vLine(0), "(")
            iSize = CInt(vItem(1))
            'If Not vItem(1) = cbCableSize.Value Then
                'If cbCableSize.Value = "" Then
                    'cbCableSize.Value = vItem(1)
                'Else
                    'GoTo Exit_Sub
                'End If
            'End If
            
            strLine = Replace(vAttList(2).TextString, "\P", " + ")
            'MsgBox strLine
            GoTo Fixed_Line
        Case Else
            GoTo Exit_Sub
    End Select
    
    vLine = Split(strLine, " / ")
    vItem = Split(vLine(0), ")")
    vTemp = Split(vItem(0), "(")
    iSize = CInt(vTemp(1))
    'If Not vTemp(1) = cbCableSize.Value Then
        'If cbCableSize.Value = "" Then
            'cbCableSize.Value = vTemp(1)
        'Else
            'GoTo Exit_Sub
        'End If
    'End If
    
    strLine = vLine(1)
    
Fixed_Line:
    
    'If cbCableSize.Value = "" Then GoTo Exit_Sub
    If strLine = "" Then GoTo Exit_Sub
    
    strLine = Replace(strLine, " + ", vbCr)
    
    tbBefore.Value = strLine
    lbBefore.Clear
    
    vLine = Split(strLine, vbCr)
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        
        lbBefore.AddItem vItem(0)
        lbBefore.List(lbBefore.ListCount - 1, 1) = vItem(1)
        If UBound(vItem) > 1 Then lbBefore.List(lbBefore.ListCount - 1, 2) = vItem(2)
    Next i
    
    'Dim strName, strSource As String
    'Dim iFiber As Integer
    
    'iFiber = 1
    'lbCounts.Clear
    'lbCounts.AddItem ""
    
    'vLine = Split(strLine, " + ")
    'For i = 0 To UBound(vLine)
        'vItem = Split(vLine(i), ": ")
        'strName = vItem(0)
        'If UBound(vItem) < 2 Then
            'strSource = ""
        'Else
            'strSource = vItem(2)
        'End If
        
        'vCounts = Split(vItem(1), "-")
        'iStart = CInt(vCounts(0))
        'If UBound(vCounts) < 1 Then
            'iEnd = iStart
        'Else
            'iEnd = CInt(vCounts(1))
        'End If
        
        'strLine = iFiber & "-"
        'iFiber = iFiber + iEnd - iStart
        'strLine = strLine & iFiber
        
        'lbCounts.AddItem strLine
        'lbCounts.List(lbCounts.ListCount - 1, 1) = strName
        'lbCounts.List(lbCounts.ListCount - 1, 2) = vItem(1)
        'lbCounts.List(lbCounts.ListCount - 1, 3) = strSource
        
        'iFiber = iFiber + 1
        
        'For j = iStart To iEnd
            'lbCounts.AddItem iFiber
            'lbCounts.List(lbCounts.ListCount - 1, 1) = strName
            'lbCounts.List(lbCounts.ListCount - 1, 2) = j
            'lbCounts.List(lbCounts.ListCount - 1, 3) = strSource
            
            'iFiber = iFiber + 1
        'Next j
    'Next i
    
Exit_Sub:
    Me.show
End Sub

Private Sub lbList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strAtt(1) As String
    Dim iIndex As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            iListIndex = lbList.ListIndex
    
            tbEndFibers.Value = lbList.List(iListIndex, 0)
            tbSource.Value = lbList.List(iListIndex, 1)
    
            cbAddToList.Caption = "Update"
        Case vbKeyUp
            iIndex = lbList.ListIndex
            If iIndex < 1 Then Exit Sub
            
            strAtt(0) = lbList.List(iIndex, 0)
            strAtt(1) = lbList.List(iIndex, 1)
            
            lbList.List(iIndex, 0) = lbList.List(iIndex - 1, 0)
            lbList.List(iIndex, 1) = lbList.List(iIndex - 1, 1)
            
            lbList.List(iIndex - 1, 0) = strAtt(0)
            lbList.List(iIndex - 1, 1) = strAtt(1)
        Case vbKeyDown
            iIndex = lbList.ListIndex
            If iIndex > lbList.ListCount - 2 Then Exit Sub
            
            strAtt(0) = lbList.List(iIndex, 0)
            strAtt(1) = lbList.List(iIndex, 1)
            
            lbList.List(iIndex, 0) = lbList.List(iIndex + 1, 0)
            lbList.List(iIndex, 1) = lbList.List(iIndex + 1, 1)
            
            lbList.List(iIndex + 1, 0) = strAtt(0)
            lbList.List(iIndex + 1, 1) = strAtt(1)
        Case vbKeyDelete
            lbList.RemoveItem lbList.ListIndex
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbCounts.ColumnCount = 4
    lbCounts.ColumnWidths = "48;96;48;90"

    lbList.ColumnCount = 2
    lbList.ColumnWidths = "72;72"

    lbBefore.ColumnCount = 3
    lbBefore.ColumnWidths = "60;24;60"
    
    cbEndFiber.AddItem "12"
    cbEndFiber.AddItem "24"
    cbEndFiber.AddItem "48"
    cbEndFiber.AddItem "72"
    cbEndFiber.AddItem "96"
    cbEndFiber.AddItem "144"
    cbEndFiber.AddItem "216"
    cbEndFiber.AddItem "288"
    cbEndFiber.AddItem "360"
    cbEndFiber.AddItem "432"
    cbEndFiber.AddItem "576"
    cbEndFiber.AddItem "876"
    
    cbCableSize.AddItem "12"
    cbCableSize.AddItem "24"
    cbCableSize.AddItem "48"
    cbCableSize.AddItem "72"
    cbCableSize.AddItem "96"
    cbCableSize.AddItem "144"
    cbCableSize.AddItem "216"
    cbCableSize.AddItem "288"
    cbCableSize.AddItem "360"
    cbCableSize.AddItem "432"
    cbCableSize.AddItem "576"
    cbCableSize.AddItem "876"
End Sub

Private Sub UpdateCallout()
    Dim vLine, vItem, vCounts As Variant
    Dim iStart, iEnd As Integer
    
    For i = 0 To lbList.ListCount - 1
        vItem = Split(lbList.List(i, 0), "-")
        iStart = CInt(vItem(0))
        iEnd = CInt(vItem(1))
        
        For j = iStart To iEnd
            
        Next j
    Next i
End Sub
