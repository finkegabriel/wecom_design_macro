VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlaceCableCCounts 
   Caption         =   "Cable Counts"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   OleObjectBlob   =   "PlaceCableCCounts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlaceCableCCounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAddUpdate_Click()
    Dim iIndex As Integer
    
    If cbAddUpdate.Caption = "Update" Then
        iIndex = lbCounts.ListIndex
        lbCounts.List(iIndex, 0) = tbName.Value
        lbCounts.List(iIndex, 1) = tbCounts.Value
        lbCounts.List(iIndex, 2) = tbSource.Value
        lbCounts.List(iIndex, 3) = cbStatus.Value
        
        GoTo Exit_Sub
    End If
    
    lbCounts.AddItem tbName.Value
    iIndex = lbCounts.ListCount - 1
    lbCounts.List(iIndex, 1) = tbCounts.Value
    lbCounts.List(iIndex, 2) = tbSource.Value
    lbCounts.List(iIndex, 3) = cbStatus.Value
    
Exit_Sub:
    cbAddUpdate.Caption = "Add New Line"
End Sub

Private Sub cbCancel_Click()
    Me.Hide
End Sub

Private Sub cbGetLine_Click()
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim vCounts As Variant
    Dim iFrom, iTo As Integer
    
    For i = 0 To lbCounts.ListCount - 1
        If lbCounts.Selected(i) = True Then
            vCounts = Split(lbCounts.List(i, 1), "-")
            If UBound(vCounts) = 0 Then Exit Sub
            
            tbFromFirst.Value = vCounts(0)
            tbToLast.Value = vCounts(1)
        End If
    Next i
    
    cbSplit.Enabled = True
    tbTo.SetFocus
End Sub

Private Sub cbQuit_Click()
    Dim strLine, strTemp As String
    
    strLine = ""
    
    For i = 0 To lbCounts.ListCount - 1
        strTemp = lbCounts.List(i, 0) & ":" & vbTab & lbCounts.List(i, 1) & ":" & vbTab & lbCounts.List(i, 2)
        'If lbCounts.List(i, 3) = "Old" Then strTemp = "(" & strTemp & ")"
        
        If strLine = "" Then
            strLine = strTemp
        Else
            strLine = strLine & vbCr & strTemp
        End If
    Next i
    
    PlaceCountCallouts.tbCableCounts.Value = strLine
    
    Me.Hide
End Sub

Private Sub cbSplit_Click()
    If tbTo.Value = "" Then
        tbFrom.Value = ""
        Exit Sub
    End If
    
    Dim iCurrent, iFrom, iTo As Integer
    
    iFrom = CInt(tbFromFirst.Value)
    iTo = CInt(tbToLast.Value)
    iCurrent = CInt(tbTo.Value)
    
    If iCurrent < iFrom + 1 Then Exit Sub
    If iCurrent > iTo - 1 Then Exit Sub
    
    Dim iIndex, iNext As Integer
    Dim strName, strLine, strSource As String
    
    iNext = iCurrent + 1
    iIndex = lbCounts.ListIndex
    lbCounts.List(iIndex, 1) = iFrom & "-" & iCurrent
    
    iIndex = iIndex + 1
    lbCounts.AddItem lbCounts.List(iIndex - 1, 0), iIndex
    lbCounts.List(iIndex, 1) = iNext & "-" & iTo
    lbCounts.List(iIndex, 2) = lbCounts.List(iIndex - 1, 2)
    lbCounts.List(iIndex, 3) = lbCounts.List(iIndex - 1, 3)
    
    tbFromFirst.Value = ""
    tbTo.Value = ""
    tbFrom.Value = ""
    tbToLast.Value = ""
End Sub

Private Sub lbCounts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim iIndex As Integer
    
    iIndex = lbCounts.ListIndex
    
    tbName.Value = lbCounts.List(iIndex, 0)
    tbCounts.Value = lbCounts.List(iIndex, 1)
    tbSource.Value = lbCounts.List(iIndex, 2)
    cbStatus.Value = lbCounts.List(iIndex, 3)
    
    cbAddUpdate.Caption = "Update"
End Sub

Private Sub lbCounts_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim str1, str2, str3, str4 As String
    Dim i, i2 As Integer
    
    Select Case KeyCode
        Case vbKeyDown
            If lbCounts.ListIndex = (lbCounts.ListCount - 1) Then Exit Sub
            i = lbCounts.ListIndex
            i2 = i + 1
            
            str1 = lbCounts.List(i, 0)
            str2 = lbCounts.List(i, 1)
            str3 = lbCounts.List(i, 2)
            str4 = lbCounts.List(i, 3)
            
            
            lbCounts.List(i, 0) = lbCounts.List(i2, 0)
            lbCounts.List(i, 1) = lbCounts.List(i2, 1)
            lbCounts.List(i, 2) = lbCounts.List(i2, 2)
            lbCounts.List(i, 3) = lbCounts.List(i2, 3)
            
            lbCounts.List(i2, 0) = str1
            lbCounts.List(i2, 1) = str2
            lbCounts.List(i2, 2) = str3
            lbCounts.List(i2, 3) = str4
            'lbUnits.ListIndex = i2
        Case vbKeyUp
            If lbCounts.ListIndex = 0 Then Exit Sub
            i = lbCounts.ListIndex
            i2 = i - 1
            
            str1 = lbCounts.List(i, 0)
            str2 = lbCounts.List(i, 1)
            str3 = lbCounts.List(i, 2)
            str4 = lbCounts.List(i, 3)
            
            lbCounts.List(i, 0) = lbCounts.List(i2, 0)
            lbCounts.List(i, 1) = lbCounts.List(i2, 1)
            lbCounts.List(i, 2) = lbCounts.List(i2, 2)
            lbCounts.List(i, 3) = lbCounts.List(i2, 3)
            
            lbCounts.List(i2, 0) = str1
            lbCounts.List(i2, 1) = str2
            lbCounts.List(i2, 2) = str3
            lbCounts.List(i2, 3) = str4
            'lbUnits.ListIndex = i2
        Case vbKeyDelete
            lbCounts.RemoveItem (lbCounts.ListIndex)
    End Select
End Sub

Private Sub tbTo_Change()
    If tbTo.Value = "" Then
        tbFrom.Value = ""
        Exit Sub
    End If
    
    Dim iCurrent, iFrom, iTo As Integer
    
    iFrom = CInt(tbFromFirst.Value)
    iTo = CInt(tbToLast.Value)
    iCurrent = CInt(tbTo.Value)
    
    If iCurrent < iFrom + 1 Then Exit Sub
    If iCurrent > iTo - 1 Then Exit Sub
    
    tbFrom.Value = iCurrent + 1
End Sub

Private Sub UserForm_Initialize()
    cbStatus.AddItem "Current"
    cbStatus.AddItem "Old"
    cbStatus.AddItem "New"
    
    lbCounts.ColumnCount = 4
    lbCounts.ColumnWidths = "96;48;96;38"
    
    Dim vCount, vItem As Variant
    Dim strCounts As String
    
    strCounts = PlaceCountCallouts.tbCableCounts.Value
    strCounts = Replace(strCounts, vbLf, "")
    strCounts = Replace(strCounts, vbTab, " ")
    vCount = Split(strCounts, vbCr)
    
    For i = 0 To UBound(vCount)
        If Not vCount(i) = "" Then
            vItem = Split(vCount(i), ": ")
            lbCounts.AddItem vItem(0)
            lbCounts.List(lbCounts.ListCount - 1, 1) = vItem(1)
            lbCounts.List(lbCounts.ListCount - 1, 2) = vItem(2)
            If InStr(vCount(i), "(") > 0 Then
                lbCounts.List(lbCounts.ListCount - 1, 3) = "Old"
            Else
                lbCounts.List(lbCounts.ListCount - 1, 3) = "Current"
            End If
        End If
    Next i
End Sub
