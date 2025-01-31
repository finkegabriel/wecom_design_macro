VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zxxValidateCounts4 
   Caption         =   "Validate Counts"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12240
   OleObjectBlob   =   "zxxValidateCounts4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zxxValidateCounts4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim objSSPoles As AcadSelectionSet
Dim objSS As AcadSelectionSet
Dim iIndex As Integer
Dim vPnt1, vPnt2 As Variant

Private Sub cbGetCallouts_Click()
    Dim objBlock As AcadBlockReference
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vAtt As Variant
    Dim vResult As Variant
    
    On Error Resume Next
    
    Me.Hide
        
    Err = 0
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    lbList.Clear
    objSS.Clear
    
    grpCode(0) = 2
    grpValue(0) = "Callout"
    filterType = grpCode
    filterValue = grpValue
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        
        vAtt = objBlock.GetAttributes
        If InStr(vAtt(1).TextString, "+HO1") > 0 Then GoTo Next_objBlock
        If vAtt(1).TextString = "CALLOUT" Then GoTo Next_objBlock
        
        lbList.AddItem vAtt(0).TextString
        lbList.List(lbList.ListCount - 1, 1) = UCase(vAtt(1).TextString)
        lbList.List(lbList.ListCount - 1, 2) = UCase(vAtt(2).TextString)
        lbList.List(lbList.ListCount - 1, 5) = i
        
        vResult = ValidateCounts(UCase(vAtt(1).TextString), UCase(vAtt(2).TextString))
        
        If vResult(0) = "Y" Then
            lbList.List(lbList.ListCount - 1, 3) = "Y"
        Else
            lbList.List(lbList.ListCount - 1, 3) = ""
        End If
        
        If vResult(1) = "Y" Then
            lbList.List(lbList.ListCount - 1, 4) = "Y"
        Else
            lbList.List(lbList.ListCount - 1, 4) = ""
        End If
Next_objBlock:
    Next i
    
Exit_Sub:
    tbListCount.Value = lbList.ListCount
    Me.show
End Sub

Private Sub cbQuit_Click()
    objSS.Clear
    objSS.Delete
    
    Me.Hide
End Sub

Private Sub cbRemove_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    For i = lbList.ListCount - 1 To 0 Step -1
        If lbList.List(i, 3) = "Y" Then
            If lbList.List(i, 4) = "Y" Then lbList.RemoveItem i
        End If
    Next i
    
    tbListCount.Value = lbList.ListCount
End Sub

Private Sub cbUpdate_Click()
    If iIndex < 0 Then GoTo Exit_Sub
    
    Dim objBlock As AcadBlockReference
    Dim vAtt As Variant
    Dim strText As String
    Dim i As Integer
    
    strText = Replace(UCase(tbCallout.Value), vbLf, "")
    strText = Replace(strText, vbCr, "\P")
    
    lbList.List(iIndex, 1) = UCase(tbCable.Value)
    lbList.List(iIndex, 2) = strText
    
    i = CInt(lbList.List(iIndex, 5))
    Set objBlock = objSS.Item(i)
    vAtt = objBlock.GetAttributes
    
    tbCable.Value = UCase(tbCable.Value)
    vAtt(1).TextString = tbCable.Value
    vAtt(2).TextString = strText
    
    objBlock.Update
        
    vResult = ValidateCounts(tbCable.Value, strText)
    If vResult(0) = "Y" Then lbList.List(iIndex, 3) = "Y"
    If vResult(1) = "Y" Then lbList.List(iIndex, 4) = "Y"
    
    Call UpdatePole(vAtt(0).TextString, vAtt(1).TextString, vAtt(2).TextString)
    
Exit_Sub:
    tbCable.Value = ""
    tbCallout.Value = ""
    
    iIndex = -1
    cbUpdate.Enabled = False
End Sub

Private Sub lbList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strText As String
    
    iIndex = lbList.ListIndex
    cbUpdate.Enabled = True
    
    strText = Replace(UCase(lbList.List(iIndex, 2)), "\P", vbCr)
    tbCallout.Value = strText
    tbCable.Value = UCase(lbList.List(iIndex, 1))
End Sub

Private Sub UserForm_Initialize()
    lbList.ColumnCount = 6
    lbList.ColumnWidths = "120;96;120;36;36;18"
    
    iIndex = -1
    
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    End If
End Sub

Private Function ValidateCounts(strCable As String, strCallout As String)
    Dim strResult, strColor As String
    Dim vResult As Variant
    Dim vLine, vItem, vCount, vTemp As Variant
    
    Dim iCable, iCount, iTCount As Integer
    Dim iLow, iHigh As Integer
    
    vLine = Split(strCable, "(")
    vItem = Split(vLine(1), ")")
    iCable = CInt(vItem(0))
    iCount = 0
    strColor = "Y"
    
    vLine = Split(strCallout, "\P")
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        vCount = Split(vItem(UBound(vItem)), "-")
        
        iTCount = iCount + 1
        While iTCount > 12
            iTCount = iTCount - 12
        Wend
        
        iLow = CInt(vCount(0))
        While iLow > 12
            iLow = iLow - 12
        Wend
        
        If Not iLow = iTCount Then strColor = "N"
        
        If UBound(vCount) = 0 Then
            iCount = iCount + 1
        Else
            iCount = iCount + CInt(vCount(1)) - CInt(vCount(0)) + 1
        End If
    Next i
    
    If iCount = iCable Then
        strResult = "Y"
    Else
        strResult = "N"
    End If
    
    'MsgBox strCable & vbCr & vbCr & strCallout
    
    strResult = strResult & "," & strColor
    vResult = Split(strResult, ",")
    
    ValidateCounts = vResult
End Function

Private Sub UpdatePole(strName As String, strCable As String, strCounts As String)
    Dim strPole, strPosition, strLine As String
    Dim strTemp As String
    Dim vLine, vItem, vTemp As Variant
    
    vLine = Split(strName, ": ")
    strPole = vLine(0)
    strPosition = vLine(1)
    strLine = strPosition & ": " & strCable & " / " & Replace(strCounts, "\P", " + ")
    
    strTemp = strPole & vbCr & strPosition & vbCr & strLine
    MsgBox "Pole not updated" & vbCr & vbCr & strTemp
End Sub
