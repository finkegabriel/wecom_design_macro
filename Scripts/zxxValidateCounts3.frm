VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zxxValidateCounts3 
   Caption         =   "Validate Counts"
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   OleObjectBlob   =   "zxxValidateCounts3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zxxValidateCounts3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Dim vLine, vItem, vTemp As Variant
    Dim strCounts As String
    Dim iAtt As Integer
    
    On Error Resume Next
    
    Me.Hide
        
    Err = 0
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    lbList.Clear
    objSS.Clear
    
    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH,sPanel"
    filterType = grpCode
    filterValue = grpValue
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    'MsgBox objSS.count
    
    For i = 0 To objSS.count - 1
        Set objBlock = objSS.Item(i)
        
        vAtt = objBlock.GetAttributes
        If vAtt(0).TextString = "" Then GoTo Next_objBlock
        If vAtt(0).TextString = "POLE" Then GoTo Next_objBlock
        If vAtt(0).TextString = "PED" Then GoTo Next_objBlock
        If vAtt(0).TextString = "HH" Then GoTo Next_objBlock
        If vAtt(0).TextString = "PANEL" Then GoTo Next_objBlock
        
        If objBlock.Name = "sPole" Then
            vLine = Split(UCase(vAtt(25).TextString), vbCr)
        Else
            vLine = Split(UCase(vAtt(5).TextString), vbCr)
        End If
        
        'MsgBox vLine(0) & vbCr & UBound(vLine) & vLine(UBound(vLine))
        
        For j = 0 To UBound(vLine)
            vItem = Split(vLine(j), " / ")
            vTemp = Split(vItem(0), ": ")
            strCounts = Replace(vItem(1), " + ", "\P")
            
            lbList.AddItem vAtt(0).TextString
            lbList.List(lbList.ListCount - 1, 1) = objBlock.Name
            lbList.List(lbList.ListCount - 1, 2) = vTemp(0)
            lbList.List(lbList.ListCount - 1, 3) = vTemp(1)
            lbList.List(lbList.ListCount - 1, 4) = strCounts
            lbList.List(lbList.ListCount - 1, 7) = i
        
            vResult = ValidateCounts(CStr(vTemp(1)), strCounts)
        
            If vResult(0) = "Y" Then
                lbList.List(lbList.ListCount - 1, 5) = "Y"
            Else
                lbList.List(lbList.ListCount - 1, 5) = ""
            End If
        
            If vResult(1) = "Y" Then
                lbList.List(lbList.ListCount - 1, 6) = "Y"
            Else
                lbList.List(lbList.ListCount - 1, 6) = ""
            End If
        Next j
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
        If lbList.List(i, 5) = "Y" Then
            If lbList.List(i, 6) = "Y" Then lbList.RemoveItem i
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
    strText = Replace(strText, vbTab, " ")
    
    lbList.List(iIndex, 3) = UCase(tbCable.Value)
    lbList.List(iIndex, 4) = strText
    
    'i = CInt(lbList.List(iIndex, 7))
    'Set objBlock = objSS.Item(i)
    'vAtt = objBlock.GetAttributes
    
    'tbCable.Value = UCase(tbCable.Value)
    'vAtt(1).TextString = tbCable.Value
    'vAtt(2).TextString = strText
    
    'objBlock.Update
        
    vResult = ValidateCounts(tbCable.Value, strText)
    If vResult(0) = "Y" Then lbList.List(iIndex, 5) = "Y"
    If vResult(1) = "Y" Then lbList.List(iIndex, 6) = "Y"
    
    'Call UpdatePole(vAtt(0).TextString, vAtt(1).TextString, vAtt(2).TextString)
    
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
    
    strText = Replace(UCase(lbList.List(iIndex, 4)), "\P", vbCr)
    strText = Replace(strText, " ", vbTab)
    tbCallout.Value = strText
    tbCable.Value = UCase(lbList.List(iIndex, 3))
End Sub

Private Sub UserForm_Initialize()
    lbList.ColumnCount = 8
    lbList.ColumnWidths = "120;36;24;96;120;36;36;18"
    
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
        If InStr(vLine(i), ")") > 0 Then
            vTemp = Split(vLine(i), ")")
            vItem = Split(vTemp(1), ": ")
        Else
            vItem = Split(vLine(i), ": ")
        End If
        vCount = Split(vItem(1), "-")
        
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
