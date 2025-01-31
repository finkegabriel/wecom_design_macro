VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewCableForm 
   Caption         =   "New Cable"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   OleObjectBlob   =   "NewCableForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewCableForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iListIndex As Integer



Private Sub cbAdd_Click()
    Dim strLine As String

    strLine = tbFFiber.Value & "-" & tbTFiber.Value
    lbCounts.AddItem strLine
    
    strLine = cbCblName.Value & ": " & tbFCount.Value & "-" & tbTCount.Value
    lbCounts.List(lbCounts.ListCount - 1, 1) = strLine
End Sub

Private Sub cbCableSize_Change()
    If lbCounts.Value = "" Then Exit Sub
    
    
End Sub

Private Sub cbCancel_Click()
    If cbCblType.Value = "" Then
        MsgBox "Need Cable Type."
        Exit Sub
    End If
    If cbCableSize.Value = "" Then
        MsgBox "Need Cable Size."
        Exit Sub
    End If
    If tbPosition.Value = "" Then
        MsgBox "Need Cable Position."
        Exit Sub
    End If
    If cbCblType.Value = "CO" Then
        If cbSuffix.Value = "" Then
            MsgBox "Need Cable Suffix."
            Exit Sub
        End If
    End If
    
    Dim vList, vFibers, vCable, vCounts As Variant
    Dim iFiber, iTotal, iAtt As Integer
    Dim iSFiber, iEFiber, iSCount, iECount As Integer
    Dim strCblName As String
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    
    If Me.Caption = "Tap Cable" Then
        strCblName = tbPosition.Value & ": " & cbCblType.Value & "(" & cbCableSize.Value & ")"
        If Not cbSuffix.Value = "" Then strCblName = strCblName & cbSuffix.Value
        strCblName = strCblName & " / "
        
        If lbCounts.ListCount > 0 Then
            strCblName = strCblName & lbCounts.List(0, 1) & ": "
            If InStr(lbCounts.List(0, 1), "XD") > 0 Then
                strCblName = strCblName & "END"
            Else
                strCblName = strCblName & CountsForm.tbPoleNumber.Value
            End If
            
            If lbCounts.ListCount > 1 Then
                For i = 1 To lbCounts.ListCount - 1
                    strCblName = strCblName & " + " & lbCounts.List(i, 1) & ": "
                    If InStr(lbCounts.List(i, 1), "XD") > 0 Then
                        strCblName = strCblName & "END"
                    Else
                        strCblName = strCblName & CountsForm.tbPoleNumber.Value
                    End If
                Next i
            End If
        End If
        
        'MsgBox strCblName
        
        Me.Hide
        
        On Error Resume Next
        
        Err = 0
        ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block:"
        If Not Err = 0 Then Exit Sub
        
        If Not TypeOf objEntity Is AcadBlockReference Then Exit Sub
        Set objBlock = objEntity
        
        Select Case objBlock.Name
            Case "sPole"
                iAtt = 25
            Case "sPed", "sHH", "sPanel"
                iAtt = 5
            Case Else
                Exit Sub
        End Select
        
        vAttList = objBlock.GetAttributes
        'If vAttList(iAtt).TextString = "" Then
            vAttList(iAtt).TextString = strCblName
        'Else
        'End If
        
        objBlock.Update
            
        Exit Sub
    End If
    
    iFiber = 1
    iTotal = CInt(cbCableSize.Value)
    strCblName = CountsForm.tbF1Name.Value
    
    CountsForm.lbCounts.AddItem ""
    For k = 1 To iTotal
        CountsForm.lbCounts.AddItem
        CountsForm.lbCounts.List(k, 0) = k
        CountsForm.lbCounts.List(k, 1) = "XD"
        CountsForm.lbCounts.List(k, 2) = k
        CountsForm.lbCounts.List(k, 3) = "<>"
        CountsForm.lbCounts.List(k, 4) = "<>"
        CountsForm.lbCounts.List(k, 5) = "<>"
        CountsForm.lbCounts.List(k, 6) = "<>"
        CountsForm.lbCounts.List(k, 7) = "<>"
    Next k
    
    For i = 0 To lbCounts.ListCount - 1
        vList = Split(lbCounts.List(i, 0), vbTab)
        vFibers = Split(vList(0), "-")
        iSFiber = CInt(vFibers(0))
        iEFiber = CInt(vFibers(1))
        
        vCable = Split(lbCounts.List(i, 1), ": ")
        vCounts = Split(vCable(1), "-")
        iSCount = CInt(vCounts(0))
        iECount = CInt(vCounts(1))
        
        If iSFiber < iFiber Then GoTo Next_I
        If iSFiber > iTotal Then GoTo Next_I
        
        For j = iSFiber To iEFiber
            CountsForm.lbCounts.List(j, 1) = vCable(0)
            CountsForm.lbCounts.List(j, 2) = iSCount
            
            'If vCable(0) = strCblName Then
                'DesignCountForm.lbCounts.List(j, 1) = iSCount
            'Else
                'DesignCountForm.lbCounts.List(j, 1) = " "
                'DesignCountForm.lbCounts.List(j, 6) = vCable(0) & ": " & iSCount
            'End If
            iSCount = iSCount + 1
        Next j
Next_I:
    Next i
    
    CountsForm.cbCableSize.Value = cbCableSize.Value

    Me.Hide
End Sub

Private Sub cbCblType_Change()
    If cbCblType.Value = "CO" Then
        cbSuffix.AddItem ""
        cbSuffix.AddItem "E"
        cbSuffix.AddItem "6M-EHS"
        cbSuffix.AddItem "10M"
        cbSuffix.Enabled = True
    Else
        cbSuffix.Clear
        cbSuffix.Value = ""
        cbSuffix.Enabled = False
    End If
End Sub

Private Sub cbCLR_Click()
    lbCounts.Clear
End Sub

Private Sub cbDelete_Click()
    lbCounts.RemoveItem (lbCounts.ListIndex)
End Sub

Private Sub cbDOWN_Click()
    Dim str1 As String
    Dim i, i2 As Integer
    
    If lbCounts.ListIndex = (lbCounts.ListCount - 1) Then Exit Sub
    i = lbCounts.ListIndex
    i2 = i + 1
    str1 = lbCounts.List(i)
    lbCounts.List(i) = lbCounts.List(i2)
    lbCounts.List(i2) = str1
    
    lbCounts.ListIndex = i2
End Sub

Private Sub cbGetCable_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt, vAttList As Variant
    Dim vLine, vItem, vCount As Variant
    Dim iStart, iEnd As Integer
    Dim iFiber, iEndFiber As Integer
    Dim iCable, iCount As Integer
    
    Me.Hide
    On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Callout:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    
    iCable = 1
    Select Case objBlock.Name
        Case "Callout"
            iCount = 2
        Case "CableCounts"
            iCount = 0
        Case Else
            GoTo Exit_Sub
    End Select
    'If Not objBlock.Name = "Callout" Then GoTo Exit_Sub
    vAttList = objBlock.GetAttributes
    
    If Not vAttList(iCable).TextString = "" Then
        vLine = Split(vAttList(iCable).TextString, "(")
        cbCblType.Value = vLine(0)
        If UBound(vLine) > 0 Then
            vItem = Split(vLine(1), ")")
            cbCableSize.Value = vItem(0)
            If UBound(vItem) > 0 Then cbSuffix.Value = vItem(1)
        End If
    End If
    
    iFiber = 1
    
    If Not vAttList(iCount).TextString = "" Then
        vLine = Split(vAttList(iCount).TextString, "\P")
        For i = 0 To UBound(vLine)
            vItem = Split(vLine(i), ": ")
            vCount = Split(vItem(1), "-")
            iStart = CInt(vCount(0))
            If UBound(vCount) > 0 Then
                iEnd = CInt(vCount(1))
            Else
                iEnd = iStart
            End If
            
            iEndFiber = iFiber + iEnd - iStart
            lbCounts.AddItem iFiber & "-" & iEndFiber
            lbCounts.List(lbCounts.ListCount - 1, 1) = vLine(i)
            iFiber = iEndFiber + 1
        Next i
    End If
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbUP_Click()
    Dim str1 As String
    Dim i, i2 As Integer
    
    If lbCounts.ListIndex = 0 Then Exit Sub
    i = lbcountss.ListIndex
    i2 = i - 1
    str1 = lbCounts.List(i)
    lbCounts.List(i) = lbCounts.List(i2)
    lbCounts.List(i2) = str1
    
    lbCounts.ListIndex = i2
End Sub

Private Sub lbCounts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim iIndex As Integer
    Dim vLine, vItem, vTemp As Variant
    
    iIndex = lbCounts.ListIndex
    
    vTemp = Split(lbCounts.List(iIndex, 0), "-")
    
    tbFFiber.Value = vTemp(0)
    tbTFiber.Value = vTemp(1)
    
    vLine = Split(lbCounts.List(iIndex, 1), ": ")
    cbCblName.Value = vLine(0)
    
    vItem = Split(vLine(1), "-")
    tbFCount.Value = vItem(0)
    
    lbCounts.RemoveItem iIndex
End Sub

Private Sub tbFCount_AfterUpdate()
    Dim iFFiber, iTFiber, iFCount, iTCount As Integer
    Dim iDiff As Integer
    Dim dTest As Double
    
    iFFiber = CInt(tbFFiber.Value)
    iTFiber = CInt(tbTFiber.Value)
    iFCount = CInt(tbFCount.Value)
    iDiff = iFCount - iFFiber
    dTest = iDiff / 12
    If Not (dTest - Int(dTest)) = 0 Then MsgBox "Off-Color"
    iDiff = iTFiber - iFFiber
    tbTCount.Value = iFCount + iDiff
End Sub

Private Sub tbFFiber_AfterUpdate()
    If CInt(tbFFiber.Value) > CInt(cbCableSize.Value) Then tbFFiber.Value = 1
End Sub

Private Sub tbTFiber_AfterUpdate()
    If CInt(tbTFiber.Value) > CInt(cbCableSize.Value) Then tbTFiber.Value = cbCableSize.Value
End Sub

Private Sub UserForm_Initialize()
    cbCblType.AddItem ""
    cbCblType.AddItem "CO"
    cbCblType.AddItem "BFO"
    cbCblType.AddItem "UO"
    
    cbCableSize.AddItem ""
    cbCableSize.AddItem "12"
    cbCableSize.AddItem "24"
    cbCableSize.AddItem "36"
    cbCableSize.AddItem "48"
    cbCableSize.AddItem "72"
    cbCableSize.AddItem "96"
    cbCableSize.AddItem "144"
    cbCableSize.AddItem "216"
    cbCableSize.AddItem "288"
    cbCableSize.AddItem "360"
    cbCableSize.AddItem "432"
    
    cbUP.Caption = Chr(225)
    cbDOWN.Caption = Chr(226)
    
    On Error Resume Next
    
    cbCblName.AddItem "XD"
    cbCblName.AddItem CountsForm.tbF1Name.Value
    cbCblName.AddItem CountsForm.tbF2Name.Value
    
    cbCblType.Value = CountsForm.cbCblType.Value
    cbCableSize.Value = CountsForm.cbCableSize.Value
    If cbCableSize.Value = "" Then cbCableSize.Value = "144"
    
    cbSuffix.Value = CountsForm.cbSuffix.Value
    
    lbCounts.ColumnCount = 2
    lbCounts.ColumnWidths = "36;108"
End Sub
