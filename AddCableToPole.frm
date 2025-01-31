VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCableToPole 
   Caption         =   "Add Cable to Pole / Buried Plant"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "AddCableToPole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCableToPole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iListIndex As Integer

Private Sub cbAdd_Click()
    If lbCounts.ListCount < 1 Then
        MsgBox "Counts empty." & vbCr & "Try add a Cable Size."
        Exit Sub
    End If
    
    If tbCblName.Value = "" Then
        MsgBox "Need to add a Count Name"
        Exit Sub
    End If
    
    If tbFFiber.Value = "" Then
        MsgBox "Missing From Fiber"
        Exit Sub
    End If
    
    If tbTFiber.Value = "" Then
        MsgBox "Missing To Fiber"
        Exit Sub
    End If
    
    If tbFCount.Value = "" Then
        MsgBox "Missing Start Count"
        Exit Sub
    End If
    
    Dim vLine, vCable, vFibers, vCounts As Variant
    Dim strLine, strOld As String
    Dim str1, str2, str3 As String
    Dim iFF, iTF As Integer
    Dim iFC, iTC As Integer
    Dim iPFF, iPTF As Integer
    Dim iPFC, iPTC As Integer
    Dim iNewTF, iNewFF As Integer
    Dim iNewTC, iNewFC As Integer
    
    str1 = "": str2 = "": str3 = ""
    
    iPFF = CInt(tbFFiber.Value)
    iPTF = CInt(tbTFiber.Value)
    
    iPFC = CInt(tbFCount.Value)
    iPTC = CInt(tbTCount.Value)
    
    For i = 0 To lbCounts.ListCount - 1
        vLine = Split(lbCounts.List(i), vbTab)
        vFibers = Split(vLine(0), "-")
        vCable = Split(vLine(1), ": ")
        vCounts = Split(vCable(1), "-")
        
        strOld = vCable(2)
        
        iFF = CInt(vFibers(0))
        If UBound(vFibers) > 0 Then
            iTF = CInt(vFibers(1))
        Else
            iTF = iFF
        End If
        
        iFC = CInt(vCounts(0))
        If UBound(vCounts) > 0 Then
            iTC = CInt(vCounts(1))
        Else
            iTC = iFC
        End If
        
        If iPFF > iTF Then GoTo Next_I
        
        Select Case (iPFF - iFF)
            Case Is = 0
                Select Case (iPTF - iTF)
                    Case Is = 0
                        str1 = iPFF & "-" & iPTF & vbTab & tbCblName.Value & ": " & iPFC  '& "-" & iTC
                        If iPFC < iPTC Then str1 = str1 & "-" & iPTC
                        str1 = str1 & ": " & tbPrevious.Value
                        lbCounts.List(i) = str1
                        GoTo Exit_Next
                    Case Is < 0
                        str1 = iPFF & "-" & iPTF & vbTab & tbCblName.Value & ": " & iPFC  '& "-" & iTC
                        If iPFC < iPTC Then str1 = str1 & "-" & iPTC
                        str1 = str1 & ": " & tbPrevious.Value
                        lbCounts.List(i) = str1
                        
                        iNewFF = iPTF + 1
                        iNewFC = iFC + iPTC - iPFC + 1
                        
                        str2 = iNewFF & "-" & iTF & vbTab & vCable(0) & ": " & iNewFC
                        If iNewFC < iTC Then str2 = str2 & "-" & iTC
                        str2 = str2 & ": " & strOld
                        lbCounts.AddItem str2, i + 1
                        GoTo Exit_Next
                    Case Is > 0
                        iNewTC = iPTC + iTF - iFF + 1
                        
                        str1 = iFF & "-" & iTF & vbTab & tbCblName.Value & ": " & iPFC & "-" & iNewTC & ": " & tbPrevious.Value
                        lbCounts.List(i) = str1
                        GoTo Exit_Next
                End Select
            Case Is > 0
                iNewTF = iPFF - 1
                iNewTC = iFC + iNewTF - iFF
                        
                str1 = iFF & "-" & iNewTF & vbTab & vCable(0) & ": " & iFC
                If iNewTC > iFC Then str1 = str1 & "-" & iNewTC
                str1 = str1 & ": " & strOld
                lbCounts.List(i) = str1
                        
                Select Case (iPTF - iTF)
                    Case Is = 0
                        str2 = iPFF & "-" & iPTF & vbTab & tbCblName.Value & ": " & iPFC
                        If iPFC < iPTC Then str2 = str2 & "-" & iPTC
                        str2 = str2 & ": " & tbPrevious.Value
                        lbCounts.AddItem str2, i + 1
                        GoTo Exit_Next
                    Case Is < 0
                        str2 = iPFF & "-" & iPTF & vbTab & tbCblName.Value & ": " & iPFC  '& "-" & iTC
                        If iPFC < iPTC Then str2 = str2 & "-" & iPTC
                        str2 = str2 & ": " & tbPrevious.Value
                        lbCounts.AddItem str2, i + 1
                        
                        iNewFF = iPTF + 1
                        iNewFC = iFC + iPTF - iFF + 1
                        
                        str3 = iNewFF & "-" & iTF & vbTab & vCable(0) & ": " & iNewFC
                        If iNewFC < iTC Then str3 = str3 & "-" & iTC
                        str3 = str3 & ": " & strOld
                        lbCounts.AddItem str3, i + 2
                        GoTo Exit_Next
                    Case Is > 0
                        iNewTC = iPTC + iTF - iFF + 1
                        
                        str2 = iFF & "-" & iTF & vbTab & tbCblName.Value & ": " & iPFC
                        If iPFC < iNewTC Then str2 = str2 & "-" & iNewTC
                        str2 = str2 & ": " & tbPrevious.Value
                        lbCounts.AddItem str2, i + 1
                        GoTo Exit_Next
                End Select
        End Select
Next_I:
    Next i
    
Exit_Next:
    'Dim strLine As String

    'strLine = tbFFiber.Value & "-" & tbTFiber.Value & vbTab
    'strLine = strLine & tbCblName.Value & ": " & tbFCount.Value & "-" & tbTCount.Value
    
    'lbCounts.AddItem strLine
End Sub

Private Sub cbAddToPole_Click()
    If lbCounts.ListCount < 1 Then
        MsgBox "Needs Counts"
        Exit Sub
    End If
    
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vLine As Variant
    Dim vReturnPnt As Variant
    Dim strLine As String
    Dim iAtt As Integer
    
    If cbCblType.Value = "CO" Then
        If cbSuffix.Value = "E" Then
            strLine = tbPosition.Value & "2: "
        Else
            strLine = tbPosition.Value & "1: "
        End If
    Else
        strLine = tbPosition.Value & "1: "
    End If
    
    strLine = strLine & cbCblType.Value & "(" & cbCableSize.Value & ")"
    If Not cbSuffix.Value = "" Then strLine = strLine & cbSuffix.Value
    strLine = strLine & " / "
    
    vLine = Split(lbCounts.List(0), vbTab)
    strLine = strLine & vLine(1)
    
    If lbCounts.ListCount > 1 Then
        For i = 1 To lbCounts.ListCount - 1
            vLine = Split(lbCounts.List(i), vbTab)
            strLine = strLine & " + " & vLine(1)
        Next i
    End If
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt
    If Not Err = 0 Then
        MsgBox "Cable not added to Block."
        Me.show
    End If
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
    Else
        MsgBox "Cable not added to Block."
        Me.show
    End If
    
    Select Case objBlock.Name
        Case "sPole"
            iAtt = 25
        Case Else
            iAtt = 5
    End Select
    
    vAttList = objBlock.GetAttributes
    
    If vAttList(iAtt).TextString = "" Then
        vAttList(iAtt).TextString = strLine
    Else
        vAttList(iAtt).TextString = vAttList(iAtt).TextString & vbCr & strLine
    End If
    
    objBlock.Update
    
    tbPrevious.Value = vAttList(0).TextString
    
    Me.show
End Sub

Private Sub cbCableSize_AfterUpdate()
    If cbCableSize.Value = "" Then Exit Sub
    
    lbCounts.Clear
    
    lbCounts.AddItem "1-" & cbCableSize.Value & vbTab & "XD: 1-" & cbCableSize.Value & ": END"
End Sub

Private Sub cbCancel_Click()
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
        cbSuffix.Enabled = False
        
        tbPosition.Value = tbPosition.Value & "1"
    End If
End Sub

Private Sub cbSendToCounts_Click()
    Me.Hide
End Sub

Private Sub cbCLR_Click()
    lbCounts.Clear
    
    lbCounts.AddItem "1-" & cbCableSize.Value & vbTab & "XD: 1-" & cbCableSize.Value & ": END"
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

Private Sub Label12_Click()
    Dim objEntity As AcadEntity
    Dim objPrevious As AcadBlockReference
    Dim vReturnPnt, vAttList As Variant
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Previous Block: "
    If TypeOf objEntity Is AcadBlockReference Then
        Set objPrevious = objEntity
        vAttList = objPrevious.GetAttributes
        
        tbPrevious.Value = vAttList(0).TextString
    End If
    
    Me.show
End Sub

Private Sub lbCounts_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbCounts.ListCount < 1 Then Exit Sub
    
    Dim str1 As String
    Dim i, i2 As Integer
    
    Select Case KeyCode
        Case vbKeyUp
            If lbCounts.ListIndex = 0 Then Exit Sub
            i = lbcountss.ListIndex
            i2 = i - 1
            str1 = lbCounts.List(i)
            lbCounts.List(i) = lbCounts.List(i2)
            lbCounts.List(i2) = str1
    
            lbCounts.ListIndex = i2
        Case vbKeyDown
            If lbCounts.ListIndex = (lbCounts.ListCount - 1) Then Exit Sub
            i = lbCounts.ListIndex
            i2 = i + 1
            str1 = lbCounts.List(i)
            lbCounts.List(i) = lbCounts.List(i2)
            lbCounts.List(i2) = str1
    
            lbCounts.ListIndex = i2
        Case vbKeyDelete
            lbCounts.RemoveItem (lbCounts.ListIndex)
    End Select
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
    If tbTFiber.Value < tbFFiber.Value Then tbTFiber.Value = tbFFiber.Value
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
    
    'cbCblName.AddItem tbF1Name.Value
    'cbCblName.AddItem tbF2Name.Value
    'cbCblName.AddItem "XD"
    
    'cbCblType.Value = CountsForm.cbCblType.Value
    'cbCableSize.Value = CountsForm.cbCableSize.Value
    'If cbCableSize.Value = "" Then cbCableSize.Value = "144"
    
    'cbSuffix.Value = CountsForm.cbSuffix.Value
End Sub
