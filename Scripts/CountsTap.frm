VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CountsTap 
   Caption         =   "Tap Cable"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   OleObjectBlob   =   "CountsTap.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CountsTap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iChanged As Integer

Private Sub cbAddToBottom_Click()
    If lbTap.ListCount < 1 Then Exit Sub
    If lbMain.ListCount < 1 Then Exit Sub
    
    Dim iMainIndex, iTapIndex As Integer
    Dim iCounter, iIndex As Integer
    
    iCounter = 0
    iIndex = 0
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.Selected(i) = True Then
            iCounter = iCounter + 1
            iIndex = i
        Else
            If iCounter > 0 Then GoTo Exit_Next
        End If
    Next i
Exit_Next:
    
    iTapIndex = iIndex - iCounter + 1
    While iTapIndex > lbTap.ListCount - 1
        iTapIndex = iTapIndex - 12
    Wend
    While iTapIndex < lbTap.ListCount - 13
        iTapIndex = iTapIndex + 12
    Wend
    iCounter = 0
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.Selected(i) = True Then
            lbTap.List(iTapIndex, 1) = lbMain.List(i, 1)
            lbTap.List(iTapIndex, 2) = lbMain.List(i, 2)
            iTapIndex = iTapIndex + 1
            lbMain.List(i, 3) = "Y " & iTapIndex
            iCounter = iCounter + 1
        Else
            If iCounter > 0 Then GoTo Exit_Sub
        End If
    Next i
    
Exit_Sub:
    
    Call GetTapCallout
End Sub

Private Sub cbAddToTop_Click()
    If lbTap.ListCount < 1 Then Exit Sub
    If lbMain.ListCount < 1 Then Exit Sub
    
    Dim iMainIndex, iTapIndex As Integer
    Dim iCounter, iIndex As Integer
    
    iCounter = 0
    'iIndex = 0
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.Selected(i) = True Then
            If iCounter = 0 Then
                iTapIndex = CInt(lbMain.List(i, 2)) - 1
                While iTapIndex > 11
                    iTapIndex = iTapIndex - 12
                Wend
            End If
            
            lbTap.List(iTapIndex, 1) = lbMain.List(i, 1)
            lbTap.List(iTapIndex, 2) = lbMain.List(i, 2)
            iTapIndex = iTapIndex + 1
            lbMain.List(i, 3) = "Y " & iTapIndex
            iCounter = iCounter + 1
        Else
            If iCounter > 0 Then GoTo Exit_Next
        End If
    Next i
    
Exit_Next:
    
    Call GetTapCallout
End Sub

Private Sub cbCableSize_Change()
    If cbCableSize.Value = "" Then Exit Sub
    
    Dim iCount As Integer
    
    lbTap.Clear
    iCount = CInt(cbCableSize.Value)
    
    For i = 0 To iCount - 1
        lbTap.AddItem i + 1
        lbTap.List(i, 1) = "XD"
        lbTap.List(i, 2) = i + 1
    Next i
    
    tbResult.Value = "XD: 1-" & cbCableSize.Value & ": END"
End Sub

Private Sub cbCblType_Change()
    Select Case cbCblType.Value
        Case "BFO", "UO"
            cbSuffix.Value = ""
        Case "CO"
            cbSuffix.Value = "6M-EHS"
    End Select
End Sub

Private Sub cbDone_Click()
    cbChanged.Value = True
    Me.Hide
End Sub

Private Sub cbIn_Click()
    If lbTap.ListCount < 1 Then Exit Sub
    If lbMain.ListCount < 1 Then Exit Sub
    
    Dim iMainIndex, iTapIndex As Integer
    Dim iTapTemp, iMainTemp As Integer
    Dim iCounter As Integer
    Dim result As Integer
    
    For i = 0 To lbTap.ListCount - 1
        If lbTap.Selected(i) = True Then
            iTapIndex = i
            GoTo Found_TapSelected
        End If
    Next i
    
    Exit Sub
Found_TapSelected:
    iTapTemp = iTapIndex
    While iTapTemp > 11
        iTapTemp = iTapTemp - 12
    Wend
    iCounter = 0
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.Selected(i) = True Then
            If iCounter = 0 Then
                iMainTemp = i
                While iMainTemp > 11
                    iMainTemp = iMainTemp - 12
                Wend
                
                If Not iTapTemp = iMainTemp Then
                    result = MsgBox("Continue off-color?", vbYesNo, "Off - Color")
                    If result = vbNo Then Exit Sub
                End If
            End If
            
            iCounter = iCounter + 1
            lbTap.List(iTapIndex, 1) = lbMain.List(i, 1) & "(IN)"
            lbTap.List(iTapIndex, 2) = lbMain.List(i, 2)
            iTapIndex = iTapIndex + 1
            lbMain.List(i, 3) = "Y " & iTapIndex
        Else
            If iCounter > 0 Then GoTo Exit_Next
        End If
    Next i
    
Exit_Next:
    
    Call GetTapCallout
End Sub

Private Sub cbOut_Click()
    If lbTap.ListCount < 1 Then Exit Sub
    If lbMain.ListCount < 1 Then Exit Sub
    
    Dim iMainIndex, iTapIndex As Integer
    Dim iTapTemp, iMainTemp As Integer
    Dim iCounter As Integer
    Dim result As Integer
    
    For i = 0 To lbTap.ListCount - 1
        If lbTap.Selected(i) = True Then
            iTapIndex = i
            GoTo Found_TapSelected
        End If
    Next i
    
    Exit Sub
Found_TapSelected:
    iTapTemp = iTapIndex
    While iTapTemp > 11
        iTapTemp = iTapTemp - 12
    Wend
    iCounter = 0
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.Selected(i) = True Then
            If iCounter = 0 Then
                iMainTemp = i
                While iMainTemp > 11
                    iMainTemp = iMainTemp - 12
                Wend
                
                If Not iTapTemp = iMainTemp Then
                    result = MsgBox("Continue off-color?", vbYesNo, "Off - Color")
                    If result = vbNo Then Exit Sub
                End If
            End If
            
            iCounter = iCounter + 1
            lbTap.List(iTapIndex, 1) = lbMain.List(i, 1) & "(OUT)"
            lbTap.List(iTapIndex, 2) = lbMain.List(i, 2)
            iTapIndex = iTapIndex + 1
            lbMain.List(i, 3) = "Y " & iTapIndex
        Else
            If iCounter > 0 Then GoTo Exit_Next
        End If
    Next i
    
Exit_Next:
    
    Call GetTapCallout
End Sub

Private Sub cbQuit_Click()
    cbChanged.Value = False
    Me.Hide
End Sub

Private Sub cbSpliceSelected_Click()
    If lbTap.ListCount < 1 Then Exit Sub
    If lbMain.ListCount < 1 Then Exit Sub
    
    Dim iMainIndex, iTapIndex As Integer
    Dim iTapTemp, iMainTemp As Integer
    Dim iCounter As Integer
    Dim result As Integer
    
    For i = 0 To lbTap.ListCount - 1
        If lbTap.Selected(i) = True Then
            iTapIndex = i
            GoTo Found_TapSelected
        End If
    Next i
    
    Exit Sub
Found_TapSelected:
    iTapTemp = iTapIndex
    While iTapTemp > 11
        iTapTemp = iTapTemp - 12
    Wend
    iCounter = 0
    
    For i = 0 To lbMain.ListCount - 1
        If lbMain.Selected(i) = True Then
            If iCounter = 0 Then
                iMainTemp = i
                While iMainTemp > 11
                    iMainTemp = iMainTemp - 12
                Wend
                
                If Not iTapTemp = iMainTemp Then
                    result = MsgBox("Continue off-color?", vbYesNo, "Off - Color")
                    If result = vbNo Then Exit Sub
                End If
            End If
            
            iCounter = iCounter + 1
            lbTap.List(iTapIndex, 1) = lbMain.List(i, 1)
            lbTap.List(iTapIndex, 2) = lbMain.List(i, 2)
            iTapIndex = iTapIndex + 1
            lbMain.List(i, 3) = "Y " & iTapIndex
        Else
            If iCounter > 0 Then GoTo Exit_Next
        End If
    Next i
    
Exit_Next:
    
    Call GetTapCallout
End Sub

Private Sub lbTap_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbTap.ListCount < 1 Then Exit Sub
    
    Dim strCable, strCount As String
    Dim iIndex As Integer
    
    Select Case KeyCode
        Case vbKeyDelete
            iIndex = lbTap.ListIndex
            
            strCable = lbTap.List(iIndex, 1)
            strCount = lbTap.List(iIndex, 2)
            
            lbTap.List(iIndex, 1) = "XD"
            lbTap.List(iIndex, 2) = lbTap.List(iIndex, 0)
            
            For i = 0 To lbMain.ListCount - 1
                If lbMain.List(i, 1) = strCable Then
                    If lbMain.List(i, 2) = strCount Then
                        lbMain.List(i, 3) = ""
                        lbMain.Selected(i) = False
                        GoTo Exit_Next
                    End If
                End If
            Next i
Exit_Next:
            If Not iIndex > lbTap.ListCount - 1 Then
                lbTap.Selected(iIndex) = False
                lbTap.ListIndex = iIndex + 1
                lbTap.Selected(iIndex + 1) = True
            End If
            
            Call GetTapCallout
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbMain.ColumnCount = 4
    lbMain.ColumnWidths = "24;72;36;30"
    
    lbTap.ColumnCount = 3
    lbTap.ColumnWidths = "24;72;30"
    
    cbCblType.AddItem "CO"
    cbCblType.AddItem "BFO"
    cbCblType.AddItem "UO"
    cbCblType.Value = "CO"
    
    cbSuffix.AddItem ""
    cbSuffix.AddItem "E"
    cbSuffix.AddItem "6M-EHS"
    cbSuffix.AddItem "6M"
    cbSuffix.AddItem "10M"
    cbSuffix.Value = "6M-EHS"
    
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
    
    cbChanged.Value = False
    
'    For i = 0 To 5
'        lbMain.AddItem i + 1
'        lbMain.List(i, 1) = "XD"
'        lbMain.List(i, 2) = i + 1
'        lbMain.List(i, 3) = ""
'    Next i
'
'    For i = 6 To 20
'        lbMain.AddItem i + 1
'        lbMain.List(i, 1) = "F100"
'        lbMain.List(i, 2) = i + 25
'        lbMain.List(i, 3) = ""
'    Next i
'
'    For i = 21 To 23
'        lbMain.AddItem i + 1
'        lbMain.List(i, 1) = "XD"
'        lbMain.List(i, 2) = i + 1
'        lbMain.List(i, 3) = ""
'    Next i
'
'    For i = 24 To 65
'        lbMain.AddItem i + 1
'        lbMain.List(i, 1) = "F100-32"
'        lbMain.List(i, 2) = i - 23
'        lbMain.List(i, 3) = ""
'    Next i
'
'    For i = 66 To 71
'        lbMain.AddItem i + 1
'        lbMain.List(i, 1) = "XD"
'        lbMain.List(i, 2) = i + 1
'        lbMain.List(i, 3) = ""
'    Next i
    
End Sub

Private Sub GetTapCallout()
    Dim strLine, strItem As String
    Dim strCurrent, strPrevious As String
    Dim iStart, iEnd As Integer
    Dim iHO1 As Integer
    
    tbResult.Value = ""
    iHO1 = 0
    
    strPrevious = lbTap.List(0, 1)
    strCurrent = strPrevious
    iStart = CInt(lbTap.List(0, 2))
    iEnd = iStart
    
    If Not strCurrent = "XD" Then iHO1 = iHO1 + 1
    
    For i = 1 To lbTap.ListCount - 1
        strCurrent = lbTap.List(i, 1)
        
        If strCurrent = strPrevious Then
            iEnd = CInt(lbTap.List(i, 2))
        Else
            strItem = strPrevious & ": " & iStart
            If iEnd > iStart Then strItem = strItem & "-" & iEnd
            If strPrevious = "XD" Then
                strItem = strItem & ": END"
            Else
                If InStr(strPrevious, " OUT") = 0 Then
                    strItem = strItem & ": " & tbStructure.Value
                Else
                    strItem = strItem & ": <<???>>"
                End If
            End If
            
            If strLine = "" Then
                strLine = strItem
            Else
                strLine = strLine & vbCr & strItem
            End If
             
            strPrevious = strCurrent
            iStart = CInt(lbTap.List(i, 2))
            iEnd = iStart
        End If
    
        If Not strCurrent = "XD" Then iHO1 = iHO1 + 1
    Next i
             
    strItem = strPrevious & ": " & iStart
    If iEnd > iStart Then strItem = strItem & "-" & iEnd
    If strPrevious = "XD" Then
        strItem = strItem & ": END"
    Else
        If InStr(strPrevious, " OUT") = 0 Then
            strItem = strItem & ": " & tbStructure.Value
        Else
            strItem = strItem & ": <<???>>"
        End If
    End If
             
    If strLine = "" Then
        strLine = strItem
    Else
        strLine = strLine & vbCr & strItem
    End If
    
    tbResult.Value = strLine
    tbHO1.Value = iHO1
End Sub
