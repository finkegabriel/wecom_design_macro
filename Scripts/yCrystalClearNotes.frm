VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} yCrystalClearNotes 
   Caption         =   "UserForm3"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "yCrystalClearNotes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "yCrystalClearNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGet_Click()
    Dim objEntity As AcadEntity
    Dim vReturnPnt As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine As Variant
    
    On Error Resume Next
    
    Me.Hide
    
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    If Not objBlock.Name = "Customer" Then GoTo Exit_Sub
    
    vAttList = objBlock.GetAttributes
        
    vLine = Split(vAttList(4).TextString, ": ")
        
    If UBound(vLine) = 0 Then
        lbFibers.AddItem "???"
        lbFibers.List(lbFibers.ListCount - 1, 1) = vLine(0)
        lbFibers.List(lbFibers.ListCount - 1, 2) = vAttList(1).TextString & " " & vAttList(2).TextString
    Else
        lbFibers.AddItem vLine(0)
        lbFibers.List(lbFibers.ListCount - 1, 1) = vLine(1)
        lbFibers.List(lbFibers.ListCount - 1, 2) = vAttList(1).TextString & " " & vAttList(2).TextString
    End If
    
    GoTo Get_Another
    
Exit_Sub:
    
    Call SortList
    cbPlace.SetFocus
    
    Me.show
End Sub

Private Sub cbPlace_Click()
    If lbFibers.ListCount < 1 Then Exit Sub
    
    Dim objMText As AcadMText
    Dim vBasePnt, vLL As Variant
    Dim strLine, strTemp, strColors As String
    Dim dInsert(2) As Double
    Dim dScale As Double
    Dim iLength As Integer
    
    On Error Resume Next
    Me.Hide
    
    vBasePnt = ThisDrawing.Utility.GetPoint(, "Select Note Placement: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    dInsert(0) = vBasePnt(0)
    dInsert(1) = vBasePnt(1)
    dInsert(2) = 0#
    
    dScale = CDbl(tbTextHeight.Value)
    
    Select Case lbFibers.List(0, 0)
        Case "LOT", "???"
            strColors = "? / ?"
            iLength = 5
        Case Else
            strColors = GetColors(CInt(lbFibers.List(0, 1)))
            iLength = GetLength(CStr(strColors))
    End Select
    
    strLine = lbFibers.List(0, 0) & ": " & lbFibers.List(0, 1)
    If Len(strLine) < 11 Then
        strLine = strLine & vbTab & vbTab & strColors
    Else
        strLine = strLine & vbTab & strColors
    End If
            
    If iLength < 7 Then
        strLine = strLine & vbTab & vbTab & lbFibers.List(0, 2)
    Else
        strLine = strLine & vbTab & lbFibers.List(0, 2)
    End If
    
    If lbFibers.ListCount > 1 Then
        For i = 1 To lbFibers.ListCount - 1
            Select Case lbFibers.List(i, 0)
                Case "LOT", "???"
                    strColors = "? / ?"
                    iLength = 5
                Case Else
                    strColors = GetColors(CInt(lbFibers.List(i, 1)))
                    iLength = GetLength(CStr(strColors))
            End Select
            
            strTemp = lbFibers.List(i, 0) & ": " & lbFibers.List(i, 1)
            If Len(strTemp) < 11 Then
                strTemp = strTemp & vbTab & vbTab & strColors
            Else
                strTemp = strTemp & vbTab & strColors
            End If
            
            If iLength < 7 Then
                strTemp = strTemp & vbTab & vbTab & lbFibers.List(i, 2)
            Else
                strTemp = strTemp & vbTab & lbFibers.List(i, 2)
            End If
            
            strLine = strLine & "\P" & strTemp
        Next i
    End If

    Set objMText = ThisDrawing.ModelSpace.AddMText(dInsert, 0, strLine)
    objMText.Layer = cbNoteLayer.Value
    objMText.Height = dScale
    objMText.InsertionPoint = dInsert
    Select Case cbJust.Value
        Case "TL"
            objMText.AttachmentPoint = acAttachmentPointTopLeft
        Case "TC"
            objMText.AttachmentPoint = acAttachmentPointTopCenter
        Case "TR"
            objMText.AttachmentPoint = acAttachmentPointTopRight
        Case "ML"
            objMText.AttachmentPoint = acAttachmentPointMiddleLeft
        Case "MC"
            objMText.AttachmentPoint = acAttachmentPointMiddleCenter
        Case "MR"
            objMText.AttachmentPoint = acAttachmentPointMiddleRight
        Case "BL"
            objMText.AttachmentPoint = acAttachmentPointBottomLeft
        Case "BC"
            objMText.AttachmentPoint = acAttachmentPointBottomCenter
        Case "BR"
            objMText.AttachmentPoint = acAttachmentPointBottomRight
    End Select
    If cbMask.Value = True Then objMText.BackgroundFill = True
    
    objMText.Update
    
Exit_Sub:
    lbFibers.Clear
    cbGet.SetFocus
    
    Me.show
End Sub

Private Sub lbFibers_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbFibers.ListCount < 1 Then Exit Sub
    
    Dim iIndex As Integer
    Dim str(2) As String
    
    Select Case KeyCode
        Case vbKeyDelete
            lbFibers.RemoveItem lbFibers.ListIndex
        Case vbKeyDown
            iIndex = lbFibers.ListIndex
            If iIndex = lbFibers.ListCount - 1 Then Exit Sub
            
            str(0) = lbFibers.List(iIndex, 0)
            str(1) = lbFibers.List(iIndex, 1)
            str(2) = lbFibers.List(iIndex, 2)
            
            lbFibers.List(iIndex, 0) = lbFibers.List(iIndex + 1, 0)
            lbFibers.List(iIndex, 1) = lbFibers.List(iIndex + 1, 1)
            lbFibers.List(iIndex, 2) = lbFibers.List(iIndex + 1, 2)
            
            lbFibers.List(iIndex + 1, 0) = str(0)
            lbFibers.List(iIndex + 1, 1) = str(1)
            lbFibers.List(iIndex + 1, 2) = str(2)
        Case vbKeyUp
            iIndex = lbFibers.ListIndex
            If iIndex < 1 Then Exit Sub
            
            str(0) = lbFibers.List(iIndex, 0)
            str(1) = lbFibers.List(iIndex, 1)
            str(2) = lbFibers.List(iIndex, 2)
            
            lbFibers.List(iIndex, 0) = lbFibers.List(iIndex - 1, 0)
            lbFibers.List(iIndex, 1) = lbFibers.List(iIndex - 1, 1)
            lbFibers.List(iIndex, 2) = lbFibers.List(iIndex - 1, 2)
            
            lbFibers.List(iIndex - 1, 0) = str(0)
            lbFibers.List(iIndex - 1, 1) = str(1)
            lbFibers.List(iIndex - 1, 2) = str(2)
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbFibers.ColumnCount = 3
    lbFibers.ColumnWidths = "48;48;120"
    
    cbJust.AddItem "TL"
    cbJust.AddItem "TC"
    cbJust.AddItem "TR"
    cbJust.AddItem "ML"
    cbJust.AddItem "MC"
    cbJust.AddItem "MR"
    cbJust.AddItem "BL"
    cbJust.AddItem "BC"
    cbJust.AddItem "BR"
    cbJust.Value = "TL"
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        cbNoteLayer.AddItem objLayer.Name
    Next objLayer
    cbNoteLayer.Value = "Integrity Notes"
End Sub

Private Function GetColors(iPort As Integer)
    If iPort > 144 Then iPort = iPort - 144
    
    Dim strLine As String
    Dim iRibbon, iStrand As Integer
    
    iFiber = iPort
    iRibbon = 0
    While iFiber > 12
        iRibbon = iRibbon + 1
        iFiber = iFiber - 12
    Wend
    iFiber = iFiber - 1
    
    'iRibbon = CInt(iPort / 12 - 0.5)
    'If iRibbon < 0 Then iRibbon = 0
    'iFiber = iPort - (iRibbon * 12) - 1
    'If iFiber < 1 Then
        'iRibbon = iRibbon - 1
        'iFiber = 0
    'End If
    
    Select Case iRibbon
        Case Is < 1
            strLine = "BL / "
        Case Is = 1
            strLine = "O / "
        Case Is = 2
            strLine = "G / "
        Case Is = 3
            strLine = "BR / "
        Case Is = 4
            strLine = "S / "
        Case Is = 5
            strLine = "W / "
        Case Is = 6
            strLine = "RD / "
        Case Is = 7
            strLine = "BK / "
        Case Is = 8
            strLine = "Y / "
        Case Is = 9
            strLine = "V / "
        Case Is = 10
            strLine = "RS / "
        Case Is = 11
            strLine = "A / "
    End Select
    
    Select Case iFiber
        Case Is < 1
            strLine = strLine & "BL"
        Case Is = 1
            strLine = strLine & "O"
        Case Is = 2
            strLine = strLine & "G"
        Case Is = 3
            strLine = strLine & "BR"
        Case Is = 4
            strLine = strLine & "S"
        Case Is = 5
            strLine = strLine & "W"
        Case Is = 6
            strLine = strLine & "RD"
        Case Is = 7
            strLine = strLine & "BK"
        Case Is = 8
            strLine = strLine & "Y"
        Case Is = 9
            strLine = strLine & "V"
        Case Is = 10
            strLine = strLine & "RS"
        Case Is = 11
            strLine = strLine & "A"
    End Select
    
    GetColors = strLine
End Function

Private Function GetLength(strLine As String)
    Dim vLine As Variant
    Dim dLength As Double
    Dim iLength As Integer
    Dim strTemp As String
    
    vLine = Split(strLine, " / ")
    
    Select Case vLine(0)
        Case "S", "Y", "V", "A"
            dLength = 4
        Case "G", "W", "O"
            dLength = 4.5
        Case Else
            dLength = 5
    End Select
    
    Select Case vLine(1)
        Case "S", "Y", "V", "A"
            dLength = dLength + 1
        Case "G", "W", "O"
            dLength = dLength + 1.5
        Case Else
            dLength = dLength + 2
    End Select
    
    strTemp = dLength
    If InStr(strTemp, ".") > 0 Then dLength = dLength + 0.5
    
    iLength = CInt(dLength)
    GetLength = iLength
End Function

Private Sub SortList()
    If lbFibers.ListCount < 3 Then Exit Sub
    
    Dim strTemp, strTotal As String
    Dim strCurrent, strNext As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 5) As String
    
    iCount = lbFibers.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            strCurrent = lbFibers.List(b, 0)
            Select Case Len(lbFibers.List(b, 1))
                Case Is = 1
                    strCurrent = strCurrent & "00" & lbFibers.List(b, 1)
                Case Is = 2
                    strCurrent = strCurrent & "0" & lbFibers.List(b, 1)
                Case Else
                    strCurrent = strCurrent & lbFibers.List(b, 1)
            End Select
            
            strNew = lbFibers.List(b + 1, 0)
            Select Case Len(lbFibers.List(b + 1, 1))
                Case Is = 1
                    strNew = strNew & "00" & lbFibers.List(b + 1, 1)
                Case Is = 2
                    strNew = strNew & "0" & lbFibers.List(b + 1, 1)
                Case Else
                    strNew = strNew & lbFibers.List(b + 1, 1)
            End Select
            
            
            If strCurrent > strNew Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbFibers.Selected(b) = True
                    lbFibers.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbFibers.List(b + 1, 0)
                strAtt(1) = lbFibers.List(b + 1, 1)
                strAtt(2) = lbFibers.List(b + 1, 2)
                
                lbFibers.List(b + 1, 0) = lbFibers.List(b, 0)
                lbFibers.List(b + 1, 1) = lbFibers.List(b, 1)
                lbFibers.List(b + 1, 2) = lbFibers.List(b, 2)
                
                lbFibers.List(b, 0) = strAtt(0)
                lbFibers.List(b, 1) = strAtt(1)
                lbFibers.List(b, 2) = strAtt(2)
            End If
        Next b
    Next a
End Sub

