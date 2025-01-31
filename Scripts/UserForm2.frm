VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   8880.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12240
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objStructure As AcadBlockReference

Private Sub cbGetStructure_Click()
    Dim objEntity As AcadEntity
    Dim vAttList, vReturnPnt As Variant
    Dim vLine, vItem, vCounts As Variant
    Dim vCable, vSplice As Variant
    Dim strText, strBack As String
    Dim strCables, strSplices, strTemp As String
    Dim result, iAtt As Integer
    Dim bTest As Boolean
    
    'bTest = False
    
    'If Not tbCableCounts.Value = strMainCC Then bTest = True
    'If Not tbClosure.Value = strMainSC Then bTest = True
    'If strMainCC = "empty" Then
        'If strMainSC = "empty" Then bTest = False
    'End If
    'If bTest = True Then
        'result = MsgBox("Save count changes to Block?", vbYesNo, "Save Changes")
        'If result = vbYes Then Call UpdateBlock(CStr(tbPosition.Value))
    'End If
    
    Me.Hide
    
    On Error Resume Next
    
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Structure:"
    If Not Err = 0 Then
        MsgBox Err.Description
        GoTo Exit_Sub
    End If
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox objEntity.ObjectName
        GoTo Exit_Sub
    End If
    
    lbCounts.Clear
    tbClosure.Value = ""
    tbCableCounts.Value = ""
    tbFutureCableCounts.Value = ""
    tbType.Value = ""
    tbNumber.Value = ""
    tbPosition.Value = ""
    
    Set objStructure = objEntity
    
    Select Case objStructure.Name
        Case "sPole"
            iAtt = 25
            tbType.Value = "POLE"
        Case "sPed"
            iAtt = 5
            tbType.Value = "PED"
        Case "sHH"
            iAtt = 5
            tbType.Value = "HH"
        Case "sMH"
            iAtt = 5
            tbType.Value = "MH"
        Case "sPanel"
            iAtt = 5
            tbType.Value = "PANEL"
        Case Else
            GoTo Exit_Sub
    End Select
    
    'cbScale.Value = GetScale(objStructure.InsertionPoint) * 100
    
    vAttList = objStructure.GetAttributes
    
    'MsgBox vAttList(iAtt).TextString
    strCables = Replace(vAttList(iAtt).TextString, vbLf, "")
    strCables = Replace(strCables, "\P", vbCr)
    vCable = Split(strCables, vbCr)
    
    cbPosition.Clear
    For i = 0 To UBound(vCable)
        vCounts = Split(vCable(i), ": ")
        cbPosition.AddItem vCounts(0)
    Next i
    cbPosition.Value = cbPosition.List(0)
    
    Call GetPositionData(cbPosition.Value)
    
    If tbNumber.Value = "" Then tbNumber.Value = vAttList(0).TextString
    
    
    strText = Replace(tbCableCounts.Value, vbLf, "")
    strText = Replace(strText, vbTab, " ")
    vCable = Split(strText, vbCr)
    
    For i = 0 To UBound(vCable)
        vLine = Split(vCable(i), ": ")
        vCable(i) = vLine(0) & ": " & vLine(1) & ": " & tbNumber.Value
    Next i
    
    strText = vCable(0)
    If UBound(vCable) > 0 Then
        For i = 1 To UBound(vCable)
            strText = strText & vbCr & vCable(i)
        Next i
    End If
    
    strText = Replace(strText, " ", vbTab)
    tbFutureCableCounts.Value = strText
    
    Call CreateList
    
    
    strText = Replace(tbClosure.Value, vbLf, "")
    strText = Replace(strText, vbTab, " ")
    vSplice = Split(strText, vbCr)
    
    For i = 0 To UBound(vSplice)
        vLine = Split(vSplice(i), ": ")
        strText = vLine(0)
    Next i
    
    tbF2.Value = strText
    
    vLine = Split(ThisDrawing.Name, " ")
    tbFileName.Value = vLine(0) & " Counts -" & strText & ".mcl"
    
    Call OpenMCL
    
Exit_Sub:
    
    Me.show
End Sub

Private Sub GetPositionData(strPosition As String)
    Dim vAttList As Variant
    Dim vLine, vItem, vCounts As Variant
    Dim vCable, vSplice As Variant
    Dim strText, strBack As String
    Dim strCables, strSplices, strTemp As String
    Dim iAtt As Integer
    
    On Error Resume Next
    
    tbCableCounts.Value = ""
    tbClosure.Value = ""
    'tbWL.Value = ""
    
    Select Case objStructure.Name
        Case "sPole"
            iAtt = 25
            tbType.Value = "POLE"
        Case "sPed"
            iAtt = 5
            tbType.Value = "PED"
        Case "sHH"
            iAtt = 5
            tbType.Value = "HH"
        Case "sPanel"
            iAtt = 5
            tbType.Value = "PANEL"
        Case "sMH"
            iAtt = 5
            tbType.Value = "MH"
        Case Else
            MsgBox objStructure.Name & vbCr & "Why?"
            GoTo Exit_Sub
    End Select
    
    vAttList = objStructure.GetAttributes
    
    strCables = Replace(vAttList(iAtt).TextString, vbLf, "")
    strCables = Replace(strCables, "\P", vbCr)
    vCable = Split(strCables, vbCr)
    
    For i = 0 To UBound(vCable)
        vItem = Split(vCable(i), ": ")
        If vItem(0) = strPosition Then GoTo Found_Cable
    Next i
    
    'MsgBox "No Cable found at that Position."
    GoTo Find_Splice
    
Found_Cable:
    vLine = Split(vCable(i), " / ")
    vItem = Split(vLine(0), ": ")
    tbNumber.Value = vAttList(0).TextString
    tbPosition.Value = vItem(0)
    tbCableType.Value = vItem(1)
            
    strText = Replace(vLine(1), " + ", vbCr)
    strText = Replace(strText, " ", vbTab)
    tbCableCounts.Value = strText
    
Find_Splice:
            
    strSplices = Replace(vAttList(iAtt + 1).TextString, vbLf, "")
    
    If Left(strSplices, 1) = vbCrLf Then strSplices = Right(strSplices, Len(strSplices) - 1)
    If Left(strSplices, 1) = vbCr Then strSplices = Right(strSplices, Len(strSplices) - 1)
    If Left(strSplices, 1) = vbLf Then strSplices = Right(strSplices, Len(strSplices) - 1)
    
    If Right(strSplices, 1) = vbCrLf Then strSplices = Left(strSplices, Len(strSplices) - 1)
    If Right(strSplices, 1) = vbCr Then strSplices = Left(strSplices, Len(strSplices) - 1)
    If Right(strSplices, 1) = vbLf Then strSplices = Left(strSplices, Len(strSplices) - 1)
    
    strSplices = Replace(strSplices, "\P", vbCr)
    
    vAttList(iAtt + 1).TextString = strSplices
    objStructure.Update
    
    vSplice = Split(strSplices, vbCr)
    
    For i = 0 To UBound(vSplice)
        If vSplice(i) = "" Then GoTo Next_I
        
        vItem = Split(vSplice(i), "] ")
        strTemp = Replace(vItem(0), "[", "")
        If tbPosition.Value = "" Then tbPosition.Value = strTemp
        If strTemp = strPosition Then GoTo Found_Splice
Next_I:
    Next i
    
    'MsgBox "No Splice found at that Position."
    GoTo Get_Wiring_Limits
    
Found_Splice:
    
    If Not vSplice(0) = "" Then
        For i = 0 To UBound(vSplice)
            vCounts = Split(vSplice(i), "] ")
            strTemp = Replace(vCounts(0), "[", "")
            If strTemp = strPosition Then
                strText = Replace(vCounts(1), " + ", vbCr)
                strText = Replace(strText, " ", vbTab)
                tbClosure.Value = strText
                
                GoTo Get_Wiring_Limits
            End If
        Next i
    End If
    
Get_Wiring_Limits:
    GoTo Clear_objSS

    tbWL.Value = ""
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    
    grpCode(0) = 2
    grpValue(0) = "Customer,SG"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        Select Case objBlock.Name
            Case "SG"
                If vAttList(2).TextString = "" Then GoTo Next_objBlock
                vLine = Split(vAttList(2).TextString, " - ")
                strText = vLine(1) & vbTab & "SG - " & vAttList(1).TextString
            Case Else
                If vAttList(4).TextString = "" Then GoTo Next_objBlock
                vLine = Split(vAttList(4).TextString, " - ")
                strText = vLine(1) & vbTab & vAttList(1).TextString & " " & vAttList(2).TextString
        End Select
        
        If vLine(0) = tbNumber.Value Then
            'vItem = Split(vLine(1), ": ")
            'strText = Replace(vItem(1), ")", "")
            
            If tbWL.Value = "" Then
                'tbWL.Value = strText & vbTab & vAttList(1).TextString & " " & vAttList(2).TextString
                tbWL.Value = strText
            Else
                'tbWL.Value = tbWL.Value & vbCr & strText & vbTab & vAttList(1).TextString & " " & vAttList(2).TextString
                tbWL.Value = tbWL.Value & vbCr & strText
            End If
            
        End If
Next_objBlock:
    Next objBlock
    
    'Call SortWL
    
Clear_objSS:
    objSS.Clear
    objSS.Delete
        
Exit_Sub:
    cbUpdate.Enabled = True
    
    strMainCC = tbCableCounts.Value
    strMainSC = tbClosure.Value
End Sub

Private Sub UserForm_Initialize()
    lbCounts.Clear
    lbCounts.ColumnCount = 9
    lbCounts.ColumnWidths = "24;72;26;72;72;36;140;48;102"
End Sub

Private Sub CreateList()
    If tbCableCounts.Value = "" Then Exit Sub
    
    Dim vLine, vItem, vCounts As Variant
    Dim strText As String
    Dim iStart, iEnd As Integer
    Dim iFiber As Integer
    
    lbCounts.AddItem ""
    iFiber = 1
    
    strText = Replace(tbCableCounts.Value, vbLf, "")
    strText = Replace(strText, vbTab, " ")
    vLine = Split(strText, vbCr)
    
    For i = 0 To UBound(vLine)
        vItem = Split(vLine(i), ": ")
        vCounts = Split(vItem(1), "-")
        
        iStart = CInt(vCounts(0))
        
        If UBound(vCounts) = 0 Then
            iEnd = iStart
        Else
            iEnd = CInt(vCounts(1))
        End If
        
        For j = iStart To iEnd
            lbCounts.AddItem iFiber
            lbCounts.List(iFiber, 1) = vItem(0)
            lbCounts.List(iFiber, 2) = j
            lbCounts.List(iFiber, 3) = vItem(2)
            lbCounts.List(iFiber, 4) = "<>"
            lbCounts.List(iFiber, 5) = "<>"
            lbCounts.List(iFiber, 6) = "<>"
            lbCounts.List(iFiber, 7) = "<>"
            lbCounts.List(iFiber, 8) = "<>"
            
            iFiber = iFiber + 1
        Next j
    Next i
End Sub

Private Sub OpenMCL()
    Dim vLine, vItem As Variant
    Dim strFile, strAllText As String
    Dim fName As String
    Dim iIndex As Integer
    
    strFile = ThisDrawing.Path & "\" & tbFileName.Value
    
    fName = Dir(strFile)
    If fName = "" Then
        tbFileStatus.Value = "missing"
        Exit Sub
    Else
        tbFileStatus.Value = "found"
    End If
    
    Open strFile For Input As #2
    
    strAllText = Input(LOF(2), 2)
    strAllText = Replace(strAllText, vbLf, "")
    vLine = Split(strAllText, vbCr)
    
    Close #2
    
    For i = 1 To lbCounts.ListCount - 1
        If lbCounts.List(i, 1) = tbF2.Value Then
            iIndex = CInt(lbCounts.List(i, 2))
            
            If iIndex > UBound(vLine) + 1 Then GoTo Next_line
            
            vItem = Split(vLine(iIndex), vbTab)
            
            lbCounts.List(i, 4) = vItem(1)
            lbCounts.List(i, 5) = vItem(2)
            lbCounts.List(i, 6) = vItem(3)
            lbCounts.List(i, 7) = vItem(4)
            lbCounts.List(i, 8) = vItem(5)
        End If
        
        If lbCounts.List(i, 4) = tbNumber.Value Then lbCounts.List(i, 4) = "current pole"
Next_line:
    Next i
    
    
    'MsgBox UBound(vLine) & vbCr & vLine(0) & vbCr & vLine(1) & vbCr & vLine(UBound(vLine) - 1)
    
    'While Not EOF(2)
        'Input #2, strLine
        
    'Wend
    
        
End Sub
