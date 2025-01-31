VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindMissingCallouts 
   Caption         =   "Find Missing Callouts"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   OleObjectBlob   =   "FindMissingCallouts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindMissingCallouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbMark_Click()
    If lbPoles.ListCount < 1 Then Exit Sub
    
    For i = 0 To lbPoles.ListCount - 1
        If Not lbPoles.List(i, 2) = "x" And Not lbPoles.List(i, 3) = "x" Then lbPoles.Selected(i) = True
        If lbPoles.List(i, 3) = "x" And i > 0 Then lbPoles.Selected(i - 1) = False
    Next i
End Sub

Private Sub cbSort_Click()
    Call SortPoles
End Sub

Private Sub cbWindow_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vItem, vLine, vTemp, vL, vR As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim strLine As String
    Dim iIndex, iPos, iLength As Integer
    
    On Error Resume Next
    
    Me.Hide
    
    lbPoles.Clear
        
    Err = 0
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    End If
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = "POLE" Then GoTo Next_objBlock
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        
        Select Case objBlock.Name
            Case "sPole"
                If vAttList(25).TextString = "" Then GoTo Next_objBlock
                iPos = 26
            Case Else
                If vAttList(5).TextString = "" Then GoTo Next_objBlock
                iPos = 6
        End Select
        
        strLine = UCase(vAttList(0).TextString)
        vTemp = Split(strLine, "/")
        vL = Split(vTemp(UBound(vTemp)), "L")
        vR = Split(vL(UBound(vL)), "R")
        
        iLength = Len(strLine) - Len(vR(UBound(vR)))
        
        lbPoles.AddItem Left(strLine, iLength)
        iIndex = lbPoles.ListCount - 1
        lbPoles.List(iIndex, 1) = vR(UBound(vR))
        If vAttList(iPos).TextString = "" Then
            lbPoles.List(iIndex, 2) = " "
        Else
            lbPoles.List(iIndex, 2) = "M"
        End If
        lbPoles.List(iIndex, 3) = " "
        lbPoles.List(iIndex, 4) = objBlock.InsertionPoint(0) & "," & objBlock.InsertionPoint(1)
Next_objBlock:
    Next objBlock
    
    objSS.Clear
    'Call SortPoles
    
    grpCode(0) = 2
    grpValue(0) = "Callout"
    filterType = grpCode
    filterValue = grpValue
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = "" Then GoTo Next_Callout
        
        strLine = vAttList(0).TextString
        vTemp = Split(strLine, ": ")
        For i = 0 To lbPoles.ListCount - 1
            If vTemp(0) = lbPoles.List(i, 0) & lbPoles.List(i, 1) Then
                If Left(vAttList(1).TextString, 4) = "+HO1" Then
                    lbPoles.List(i, 2) = "x"
                Else
                    lbPoles.List(i, 3) = "x"
                End If
                GoTo Next_Callout
            End If
        Next i
        
Next_Callout:
    Next objBlock
    
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub lbPoles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vCoords As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    On Error Resume Next
    
    vCoords = Split(lbPoles.List(lbPoles.ListIndex, 4), ",")
    
    viewCoordsB(0) = vCoords(0) - 300
    viewCoordsB(1) = vCoords(1) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = vCoords(0) + 300
    viewCoordsE(1) = vCoords(1) + 300
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    
    Load PlaceCountCallouts
        PlaceCountCallouts.show
    Unload PlaceCountCallouts
    
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbPoles.ColumnCount = 5
    lbPoles.ColumnWidths = "120;48;36;30;6"
End Sub

Private Sub SortPoles()
    Dim strTemp, strTotal As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 4) As String
    'Dim strTemp As String
    Dim iB, iB1 As Integer
    
    iCount = lbPoles.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    'On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            Select Case lbPoles.List(b, 0)
                Case Is < lbPoles.List(b + 1, 0)
                    GoTo Next_line
                Case Is = lbPoles.List(b + 1, 0)
                    'If CInt(lbPoles.List(b, 1)) < CInt(lbPoles.List(b + 1, 1)) Then GoTo Next_Line
                    strTemp = Replace(lbPoles.List(b, 1), "X", "")
                    iB = CInt(strTemp)
                    
                    strTemp = Replace(lbPoles.List(b + 1, 1), "X", "")
                    iB1 = CInt(strTemp)
                        
                    If iB < iB1 Then GoTo Next_line
            End Select
                
            strAtt(0) = lbPoles.List(b + 1, 0)
            strAtt(1) = lbPoles.List(b + 1, 1)
            strAtt(2) = lbPoles.List(b + 1, 2)
            strAtt(3) = lbPoles.List(b + 1, 3)
            strAtt(4) = lbPoles.List(b + 1, 4)
                
            lbPoles.List(b + 1, 0) = lbPoles.List(b, 0)
            lbPoles.List(b + 1, 1) = lbPoles.List(b, 1)
            lbPoles.List(b + 1, 2) = lbPoles.List(b, 2)
            lbPoles.List(b + 1, 3) = lbPoles.List(b, 3)
            lbPoles.List(b + 1, 4) = lbPoles.List(b, 4)
                
            lbPoles.List(b, 0) = strAtt(0)
            lbPoles.List(b, 1) = strAtt(1)
            lbPoles.List(b, 2) = strAtt(2)
            lbPoles.List(b, 3) = strAtt(3)
            lbPoles.List(b, 4) = strAtt(4)
Next_line:
        Next b
    Next a
End Sub
