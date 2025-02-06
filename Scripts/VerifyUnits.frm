VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VerifyUnits 
   Caption         =   "Verify Units"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10335
   OleObjectBlob   =   "VerifyUnits.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VerifyUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbExport_Click()
    If lbUnits.ListCount < 1 Then Exit Sub
    
    Dim strPath, strName, strFile As String
    Dim vTemp As Variant
    
    strPath = ThisDrawing.Path
    strName = ThisDrawing.Name
    vTemp = Split(strName, " ")
    
    strFile = strPath & "\" & vTemp(0) & " Unit Errors.csv"
    
    Open strFile For Output As #1
    
    Print #1, "Structure Number,Block Units,Callout Units"
    
    For i = 0 To lbUnits.ListCount - 1
        Print #1, lbUnits.List(i, 0) & "," & lbUnits.List(i, 1) & "," & lbUnits.List(i, 2)
    Next i
    
    Close #1
End Sub

Private Sub cbFindDiffer_Click()
    If lbUnits.ListCount < 1 Then Exit Sub
    
    Dim vPole, vUnit As Variant
    Dim strPole, strUnit As String
    
    For i = lbUnits.ListCount - 1 To 0 Step -1
        vPole = Split(lbUnits.List(i, 1), ";;")
        vUnit = Split(lbUnits.List(i, 2), ";;")
        
        For j = 0 To UBound(vPole)
            For k = 0 To UBound(vUnit)
                If vPole(j) = vUnit(k) Then
                    vPole(j) = ""
                    vUnit(k) = ""
                    GoTo Next_Pole
                End If
            Next k
Next_Pole:
        Next j
        
        strPole = ""
        For j = 0 To UBound(vPole)
            If Not vPole(j) = "" Then
                If strPole = "" Then
                    strPole = vPole(j)
                Else
                    strPole = strPole & " & " & vPole(j)
                End If
            End If
        Next j
        strPole = Replace(strPole, "+", "")
        
        strUnit = ""
        For j = 0 To UBound(vUnit)
            If Not vUnit(j) = "" Then
                If strUnit = "" Then
                    strUnit = vUnit(j)
                Else
                    strUnit = strUnit & " & " & vUnit(j)
                End If
            End If
        Next j
        strUnit = Replace(strUnit, "+", "")
        
        If strUnit = "" Then
            If strPole = "" Then
                lbUnits.RemoveItem i
                GoTo Next_I
            End If
        End If
        
        lbUnits.List(i, 1) = strPole
        
        lbUnits.List(i, 2) = strUnit
Next_I:
    Next i
    
    tbListcount.Value = lbUnits.ListCount
End Sub

Private Sub cbGetBlocks_Click()
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim filterType, filterValue As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double

    grpCode(0) = 2
    grpValue(0) = "sPole,sPed,sHH,sFP,sPanel,sMH"
    'grpValue(0) = "pole_unit"
    filterType = grpCode
    filterValue = grpValue
    
    lbUnits.Clear
    
  On Error Resume Next
    Me.Hide
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    dPnt1(0) = vPnt1(0)
    dPnt1(1) = vPnt1(1)
    dPnt1(2) = vPnt1(2)
    
    dPnt2(0) = vPnt2(0)
    dPnt2(1) = vPnt2(1)
    dPnt2(2) = vPnt2(2)
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
        Err = 0
    End If
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        If vAttList(0).TextString = "POLE" Then GoTo Next_objBlock
        If vAttList(0).TextString = "PED" Then GoTo Next_objBlock
        If vAttList(0).TextString = "HH" Then GoTo Next_objBlock
        If vAttList(0).TextString = "PANEL" Then GoTo Next_objBlock
        If vAttList(0).TextString = "MH" Then GoTo Next_objBlock
        
        lbUnits.AddItem vAttList(0).TextString
        Select Case objBlock.Name
            Case "sPole"
                lbUnits.List(lbUnits.ListCount - 1, 1) = vAttList(27).TextString
            Case Else
                lbUnits.List(lbUnits.ListCount - 1, 1) = vAttList(7).TextString
        End Select
        lbUnits.List(lbUnits.ListCount - 1, 2) = ""
        lbUnits.List(lbUnits.ListCount - 1, 3) = objBlock.InsertionPoint(0) & "," & objBlock.InsertionPoint(1)
Next_objBlock:
    Next objBlock
    
    objSS.Clear

    grpCode(0) = 2
    grpValue(0) = "pole_unit"
    filterType = grpCode
    filterValue = grpValue
    
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        For i = 0 To lbUnits.ListCount - 1
            If vAttList(0).TextString = lbUnits.List(i, 0) Then
                If lbUnits.List(i, 2) = "" Then
                    lbUnits.List(i, 2) = vAttList(3).TextString
                Else
                    lbUnits.List(i, 2) = lbUnits.List(i, 2) & ";;" & vAttList(3).TextString
                End If
                
                GoTo Next_Pole_Unit
            End If
        Next i
            
        lbUnits.AddItem vAttList(0).TextString
        lbUnits.List(lbUnits.ListCount - 1, 1) = ""
        lbUnits.List(lbUnits.ListCount - 1, 2) = vAttList(3).TextString
        lbUnits.List(lbUnits.ListCount - 1, 3) = objBlock.InsertionPoint(0) & "," & objBlock.InsertionPoint(1)
        
Next_Pole_Unit:
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    tbListcount.Value = lbUnits.ListCount
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub lbUnits_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim vCoords As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    vCoords = Split(lbUnits.List(lbUnits.ListIndex, 3), ",")
    
    viewCoordsB(0) = CDbl(vCoords(0)) - 300
    viewCoordsB(1) = CDbl(vCoords(1)) - 300
    viewCoordsB(2) = 0#
    viewCoordsE(0) = CDbl(vCoords(0)) + 300
    viewCoordsE(1) = CDbl(vCoords(1)) + 300
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbUnits.ColumnCount = 4
    lbUnits.ColumnWidths = "120;192;180;6"
End Sub
