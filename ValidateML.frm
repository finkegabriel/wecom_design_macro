VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateML 
   Caption         =   "Validate Matchlines"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7485
   OleObjectBlob   =   "ValidateML.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbValidate_Click()
    Dim iDWG, iFrom, iTestDWG As Integer
    Dim strFrom, strTo, strDWG As String
    Dim strTestFrom, strTestTo, strTestDWG As String
    Dim vFrom, vTo As Variant
    
    For i = 0 To lbMatches.ListCount - 1
        strFrom = lbMatches.List(i, 0)
        strDWG = lbMatches.List(i, 1)
        strTo = lbMatches.List(i, 2)
        
        For j = 0 To lbMatches.ListCount - 1
            strTestFrom = lbMatches.List(j, 0)
            strTestDWG = lbMatches.List(j, 1)
            strTestTo = lbMatches.List(j, 2)
            
            If InStr(strFrom, strTestDWG) Then
                If InStr(strTestTo, strDWG) Then
                    strTestTo = Replace(strTestTo, strDWG, "")
                    strTestTo = Replace(strTestTo, "  ", " ")
                    lbMatches.List(j, 2) = strTestTo
                    
                    strFrom = Replace(strFrom, strTestDWG, "")
                    strFrom = Replace(strFrom, "  ", " ")
                    lbMatches.List(i, 0) = strFrom
                End If
                
            End If
        Next j
    Next i
    
    For i = lbMatches.ListCount - 1 To 0 Step -1
        strFrom = Replace(lbMatches.List(i, 0), " ", "")
        strTo = Replace(lbMatches.List(i, 2), " ", "")
        If strFrom = "" Then
            If strTo = "" Then lbMatches.RemoveItem i
        End If
    Next i
End Sub

Private Sub lbMatches_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim str As String
    Dim strArray As Variant
    Dim viewCoordsB(0 To 2) As Double
    Dim viewCoordsE(0 To 2) As Double
    
    Me.Hide
    
    'str = lbMatches.List(lbMatches.ListIndex)
    'strArray = Split(str, vbTab)
    
    viewCoordsB(0) = lbMatches.List(lbMatches.ListIndex, 3)
    viewCoordsB(1) = lbMatches.List(lbMatches.ListIndex, 4)
    viewCoordsB(2) = 0#
    viewCoordsE(0) = viewCoordsB(0) + 1650 * CDbl(lbMatches.List(lbMatches.ListIndex, 5))
    viewCoordsE(1) = viewCoordsB(1) + 1050 * CDbl(lbMatches.List(lbMatches.ListIndex, 5))
    viewCoordsE(2) = 0#
    
    ThisDrawing.Application.ZoomWindow viewCoordsB, viewCoordsE
    Me.show
End Sub

Private Sub UserForm_Initialize()
    lbMatches.Clear
    lbMatches.ColumnCount = 6
    lbMatches.ColumnWidths = "42;42;200;80;80;20"
    
    Dim objSS3 As AcadSelectionSet
    Dim entBlock As AcadObject
    Dim obrTemp As AcadBlockReference
    Dim attItem2, vTemp As Variant
    Dim mode As Integer
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim str, strList As String
    Dim iCount As Integer
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    
  On Error Resume Next
  
    strList = ""
    iCount = 0
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    dPnt1(0) = vPnt1(0)
    dPnt1(1) = vPnt1(1)
    dPnt1(2) = vPnt1(2)
    
    dPnt2(0) = vPnt2(0)
    dPnt2(1) = vPnt2(1)
    dPnt2(2) = vPnt2(2)
    
    grpCode(0) = 2
    grpValue(0) = "SS-11x17"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS3 = ThisDrawing.SelectionSets.Add("objSS3")
    If Not Err = 0 Then Set objSS3 = ThisDrawing.SelectionSets.List("objSS3")
    
    objSS3.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    'MsgBox objSS3.count
    
    Err = 0
    For Each entBlock In objSS3
        Set obrTemp = entBlock
        attItem2 = obrTemp.GetAttributes
        
        vTemp = Split(attItem2(0).TextString, " ")
        
        If UBound(vTemp) < 1 Then GoTo Next_entBlock
        
        lbMatches.AddItem
        lbMatches.List(iCount, 0) = ""
        lbMatches.List(iCount, 1) = vTemp(1)
        lbMatches.List(iCount, 2) = ""
        lbMatches.List(iCount, 3) = obrTemp.InsertionPoint(0)
        lbMatches.List(iCount, 4) = obrTemp.InsertionPoint(1)
        lbMatches.List(iCount, 5) = obrTemp.XScaleFactor
        
        iCount = iCount + 1
Next_entBlock:
        Err = 0
    Next entBlock
    
    tbDWGType.Value = vTemp(0)
    
Exit_Sub:
    objSS3.Delete
    
    Call GetMatchlines
    
End Sub

Private Sub GetMatchlines()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objText As AcadText
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim strLayer As String
    
    Dim iDWG As Integer
    Dim strDWG As String
    Dim dLL(0 To 2) As Double
    Dim dUR(0 To 2) As Double
    Dim dScale As Double
    
    Dim vTemp As Variant
    Dim iTest As Integer
    Dim strText As String
    
    On Error Resume Next
    
    Select Case tbDWGType.Value
        Case ""
            strLayer = "Integrity Sheets"
        Case "DWG"
            strLayer = "Integrity Sheets"
        Case Else
            strLayer = "Integrity Permits-" & tbDWGType.Value
    End Select
    
    grpCode(0) = 8
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Clear
    
    For i = 0 To lbMatches.ListCount - 1
        If lbMatches.List(i, 1) = "" Then GoTo Next_I
        
        strDWG = lbMatches.List(i, 1)
        iDWG = CInt(lbMatches.List(i, 1))
        dLL(0) = CDbl(lbMatches.List(i, 3))
        dLL(1) = CDbl(lbMatches.List(i, 4))
        dLL(2) = 0#
        
        dScale = CDbl(lbMatches.List(i, 5))
        dUR(0) = dLL(0) + (1652 * dScale)
        dUR(1) = dLL(1) + (1052 * dScale)
        dUR(2) = 0#
        
        'MsgBox iDWG & vbCr & dLL(0) & vbCr & dLL(1) & vbCr & dScale
        'MsgBox dUR(0) & vbCr & dUR(1)
        
        objSS.Select acSelectionSetWindow, dLL, dUR, filterType, filterValue
        If Not Err = 0 Then
            MsgBox Err.Description
            Err = 0
        End If
        
        'MsgBox objSS.count
        'GoTo Next_I
        
        For Each objEntity In objSS
            'MsgBox objEntity.ObjectName
            If TypeOf objEntity Is AcadText Then
                Set objText = objEntity
                vTemp = Split(objText.TextString, " ")
                strTest = vTemp(2)
                iTest = CInt(vTemp(2))
                
                'MsgBox objText.TextString & vbCr & vTemp(2) & vbCr & iTest
                
                If iTest > iDWG Then
                    If lbMatches.List(i, 2) = "" Then
                        lbMatches.List(i, 2) = strTest
                    Else
                        lbMatches.List(i, 2) = lbMatches.List(i, 2) & " " & strTest
                    End If
                Else
                    If lbMatches.List(i, 0) = "" Then
                        lbMatches.List(i, 0) = strTest
                    Else
                        lbMatches.List(i, 0) = lbMatches.List(i, 0) & " " & strTest
                    End If
                End If
            End If
        Next objEntity
        
        objSS.Clear
Next_I:
    Next i
End Sub
