VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArcViewCustomer 
   Caption         =   "Customers"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11280
   OleObjectBlob   =   "ArcViewCustomer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArcViewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbClear_Click()
    lbCustomers.Clear
End Sub

Private Sub cbConvert_Click()
    Dim objSSTemp4 As AcadSelectionSet
    Dim entItem As AcadEntity
    Dim obrTemp, objBlock As AcadBlockReference
    Dim attList, vAttList2 As Variant
    Dim vTemp As Variant
    Dim filterType, filterValue As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim strType, strSymbol As String

    grpCode(0) = 2
    grpValue(0) = "RES,BUSINESS,TRLR,MDU,SCHOOL,CHURCH,LOT"
    filterType = grpCode
    filterValue = grpValue
    
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
  
    'Err = 0
    
    Set objSSTemp4 = ThisDrawing.SelectionSets.Add("objSSTemp4")
    If Not Err = 0 Then
        Set objSSTemp4 = ThisDrawing.SelectionSets.Item("objSSTemp4")
        Err = 0
    End If
    objSSTemp4.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    For Each entItem In objSSTemp4
        If TypeOf entItem Is AcadBlockReference Then
            Set obrTemp = entItem
            
            'If obrTemp.Layer = "Integrity Future" Then GoTo Next_Object
            
            Select Case obrTemp.Name
                Case "RES"
                    strType = "RESIDENCE"
                    strSymbol = ""
                Case "TRLR"
                    strType = "TRAILER"
                    strSymbol = "T"
                Case "BUSINESS", "SCHOOL", "CHURCH", "MDU", "LOT", "VACANT", "ABANDONED"
                    strType = obrTemp.Name
                    strSymbol = Left(strType, 1)
                Case Else
                    GoTo Next_Object
            End Select
            
            attList = obrTemp.GetAttributes
            vTemp = obrTemp.InsertionPoint
            
            Set objBlock = ThisDrawing.ModelSpace.InsertBlock(vTemp, "Customer", 1#, 1#, 1#, 0#)
            objBlock.Layer = "Customers"
            
            vAttList2 = objBlock.GetAttributes
            vAttList2(0).TextString = strType
            vAttList2(1).TextString = attList(0).TextString
            vAttList2(2).TextString = attList(1).TextString
            vAttList2(3).TextString = attList(2).TextString
            vAttList2(4).TextString = ""
            vAttList2(5).TextString = strSymbol
            
            objBlock.Update
        End If
Next_Object:
    Next entItem
    
Exit_Sub:
    objSSTemp4.Clear
    objSSTemp4.Delete
    Me.show
End Sub

Private Sub cbGetCustomers_Click()
    Dim objSSTemp4 As AcadSelectionSet
    Dim entItem As AcadEntity
    Dim obrTemp, objBlock As AcadBlockReference
    Dim vAttList, vAttList2 As Variant
    Dim vTemp As Variant
    Dim filterType, filterValue As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
    Dim strType, strSymbol As String
    Dim strTemp As String
    Dim iIndex As Integer

    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    
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
  
    'Err = 0
    
    Set objSSTemp4 = ThisDrawing.SelectionSets.Add("objSSTemp4")
    If Not Err = 0 Then
        Set objSSTemp4 = ThisDrawing.SelectionSets.Item("objSSTemp4")
        Err = 0
    End If
    objSSTemp4.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    
    For Each entItem In objSSTemp4
        If TypeOf entItem Is AcadBlockReference Then
            Set obrTemp = entItem
            
            vAttList = obrTemp.GetAttributes
            
            If vAttList(4).TextString = "" Then
                strTemp = "<empty> - <empty>"
            Else
                strTemp = vAttList(4).TextString
            End If
            
            If InStr(strTemp, " (") > 0 Then
                strTemp = Replace(strTemp, " (", " RC ")
                strTemp = Replace(strTemp, ")", "")
            End If
            
            vTemp = Split(strTemp, " - ")
            lbCustomers.AddItem vTemp(0)
            
            iIndex = lbCustomers.ListCount - 1
            lbCustomers.List(iIndex, 1) = vTemp(1)
            lbCustomers.List(iIndex, 2) = vAttList(1).TextString
            lbCustomers.List(iIndex, 3) = vAttList(2).TextString
            If Not vAttList(0).TextString = "" Then lbCustomers.List(iIndex, 4) = vAttList(3).TextString
            lbCustomers.List(iIndex, 5) = vAttList(0).TextString
        End If
    Next entItem
    
Exit_Sub:
    objSSTemp4.Clear
    objSSTemp4.Delete
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub Label42_Click()
    If lbCustomers.ListCount = 0 Then Exit Sub
    
    Dim strPole1, strCount, strHSE As String
    Dim strRoad, strNote, strType As String
    Dim strTest1, strTest2 As String
    Dim vTest As Variant
    
    Dim strPole2, strOwner2, strDWG2 As String
    Dim strLower2, strRTA2 As Integer
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    
    On Error Resume Next
    Err = 0
    
    iCount = lbCustomers.ListCount - 1
    
    'For c = 0 To iCount
        'If IsNull(lbCustomers.List(c, 0)) Then lbCustomers.List(c, 0) = "-"
        'If IsNull(lbCustomers.List(c, 1)) Then lbCustomers.List(c, 1) = "-"
        'If IsNull(lbCustomers.List(c, 2)) Then lbCustomers.List(c, 2) = "0"
        'If IsNull(lbCustomers.List(c, 3)) Then lbCustomers.List(c, 3) = "0"
        'If IsNull(lbCustomers.List(c, 4)) Then lbCustomers.List(c, 4) = "-"
        
        'If lbCustomers.List(c, 0) = "" Then lbCustomers.List(c, 0) = "-"
        'If lbCustomers.List(c, 1) = "" Then lbCustomers.List(c, 1) = "-"
        'If lbCustomers.List(c, 2) = "" Then lbCustomers.List(c, 2) = "0"
        'If lbCustomers.List(c, 3) = "" Then lbCustomers.List(c, 3) = "0"
        'If lbCustomers.List(c, 4) = "" Then lbCustomers.List(c, 4) = "-"
    'Next c
    
    For a = iCount To 1 Step -1
        For b = 0 To a - 1
            strTest1 = Replace(lbCustomers.List(b, 1), ")", "")
            vTest = Split(strTest1, ": ")
            strTest1 = vTest(1)
            
            strTest2 = Replace(lbCustomers.List(b + 1, 1), ")", "")
            vTest = Split(strTest2, ": ")
            strTest2 = vTest(1)
                        
            'If lbCustomers.List(b, 1) > lbCustomers.List(b + 1, 1) Then
            If CInt(strTest1) > CInt(strTest2) Then
                strPole1 = lbCustomers.List(b + 1, 0)
                strCount = lbCustomers.List(b + 1, 1)
                strHSE = lbCustomers.List(b + 1, 2)
                strRoad = lbCustomers.List(b + 1, 3)
                strNote = lbCustomers.List(b + 1, 4)
                strType = lbCustomers.List(b + 1, 5)
                
                lbCustomers.List(b + 1, 0) = lbCustomers.List(b, 0)
                lbCustomers.List(b + 1, 1) = lbCustomers.List(b, 1)
                lbCustomers.List(b + 1, 2) = lbCustomers.List(b, 2)
                lbCustomers.List(b + 1, 3) = lbCustomers.List(b, 3)
                lbCustomers.List(b + 1, 4) = lbCustomers.List(b, 4)
                lbCustomers.List(b + 1, 5) = lbCustomers.List(b, 5)
                                
                lbCustomers.List(b, 0) = strPole1
                lbCustomers.List(b, 1) = strCount
                lbCustomers.List(b, 2) = strHSE
                lbCustomers.List(b, 3) = strRoad
                lbCustomers.List(b, 4) = strNote
                lbCustomers.List(b, 5) = strType
            End If
        Next b
    Next a
End Sub

Private Sub UserForm_Initialize()
    lbCustomers.ColumnCount = 6
    lbCustomers.ColumnWidths = "102;90;36;140;114;68"
End Sub
