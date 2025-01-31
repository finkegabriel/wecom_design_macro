VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddClearances 
   Caption         =   "Add CLR"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1905
   OleObjectBlob   =   "AddClearances.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddClearances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAdd_Click()
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objBlock As AcadBlockReference
    Dim objCLR As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine, vAttach, vTemp As Variant
    Dim iTest, iClr As Integer
    Dim strLine As String
    Dim dRotate As Double
    Dim vInsert As Variant
    
    vLine = Split(tbCLR.Value, "-")
    iTest = CInt(vLine(0)) * 12
    If UBound(vLine) > 0 Then iTest = iTest + CInt(vLine(1))
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "cable_span"
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
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        
        vLine = Split(UCase(vAttList(0).TextString), " ")
        If UBound(vLine) < 1 Then GoTo Next_objBlock
        
        Select Case vLine(0)
            Case "M", "MS"
                strLine = "MIDSPAN CLR"
            Case "D", "DW"
                strLine = "DRIVEWAY CLR"
            Case "R", "RD"
                strLine = "ROAD CLR"
        End Select
        
        vAttach = Split(vLine(1), "-")
        iClr = CInt(vAttach(0)) * 12
        If UBound(vAttach) > 0 Then iClr = iClr + CInt(vAttach(1))
        
        If iClr > iTest Then GoTo Next_objBlock
        
        dRotate = objBlock.Rotation
        vInsert = objBlock.InsertionPoint
        
        Set objCLR = ThisDrawing.ModelSpace.InsertBlock(vInsert, "clr", 1#, 1#, 1#, dRotate)
        objCLR.Layer = "Integrity Roads-Clearance"
        
        vAttList = objCLR.GetAttributes
        vAttList(0).TextString = strLine
        vAttList(1).TextString = vAttach(0) & "'"
        If UBound(vAttach) > 0 Then vAttList(1).TextString = vAttList(1).TextString & vAttach(1) & """"
        
        objCLR.Update
Next_objBlock:
    Next objBlock
    
    objSS.Clear
    objSS.Delete
    
    Me.Hide
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub CommandButton1_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim objCLR As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    Dim vLine, vAttach, vTemp As Variant
    Dim strLine, strHeight As String
    Dim dRotate As Double
    Dim vInsert As Variant
    
    On Error Resume Next
    
    Me.Hide
    
Place_Block:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Span: "
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    
    If Not objBlock.Name = "cable_span" Then GoTo Exit_Sub
    
    vAttList = objBlock.GetAttributes
    
    If vAttList(0).TextString = "" Then
        MsgBox "No clearance"
        GoTo Place_Block
    End If
        
    vLine = Split(UCase(vAttList(0).TextString), " ")
    If UBound(vLine) < 1 Then
        MsgBox "Invalid clearance"
        GoTo Place_Block
    End If
        
    Select Case vLine(0)
        Case "M", "MS"
            strLine = "MIDSPAN CLR"
        Case "D", "DW"
            strLine = "DRIVEWAY CLR"
        Case "R", "RD"
            strLine = "ROAD CLR"
    End Select
    
    strHeight = Replace(vLine(1), "-", "'") & """"
        
    dRotate = objBlock.Rotation
    vInsert = objBlock.InsertionPoint
        
    Set objCLR = ThisDrawing.ModelSpace.InsertBlock(vInsert, "clr", 1#, 1#, 1#, dRotate)
    objCLR.Layer = "Integrity Roads-Clearance"
        
    vAttList = objCLR.GetAttributes
    vAttList(0).TextString = strLine
    vAttList(1).TextString = strHeight
        
    objCLR.Update
    
    GoTo Place_Block
    
Exit_Sub:
    Me.Hide
End Sub
