VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddBLDGForm 
   Caption         =   "Add Buildings"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5295
   OleObjectBlob   =   "AddBLDGForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddBLDGForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbPlace_Click()
    Dim objBldg As AcadBlockReference
    Dim returnPnt, vAttList As Variant
    Dim dScale As Double
    'Dim strBlock As String
    
    Me.Hide
  On Error Resume Next
    
    dScale = CDbl(cbScale.Value) / 100
    If Not Err = 0 Then
        dScale = 1#
        Err = 0
    End If
    
Place_Block:
    Err = 0
  
    returnPnt = ThisDrawing.Utility.GetPoint(, "Place " & cbType.Value & ": ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    Set objBldg = ThisDrawing.ModelSpace.InsertBlock(returnPnt, "Customer", dScale, dScale, dScale, 0#)
    objBldg.Layer = "Customers"
    
    vAttList = objBldg.GetAttributes
    
    
    Select Case cbType.Value
        Case "RES"
            vAttList(0).TextString = "RESIDENCE"
            vAttList(5).TextString = ""
        Case "BUS"
            vAttList(0).TextString = "BUSINESS"
            vAttList(5).TextString = "B"
        Case "TRLR"
            vAttList(0).TextString = "TRAILER"
            vAttList(5).TextString = "T"
        Case "MDU"
            vAttList(0).TextString = "MDU"
            vAttList(5).TextString = "M"
        Case "CHURCH"
            vAttList(0).TextString = "CHURCH"
            vAttList(5).TextString = "C"
        Case "SCHOOL"
            vAttList(0).TextString = "SCHOOL"
            vAttList(5).TextString = "S"
        Case "EXT"
            vAttList(0).TextString = "EXTENSION"
            vAttList(5).TextString = "X"
        Case Else
    End Select
    
    vAttList(1).TextString = tbHNum.Value
    vAttList(2).TextString = tbStreet.Value
    If Not tbDescription.Value = "" Then vAttList(3).TextString = tbDescription.Value
    
    objBldg.Update
    
    If Not tbPlus.Value = "" Then tbHNum.Value = CDbl(tbHNum.Value) + tbPlus.Value
        
    GoTo Place_Block
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRename_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim returnPnt, vAttList As Variant
    'Dim dScale As Double
    'Dim strBlock As String
    
    Me.Hide
  On Error Resume Next
    
    dScale = CDbl(cbScale.Value) / 100
    If Not Err = 0 Then
        dScale = 1#
        Err = 0
    End If
    
Place_Block:
    Err = 0
    ThisDrawing.Utility.GetEntity objEntity, returnPnt, vbCr & "Select Existing Building: "
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    Select Case objBlock.Name
        Case "Customer"
            vAttList = objBlock.GetAttributes
        Case Else
            GoTo Exit_Sub
    End Select
    
    
    If Not tbHNum.Value = "" Then
        vAttList(1).TextString = tbHNum.Value
        If Not tbPlus.Value = "" Then tbHNum.Value = CDbl(tbHNum.Value) + CDbl(tbPlus.Value)
    End If
    
    If Not tbStreet.Value = "" Then vAttList(2).TextString = tbStreet.Value
    If Not tbDescription.Value = "" Then vAttList(3).TextString = tbDescription.Value
    
    Select Case cbType.Value
        Case "RES"
            vAttList(0).TextString = "RESIDENCE"
            vAttList(5).TextString = ""
        Case "BUS"
            vAttList(0).TextString = "BUSINESS"
            vAttList(5).TextString = "B"
        Case "TRLR"
            vAttList(0).TextString = "TRAILER"
            vAttList(5).TextString = "T"
        Case "MDU"
            vAttList(0).TextString = "MDU"
            vAttList(5).TextString = "M"
        Case "CHURCH"
            vAttList(0).TextString = "CHURCH"
            vAttList(5).TextString = "C"
        Case "SCHOOL"
            vAttList(0).TextString = "SCHOOL"
            vAttList(5).TextString = "S"
        Case "EXT"
            vAttList(0).TextString = "EXTENSION"
            vAttList(5).TextString = "X"
        Case Else
    End Select
    
    objBlock.Update
        
    GoTo Place_Block
    
Exit_Sub:
    Me.show
End Sub

Private Sub Label1_Click()
    Dim objRoadname As AcadObject
    Dim obrGP As AcadBlockReference
    Dim basePnt, vList As Variant
    Dim objText As AcadText
    Dim objMText As AcadMText
    
    Me.Hide
On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objRoadname, basePnt, "Select Address: "
    'MsgBox objRoadname.ObjectName
    
    Select Case objRoadname.ObjectName
        Case "AcDbText"
            Set objText = objRoadname
            tbHNum.Text = objText.TextString
        Case "AcDbMText"
            Set objMText = objRoadname
            tbHNum.Text = objMText.TextString
        'Case "AcDbBlockReference"
            'Set obrGP = objRoadname
            'vList = obrGP.GetAttributes
            'tbHNum.Value = vList(0).TextString
    End Select
    
    cbAddWL.SetFocus
    Me.show
End Sub

Private Sub Label2_Click()
    Dim objRoadname As AcadObject
    'Dim obrGP As AcadBlockReference
    Dim basePnt As Variant
    Dim objText As AcadText
    Dim objMText As AcadMText
    
    Me.Hide
On Error Resume Next
    
    ThisDrawing.Utility.GetEntity objRoadname, basePnt, "Select Road Name: "
    'MsgBox objRoadname.ObjectName
    
    Select Case objRoadname.ObjectName
        Case "AcDbText"
            Set objText = objRoadname
            tbStreet.Text = objText.TextString
        Case "AcDbMText"
            Set objMText = objRoadname
            tbStreet.Text = objMText.TextString
    End Select
    Me.show
End Sub

Private Sub UserForm_Initialize()
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    cbType.AddItem "RES"
    cbType.AddItem "BUS"
    cbType.AddItem "TRLR"
    cbType.AddItem "MDU"
    cbType.AddItem "CHURCH"
    cbType.AddItem "SCHOOL"
    cbType.AddItem "EXT"
    cbType.Value = "RES"
End Sub
