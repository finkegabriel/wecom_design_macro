VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReplaceF1 
   Caption         =   "Replace F1 Counts"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   OleObjectBlob   =   "ReplaceF1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReplaceF1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vPnt1, vPnt2 As Variant

Private Sub cbClearData_Click()
    Dim result As Integer
    
    result = MsgBox("Are you sure you want to Clear the Cables and Splices from the Checked Boxes?", vbYesNo, "Clear Data!!!")
    If result = vbNo Then
        Exit Sub
    End If
    
    On Error Resume Next
    Me.Hide
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, vbCr & "Get a corner:")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get opposite corner:")
    
    If cbPoles.Value = True Then Call ClearBlocks("sPole")
    If cbPoles.Value = True Then Call ClearBlocks("sPed")
    If cbPoles.Value = True Then Call ClearBlocks("sHH")
    If cbPoles.Value = True Then Call ClearBlocks("Customer")
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub ClearBlocks(strName As String)
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    grpCode(0) = 2
    grpValue(0) = strName
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    If objSS.count < 1 Then GoTo Exit_Sub
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        Select Case objBlock.Name
            Case "sPole"
                vAttList(25).TextString = ""
                vAttList(26).TextString = ""
            Case "sPed", "sHH"
                vAttList(5).TextString = ""
                vAttList(6).TextString = ""
            Case "Customer"
                vAttList(4).TextString = ""
        End Select
        
        objBlock.Update
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub

