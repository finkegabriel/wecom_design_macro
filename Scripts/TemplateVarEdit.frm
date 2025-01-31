VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateVarEdit 
   Caption         =   "Edit Variable"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "TemplateVarEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TemplateVarEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbUpdate_Click()
    Me.Hide
End Sub

Private Sub cbVar_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Me.Caption = "New Variable" Then Exit Sub
    If cbVar.ListCount = 0 Then Exit Sub
    
    Dim strTemp As String
    
    For i = 0 To cbVar.ListCount - 1
        If cbVar.List(i) = cbVar.Value Then
            For j = 0 To TemplateEditor.lbVar.ListCount - 1
                If TemplateEditor.lbVar.List(j, 0) = cbVar.Value Then
                    strTemp = Replace(TemplateEditor.lbVar.List(j, 1), ",", vbCr)
                    tbList.Value = strTemp
                    Exit Sub
                End If
            Next j
        End If
    Next i
End Sub

