VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectAttach 
   Caption         =   "Select Attachment"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3585
   OleObjectBlob   =   "SelectAttach.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectAttach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbAttach_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub

Private Sub lbAttach_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Me.Hide
    End Select
End Sub

Private Sub UserForm_Initialize()
    lbAttach.ColumnCount = 3
    lbAttach.ColumnWidths = "72;48;48"
End Sub
