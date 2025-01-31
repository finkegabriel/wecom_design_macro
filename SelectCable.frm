VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectCable 
   Caption         =   "Select Active Cable"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   OleObjectBlob   =   "SelectCable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectCable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbNewCable_Click()
    lbCable.Clear
    
    Me.Hide
End Sub

Private Sub lbCable_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
