VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CountFileView 
   Caption         =   "Count File View"
   ClientHeight    =   7819
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   OleObjectBlob   =   "CountFileView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CountFileView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    lbCounts.ColumnCount = 6
    lbCounts.ColumnWidths = "26;70;36;140;24;84"
End Sub
