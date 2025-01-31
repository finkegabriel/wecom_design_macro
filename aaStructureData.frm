VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} aaStructureData 
   Caption         =   "Structure Data"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   OleObjectBlob   =   "aaStructureData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "aaStructureData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iSave As Integer

Private Sub cbUpdate_Click()
    iSave = 1
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    iSave = 0
End Sub
