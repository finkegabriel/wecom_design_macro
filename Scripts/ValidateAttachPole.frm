VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateAttachPole 
   Caption         =   "Fielded Pole Attachments"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3750
   OleObjectBlob   =   "ValidateAttachPole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateAttachPole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbMark_Click()
    Dim objCircle As AcadCircle
    Dim dInsert(2) As Double
    Dim vTemp As Variant

    vTemp = Split(tbCoords.Value, ",")
    dInsert(0) = CDbl(vTemp(0))
    dInsert(1) = CDbl(vTemp(1))
    dInsert(2) = 0#
                    
    Set objCircle = ThisDrawing.ModelSpace.AddCircle(dInsert, 80)
    objCircle.Layer = "Integrity Notes"
    objCircle.Update
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    Err = 0
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

