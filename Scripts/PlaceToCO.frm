VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlaceToCO 
   Caption         =   "Place Block"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2160
   OleObjectBlob   =   "PlaceToCO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlaceToCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbPlaceBlock_Click()
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strName As String
    Dim vBasePnt, vReturnPnt As Variant
    Dim dInsert(2) As Double
    Dim dDiffX, dDiffY, dZ As Double
    Dim dRotate, dRatio As Double
    
    On Error Resume Next
    Me.Hide
    
Add_Callout:
    Err = 0
    
    vBasePnt = ThisDrawing.Utility.GetPoint(, "Select First Point: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    vReturnPnt = ThisDrawing.Utility.GetPoint(vBasePnt, "Select First Point: ")
    If Not Err = 0 Then GoTo Exit_Sub
    
    dDiffX = vReturnPnt(0) - vBasePnt(0)
    dDiffY = vReturnPnt(1) - vBasePnt(1)
    dZ = Sqr((dDiffX * dDiffX) + (dDiffY * dDiffY))
    dRatio = 40.8076 / dZ
    
    vReturnPnt(0) = vBasePnt(0) + (dDiffX * dRatio)
    vReturnPnt(1) = vBasePnt(1) + (dDiffY * dRatio)
    
    
    Select Case dDiffX
        Case Is > 0
            strName = "ToLOCr"
            dInsert(0) = vBasePnt(0)
            dInsert(1) = vBasePnt(1)
            dInsert(2) = 0#
            
            dRotate = Atn(dDiffY / dDiffX)
        Case Is = 0
            Select Case dDiffY
                Case Is > 0
                    strName = "ToLOCr"
                    dInsert(0) = vBasePnt(0)
                    dInsert(1) = vBasePnt(1)
                    dInsert(2) = 0#
            
                    dRotate = 1.5707963267949
                Case Is = 0
                    strName = "ToLOCl"
                    dInsert(0) = vBasePnt(0)
                    dInsert(1) = vBasePnt(1)
                    dInsert(2) = 0#
            
                    dRotate = 0
                Case Is < 0
                    strName = "ToLOCl"
                    dInsert(0) = vBasePnt(0)
                    dInsert(1) = vBasePnt(1)
                    dInsert(2) = 0#
            
                    dRotate = 1.5707963267949
            End Select
        Case Is < 0
            strName = "ToLOCl"
            dInsert(0) = vBasePnt(0)
            dInsert(1) = vBasePnt(1)
            dInsert(2) = 0#
            
            dDiffX = 0 - dDiffX
            dDiffY = 0 - dDiffY
            dRotate = Atn(dDiffY / dDiffX) - 3.14159265359
    End Select
    
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dInsert, strName, 1#, 1#, 1#, dRotate)
    vAttList = objBlock.GetAttributes
    vAttList(0).TextString = "TO " & cbTo.Value
    vAttList(1).TextString = tbName.Value
    
    objBlock.Layer = "Integrity Sheets"
    objBlock.Update
    
    GoTo Add_Callout
    
Exit_Sub:
    Me.show
End Sub

Private Sub UserForm_Initialize()
    cbTo.AddItem "C.O."
    cbTo.AddItem "RST"
    cbTo.AddItem "FDH"
    cbTo.Value = "C.O."
End Sub
