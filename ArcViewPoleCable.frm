VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArcViewPoleCable 
   Caption         =   "Cable Form"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3225
   OleObjectBlob   =   "ArcViewPoleCable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArcViewPoleCable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbAddData_Click()
    Dim amap As AcadMap
    Dim tbl As ODTable
    Dim tbls As ODTables
    Dim ODRcs As ODRecords
    Dim ODRc As ODRecord
    Dim boolVal As Boolean
    
    Dim objEntity As AcadEntity
    Dim vReturnPnt As Variant
    Dim objLine As AcadLine
    Dim objLWP As AcadLWPolyline
    Dim iLength As Integer
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
    Set tbls = amap.Projects(ThisDrawing).ODTables
    If tbls.count > 0 Then
        For Each tbl In tbls
            If tbl.Name = "Cables" Then GoTo Exit_For
        Next
    End If
    
    MsgBox "Object Data Table not found."
    Exit Sub
    
Exit_For:
    Set ODRcs = tbl.GetODRecords
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Cable: "
    
    If TypeOf objEntity Is AcadLine Then
        Set objLine = objEntity
        iLength = objLine.Length
    ElseIf TypeOf objEntity Is AcadLWPolyline Then
        Set objLWP = objEntity
        iLength = objLWP.Length
    Else
        MsgBox "Wrong type of Entity."
        Me.show
        Exit Sub
    End If
    
    Set ODRc = tbl.CreateRecord
    
    ODRc.Item(0).Value = tbCable.Value
    ODRc.Item(1).Value = iLength
    ODRc.Item(2).Value = tbCounts.Value
    
    MsgBox objEntity.ObjectID
    'ODRc.AttachTo objEntity.ObjectID
    If TypeOf objEntity Is AcadLine Then
        MsgBox "Line ID  " & objLine.ObjectID
        'ODRc.AttachTo (objLine.ObjectID)
    Else
        MsgBox "LWP ID  " & objLWP.ObjectID
        'ODRc.AttachTo (objLWP.ObjectID)
    End If
    
    Me.show
End Sub

Private Sub cbCancel_Click()
    tbCounts.Value = ""
    
    Me.Hide
End Sub

Private Sub cbUpdate_Click()
    Me.Hide
End Sub
