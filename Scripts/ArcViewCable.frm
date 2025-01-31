VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArcViewCable 
   Caption         =   "Cable Form"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3225
   OleObjectBlob   =   "ArcViewCable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArcViewCable"
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
    Dim strCounts As String
    
    Me.Hide
    
    strCounts = Replace(tbCounts.Value, vbCr, " + ")
    strCounts = Replace(strCounts, vbLf, "")
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
    Set tbls = amap.Projects(ThisDrawing).ODTables
    If tbls.count > 0 Then
        For Each tbl In tbls
            If tbl.Name = "Cables" Then GoTo Exit_For
        Next
    End If
    
    MsgBox "Object Data Table not found."
    GoTo Exit_Sub
    
Exit_For:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Cable: "
    
    If TypeOf objEntity Is AcadLine Then
        Set objLine = objEntity
        iLength = CInt(objLine.Length + 0.5)
    ElseIf TypeOf objEntity Is AcadLWPolyline Then
        Set objLWP = objEntity
        iLength = CInt(objLWP.Length + 0.5)
    Else
        MsgBox "Wrong type of Entity."
        GoTo Exit_Sub
    End If
    
    Set ODRcs = tbl.GetODRecords
            
    boolVal = ODRcs.Init(objEntity, True, False)
    Set ODRc = ODRcs.Record
    
    ODRc.Item(0).Value = tbCable.Value
    ODRc.Item(1).Value = iLength
    ODRc.Item(2).Value = strCounts
    
    boolVal = ODRcs.Update(ODRc)
    
    'GoTo Exit_For
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbAddLines_Click()
    Dim objObject As AcadObject
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim basePnt As Variant
    Dim X(0 To 2), Y(0 To 2), Z(0 To 2) As Double
    Dim leaderPnt(0 To 3) As Double
    Dim counter, iTemp As Integer
    Dim lineObj As AcadLWPolyline
    Dim layerObj As AcadLayer
    Dim strCommand As String
    
    Me.Hide
    ThisDrawing.SetVariable "CMDDIA", 0
    
  On Error Resume Next
    Set layerObj = ThisDrawing.Layers.Add("Cables - Aerial")
    ThisDrawing.ActiveLayer = layerObj
    
    ThisDrawing.Utility.GetEntity objObject, basePnt, "From Pole: "
    If TypeOf objObject Is AcadBlockReference Then
        Set objBlock = objObject
    Else
        GoTo Exit_While
    End If
    X(1) = objBlock.InsertionPoint(0)
    Y(1) = objBlock.InsertionPoint(1)
    Z(1) = 0#
    
    ThisDrawing.Utility.GetEntity objObject, basePnt, "To Pole: "
    If TypeOf objObject Is AcadBlockReference Then
        Set objBlock = objObject
    Else
        GoTo Exit_While
    End If
    If Not Err = 0 Then GoTo Exit_While
    
    X(0) = objBlock.InsertionPoint(0)
    Y(0) = objBlock.InsertionPoint(1)
    Z(0) = 0#
    X(2) = X(0)
    Y(2) = Y(0)
    
    leaderPnt(0) = X(1)
    leaderPnt(1) = Y(1)
    leaderPnt(2) = X(0)
    leaderPnt(3) = Y(0)
    
    Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(leaderPnt)
    lineObj.Update
    
    strCommand = "_adeattachdata" & vbCr & "Cables" & vbCr & "a" & vbCr & "n" & vbCr & "l" & vbCr & vbCr & vbCr
    ThisDrawing.SendCommand strCommand
    
    
    
    
    
    
    
    Select Case objBlock.Name
        Case "sPED", "sHH"
            vAttList = objBlock.GetAttributes
            
            If vAttList(7).TextString = "" Then
                vAttList(7).TextString = "+" & tbCable.Value & "="
            Else
                vAttList(7).TextString = vAttList(7).TextString & ";;" & "+" & tbCable.Value & "="
            End If
        Case "sPole"
            vAttList = objBlock.GetAttributes
            
            If vAttList(27).TextString = "" Then
                vAttList(27).TextString = "+" & tbCable.Value & "="
            Else
                vAttList(27).TextString = vAttList(27).TextString & ";;" & "+" & tbCable.Value & "="
            End If
    End Select
    
    
    
    
    
    
    
    
    Err = 0
    
    While Err = 0
        X(1) = X(0)
        Y(1) = Y(0)
        
        Err = 0
        
        ThisDrawing.Utility.GetEntity objObject, basePnt, "To Pole: "
        
        If Not Err = 0 Then GoTo Exit_While
        
        If TypeOf objObject Is AcadBlockReference Then
            Set objBlock = objObject
        Else
            GoTo Exit_While
        End If
        
        X(0) = objBlock.InsertionPoint(0)
        Y(0) = objBlock.InsertionPoint(1)
    
        leaderPnt(0) = X(1)
        leaderPnt(1) = Y(1)
        leaderPnt(2) = X(0)
        leaderPnt(3) = Y(0)
    
        Set lineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(leaderPnt)
        lineObj.Update
        
        ThisDrawing.SendCommand strCommand
    Wend
    
Exit_While:
    ThisDrawing.SetVariable "CMDDIA", 1
    Me.show
End Sub

Private Sub cbCancel_Click()
    tbCounts.Value = ""
    
    Me.Hide
End Sub

Private Sub cbGet_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    Dim vLine As Variant
    
    On Error Resume Next
    
    Me.Hide
        
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Cable Callout: "
        
    If Not Err = 0 Then GoTo Exit_Sub
    
    If TypeOf objEntity Is AcadBlockReference Then
        Set objBlock = objEntity
    Else
        GoTo Exit_Sub
    End If
    
    If Not objBlock.Name = "CableCounts" Then GoTo Exit_Sub
    
    vAttList = objBlock.GetAttributes
    
    tbCable.Value = vAttList(1).TextString
    tbCounts.Value = Replace(vAttList(0).TextString, "\P", vbCr)
Exit_Sub:
    Me.show
End Sub

Private Sub cbUpdate_Click()
    Me.Hide
End Sub

