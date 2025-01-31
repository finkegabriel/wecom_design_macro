Attribute VB_Name = "Module3"
Public Sub ChangeODataLayers()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    
    Dim objEntity As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objLine As AcadLine
    Dim strFeeder, strLayer As String
    Dim iCount As Integer
    
    Dim amap As AcadMap
    Dim ODRcs As ODRecords
    Dim ODRc As ODRecord
    Dim tbl As ODTable
    Dim tbls As ODTables
    Dim boolVal As Boolean
    
    On Error Resume Next
    
    grpCode(0) = 8
    grpValue(0) = "OH_Primary"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
        Err = 0
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    MsgBox "Objects found:  " & objSS.count
        
    Err = 0
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
    Set tbls = amap.Projects(ThisDrawing).ODTables
    If tbls.count > 0 Then
        For Each tbl In tbls
            If tbl.Name = "OH_Primary" Then GoTo Exit_For
        Next
    End If
    
Exit_For:
    
    iCount = 0
    For Each objEntity In objSS
        iCount = iCount + 1
        If iCount > 10 Then GoTo Exit_Sub
        
        Set ODRcs = tbl.GetODRecords
        boolVal = ODRcs.Init(objEntity, True, False)
        Set ODRc = ODRcs.Record
        
        MsgBox Err.Description
        MsgBox iCount & vbCr & ODRc.Item(2).Value
        
        If Not Err = 0 Then GoTo Next_objEntity
        
        strFeeder = ODRc.Item(7).Value
        If strFeeder = "" Then GoTo Next_objEntity
        
        Select Case CInt(strFeeder)
            Case Is < 200
                objEntity.Layer = "OH_Primary - 100"
            Case Is < 300
                If CInt(strFeeder) > 200 Then strLayer = "OH_Primary - 200"
            Case Is < 400
                If CInt(strFeeder) > 300 Then strLayer = "OH_Primary - 300"
            Case Is < 500
                If CInt(strFeeder) > 400 Then strLayer = "OH_Primary - 400"
            Case Is < 600
                If CInt(strFeeder) > 500 Then strLayer = "OH_Primary - 500"
            Case Is < 700
                If CInt(strFeeder) > 600 Then strLayer = "OH_Primary - 600"
            Case Is < 800
                If CInt(strFeeder) > 700 Then strLayer = "OH_Primary - 700"
            Case Is < 900
                If CInt(strFeeder) > 800 Then strLayer = "OH_Primary - 800"
            Case Is < 1000
                If CInt(strFeeder) > 900 Then strLayer = "OH_Primary - 900"
            Case Else
                If CInt(strFeeder) > 1000 Then strLayer = "OH_Primary - 1000"
        End Select
        
        If TypeOf objEntity Is AcadLine Then
            Set objLine = objEntity
            objLine.Layer = strLayer
            
            objLine.Update
        Else
            Set objLWP = objEntity
            objLWP.Layer = strLayer
            
            objLWP.Update
        End If
        
Next_objEntity:
        'Set ODRc = Nothing
        Err = 0
    Next objEntity
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub
