VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadBlocks 
   Caption         =   "Place Load Blocks"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   OleObjectBlob   =   "LoadBlocks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbClear_Click()
    tbRes.Value = "0"
    tbSG.Value = "0"
    tbBus.Value = "0"
    tbSGNeed.Value = ""
    tbRESNeed.Value = ""
    tbBUSNeed.Value = ""
    tbFibers.Value = ""
    tbSize.Value = ""
End Sub

Private Sub cbCustomers_Click()
    Dim basePoint(0 To 2) As Double
    Dim scaleFactor As Double
    Dim ObjInSS As AcadObject
    Dim obrGP As AcadBlockReference
    Dim vAttItem, basePnt As Variant
    Dim ssObj As AcadSelectionSet
    Dim iSG As Integer
    Dim iRes, iBus, iSize As Integer
    Dim dNeeded As Double
    
    Me.Hide
    On Error Resume Next
    
    iSG = CInt(tbSG.Value)
    iRes = CInt(tbRes.Value)
    iBus = CInt(tbBus.Value)
    
    Set ssObj = ThisDrawing.SelectionSets.Add("LoadBlocks")
    If Not Err = 0 Then
        Set ssObj = ThisDrawing.SelectionSets.Item("LoadBlocks")
        Err = 0
    End If
    ssObj.SelectOnScreen
    
    For Each ObjInSS In ssObj
        If TypeOf ObjInSS Is AcadBlockReference Then
            Set obrGP = ObjInSS
            Select Case obrGP.Name
                Case "RES", "LOT", "TRLR", "MDU", "CHURCH", "SCHOOL"
                    iRes = iRes + 1
                Case "BUSINESS", "SCHOOL"
                    iBus = iBus + 1
                Case "SG"
                    iSG = iSG + 1
                Case "Customer"
                    vAttList = obrGP.GetAttributes
                    Select Case vAttList(5).TextString
                        Case "", "R", "T", "M", "C"
                            iRes = iRes + 1
                        Case "B", "S"
                            iBus = iBus + 1
                    End Select
                'Case "testblock"    '<------------------------------------- need to get SG in existing load block
                    'attItem = obrGP.GetAttributes
                    'iRes = iRes + cint(attItem(0).TextString)
                    'iBus = iBus + cint(attItem(1).TextString)
            End Select
        End If
        Set obrGP = Nothing
    Next ObjInSS
    
    tbSG.Value = iSG
    tbRes.Value = iRes
    tbBus.Value = iBus
    tbAll.Value = iSG + iRes + iBus
    tbSGNeed.Value = CInt(iSG * CDbl(TextBox3.Value) + 0.5) 'iSG
    tbRESNeed.Value = CInt(iRes * CDbl(tbGFRES.Value) + 0.5)
    tbBUSNeed.Value = CInt(iBus * CDbl(tbGFBUS.Value) + 0.5)
    tbFibers.Value = CInt(tbSGNeed.Value) + CInt(tbRESNeed.Value) + CInt(tbBUSNeed.Value)
    
    Select Case CInt(tbFibers.Value)
        Case Is < 25
            tbSize.Value = "24"
        Case Is < 37
            tbSize.Value = "36"
        Case Is < 49
            tbSize.Value = "48"
        Case Is < 73
            tbSize.Value = "72"
        Case Is < 97
            tbSize.Value = "96"
        Case Is < 145
            tbSize.Value = "144"
        Case Is < 217
            tbSize.Value = "216"
        Case Is < 289
            tbSize.Value = "288"
        Case Is < 361
            tbSize.Value = "360"
        Case Is < 433
            tbSize.Value = "432"
    End Select
    
    ssObj.Clear
    ssObj.Delete
    'ThisDrawing.SelectionSets.Item("LoadBlocks").Delete
    Me.show
End Sub

Private Sub cbExport_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim vCoords, vArray As Variant
    Dim dCoords() As Double
    Dim strFileName As String
    Dim vTemp As Variant
    
    'Dim iRES, iBUS, iSG, iTemp, iCounter As Integer
    
    iRes = 0: iBus = 0: iSG = 0
    
    On Error Resume Next
    
    vTemp = Split(ThisDrawing.Name, " ")
    strFileName = ThisDrawing.Path & "\" & vTemp(0) & " Homes Passed.csv"
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Trim Border: "
    If Not objEntity.ObjectName = "AcDbPolyline" Then
        MsgBox "Error: Invalid Selection."
        Me.show
        Exit Sub
    End If
    
    Set objLWP = objEntity
    vCoords = objLWP.Coordinates
    
    iTemp = (UBound(vCoords) + 1) / 2 * 3 - 1
    ReDim dCoords(iTemp) As Double
    
    iCounter = 0
    For i = 0 To UBound(vCoords) Step 2
        dCoords(iCounter) = vCoords(i)
        iCounter = iCounter + 1
        dCoords(iCounter) = vCoords(i + 1)
        iCounter = iCounter + 1
        dCoords(iCounter) = 0#
        iCounter = iCounter + 1
    Next i
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    
    objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords
    
    If objSS.count < 1 Then GoTo Exit_Sub
    
    Open strFileName For Output As #1
    Print #1, "Type, House #,Street Name,Note"
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadBlockReference Then
            Set objBlock = objEntity
            Select Case objBlock.Name
                Case "RES", "LOT", "TRLR", "MDU", "CHURCH", "SCHOOL", "BUSINESS", "SCHOOL"
                    vTemp = objBlock.GetAttributes
                    
                    Print #1, objBlock.Name & "," & vTemp(0).TextString & "," & vTemp(1).TextString & "," & vTemp(2).TextString
                Case "Customer"
                    vTemp = objBlock.GetAttributes
                    
                    Print #1, vTemp(0).TextString & "," & vTemp(1).TextString & "," & vTemp(2).TextString & "," & vTemp(3).TextString
            End Select
        Set objBlock = Nothing
        End If
    Next objEntity
    
    Close #1
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbGetLB_Click()
    Dim basePoint(0 To 2) As Double
    Dim scaleFactor As Double
    Dim ObjInSS As AcadObject
    Dim obrGP As AcadBlockReference
    Dim attItem, basePnt As Variant
    Dim ssObj As AcadSelectionSet
    Dim iSG As Integer
    Dim iRes, iBus, iSize As Integer
    Dim dNeeded As Double
    
    Me.Hide
    On Error Resume Next
    
    iSG = CInt(tbSG.Value)
    iRes = CInt(tbRes.Value)
    iBus = CInt(tbBus.Value)
    
    Set ssObj = ThisDrawing.SelectionSets.Add("LoadBlocks")
    ssObj.SelectOnScreen
    
    For Each ObjInSS In ssObj
        If TypeOf ObjInSS Is AcadBlockReference Then
            Set obrGP = ObjInSS
            
            If obrGP.Name = "loadblock2" Then
                attItem = obrGP.GetAttributes
                iSG = iSG + CInt(attItem(0).TextString)
                iRes = iRes + CInt(attItem(1).TextString)
                iBus = iBus + CInt(attItem(2).TextString)
            End If
        End If
        Set obrGP = Nothing
    Next ObjInSS
    
    tbSG.Value = iSG
    tbRes.Value = iRes
    tbBus.Value = iBus
    tbAll.Value = iSG + iRes + iBus
    tbSGNeed.Value = CInt(iSG * CDbl(TextBox3.Value) + 0.5) 'iSG
    tbRESNeed.Value = CInt(iRes * CDbl(tbGFRES.Value) + 0.5)
    tbBUSNeed.Value = CInt(iBus * CDbl(tbGFBUS.Value) + 0.5)
    tbFibers.Value = CInt(tbSGNeed.Value) + CInt(tbRESNeed.Value) + CInt(tbBUSNeed.Value)
    
    Select Case CInt(tbFibers.Value)
        Case Is < 25
            tbSize.Value = "24"
        Case Is < 37
            tbSize.Value = "36"
        Case Is < 49
            tbSize.Value = "48"
        Case Is < 73
            tbSize.Value = "72"
        Case Is < 97
            tbSize.Value = "96"
        Case Is < 145
            tbSize.Value = "144"
        Case Is < 217
            tbSize.Value = "216"
        Case Is < 289
            tbSize.Value = "288"
        Case Is < 361
            tbSize.Value = "360"
        Case Is < 433
            tbSize.Value = "432"
    End Select
    
    ThisDrawing.SelectionSets.Item("LoadBlocks").Delete
    Me.show
End Sub

Private Sub cbPlace_Click()
    Dim objBlock As AcadBlockReference
    Dim returnPnt, attItem As Variant
    Dim dScale As Double
    Dim iRes As Integer
    
    dScale = 2#
    
    Me.Hide
    
    returnPnt = ThisDrawing.Utility.GetPoint(, "Select Point: ")
    
    Set objBlock = ThisDrawing.ModelSpace.InsertBlock(returnPnt, "loadblock2", dScale, dScale, dScale, 0)
    
    'iRes = CInt(tbRES.Value) + CInt(tbSG.Value)
    
    attItem = objBlock.GetAttributes
    attItem(0).TextString = tbSG.Value
    attItem(1).TextString = tbRes.Value
    attItem(2).TextString = tbBus.Value
    attItem(3).TextString = tbFibers.Value
    attItem(4).TextString = tbSize.Value
    objBlock.Update
    
    tbRes.Value = ""
    tbBus.Value = ""
    tbRESNeed.Value = ""
    tbBUSNeed.Value = ""
    tbFibers.Value = ""
    tbSize.Value = ""
    
    Me.show
End Sub

Private Sub cbPolygon_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objLWP As AcadLWPolyline
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt As Variant
    Dim vCoords, vArray As Variant
    Dim vAttList As Variant
    Dim dCoords() As Double
    Dim iRes, iBus, iSG, iTemp, iCounter As Integer
    
    iRes = 0: iBus = 0: iSG = 0
    
    On Error Resume Next
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, "Select Polygon Border: "
    If Not objEntity.ObjectName = "AcDbPolyline" Then
        MsgBox "Error: Invalid Selection."
        Me.show
        Exit Sub
    End If
    
    Set objLWP = objEntity
    vCoords = objLWP.Coordinates
    
    iTemp = (UBound(vCoords) + 1) / 2 * 3 - 1
    ReDim dCoords(iTemp) As Double
    
    iCounter = 0
    For i = 0 To UBound(vCoords) Step 2
        dCoords(iCounter) = vCoords(i)
        iCounter = iCounter + 1
        dCoords(iCounter) = vCoords(i + 1)
        iCounter = iCounter + 1
        dCoords(iCounter) = 0#
        iCounter = iCounter + 1
    Next i
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    
    objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadBlockReference Then
            Set objBlock = objEntity
            Select Case objBlock.Name
                Case "RES", "LOT", "TRLR", "MDU", "CHURCH", "SCHOOL"
                    iRes = iRes + 1
                Case "BUSINESS", "SCHOOL"
                    iBus = iBus + 1
                Case "SG"
                    iSG = iSG + 1
                Case "Customer"
                    vAttList = objBlock.GetAttributes
                    
                    Select Case vAttList(5).TextString
                        Case "", "T", "M", "C"
                            iRes = iRes + 1
                        Case "B", "S"
                            iBus = iBus + 1
                        Case "!"
                            iSG = iSG + 1
                    End Select
            End Select
        Set objBlock = Nothing
        End If
    Next objEntity
    
    objSS.Clear
    objSS.Delete
    
    tbSG.Value = CInt(tbSG.Value) + iSG
    tbRes.Value = CInt(tbRes.Value) + iRes
    tbBus.Value = CInt(tbBus.Value) + iBus
    tbAll.Value = CInt(tbSG.Value) + CInt(tbRes.Value) + CInt(tbBus.Value)
    tbSGNeed.Value = CInt(iSG * CDbl(TextBox3.Value) + 0.5) 'iSG
    tbRESNeed.Value = CInt(iRes * CDbl(tbGFRES.Value) + 0.5)
    tbBUSNeed.Value = CInt(iBus * CDbl(tbGFBUS.Value) + 0.5)
    tbFibers.Value = CInt(tbSGNeed.Value) + CInt(tbRESNeed.Value) + CInt(tbBUSNeed.Value)
    
    Select Case CInt(tbFibers.Value)
        Case Is < 25
            tbSize.Value = "24"
        Case Is < 37
            tbSize.Value = "36"
        Case Is < 49
            tbSize.Value = "48"
        Case Is < 73
            tbSize.Value = "72"
        Case Is < 97
            tbSize.Value = "96"
        Case Is < 145
            tbSize.Value = "144"
        Case Is < 217
            tbSize.Value = "216"
        Case Is < 289
            tbSize.Value = "288"
        Case Is < 361
            tbSize.Value = "360"
        Case Is < 433
            tbSize.Value = "432"
    End Select
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbSenerio_Change()
    Select Case cbSenerio.Value
        Case "UTC"
            TextBox3.Value = "1"
            tbGFRES.Value = "1.25"
            tbGFBUS.Value = "2"
        Case "Lor Rural"
            TextBox3.Value = "1"
            tbGFRES.Value = "1.5"
            tbGFBUS.Value = "1.5"
        Case "Lor Urban"
            TextBox3.Value = "1"
            tbGFRES.Value = "1.75"
            tbGFBUS.Value = "1.75"
    End Select
    
    tbSGNeed.Value = CInt(CInt(tbSG.Value) * CDbl(TextBox3.Value) + 0.5)
    tbRESNeed.Value = CInt(CInt(tbRes.Value) * CDbl(tbGFRES.Value) + 0.5)
    tbBUSNeed.Value = CInt(CInt(tbBus.Value) * CDbl(tbGFBUS.Value) + 0.5)
    tbFibers.Value = CInt(tbSGNeed.Value) + CInt(tbRESNeed.Value) + CInt(tbBUSNeed.Value)
End Sub

Private Sub UserForm_Initialize()
    cbSenerio.AddItem ""
    cbSenerio.AddItem "UTC"
    cbSenerio.AddItem "Lor Rural"
    cbSenerio.AddItem "Lor Urban"
    
    cbSenerio.Value = "UTC"
End Sub
