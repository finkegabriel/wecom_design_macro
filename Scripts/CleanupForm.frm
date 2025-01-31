VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CleanupForm 
   Caption         =   "Resize Blocks"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "CleanupForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CleanupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vPnt1, vPnt2 As Variant

Private Sub cbParcels_Click()
    Dim objSS6 As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objLWP As AcadPolyline
    Dim objLayer As AcadLayer
    Dim strLineType As String
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    
  On Error Resume Next
    
    Me.Hide
    
    strLineType = "PHANTOM"
    ThisDrawing.Linetypes.Load strLineType, "acad.lin"
    
    strLineType = "DASHED"
    ThisDrawing.Linetypes.Load strLineType, "acad.lin"
    
    Set objLayer = ThisDrawing.Layers("Parcels")
    objLayer.color = acYellow
    objLayer.Linetype = "PHANTOM"
    
    ThisDrawing.SendCommand "REGEN" & vbCr
    
    Exit Sub

    grpCode(0) = 0
    grpValue(0) = "AcadLWPolyline"
    grpCode(0) = 8
    grpValue(0) = "Parcels"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS6 = ThisDrawing.SelectionSets.Add("objSS6")
    objSS6.Select acSelectionSetAll, , , filterType, filterValue
    If Not Err = 0 Then GoTo Exit_Sub
    
    For Each objEntity In objSS6
        Set objLWP = objEntity
        objLWP.LinetypeGeneration = True 'Not objLWP.LinetypeGeneration
        objLWP.Update
    Next objEntity
    
    Err = 0
    ThisDrawing.SendCommand "-OVERKILL" & vbCr & "P" & vbCr & vbCr & vbCr
    If Not Err = 0 Then
        MsgBox "Broke"
        GoTo Exit_Sub
    End If
    
Exit_Sub:
    objSS6.Clear
    objSS6.Delete
    
    Me.show
End Sub

Private Sub cbResize_Click()
    Dim objSS6 As AcadSelectionSet
    
    Me.Hide
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    If cbPoles Then ResizeBldgType ("sPole")
    If cbBldg Then Call ResizeBldgs
    If cbGuy Then Call ResizeGuys
    If cbPEDs Then Call ResizePeds
    If cbRoads Then Call ResizeRoads
    If cbDynamic Then Call ResizeDynamicType
    
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    cbScale.AddItem ""
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "75"
End Sub

Private Sub ResizeBldgs()
    ResizeBldgType ("Customer")
    'ResizeBldgType ("RES")
    'ResizeBldgType ("TRLR")
    'ResizeBldgType ("MDU")
    'ResizeBldgType ("SCHOOL")
    'ResizeBldgType ("CHURCH")
    'ResizeBldgType ("EXTENTION")
    'ResizeBldgType ("NONRES")
End Sub

Private Sub ResizePeds()
    ResizeBldgType ("sPED")
    ResizeBldgType ("sFP")
    ResizeBldgType ("sHH")
    'ResizeDynamicType ("dHH")
End Sub

Private Sub ResizeGuys()
    ResizeBldgType ("ExAncOL")
    ResizeBldgType ("ExAncOR")
    
    ResizeBldgType ("ExGuyOL")
    ResizeBldgType ("ExGuyOR")
    
    ResizeBldgType ("ohgL")
    ResizeBldgType ("ohgR")
End Sub

Private Sub ResizeRoads()
    Dim objSS5 As AcadSelectionSet
    'Dim objPointBlock As AcadBlockReference
    Dim txtRoads As AcadMText
    Dim objEntity As AcadEntity
    Dim filterType, filterValue As Variant
    Dim grpCode(0 To 1) As Integer
    Dim grpValue(0 To 1) As Variant
    Dim dScale As Double
    
    On Error Resume Next
    
    dScale = CDbl(cbScale.Value) / 100

    grpCode(0) = 0
    grpValue(0) = "MTEXT"
    grpCode(1) = 8
    grpValue(1) = "Roads_MText"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS5 = ThisDrawing.SelectionSets.Add("objSS5")
    'objSS5.Select acSelectionSetAll, , , filterType, filterValue
    objSS5.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    If Not Err = 0 Then GoTo Exit_Sub
    
    'MsgBox strLayer & ": " & objSS5.count
        
    For Each objEntity In objSS5
        'If Not objEntity.ObjectName = "AcDbBlockReference" Then GoTo Next_objEntity
        Set txtRoads = objEntity
        
        txtRoads.Height = 8
Next_objEntity:
    Next objEntity
    
Exit_Sub:
    objSS5.Clear
    objSS5.Delete
End Sub

Private Sub ResizeBldgType(strLayer As String)
    Dim objSS5 As AcadSelectionSet
    Dim objPointBlock As AcadBlockReference
    Dim objEntity As AcadEntity
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim dScale As Double
    
    On Error Resume Next
    
    dScale = CDbl(cbScale.Value) / 100

    grpCode(0) = 2
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS5 = ThisDrawing.SelectionSets.Add("objSS5")
    'objSS5.Select acSelectionSetAll, , , filterType, filterValue
    objSS5.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    If Not Err = 0 Then GoTo Exit_Sub
    
    'MsgBox strLayer & ": " & objSS5.count
        
    For Each objEntity In objSS5
        If Not objEntity.ObjectName = "AcDbBlockReference" Then GoTo Next_objEntity
        Set objPointBlock = objEntity
        
        objPointBlock.XScaleFactor = dScale
        objPointBlock.YScaleFactor = dScale
        objPointBlock.ZScaleFactor = dScale
Next_objEntity:
    Next objEntity
    
Exit_Sub:
    objSS5.Clear
    objSS5.Delete
End Sub

Private Sub ResizeDynamicType()
    Dim objSS5 As AcadSelectionSet
    Dim objPointBlock As AcadBlockReference
    Dim objEntity As AcadEntity
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vProp As Variant
    Dim dScale As Double
    Dim strDBName As String
    Dim strDBList As String
    
    On Error Resume Next
    
    dScale = CDbl(cbScale.Value) / 100

    grpCode(0) = 0
    grpValue(0) = "INSERT"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS5 = ThisDrawing.SelectionSets.Add("objSS5")
    objSS5.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    If Not Err = 0 Then GoTo Exit_Sub
    
    For Each objEntity In objSS5
        Set objPointBlock = objEntity
        
        If objPointBlock.IsDynamicBlock Then
            strDBName = objPointBlock.Name
            MsgBox strDBName

            'If Left(strDBName, 1) = "*" Then
                strDBName = "`" + strDBName
                If strDBList = "" Then
                    strDBList = strDBName
                Else
                    If InStr(strDBList, strDBName) = 0 Then
                        strDBList = strDBList & "," & strDBName
                    End If
                End If
            'End If
        End If
    Next objEntity
    
    MsgBox strDBList
    
    objSS5.Clear
    
    Dim grpCode2(0 To 1)
    Dim grpValue2(0 To 1)
    grpCode2(0) = 0: grpValue2(0) = "Insert"
    grpCode2(1) = 2: grpValue2(1) = strDBList
    filterType = grpCode2
    filterValue = grpValue2
    
    objSS5.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
        
    For Each objEntity In objSS5
        'If Not objEntity.ObjectName = "AcDbBlockReference" Then GoTo Next_objEntity
        Set objPointBlock = objEntity
        
        objPointBlock.XScaleFactor = dScale
        objPointBlock.YScaleFactor = dScale
        objPointBlock.ZScaleFactor = dScale
Next_objEntity:
    Next objEntity
    
Exit_Sub:
    objSS5.Clear
    objSS5.Delete
End Sub
