VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PoleOwnerNote 
   Caption         =   "Pole Owners Note"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2880
   OleObjectBlob   =   "PoleOwnerNote.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PoleOwnerNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vPnt1, vPnt2 As Variant

Private Sub cbGetPoles_Click()
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    'Dim vPnt1, vPnt2 As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine, vItem, vTemp As Variant

    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
  On Error Resume Next
  
    Me.Hide
    
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
  
    Err = 0
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    If Not Err = 0 Then
        MsgBox "Error: " & Err.Number & vbCr & Err.Description
        Me.show
        Exit Sub
    End If
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(0).TextString = "" Then GoTo Next_objBlock
        If vAttList(0).TextString = "POLE" Then GoTo Next_objBlock
        If vAttList(1).TextString = "INCOMPLETE" Then GoTo Next_objBlock
        
        If lbOwner.ListCount > 0 Then
            For i = 0 To lbOwner.ListCount - 1
                If lbOwner.List(i, 0) = vAttList(2).TextString Then
                    lbOwner.List(i, 1) = CInt(lbOwner.List(i, 1)) + 1
                    GoTo Next_objBlock
                End If
            Next i
        End If
        
        lbOwner.AddItem vAttList(2).TextString
        lbOwner.List(lbOwner.ListCount - 1, 1) = "1"
        
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    Me.show
End Sub

Private Sub cbPlaceNotes_Click()
    If lbOwner.ListIndex < 0 Then Exit Sub
    
    Dim objLayer As AcadLayer
    Dim objSSPole As AcadSelectionSet
    Dim objSSInfo As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objPole As AcadBlockReference
    Dim objInfo As AcadBlockReference
    Dim objMText As AcadMText
    Dim objLWP As AcadLWPolyline
    Dim vAttList As Variant
    Dim vAttPole As Variant
    Dim vAttInfo As Variant
    Dim strOwner, strNote As String
    Dim vCoords As Variant
    Dim dCoords(0 To 2) As Double
    Dim dBorder(0 To 9) As Double
    Dim dWidth, dScale As Double
    Dim vMinExt, vMaxExt As Variant
    
  On Error Resume Next
    Set objLayer = ThisDrawing.Layers.Add("Integrity Pole-Owner")
    If Err = 0 Then
        objLayer.color = acRed
    Else
        Err = 0
    End If
    
    Set objSSPole = ThisDrawing.SelectionSets.Add("objSSpole")
    If Not Err = 0 Then
        Set objSSPole = ThisDrawing.SelectionSets.Item("objSSpole")
        Err = 0
    End If

    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    objSSPole.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    If Not Err = 0 Then
        MsgBox "Error: " & Err.Number & vbCr & Err.Description
        Me.show
        Exit Sub
    End If
    
    Set objSSInfo = ThisDrawing.SelectionSets.Add("objSSinfo")
    If Not Err = 0 Then
        Set objSSInfo = ThisDrawing.SelectionSets.Item("objSSinfo")
        Err = 0
    End If

    grpCode(0) = 2
    grpValue(0) = "pole_info"
    filterType = grpCode
    filterValue = grpValue
    
    objSSInfo.Select acSelectionSetWindow, vPnt1, vPnt2, filterType, filterValue
    If Not Err = 0 Then
        MsgBox "Error: " & Err.Number & vbCr & Err.Description
        Me.show
        Exit Sub
    End If
    
    strOwner = lbOwner.List(lbOwner.ListIndex)
    
    'MsgBox "Poles: " & objSSPole.count & vbCr & "Info: " & objSSInfo.count
    
    For Each objInfo In objSSInfo
        vAttInfo = objInfo.GetAttributes
        
        If Not vAttInfo(1).TextString = "0.1" Then GoTo Next_objInfo
        If Not vAttInfo(2).TextString = strOwner Then GoTo Next_objInfo
        
        'MsgBox vAttInfo(1).TextString & vbCr & vAttInfo(2).TextString
        
        dScale = objInfo.XScaleFactor
        vCoords = objInfo.InsertionPoint
        dCoords(0) = vCoords(0) + (90 * dScale)
        dCoords(1) = vCoords(1)
        dCoords(2) = 0#
        
        If cbFormat.Value = "POLE" Then
            strNote = strOwner & " POLE"
            GoTo Place_Note
        End If
        
        For Each objPole In objSSPole
            vAttPole = objPole.GetAttributes
            
            If vAttPole(0).TextString = vAttInfo(0).TextString Then
                If Not vAttPole(15).TextString = "" Then
                    strNote = strOwner & " NEW"
                    GoTo Place_Note
                End If
                
                If cbFormat.Value = "NEW / LASH" Then
                    strNote = strOwner & " LASH"
                    GoTo Place_Note
                End If
            End If
        Next objPole
        
        GoTo Next_objInfo
        
Place_Note:
        
        Set objMText = ThisDrawing.ModelSpace.AddMText(dCoords, 0, strNote)
        objMText.Layer = "Integrity Pole-Owner"
        objMText.Height = (6 * dScale)
        objMText.AttachmentPoint = acAttachmentPointBottomLeft
        objMText.InsertionPoint = dCoords
        objMText.Rotation = 0#
        'objMText.BackgroundFill = True
        objMText.Update
        
        If InStr(strNote, "LASH") > 0 Then GoTo Next_objInfo
        
        objMText.GetBoundingBox vMinExt, vMaxExt
        dWidth = vMaxExt(0) - vMinExt(0)
        
        dBorder(0) = dCoords(0) - (2 * dScale)
        dBorder(1) = dCoords(1) - (2 * dScale)
        dBorder(2) = dBorder(0) + dWidth + (4 * dScale)
        dBorder(3) = dBorder(1)
        dBorder(4) = dBorder(2)
        dBorder(5) = dBorder(1) + (10 * dScale)
        dBorder(6) = dBorder(0)
        dBorder(7) = dBorder(5)
        dBorder(8) = dBorder(0)
        dBorder(9) = dBorder(1)
        
    
        Set objLWP = ThisDrawing.ModelSpace.AddLightWeightPolyline(dBorder)
        objLWP.Layer = "Integrity Pole-Owner"
        objLWP.Update
        
Next_objInfo:
    Next objInfo
    
    
Exit_Sub:
    objSSPole.Clear
    objSSPole.Delete
    objSSInfo.Clear
    objSSInfo.Delete
    'Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    lbOwner.ColumnCount = 2
    lbOwner.ColumnWidths = "96;30"
    
    cbFormat.AddItem "POLE"
    cbFormat.AddItem "NEW"
    cbFormat.AddItem "NEW / LASH"
    cbFormat.Value = "NEW"
    
    tbFormat.Value = "<OWNER> NEW"
End Sub
