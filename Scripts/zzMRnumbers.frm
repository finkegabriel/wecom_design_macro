VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zzMRnumbers 
   Caption         =   "UserForm1"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5025
   OleObjectBlob   =   "zzMRnumbers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zzMRnumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGetPoles_Click()
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim vPnt1, vPnt2 As Variant
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim vLine, vItem, vTemp As Variant
    Dim iIndex, iTemp As Integer
    
    'Dim dN, dE As Double
    'Dim vLL As Variant
    
    iIndex = -1

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
                    GoTo Found_Owner
                End If
            Next i
        End If
        
        lbOwner.AddItem vAttList(2).TextString
        lbOwner.List(lbOwner.ListCount - 1, 1) = "1"
        
Found_Owner:
        
        For i = 16 To 23
            If vAttList(i).TextString = "" Then GoTo Next_I
            
            vLine = Split(vAttList(i).TextString, "=")
            If UBound(vLine) > 0 Then
                If vLine(1) = "" Then GoTo Next_I
                vItem = Split(vLine(1), " ")
                If lbMR.ListCount > 0 Then
                    For j = 0 To lbMR.ListCount - 1
                        If lbMR.List(j, 0) = vLine(0) Then
                            lbMR.List(j, 1) = CInt(lbMR.List(j, 1)) + 1
                                
                            iIndex = j
                            GoTo Found_MR
                        End If
                    Next j
                End If
                        
                lbMR.AddItem vLine(0)
                iIndex = lbMR.ListCount - 1
                lbMR.List(iIndex, 1) = "1"
                lbMR.List(iIndex, 2) = "0"
                lbMR.List(iIndex, 3) = "0"
                lbMR.List(iIndex, 4) = "0"
Found_MR:
                For j = 0 To UBound(vItem)
                    iTemp = CInt(lbMR.List(iIndex, 2)) + 1
                    lbMR.List(iIndex, 2) = iTemp
                    
                    'MsgBox iTemp
                            
                    If InStr(vItem(j), ")") > 0 Then
                        iTemp = CInt(lbMR.List(iIndex, 3)) + 1
                        lbMR.List(iIndex, 3) = iTemp
                        'MsgBox iTemp
                    Else
                        If InStr(UCase(vItem(j)), "X") > 0 Then
                            iTemp = CInt(lbMR.List(iIndex, 3)) + 1
                            lbMR.List(iIndex, 3) = iTemp
                            'MsgBox iTemp
                        End If
                    End If
                Next j
            End If
Next_I:
        Next i
        
Next_objBlock:
    Next objBlock
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    lbMR.ColumnCount = 5
    lbMR.ColumnWidths = "96;36;36;36;30"
    
    lbOwner.ColumnCount = 2
    lbOwner.ColumnWidths = "96;30"
End Sub
