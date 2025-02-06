VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} aaPlaceDrops 
   Caption         =   "Place Drops"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4830
   OleObjectBlob   =   "aaPlaceDrops.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "aaPlaceDrops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objStructure As AcadBlockReference

Private Sub cbGetCustomers_Click()
    Call GetSub
    Exit Sub
    
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strLine As String
    Dim iCount As Integer
    
    iCount = lbSubs.ListCount + 1
    
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
        Err = 0
    End If
    
    Me.Hide
    
    objSS.SelectOnScreen
    
    For Each objBlock In objSS
        If objBlock.Name = "Customer" Then
            vAttList = objBlock.GetAttributes
            If InStr(vAttList(4).TextString, tbStructure.Value & "-") > 0 Then GoTo Next_Customer
            
            vAttList(4).TextString = tbStructure.Value & " - " & iCount
            objBlock.Update
            
            strLine = vAttList(1).TextString & " " & vAttList(2).TextString
            lbSubs.AddItem strLine
            lbSubs.List(lbSubs.ListCount - 1, 1) = Left(vAttList(0).TextString, 1)
            
            iCount = iCount + 1
        End If
Next_Customer:
    Next objBlock
    
    tbListcount.Value = lbSubs.ListCount
    
    
    If lbSubs.ListCount > 0 Then
        Dim iR, iB As Integer
        
        iR = 0: iB = 0
        
        For i = 0 To lbSubs.ListCount - 1
            Select Case lbSubs.List(i, 1)
                Case "R", "C", "M", "T"
                    iR = iR + 1
                Case "B", "S"
                    iB = iB + 1
            End Select
        Next i
        
        Dim vData As Variant
        
        vData = Split(tbData.Value, vbCr)
        vData(5) = iR & "," & iB
        strLine = vData(0)
        For i = 1 To UBound(vData)
            strLine = strLine & vbCr & vData(i)
        Next i
        
        tbData.Value = strLine
    End If
    
    Me.show
End Sub

Private Sub cbGetStructure_Click()
    Dim objEntity As AcadEntity
    Dim vAttList As Variant
    Dim vReturnPnt As Variant
    
    On Error Resume Next
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt
    If Not Err = 0 Then
        MsgBox "Error"
        GoTo Exit_Sub
    End If
    
    If Not TypeOf objEntity Is AcadBlockReference Then
        MsgBox "Not a block"
        GoTo Exit_Sub
    End If
    
    Set objStructure = objEntity
    
    Select Case objStructure.Name
        Case "sPole"
            vAttList = objStructure.GetAttributes
            tbStructure.Value = vAttList(0).TextString
        Case Else
            MsgBox "Invalid block"
    End Select
    
Exit_Sub:
    
    lbSubs.Clear
    If Not vAttList(28).TextString = "" Then tbData.Value = Replace(vAttList(28).TextString, ";;", vbCr)
    
    Me.show
End Sub

Private Sub cbSaveStructure_Click()
    If tbData.Value = "" Then Exit Sub
    
    Dim vAttList As Variant
    Dim strLine As String
    
    strLine = Replace(tbData.Value, vbCr, ";;")
    strLine = Replace(strLine, vbLf, "")
    
    vAttList = objStructure.GetAttributes
    vAttList(28).TextString = strLine
    
    objStructure.Update
End Sub

Private Sub UserForm_Initialize()
    lbSubs.ColumnCount = 2
    lbSubs.ColumnWidths = "168;54"
End Sub

Private Sub GetSub()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    Dim vAttList, vReturnPnt As Variant
    Dim strResult As String
    Dim iCount, iAerial As Integer
    
    On Error Resume Next
    Me.Hide
    
Get_Customer:
    iCount = lbSubs.ListCount + 1
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Customer:"
    If Not Err = 0 Then GoTo Exit_Sub
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    If Not objBlock.Name = "Customer" Then GoTo Exit_Sub
    
    vAttList = objBlock.GetAttributes
    vAttList(4).TextString = tbStructure.Value & " - " & iCount
    objBlock.Update
    
    lbSubs.AddItem vAttList(1).TextString & " " & vAttList(2).TextString
    lbSubs.List(iCount - 1, 1) = Left(vAttList(0).TextString, 1)
    
    If cbDrops.Value = False Then GoTo Get_Customer
    
    strResult = ThisDrawing.Utility.GetString(False, vbCr & "Aerial/Buried segment:")
    Select Case Left(UCase(strResult), 1)
        Case "B"
        Case Else
    End Select
    
    Err = 0
    GoTo Get_Customer
    
Exit_Sub:
    
    If lbSubs.ListCount > 0 Then
        Dim iR, iB As Integer
        
        iR = 0: iB = 0
        
        For i = 0 To lbSubs.ListCount - 1
            Select Case lbSubs.List(i, 1)
                Case "R", "C", "M", "T"
                    iR = iR + 1
                Case "B", "S"
                    iB = iB + 1
            End Select
        Next i
        
        Dim vData As Variant
        
        vData = Split(tbData.Value, vbCr)
        vData(5) = iR & "," & iB
        strLine = vData(0)
        For i = 1 To UBound(vData)
            strLine = strLine & vbCr & vData(i)
        Next i
        
        tbData.Value = strLine
    End If
    
    Me.show
End Sub
