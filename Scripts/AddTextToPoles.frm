VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddTextToPoles 
   Caption         =   "Add Text to Poles"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945.001
   OleObjectBlob   =   "AddTextToPoles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddTextToPoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAddToPole_Click()
    Dim objMText As AcadMText
    Dim objBlock As AcadBlockReference
    Dim objEntity As AcadEntity
    Dim vAttList As Variant
    Dim vReturnPnt As Variant
    Dim strLine As String
    Dim vLine, vItem As Variant
    Dim iCOMM As Integer
    
    On Error Resume Next
    Me.Hide
    
Get_Data:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select MText:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadMText Then GoTo Exit_Sub
    
    Set objMText = objEntity
    strLine = UCase(objMText.TextString)
    strLine = Replace(strLine, "\P", vbCr)
    strLine = Replace(strLine, vbLf, vbCr)
    strLine = Replace(strLine, vbTab, "")
    strLine = Replace(strLine, ": ", ":")
    strLine = Replace(strLine, "= ", "=")
    vLine = Split(strLine, vbCr)
    'tbMText.Value = strLine
    'MsgBox UBound(vLine) & vbCr & vbCr & strLine
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select a Block:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    Select Case objBlock.Name
        Case "sPole"
            vAttList = objBlock.GetAttributes
            iCOMM = 1
            
            For i = 0 To UBound(vLine)
                If InStr(vLine(i), tbStatus.Value) > 0 Then
                    vItem = Split(vLine(i), tbStatus.Value)
                    If vItem(0) = "" Then
                        vAttList(1).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbOwner.Value) > 0 Then
                    vItem = Split(vLine(i), tbOwner.Value)
                    If vItem(0) = "" Then
                        vAttList(2).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbOwnerNumber.Value) > 0 Then
                    vItem = Split(vLine(i), tbOwnerNumber.Value)
                    If vItem(0) = "" Then
                        vAttList(3).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbOtherNumber.Value) > 0 Then
                    vItem = Split(vLine(i), tbOtherNumber.Value)
                    If vItem(0) = "" Then
                        vAttList(4).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbHC.Value) > 0 Then
                    vItem = Split(vLine(i), tbHC.Value)
                    If vItem(0) = "" Then
                        vAttList(5).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbGRD.Value) > 0 Then
                    vItem = Split(vLine(i), tbGRD.Value)
                    If vItem(0) = "" Then
                        vAttList(8).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbN.Value) > 0 Then
                    vItem = Split(vLine(i), tbN.Value)
                    If vItem(0) = "" Then
                        vAttList(9).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbT.Value) > 0 Then
                    vItem = Split(vLine(i), tbT.Value)
                    If vItem(0) = "" Then
                        vAttList(10).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbLP.Value) > 0 Then
                    vItem = Split(vLine(i), tbLP.Value)
                    If vItem(0) = "" Then
                        vAttList(11).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbA.Value) > 0 Then
                    vItem = Split(vLine(i), tbA.Value)
                    If vItem(0) = "" Then
                        vAttList(12).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbSLC.Value) > 0 Then
                    vItem = Split(vLine(i), tbSLC.Value)
                    If vItem(0) = "" Then
                        vAttList(13).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbSL.Value) > 0 Then
                    vItem = Split(vLine(i), tbSL.Value)
                    If vItem(0) = "" Then
                        vAttList(14).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                If InStr(vLine(i), tbNew.Value) > 0 Then
                    vItem = Split(vLine(i), tbNew.Value)
                    If vItem(0) = "" Then
                        vAttList(15).TextString = vItem(1)
                        GoTo Next_line
                    End If
                End If
                
                vAttList(15 + iCOMM).TextString = vLine(i)
                iCOMM = iCOMM + 1
                
Next_line:
            Next i
            
            objBlock.Update
        Case "cable_span"
            vAttList = objBlock.GetAttributes
            vAttList(0).TextString = vLine(0)
            objBlock.Update
        Case "Existing_Guys"
            vAttList = objBlock.GetAttributes
            iCOMM = 2
            
            For i = 0 To UBound(vLine)
                If InStr(vLine(i), tbAStatus.Value) > 0 Then
                    vItem = Split(vLine(i), tbAStatus.Value)
                    If vItem(0) = "" Then
                        vAttList(9).TextString = vItem(1)
                        GoTo Next_Line2
                    End If
                End If
                
                If InStr(vLine(i), tbAOffset.Value) > 0 Then
                    vItem = Split(vLine(i), tbAOffset.Value)
                    If vItem(0) = "" Then
                        vAttList(8).TextString = vItem(1)
                        GoTo Next_Line2
                    End If
                End If
                
                If InStr(vLine(i), tbAPower.Value) > 0 Then
                    vItem = Split(vLine(i), tbAPower.Value)
                    If vItem(0) = "" Then
                        vAttList(0).TextString = vItem(1)
                        GoTo Next_Line2
                    End If
                End If
                
                If InStr(vLine(i), tbANew.Value) > 0 Then
                    vItem = Split(vLine(i), tbANew.Value)
                    If vItem(0) = "" Then
                        vAttList(1).TextString = vItem(1)
                        GoTo Next_Line2
                    End If
                End If
                
                vAttList(iCOMM).TextString = vLine(i)
                iCOMM = iCOMM + 1
                
Next_Line2:
            Next i
        Case ""
        Case Else
            GoTo Exit_Sub
    End Select
    
    GoTo Get_Data
    
Exit_Sub:
    
    Me.show
End Sub

Private Sub cbGetText_Click()
    Dim objMText As AcadMText
    Dim objEntity As AcadEntity
    Dim vReturnPnt As Variant
    Dim strLine As String
    
    On Error Resume Next
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select MText:"
    If Not Err = 0 Then GoTo Exit_Sub
    If Not TypeOf objEntity Is AcadMText Then GoTo Exit_Sub
    
    Set objMText = objEntity
    strLine = UCase(objMText.TextString)
    'tbMText.Value = strLine
    strLine = Replace(strLine, vbLf, vbCr)
    strLine = Replace(strLine, vbTab, "")
    strLine = Replace(strLine, ": ", ":")
    strLine = Replace(strLine, "= ", "=")
    tbMText.Value = Replace(strLine, "\P", vbCr)
        
Exit_Sub:
        
        Me.show
End Sub
