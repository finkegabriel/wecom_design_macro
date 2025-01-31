VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertKML 
   Caption         =   "Convert KML"
   ClientHeight    =   8385.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   OleObjectBlob   =   "ConvertKML.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConvertKML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbConvert_Change()
    If lbDesc.ListCount < 1 Then Exit Sub
    
    If cbConvert.Value Then
        lbDesc.AddItem "<TN83F>"
    Else
        lbDesc.RemoveItem lbDesc.ListCount - 1
    End If
End Sub

Private Sub cbConvertKML_Click()
    If tbFile.Value = "" Then Exit Sub
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    Dim strFile, strTemp, strPre As String
    Dim strValue As String
    Dim vFile, vTemp As Variant
    Dim lCount, lNum As Long
    Dim strAtt() As String
    Dim iCount As Integer
    Dim dLat, dLong As Double
    Dim vNE As Variant
    Dim dInsert(0 To 2) As Double
    Dim dScale As Double
    
    dScale = CDbl(cbScale.Value)
    For i = 0 To lbTo.ListCount - 1
        If Not lbTo.List(i, 2) = "" Then iCount = i
    Next i
    
    ReDim strAtt(iCount)
    
    For i = 0 To iCount
        strAtt(i) = lbTo.List(i, 2)
    Next i
    
    strFile = tbFile.Value
    
    On Error Resume Next
    
    Open strFile For Input As #1
    If Not Err = 0 Then
        MsgBox "Error opening file."
        Exit Sub
    End If
    
    vFile = Split(Input$(LOF(1), 1), vbLf)
    Close #1
    
    If Not Err = 0 Then
        MsgBox Err.Description
        Err = 0
    End If
    
    For lNum = 0 To UBound(vFile)
        strTemp = Replace(vFile(lNum), vbTab, "")
        If LCase(Left(strTemp, 5)) = "</pla" Then GoTo Next_L
        
        If Not LCase(Left(strTemp, 4)) = "<pla" Then
            GoTo Next_lNum
        Else
            lNum = lNum + 1
            
            strTemp = Replace(vFile(lNum), vbTab, "")
            'MsgBox strTemp
            While Not LCase(Left(strTemp, 5)) = "</pla"
                If LCase(Left(strTemp, 5)) = "<name" Then
                    'MsgBox strTemp & vbCr & InStr(lbTo.List(0, 2), "<name>")
                    If InStr(lbTo.List(0, 2), "<name>") > 0 Then
                        strTemp = Replace(strTemp, "<name>", "")
                        strTemp = Replace(strTemp, "</name>", "")
                        
                        For j = 0 To iCount
                            If InStr(strAtt(j), "<name>") > 0 Then
                                strPre = "{<name>}"

                                If Left(strTemp, 1) = "&" Then
                                    strAtt(j) = Replace(strAtt(j), strPre, "NA")
                                Else
                                    strAtt(j) = Replace(strAtt(j), strPre, strTemp)
                                End If
                            End If
                        Next j
                    End If
                End If
                
                If LCase(Left(strTemp, 3)) = "<th" Then
                    strValue = Replace(strTemp, "<th>", "")
                    strValue = Replace(strValue, "</th>", "")
                    
                    lNum = lNum + 1
                    strTemp = Replace(vFile(lNum), vbTab, "")
                    If Left(strTemp, 3) = "<td" Then
                        strTemp = Replace(strTemp, "<td>", "")
                        strTemp = Replace(strTemp, "</td>", "")
                        
                        For j = 0 To iCount
                            If InStr(strAtt(j), strValue) > 0 Then
                                strPre = "{" & strValue & "}"

                                If Left(strTemp, 1) = "&" Then
                                    strAtt(j) = Replace(strAtt(j), strPre, "NA")
                                Else
                                    strAtt(j) = Replace(strAtt(j), strPre, strTemp)
                                End If
                            End If
                        Next j
                    End If
                End If
                
'                If LCase(Left(strTemp, 3)) = "<tr" Then
'                    lNum = lNum + 2
'                    strTemp = Replace(vFile(lNum), vbTab, "")
'                    If Left(strTemp, 3) = "<td" Then
'                        strTemp = Replace(strTemp, "<td>", "")
'                        strTemp = Replace(strTemp, "</td>", "")
'
'                        For j = 0 To iCount
'                            If InStr(strAtt(j), strTemp) > 0 Then
'                                strPre = "{" & strTemp & "}"
'
'                                'lNum = lNum + 2
'                                'strTemp = Replace(vFile(lNum), vbTab, "")
'                                'strTemp = Replace(strTemp, "<td>", "")
'                                'strTemp = Replace(strTemp, "</td>", "")
'                                If Left(strTemp, 1) = "&" Then
'                                    strAtt(j) = Replace(strAtt(j), strPre, "NA")
'                                Else
'                                    strAtt(j) = Replace(strAtt(j), strPre, strTemp)
'                                End If
'                            End If
'                        Next j
'                    End If
'                End If
                
                If LCase(Left(strTemp, 5)) = "<coor" Then
                    
                    'MsgBox "Here"
                    strTemp = Replace(strTemp, "<coordinates>", "")
                    strTemp = Replace(strTemp, "</coordinates>", "")
                    vTemp = Split(strTemp, ",")
                        
                    For j = 0 To iCount
                        If lbTo.List(j, 2) = "{<LL>}" Then
                            strAtt(j) = vTemp(1) & "," & vTemp(0)
                        End If
                    Next j
                        
                    If cbConvert.Value Then
                        vNE = LLtoTN83F(CDbl(vTemp(1)), CDbl(vTemp(0)))
                        'MsgBox vTemp(1) & " , " & vTemp(0) & vbCr & vNE(0) & vbCr & vNE(1)
                        
                        For j = 0 To iCount
                            If lbTo.List(j, 2) = "{<TN83F>}" Then
                                strAtt(j) = vNE(0) & "," & vNE(1)
                            End If
                        Next j
                            
                        dInsert(0) = CDbl(vNE(1))
                        dInsert(1) = CDbl(vNE(0))
                        dInsert(2) = 0#
                    Else
                        dInsert(0) = CDbl(vTemp(0))
                        dInsert(1) = CDbl(vTemp(1))
                        dInsert(2) = 0#
                    End If
                    'MsgBox dInsert(0) & vbCr & dInsert(1) & vbCr & dInsert(2)
                End If
                
                lNum = lNum + 1
                strTemp = Replace(vFile(lNum), vbTab, "")
            Wend
        End If
Next_L:
        'MsgBox dInsert(0) & vbCr & dInsert(1) & vbCr & dInsert(2)
        'MsgBox cbToList.Value & vbCr & strAtt(0) & vbCr & UBound(strAtt)
        'Exit Sub
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dInsert, cbToList.Value, dScale, dScale, dScale, 0#)
        objBlock.Layer = cbToLayer.Value
        vAttList = objBlock.GetAttributes
        
        For k = 0 To iCount
            vAttList(k).TextString = strAtt(k)
        Next k
        objBlock.Update
    
        For i = 0 To iCount
            strAtt(i) = lbTo.List(i, 2)
        Next i
Next_lNum:
    Next lNum
Exit_For:

    MsgBox "Done"
    
End Sub

Private Sub cbFile_Click()
    Dim objSS As AcadSelectionSet
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    
    Dim objBlock As AcadBlockReference
    Dim vAtt As Variant
    
    Dim strPath, strFile As String
    Dim strLine As String
    
  On Error Resume Next
    
    strPath = ThisDrawing.Path & "\"
    strFile = strPath & "Pole Data.txt"
    
    Open strFile For Output As #1
    If Not Err = 0 Then
        MsgBox "Error opening file."
        Exit Sub
    End If

    grpCode(0) = 2
    grpValue(0) = "iPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    If Not Err = 0 Then
        MsgBox "Error: " & Err.Number & vbCr & Err.Description
        Exit Sub
    End If
    
    For Each objBlock In objSS
        vAtt = objBlock.GetAttributes
        
        If vAtt(0).TextString = "" Then GoTo Next_objBlock
        
        strLine = vAtt(0).TextString & vbTab & vAtt(1).TextString & vbTab & vAtt(5).TextString & vbTab & vAtt(7).TextString
        
        Print #1, strLine
        
Next_objBlock:
    Next objBlock
    
    Close #1
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbToList_Change()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS2 As AcadSelectionSet
    
    On Error Resume Next
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    grpCode(0) = 2
    grpValue(0) = cbToList.Value
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
    If Not Err = 0 Then
        Set objSS2 = ThisDrawing.SelectionSets.Item("objSS2")
        objSS2.Clear
    End If
    
    objSS2.Select acSelectionSetAll, , , filterType, filterValue
        
    For Each objBlock In objSS2
        vAttList = objBlock.GetAttributes
        lbTo.Clear
        For i = 0 To UBound(vAttList)
            lbTo.AddItem
            lbTo.List(i, 0) = i
            lbTo.List(i, 1) = vAttList(i).TagString
            lbTo.List(i, 2) = ""
        Next i
        GoTo Exit_Next
    Next objBlock
Exit_Next:
End Sub

Private Sub cbUpdate_Click()
    lbTo.List(lbTo.ListIndex, 1) = lAttTag.Caption
    lbTo.List(lbTo.ListIndex, 2) = tbValue.Value
    cbUpdate.Enabled = False
    
    lAttTag.Caption = "Attribute Tag"
    tbValue.Value = ""
End Sub

Private Sub lbDesc_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If cbUpdate.Enabled = True Then
        tbValue = tbValue & "{" & lbDesc.List(lbDesc.ListIndex, 0) & "}"
        tbValue.SetFocus
    End If
End Sub

Private Sub lbFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strFile, strTemp As String
    Dim vFile, vTemp As Variant
    Dim lCount, l As Long
    
    strFile = ThisDrawing.Path & "\" & lbFiles.List(lbFiles.ListIndex)
    tbFile.Value = strFile
    
    On Error Resume Next
    
    Open strFile For Input As #1
    If Not Err = 0 Then
        MsgBox "Error opening file."
        Exit Sub
    End If
    
    lbDesc.Clear
    
    vFile = Split(Input$(LOF(1), 1), vbLf)
    Close #1
    
    For l = 0 To UBound(vFile)
        strTemp = Replace(vFile(l), vbTab, "")
        If LCase(Left(strTemp, 5)) = "</pla" Then GoTo Exit_For
        
        If Left(strTemp, 3) = "<th" Then
            'l = l + 2
            'strTemp = Replace(vFile(l), vbTab, "")
            'If Left(strTemp, 3) = "<td" Then
                strTemp = Replace(strTemp, "<th>", "")
                strTemp = Replace(strTemp, "</th>", "")
                
                If Not strTemp = "" Then lbDesc.AddItem strTemp
            'End If
        End If
        
        'If Left(strTemp, 3) = "<tr" Then
            'l = l + 2
            'strTemp = Replace(vFile(l), vbTab, "")
            'If Left(strTemp, 3) = "<td" Then
                'strTemp = Replace(strTemp, "<td>", "")
                'strTemp = Replace(strTemp, "</td>", "")
                
                'If Not strTemp = "" Then lbDesc.AddItem strTemp
            'End If
        'End If
    Next l
Exit_For:
    
    lbDesc.AddItem "<name>"
    lbDesc.AddItem "<LL>"
    lbDesc.AddItem "<TN83F>"
    
    'Close #1
End Sub

Private Sub lbTo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    lAttTag.Caption = lbTo.List(lbTo.ListIndex, 1)
    tbValue.Value = lbTo.List(lbTo.ListIndex, 2)
    cbUpdate.Enabled = True
    tbValue.SetFocus
End Sub

Private Sub UserForm_Initialize()
    cbScale.AddItem ""
    cbScale.AddItem "0.5"
    cbScale.AddItem "0.75"
    cbScale.AddItem "1.0"
    cbScale.AddItem "2.0"
    cbScale.AddItem "10"
    cbScale.AddItem "12"
    cbScale.Value = "1.0"
    
    lbTo.Clear
    lbTo.ColumnCount = 3
    lbTo.ColumnWidths = "20;80;70"
    
    
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Integer
    Dim vName As Variant
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getfolder(ThisDrawing.Path)
    
    For Each objFile In objFolder.Files
        vName = Split(objFile.Name, ".")
        If LCase(vName(UBound(vName))) = "kml" Then lbFiles.AddItem objFile.Name
    Next objFile
    
    
    Dim objBlocks As AcadBlocks
    Dim strLine As String
            
    Set objBlocks = ThisDrawing.Blocks
    For i = 0 To objBlocks.count - 1
        strLine = objBlocks(i).Name
        If Not Left(strLine, 1) = "*" Then cbToList.AddItem objBlocks(i).Name
    Next i
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        cbToLayer.AddItem objLayer.Name
    Next objLayer
    cbToLayer.Value = "0"
    
End Sub

Private Function TN83FtoLL(dNorth As Double, dEast As Double)
    Dim dLat, dLong, dDLat As Double
    Dim dDiffE, dEast0 As Double
    Dim dDiffN, dNorth0 As Double
    Dim dR, dCAR, dCA, dU, dK As Double
    Dim LL(2) As Double
    
    dEast = dEast * 0.3048006096
    dNorth = dNorth * 0.3048006096
    
    dDiffE = dEast - 600000
    dDiffN = dNorth - 166504.1691
    
    dR = 8842127.1422 - dDiffN
    dCAR = Atn(dDiffE / dR)
    dCA = dCAR * 180 / 3.14159265359
    dLong = -86 + dCA / 0.585439726459
    
    dU = dDiffN - dDiffE * Tan(dCAR / 2)
    dDLat = dU * (0.00000901305249 + dU * (-6.77268E-15 + dU * (-3.72351E-20 + dU * -9.2828E-28)))
    dLat = 35.8340607459 + dDLat
    
    dK = 0.999948401424 + (1.23188E-14 * dU * dU) + (4.54E-22 * dU * dU * dU)
    
    LL(0) = dLat
    LL(1) = dLong
    LL(2) = dK
    
    TN83FtoLL = LL
End Function

Private Function LLtoTN83F(dLat As Double, dLong As Double)
    Dim dDLat As Double
    Dim dEast, dDiffE, dEast0 As Double
    Dim dNorth, dDiffN, dNorth0 As Double
    Dim dU, dR, dCA, dK As Double
    Dim NE(2) As Double
    
    dDLat = dLat - 35.8340607459
    dU = dDLat * (110950.2019 + dDLat * (9.25072 + dDLat * (5.64572 + dDLat * 0.017374)))
    dR = 8842127.1422 - dU
    dCA = ((86 + dLong) * 0.585439726459) * 3.14159265359 / 180
    
    dDiffE = dR * Sin(dCA)
    dDiffN = dU + dDiffE * Tan(dCA / 2)
    
    dEast = (dDiffE + 600000) / 0.3048006096
    dNorth = (dDiffN + 166504.1691) / 0.3048006096
    
    'dEast = (dDiffE + 600000) / 0.3048
    'dNorth = (dDiffN + 166504.1691) / 0.3048
    
    dK = 0.999948401424 + (1.23188E-14 * dU * dU) + (4.54E-22 * dU * dU * dU)
    
    NE(0) = dNorth
    NE(1) = dEast
    NE(2) = dK
    
    LLtoTN83F = NE
End Function
