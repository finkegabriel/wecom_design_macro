VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zMisc 
   Caption         =   "Miscellaneous Programs"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   OleObjectBlob   =   "zMisc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAddMissingCo_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    Dim strLine, strTemp As String
    Dim vLine, vTemp As Variant
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(4).TextString = "" Then GoTo Next_objBlock
        
        If vAttList(2).TextString = tbAtt2.Value Then
            vLine = Split(vAttList(4).TextString, " ")
            For i = 0 To UBound(vLine)
                vTemp = Split(vLine(i), "=")
                If UBound(vTemp) < 1 Then vLine(i) = tbAtt4.Value & "=" & vTemp(0)
            Next i
            
            strTemp = vLine(0)
            If UBound(vLine) > 0 Then
                For i = 0 To UBound(vLine)
                    strTemp = strTemp & " " & vLine(i)
                Next i
            End If
            
            vAttList(4).TextString = strTemp
            objBlock.Update
        End If
        
Next_objBlock:
    Next objBlock
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub cbAddRoadNames_Click()
    Dim amap As AcadMap
    Dim ODRcs As ODRecords
    Dim tbl As ODTable
    Dim tbls As ODTables
    Dim boolVal As Boolean
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    'Dim objPolyline As AcadPolyline
    Dim objLine As AcadLine
    Dim objMText As AcadMText
    Dim vReturnPnt, vCoords As Variant
    Dim vStartPnt, vEndPnt As Variant
    Dim dDiff(0 To 2) As Double
    Dim dStartPnt(0 To 2) As Double
    Dim dEndPnt(0 To 2) As Double
    Dim dToText(0 To 2) As Double
    Dim dInsertPnt(0 To 2) As Double
    Dim dTest As Double
    Dim dFromText(0 To 2) As Double
    Dim dRotate, dScale As Double
    Dim dDistScale As Double
    Dim Pi As Double
    Dim iPosition, iLength As Integer
    Dim iScale, iHeight As Integer
    Dim strName As String
    
  On Error Resume Next
    
    If IsNumeric(tbRoadTextHeight.Value) Then
        iHeight = CInt(tbRoadTextHeight.Value)
    Else
        iHeight = 10
    End If
  
    Pi = 3.14159265359
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
    Set tbls = amap.Projects(ThisDrawing).ODTables
    
    Err = 0
    If tbls.count > 0 Then
        For Each tbl In tbls
            If tbl.Name = "IS_Streets" Then GoTo exit_for1
        Next
    End If
    
    Exit Sub
exit_for1:

    Set ODRcs = tbl.GetODRecords
    
    Me.Hide
    
Get_Street:

    Err = 0
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Street: "
    
    If Not Err = 0 Then GoTo Exit_Sub
    If Not objEntity.Layer = "IS_Streets" Then GoTo Get_Street
    
    Select Case objEntity.ObjectName
        Case "AcDbLine"
            Set objLine = objEntity
            iLength = CInt(objLine.Length)
            vStartPnt = objLine.startPoint
            vEndPnt = objLine.endPoint
            
            dDiff(0) = vEndPnt(0) - vStartPnt(0)
            dDiff(1) = vEndPnt(1) - vStartPnt(1)
            dDiff(2) = (dDiff(0) * dDiff(0)) + (dDiff(1) * dDiff(1))
            dDiff(2) = Sqr(dDiff(2))
            
            If vEndPnt(0) > vStartPnt(0) Then
                dToText(0) = vReturnPnt(0) - vStartPnt(0)
                dToText(1) = vReturnPnt(1) - vStartPnt(1)
            Else
                dToText(0) = vStartPnt(0) - vReturnPnt(0)
                dToText(1) = vStartPnt(1) - vReturnPnt(1)
            End If
            dToText(2) = Sqr((dToText(0) * dToText(0)) + (dToText(1) * dToText(1)))
            
            dDistScale = dToText(2) / dDiff(2)
            dInsertPnt(0) = vStartPnt(0) + (dDiff(0) * dDistScale)
            dInsertPnt(1) = vStartPnt(1) + (dDiff(1) * dDistScale)
            dInsertPnt(2) = 0#
            
            If dToText(0) = 0 Then
                dRotate = Pi / 2
            Else
                dRotate = Atn(dDiff(1) / dDiff(0))
            End If
            
            boolVal = ODRcs.Init(objEntity, True, False)
            
            strName = ""
            If Not ODRcs.Record.Item(6).Value = "" Then strName = ODRcs.Record.Item(6).Value & " "
            strName = strName & ODRcs.Record.Item(5).Value
            If Not ODRcs.Record.Item(11).Value = "" Then strName = strName & " " & ODRcs.Record.Item(11).Value & " "
            
        Case "AcDbPolyline"
            Set objLWP = objEntity
            iLength = CInt(objLWP.Length)
            vCoords = objLWP.Coordinates
            
            For i = 3 To UBound(vCoords) Step 2
                dDiff(0) = vCoords(i - 1) - vCoords(i - 3)
                dDiff(1) = vCoords(i) - vCoords(i - 2)
                dDiff(2) = Sqr((dDiff(0) * dDiff(0)) + (dDiff(1) * dDiff(1)))
            
                dToText(0) = vReturnPnt(0) - vCoords(i - 3)
                dToText(1) = vReturnPnt(1) - vCoords(i - 2)
                dToText(2) = Sqr((dToText(0) * dToText(0)) + (dToText(1) * dToText(1)))
                
                If dDiff(2) > dToText(2) Then
                    dFromText(0) = vReturnPnt(0) - vCoords(i - 1)
                    dFromText(1) = vReturnPnt(1) - vCoords(i)
                    dFromText(2) = Sqr((dFromText(0) * dFromText(0)) + (dFromText(1) * dFromText(1)))
                    
                    dTemp = (dFromText(2) + dToText(2)) / dDiff(2)
                    If dTemp < 1.01 Then
                        GoTo Exit_First_Next
                    End If
                End If
            Next i
Exit_First_Next:
            
            dDistScale = Abs(dToText(2) / dDiff(2))
            dInsertPnt(0) = vCoords(i - 3) + (dDiff(0) * dDistScale)
            dInsertPnt(1) = vCoords(i - 2) + (dDiff(1) * dDistScale)
            dInsertPnt(2) = 0#
            
            If dToText(0) = 0 Then
                dRotate = Pi / 2
            Else
                dRotate = Atn(dDiff(1) / dDiff(0))
            End If
            
            boolVal = ODRcs.Init(objEntity, True, False)
            
            strName = ""
            If Not ODRcs.Record.Item(6).Value = "" Then strName = ODRcs.Record.Item(6).Value & " "
            strName = strName & ODRcs.Record.Item(5).Value
            If Not ODRcs.Record.Item(11).Value = "" Then strName = strName & " " & ODRcs.Record.Item(11).Value & " "
            
        Case Else
            Me.show
            Exit Sub
    End Select
    
    iScale = CInt(Len(strName) * 9 * iHeight / 10)

    'Set objMText = ThisDrawing.ModelSpace.AddMText(dInsertPnt, iScale, strName)
    Set objMText = ThisDrawing.ModelSpace.AddMText(dInsertPnt, 0, strName)
    objMText.Layer = "IS_StreetsText"
    objMText.Height = iHeight
    objMText.AttachmentPoint = acAttachmentPointMiddleCenter
    objMText.InsertionPoint = dInsertPnt
    objMText.Rotation = dRotate
    'objMText.BackgroundFill = True
    objMText.Update
    
    GoTo Get_Street
    
Exit_Sub:
    'MsgBox Err.Number & vbCr & Err.Description
    
    Err.Clear
    
    Me.show
End Sub

Private Sub cbAddRoute_Click()
    Dim objSS2 As AcadSelectionSet
    Dim entObjects As AcadEntity
    Dim objPole As AcadBlockReference
    Dim vAttList As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    
    Me.Hide
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
    objSS2.SelectOnScreen filterType, filterValue
    'objSS2.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objPole In objSS2
        vAttList = objPole.GetAttributes
        vAttList(0).TextString = UCase(tbPoleRoute.Value) & "/" & vAttList(0).TextString
        objPole.Update
    Next objPole
    
    objSS2.Clear
    objSS2.Delete
    Me.show
End Sub

Private Sub cbBlockName_Change()
    If cbBlockName.Value = "MiscUtility" Then tbAttribute.Enabled = True
End Sub

Private Sub cbConvertToCustomer_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim objBlock As AcadBlockReference
    Dim objCustomer As AcadBlockReference
    Dim vAttList, vCustList As Variant
    Dim iCount As Integer
    
    On Error Resume Next
    
    iCount = 0
    
    grpCode(0) = 2
    grpValue(0) = "RES,TRLR,MDU,BUSINESS,CHURCH,SCHOOL,EXTENTION"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    MsgBox "Found:  " & objSS.count
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        Set objCustomer = ThisDrawing.ModelSpace.InsertBlock(objBlock.InsertionPoint, "Customer", 1#, 1#, 1#, 0#)
        vCustList = objCustomer.GetAttributes
        
        vCustList(1).TextString = vAttList(0).TextString
        vCustList(2).TextString = vAttList(1).TextString
        vCustList(3).TextString = vAttList(2).TextString
        
        Select Case objBlock.Name
            Case "BUSINESS"
                vCustList(0).TextString = "BUSINESS"
                vCustList(5).TextString = "B"
            Case "CHURCH"
                vCustList(0).TextString = "CHURCH"
                vCustList(5).TextString = "C"
            Case "EXTENTION"
                vCustList(0).TextString = "EXTENSION"
                vCustList(5).TextString = "X"
            Case "MDU"
                vCustList(0).TextString = "MDU"
                vCustList(5).TextString = "M"
            Case "RES"
                vCustList(0).TextString = "RESIDENCE"
                vCustList(5).TextString = ""
            Case "SCHOOL"
                vCustList(0).TextString = "SCHOOL"
                vCustList(5).TextString = "S"
            Case "TRLR"
                vCustList(0).TextString = "TRAILER"
                vCustList(5).TextString = "T"
        End Select
        
        objCustomer.Layer = "Customers"
        objCustomer.Update
        
        iCount = iCount + 1
    Next objBlock
    
    MsgBox "Converted  " & iCount & "  customers"
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub cbConvertTosFP_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    
    Dim objPole As AcadBlockReference
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    Dim strAtt(0 To 7) As String
    Dim iAtt As Integer
    Dim vLine, vTemp As Variant
    Dim dScale As Double
    
    On Error Resume Next
    
    'iAtt = 16
    
    grpCode(0) = 2
    grpValue(0) = "dFP"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objPole In objSS
        For i = 0 To 7
            strAtt(i) = ""
        Next i
        
        dScale = objPole.XScaleFactor
        
        vAttList = objPole.GetAttributes
        
        strAtt(0) = vAttList(0).TextString
        strAtt(1) = vAttList(0).TextString
        strAtt(2) = vAttList(2).TextString
        If Not vAttList(3).TextString = "" Then strAtt(3) = vAttList(3).TextString
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(objPole.InsertionPoint, "sFP", dScale, dScale, dScale, 0#)
        vAttList = objBlock.GetAttributes
        
        For i = 0 To 7
            If strAtt(i) = "" Then
                vAttList(i).TextString = ""
            Else
                vAttList(i).TextString = strAtt(i)
            End If
        Next i
        
        objBlock.Layer = objPole.Layer
        objBlock.Update
        
        objPole.Layer = "Integrity Delete"
        objPole.Update
        
        Err = 0
    Next objPole
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub cbConvertTosHH_Click()
    Dim objEntity As AcadEntity
    Dim objBlock, objHH As AcadBlockReference
    Dim vAttList, vHH, vReturnPnt As Variant
    
    On Error GoTo Exit_Sub
    Me.Hide
    
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Handhole:"
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    vAttList = objBlock.GetAttributes
    
    Set objHH = ThisDrawing.ModelSpace.InsertBlock(objBlock.InsertionPoint, "sHH", 1#, 1#, 1#, 0#)
    objHH.Layer = "Integrity Proposed"
    vHH = objHH.GetAttributes
    
    vHH(0).TextString = vAttList(0).TextString
    vHH(2).TextString = vAttList(1).TextString
    
    objHH.Update
    
    objBlock.Layer = "Integrity Delete"
    objBlock.Update
    
    GoTo Get_Another
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbConvertTosPed_Click()
    Dim objEntity As AcadEntity
    Dim objBlock, objHH As AcadBlockReference
    Dim vAttList, vHH, vReturnPnt As Variant
    
    On Error GoTo Exit_Sub
    Me.Hide
    
Get_Another:
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Ped:"
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    
    Set objBlock = objEntity
    vAttList = objBlock.GetAttributes
    
    Set objHH = ThisDrawing.ModelSpace.InsertBlock(objBlock.InsertionPoint, "sPed", 1#, 1#, 1#, 0#)
    objHH.Layer = "Integrity Proposed"
    vHH = objHH.GetAttributes
    
    vHH(0).TextString = vAttList(0).TextString
    vHH(2).TextString = vAttList(1).TextString
    
    objHH.Update
    
    objBlock.Layer = "Integrity Delete"
    objBlock.Update
    
    GoTo Get_Another
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbConvertTosPole_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    
    Dim objPole As AcadBlockReference
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    Dim strAtt(0 To 23) As String
    Dim iAtt As Integer
    Dim vLine, vTemp As Variant
    Dim dScale As Double
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "iPole"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objPole In objSS
        For i = 0 To 23
            strAtt(i) = ""
        Next i
        iAtt = 16
        
        dScale = objPole.XScaleFactor
        
        vAttList = objPole.GetAttributes
        
        strAtt(0) = vAttList(0).TextString
        strAtt(1) = vAttList(1).TextString
        strAtt(2) = vAttList(2).TextString
        strAtt(3) = vAttList(3).TextString
        strAtt(4) = vAttList(4).TextString
        strAtt(5) = vAttList(5).TextString
        strAtt(6) = vAttList(6).TextString
        strAtt(7) = vAttList(7).TextString
        strAtt(8) = vAttList(8).TextString
        strAtt(9) = vAttList(9).TextString
        strAtt(10) = vAttList(10).TextString
        strAtt(11) = vAttList(11).TextString
        strAtt(12) = vAttList(12).TextString
        strAtt(13) = vAttList(13).TextString
        strAtt(14) = vAttList(14).TextString
        
        For i = 15 To 26
            If Not vAttList(i).TextString = "" Then
                If InStr(vAttList(i).TextString, "=") > 1 Then
                    If i = 18 Then
                        If InStr(UCase(vAttList(i).TextString), "T") > 0 Then
                            vLine = Split(vAttList(i).TextString, "=")
                            
                            strAtt(15) = vLine(1)
                        Else
                            strAtt(iAtt) = vAttList(i).TextString
                            iAtt = iAtt + 1
                        End If
                    Else
                        strAtt(iAtt) = vAttList(i).TextString
                        iAtt = iAtt + 1
                    End If
                Else
                    Select Case i
                        Case Is = 15
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "CATV=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 16
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "ATT=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 17
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "TDS=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 18
                            If Not vAttList(i).TextString = "" Then
                                If InStr(UCase(vAttList(18).TextString), "T") > 0 Then
                                    strAtt(15) = vAttList(18).TextString
                                Else
                                    strAtt(iAtt) = "UTC=" & vAttList(i).TextString
                                    iAtt = iAtt + 1
                                End If
                            End If
                        Case Is = 19
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "XO=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 20
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "ZAYO=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 21
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "CITY=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 22
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "TELCO=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 23
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "CITY=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 24
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "TRAFFIC=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 25
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "OHG=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                        Case Is = 26
                            If Not vAttList(i).TextString = "" Then
                                strAtt(iAtt) = "OTHER=" & vAttList(i).TextString
                                iAtt = iAtt + 1
                            End If
                    End Select
                End If
            End If
        Next i
        
        For i = iAtt To 23
            strAtt(i) = ""
        Next i
        
        Set objBlock = ThisDrawing.ModelSpace.InsertBlock(objPole.InsertionPoint, "sPole", dScale, dScale, dScale, 0#)
        vAttList = objBlock.GetAttributes
        
        For i = 0 To 23
            If strAtt(i) = "" Then
                vAttList(i).TextString = ""
            Else
                vAttList(i).TextString = strAtt(i)
            End If
        Next i
        
        objBlock.Layer = objPole.Layer
        objBlock.Update
        
        objPole.Layer = "Integrity Delete"
        objPole.Update
        
        Err = 0
    Next objPole
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub cbGetAttribute_Click()
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vReturnPnt, vAttList As Variant
    Dim strLine As String
    Dim iAtt As Integer
        
    On Error Resume Next
    
    If tbAttNumber.Value = "" Then
        iAtt = 0
    Else
        iAtt = CInt(tbAttNumber.Value)
    End If
    
    Me.Hide
    
    ThisDrawing.Utility.GetEntity objEntity, vReturnPnt, vbCr & "Select Block with Attributes: "
    
    If Not TypeOf objEntity Is AcadBlockReference Then GoTo Exit_Sub
    Set objBlock = objEntity
    
    vAttList = objBlock.GetAttributes
    strLine = vAttList(iAtt).TextString
    If strLine = "" Then GoTo Exit_Sub
    
    'strLine = Replace(strLine, vbCrLf, "<cl>")
    'strLine = Replace(strLine, vbCr, "<c>")
    'strLine = Replace(strLine, vbLf, "<l>")
    'strLine = Replace(strLine, "<c>", "<c>" & vbCr)
    
    tbAttributeValue.Value = strLine
    
Exit_Sub:
    Me.show
End Sub

Private Sub cbPntToBlock_Click()
    If lbLayers.ListIndex < 0 Then Exit Sub
    If cbBlockName.Value = "" Then Exit Sub
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    
    Dim objEntity As AcadEntity
    Dim objPoint As AcadPoint
    Dim objBlock As AcadBlockReference
    Dim vAtt As Variant
    Dim strBlockName, strLayer As String
    Dim dInsertPnt(0 To 2) As Double
    Dim dScale As Double
    
    If cbScale.Value = "" Then
        dScale = 1#
    Else
        dScale = CDbl(cbScale.Value) / 100
    End If
    
    strBlockName = cbBlockName.Value
    strLayer = lbLayers.List(lbLayers.ListIndex)
    
    grpCode(0) = 8
    grpValue(0) = strLayer
    filterType = grpCode
    filterValue = grpValue
    
    On Error Resume Next
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    MsgBox lbLayers.List(lbLayers.ListIndex) & " found:  " & objSS.count
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadPoint Then
            Set objPoint = objEntity
            
            dInsertPnt(0) = objPoint.Coordinates(0)
            dInsertPnt(1) = objPoint.Coordinates(1)
            dInsertPnt(2) = objPoint.Coordinates(2)
            
            Set objBlock = ThisDrawing.ModelSpace.InsertBlock(dInsertPnt, strBlockName, dScale, dScale, dScale, 0#)
            objBlock.Layer = strLayer
            If cbBlockName.Value = "MiscUtility" Then
                vAtt = objBlock.GetAttributes
                If Not tbAttribute.Value = "" Then vAtt(0).TextString = UCase(Left(tbAttribute.Value, 2))
            End If
            objBlock.Update
            
            If cbDeletePoint.Value = True Then objPoint.Delete
        End If
    Next objEntity
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub CommandButton1_Click()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    Dim strLine, strTemp As String
    Dim vLine, vTemp As Variant
    
    On Error Resume Next
    
    grpCode(0) = 2
    grpValue(0) = "sPole"
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        If vAttList(4).TextString = "" Then GoTo Next_objBlock
        
        If vAttList(2).TextString = tbAtt2.Value Then
            vLine = Split(vAttList(4).TextString, " ")
            For i = 0 To UBound(vLine)
                vTemp = Split(vLine(i), "=")
                If UBound(vTemp) < 1 Then vLine(i) = tbAtt4.Value & "=" & vTemp(0)
            Next i
            
            strTemp = vLine(0)
            If UBound(vLine) > 0 Then
                For i = 0 To UBound(vLine)
                    strTemp = strTemp & " " & vLine(i)
                Next i
            End If
            
            vAttList(4).TextString = strTemp
            objBlock.Update
        End If
        
Next_objBlock:
    Next objBlock
    
    objSS.Clear
    objSS.Delete
End Sub

Private Sub UserForm_Initialize()
    cbScale.AddItem "50"
    cbScale.AddItem "75"
    cbScale.AddItem "100"
    cbScale.Value = "100"
    
    Dim objBlocks As AcadBlocks
    Dim strLine As String
            
    Set objBlocks = ThisDrawing.Blocks
    For i = 0 To objBlocks.count - 1
        strLine = objBlocks(i).Name
        If Not Left(strLine, 1) = "*" Then cbBlockName.AddItem objBlocks(i).Name
    Next i
    
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    
    Set objLayers = ThisDrawing.Layers
    For Each objLayer In objLayers
        lbLayers.AddItem objLayer.Name
    Next objLayer
    
End Sub
