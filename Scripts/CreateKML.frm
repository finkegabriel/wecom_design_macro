VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateKML 
   Caption         =   "Create KML"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   OleObjectBlob   =   "CreateKML.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateKML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAddDescription_Click()
    If lbFrom.ListIndex < 0 Then Exit Sub
    
    tbDescription.Value = tbDescription.Value & "{" & lbFrom.ListIndex & "}"
End Sub

Private Sub cbAddName_Click()
    If lbFrom.ListIndex < 0 Then Exit Sub
    
    tbName.Value = tbName.Value & "{" & lbFrom.ListIndex & "}"
End Sub

Private Sub cbConvert_Click()
    If tbFolder.Value = "" Then Exit Sub
    If tbFile.Value = "" Then Exit Sub
    
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    
    Dim objEntity As AcadEntity
    Dim objPoint As AcadPoint
    Dim objBlock As AcadBlockReference
    Dim vAttList As Variant
    
    Dim amap As AcadMap
    Dim ODRcs As ODRecords
    Dim ODRc As ODRecord
    Dim tbl As ODTable
    Dim tbls As ODTables
    
    Dim strName, strDescription, strCoords As String
    Dim dNort, dEast As Double
    Dim vLL As Variant
    Dim vLine, vItem, vTemp As Variant
    Dim iIndex As Integer
    Dim strFind, strReplace As String
    Dim strFile, strItem As String
    Dim strFileName As String
    Dim strStyle As String
    
    strFileName = tbFolder.Value & tbFile.Value & ".kml"
    
    strFile = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCr
    strFile = strFile & "<kml xmlns=""http://www.opengis.net/kml/2.2"">" & vbCr
    strFile = strFile & "<Document>" & vbCr & vbCr
    
    If Not tbName.Value = "" Then
        strFile = strFile & "<Style id=""size_label"">" & vbCr
        strFile = strFile & "<LabelStyle>" & vbCr
        strFile = strFile & "<scale>" & tbNameScale.Value & "</scale>" & vbCr
        strFile = strFile & "</LabelStyle>" & vbCr
        strFile = strFile & "</Style>" & vbCr & vbCr
    End If
    
    If Not cbIconFolder.Value = "" Then
        If Not cbIcon.Value = "" Then
            strStyle = "<Style>" & vbCr & "<IconStyle>" & vbCr & "<Scale>" & tbScale.Value & "</Scale>" & vbCr
            strStyle = strStyle & "<Icon>" & vbCr & "<href>"
            strStyle = strStyle & "http://maps.google.com/mapfiles/kml/" & cbIconFolder.Value & "/" & cbIcon.Value & ".png</href>"
            strStyle = strStyle & vbCr & "</Icon>" & vbCr & "</IconStyle>" & vbCr & "</Style>" & vbCr
        End If
    End If
    
    On Error Resume Next
    
    If cbFromType.Value = "Block" Then
        grpCode(0) = 2
    Else
        grpCode(0) = 8
    
        Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
        Set tbls = amap.Projects(ThisDrawing).ODTables
        If tbls.count > 0 Then
            For Each tbl In tbls
                If tbl.Name = cbFromList.Value Then GoTo Exit_For
            Next
        End If
Exit_For:

        Set ODRcs = tbl.GetODRecords
    End If
    
    grpValue(0) = cbFromList.Value
    filterType = grpCode
    filterValue = grpValue
    
    Err = 0
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.Select acSelectionSetAll, , , filterType, filterValue
    MsgBox cbFromType.Value & " found:  " & objSS.count
    
    For Each objEntity In objSS
        strName = tbName.Value
        strDescription = Replace(tbDescription.Value, vbLf, "")
        strDescription = Replace(strDescription, vbCr, "<br/>")
        
        If TypeOf objEntity Is AcadPoint Then
            Set objPoint = objEntity
            dEast = objPoint.Coordinates(0)
            dNorth = objPoint.Coordinates(1)
            vLL = TN83FtoLL(CDbl(dNorth), CDbl(dEast))
            
            If InStr(strName, "{") > 0 Or InStr(strDescription, "{") > 0 Then
                boolVal = ODRcs.Init(objEntity, True, False)
                Set ODRc = ODRcs.Record
                
                If InStr(strName, "}") > 0 Then
                    vLine = Split(strName, "}")
                    For i = 0 To UBound(vLine)
                        vItem = Split(vLine(i), "{")
                        iIndex = CInt(vItem(1))
                        
                        strFind = "{" & iIndex & "}"
                        strReplace = ODRc.Item(iIndex).Value
                        strName = Replace(strName, strFind, strReplace)
                    Next i
                End If
                
                If InStr(strDescription, "}") > 0 Then
                    vLine = Split(strDescription, "}")
                    For i = 0 To UBound(vLine)
                        vItem = Split(vLine(i), "{")
                        iIndex = CInt(vItem(1))
                        
                        strFind = "{" & iIndex & "}"
                        strReplace = ODRc.Item(iIndex).Value
                        strDescription = Replace(strDescription, strFind, strReplace)
                    Next i
                End If
            End If
            
            GoTo Add_Point
        End If
        
        If TypeOf objEntity Is AcadBlockReference Then
            Set objBlock = objEntity
            vAttList = objBlock.GetAttributes
            If vAttList(0).TextString = "POLE" Then GoTo Next_Entity
            
            dEast = objBlock.InsertionPoint(0)
            dNorth = objBlock.InsertionPoint(1)
            vLL = TN83FtoLL(CDbl(dNorth), CDbl(dEast))
            
            If InStr(strName, "}") > 0 Then
                vLine = Split(strName, "}")
                For i = 0 To UBound(vLine)
                    vItem = Split(vLine(i), "{")
                    iIndex = CInt(vItem(1))
                    
                    strFind = "{" & iIndex & "}"
                    strReplace = vAttList(iIndex).TextString
                    strName = Replace(strName, strFind, strReplace)
                Next i
            End If
            
            If InStr(strDescription, "}") > 0 Then
                vLine = Split(strDescription, "}")
                For i = 0 To UBound(vLine)
                    vItem = Split(vLine(i), "{")
                    iIndex = CInt(vItem(1))
                    
                    strFind = "{" & iIndex & "}"
                    strReplace = vAttList(iIndex).TextString
                    strDescription = Replace(strDescription, strFind, strReplace)
                Next i
            End If
            
            GoTo Add_Point
        End If
        
        GoTo Next_Entity
        
Add_Point:
            
        strItem = "<Placemark>" & vbCr
            
        If Not strName = "" Then strItem = strItem & "<name>" & strName & "</name>" & vbCr & "<styleUrl>#size_label</styleUrl>" & vbCr
        If Not strDescription = "" Then strItem = strItem & "<description>" & strDescription & "</description>" & vbCr
        strItem = strItem & strStyle
        
        strItem = strItem & "<Point>" & vbCr & "<coordinates>" & vLL(1) & "," & vLL(0) & "</coordinates>" & vbCr
        strItem = strItem & "</Point>" & vbCr & "</Placemark>" & vbCr
        
        strFile = strFile & strItem & vbCr
Next_Entity:
    Next objEntity
    
    strFile = strFile & "</Document>" & vbCr & "</kml>"
    
    'tbOutput.Value = strFile
    Open strFileName For Output As #1
    
    Print #1, strFile
    Close #1
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
End Sub

Private Sub cbFromList_Change()
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS2 As AcadSelectionSet
    
    On Error Resume Next
    
    If cbFromType.Value = "Block" Then
        Dim objBlock As AcadBlockReference
        Dim vAttList As Variant
    
        grpCode(0) = 2
        grpValue(0) = cbFromList.Value
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
            lbFrom.Clear
            For i = 0 To UBound(vAttList)
                lbFrom.AddItem
                lbFrom.List(i, 0) = i
                lbFrom.List(i, 1) = vAttList(i).TagString
            Next i
            GoTo Exit_Next
        Next objBlock
Exit_Next:
    Else
        Dim amap As AcadMap
        Dim ODRcs As ODRecords
        Dim ODRc As ODRecord
        Dim tbl As ODTable
        Dim tbls As ODTables
        Dim refColumn As ODFieldDef
        Dim objEntity As AcadEntity
        Dim boolVal As Boolean
        Dim iCount As Integer
        Dim strTest As String
    
        Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
        Set tbls = amap.Projects(ThisDrawing).ODTables
        If tbls.count > 0 Then
            For Each tbl In tbls
                If tbl.Name = cbFromList.Value Then GoTo Exit_For
            Next
        End If
Exit_For:

        Set ODRcs = tbl.GetODRecords
        
        iCount = 0
        Err = 0
        
        While Err = 0
            Set refColumn = tbl.ODFieldDefs(iCount)
            strTest = refColumn.Name
            If Not Err = 0 Then GoTo Exit_objEntity
            
            lbFrom.AddItem
            lbFrom.List(iCount, 0) = iCount
            lbFrom.List(iCount, 1) = strTest
            iCount = iCount + 1
        Wend
        
Exit_objEntity:
    End If
    
    objSS2.Clear
    objSS2.Delete
End Sub

Private Sub cbFromType_Change()
    Select Case cbFromType.Value
        Case "Block"
            Dim objBlocks As AcadBlocks
            Dim strLine As String
            
            Set objBlocks = ThisDrawing.Blocks
            For i = 0 To objBlocks.count - 1
                strLine = objBlocks(i).Name
                If Not Left(strLine, 1) = "*" Then cbFromList.AddItem objBlocks(i).Name
            Next i
        Case "Point w/OD"
            Dim amap As AcadMap
            Dim tbl As ODTable
            Dim tbls As ODTables
    
            Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    
            Set tbls = amap.Projects(ThisDrawing).ODTables
            If tbls.count > 0 Then
                cbFromList.Clear
                For Each tbl In tbls
                    cbFromList.AddItem tbl.Name
                Next
            End If
    End Select
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cbIconFolder_Change()
    cbIcon.Clear
    
    Select Case cbIconFolder.Value
        Case "shapes"
            cbIcon.AddItem "placemark_circle"
            cbIcon.AddItem "placemark_square"
            cbIcon.AddItem "square"
            cbIcon.AddItem "star"
            cbIcon.AddItem "polygon"
            cbIcon.AddItem "open-diamond"
            cbIcon.AddItem "shaded_dot"
            cbIcon.AddItem "target"
            cbIcon.AddItem "triangle"
            cbIcon.AddItem "donut"
            cbIcon.AddItem "cross-hairs"
            
            cbIcon.AddItem "phone"
            cbIcon.AddItem "homegardenbusiness"
            cbIcon.AddItem "schools"
            cbIcon.AddItem "church"
            cbIcon.AddItem "arrow"
            cbIcon.AddItem "flag"
            cbIcon.AddItem "caution"
            cbIcon.AddItem "marina"
            cbIcon.AddItem "mechanic"
        Case "pushpin"
            cbIcon.AddItem "blue-pushpin"
            cbIcon.AddItem "grn-pushpin"
            cbIcon.AddItem "ltblu-pushpin"
            cbIcon.AddItem "pink-pushpin"
            cbIcon.AddItem "purple-pushpin"
            cbIcon.AddItem "red-pushpin"
            cbIcon.AddItem "wht-pushpin"
            cbIcon.AddItem "ylw-pushpin"
        Case "paddle"
            cbIcon.AddItem "1"
            cbIcon.AddItem "2"
            cbIcon.AddItem "3"
            cbIcon.AddItem "4"
            cbIcon.AddItem "5"
            cbIcon.AddItem "6"
            cbIcon.AddItem "7"
            cbIcon.AddItem "8"
            cbIcon.AddItem "9"
            cbIcon.AddItem "10"
            cbIcon.AddItem "A"
            cbIcon.AddItem "B"
            cbIcon.AddItem "C"
            cbIcon.AddItem "D"
            cbIcon.AddItem "E"
            cbIcon.AddItem "F"
            cbIcon.AddItem "G"
            cbIcon.AddItem "H"
            cbIcon.AddItem "red-blank"
            cbIcon.AddItem "red-circle"
            cbIcon.AddItem "red-diamond"
            cbIcon.AddItem "red-square"
            cbIcon.AddItem "red-stars"
            cbIcon.AddItem "orange-blank"
            cbIcon.AddItem "orange-circle"
            cbIcon.AddItem "orange-diamond"
            cbIcon.AddItem "orange-square"
            cbIcon.AddItem "orange-stars"
            cbIcon.AddItem "ylw-blank"
            cbIcon.AddItem "ylw-circle"
            cbIcon.AddItem "ylw-diamond"
            cbIcon.AddItem "ylw-square"
            cbIcon.AddItem "ylw-stars"
            cbIcon.AddItem "grn-circle"
            cbIcon.AddItem "grn-diamond"
            cbIcon.AddItem "grn-square"
            cbIcon.AddItem "grn-stars"
            cbIcon.AddItem "blu-blank"
            cbIcon.AddItem "blu-circle"
            cbIcon.AddItem "blu-diamond"
            cbIcon.AddItem "blu-square"
            cbIcon.AddItem "blu-stars"
            cbIcon.AddItem "grn-blank"
            cbIcon.AddItem "purple-blank"
            cbIcon.AddItem "purple-circle"
            cbIcon.AddItem "purple-diamond"
            cbIcon.AddItem "purple-square"
            cbIcon.AddItem "purple-stars"
    End Select
End Sub

Private Sub UserForm_Initialize()
    cbFromType.AddItem "Block"
    cbFromType.AddItem "Point w/OD"
    cbFromType.Value = "Point w/OD"
    
    lbFrom.Clear
    lbFrom.ColumnCount = 2
    lbFrom.ColumnWidths = "20;90"
    
    cbIconFolder.AddItem "shapes"
    cbIconFolder.AddItem "pushpin"
    cbIconFolder.AddItem "paddle"
    'cbIconFolder.AddItem ""
    
    tbFolder.Value = ThisDrawing.Path & "\"
    
    'Dim strLine As String
    
    'strLine = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCr
    'strLine = strLine & "<kml xmlns=""http://www.opengis.net/kml/2.2"">" & vbCr
    'strLine = strLine & "<Document>" & vbCr & "Placemark>" & vbCr
    'strLine = strLine & "<name>Test</name>" & vbCr
    'strLine = strLine & "<description> This is a test point</description>" & vbCr
    'strLine = strLine & "<Point>" & vbCr & "<coordinates>-86.36249947,36.22776216</coordinates>" & vbCr
    'strLine = strLine & "</Point>" & vbCr & "</Placemark>" & vbCr & "</Document>" & vbCr & "</kml>"
    
    'tbOutput.Value = strLine
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
