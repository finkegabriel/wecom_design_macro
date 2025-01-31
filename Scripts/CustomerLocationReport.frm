VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerLocationReport 
   Caption         =   "Customer Location Report"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9375.001
   OleObjectBlob   =   "CustomerLocationReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustomerLocationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCreateReport_Click()
    If lbList.ListCount < 1 Then Exit Sub
    
    Dim strPath As String
    
    strPath = ThisDrawing.Path & "\Customer Location Report.csv"
    Open strPath For Output As #1
    
    Print #1, "Hse #,Street Name,Latitude,Longitude,Type"
    For i = 0 To lbList.ListCount - 1
        strLine = lbList.List(i, 0) & "," & lbList.List(i, 1) & "," & lbList.List(i, 2) & "," & lbList.List(i, 3) & "," & lbList.List(i, 4)
        Print #1, strLine
    Next i
    
    Close #1
End Sub

Private Sub cbGetCustomers_Click()
    Dim objSS As AcadSelectionSet
    Dim objLWP As AcadLWPolyline
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vReturnPnt As Variant
    
    Dim vCoords As Variant
    Dim vNE, vLL As Variant
    Dim iTemp, iCounter As Integer
    
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
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    objSS.SelectByPolygon acSelectionSetWindowPolygon, dCoords
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadBlockReference Then
            Set objBlock = objEntity
            If objBlock.Name = "Customer" Then
                vAttList = objBlock.GetAttributes
                
                If Not vAttList(1).TextString = "" Then
                    lbList.AddItem vAttList(1).TextString
                    lbList.List(lbList.ListCount - 1, 1) = vAttList(2).TextString
                    
                    vLL = TN83FtoLL(CDbl(objBlock.InsertionPoint(1)), CDbl(objBlock.InsertionPoint(0)))
                    lbList.List(lbList.ListCount - 1, 2) = vLL(0)
                    lbList.List(lbList.ListCount - 1, 3) = vLL(1)
                    lbList.List(lbList.ListCount - 1, 4) = vAttList(0).TextString
                    
                End If
            End If
        End If
    Next objEntity
    
    tbListCount = lbList.ListCount
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub cbSort_Click()
    If lbList.ListCount < 2 Then Exit Sub
    
    Call SortList
End Sub

Private Sub UserForm_Initialize()
    lbList.ColumnCount = 5
    lbList.ColumnWidths = "60;144;96;96;54"
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

Private Sub SortList()
    Dim strTemp, strTotal As String
    Dim iCount, iOffset As Integer
    Dim iIndex As Integer
    Dim strAtt(0 To 4) As String
    
    iCount = lbList.ListCount - 1
    If iCount < 1 Then Exit Sub
    
    On Error Resume Next
    
    For a = iCount To 0 Step -1
        For b = 0 To a - 1
            If lbList.List(b, 1) > lbList.List(b + 1, 1) Then
                If Not Err = 0 Then
                    MsgBox "Error sorting list"
                    lbList.Selected(b) = True
                    lbList.ListIndex = b
                    Exit Sub
                End If
                
                strAtt(0) = lbList.List(b + 1, 0)
                strAtt(1) = lbList.List(b + 1, 1)
                strAtt(2) = lbList.List(b + 1, 2)
                strAtt(3) = lbList.List(b + 1, 3)
                strAtt(4) = lbList.List(b + 1, 4)
                
                lbList.List(b + 1, 0) = lbList.List(b, 0)
                lbList.List(b + 1, 1) = lbList.List(b, 1)
                lbList.List(b + 1, 2) = lbList.List(b, 2)
                lbList.List(b + 1, 3) = lbList.List(b, 3)
                lbList.List(b + 1, 4) = lbList.List(b, 4)
                
                lbList.List(b, 0) = strAtt(0)
                lbList.List(b, 1) = strAtt(1)
                lbList.List(b, 2) = strAtt(2)
                lbList.List(b, 3) = strAtt(3)
                lbList.List(b, 4) = strAtt(4)
            End If
            
            If lbList.List(b, 1) = lbList.List(b + 1, 1) Then
                If lbList.List(b, 0) > lbList.List(b + 1, 0) Then
                    strAtt(0) = lbList.List(b + 1, 0)
                    strAtt(1) = lbList.List(b + 1, 1)
                    strAtt(2) = lbList.List(b + 1, 2)
                    strAtt(3) = lbList.List(b + 1, 3)
                    strAtt(4) = lbList.List(b + 1, 4)
                
                    lbList.List(b + 1, 0) = lbList.List(b, 0)
                    lbList.List(b + 1, 1) = lbList.List(b, 1)
                    lbList.List(b + 1, 2) = lbList.List(b, 2)
                    lbList.List(b + 1, 3) = lbList.List(b, 3)
                    lbList.List(b + 1, 4) = lbList.List(b, 4)
                
                    lbList.List(b, 0) = strAtt(0)
                    lbList.List(b, 1) = strAtt(1)
                    lbList.List(b, 2) = strAtt(2)
                    lbList.List(b, 3) = strAtt(3)
                    lbList.List(b, 4) = strAtt(4)
                End If
            End If
        Next b
    Next a
End Sub
