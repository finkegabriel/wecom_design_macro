VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateVHLEForms 
   Caption         =   "Create/Update VHLE Forms"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "CreateVHLEForms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateVHLEForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dProjectCost As Double

Private Sub cbPOR_Click()
    If cbPOR.Value = True Then cbVHLE.Value = True
End Sub

Private Sub cbRun_Click()
    If cbVHLE.Value = True Then
        Call FilloutVHLE
        If cbPOR.Value = True Then Call FilloutPORequest
    End If
    If cbMSAT.Value = True Then Call FilloutMSAT
End Sub

Private Sub cbVHLE_Click()
    If cbVHLE.Value = False Then cbPOR.Value = False
End Sub

Private Sub Label10_Click()
    Dim dTotal As Double
    
    Me.Hide
    Load GetCableLengths
        GetCableLengths.show
        
        dTotal = CDbl(GetCableLengths.tbTotalFeet.Value) * 1000
        tbBP.Value = dTotal
    Unload GetCableLengths
    Me.show
End Sub

Private Sub Label4_Click()
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
    
    tbRU.Value = iRes
    tbBU.Value = iBus
    tbSG.Value = iSG
    
    iRes = iRes + iBus + iSG
    tbCTotal.Value = iRes
    
    Me.show
End Sub

Private Sub Label7_Click()
    Dim dTotal As Double
    
    Me.Hide
    Load GetCableLengths
        GetCableLengths.show
        
        dTotal = CDbl(GetCableLengths.tbTotalFeet.Value) * 1000
        tbAS.Value = dTotal
    Unload GetCableLengths
    Me.show
End Sub

Private Sub Label8_Click()
    Dim dTotal As Double
    
    Me.Hide
    Load GetCableLengths
        GetCableLengths.show
        
        dTotal = CDbl(GetCableLengths.tbTotalFeet.Value) * 1000
        tbAL.Value = dTotal
    Unload GetCableLengths
    Me.show
End Sub

Private Sub Label9_Click()
    Dim dTotal As Double
    
    Me.Hide
    Load GetCableLengths
        GetCableLengths.show
        
        dTotal = CDbl(GetCableLengths.tbTotalFeet.Value) * 1000
        tbBC.Value = dTotal
    Unload GetCableLengths
    Me.show
End Sub

Private Sub UserForm_Initialize()
    cbType.AddItem "FEEDER"
    cbType.AddItem "DISTRIBUTION"
    cbType.AddItem "RURAL DISTRIBUTION"
    cbType.AddItem "DROP-AERIAL"
    cbType.AddItem "DROP-UG"
    
    cbLocation.AddItem "ILEC"
    cbLocation.AddItem "CLEC"
    
    cbProject.AddItem "FTTH"
    cbProject.AddItem "FTTB"
    cbProject.AddItem "FTTT"
    cbProject.AddItem "FFIB"
    
    cbField.AddItem ""
    cbField.AddItem "GF"
    cbField.AddItem "GF-SG"
    cbField.AddItem "OVERBUILD"
    cbField.AddItem "RURAL OVERBUILD"
    
    cbSG.AddItem "YES"
    cbSG.AddItem "NO"
    
    dProjectCost = 0#
    
    Dim strLine As String
    Dim vLine As Variant
    
    strLine = Replace(UCase(ThisDrawing.Name), ".DWG", "")
    vLine = Split(strLine, " ")
    tbNumber.Value = vLine(0)
    
    If UBound(vLine) > 0 Then
        strLine = vLine(1)
        If UBound(vLine) > 1 Then
            For i = 2 To UBound(vLine)
                strLine = strLine & " " & vLine(i)
            Next i
        End If
        
        tbDescription.Value = strLine
    End If
End Sub

Private Sub FilloutVHLE()
    Dim strExistingName, strFileName As String
    Dim strTemp, strFormName As String
    Dim vTemp, vTemp2, vLine As Variant
    Dim fName, strLocal As String
    Dim objExcel As Workbook
    Dim strResult As String
    
    'vTemp = Split(ThisDrawing.Name, " ")
    'strExistingName = ThisDrawing.Path & "\VHLE\" & vTemp(0) & " VHLE.xlsx"
    strExistingName = ThisDrawing.Path & "\" & tbNumber.Value & " " & tbDescription.Value & " VHLE.xlsx"
    strLocal = "C:\Integrity\VBA\Forms\"
    
    fName = Dir(strExistingName)
    If fName = "" Then
        strTemp = ThisDrawing.Path
        If InStr(LCase(strTemp), "dropbox") > 0 Then
            vTemp2 = Split(LCase(strTemp), "dropbox")
            strFormName = vTemp2(0) & "\Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\Blank VHLE.xlsx"
        Else
            strFormName = strLocal & "Blank VHLE.xlsx"
        End If
        
        Set objExcel = Workbooks.Open(strFormName)
        'MsgBox strExistingName
        objExcel.SaveAs (strExistingName)
        MsgBox "Created New VHLE."
    Else
        Set objExcel = Workbooks.Open(strExistingName)
        'MsgBox "Opened Existing File"
    End If
    
    objExcel.Sheets("VHLE").Cells(2, 4).Value = tbNumber.Value
    objExcel.Sheets("VHLE").Cells(2, 5).Value = tbDescription.Value
    
    If cbProject.Value = "FFIB" Then
        objExcel.Sheets("VHLE").Cells(2, 7).Value = "FEEDER"
    Else
        Select Case cbField.Value
            Case "RURAL OVERBUILD"
                objExcel.Sheets("VHLE").Cells(2, 7).Value = "RURAL DISTRIBUTION"
            Case Else
                objExcel.Sheets("VHLE").Cells(2, 7).Value = "DISTRIBUTION"
        End Select
    End If
    
    'objExcel.Sheets("VHLE").Cells(2, 7).Value = cbType.Value
    objExcel.Sheets("VHLE").Cells(2, 8).Value = tbRU.Value
    objExcel.Sheets("VHLE").Cells(2, 9).Value = tbBU.Value
    objExcel.Sheets("VHLE").Cells(2, 10).Value = tbSG.Value
    objExcel.Sheets("VHLE").Cells(2, 19).Value = tbAS.Value
    objExcel.Sheets("VHLE").Cells(2, 21).Value = tbAL.Value
    objExcel.Sheets("VHLE").Cells(2, 23).Value = tbBC.Value
    objExcel.Sheets("VHLE").Cells(2, 25).Value = tbBP.Value
    
    dProjectCost = CDbl(objExcel.Sheets("VHLE").Cells(2, 40))
        
    objExcel.Save
    objExcel.Close
    
    MsgBox "VHLE form has been updated."
End Sub

Private Sub FilloutPORequest()
    Dim strExistingName, strFileName As String
    Dim strTemp, strFormName As String
    Dim vTemp, vTemp2, vLine As Variant
    Dim fName, strLocal As String
    Dim objExcel As Workbook
    
    'vTemp = Split(ThisDrawing.Name, " ")
    'strExistingName = ThisDrawing.Path & "\VHLE\" & vTemp(0) & " VHLE.xlsx"
    strExistingName = ThisDrawing.Path & "\" & tbNumber.Value & " " & tbDescription.Value & " ICS PO Request.xlsx"
    strLocal = "C:\Integrity\VBA\Forms\"
    
    fName = Dir(strExistingName)
    If fName = "" Then
        strTemp = ThisDrawing.Path
        If InStr(LCase(strTemp), "dropbox") > 0 Then
            vTemp2 = Split(LCase(strTemp), "dropbox")
            strFormName = vTemp2(0) & "\Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\Blank PO Request.xlsx"
        Else
            strFormName = strLocal & "Blank PO Request.xlsx"
        End If
        
        Set objExcel = Workbooks.Open(strFormName)
        'MsgBox strExistingName
        objExcel.SaveAs (strExistingName)
        MsgBox "Created New P.O. Request."
    Else
        Set objExcel = Workbooks.Open(strExistingName)
        'MsgBox "Opened Existing File"
    End If
    
    objExcel.Sheets("ICS").Cells(4, 2).Value = Date
    objExcel.Sheets("ICS").Cells(7, 2).Value = tbNumber.Value
    objExcel.Sheets("ICS").Cells(8, 2).Value = tbCTotal.Value
    objExcel.Sheets("ICS").Cells(9, 2).Value = dProjectCost
    objExcel.Sheets("ICS").Cells(15, 2).Value = tbAS.Value
    objExcel.Sheets("ICS").Cells(16, 2).Value = tbAL.Value
    objExcel.Sheets("ICS").Cells(17, 2).Value = tbBC.Value
    objExcel.Sheets("ICS").Cells(18, 2).Value = tbBP.Value
        
    objExcel.Save
    objExcel.Close
    
    MsgBox "P.O. Request form has been updated."
End Sub

Private Sub FilloutMSAT()
    Dim strExistingName, strFileName As String
    Dim strTemp, strFormName, strLine As String
    Dim vTemp, vTemp2, vLine As Variant
    Dim fName As String
    Dim book1 As Word.Application
    Dim sheet1 As Word.Document
    
    strExistingName = ThisDrawing.Path & "\" & tbNumber.Value & " " & tbDescription.Value & " COVER SHEET.docx"
    
    fName = Dir(strExistingName)
    If fName = "" Then
        strTemp = ThisDrawing.Path
        If InStr(LCase(strTemp), "dropbox") > 0 Then
            vTemp2 = Split(LCase(strTemp), "dropbox")
            strFormName = vTemp2(0) & "\Dropbox\UNITED COMMUNICATIONS JOB INFORMATION\VBA\Integrity\VBA\Forms\Blank COVER SHEET.docx"
        Else
            strFormName = "C:\Integrity\VBA\Forms\Blank COVER SHEET.docx"
        End If
        
        FileCopy strFormName, strExistingName
        MsgBox "Created New COVER SHEET."
    End If
    
    Set book1 = CreateObject("word.application")
    book1.Visible = True
    Set sheet1 = book1.Documents.Open(strExistingName)

    With sheet1.Content.Paragraphs(1).Range.Find
        .Text = "<<blank date>>"
        .Replacement.Text = Date
        '.Forward = True
        .Wrap = wdFindContinue
        '.Format = False
        '.MatchCase = False
        .Execute Replace:=wdReplaceAll
        
        .Text = "<<blank project>>"
        .Replacement.Text = tbNumber.Value & " " & tbDescription.Value
        '.Forward = True
        .Wrap = wdFindContinue
        '.Format = False
        '.MatchCase = False
        .Execute Replace:=wdReplaceAll
        
        strLine = cbLocation.Value & " " & cbProject.Value & " " & cbField.Value
        .Text = "<<blank type>>"
        .Replacement.Text = strLine
        '.Forward = True
        .Wrap = wdFindContinue
        '.Format = False
        '.MatchCase = False
        .Execute Replace:=wdReplaceAll
        
        If Not tbSource.Value = "" Then
            .Text = "<<blank source>>"
            .Replacement.Text = tbSource.Value
            '.Forward = True
            .Wrap = wdFindContinue
            '.Format = False
            '.MatchCase = False
            .Execute Replace:=wdReplaceAll
        End If
        
        If Not tbFeeder.Value = "" Then
            .Text = "<<blank feeder>>"
            .Replacement.Text = tbFeeder.Value
            '.Forward = True
            .Wrap = wdFindContinue
            '.Format = False
            '.MatchCase = False
            .Execute Replace:=wdReplaceAll
        End If
        
        If Not tbScope.Value = "" Then
            .Text = "<<blank scope>>"
            .Replacement.Text = tbScope.Value
            '.Forward = True
            .Wrap = wdFindContinue
            '.Format = False
            '.MatchCase = False
            .Execute Replace:=wdReplaceAll
        End If
    End With
    
    sheet1.Save
    sheet1.Close
    book1.Quit
    
    MsgBox "MSAT Cover Sheet has been updated."
End Sub
