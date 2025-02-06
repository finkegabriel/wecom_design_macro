VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReplaceUnits 
   Caption         =   "Replace Units"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "ReplaceUnits.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReplaceUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAddToList_Click()
    If tbFind.Value = "" Then Exit Sub
    If tbReplace.Value = "" Then Exit Sub
    
    Dim strLine As String
    Dim vLine, vItem As Variant
    Dim strType, strTest As String
    
    strLine = Replace(UCase(tbReplace.Value), vbLf, "")
    vLine = Split(strLine, vbCr)
    For i = 0 To UBound(vLine)
        If vLine(i) = "" Then GoTo Next_line
        If Not Left(vLine(i), 1) = "+" Then vLine(i) = "+" & vLine(i)
        
        If InStr(vLine(i), "=") > 0 Then
            vItem = Split(vLine(i), "=")
            
            If Left(vItem(1), 1) = " " Then
                vLine(i) = vItem(0) & "={A} " & vItem(1)
            'Else
                'vLine(i) = vItem(0) & "=" & vItem(1)
            End If
        Else
            vLine(i) = vLine(i) & "=1"
        End If
        
Next_line:
    Next i
    
    strLine = ""
    For i = 0 To UBound(vLine)
        If vLine(i) = "" Then GoTo Next_Unit
        
        If strLine = "" Then
            strLine = vLine(i)
        Else
            strLine = strLine & ";;" & vLine(i)
        End If
        
Next_Unit:
    Next i
    
    If InStr(tbFind.Value, "+") > 0 Then
        strType = "f"
    Else
        strType = "p"
    End If
    
    If InStr(tbFind.Value, "= ") > 0 Then
        strType = strType & "f"
    Else
        strType = strType & "p"
    End If
    
    strTest = Left(tbFind.Value, 1)
    If Not strTest = "=" And Not strTest = "+" Then strLine = "+" & strLine
    
    lbList.AddItem UCase(tbFind.Value)
    lbList.List(lbList.ListCount - 1, 1) = strLine
    lbList.List(lbList.ListCount - 1, 2) = strType
    'MsgBox strType
End Sub

Private Sub cbClearText_Click()
    tbFind.Value = ""
    tbReplace.Value = ""
End Sub

Private Sub cbHelp_Click()
    If cbHelp.Caption = "Help" Then
        Me.Height = 366
        
        cbHelp.Caption = "Help Off"
    Else
        Me.Height = 234
        
        cbHelp.Caption = "Help"
    End If
End Sub

Private Sub cbReplace_Click()
    Dim vDwgLL, vDwgUR As Variant
    Dim filterType, filterValue As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0) As Variant
    Dim objSS As AcadSelectionSet
    Dim strType, strFind, strReplace As String
    Dim strStructures, strExtra, strSize As String
    Dim strAmount, strUnit As String
    Dim vFind, vReplace As Variant
    Dim vList, vUnit, vExtra As Variant
    Dim vLine, vItem, vTemp As Variant
    Dim iPos As Integer
    
    On Error Resume Next
    
    strStructures = ""
    If cbPoles.Value = True Then strStructures = "sPole"
    If cbPEDs.Value = True Then
        If strStructures = "" Then
            strStructures = "sPED"
        Else
            strStructures = strStructures & ",sPED"
        End If
    End If
    If cbHHs.Value = True Then
        If strStructures = "" Then
            strStructures = "sHH"
        Else
            strStructures = strStructures & ",sHH"
        End If
    End If
    If cbPanels.Value = True Then
        If strStructures = "" Then
            strStructures = "sPanel"
        Else
            strStructures = strStructures & ",sPanel"
        End If
    End If
    If cbFPs.Value = True Then
        If strStructures = "" Then
            strStructures = "sFP"
        Else
            strStructures = strStructures & ",sFP"
        End If
    End If
    
    Me.Hide
    
    vDwgLL = ThisDrawing.Utility.GetPoint(, "Get DWG LL Corner: ")
    vDwgUR = ThisDrawing.Utility.GetCorner(vDwgLL, vbCr & "Get DWG UR Corner: ")
    If Not Err = 0 Then
        Me.show
        Exit Sub
    End If
    
    grpCode(0) = 2
    grpValue(0) = strStructures
    filterType = grpCode
    filterValue = grpValue
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
    End If
    objSS.Select acSelectionSetWindow, vDwgLL, vDwgUR, filterType, filterValue
    If objSS.count < 1 Then GoTo Exit_Sub
    
    MsgBox objSS.count & "  " & strStructures & " found."
    
    'GoTo Exit_Sub
    
    For Each objBlock In objSS
        vAttList = objBlock.GetAttributes
        
        Select Case objBlock.Name
            Case "sPole"
                If vAttList(27).TextString = "" Then GoTo Next_objBlock
                vList = Split(vAttList(27).TextString, ";;")
            Case Else
                If vAttList(7).TextString = "" Then GoTo Next_objBlock
                vList = Split(vAttList(7).TextString, ";;")
        End Select
        
        For i = 0 To UBound(vList)
            strUnit = vList(i)
            
            For j = 0 To lbList.ListCount - 1
                vFind = Split(lbList.List(j, 0), "=")
                
                'Check Left of =
                If InStr(vFind(0), "()") > 0 Then
                    vItem = Split(vFind(0), "()")
                    If InStr(strUnit, vItem(0)) > 0 Then
                        If InStr(strUnit, vItem(1)) < 1 Then GoTo Next_Find
                    Else
                        GoTo Next_Find
                    End If
                Else
                    If InStr(strUnit, vFind(0)) < 1 Then GoTo Next_Find
                End If
                
                'Check Right of = and get strAmount
                If InStr(vFind(1), "  ") < 1 Then
                    If InStr(vFind(1), " ") > 0 Then
                        vItem = Split(vFind(1), " ")
                        vFind(1) = Replace(vFind(1), vItem(0), vItem(0) & " ")
                        vItem = Split(vFind(1), "  ")
                    End If
                    
                    vUnit = Split(strUnit, "=")
                    vTemp = Split(vUnit(1), "  ")
                    If UBound(vTemp) < 1 Then GoTo Next_Find
                    
                    If InStr(vTemp(1), vItem(1)) < 1 Then GoTo Next_Find
                    strAmount = vTemp(0)
                Else
                    strAmount = vUnit(1)
                End If
                
                'FOUND UNIT: Now replace the units
                strReplace = Replace(lbList.List(j, 1), "{A}", strAmount)
                If InStr(strReplace, "()") > 0 Then
                    If InStr(strUnit, ")") > 0 Then
                        vLine = Split(strUnit, ")")
                        vTemp = Split(vLine(0), "(")
                        strSize = "(" & vTemp(1) & ")"
                        
                        strReplace = Replace(strReplace, "()", strSize)
                    End If
                End If
                
                vList(i) = strReplace
                
Next_Find:
            Next j
            
        Next i
        
        strLine = ""
        For i = 0 To UBound(vList)
            If Not vList(i) = "" Then
                If strLine = "" Then
                    strLine = vList(i)
                Else
                    strLine = strLine & ";;" & vList(i)
                End If
            End If
        Next i
        
        Select Case objBlock.Name
            Case "sPole"
                vAttList(27).TextString = strLine
            Case Else
                vAttList(7).TextString = strLine
        End Select
        
        objBlock.Update
        
Next_objBlock:
    Next objBlock
    
    
Exit_Sub:
    objSS.Clear
    objSS.Delete
    
    Me.show
End Sub

Private Sub tbFind_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If tbFind.Value = "" Then Exit Sub
    If Not Left(tbFind.Value, 1) = "+" Then tbFind.Value = "+" & tbFind.Value
    If InStr(tbFind.Value, "=") < 1 Then
        tbFind.Value = tbFind.Value & "={a}"
    Else
        If InStr(tbFind.Value, "= ") < 1 Then
            If InStr(tbFind.Value, "={") < 1 Then tbFind.Value = Replace(tbFind.Value, "=", "={A} ")
        End If
    End If
        
    If tbReplace.Value = "" Then
        tbReplace.Value = tbFind.Value
        Exit Sub
    End If
    
    If InStr(tbFind.Value, "(") < 1 Then Exit Sub
    If InStr(tbReplace.Value, "(") < 1 Then Exit Sub
    
    Dim vLine, vItem, vTemp As Variant
    Dim strLine, strSize, strFind, strReplace As String
    
    vLine = Split(tbFind.Value, "(")
    vItem = Split(vLine(1), ")")
    'strSize = vItem(0)
    strReplace = "(" & vItem(0) & ")"
    
    vTemp = Split(tbReplace.Value, vbCr)
    
    vLine = Split(vTemp(0), "(")
    vItem = Split(vLine(1), ")")
    strFind = "(" & vItem(0) & ")"
    
    strLine = vLine(0) & strReplace & vItem(1)
    If UBound(vTemp) > 0 Then
        For i = 1 To UBound(vTemp)
            vTemp(i) = Replace(vTemp(i), strFind, strReplace)
            strLine = strLine & vbCr & vTemp(i)
        Next i
    End If
    
    tbReplace.Value = strLine
End Sub

Private Sub UserForm_Initialize()
    lbList.ColumnCount = 3
    lbList.ColumnWidths = "84;180;6"
    
    Dim strLine As String
    
    strLine = "Enter the unit name you want to find.  It is not case senitive.  If you don't start with a plus symbol, it will be added for you."
    strLine = strLine & "  If you want to find all closures or cables placed in the same way, do not put a size inbetween the parenthesis."
    strLine = strLine & "  If you need to find something with a description add an equal sign and a space then enter the description."
    strLine = strLine & vbCr & vbCr & "Find Examples:" & vbCr & "+haco(12)" & vbCr & "haco()= g5"
    strLine = strLine & vbCr & "+co()e" & vbCr & "+co()e= loop" & vbCr & vbCr
    strLine = strLine & ""
    strLine = strLine & "Enter the replacement unit first then any other units that need to added.  Seperated the units on different lines."
    strLine = strLine & "  If the Replace text box is empty, it will copy the Find text into the box."
    strLine = strLine & "  If you want the amount of any units to be equal to a know number add an equal sign and the number it will be."
    strLine = strLine & "  If you need to add a description, make sure there is an equal sign and after the known amount, if any, add the description."
    strLine = strLine & "  The plus symbol is not necessary on the any of the replace units."
    strLine = strLine & ""
    strLine = strLine & ""
    strLine = strLine & ""
    
    tbHelp.Value = strLine
End Sub
