VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateNESPoles 
   Caption         =   "Create NES Poles"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9480.001
   OleObjectBlob   =   "CreateNESPoles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateNESPoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbCompany_Change()
    If tbPaste.Value = "" Then Exit Sub
    
    If cbCompany.Value = "NES" Then
        Call GetNESData
    Else
        Call GetMTEData
    End If
End Sub

Private Sub cbPlace_Click()
    If tbCoords.Value = "" Then Exit Sub
    
    Dim obrNES As AcadBlockReference
    Dim vAttList, vTemp As Variant
    Dim dCoords(0 To 2) As Double
    Dim dPoint(2) As Double
    Dim vNE As Variant
    Dim iCount As Integer
    
    vTemp = Split(tbCoords.Value, ", ")
    vNE = LLtoTN83F(CDbl(vTemp(0)), CDbl(vTemp(1)))
    dCoords(0) = vNE(1)
    dCoords(1) = vNE(0)
    dCoords(2) = 0#
    
    iCount = 16
    
    Set obrNES = ThisDrawing.ModelSpace.InsertBlock(dCoords, "sPole", 1#, 1#, 1#, 0#)
    vAttList = obrNES.GetAttributes
    
    vAttList(0).TextString = tbOwner.Value & " POLE"
    If Not tbType.Value = "" Then vAttList(1).TextString = tbType.Value
    vAttList(2).TextString = tbOwner.Value
    vAttList(3).TextString = tbNumber.Value
    'vAttList(25).TextString = tbID.Value
    vAttList(5).TextString = tbHC.Value
    vAttList(7).TextString = tbCoords.Value
    
    Select Case cbCompany.Value
        Case "NES"
            If tbOwner.Value = "NES" Then
                obrNES.Layer = "Integrity Poles-Power"
            Else
                obrNES.Layer = "Integrity Poles-Other"
            End If
        Case Else
            If InStr(tbOwner.Value, "MTE") > 0 Then
                obrNES.Layer = "Integrity Poles-Power"
            Else
                obrNES.Layer = "Integrity Poles-Other"
            End If
    End Select
    
    If lbCOMMs.ListCount > 0 Then
        For i = 0 To lbCOMMs.ListCount - 1
            vAttList(iCount).TextString = lbCOMMs.List(i) & "="
            iCount = iCount + 1
        Next i
    End If
    
    'vAttList(16).TextString = "CLIENT="
    
    'If lbCOMMs.ListCount > 0 Then
        
    'End If
    
    obrNES.Update
End Sub

Private Sub cbQuit_Click()
    ThisDrawing.SendCommand "_QSAVE" & vbCr
    
    Me.Hide
End Sub

Private Sub LabelPan_Click()
    Dim entPole As AcadObject
    Dim basePnt As Variant
    
    Me.Hide
  On Error Resume Next
    
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Left Click to Exit: "
    
    Me.show
End Sub

Private Sub tbPaste_Change()
    If tbPaste.Value = "" Then Exit Sub
    
    If cbCompany.Value = "NES" Then
        Call GetNESData
    Else
        Call GetMTEData
    End If
    
    Exit Sub
    
    Dim strAll, strTemp As String
    Dim vList As Variant
    Dim iCOMMs As Integer
    
    strAll = Replace(tbPaste.Value, vbLf, "")
    vList = Split(strAll, vbCr)
    iCOMMs = 0
    
    For i = 0 To UBound(vList)
        Select Case vList(i)
            Case "Location"
                tbCoords.Value = vList(i + 1)
            Case "Pole Number"
                tbNumber.Value = vList(i + 1)
            Case "Class"
                strTemp = vList(i + 3) & "-" & vList(i + 1)
                tbHC.Value = strTemp
            Case "Owner"
                tbOwner.Value = vList(i + 1)
            Case "Pole Oid"
                tbID.Value = vList(i + 1)
            Case "Type"
                If iCOMMs = 0 Then tbType.Value = vList(i + 1)
            Case "Asset Attachments"
                iCOMMs = 1
                i = i + 1
                lbCOMMs.AddItem vList(i)
                i = i + 4
            Case Else
                If iCOMMs = 1 Then
                    For j = 0 To lbCOMMs.ListCount - 1
                        If lbCOMMs.List(j) = vList(i) Then GoTo Found_COMM
                    Next j
                    
                    lbCOMMs.AddItem vList(i)
Found_COMM:
                    i = i + 4
                End If
        End Select
    Next i
End Sub

Private Sub tbPaste_Enter()
   tbPaste.Value = ""
    lbCOMMs.Clear
End Sub

Private Sub tbPaste_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    tbPaste.Value = ""
    lbCOMMs.Clear
End Sub

Private Sub UserForm_Initialize()
    cbCompany.AddItem "NES"
    cbCompany.AddItem "MTE"
    cbCompany.Value = "NES"
End Sub

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
    
    dK = 0.999948401424 + (1.23188E-14 * dU * dU) + (4.54E-22 * dU * dU * dU)
    
    NE(0) = dNorth
    NE(1) = dEast
    NE(2) = dK
    
    LLtoTN83F = NE
End Function

Private Sub GetNESData()
    If tbPaste.Value = "" Then Exit Sub
    
    Dim strAll, strTemp As String
    Dim vList As Variant
    Dim iCOMMs As Integer
    
    strAll = Replace(tbPaste.Value, vbLf, "")
    vList = Split(strAll, vbCr)
    iCOMMs = 0
    
    For i = 0 To UBound(vList)
        Select Case vList(i)
            Case "Location"
                tbCoords.Value = vList(i + 1)
            Case "Pole Number"
                tbNumber.Value = vList(i + 1)
            Case "Class"
                strTemp = vList(i + 3) & "-" & vList(i + 1)
                tbHC.Value = strTemp
            Case "Owner"
                tbOwner.Value = vList(i + 1)
            Case "Pole Oid"
                tbID.Value = vList(i + 1)
            Case "Type"
                If iCOMMs = 0 Then tbType.Value = vList(i + 1)
            Case "Asset Attachments"
                iCOMMs = 1
                i = i + 1
                lbCOMMs.AddItem vList(i)
                i = i + 4
            Case Else
                If iCOMMs = 1 Then
                    For j = 0 To lbCOMMs.ListCount - 1
                        If lbCOMMs.List(j) = vList(i) Then GoTo Found_COMM
                    Next j
                    
                    lbCOMMs.AddItem vList(i)
Found_COMM:
                    i = i + 4
                End If
        End Select
    Next i
End Sub

Private Sub GetMTEData()
    If tbPaste.Value = "" Then Exit Sub
    
    Dim strAll, strTemp As String
    Dim vList As Variant
    Dim iCOMMs As Integer
    
    strAll = Replace(tbPaste.Value, vbLf, "")
    vList = Split(strAll, vbCr)
    iCOMMs = 0
    
    For i = 0 To UBound(vList)
        Select Case vList(i)
            Case "Location"
                tbCoords.Value = vList(i + 1)
            Case "Pole Number"
                tbNumber.Value = vList(i + 1)
            Case "Class"
                strTemp = vList(i + 3) & "-" & vList(i + 1)
                tbHC.Value = strTemp
            Case "Ownership Code"
                tbOwner.Value = vList(i + 1)
            Case "Fid Number"
                tbID.Value = vList(i + 1)
        End Select
    Next i
    
    tbType.Value = "INCOMPLETE"
End Sub
