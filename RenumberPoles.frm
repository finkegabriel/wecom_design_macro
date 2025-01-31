VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenumberPoles 
   Caption         =   "Renumber Poles"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   OleObjectBlob   =   "RenumberPoles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenumberPoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbBlocks_Click()
    Dim objSS As AcadSelectionSet
    Dim objEntity As AcadEntity
    Dim objBlock As AcadBlockReference
    Dim vAttList, vList As Variant
    Dim str1, strList As String
    Dim strNumber As String
    
    Me.Hide
    strNumber = tbRoute.Value & tbPole.Value
    'strList = ""
    
  On Error Resume Next
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        Err = 0
    End If
    
    objSS.SelectOnScreen
    
    For Each objEntity In objSS
        If TypeOf objEntity Is AcadBlockReference Then
            Set objBlock = objEntity
            vAttList = objBlock.GetAttributes
            
            Select Case objBlock.Name
                Case "iPole", "sPole", "sPed", "sHH", "pole_attach", "pole_attach_title", "pole_unit"
                    vAttList(0).TextString = strNumber
                Case "pole_info"
                    vAttList(0).TextString = strNumber
                    
                    If vAttList(1).TextString = "0.0" Then
                        vAttList(2).TextString = strNumber
                    End If
            End Select
            
            objBlock.Update
        End If
    Next objEntity
    
    tbPole.Value = CInt(tbPole.Value) + 1
    
    objSS.Clear
    objSS.Delete
    Me.show
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbRenumPole_Click()
    Dim entPole As AcadObject
    Dim entObj As AcadEntity
    Dim obrPole As AcadBlockReference
    Dim vList As Variant
    Dim str1 As String
    Dim iNum As Integer
    Dim dN, dE As Double
    Dim vLL As Variant
    Dim iPlus As Integer
    Dim iChr As Integer
    
    
    Me.Hide
    iPlus = CInt(tbIncrement.Value)
    iNum = CInt(tbPole.Value) - iPlus
    iChr = 65
    
  On Error Resume Next
    While Err = 0
        ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Pole: "
        If Not Err = 0 Then GoTo Exit_Sub
        If TypeOf entPole Is AcadBlockReference Then
            Set obrPole = entPole
        Else
            tbPole.Value = iNum
            Me.show
            Exit Sub
        End If
        
        'MsgBox obrPole.Name
        
        Select Case obrPole.Name
            Case "sPole"
                iNum = iNum + iPlus
                iChr = 65
                
                vList = obrPole.GetAttributes
                vList(0).TextString = tbRoute.Value & iNum
        
                If cbLL.Value = True Then
                    If vList(7).TextString = "" Then
                        vLL = TN83FtoLL(CDbl(obrPole.InsertionPoint(0)), CDbl(obrPole.InsertionPoint(1)))
                        vList(7).TextString = vLL(0) & "," & vLL(1)
                    End If
                End If
            Case "sPed", "sHH", "sPanel", "sMH"
                iNum = iNum + iPlus
                iChr = 65
                
                vList = obrPole.GetAttributes
                vList(0).TextString = tbRoute.Value & iNum
        
                If cbLL.Value = True Then
                    If vList(3).TextString = "" Then
                        vLL = TN83FtoLL(CDbl(obrPole.InsertionPoint(0)), CDbl(obrPole.InsertionPoint(1)))
                        vList(3).TextString = vLL(0) & "," & vLL(1)
                    End If
                End If
            Case "sFP"
                vList = obrPole.GetAttributes
                vList(0).TextString = tbRoute.Value & iNum & Chr(iChr)
                vList(1).TextString = iNum & Chr(iChr)
                iChr = iChr + 1
        
                If cbLL.Value = True Then
                    If vList(3).TextString = "" Then
                        vLL = TN83FtoLL(CDbl(obrPole.InsertionPoint(0)), CDbl(obrPole.InsertionPoint(1)))
                        vList(3).TextString = vLL(0) & "," & vLL(1)
                    End If
                End If
        End Select
        
        obrPole.Update
    Wend
Exit_Sub:
    tbPole.Value = iNum + iPlus
    Me.show
End Sub

Private Sub Label1_Click()
    Dim entPole As AcadObject
    Dim entObj As AcadEntity
    Dim obrPole As AcadBlockReference
    Dim vList As Variant
    Dim str1 As String
    Dim iNum As Integer
    
    Me.Hide
    
  On Error Resume Next
    ThisDrawing.Utility.GetEntity entPole, basePnt, "Select Pole: "
    If TypeOf entPole Is AcadBlockReference Then
        Set obrPole = entPole
    Else
        Me.show
        Exit Sub
    End If
        
    vList = obrPole.GetAttributes
    tbRoute.Value = vList(0).TextString
    tbPole.Value = "1"
    
    Me.show
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
