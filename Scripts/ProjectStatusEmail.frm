VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectStatusEmail 
   Caption         =   "Project Status Email"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490.001
   OleObjectBlob   =   "ProjectStatusEmail.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectStatusEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddRouteInfo_Click()
    Me.Hide
    Load GetCableLengths
        GetCableLengths.show
        
        tbBody.Value = tbBody.Value & vbCr & vbCr & "<b>Layer:</b>" & vbTab & GetCableLengths.lbLayers.Value
        tbBody.Value = tbBody.Value & vbCr & "<b>Route Miles =</b>" & vbTab & GetCableLengths.tbTotalMiles.Value
        tbBody.Value = tbBody.Value & vbCr & "<b>Route KF =</b>" & vbTab & GetCableLengths.tbTotalFeet.Value
    Unload GetCableLengths
    Me.show
End Sub

Private Sub cbAddRuBu_Click()
    Me.Hide
    Load LoadBlocks
        LoadBlocks.show
        
        tbBody.Value = tbBody.Value & vbCr & vbCr & "<b>RU =</b>" & vbTab & LoadBlocks.tbRES.Value
        tbBody.Value = tbBody.Value & vbCr & "<b>BU =</b>" & vbTab & LoadBlocks.tbBUS.Value
        tbBody.Value = tbBody.Value & vbCr & "<b>SG =</b>" & vbTab & LoadBlocks.tbSG.Value
    Unload LoadBlocks
    Me.show
End Sub

Private Sub cbCancel_Click()
    tbBody.Value = ""
    Me.Hide
End Sub

Private Sub cbConvert_Click()
    Dim strText As String
    
    strText = tbBody.Value
    strText = Replace(strText, "<br>", vbCr)
    
    tbBody.Value = strText
End Sub

Private Sub cbSend_Click()
    Me.Hide
End Sub

Private Sub lbCC_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lbCC.ListCount < 0 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            lbCC.RemoveItem lbCC.ListIndex
    End Select
End Sub

Private Sub lbEmails_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    lbCC.AddItem lbEmails.List(lbEmails.ListIndex)
End Sub

Private Sub UserForm_Initialize()
    lbCC.Clear
    lbEmails.AddItem "Adam.Kemper@Integrity-US.com"
    'lbEmails.AddItem "Alex.Skae@Integrity-US.com"
    lbEmails.AddItem "Byron.Auer@Integrity-US.com"
    lbEmails.AddItem "Daniel.Campbell@Integrity-US.com"
    lbEmails.AddItem "Drew.Curtis@Integrity-US.com"
    lbEmails.AddItem "Dylan.Spears@Integrity-US.com"
    lbEmails.AddItem "Franklin.Angulo@Integrity-US.com"
    lbEmails.AddItem "Jason.Pafford@Integrity-US.com"
    lbEmails.AddItem "Jay.Penny@Integrity-US.com"
    lbEmails.AddItem "Jeremy.Pafford@Integrity-US.com"
    lbEmails.AddItem "Jon.Wilburn@Integrity-US.com"
    lbEmails.AddItem "Matt.Snyder@Integrity-US.com"
    lbEmails.AddItem "Rich.Taylor@Integrity-US.com"
    lbEmails.AddItem "Ronn.Elliott@Integrity-US.com"
    lbEmails.AddItem "Sam.Jackson@Integrity-US.com"
    lbEmails.AddItem "Tara.Taylor@Integrity-US.com"
    lbEmails.AddItem "Wade.Hampton@Integrity-US.com"
End Sub
