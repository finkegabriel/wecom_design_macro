VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VerifyACAtributes 
   Caption         =   "Block Attributes"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10440
   OleObjectBlob   =   "VerifyACAtributes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VerifyACAtributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iPoleIndex, iCalloutIndex As Integer
Dim iList As Integer

Private Sub cbCalloutUpdate_Click()
    If Not lbCallout.List(iCalloutIndex, 0) = tbCompany.Value Then
        lbCallout.List(iCalloutIndex, 0) = tbCompany.Value
        lbCallout.List(iCalloutIndex, 4) = "Y"
    End If
    
    If Not lbCallout.List(iCalloutIndex, 1) = tbHeight.Value Then
        lbCallout.List(iCalloutIndex, 1) = tbHeight.Value
        lbCallout.List(iCalloutIndex, 4) = "Y"
    End If
    
    If Not lbCallout.List(iCalloutIndex, 2) = tbMR.Value Then
        lbCallout.List(iCalloutIndex, 2) = tbMR.Value
        lbCallout.List(iCalloutIndex, 4) = "Y"
    End If
    
    cbCalloutUpdate.Enabled = False
End Sub

Private Sub cbCancel_Click()
    lbPole.Clear
    lbCallout.Clear
    
    Me.Hide
End Sub

Private Sub cbPoleUpdate_Click()
    If Not lbPole.List(iPoleIndex, 1) = tbText.Value Then
        lbPole.List(iPoleIndex, 1) = tbText.Value
        lbPole.List(iPoleIndex, 2) = "Y"
    End If
    
    tbText.Value = ""
    cbPoleUpdate.Enabled = False
End Sub

Private Sub cbSave_Click()
    Me.Hide
End Sub

Private Sub lbCallout_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    iList = 1
    iCalloutIndex = lbCallout.ListIndex
    
    tbCompany.Value = lbCallout.List(iCalloutIndex, 0)
    tbHeight.Value = lbCallout.List(iCalloutIndex, 1)
    tbMR.Value = lbCallout.List(iCalloutIndex, 2)
    
    cbCalloutUpdate.Enabled = True
    tbCompany.SetFocus
End Sub

Private Sub lbPole_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    iList = 0
    iPoleIndex = lbPole.ListIndex
    tbText.Value = lbPole.List(iPoleIndex, 1)
    
    cbPoleUpdate.Enabled = True
    tbText.SetFocus
End Sub

Private Sub UserForm_Initialize()
    lbPole.ColumnCount = 3
    lbPole.ColumnWidths = "72;108;24"
    
    lbCallout.ColumnCount = 5
    lbCallout.ColumnWidths = "96;36;84;36;36"
End Sub
