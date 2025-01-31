VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidateChecklist 
   Caption         =   "Validate Checklist"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   OleObjectBlob   =   "ValidateChecklist.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidateChecklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCountsCallout_Click()
    Me.Hide
        Load PlaceCountCallouts
        PlaceCountCallouts.show
        Unload PlaceCountCallouts
    Me.show
End Sub

Private Sub cbExtraHeights_Click()
    Me.Hide
        Load ExtraHeightForm
        ExtraHeightForm.show
        Unload ExtraHeightForm
    Me.show
End Sub

Private Sub cbGetSpansReel_Click()
    Me.Hide
        Load GetSpansReel
        GetSpansReel.show
        Unload GetSpansReel
    Me.show
End Sub

Private Sub cbQuit_Click()
    Call SaveData
    Me.Hide
End Sub

Private Sub cbTabGeneric_Click()
    Me.Hide
        Load TabGeneric
            TabGeneric.show
        Unload TabGeneric
    Me.show
End Sub

Private Sub cbTDOT_Click()
    Me.Hide
        Load TDOT
            TDOT.show
        Unload TDOT
    Me.show
End Sub

Private Sub cbTransfer_Click()
    Me.Hide
        Load TransferToMap
        TransferToMap.show
        Unload TransferToMap
    Me.show
End Sub

Private Sub cbValidateCounts_Click()
    Me.Hide
        Load ValidateCounts
        ValidateCounts.show
        Unload ValidateCounts
    Me.show
End Sub

Private Sub cbValidateCustomers_Click()
    Me.Hide
        Load ValidateCustomers
        ValidateCustomers.show
        Unload ValidateCustomers
    Me.show
End Sub

Private Sub cbValidateHO1_Click()
    Me.Hide
        Load ValidateHO1
        ValidateHO1.show
        Unload ValidateHO1
    Me.show
End Sub

Private Sub cbValidateMCL_Click()
    Me.Hide
        Load ValidateMCL
        ValidateMCL.show
        Unload ValidateMCL
    Me.show
End Sub

Private Sub cbValidateML_Click()
    Me.Hide
        Load ValidateML
        ValidateML.show
        Unload ValidateML
    Me.show
End Sub

Private Sub cbVerifyUnits_Click()
    Me.Hide
        Load VerifyUnits
            VerifyUnits.show
        Unload VerifyUnits
    Me.show
End Sub

Private Sub UserForm_Initialize()
    Dim strFileName As String
    Dim vLine As Variant
    Dim fName As String
    
    strFileName = ThisDrawing.Path & "\xxQC Checklist.qcl"
    
    fName = Dir(strFileName)
    If fName = "" Then
        MsgBox "No File found."
        Exit Sub
    End If
    
    Open strFileName For Input As #2
    
    While Not EOF(2)
        Line Input #2, strLine
        vLine = Split(strLine, vbTab)
        
        Select Case vLine(0)
            Case "Validate Matchlines"
                If vLine(1) = "Y" Then chVML.Value = True
            Case "Validate Customers"
                If vLine(1) = "Y" Then chVCustomers.Value = True
            Case "Validate MCL"
                If vLine(1) = "Y" Then chVMCL.Value = True
            Case "Counts Callouts"
                If vLine(1) = "Y" Then chCC.Value = True
            Case "Validate Callouts"
                If vLine(1) = "Y" Then chVCallouts.Value = True
            Case "Validate Splices"
                If vLine(1) = "Y" Then chVHO1.Value = True
            Case "Validate Units"
                If vLine(1) = "Y" Then chVUnits.Value = True
            Case "Tab"
                If vLine(1) = "Y" Then chTab.Value = True
            Case "Extra Heights"
                If vLine(1) = "Y" Then chEH.Value = True
            Case "Spans for Reels"
                If vLine(1) = "Y" Then chS4R.Value = True
            Case "Print DWGs"
                If vLine(1) = "Y" Then chPrint.Value = True
            Case "Transfer Data to Maps"
                If vLine(1) = "Y" Then chTransfer.Value = True
        End Select
    Wend
    
Exit_Sub:
    Close #2
End Sub

Private Sub SaveData()
    Dim strFileName As String
    Dim vLine As Variant
    Dim strLine As String
    
    strFileName = ThisDrawing.Path & "\xxQC Checklist.qcl"
    
    Open strFileName For Output As #2
    
    If chVML.Value = True Then
        strLine = "Validate Matchlines" & vbTab & "Y"
    Else
        strLine = "Validate Matchlines" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chVCustomers.Value = True Then
        strLine = "Validate Customers" & vbTab & "Y"
    Else
        strLine = "Validate Customers" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chVMCL.Value = True Then
        strLine = "Validate MCL" & vbTab & "Y"
    Else
        strLine = "Validate MCL" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chCC.Value = True Then
        strLine = "Counts Callouts" & vbTab & "Y"
    Else
        strLine = "Counts Callouts" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chVCallouts.Value = True Then
        strLine = "Validate Callouts" & vbTab & "Y"
    Else
        strLine = "Validate Callouts" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chVHO1.Value = True Then
        strLine = "Validate Splices" & vbTab & "Y"
    Else
        strLine = "Validate Splices" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chVUnits.Value = True Then
        strLine = "Validate Units" & vbTab & "Y"
    Else
        strLine = "Validate Units" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chTab.Value = True Then
        strLine = "Tab" & vbTab & "Y"
    Else
        strLine = "Tab" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chEH.Value = True Then
        strLine = "Extra Heights" & vbTab & "Y"
    Else
        strLine = "Extra Heights" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chS4R.Value = True Then
        strLine = "Spans for Reels" & vbTab & "Y"
    Else
        strLine = "Spans for Reels" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chPrint.Value = True Then
        strLine = "Print DWGs" & vbTab & "Y"
    Else
        strLine = "Print DWGs" & vbTab & "N"
    End If
    Print #2, strLine
    
    If chTransfer.Value = True Then
        strLine = "Transfer Data to Maps" & vbTab & "Y"
    Else
        strLine = "Transfer Data to Maps" & vbTab & "N"
    End If
    Print #2, strLine
    
    Close #2
End Sub
