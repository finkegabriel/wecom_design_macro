VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zzzJMS 
   Caption         =   "Job Management Form"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9015.001
   OleObjectBlob   =   "zzzJMS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zzzJMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    lbJobs.ColumnCount = 2
    lbJobs.ColumnWidths = "48;12"
    
    lbJobs.AddItem "2019-555"
        lbJobs.List(0, 1) = "X"
    lbJobs.AddItem "2020-756"
        lbJobs.List(1, 1) = "A"
    lbJobs.AddItem "2021-010"
        lbJobs.List(2, 1) = "C"
    lbJobs.AddItem "2021-015"
        lbJobs.List(3, 1) = "D"
    lbJobs.AddItem "2021-559"
        lbJobs.List(4, 1) = "H"
    
    
    lbFolder.ColumnCount = 2
    lbFolder.ColumnWidths = "16;80"
    
    lbFolder.AddItem "*"
        lbFolder.List(0, 1) = "Planning"
    lbFolder.AddItem ""
        lbFolder.List(1, 1) = "      Information"
    lbFolder.AddItem "*"
        lbFolder.List(2, 1) = "Design"
    lbFolder.AddItem ""
        lbFolder.List(3, 1) = "      CADD"
    lbFolder.AddItem ""
        lbFolder.List(4, 1) = "      Reports"
    lbFolder.AddItem "!"
        lbFolder.List(5, 1) = "Permits"
    lbFolder.AddItem ""
        lbFolder.List(6, 1) = "      Submitted"
    lbFolder.AddItem ""
        lbFolder.List(7, 1) = "      Approved"
    lbFolder.AddItem ""
        lbFolder.List(8, 1) = "Construction"
    lbFolder.AddItem ""
        lbFolder.List(9, 1) = "      Sheets"
    lbFolder.AddItem ""
        lbFolder.List(10, 1) = "      Reports"
    lbFolder.AddItem ""
        lbFolder.List(11, 1) = "Asbuilts"
    lbFolder.AddItem ""
        lbFolder.List(12, 1) = "      Notes"
    lbFolder.AddItem ""
        lbFolder.List(13, 1) = "      Reports"
    
    
    lbFiles.ColumnCount = 2
    lbFiles.ColumnWidths = "180;60"
    
    lbFiles.AddItem "2021-559 MR - All.dwg"
    lbFiles.List(0, 1) = "CADD File"
    lbFiles.AddItem "2021-559 MAKE READY - ATT.pdf"
    lbFiles.List(1, 1) = "MR Submit"
    lbFiles.AddItem "2021-559 MAKE READY - COMCAST.pdf"
    lbFiles.List(2, 1) = "MR Submit"
    lbFiles.AddItem "2021-559 TDOT PERMIT.dwg"
    lbFiles.List(3, 1) = "CADD File"
    lbFiles.AddItem "2021-559 TDOT PERMIT - SR99.pdf"
    lbFiles.List(4, 1) = "Permit Submit"
    lbFiles.AddItem "2021-559 TDOT PERMIT - COVER.pdf"
    lbFiles.List(5, 1) = "Permit Submit"
    
    Dim strText As String
    
    strText = "2021-559 SUBDIVISION NAME - PHASE #" & vbCr
    strText = strText & "Exchange:" & vbTab & "Full Name, Code, or CLLI" & vbCr
    strText = strText & "Type:" & vbTab & vbTab & "GF-DSPLIT-SFU" & vbCr & vbCr
    strText = strText & "Contact:" & vbTab & "Project Manager" & vbCr
    strText = strText & "Phone:" & vbTab & "888-555-1212" & vbCr
    strText = strText & "Email:" & vbTab & "project.manager@contractor.com" & vbCr
    strText = strText & "Address:" & vbTab & "N/A" & vbCr & vbCr
    strText = strText & "05/04/2021 UC-Engineer: Handoff to Integrity" & vbCr
    strText = strText & "06/08/2021 Integrity-Planner: *** PLANNING COMPLETE" & vbCr
    strText = strText & "06/14/2021 Integrity-Manager: Conduit placement started but not complete." & vbCr
    strText = strText & "06/30/2021 Integrity-Manager: Conduit placement complete." & vbCr
    strText = strText & "07/16/2021 Integrity-Manager: *** DESIGN COMPLETE" & vbCr
    strText = strText & "07/20/2021 UC-Manager: *** PROJECT APPROVED" & vbCr
    strText = strText & "07/23/2021 UC-Permitting: *** TDOT PERMIT SUBMITTED" & vbCr
    strText = strText & "07/23/2021 UC-Permitting: *** ATT MR PACKAGE SUBMITTED" & vbCr
    strText = strText & "07/23/2021 UC-Permitting: *** COMCAST MR PACKAGE SUBMITTED" & vbCr
    strText = strText & "09/17/2021 UC-Manager: *** REVISION REQUIRED" & vbCr
    strText = strText & vbTab & vbTab & "Why:" & vbTab & "Cable need to be rerouted due to Future TDOT Project" & vbCr
    strText = strText & vbTab & vbTab & "Where:" & vbTab & "DWG 003-Pole F5000A/10 to DWG 005-Pole F5000A/21" & vbCr
    strText = strText & "09/30/2021 Integrity-Manager: *** REVISION 09/17/2021 COMPLETE" & vbCr
    
    tbProjectNotes.Value = strText
End Sub
