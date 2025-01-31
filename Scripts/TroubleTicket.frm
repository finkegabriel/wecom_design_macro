VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TroubleTicket 
   Caption         =   "Trouble Ticket"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7320
   OleObjectBlob   =   "TroubleTicket.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TroubleTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strSubject As String

Private Sub cbApp_Change()
    tbSubject.Value = strSubject & cbApp.Value
End Sub

Private Sub cbQuit_Click()
    Me.Hide
End Sub

Private Sub cbSend_Click()
    If cbType.Value = "Complaints" Then
        
    End If
    
    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem
    Dim strBody As String

    Set objOutlook = New Outlook.Application
    Set objMail = objOutlook.CreateItem(olMailItem)
    
    Call ResetSubject
    strSubject = strSubject & cbApp.Value & " --> " & cbType.Value
    
    strBody = tbNote.Value '& vbCr & vbCr & "Example:" & vbCr & tbExample.Value
    If Not tbExample.Value = "" Then strBody = strBody & vbCr & vbCr & "Examples:" & vbCr & tbExample.Value
    strBody = Replace(strBody, vbCr, "<br>")
    strBody = Replace(strBody, vbLf, "")
    
    objMail.To = "rich.taylor@integrity-us.com"
    'objMail.To = objOutlook.Session.CurrentUser.Address
    objMail.CC = objOutlook.Session.CurrentUser.Address
    objMail.Subject = strSubject
    objMail.HTMLBody = strBody
    
    If InStr(cbType.Value, "High Priority") > 0 Then objMail.FlagRequest = "Follow up"
    
    If cbAttachment.Value = True Then
        objMail.Display
    Else
        objMail.Send
    End If
    
    Me.Hide
End Sub

Private Sub cbType_Change()
    tbSubject.Value = strSubject & cbApp.Value & " --> " & cbType.Value
    'If cbType.Value = "Complaints" Then tbNote.Value = "Hey Asshole," & vbCr & vbCr
End Sub

Private Sub UserForm_Initialize()
    cbApp.AddItem "Integrity Tools"
    cbApp.AddItem "HLE"
    cbApp.AddItem "AutoCAD"
    
    cbType.AddItem "Bug - High Priority"
    cbType.AddItem "Bug"
    'cbType.AddItem "Bug - Low Priority"
    cbType.AddItem "Request - High Priority"
    cbType.AddItem "Request / Suggestion"
    cbType.AddItem "Assistance - High Priority"
    cbType.AddItem "Assistance / Training"
    'cbType.AddItem "Complaints"
    
    'tbSubject.Value = "<Ticket #"
    Call ResetSubject
    tbSubject.Value = strSubject
End Sub

Private Sub ResetSubject()
    Dim strNumber As String
    Dim iTemp As Integer
    
    iTemp = Year(Date) - 2000
    strNumber = CStr(iTemp)
    
    iTemp = Month(Date)
    If iTemp < 10 Then
        strNumber = strNumber & "0" & CStr(iTemp)
    Else
        strNumber = strNumber & CStr(iTemp)
    End If
    
    iTemp = Day(Date)
    If iTemp < 10 Then
        strNumber = strNumber & "0" & CStr(iTemp) & "-"
    Else
        strNumber = strNumber & CStr(iTemp) & "-"
    End If
    
    iTemp = Hour(Time)
    If iTemp < 10 Then
        strNumber = strNumber & "0" & CStr(iTemp)
    Else
        strNumber = strNumber & CStr(iTemp)
    End If
    
    iTemp = Minute(Time)
    If iTemp < 10 Then
        strNumber = strNumber & "0" & CStr(iTemp) & "."
    Else
        strNumber = strNumber & CStr(iTemp) & "."
    End If
    
    iTemp = Second(Time)
    If iTemp < 10 Then
        strNumber = strNumber & "0" & CStr(iTemp)
    Else
        strNumber = strNumber & CStr(iTemp)
    End If
    
    strSubject = "<Ticket #" & strNumber & "> "
    
End Sub
