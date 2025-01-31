VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LotUpdate 
   Caption         =   "Lot Update"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6540
   OleObjectBlob   =   "LotUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LotUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objSS, objSS2 As AcadSelectionSet

Private Sub btn_GenerateMCL_Click()
    Dim filterType, filterValue, filterValue2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0), grpValue2(0) As Variant
    
    Dim vPnt1, vPnt2 As Variant
    Dim dCoords() As Double
    Dim vReturnPnt, vCoords As Variant
    Dim iCounter As Integer
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double
           
    Dim objEntity As AcadEntity
    Dim objPoint As AcadPoint
    Dim objBlock As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    Dim vBlockAtt As Variant

    Dim strAttList() As String
    Dim strAttTemp() As String
    Dim strLine, strTemp As Variant
    Dim str1, str2, str3 As String
    Dim iCount, iTest As Integer
    Dim iTemp As Integer
    Dim vLine, vTemp, vForm As Variant
    Dim vData As Variant
    On Error Resume Next
    
    'WINDOW SELECT SETTTINGS
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If

    
    grpCode(0) = 2
    grpValue(0) = "Customer"
    filterType = grpCode
    filterValue = grpValue
    objSS.Clear
    
    Err = 0
    
    Me.Hide
    
    'WINDOW SELECT SUBROUTINE
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    dPnt1(0) = vPnt1(0)
    dPnt1(1) = vPnt1(1)
    dPnt1(2) = vPnt1(2)
    
    dPnt2(0) = vPnt2(0)
    dPnt2(1) = vPnt2(1)
    dPnt2(2) = vPnt2(2)
    
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    MsgBox "Customer" & " found:  " & objSS.count
    
    'CODE TO GRAB THE DATA FROM THE MCL FILE
    Dim strFile, strFolder As String
    Dim strFileName As String
    Dim vName, vItem, strText, vText As Variant
    Dim strTab, strCable As String
    Dim fName As String
    Dim iIndex, j As Integer
    
    strFolder = ThisDrawing.Path & "\*.*"
    
    strFile = Dir$(strFolder)
    
    Do While strFile <> ""      '<> is not equal to
        If InStr(strFile, ".mcl") Then
            strTab = Replace(strFile, ".mcl", "")
            vLine = Split(strTab, " -")
            'TabStrip1.Tabs.Add vLine(1), vLine(1)   'THIS CREATES NEW TABS FOR THE LISTS WHICH WE WONT NEED
            
            strFileName = ThisDrawing.Path & "\" & strFile
            
            'OPEN THE FILE AND SAVE THE CONTENTS AND REMOVE VBLF(NEXT LINE) AND VBCR(PARAGRAPH ENTER)
            Open strFileName For Input As #1
                strText = Input(LOF(1), 1)
            Close #1
            
            strText = Replace(strText, vbLf, "")
            vText = Split(strText, vbCr)
            
            txtbx_test = vText(0)
            
            'UPDATING THE FIRST HEADER LINE OF THE TEXT FILE IF IT ALSO HAS A LOT NUMBER (AS OF 10.4.23 MAY NOT BE NEEDED AS WE WILL USE GPS INSTEAD)
            vTemp = Split(vText(0), " ")
            For Each objEntity In objSS
                Set objBlock = objEntity
                vBlockAtt = objBlock.GetAttributes
                
                If InStr(1, vBlockAtt(3).TextString, ", ") > 0 Then
                    strTemp = Split(vBlockAtt(3).TextString, ", ")
                Else
                    GoTo No_Match1
                End If
                
                If vTemp(1) = strTemp(1) Then
                    vTemp(1) = vBlockAtt(1).TextString
                    vText(0) = vTemp(0) & " " & vTemp(1) & " " & vTemp(2)
                End If
No_Match1:
            Next objEntity
            
            'FIND, MATCH, AND REPLACE EACH ROW OF THE TEXT FILE IF THE ADDRESS ROW MATCHES THE LOT
            For j = 1 To UBound(vText)
                    vItem = Split(vText(j), vbTab)
                    
                    For Each objEntity In objSS
                        Set objBlock = objEntity
                        vBlockAtt = objBlock.GetAttributes
                        
                        If InStr(1, vBlockAtt(3).TextString, ", ") > 0 Then
                            strTemp = Split(vBlockAtt(3).TextString, ", ")
                        Else
                            GoTo No_Match2
                        End If

                        If strTemp(1) = vItem(2) Then
                            If Not vItem(2) = "<>" Then
                            
                                vItem(2) = vBlockAtt(1).TextString
                                vItem(3) = vBlockAtt(2).TextString
                                vItem(5) = vBlockAtt(3).TextString
                                
                                
                                vText(j) = vItem(0) & vbTab & vItem(1) & vbTab & vItem(2) & vbTab & vItem(3) & vbTab & vItem(4) & vbTab & vItem(5)
                                
                                GoTo Found_Match
                            End If
                        End If
No_Match2:
                    Next objEntity
Found_Match:

            Next j
            
            'COMBINING THE TEXT FILE ROWS TO MAKE A COMPLETE FILE
            strText = vText(0)
            If UBound(vText) > 0 Then
                For j = 1 To UBound(vText)
                    If Not vText(j) = "" Then strText = strText & vbCr & vText(j)
                Next j
            End If
            
            
            Open strFileName For Output As #1
                Print #1, strText
            Close #1
            

        End If
        strFile = Dir$
    Loop
    
    strTemp.Clear
    vTemp.Clear
    vItem.Clear
    Me.show
    
End Sub

Private Sub btn_UpdateLots_Click()
    Dim vLotsList As Variant
    Dim vStreetNumbersList As Variant
    Dim vAddressesList As Variant
    Dim vTempList1 As String
    Dim vTempList2 As String
    Dim vTempList3 As String
    Dim i, j As Integer
    
    Dim filterType, filterValue, filterValue2 As Variant
    Dim grpCode(0) As Integer
    Dim grpValue(0), grpValue2(0) As Variant
    
    Dim vPnt1, vPnt2 As Variant
    Dim dCoords() As Double
    Dim vReturnPnt, vCoords As Variant
    Dim iCounter As Integer
    Dim dPnt1(0 To 2) As Double
    Dim dPnt2(0 To 2) As Double

    Dim objEntity As AcadEntity
    Dim objPoint As AcadPoint
    Dim objBlock As AcadBlockReference
    Dim objLWP As AcadLWPolyline
    Dim vBlockAtt As Variant

    Dim strAttList() As String
    Dim strAttTemp() As String
    Dim strLine, strTemp As String
    Dim str1, str2, str3 As String
    Dim iCount, iTest As Integer
    Dim iTemp As Integer
    Dim vLine, vTemp, vForm As Variant
    Dim vData As Variant
    
    On Error Resume Next
    
    Set objSS = ThisDrawing.SelectionSets.Add("objSS")
    If Not Err = 0 Then
        Set objSS = ThisDrawing.SelectionSets.Item("objSS")
        objSS.Clear
    End If
    
    Set objSS2 = ThisDrawing.SelectionSets.Add("objSS2")
    If Not Err = 0 Then
        Set objSS2 = ThisDrawing.SelectionSets.Item("objSS2")
        objSS2.Clear
    End If
    
    'create a itemized list of the copy pasted row of Lots
    vTempList1 = txtbx_Lots.Text
    vLotsList = Split(vTempList1, vbTab)
    
    'create an itemized list of the copy pasted street numbers
    vTempList2 = txtbx_StreetNumbers.Text
    vStreetNumbersList = Split(vTempList2, vbTab)
    
    'create an itemized list of the copy pasted Addresses (not needed at this time)
    vTempList3 = txtbx_Addresses.Text
    vAddressesList = Split(vTempList3, vbTab)
        
    grpCode(0) = 2
    grpValue(0) = "Customer"
    grpValue2(0) = "iSplitter"
    filterType = grpCode
    filterValue = grpValue
    filterValue2 = grpValue2
    objSS.Clear
    objSS2.Clear
    
    Err = 0
    
    Me.Hide
    'select all
    'objSS.Select acSelectionSetAll, , , filterType, filterValue
    
    'window select
    vPnt1 = ThisDrawing.Utility.GetPoint(, "Get BL Corner: ")
    vPnt2 = ThisDrawing.Utility.GetCorner(vPnt1, vbCr & "Get UR Corner: ")
    
    dPnt1(0) = vPnt1(0)
    dPnt1(1) = vPnt1(1)
    dPnt1(2) = vPnt1(2)
    
    dPnt2(0) = vPnt2(0)
    dPnt2(1) = vPnt2(1)
    dPnt2(2) = vPnt2(2)
    
    objSS.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue
    objSS2.Select acSelectionSetWindow, dPnt1, dPnt2, filterType, filterValue2
    
    MsgBox "Customer" & " found:  " & objSS.count
    
    For Each objEntity In objSS
        Set objBlock = objEntity
        vBlockAtt = objBlock.GetAttributes
        
        For i = 0 To UBound(vLotsList)
            If vLotsList(i) = vBlockAtt(1).TextString Then
                vBlockAtt(3).TextString = vBlockAtt(3).TextString & ", " & vLotsList(i)
                vBlockAtt(1).TextString = vStreetNumbersList(i)
                vBlockAtt(2).TextString = vAddressesList(i)
                
                GoTo match_found
            End If
        Next i
match_found:
        objBlock.Update
    
    Next objEntity
    
    For Each objEntity In objSS2
        Set objBlock = objEntity
        vBlockAtt = objBlock.GetAttributes
        
        For i = 0 To UBound(vLotsList)
            vTemp = Split(vBlockAtt(1).TextString, " ")
            
            If vLotsList(i) = vTemp(1) Then
                vTemp(1) = vStreetNumbersList(i)
                vTemp(2) = vAddressesList(i)
                vBlockAtt(1).TextString = vTemp(0) & " " & vTemp(1) & " " & vTemp(2)


                GoTo match_found2
            End If
        Next i
match_found2:
    objBlock.Update
    
    Next objEntity
    
    Me.show
    
End Sub

