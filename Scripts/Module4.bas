Attribute VB_Name = "Module4"
Public Sub Export()
    Dim vbe As vbe
    Set vbe = ThisDrawing.Application.vbe
    Dim comp As VBComponent
    Dim outDir As String
    outDir = "C:\Users\GabrielFinke\OneDrive - Wecom LLC\Documents\autocad"
    If Dir(outDir, vbDirectory) = "" Then
        MkDir outDir
    End If
    For Each comp In vbe.ActiveVBProject.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule
                comp.Export outDir & "\" & comp.Name & ".bas"
            Case vbext_ct_Document, vbext_ct_ClassModule
                comp.Export outDir & "\" & comp.Name & ".cls"
            Case vbext_ct_MSForm
                comp.Export outDir & "\" & comp.Name & ".frm"
            Case Else
                comp.Export outDir & "\" & comp.Name
        End Select
    Next comp

     MsgBox "VBA files were exported to : " & outDir
End Sub
