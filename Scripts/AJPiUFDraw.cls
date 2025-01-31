VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AJPiUFDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_frmUFCanvas As MSForms.UserForm
Private m_objCanvas As Object
Private m_shtHolder As Worksheet
Private m_blnUseUF As Boolean
Public Function Box(Left As Single, Top As Single, Width As Single, Height As Single) As Shape
'
' Add a box to the drawing
'
    On Error GoTo ErrBox
    Set Box = m_shtHolder.Shapes.AddShape(msoShapeRectangle, Left, Top, Width, Height)
    Exit Function
ErrBox:
    Set Box = Nothing
    Exit Function
End Function
Public Sub CleanCanvas()
'
' Clear both userform and temporary work sheet
'
    On Error Resume Next
    m_ClearHolder
    WipeCanvas
End Sub
'Public Function TextEffect(Left As Single, Top As Single, _
'                           Effect As MsoPresetTextEffect, Text As String, _
'                           Optional FontName As String = "Arial", _
'                           Optional FontSize As Single = 10, _
'                           Optional FontBold As MsoTriState = msoFalse, _
'                           Optional FontItalic As MsoTriState = msoFalse) As Shape
''
'' Add a Wordart to the drawing
''
'    On Error GoTo ErrTextEffect
'    Set TextEffect = m_shtHolder.Shapes.AddTextEffect(Effect, Text, _
'         FontName, FontSize, FontBold, FontItalic, Left, Top)
'    Exit Function
'ErrTextEffect:
'    Set TextEffect = Nothing
'    Exit Function
'End Function
Public Function Curve(Xs() As Single, Ys() As Single) As Shape
'
' Add a curve to the drawing
'
    Dim intIndex As Integer
    
    On Error GoTo ErrCurve
    With m_shtHolder.Shapes.BuildFreeform(msoEditingAuto, Xs(LBound(Xs)), Ys(LBound(Ys)))
        For intIndex = LBound(Xs) + 1 To UBound(Xs)
            .AddNodes msoSegmentCurve, msoEditingAuto, Xs(intIndex), Ys(intIndex)
        Next
        Set Curve = .ConvertToShape
    End With
    Exit Function
ErrCurve:
    Set Curve = Nothing
    Exit Function
End Function
Public Function FreeForm(Xs() As Single, Ys() As Single, Closed As Boolean) As Shape
'
' Add a Freeform to the drawing
'
    Dim intIndex As Integer
        
    On Error GoTo ErrFreeForm
    With m_shtHolder.Shapes.BuildFreeform(msoEditingAuto, Xs(LBound(Xs)), Ys(LBound(Ys)))
        For intIndex = LBound(Xs) + 1 To UBound(Xs)
            .AddNodes msoSegmentLine, msoEditingAuto, Xs(intIndex), Ys(intIndex)
        Next
        If Closed Then
            .AddNodes msoSegmentLine, msoEditingAuto, Xs(LBound(Xs)), Ys(LBound(Ys))
        End If
        Set FreeForm = .ConvertToShape
    End With
    Exit Function
ErrFreeForm:
    Set FreeForm = Nothing
    Exit Function
End Function
Public Function Oval(CenterX As Single, CenterY As Single, RadiusX As Single, RadiusY As Single) As Shape
'
' Add a circle to the drawing
'
    Dim Left As Single
    Dim Top As Single
    Dim Width As Single
    Dim Height As Single
    
    On Error GoTo ErrOval
    Left = CenterX - RadiusX
    Top = CenterY - RadiusY
    Width = RadiusX * 2
    Height = RadiusY * 2
    Set Oval = m_shtHolder.Shapes.AddShape(msoShapeOval, Left, Top, Width, Height)
    Exit Function
ErrOval:
    Set Oval = Nothing
    Exit Function
End Function
Public Function Round(CenterX As Single, CenterY As Single, Radius As Single) As Shape
'
' Add a circle to the drawing
'
    Dim Left As Single
    Dim Top As Single
    Dim Width As Single
    Dim Height As Single
    
    On Error GoTo ErrRound
    Left = CenterX - Radius
    Top = CenterY - Radius
    Width = Radius * 2
    Height = Radius * 2
    Set Round = m_shtHolder.Shapes.AddShape(msoShapeOval, Left, Top, Width, Height)
    Exit Function
ErrRound:
    Set Round = Nothing
    Exit Function
End Function
Public Function Line(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Shape
'
' Add a Line to the drawing
' always draw line with X2 & Y2 largest
'
    Dim sngX1 As Single
    Dim sngY1 As Single
    Dim sngX2 As Single
    Dim sngY2 As Single
    
    On Error GoTo ErrLine
    sngX1 = X1
    If X2 < sngX1 Then
        sngX1 = X2
        sngX2 = X1
    Else
        sngX2 = X2
    End If
    sngY1 = Y1
    If Y2 < sngY1 Then
        sngY1 = Y2
        sngY2 = Y1
    Else
        sngY2 = Y2
    End If
    
    Set Line = m_shtHolder.Shapes.AddLine(X1, Y1, X2, Y2)
    Exit Function
ErrLine:
    Set Line = Nothing
    Exit Function
End Function
Public Property Set CanvasGeneric(MyCanvas As Object)
'
' Create a pointer to an object
' this allows use to use any control that has a picture property
' set picture sizing and alignment
'
    On Error GoTo ErrCanvas
    m_blnUseUF = False
    Set m_objCanvas = MyCanvas
    m_objCanvas.PictureSizeMode = fmPictureSizeModeClip
    m_objCanvas.PictureAlignment = fmPictureAlignmentTopLeft
ErrCanvas:
    Exit Property
End Property
Public Property Set CanvasUserform(MyUserform As MSForms.UserForm)
'
' Create a pointer to the userform
' set picture sizing and alignment
'
    On Error GoTo ErrUFCanvas
    m_blnUseUF = True
    Set m_frmUFCanvas = MyUserform
    m_frmUFCanvas.PictureSizeMode = fmPictureSizeModeClip
    m_frmUFCanvas.PictureAlignment = fmPictureAlignmentTopLeft
    Exit Property
ErrUFCanvas:
    Set CanvasUserform = Nothing
    Exit Property
End Property
Public Sub WipeCanvas()
' clear canvas
    On Error Resume Next
    If m_blnUseUF Then
        m_frmUFCanvas.Picture = LoadPicture("")
    Else
        m_objCanvas.Picture = LoadPicture("")
    End If
End Sub
Private Sub m_ClearHolder()
'
' Remove all shapes from holding worksheet
'
    On Error GoTo ErrClearHolder
    Do While m_shtHolder.Shapes.count > 0
        m_shtHolder.Shapes(1).Delete
    Loop
ErrClearHolder:
    Exit Sub
End Sub
Private Function m_CreateFrame() As Shape
'
' Create a shape top left in order to keep the other
' elements in the correct offset positions
'
    Dim lngColor As Long

    On Error GoTo ErrCreateFrame
    If m_blnUseUF Then
        If m_frmUFCanvas.BackColor >= 0 Then
            lngColor = m_frmUFCanvas.BackColor
        Else
            lngColor = m_frmUFCanvas.BackColor And &HFF
            lngColor = GetSysColor(lngColor)
        End If
        Set m_CreateFrame = m_shtHolder.Shapes.AddShape(msoShapeRectangle, 0, 0, m_frmUFCanvas.InsideWidth - 1, m_frmUFCanvas.InsideHeight - 1)
    Else
        If m_objCanvas.BackColor >= 0 Then
            lngColor = m_objCanvas.BackColor
        Else
            lngColor = m_objCanvas.BackColor And &HFF
            lngColor = GetSysColor(lngColor)
        End If
        Set m_CreateFrame = m_shtHolder.Shapes.AddShape(msoShapeRectangle, 0, 0, m_objCanvas.Width - 1, m_objCanvas.Height - 1)
    End If
    With m_CreateFrame
        With .Fill
            .ForeColor.RGB = lngColor
            .BackColor.RGB = lngColor
        End With
        .Line.BackColor.RGB = lngColor
        .Line.ForeColor.RGB = lngColor
        .ZOrder msoSendToBack
    End With
    Exit Function
ErrCreateFrame:
    Set m_CreateFrame = Nothing
    Exit Function
End Function
Public Sub Paint()
'
' Group all shapes and copy to picture
'
    Dim shpFrame As Shape
    Dim intIndex As Integer
    
    On Error Resume Next
    Set shpFrame = m_CreateFrame
    
    ReDim vntItems(m_shtHolder.Shapes.count - 1) As Variant
    For Each shpFrame In m_shtHolder.Shapes
        vntItems(intIndex) = shpFrame.Name
        intIndex = intIndex + 1
    Next
    m_shtHolder.Shapes.Range(vntItems).Group
    Set shpFrame = m_shtHolder.Shapes(1)
    shpFrame.CopyPicture xlScreen, xlBitmap

    'Paste the picture from the clipboard into our image control
    If m_blnUseUF Then
        Set m_frmUFCanvas.Picture = PastePicture(xlBitmap)
    Else
        Set m_objCanvas.Picture = PastePicture(xlBitmap)
    End If
    
End Sub
Private Sub Class_Initialize()

    Set m_shtHolder = ThisWorkbook.Worksheets("AJPiHolder")
    m_ClearHolder
    
End Sub
Private Sub Class_Terminate()
    ' remove any shapes
    On Error Resume Next
    m_ClearHolder
    ThisWorkbook.Saved = True
    Set m_shtHolder = Nothing
    Set m_frmUFCanvas = Nothing
    Set m_objCanvas = Nothing
End Sub

