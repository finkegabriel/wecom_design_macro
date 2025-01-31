VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zzDrawTest 
   Caption         =   "UserForm1"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8700.001
   OleObjectBlob   =   "zzDrawTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zzDrawTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' use the
Private m_objDrawing As AJPiUFDraw


Private Sub cbDrawLine_Click()
    Dim shpTemp As Shape
    
    Set shpTemp = m_objDrawing.Line(50, 150, 100, 150)
    If Not shpTemp Is Nothing Then shpTemp.Line.Weight = 1
    Set shpTemp = m_objDrawing.Line(100, 150, 100, 200)
    If Not shpTemp Is Nothing Then shpTemp.Line.Weight = 2
    Set shpTemp = m_objDrawing.Line(100, 200, 50, 200)
    If Not shpTemp Is Nothing Then shpTemp.Line.Weight = 3
    Set shpTemp = m_objDrawing.Line(50, 200, 50, 150)
    If Not shpTemp Is Nothing Then shpTemp.Line.Weight = 4
    
    Set shpTemp = m_objDrawing.Line(50, 150, 75, 175)
    If Not shpTemp Is Nothing Then shpTemp.Line.ForeColor.RGB = RGB(255, 0, 0)
    Set shpTemp = m_objDrawing.Line(100, 150, 75, 175)
    If Not shpTemp Is Nothing Then shpTemp.Line.ForeColor.RGB = RGB(0, 255, 0)
    Set shpTemp = m_objDrawing.Line(100, 200, 75, 175)
    If Not shpTemp Is Nothing Then shpTemp.Line.ForeColor.RGB = RGB(0, 0, 255)
    Set shpTemp = m_objDrawing.Line(50, 200, 75, 175)
    If Not shpTemp Is Nothing Then shpTemp.Line.ForeColor.RGB = RGB(255, 255, 255)
    
    m_objDrawing.Paint
End Sub

Private Sub UserForm_Initialize()
    Set m_objDrawing = New AJPiUFDraw
    Set m_objDrawing.CanvasUserform = Me
End Sub
