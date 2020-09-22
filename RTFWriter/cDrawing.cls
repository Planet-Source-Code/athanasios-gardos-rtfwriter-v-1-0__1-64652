VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------
'     © 2006, Athanasios Gardos
'You may freely use, modify and distribute this source code
'
'Last update: March 14, 2006
'Please visit:
'     http://business.hol.gr/gardos/
'for development tools and more source code
'-----------------------------------------------------------

Option Explicit

Private m_count As Long

Private m_xMin As Long
Private m_yMin As Long
Private m_xMax As Long
Private m_yMax As Long

Private m_Drawing As String

Private Sub Class_Initialize()
    m_count = 0
End Sub

Public Function InsertText(Text As cDrawingText) As Boolean
    If (Text Is Nothing) Then Exit Function
    m_Drawing = m_Drawing + Text.RtfDrText
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = Text.x
       m_yMin = Text.y
       m_xMax = Text.x + Text.Width
       m_yMax = Text.y + Text.Height
    End If
    If m_xMin > Text.x Then m_xMin = Text.x
    If m_yMin > Text.y Then m_yMin = Text.y
    If m_xMax < Text.x + Text.Width Then m_xMax = Text.x + Text.Width
    If m_yMax < Text.y + Text.Height Then m_yMax = Text.y + Text.Height
    InsertText = True
End Function

Public Function InsertEllipse(Ellipse As cEllipse) As Boolean
    If (Ellipse Is Nothing) Then Exit Function
    m_Drawing = m_Drawing + Ellipse.RtfEllipse
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = Ellipse.x
       m_yMin = Ellipse.y
       m_xMax = Ellipse.x + Ellipse.Width
       m_yMax = Ellipse.y + Ellipse.Height
    End If
    If m_xMin > Ellipse.x Then m_xMin = Ellipse.x
    If m_yMin > Ellipse.y Then m_yMin = Ellipse.y
    If m_xMax < Ellipse.x + Ellipse.Width Then m_xMax = Ellipse.x + Ellipse.Width
    If m_yMax < Ellipse.y + Ellipse.Height Then m_yMax = Ellipse.y + Ellipse.Height
    InsertEllipse = True
End Function

Public Function InsertPolygon(Polygon As cPolygon) As Boolean
    If (Polygon Is Nothing) Then Exit Function
    m_Drawing = m_Drawing + Polygon.RtfPolygon
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = Polygon.xMin
       m_yMin = Polygon.yMin
       m_xMax = Polygon.xMax
       m_yMax = Polygon.yMax
    End If
    If m_xMin > Polygon.xMin Then m_xMin = Polygon.xMin
    If m_yMin > Polygon.yMin Then m_yMin = Polygon.yMin
    If m_xMax < Polygon.xMax Then m_xMax = Polygon.xMax
    If m_yMax < Polygon.yMax Then m_yMax = Polygon.yMax
    InsertPolygon = True
End Function

Public Function InsertPolyline(Polyline As cPolyline) As Boolean
    If (Polyline Is Nothing) Then Exit Function
    m_Drawing = m_Drawing + Polyline.RtfPolyline
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = Polyline.xMin
       m_yMin = Polyline.yMin
       m_xMax = Polyline.xMax
       m_yMax = Polyline.yMax
    End If
    If m_xMin > Polyline.xMin Then m_xMin = Polyline.xMin
    If m_yMin > Polyline.yMin Then m_yMin = Polyline.yMin
    If m_xMax < Polyline.xMax Then m_xMax = Polyline.xMax
    If m_yMax < Polyline.yMax Then m_yMax = Polyline.yMax
    InsertPolyline = True
End Function

Public Function InsertLine(Line As cLine) As Boolean
    If (Line Is Nothing) Then Exit Function
    m_Drawing = m_Drawing + Line.RtfLine
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = Line.x1
       m_yMin = Line.y1
       m_xMax = Line.x2
       m_yMax = Line.y2
    End If
    If m_xMin > Line.x1 Then m_xMin = Line.x1
    If m_xMin > Line.x2 Then m_xMin = Line.x2
    If m_yMin > Line.y1 Then m_yMin = Line.y1
    If m_yMin > Line.y2 Then m_yMin = Line.y2
    If m_xMax < Line.x1 Then m_xMax = Line.x1
    If m_xMax < Line.x2 Then m_xMax = Line.x2
    If m_yMax < Line.y1 Then m_yMax = Line.y1
    If m_yMax < Line.y2 Then m_yMax = Line.y2
    InsertLine = True
End Function

Public Function InsertRect(Rect As cRect) As Boolean
    If (Rect Is Nothing) Then Exit Function
    m_Drawing = m_Drawing + Rect.RtfRect
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = Rect.x
       m_yMin = Rect.y
       m_xMax = Rect.x + Rect.Width
       m_yMax = Rect.y + Rect.Height
    End If
    If m_xMin > Rect.x Then m_xMin = Rect.x
    If m_xMin > Rect.x + Rect.Width Then m_xMin = Rect.x + Rect.Width
    If m_yMin > Rect.y Then m_yMin = Rect.y
    If m_yMin > Rect.y + Rect.Height Then m_yMin = Rect.y + Rect.Height
    If m_xMax < Rect.x Then m_xMax = Rect.x
    If m_xMax < Rect.x + Rect.Width Then m_xMax = Rect.x + Rect.Width
    If m_yMax < Rect.y Then m_yMax = Rect.y
    If m_yMax < Rect.y + Rect.Height Then m_yMax = Rect.y + Rect.Height
    InsertRect = True
End Function

Friend Property Get RtfDrawing() As String
    Dim sTmp As String
    Dim lWidth As Long
    Dim lHeight As Long
    lWidth = Abs(m_xMax - m_xMin)
    lHeight = Abs(m_yMax - m_yMin)
    If lWidth = 0 Then lWidth = 1
    If lHeight = 0 Then lHeight = 1
    sTmp = "{\*\do\dobxcolumn\dobypara\dodhgt" & Format$(Abs(m_yMax - m_yMin))
    sTmp = sTmp & "\dpx0\dpy0\dpxsize" & Format$(lWidth) & "\dpysize" & Format$(lHeight)
    sTmp = sTmp & "\dpgroup\dpcount" & Format$(m_count + 1)
    sTmp = sTmp & "\dpx" & Format$(m_xMin) & "\dpy" & Format$(m_yMin) & "\dpxsize" & Format$(lWidth) & "\dpysize" & Format$(lHeight)
    sTmp = sTmp & (m_Drawing & "\dpendgroup\dpx0\dpy0\dpxsize0\dpysize0}")
    RtfDrawing = sTmp
End Property

