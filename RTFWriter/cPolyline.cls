VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPolyline"
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
Public PenWidth As Long
Public PenRGBColor As Long
Public FillBackRGBColor As Long
Public FillForeRGBColor As Long
Public FillPattern As Long

Private m_Polyline As String

Private m_xMin As Long
Private m_yMin As Long
Private m_xMax As Long
Private m_yMax As Long

Friend Property Get xMin() As Long
    xMin = m_xMin
End Property

Friend Property Let xMin(ByVal v As Long)
    m_xMin = v
End Property

Friend Property Get yMin() As Long
    yMin = m_yMin
End Property

Friend Property Let yMin(ByVal v As Long)
    m_yMin = v
End Property

Friend Property Get xMax() As Long
    xMax = m_xMax
End Property

Friend Property Let xMax(ByVal v As Long)
    m_xMax = v
End Property

Friend Property Get yMax() As Long
    yMax = m_yMax
End Property

Friend Property Let yMax(ByVal v As Long)
    m_yMax = v
End Property

Private Sub Class_Initialize()
    m_count = 0
    PenWidth = 15
    PenRGBColor = vbBlack
    FillBackRGBColor = vbWhite
    FillForeRGBColor = vbBlack
    FillPattern = 0
End Sub

Public Function InsertPoint(Point As cPoint) As Boolean
    If (Point Is Nothing) Then Exit Function
    m_Polyline = m_Polyline + Point.RtfPoint
    m_count = m_count + 1
    If m_count = 1 Then
       xMin = Point.x
       yMin = Point.y
       xMax = Point.x
       yMax = Point.y
    End If
    If xMin > Point.x Then xMin = Point.x
    If yMin > Point.y Then yMin = Point.y
    If xMax < Point.x Then xMax = Point.x
    If yMax < Point.y Then yMax = Point.y
    InsertPoint = True
End Function

Friend Property Get RtfPolyline() As String
    Dim sTmp As String
    Dim r1 As Byte, g1 As Byte, b1 As Byte, d1 As Byte
    Dim Width As Long, Height As Long
    Width = Abs(xMax - xMin)
    Height = Abs(yMax - yMin)
    sTmp = "\dpx" & Format$(xMin) & "\dpy" & Format$(yMin) & "\dpxsize" & Format$(Width) & "\dpysize" & Format$(Height)
    Call LongToBytes(FillForeRGBColor, r1, g1, b1, d1)
    sTmp = sTmp & "\dpfillfgcr" & Format$(r1) & "\dpfillfgcg" & Format$(g1) & "\dpfillfgcb" & Format$(b1)
    Call LongToBytes(FillBackRGBColor, r1, g1, b1, d1)
    sTmp = sTmp & "\dpfillbgcr" & Format$(r1) & "\dpfillbgcg" & Format$(g1) & "\dpfillbgcb" & Format$(b1)
    sTmp = sTmp & "\dpfillpat" & Format$(FillPattern)
    Call LongToBytes(PenRGBColor, r1, g1, b1, d1)
    sTmp = sTmp & "\dplinew" & Format$(PenWidth) & "\dplinecor" & Format$(r1) & "\dplinecog" & Format$(g1) & "\dplinecob" & Format$(b1)
    RtfPolyline = "\dppolyline\dppolycount" & Format$(m_count) & m_Polyline & sTmp
End Property

