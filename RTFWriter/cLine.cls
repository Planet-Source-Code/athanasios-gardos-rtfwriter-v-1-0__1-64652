VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLine"
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

Public x1 As Long
Public y1 As Long
Public x2 As Long
Public y2 As Long
Public PenWidth As Long
Public PenRGBColor As Long

Friend Property Get RtfLine() As String
    Dim sTmp As String
    Dim r1 As Byte, g1 As Byte, b1 As Byte, d1 As Byte
    Call LongToBytes(PenRGBColor, r1, g1, b1, d1)
    sTmp = "\dplinew" & Format$(PenWidth) & "\dplinecor" & Format$(r1) & "\dplinecog" & Format$(g1) & "\dplinecob" & Format$(b1)
    RtfLine = "\dpline\dpptx" & Format$(x1) & "\dppty" & Format$(y1) & "\dpptx" & Format$(x2) & "\dppty" & Format$(y2) & "\dpx" & Format$(x1) & "\dpy" & Format$(y1) & "\dpxsize" & Format$(Abs(x2 - x1)) & "\dpysize" & Format$(Abs(y2 - y1)) & sTmp
End Property

Private Sub Class_Initialize()
    PenWidth = 15
    PenRGBColor = vbBlack
End Sub
