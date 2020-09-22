VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrawingText"
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

Public x As Long
Public y As Long
Public Width As Long
Public Height As Long
Public PenWidth As Long
Public PenRGBColor As Long
Public FillBackRGBColor As Long
Public FillForeRGBColor As Long
Public FillPattern As Long

Private m_Text As cText

Private Sub Class_Initialize()
    PenWidth = 15
    PenRGBColor = vbBlack
    FillBackRGBColor = vbWhite
    FillForeRGBColor = vbBlack
    FillPattern = 0
End Sub

Friend Property Get RtfDrText() As String
    Dim sTmp As String
    Dim r1 As Byte, g1 As Byte, b1 As Byte, d1 As Byte
    sTmp = "\dpx" & Format$(x) & "\dpy" & Format$(y) & "\dpxsize" & Format$(Width) & "\dpysize" & Format$(Height)
    Call LongToBytes(FillForeRGBColor, r1, g1, b1, d1)
    sTmp = sTmp & "\dpfillfgcr" & Format$(r1) & "\dpfillfgcg" & Format$(g1) & "\dpfillfgcb" & Format$(b1)
    Call LongToBytes(FillBackRGBColor, r1, g1, b1, d1)
    sTmp = sTmp & "\dpfillbgcr" & Format$(r1) & "\dpfillbgcg" & Format$(g1) & "\dpfillbgcb" & Format$(b1)
    sTmp = sTmp & "\dpfillpat" & Format$(FillPattern)
    Call LongToBytes(PenRGBColor, r1, g1, b1, d1)
    sTmp = sTmp & "\dplinew" & Format$(PenWidth) & "\dplinecor" & Format$(r1) & "\dplinecog" & Format$(g1) & "\dplinecob" & Format$(b1)
    RtfDrText = "\dptxbx" & sTmp & "{\dptxbxtext\pard\plain " & RtfText(m_Text) & "\par}"
End Property

Public Property Let Text(ByVal v As cText)
    Set m_Text = v
End Property



