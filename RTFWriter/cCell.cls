VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCell"
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

Public Border As cBorder
Public BackColorIndex As Long
Public Width As Long

Private m_Cell As String

Private Sub Class_Initialize()
    m_Cell = ""
    BackColorIndex = WhiteColorIndex
    Set Border = New cBorder
    Border.LeftVisible = True
    Border.RightVisible = True
    Border.TopVisible = True
    Border.BottomVisible = True
    Width = 720
End Sub

Public Function InsertParagraph(Paragraph As cParagraph) As Boolean
    If Paragraph Is Nothing Then Exit Function
    m_Cell = m_Cell + Paragraph.RtfParagraph
    InsertParagraph = True
End Function

Friend Function RtfCell() As String
    RtfCell = m_Cell
End Function

Private Sub Class_Terminate()
    Set Border = Nothing
End Sub
