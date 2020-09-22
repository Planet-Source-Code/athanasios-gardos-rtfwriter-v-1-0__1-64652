VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cText"
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

Public FontIndex As Long
Public Height As Long
Public Bold As Boolean
Public Italic As Boolean
Public Strike As Boolean
Public ForeColorIndex As Long
Public BackColorIndex As Long

Private m_Underline As Boolean
Private m_UnderlineDot As Boolean
Private m_UnderlineDouble As Boolean
Private m_UnderlineWords As Boolean
Private m_Text As String

Public Property Get Underline() As Boolean
    Underline = m_Underline
End Property

Public Property Let Underline(ByVal v As Boolean)
    m_Underline = v
    If v = True Then
       UnderlineWords = False
       UnderlineDot = False
       UnderlineDouble = False
    End If
End Property

Public Property Get UnderlineDot() As Boolean
    UnderlineDot = m_UnderlineDot
End Property

Public Property Let UnderlineDot(ByVal v As Boolean)
    m_UnderlineDot = v
    If v = True Then
       Underline = False
       UnderlineDouble = False
    End If
End Property

Public Property Get UnderlineDouble() As Boolean
    UnderlineDouble = m_UnderlineDouble
End Property

Public Property Let UnderlineDouble(ByVal v As Boolean)
    m_UnderlineDouble = v
    If v = True Then
       Underline = False
       UnderlineDot = False
    End If
End Property

Public Property Get UnderlineWords() As Boolean
    UnderlineWords = m_UnderlineWords
End Property

Public Property Let UnderlineWords(ByVal v As Boolean)
    m_UnderlineWords = v
    If v = True Then Underline = False
End Property

Friend Property Get RtfText() As String
    RtfText = m_Text
End Property

Public Property Let Text(ByVal v As String)
    Call replacement(v, "\", Chr$(0))
    Call replacement(v, Chr$(0), "\\")
    Call replacement(v, "{", Chr$(0))
    Call replacement(v, Chr$(0), "\{")
    Call replacement(v, "}", Chr$(0))
    Call replacement(v, Chr$(0), "\}")
    Call replacement(v, Chr$(9), "\tab")
    Call replacement(v, vbCrLf, "\line")
    Call replacement(v, Chr$(13), "\line")
    Call replacement(v, Chr$(149), "\bullet")
    m_Text = v
End Property

Private Sub Class_Initialize()
    m_Text = ""
    FontIndex = 0
    Height = 10
    Bold = False
    Italic = False
    Underline = False
    UnderlineDot = False
    UnderlineDouble = False
    UnderlineWords = False
    Strike = False
    ForeColorIndex = BlackColorIndex
    BackColorIndex = WhiteColorIndex
End Sub
