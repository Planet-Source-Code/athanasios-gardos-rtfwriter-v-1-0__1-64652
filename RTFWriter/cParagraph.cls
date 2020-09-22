VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParagraph"
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

Public Align As gsRTFAlignment
Public Border As cBorder
Public IndentFirstLine As Long
Public IndentLeft As Long
Public IndentRight As Long
Public SpaceBefore As Long
Public SpaceAfter As Long
Public SpaceBetweenLines As Long

Const MM_ANISOTROPIC = 8
Const STRETCH_DELETESCANS = 3
Const SRCCOPY = 13369376

Private m_Paragraph As String

Private Sub Class_Initialize()
    m_Paragraph = ""
    Align = aLeft
    IndentFirstLine = 0
    IndentLeft = 0
    IndentRight = 0
    SpaceBefore = 0
    SpaceAfter = 0
    SpaceBetweenLines = 0
    Set Border = New cBorder
End Sub

Friend Function RtfParagraph() As String
    Dim sIdent As String
    Dim sBrdr As String
    sBrdr = GetBorder(Border)
    If IndentFirstLine <> 0 Then sIdent = sIdent + "\fi" & Format$(IndentFirstLine)
    If IndentLeft <> 0 Then sIdent = sIdent + "\li" & Format$(IndentLeft)
    If IndentRight <> 0 Then sIdent = sIdent + "\ri" & Format$(IndentRight)
    If SpaceBefore <> 0 Then sIdent = sIdent + "\sb" & Format$(SpaceBefore)
    If SpaceAfter <> 0 Then sIdent = sIdent + "\sa" & Format$(SpaceAfter)
    If SpaceBetweenLines <> 0 Then sIdent = sIdent + "\sl" & Format$(SpaceBetweenLines)
    If Align = aLeft Then
       m_Paragraph = "{\pard\plain" & m_Paragraph & "}" & sIdent & sBrdr & "\ql\par "
    ElseIf Align = aRight Then
       m_Paragraph = "{\pard\plain" & m_Paragraph & "}" & sIdent & sBrdr & "\qr\par "
    ElseIf Align = aCenter Then
       m_Paragraph = "{\pard\plain" & m_Paragraph & "}" & sIdent & sBrdr & "\qc\par "
    ElseIf Align = aJustified Then
       m_Paragraph = "{\pard\plain" & m_Paragraph & "}" & sIdent & sBrdr & "\qj\par "
    End If
    RtfParagraph = m_Paragraph
End Function

Public Function InsertText(Text As cText) As Boolean
    m_Paragraph = m_Paragraph & RtfText(Text)
    InsertText = True
End Function

Public Function InsertDrawing(Drawing As cDrawing) As Boolean
    If Drawing Is Nothing Then Exit Function
    m_Paragraph = m_Paragraph & Drawing.RtfDrawing
    InsertDrawing = True
End Function


Private Sub ResetFrmRtfPictures(frmRTF1 As Object)
    With frmRTF1
        .Picture1.ScaleMode = vbPixels
        .Picture1.AutoRedraw = True
        .Picture1.AutoSize = True
        .Picture2.ScaleMode = vbPixels
        .Picture2.AutoRedraw = True
        .Picture2.AutoSize = True
        Set .Picture1.Picture = LoadPicture("")
        Set .Picture2.Picture = LoadPicture("")
        .Picture1.Width = 1
        .Picture1.Height = 1
        .Picture2.Width = 1
        .Picture2.Height = 1
    End With
End Sub


Private Sub Class_Terminate()
    Set Border = Nothing
End Sub
