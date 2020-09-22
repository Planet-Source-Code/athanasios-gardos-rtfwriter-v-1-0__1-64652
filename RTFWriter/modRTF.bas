Attribute VB_Name = "modRTF"
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

Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long

Const GMEM_MOVEABLE = 2

Private Type RECTS
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Function GetBorder(Border As cBorder) As String
    Dim sBrdr As String
    Dim sBorderType As String
    Select Case Border.BorderType
    Case gsRTFBorderType.bSingle
        sBorderType = "brdrs"
    Case gsRTFBorderType.bDouble
        sBorderType = "brdrdb"
    Case gsRTFBorderType.bThick
        sBorderType = "brdrth"
    Case gsRTFBorderType.bShadow
        sBorderType = "brdrsh"
    Case gsRTFBorderType.bDot
        sBorderType = "brdrdot"
    Case gsRTFBorderType.bHairline
        sBorderType = "brdrhair"
    End Select
    sBrdr = ""
    If Border.TopVisible = True Then
       sBrdr = sBrdr + "\brdrt\" & sBorderType & "\brdrcf" & Format$(Border.BorderColorIndex) & "\brdrw" + Format$(Border.TopWidth)
    Else
       sBrdr = sBrdr + "\brdrt\brdrnone"
    End If
    If Border.BottomVisible = True Then
       sBrdr = sBrdr + "\brdrb\" & sBorderType & "\brdrcf" & Format$(Border.BorderColorIndex) & "\brdrw" + Format$(Border.BottomWidth)
    Else
       sBrdr = sBrdr + "\brdrb\brdrnone"
    End If
    If Border.LeftVisible = True Then
       sBrdr = sBrdr + "\brdrl\" & sBorderType & "\brdrcf" & Format$(Border.BorderColorIndex) & "\brdrw" + Format$(Border.LeftWidth)
    Else
       sBrdr = sBrdr + "\brdrl\brdrnone"
    End If
    If Border.RightVisible = True Then
       sBrdr = sBrdr + "\brdrr\" & sBorderType & "\brdrcf" & Format$(Border.BorderColorIndex) & "\brdrw" + Format$(Border.RightWidth)
    Else
       sBrdr = sBrdr + "\brdrr\brdrnone"
    End If
    GetBorder = sBrdr
End Function

Public Function RtfText(Text As cText) As String
    Dim s1 As String
    If Text Is Nothing Then Exit Function
    If Text.RtfText = "" Then Exit Function
    If Text.Bold = True Then s1 = s1 + "\b"
    If Text.Italic = True Then s1 = s1 + "\i"
    If Text.UnderlineDot = True Then s1 = s1 + "\uld"
    If Text.UnderlineDouble = True Then s1 = s1 + "\uldb"
    If Text.UnderlineWords = True And Text.Underline = False Then s1 = s1 + "\ulw"
    If Text.Underline = True Then s1 = s1 + "\ul"
    If Text.Strike = True Then s1 = s1 + "\strike"
    s1 = s1 & "\f" + Format$(Text.FontIndex)
    If Text.BackColorIndex <> WhiteColorIndex Then s1 = s1 & ("\chcfpat0\chcbpat" + Format$(Text.BackColorIndex) & "\cb" & Format$(Text.BackColorIndex))
    If Text.ForeColorIndex <> BlackColorIndex Then s1 = s1 & ("\cf" + Format$(Text.ForeColorIndex))
    If Text.Height <> 0 Then s1 = s1 + ("\fs" + Format$(2 * Text.Height))
    RtfText = ("{" + s1 + " " + Text.RtfText + "}")
End Function


