VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDocument"
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

Public Author As String
Public Operator As String
Public Title As String
Public Subject As String
Public PageWidth As Long
Public PageHeight As Long
Public PageMarginLeft As Long
Public PageMarginRight As Long
Public PageMarginTop As Long
Public PageMarginBottom As Long
Public PageLandscape As Boolean
Public PageNumbersOn As Boolean
Public PageNumberStart As Long
Public PageNumberMarginRight As Long
Public PageNumberMarginTop As Long
Public FontDefaultName As String
Public FontDefaultCharSet As Long
Public FontDefaultHeight As Long
Public Header As cParagraph
Public Footer As cParagraph

Private m_Document As String
Private m_FontName() As String
Private m_FontCharSet() As Long
Private m_Colors() As Long
Private m_PageCnt As Long

Public Enum gsRTFAlignment
    aLeft = 0
    aRight = 1
    aCenter = 2
    aJustified = 3
End Enum

Private Sub Class_Initialize()
    PageWidth = 11906
    PageHeight = 16837
    PageLandscape = False
    PageMarginLeft = 1440
    PageMarginRight = 1440
    PageMarginTop = 1440
    PageMarginBottom = 1440
    FontDefaultName = "Arial"
    FontDefaultHeight = 10
    FontDefaultCharSet = 0
    PageNumberStart = 1
    PageNumberMarginRight = 720
    PageNumberMarginTop = 720
    PageNumbersOn = False
    ReDim m_FontName(0) As String
    ReDim m_Colors(0) As Long
    ReDim m_FontCharSet(0) As Long
    m_Document = ""
    BlackColorIndex = ColorIndex(vbBlack)
    WhiteColorIndex = ColorIndex(vbWhite)
End Sub

Public Property Get LastPage() As Long
    LastPage = m_PageCnt
End Property

Public Function Save(sRTFFile As String) As Boolean
    Dim iFr As Integer
    If sRTFFile = "" Then Exit Function
    iFr = FreeFile
    Open sRTFFile For Output As #iFr
    Print #iFr, RtfDocument
    Close #iFr
    Save = IsFile(sRTFFile)
End Function

Public Function Rtf() As String
    Rtf = RtfDocument
End Function

Friend Function RtfDocument() As String
    Dim sTmp As String
    Dim aplast As Single
    Dim apo As Single
    Dim sFrace As String
    sFrace = "\page "
    sTmp = m_Document
    apo = 1
    Call findinstrlast(apo, sTmp, sFrace, aplast)
    If aplast <> 0 Then
       Mid$(sTmp, aplast, Len(sFrace)) = Space$(Len(sFrace))
    End If
    RtfDocument = RTFHeader & ("{" + RTrim$(sTmp) + "}}")
End Function

Public Function InsertPage(Page As cPage) As Boolean
    If Page Is Nothing Then Exit Function
    m_PageCnt = m_PageCnt + 1
    m_Document = m_Document + Page.RtfPage
    InsertPage = True
End Function

Public Function FontIndex(FontName As String, Optional CharSet As Long = -1) As Long
    Dim l As Long
    If FontName = "" Then FontIndex = 0: Exit Function
    m_FontName(0) = FontName
    For l = 0 To UBound(m_FontName())
        If UCase$(FontName) = UCase$(m_FontName(l)) Then
           FontIndex = l
           Exit Function
        End If
    Next l
    ReDim Preserve m_FontName(UBound(m_FontName()) + 1)
    ReDim Preserve m_FontCharSet(UBound(m_FontCharSet()) + 1)
    m_FontName(UBound(m_FontName())) = FontName
    If CharSet = -1 Then
       m_FontCharSet(UBound(m_FontCharSet())) = FontDefaultCharSet
    Else
       m_FontCharSet(UBound(m_FontCharSet())) = CharSet
    End If
    FontIndex = UBound(m_FontName())
End Function

Public Function ColorIndex(RGBColor As Long) As Long
    Dim l As Long
    m_Colors(0) = 0
    For l = 0 To UBound(m_Colors())
        If RGBColor = m_Colors(l) Then
           ColorIndex = l + 1
           Exit Function
        End If
    Next l
    ReDim Preserve m_Colors(UBound(m_Colors()) + 1)
    m_Colors(UBound(m_Colors())) = RGBColor
    ColorIndex = UBound(m_Colors()) + 1
End Function

Private Function RTFHeader() As String
    On Error Resume Next
    Dim hr$, Mnt$, Yr$
    Dim dy$, mn$, nowdt$
    Dim PaperW&, hder$, sLsp$
    Dim RtfAddToHeader$
    Dim i As Long, l As Long
    Dim r1 As Byte, g1 As Byte
    Dim b1 As Byte, d1 As Byte
    hr$ = FnstM(Hour(Time))
    Mnt$ = FnstM(Minute(Time))
    Yr$ = FnstM(Year(Date))
    mn$ = FnstM(Month(Date))
    dy$ = FnstM(Day(Date))
    nowdt$ = "\yr" & Yr$ & "\mo" & mn$ & "\dy" & dy$ & "\hr" & hr$ & "\min" & Mnt$
    PaperW& = PageWidth
    hder$ = "{\rtf1\ansi\ansicpg1253\uc0\deff0"
    hder$ = hder$ & "{\fonttbl{\f0\fnil\fcharset"
    hder$ = hder$ & (Format$(FontDefaultCharSet) & "\fprq2 " & FontDefaultName & ";}")
    For l = 1 To UBound(m_FontName())
        hder$ = hder$ & ("{\f" & Format$(l) & "\\fnil\fcharset" & Format$(m_FontCharSet(l)) & "\fprq2 " & m_FontName(l) & ";}")
    Next l
    hder$ = hder$ & "}"
    hder$ = hder$ & "{\colortbl;"
    For i = 0 To UBound(m_Colors)
        Call LongToBytes(m_Colors(i), r1, g1, b1, d1)
        hder$ = hder$ & ("\red" & Format$(r1) & "\green" & Format$(g1) & "\blue" & Format$(b1) & ";")
    Next i
    hder$ = hder$ & "}"
    hder$ = hder$ & ("{\stylesheet{\fs" & Format$(FontDefaultHeight * 2))
    hder$ = hder$ & " \snext0 Normal;}}"
    hder$ = hder$ & ("{\info{\author " & Author & "}")
    hder$ = hder$ & ("{\operator " & Operator & "}")
    hder$ = hder$ & ("{\title " & Title & "}")
    hder$ = hder$ & ("{\subject " & Subject & "}")
    hder$ = hder$ & "{\comment Athanasios Gardos - RTFWriter.dll}"
    hder$ = hder$ & ("{\creatim" & nowdt$ & "}")
    hder$ = hder$ & ("{\revtim" & nowdt$ & "}{\version2}")
    hder$ = hder$ & ("{\edmins7}{\nofpages" & Format$(m_PageCnt) & "}{\nofwords0}")
    hder$ = hder$ & "{\nofchars0}{\vern8351}}"
    If PageLandscape = False Then
       hder$ = hder$ & ("\paperw" & FnstL(PageWidth) & "\paperh" & FnstL(PageHeight))
       hder$ = hder$ & ("\margl" & FnstL(PageMarginLeft) & "\margr" & FnstL(PageMarginRight))
       hder$ = hder$ & ("\margt" & FnstL(PageMarginTop) & "\margb" & FnstL(PageMarginBottom))
    Else
       hder$ = hder$ & ("\paperw" & FnstL(PageHeight) & "\paperh" & FnstL(PageWidth))
       hder$ = hder$ & ("\margl" & FnstL(PageMarginBottom) & "\margr" & FnstL(PageMarginTop))
       hder$ = hder$ & ("\margt" & FnstL(PageMarginLeft) & "\margb" & FnstL(PageMarginRight))
    End If
    If PageNumbersOn = True Then
       hder$ = hder$ & ("\pgnstarts" & FnstL(PageNumberStart) & "\pgncont\pgnrestart")
       hder$ = hder$ & ("\pgnx" & FnstL(PageNumberMarginRight))
       hder$ = hder$ & ("\pgny" & FnstL(PageNumberMarginTop) & "\pgndec")
    End If
    If Not (Header Is Nothing) Then
       hder$ = hder$ & "\pgbrdrhead\headery720 "
    End If
    If Not (Footer Is Nothing) Then
       hder$ = hder$ & "\pgbrdrfoot\footery720 "
    End If
    hder$ = hder$ & "\widowctrl\ftnbj \sectd "
    hder$ = hder$ & RtfAddToHeader$
    If Not (Header Is Nothing) Then
       hder$ = hder$ & ("{\header" & Header.RtfParagraph & "}")
    End If
    If Not (Footer Is Nothing) Then
       hder$ = hder$ & ("{\footer" & Footer.RtfParagraph & "}")
    End If
    hder$ = hder$ & "\pmmetafile28\endnhere"
    hder$ = hder$ & "\pard\plain\ql"
    hder$ = hder$ & ("\f0\fs" & Format$(FontDefaultHeight * 2))
    hder$ = hder$ & ("\cb" & Format$(WhiteColorIndex) & "\cf" & Format$(BlackColorIndex) & " ")
    nowdt$ = ""
    RTFHeader = hder$
End Function

Private Sub Class_Terminate()
    If Not (Header Is Nothing) Then Set Header = Nothing
    If Not (Footer Is Nothing) Then Set Footer = Nothing
End Sub
