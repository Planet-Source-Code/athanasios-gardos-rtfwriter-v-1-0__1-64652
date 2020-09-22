VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPage"
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

Private m_Page As String

Private Sub Class_Initialize()
    m_Page = "{"
End Sub

Friend Function RtfPage() As String
    RtfPage = m_Page + "}\page "
End Function

Public Function InsertParagraph(Paragraph As cParagraph) As Boolean
    If Paragraph Is Nothing Then Exit Function
    m_Page = m_Page + Paragraph.RtfParagraph
    InsertParagraph = True
End Function

Public Function InsertTable(oCells() As cCell) As Boolean
    Dim RtfFr1%, sFileN As String, sTmp As String
    sFileN = GetTempFile
    InsertTable = RtfTableInFile(sFileN, oCells())
    If IsFile(sFileN) = True Then
       RtfFr1% = FreeFile
       Open sFileN For Binary As #RtfFr1%
       sTmp = Space$(LOF(RtfFr1%))
       Get #RtfFr1%, , sTmp
       Close #RtfFr1%
       m_Page = m_Page + sTmp
    End If
    Call DeleteFile(sFileN)
End Function

Private Function RtfTableInFile(sFileN As String, oCells() As cCell) As Boolean
    Dim mRow As Integer
    Dim mCol As Integer
    Dim iCntC As Integer
    Dim iCntR As Integer
    Dim par1 As Integer
    Dim sPictureFile As String
    Dim i As Integer, frr%, th$
    Dim head$, mHead$
    Dim px&, v$, nn$
    Dim sTmp As String
    Dim appi2 As Integer, lTok As Long
    Dim sArr() As String, sScale As Single
    If sFileN = "" Then Exit Function
    mCol = UBound(oCells, 1)
    mRow = UBound(oCells, 2)
    If mRow = 0 Or mCol = 0 Then Exit Function
    '========
    frr% = FreeFile
    Call DeleteFile(sFileN)
    Open sFileN For Binary As #frr%
    par1 = 1
    head$ = ""
    mHead$ = ""
    mHead$ = mHead$ + "\trowd\trgaph0"
    mHead$ = mHead$ + "\trrh0"
    mHead$ = mHead$ + "\trleft0"
    For iCntR = 1 To mRow
        head$ = head$ + mHead$
        Put #frr%, , head$ & vbCrLf
        head$ = ""
        px& = 0
        For iCntC = 1 To mCol
            If oCells(iCntC, iCntR).Border.TopVisible = True Then
               head$ = head$ + "\clbrdrt\brdrs\brdrw" + FnstM(oCells(iCntC, iCntR).Border.TopWidth) & " "
            Else
               head$ = head$ + "\clbrdrt\brdrnone"
            End If
            If oCells(iCntC, iCntR).Border.BottomVisible = True Then
               head$ = head$ + "\clbrdrb\brdrs\brdrw" + FnstM(oCells(iCntC, iCntR).Border.BottomWidth) & " "
            Else
               head$ = head$ + "\clbrdrb\brdrnone"
            End If
            If oCells(iCntC, iCntR).Border.LeftVisible = True Then
               head$ = head$ + "\clbrdrl\brdrs\brdrw" + FnstM(oCells(iCntC, iCntR).Border.LeftWidth) & " "
            Else
               head$ = head$ + "\clbrdrl\brdrnone"
            End If
            If oCells(iCntC, iCntR).Border.RightVisible = True Then
               head$ = head$ + "\clbrdrr\brdrs\brdrw" + FnstM(oCells(iCntC, iCntR).Border.RightWidth) & " "
            Else
               head$ = head$ + "\clbrdrr\brdrnone"
            End If
            If oCells(iCntC, iCntR).BackColorIndex <> WhiteColorIndex Then
               head$ = head$ + "\clcbpat" + FnstM(oCells(iCntC, iCntR).BackColorIndex)
            End If
            px& = px& + CLng(oCells(iCntC, iCntR).Width)
            If px& = 0 Then v$ = "0" Else v$ = FnstL(px&)
            head$ = head$ + "\cellx" + v$
            Put #frr%, , head$ & vbCrLf
            head$ = ""
        Next iCntC
        For iCntC = 1 To mCol
            nn$ = oCells(iCntC, iCntR).RtfCell
            If iCntC = mCol Then
               head$ = head$ + "\pard\plain\intbl" + nn$ + "\cell\row"
            Else
               head$ = head$ + "\pard\plain\intbl" + nn$ + "\cell"
            End If
            Put #frr%, , head$ & vbCrLf
            head$ = ""
        Next iCntC
    Next iCntR
    head$ = head$ + "\pard"
    If par1 <> 0 Then head$ = head$ + String$(par1, "}")
    Put #frr%, , head$ & vbCrLf
    head$ = ""
    Close #frr%
    RtfTableInFile = IsFile(sFileN)
End Function

