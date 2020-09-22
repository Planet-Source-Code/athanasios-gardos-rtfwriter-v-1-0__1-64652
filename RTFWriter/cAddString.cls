VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Dim m_Diast As Long
Dim m_MaxBytes As Long
Dim m_Buffer() As Byte
Dim m_pointer As Long

Public Property Let MaxBytes(ByVal v As Long)
    If v <= 0 Then Exit Property
    m_MaxBytes = v
    m_Diast = 2 * MaxBytes - 1
    ReDim m_Buffer(m_Diast) As Byte
End Property

Public Property Get MaxBytes() As Long
    MaxBytes = m_MaxBytes
End Property

Function BeginAdd() As Boolean
    m_pointer = 0
    BeginAdd = True
End Function

Public Function Length() As Long
    Length = m_pointer / 2
End Function

Function AddString(sString As String) As Boolean
    Dim llen As Long
    llen = LenB(sString)
    If m_pointer + llen < m_Diast Then
       If llen <> 0 Then
          Call CopyMemory(ByVal VarPtr(m_Buffer(m_pointer)), ByVal StrPtr(sString), llen)
          m_pointer = m_pointer + llen
       End If
       AddString = True
    Else
       m_Diast = m_Diast * 2
       ReDim Preserve m_Buffer(m_Diast) As Byte
       AddString = AddString(sString)
    End If
End Function

Function Add2Strings(sString1 As String, sString2 As String) As Boolean
    If AddString(sString1) = False Then Exit Function
    Add2Strings = AddString(sString2)
End Function

Function CurString() As String
    If m_pointer = 0 Then Exit Function
    CurString = Space$(m_pointer / 2)
    Call CopyMemory(ByVal StrPtr(CurString), ByVal VarPtr(m_Buffer(0)), m_pointer)
End Function

Function EndAdd() As String
    If m_pointer = 0 Then Exit Function
    EndAdd = Space$(m_pointer / 2)
    Call CopyMemory(ByVal StrPtr(EndAdd), ByVal VarPtr(m_Buffer(0)), m_pointer)
    Call BeginAdd
End Function

Private Sub Class_Initialize()
    m_pointer = 0
    MaxBytes = 128000
End Sub

Private Sub Class_Terminate()
    m_pointer = 0
    Erase m_Buffer
End Sub
