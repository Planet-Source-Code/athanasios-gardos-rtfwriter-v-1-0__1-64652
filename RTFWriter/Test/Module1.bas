Attribute VB_Name = "Module1"
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

Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory _
    As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL As Long = 1
Public Const HELP_CONTENTS = &H3&

Sub OpenWeb(sWebAddr As String)
    If Len(sWebAddr) = 0 Then Exit Sub
    Call ShellExecute(0, "open", sWebAddr, vbNullString, CurDir$, SW_SHOWNORMAL)
End Sub

Public Sub SendMail(Optional Address As String, _
    Optional Subject As String, Optional Body As String, _
    Optional cc As String, Optional BCC As String, Optional hwnd As Long)
    Dim strCommand As String
    If Len(Subject) Then strCommand = "&Subject=" & Subject
    If Len(Body) Then strCommand = strCommand & "&Body=" & Body
    If Len(cc) Then strCommand = strCommand & "&CC=" & cc
    If Len(BCC) Then strCommand = strCommand & "&BCC=" & BCC
    strCommand = "mailto:" & Address & strCommand
    Call ShellExecute(hwnd, "open", strCommand, _
        vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

