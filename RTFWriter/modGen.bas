Attribute VB_Name = "ModGen"
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

Function LongToBytes(myLong As Long, a1 As Byte, b1 As Byte, c1 As Byte, d1 As Byte)
    Call CopyMemory(ByVal VarPtr(a1), ByVal VarPtr(myLong), 1&)
    Call CopyMemory(ByVal VarPtr(b1), ByVal VarPtr(myLong) + 1, 1&)
    Call CopyMemory(ByVal VarPtr(c1), ByVal VarPtr(myLong) + 2, 1&)
    Call CopyMemory(ByVal VarPtr(d1), ByVal VarPtr(myLong) + 3, 1&)
End Function

Function Fnexist(sFile As String) As Single
    Dim TempAttr As Long
    Dim myDir As String
    On Local Error GoTo ErrorFileExist
    If InStr(UCase$(sFile), "\NUL") > 1 Then
       myDir = Mid$(sFile, 1, Len(sFile) - 4)
       If IsDirectory(myDir) = True Then
          Fnexist = 0
       Else
          Fnexist = 1
       End If
    ElseIf InStr(sFile, "*.") <> 0 Then
       If Dir(sFile, vbNormal) <> "" Then
          Fnexist = 0
       Else
          Fnexist = 1
       End If
    Else
       TempAttr = GetAttr(sFile)
       If TempAttr = vbNormal Then Fnexist = 0
    End If
    Exit Function
ErrorFileExist:
    Fnexist = 1
End Function

Function IsDirectory(myDir As String) As Boolean
    Dim sDir As String
    sDir = Trim$(myDir)
    If sDir = "" Then Exit Function
    If sDir = "\" Then sDir = Mid$(CurDir, 1, 3)
    If Right$(sDir, 1) = "\" Then
       sDir = Mid$(sDir, 1, Len(sDir) - 1)
    End If
    IsDirectory = False
    If sDir = "" Then Exit Function
    Dim TempAttr As Long
    On Local Error GoTo ErrorDirExist
    TempAttr = GetAttr(sDir)
    If (TempAttr And vbDirectory) = vbDirectory Then
       IsDirectory = True
    End If
    Exit Function
ErrorDirExist:
    IsDirectory = False
End Function

Function NormalizePath(sPath As String) As String
    NormalizePath = sPath
    If sPath = "" Then Exit Function
    If Right$(sPath, 1) <> "\" Then
       NormalizePath = sPath & "\"
    End If
End Function

Function DeNormalizePath(sPath As String) As String
    DeNormalizePath = sPath
    If sPath = "" Then Exit Function
    If Right$(sPath, 1) = "\" Then
       DeNormalizePath = Mid$(sPath, 1, Len(sPath) - 1)
    End If
End Function

Sub DeleteFile(sFile As String)
    On Local Error Resume Next
    Kill sFile
End Sub

Function FShortPathName(sFileName As String) As String
    
    Dim l As Long
    Dim Short As String
    Short = Space$(1024)
    l = GetShortPathName(sFileName, Short, 1024)
    FShortPathName = Left$(Short, l)
    
End Function


Function IsFile(sFile As String) As Boolean
    Dim aLen As Long
    On Local Error Resume Next
    aLen = FileLen(sFile)
    If aLen <> 0 Then
       IsFile = True
    End If
End Function

Sub replacement(k$, apo$, se$)
    Dim ReplOnce%
    Dim ap As Single
    Dim start As Single
    Dim lf$
    Dim yp As Single
    Dim rg$
    If k$ = "" Or apo$ = "" Then Exit Sub
    ReplOnce% = 0
    If se$ <> "" Then If InStr(se$, apo$) <> 0 Then ReplOnce% = 1
    ap = 1: start = 1
    While ap <> 0
        ap = InStr(start, k$, apo$)
        If ap > 0 Then
            If ap > 1 Then lf$ = Left$(k$, ap - 1)
            yp = ap + Len(apo$) - 1
            If yp < Len(k$) Then rg$ = Right$(k$, Len(k$) - yp)
            k$ = lf$ + se$ + rg$
            lf$ = "": rg$ = ""
            If ReplOnce% = 1 Then start = ap + Len(se$)
        End If
    Wend
End Sub

Sub AllagiCha(kh$, a1$, a2$)
    Dim ap As Long, ap1 As Long
    ap = 1
    ap1 = 1
    While ap <> 0
       ap = InStr(ap1, kh$, a1$)
       If ap <> 0 Then
      Mid$(kh$, ap, 1) = a2$
      ap1 = ap + 1
       End If
    Wend
    a1$ = ""
    a2$ = ""
End Sub

Function GetToken(sString As String, sArr() As String, sDel As String) As Long
    Dim sTmp As String, lCount As Long
    Dim ap As Long, ap1 As Long
    sTmp = sString
    If sTmp = "" Then
       ReDim sArr(0) As String
    Else
       If Right$(sTmp, 1) <> sDel Then sTmp = sTmp & sDel
       ap = 1
       ap1 = 1
       While ap <> 0
          ap = InStr(ap1, sTmp, sDel)
          If ap <> 0 Then
             lCount = lCount + 1
             ReDim Preserve sArr(lCount) As String
             If ap > ap1 Then
                sArr(lCount) = Mid$(sTmp, ap1, ap - ap1)
             Else
                sArr(lCount) = ""
             End If
             ap1 = ap + 1
          End If
       Wend
    End If
    GetToken = lCount
End Function

Sub Swap(a As Variant, b As Variant)
    Dim c As Variant
    c = b
    b = a
    a = c
End Sub

Public Function FnstM(x As Integer) As String
    FnstM = Format$(x)
    If x = 0 Then FnstM = "0"
End Function

Public Function FnstL(x As Long) As String
    FnstL = Format$(x)
    If x = 0 Then FnstL = "0"
End Function

Function CDblx(var1 As Variant) As Double
    On Local Error Resume Next
    Dim a As Double
    a = CDbl(var1)
    CDblx = a
End Function

Function LongToInteger(x As Long) As Integer
    Dim i As Long
    i = x And CLng(65535)
    If i > 32767 Then i = i - 65536
    LongToInteger = i
End Function

Function CLngx(var1 As Variant) As Long
    On Local Error Resume Next
    Dim a As Long
    a = CLng(var1)
    CLngx = a
End Function

Function CSngx(var1 As Variant) As Single
    On Local Error Resume Next
    Dim a As Single
    a = CSng(var1)
    CSngx = a
End Function

Function CIntx(var1 As Variant) As Integer
    On Local Error Resume Next
    Dim a As Integer
    a = CInt(var1)
    CIntx = a
End Function

Function fni(x As Variant) As Single
    If x > 32767 Then
       fni = x - 65536
    Else
       fni = x
    End If
End Function

Function fnINrange(num#, min#, max#) As Double
    If num# < min# Then
       num# = min#
    ElseIf num# > max# Then
       num# = max#
    End If
    fnINrange# = num#
End Function

Function fnr(int1 As Variant) As Single
    If int1 < 0 Then
       fnr = 65536 + int1
    Else
       fnr = int1
    End If
End Function

Function fnst(a As Single) As String
    Dim b$
    If a > 0 Then
       b$ = Str$(a)
       fnst$ = Mid$(b$, 2, Len(b$) - 1)
       If a < 1 Then fnst$ = Chr$(48) + Mid$(b$, 2, Len(b$) - 1)
    Else
       If a <> 0 Then
          fnst$ = Str$(a)
       Else
          fnst$ = Space$(1)
       End If
    End If
    b$ = ""
End Function

Function fnsti(a As Integer) As String
    Dim b$
    If a > 0 Then
       b$ = Str$(a)
       fnsti$ = Mid$(b$, 2, Len(b$) - 1)
    Else
       If a <> 0 Then
          fnsti$ = Str$(a)
       Else
          fnsti$ = Space$(1)
       End If
    End If
    b$ = ""
End Function

Function fnstd(a#) As String
    Dim a1$
    If a# > 0 Then
       a1$ = Str$(a#): fnstd$ = Mid$(a1$, 2, Len(a1$) - 1)
       If a# < 1 Then fnstd$ = Chr$(48) + Mid$(a1$, 2, Len(a1$) - 1)
    Else
       If a# <> 0 Then
          fnstd$ = Str$(a#)
       Else
          fnstd$ = Space$(1)
       End If
    End If
    a1$ = ""
End Function

Sub findinstrlast(apo As Single, keim$, fch$, aplast As Single)
    Dim apl As Single, ap1 As Single
        
    aplast = 0
    If keim$ = "" Or fch$ = "" Then Exit Sub
    apl = 1: ap1 = apo: If ap1 = 0 Then ap1 = 1
    While apl <> 0
       apl = InStr(ap1, keim$, fch$)
       If apl <> 0 Then
          aplast = apl
          ap1 = apl + 1
       End If
    Wend
End Sub

Function FnrHex(int1 As Long) As Single
    If int1 < 0 Then
       FnrHex = 65536 + int1
    Else
       FnrHex = int1
    End If
End Function

Rem ****

