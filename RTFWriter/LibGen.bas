Attribute VB_Name = "LibGen"
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

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
            (ByVal hwnd As Long, _
             ByVal lpHelpFile As String, _
             ByVal wCommand As Long, _
             ByVal dwData As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal strBuffer As String, ByVal nBufLen As Long) As Long

Private Const MAX_PATH = 260
Public Const SW_SHOWNORMAL As Long = 1
Public Const HELP_CONTENTS = &H3&

Public BlackColorIndex As Long
Public WhiteColorIndex As Long

Public Const sEmpty = ""
Public Const cMaxPath = 260
Public Const c_CompanyName As String = "Gardos Software"

Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type POINTAPI
        x As Long
        y As Long
End Type

Type SIZE
        cx As Long
        cy As Long
End Type

Type METAFILEPICT
        mm As Long
        xExt As Long
        yExt As Long
        hMF As Long
End Type

Type METARECORD
        rdSize As Long
        rdFunction As Integer
        rdParm(1) As Integer
End Type

Declare Function CreateMetaFile& Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpstring As String)
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function Arc& Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long)
Declare Function Chord& Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long)
Declare Function CloseClipboard& Lib "user32" ()
Declare Function CloseMetaFile& Lib "gdi32" (ByVal hMF As Long)
Declare Function CreateHatchBrush& Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long)
Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Declare Function DeleteMetaFile& Lib "gdi32" (ByVal hMF As Long)
Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Declare Function DrawFocusRect& Lib "user32" (ByVal hdc As Long, lpRect As Rect)
Declare Function Ellipse& Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Declare Function EmptyClipboard& Lib "user32" ()
Declare Function EnumMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long, ByVal lpCallbackFunc As Long, ByVal lpClientData As Long) As Long
Declare Function GetClientRect& Lib "user32" (ByVal hwnd As Long, lpRect As Rect)
Declare Function GetMetaFileBitsEx& Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any)
Declare Function GlobalAlloc& Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long)
Declare Function GlobalFree& Lib "kernel32" (ByVal hMem As Long)
Declare Function GlobalLock& Lib "kernel32" (ByVal hMem As Long)
Declare Function GetObjectType& Lib "gdi32" (ByVal hgdiobj As Long)
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem As Long)
Declare Function InflateRect& Lib "user32" (lpRect As Rect, ByVal x As Long, ByVal y As Long)
Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long)
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Declare Function OpenClipboard& Lib "user32" (ByVal hwnd As Long)
Declare Function Pie& Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long)
Declare Function PlayMetaFile& Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long)
Declare Function PlayMetaFileRecord& Lib "gdi32" (ByVal hdc As Long, ByVal lpHandletable As Long, lpMetaRecord As Any, ByVal nHandles As Long)
Declare Function Polyline& Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long)
Declare Function Polygon& Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long)
Declare Function Rectangle& Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Declare Function RestoreDC& Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long)
Declare Function SaveDC& Lib "gdi32" (ByVal hdc As Long)
Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long)
Declare Function SetClipboardData& Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long)
Declare Function SetMapMode& Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long)
Declare Function SetMetaFileBitsEx& Lib "gdi32" (ByVal nSize As Long, lpData As Byte)
Declare Function SetMetaFileBitsBuffer& Lib "gdi32" Alias "SetMetaFileBitsEx" (ByVal nSize As Long, ByVal lpData As Long)
Declare Function SetPolyFillMode& Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long)
Declare Function SetViewportExtEx& Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE)
Declare Function SetViewportOrgEx& Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI)
Declare Function SetWindowOrgEx& Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI)
Declare Function SetWindowExtEx& Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE)

Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long

Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long

Sub FileNotFound(fl$)
    MsgBox "file not found =" & fl$, vbCritical, "Error"
End Sub

Function GetTempFile(Optional Prefix As String, Optional PathName As String) As String
    If Prefix = sEmpty Then Prefix = "~gv"
    If PathName = sEmpty Then PathName = GetTempDir
    Dim sRet As String
    sRet = String(cMaxPath, 0)
    Call GetTempFileName(PathName, Prefix, 0, sRet)
    Call AllagiCha(sRet, Chr$(0), Chr$(32))
    sRet = RTrim$(sRet)
    GetTempFile = sRet
End Function

Function GetTempDir() As String
    Dim sRet As String
    Dim c As Long
    Dim sDir As String
    sRet = String(cMaxPath, 0)
    c = GetTempPath(cMaxPath, sRet)
    If c <> 0 Then
       sDir = Left$(sRet, c)
       GetTempDir = NormalizePath(sDir)
    End If
End Function

Function UboundVarX(a() As Variant) As Long
    On Local Error GoTo Lab_Error
    UboundVarX = UBound(a)
    Exit Function
Lab_Error:
    UboundVarX = 0
End Function

Function UboundLngX(a() As Long) As Long
    On Local Error GoTo Lab_Error
    UboundLngX = UBound(a)
    Exit Function
Lab_Error:
    UboundLngX = 0
End Function

Function UboundSngX(a() As Single) As Long
    On Local Error GoTo Lab_Error
    UboundSngX = UBound(a)
    Exit Function
Lab_Error:
    UboundSngX = 0
End Function

Function UboundByteX(a() As Byte) As Long
    On Local Error GoTo Lab_Error
    UboundByteX = UBound(a)
    Exit Function
Lab_Error:
    UboundByteX = 0
End Function

Function UboundStrX(sArray() As String) As Integer
    On Local Error GoTo Lab_Err
    UboundStrX = UBound(sArray)
    Exit Function
Lab_Err:
    UboundStrX = 0
End Function

Function QBColorX(clr As Integer) As Long
    Dim b As Integer
    b = clr Mod 16
    QBColorX = QBColor(b)
End Function

Function SystemDirectory() As String
    Dim buffer As String * 512, Length As Long
    Length = GetSystemDirectory(buffer, Len(buffer))
    SystemDirectory = Left$(buffer, Length)
End Function

Sub Main()

End Sub

