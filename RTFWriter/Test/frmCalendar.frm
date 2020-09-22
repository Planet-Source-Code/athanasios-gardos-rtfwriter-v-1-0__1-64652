VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   10005
   Icon            =   "frmCalendar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4215
   Begin VB.CommandButton CreateRTF 
      Caption         =   "Create RTF calendar"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3720
      TabIndex        =   7
      ToolTipText     =   "Select RTF File"
      Top             =   3480
      Width           =   372
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   3480
      Width           =   2655
   End
   Begin VB.ComboBox cboYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox picMover 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox F 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "2003"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cboMonth 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picCal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2745
      ScaleWidth      =   3930
      TabIndex        =   0
      Top             =   480
      Width           =   3962
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4440
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "RTF File ="
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   160
      Width           =   3855
   End
   Begin VB.Menu mnu_About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory _
    As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL As Long = 1

Dim dStartDate As Date
Dim dX As Integer
Dim Dy As Integer
Dim mfLoaded As Boolean
Dim dClickStart As Double
Dim dDifference As Double
Const cnLtGrey = &HC0C0C0
Const cnDkGrey = &H808080
Const cnBlack = &H0&
Dim mnDate As Long
Dim mDay As Long
Dim DayX As Single
Dim DayY As Single
Dim DayNumbers() As String

Private Sub cboMonth_Click()
    lblHeader.Caption = cboMonth.Text & " " & cboYear.Text
    picCal.Cls
    DrawMonthHeading
    DrawDays
End Sub

Private Sub cboYear_Click()
    lblHeader.Caption = cboMonth.Text & " " & cboYear.Text
    picCal.Cls
    DrawMonthHeading
    DrawDays
End Sub

Private Sub CreateRTF_Click()
    Dim oDocument As RTFWriter.cDocument
    Dim oPage As RTFWriter.cPage
    Dim oParagraph As RTFWriter.cParagraph
    Dim oText As RTFWriter.cText
    Dim oCell() As RTFWriter.cCell
    ReDim oCell(7, 7) As RTFWriter.cCell
    Dim cc As Long, rr As Long, sText As String
    If Text2.Text = "" Then Exit Sub
    Set oDocument = New RTFWriter.cDocument
    Set oPage = New RTFWriter.cPage
    Set oText = New RTFWriter.cText
    Set oParagraph = New RTFWriter.cParagraph
    oText.Text = cboMonth.Text & " " & cboYear.Text
    oText.Height = 20
    oText.Bold = True
    oParagraph.InsertText oText
    oPage.InsertParagraph oParagraph
    For cc = 1 To 7
        For rr = 1 To 7
            Set oCell(cc, rr) = New RTFWriter.cCell
            oCell(cc, rr).Width = 720
            Set oText = New RTFWriter.cText
            oText.ForeColorIndex = oDocument.ColorIndex(cnBlack)
            Set oParagraph = New RTFWriter.cParagraph
            oParagraph.Align = aCenter
            If rr = 1 Then
               sText = Choose(cc, "Sun", "Mon", "Teu", "Thu", "Wed", "Fri", "Sat")
               oCell(cc, rr).BackColorIndex = oDocument.ColorIndex(cnLtGrey)
            Else
               sText = DayNumbers(cc, rr)
            End If
            oText.Text = sText
            oParagraph.InsertText oText
            oCell(cc, rr).InsertParagraph oParagraph
        Next rr
    Next cc
    oPage.InsertTable oCell()
    oDocument.InsertPage oPage
     If oDocument.Save(Text2.Text) = True Then
       ShellExecute 0, vbNullString, Text2.Text, vbNullString, vbNullString, 1
    Else
       MsgBox "Error", Me.Caption
    End If
    For cc = 1 To 7
        For rr = 1 To 7
            Set oCell(cc, rr) = Nothing
        Next rr
    Next cc
    Set oText = Nothing
    Set oParagraph = Nothing
    Set oPage = Nothing
    Set oDocument = Nothing
End Sub

Private Sub Command2_Click()
    On Local Error Resume Next
    CommonDialog1.FileName = ""
    CommonDialog1.InitDir = App.Path & "\"
    With CommonDialog1
        .Filter = "RTF Files (*.RTF)|*.RTF|"
        .DialogTitle = "Save to RTF File..."
        .ShowSave
    End With
    If CommonDialog1.FileName <> "" Then
       Text2.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub F_LostFocus()
    picCal.Cls
    DrawMonthHeading
    DrawDays
End Sub

Private Sub Form_Activate()
    If Not mfLoaded Then
        mfLoaded = True
        GetScaleFactor
        DrawMonthHeading
        DrawDays
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim mYear As Long
    Dim mMonth As Long
    Dim iYear As Integer
    For iYear = 1900 To 2100
        cboYear.AddItem iYear
    Next iYear
    dStartDate = Now
    mYear = Year(dStartDate)
    mMonth = Month(dStartDate)
    mDay = Day(dStartDate)
    cboYear.ListIndex = mYear - 1900
    cboMonth.AddItem "January"
    cboMonth.AddItem "February"
    cboMonth.AddItem "March"
    cboMonth.AddItem "April"
    cboMonth.AddItem "May"
    cboMonth.AddItem "June"
    cboMonth.AddItem "July"
    cboMonth.AddItem "August"
    cboMonth.AddItem "September"
    cboMonth.AddItem "October"
    cboMonth.AddItem "November"
    cboMonth.AddItem "December"
    cboMonth.ListIndex = mMonth - 1
    lblHeader.Caption = Format$(dStartDate, "mmmm yyyy")
    Text2.Text = App.Path & "\report.rtf"
End Sub

Private Sub GetScaleFactor()
    dX = picCal.Width / 7
    Dy = picCal.Height / 7
    picMover.Height = Dy
    picMover.Width = dX
End Sub

Private Sub DrawMonthHeading()
    Dim sText As String
    Dim i As Integer
    Dim iMonth As Integer
    Dim X1 As Integer
    Dim Y1 As Integer
    On Local Error Resume Next
    iMonth = cboMonth.ListIndex + 1
    X1 = picCal.Width / 7
    Y1 = picCal.Height / 7
    picCal.Line (0, 0)-(picCal.Width, Dy), cnLtGrey, BF
    picCal.ForeColor = cnDkGrey
    picCal.Line (X1, Y1)-(X1, picCal.Height)
    picCal.Line (2 * X1, Y1)-(2 * X1, picCal.Height)
    picCal.Line (3 * X1, Y1)-(3 * X1, picCal.Height)
    picCal.Line (4 * X1, Y1)-(4 * X1, picCal.Height)
    picCal.Line (5 * X1, Y1)-(5 * X1, picCal.Height)
    picCal.Line (6 * X1, Y1)-(6 * X1, picCal.Height)
    picCal.Line (0, Y1)-(picCal.Width, Y1)
    picCal.Line (0, 2 * Y1)-(picCal.Width, 2 * Y1)
    picCal.Line (0, 3 * Y1)-(picCal.Width, 3 * Y1)
    picCal.Line (0, 4 * Y1)-(picCal.Width, 4 * Y1)
    picCal.Line (0, 5 * Y1)-(picCal.Width, 5 * Y1)
    picCal.Line (0, 6 * Y1)-(picCal.Width, 6 * Y1)
    picCal.ForeColor = cnBlack
    picCal.FontBold = True
    For i = 1 To 7
        sText = Choose(i, "Sun", "Mon", "Teu", "Thu", "Wed", "Fri", "Sat")
        picCal.CurrentY = 0.5 * (Dy - picCal.TextHeight(sText))
        picCal.CurrentX = ((i - 1) * dX) + 0.5 * (dX - picCal.TextWidth(sText))
        picCal.Print sText
    Next
    picCal.FontBold = False
End Sub

Private Sub DrawDays()
    Dim Button As Integer, Shift As Integer
    Dim nDate As Long
    Dim i As Long
    Dim iLast As Integer
    Dim iRow As Integer
    Dim iMonth As Integer
    Dim sText As String
    On Local Error Resume Next
    If Not mfLoaded Then Exit Sub
    iMonth = cboMonth.ListIndex + 1
    sText = cboYear.Text
    nDate = DateValue("01/" & Format$(iMonth, "0") & "/" & sText)
    GetLastDay iMonth, iLast
    iRow = 1
    ReDim DayNumbers(7, 7) As String
    For i = nDate To nDate + iLast - 1
        If Weekday(i) = vbSunday Then
            If i > nDate Then
                iRow = iRow + 1
            End If
        End If
        sText = Format$(Day(i), "0")
        picCal.CurrentY = (Dy * iRow) + (0.5 * (Dy - picCal.TextHeight(sText)))
        picCal.CurrentX = dX * (Weekday(i) - 1) + (0.5 * (dX - picCal.TextWidth(sText)))
        DayNumbers(Weekday(i), iRow + 1) = sText
        picCal.Print sText
    Next
End Sub

Private Sub GetLastDay(iMonth, iLast)
    Select Case iMonth
        Case 4, 6, 9, 11
            iLast = 30
        Case 1, 3, 5, 7, 8, 10, 12
            iLast = 31
        Case 2
            If Val(F.Text) Mod 4 = 0 Then iLast = 29 Else iLast = 28
    End Select
End Sub

Private Sub mnu_About_Click()
    frmAbout.Show 1
End Sub

Private Sub picCal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iCol As Integer
    Dim iRow As Integer
    Dim i As Long
    Dim R As Integer
    Dim C As Integer
    Dim iMonth As Integer
    Dim iLast As Integer
    Dim nDate As Long
    Dim sText As String
    On Local Error Resume Next
    iMonth = cboMonth.ListIndex + 1
    picMover.Visible = False
    GetLastDay iMonth, iLast
    iCol = 7 * (X) \ picCal.Width + 1
    iRow = 7 * (Y) \ picCal.Height - 1
    If iRow < 0 Then Exit Sub
    nDate = DateValue("01/" & Format$(iMonth, "0") & "/" & cboYear.Text)
    R = 0
    For i = nDate To nDate + iLast - 1
        If Weekday(i) = vbSunday Then
            If i > nDate Then
                R = R + 1
            End If
        End If
        C = Weekday(i)
        If R = iRow And C = iCol Then
            mnDate = i
            picMover.Cls
            sText = Day(mnDate)
            picMover.Left = (picCal.Left + 20) + ((C - 1) * dX)
            picMover.Top = (picCal.Top + 20) + ((R + 1) * Dy)
            If C = 7 Then picMover.Width = dX - 40 Else picMover.Width = dX
            If R = 5 Then picMover.Height = Dy - 20 Else picMover.Height = Dy
            picMover.CurrentX = 0.5 * (picMover.Width - picCal.TextWidth(sText))
            picMover.CurrentY = 0.5 * (picMover.Height - picCal.TextHeight(sText))
            picMover.Print sText
            picMover.Visible = True
            Exit For
        End If
    Next
    dClickStart = Now
End Sub

Private Sub picMover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dDifference = Now
    dDifference = dDifference - dClickStart
End Sub

Public Sub PositionPickListForm(X As Form, iTop As Integer, iLeft As Integer)
    On Local Error Resume Next
    If iTop <> 0 Or iLeft <> 0 Then
        If iTop <> 0 Then
            If iTop + X.Height > Screen.Height Then
                X.Top = iTop - X.Height
            Else
                X.Top = iTop
            End If
        End If
        If iLeft <> 0 Then
            If iLeft + X.Width > Screen.Width Then
                X.Left = iLeft - X.Width
            Else
                X.Left = iLeft
            End If
        End If
    Else
        X.Left = 0.5 * (Screen.Width - X.Width)
        X.Top = 0.5 * (Screen.Height - X.Height)
    End If
End Sub

Private Sub Text2_Change()
    Text2.ToolTipText = Text2.Text
End Sub
