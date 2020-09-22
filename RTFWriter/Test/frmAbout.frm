VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2130
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   120
         Top             =   1320
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "http://business.hol.gr/gardos/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   5
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "gardos@hol.gr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1800
         MousePointer    =   10  'Up Arrow
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   " Athanasios Gardos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   360
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   4200
      End
   End
End
Attribute VB_Name = "frmAbout"
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

Const c_DefURL As String = "http://business.hol.gr/gardos/"

Private Sub Label3_Click()
    Screen.MousePointer = vbHourglass
    Call SendMail("gardos@hol.gr")
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Label5_Click()
    Screen.MousePointer = vbHourglass
    Call OpenWeb(c_DefURL)
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

