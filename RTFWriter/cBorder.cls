VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBorder"
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

Public TopWidth As Long
Public LeftWidth As Long
Public BottomWidth As Long
Public RightWidth As Long
Public TopVisible As Boolean
Public LeftVisible As Boolean
Public BottomVisible As Boolean
Public RightVisible As Boolean
Public BorderType As gsRTFBorderType
Public BorderColorIndex As Long

Public Enum gsRTFBorderType
    bSingle = 1
    bDouble = 2
    bThick = 3
    bShadow = 4
    bDot = 5
    bHairline = 6
End Enum

Private Sub Class_Initialize()
    TopWidth = 20
    LeftWidth = 20
    BottomWidth = 20
    RightWidth = 20
    TopVisible = False
    LeftVisible = False
    BottomVisible = False
    RightVisible = False
    BorderType = gsRTFBorderType.bSingle
    BorderColorIndex = BlackColorIndex
End Sub
