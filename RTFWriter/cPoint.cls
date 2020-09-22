VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPoint"
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

Public x As Long
Public y As Long

Friend Property Get RtfPoint() As String
    RtfPoint = "\dpptx" & Format$(x) & "\dppty" & Format$(y)
End Property


