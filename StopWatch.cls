VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngStart As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub StartTimer()
    mlngStart = GetTickCount
End Sub

Public Function EndTimer() As Long
    EndTimer = (GetTickCount - mlngStart)
End Function

Public Function EndSeconds() As String
    EndSeconds = CStr(EndTimer / 1000#) & " s"
End Function

