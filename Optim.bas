Attribute VB_Name = "Optim"
Option Explicit

Public Sub myOptimize(bUse As Boolean, Optional bIsStart As Boolean = True)
    If bUse Then
        If bIsStart Then
            Application.Optimization = True
            EventsEnabled = False
            ActiveDocument.SaveSettings
            ActiveDocument.PreserveSelection = False
        Else
            ActiveDocument.PreserveSelection = True
            ActiveDocument.RestoreSettings
            EventsEnabled = True
            Application.Optimization = False
            ActiveWindow.Refresh
        End If
    End If
End Sub

Public Sub Refresh()
    myOptimize True, True
End Sub
