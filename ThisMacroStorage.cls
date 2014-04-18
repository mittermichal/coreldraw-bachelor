VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_QueryDocumentClose(ByVal Doc As Document, Cancel As Boolean)

End Sub
'https://www.youtube.com/watch?v=MuMCmX8vWVw
Private Sub GlobalMacroStorage_SelectionChange()
If ActiveDocument Is Nothing Then Exit Sub
If Not LengthAreaFrm.Visible Then Exit Sub
If ActiveSelectionRange.Count = 0 Then Exit Sub
If LengthAreaFrm.CheckBox1.Value Then
     'myOptimize True, True
     'EventsEnabled = False
     LengthArea.updateForm
     'EventsEnabled = True
     'myOptimize True, False
End If
End Sub

