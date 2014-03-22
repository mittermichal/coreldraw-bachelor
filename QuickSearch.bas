Attribute VB_Name = "QuickSearch"

Sub eval()
    Dim Doc As Document
    Set Doc = ActiveDocument
    'MsgBox Application.Evaluate("vba.GlobalMacros.Module1.doc.FileName")CreateColorSwatch()
    MsgBox Application.Evaluate("Math") '.Module1.CreateColorSwatch()
    'MsgBox Application.Evaluate("vba.Application.ActiveDocument.ActiveLayer.CreateEllipse(0, 0, 10, 10, 90, 90, false)")
End Sub
'coment test
Public Function foo() As String
    'Set VBScript = New MSScriptControl.ScriptControl
    'FExecuteLine ("1")
    'ActiveLayer.CustomCommand
    'Dim sc As New ScriptControl
    'Call sc.Execute("vba.console.Module1")
    'Call ActiveLayer.CreateEllipse(0, 0, 10, 10, 90, 90, False)
    foo = "test"
End Function

Public Sub QuickSearch()
    QuickSearchFrm.Show
End Sub
'creates selection of objects found by query string
Sub Query(q As String, add As Boolean, subs As Boolean)
    Dim Doc As Document
    Set Doc = ActiveDocument
    Dim sr As ShapeRange
    Set sr = ActivePage.Shapes.FindShapes(Query:="" + q)
    If add Then
        sr.AddToSelection
    ElseIf subs Then
        sr.RemoveFromSelection
    Else
        sr.CreateSelection
    End If
End Sub
'https://coreldraw.com/forums/p/26065/122174.aspx
Sub myOptimize(bUse As Boolean, Optional bIsStart As Boolean = True)
    If bUse Then
        If bIsStart Then
            Optimization = True
            EventsEnabled = False
            ActiveDocument.SaveSettings
            ActiveDocument.PreserveSelection = False
        Else
            ActiveDocument.PreserveSelection = True
            ActiveDocument.RestoreSettings
            EventsEnabled = True
            Optimization = False
            ActiveWindow.Refresh
        End If
    End If
End Sub

Public Sub OnErrorDemo()
   On Error GoTo ErrorHandler   ' Enable error-handling routine.
   Dim x As Integer
   Dim y As Integer
   x = 32
   y = 1
   Dim z As Integer
   z = x / y   ' Creates a divide by zero error
   On Error GoTo 0   ' Turn off error trapping.
   On Error Resume Next   ' Defer error trapping.
   z = x / y   ' Creates a divide by zero error again
   If Err.Number = 6 Then
      ' Tell user what happened. Then clear the Err object.
      Dim Msg As String
      Msg = "There was an error attempting to divide by zero!"
      Call MsgBox(Msg, , "Divide by zero error")
      Call Err.Clear   ' Clear Err object fields.
   End If
Exit Sub      ' Exit to avoid handler.
ErrorHandler:  ' Error-handling routine.
   Select Case Err.Number   ' Evaluate error number.
      Case 6   ' Divide by zero error
         MsgBox ("You attempted to divide by zero!")
         ' Insert code to handle this error
      Case Else
         ' Insert code to handle other situations here...
   End Select
   Resume Next  ' Resume execution at same line
                ' that caused the error.
End Sub
