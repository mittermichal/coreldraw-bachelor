Attribute VB_Name = "QuickSearch"
Public Sub QuickSearchForm()
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
