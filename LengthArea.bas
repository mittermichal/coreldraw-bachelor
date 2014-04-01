Attribute VB_Name = "LengthArea"
Option Explicit

Function AreaSum(unit As cdrUnit) As Double
    Dim sum As Double
    ActiveDocument.SaveSettings
    ActiveDocument.unit = unit
    'Set sr = sr.Duplicate
    'Set sr = sr.UngroupAllEx
    Dim sh As Shape
    Dim sr As ShapeRange
    
    Set sr = ActiveSelectionRange.Duplicate.UngroupAllEx
    'Set sr = sh.Shapes.All
    'sr.AddRange sh.Shapes.All.UngroupAllEx
    For Each sh In sr
        sum = sum + sh.DisplayCurve.Area
    Next sh
    sr.Delete
    ActiveDocument.RestoreSettings
    AreaSum = sum
End Function

Sub updateForm()
    LengthAreaFrm.Label1.Caption = CStr(AreaSum(cdrMillimeter)) & "mm2"
End Sub

Sub showForm()
    LengthAreaFrm.Show
End Sub

Sub Test()
    Dim sh As Shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange.UngroupAllEx
    'ActiveShape.Shapes.All.UngroupAllEx.CreateSelection
End Sub
