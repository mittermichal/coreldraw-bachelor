Attribute VB_Name = "LengthArea"
Option Explicit


Function AreaLengthSum(sr As ShapeRange, unit As cdrUnit, ByRef l As Double, ByRef a As Double)
    Dim a_sum As Double, l_sum As Double
    
    If sr.Count = 0 Then
        Exit Function
    End If

    
    ActiveDocument.SaveSettings
    ActiveDocument.unit = unit

    Dim sh As Shape
    Set sr = sr.Duplicate.UngroupAllEx
    
    For Each sh In sr
        a_sum = a_sum + sh.DisplayCurve.Area
        l_sum = l_sum + sh.DisplayCurve.Length
    Next sh
    sr.Delete
    ActiveDocument.RestoreSettings
    l = l_sum
    a = a_sum
End Function

Sub updateForm()
    myOptimize True, True
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    Dim l As Double, a As Double
    AreaLengthSum sr, cdrMillimeter, l, a
    LengthAreaFrm.Label1.Caption = "Obsah: " & CStr(a) & " mm2"
    LengthAreaFrm.Label2.Caption = "Dlzka: " & CStr(l) & " mm"
    myOptimize True, False
End Sub

Sub showForm()
    LengthAreaFrm.Show
End Sub


