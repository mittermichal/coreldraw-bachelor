Attribute VB_Name = "Intersection"
Option Explicit

Sub Overlaps()
    Dim txt As String
    Dim res As Boolean
    Dim c As CrossPoints
    Dim point As CrossPoint
    res = ActiveSelectionRange(1).Curve.IntersectsWith(ActiveSelectionRange(2).Curve)
    'c = 0
    Set c = ActiveSelectionRange(1).Curve.SubPaths(1).GetIntersections(ActiveSelectionRange(2).Curve.SubPaths(1), 0)
    For Each point In c
        'Call ActiveLayer.CreateEllipse2(point.PositionX, point.PositionY, point.Offset)
        txt = txt + "[" + CStr(point.Offset * 25.4) + " " + CStr(point.Offset2 * 25.4) + "]"
    Next
    If res Then
        MsgBox "pretina body:" + CStr(c.Count) + txt
    End If
End Sub

Sub Poly()
    'ActiveSelectionRange(1).Curve.CopyAssign (ActiveSelectionRange(1).Curve.GetPolyline(2))
    Dim s As Curve
    Set s = ActiveSelection.Shapes.First.DisplayCurve.GetCopy.GetPolyline(5).GetCopy
    ActiveLayer.CreateCurve (s)
    MsgBox CStr(ActiveSelectionRange(1).Curve.Nodes.Count) + " -> " + CStr(ActiveSelectionRange(1).Curve.GetPolyline(5).Nodes.Count)
End Sub
