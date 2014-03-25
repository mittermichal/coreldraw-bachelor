Attribute VB_Name = "Guides"
Option Explicit

Sub GuidesFromBoundingBox()
    Dim bBox As Rect
    Set bBox = ActiveSelectionRange.BoundingBox
    'todo: begincommandgroup
    'activeselectionrange sa pravd. zmeni po vytvoreni vodiacej ciary
    Call ActivePage.GuidesLayer.CreateGuideAngle( _
        bBox.Left, _
        bBox.Top, _
        0)
    Call ActivePage.GuidesLayer.CreateGuideAngle( _
        bBox.Left, _
        bBox.Top, _
        90)
    Call ActivePage.GuidesLayer.CreateGuideAngle( _
        bBox.Right, _
        bBox.Bottom, _
        0)
    Call ActivePage.GuidesLayer.CreateGuideAngle( _
        bBox.Right, _
        bBox.Bottom, _
        90)
    'todo: activate layer
End Sub

'todo: Sub CreateGuideWithoutDuplicates

Sub GuidesOnTangents()
 Dim x As Double, y As Double
 Dim a As Double
 Dim sr As ShapeRange
 Dim sh As Shape
 Dim sp As SubPath
 Dim s As Segment
 Set sr = ActiveSelectionRange
 If sr.Count = 0 Then
    Exit Sub
 End If
 ActiveDocument.BeginCommandGroup "Guides On Tangents"
 For Each sh In ActiveSelectionRange.Shapes
    Set sh = sh.Duplicate
    sh.ConvertToCurves
    For Each sp In sh.Curve.SubPaths
       For Each s In sp.Segments
           s.GetPointPositionAt x, y, 0.5, cdrRelativeSegmentOffset
           a = s.GetTangentAt(0.5, cdrRelativeSegmentOffset)
           ActivePage.GuidesLayer.CreateGuideAngle x, y, a
       Next s
    Next sp
    sh.Delete
 Next sh
 ActiveDocument.EndCommandGroup
 sr.CreateSelection
End Sub

