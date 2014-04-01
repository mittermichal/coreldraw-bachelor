Attribute VB_Name = "Guides"
Option Explicit

Sub GuidesFromBoundingBox()
    Dim bBox As Rect
    Set bBox = ActiveSelectionRange.BoundingBox
    'todo: begincommandgroup
    'activeselectionrange sa pravd. zmeni po vytvoreni vodiacej ciary
    ActivePage.GuidesLayer.CreateGuideAngle bBox.Left, bBox.Top, 0
    ActivePage.GuidesLayer.CreateGuideAngle bBox.Left, bBox.Top, 90
    ActivePage.GuidesLayer.CreateGuideAngle bBox.Right, bBox.Bottom, 0
    ActivePage.GuidesLayer.CreateGuideAngle bBox.Right, bBox.Bottom, 90
    'todo: activate layer
End Sub

'todo: Sub CreateGuideWithoutDuplicates

Sub GuidesOnTangents()
 Dim x As Double, y As Double
 Dim a As Double
 Dim sr As ShapeRange
 Dim sh As Shape
 Dim duplicated As Boolean
 duplicated = False
 Dim sp As SubPath
 Dim s As Segment
 
 Set sr = ActiveSelectionRange
 If sr.Count = 0 Then
    Exit Sub
 End If
 
 Dim offset As Double
 offset = StrToDbl(GetSetting("CorelDrawBachelor", "Guides", "tangentOffset", "0.5"))
 
 
 ActiveDocument.BeginCommandGroup "Guides On Tangents"
 For Each sh In ActiveSelectionRange.Shapes
 
    If sh.Type <> cdrCurveShape Then
        Set sh = sh.Duplicate
        duplicated = True
        sh.ConvertToCurves
    End If
    
    For Each sp In sh.Curve.SubPaths
       For Each s In sp.Segments
           s.GetPointPositionAt x, y, offset, cdrRelativeSegmentOffset
           a = s.GetTangentAt(offset, cdrRelativeSegmentOffset)
           ActivePage.GuidesLayer.CreateGuideAngle x, y, a
       Next s
    Next sp
    
    If duplicated Then
        sh.Delete
    End If
 Next sh
 ActiveDocument.EndCommandGroup
 sr.CreateSelection
End Sub

Sub FrmShow()
    GuidesFrm.Show
End Sub

Private Function StrToDbl(s As String)
    StrToDbl = Val(Replace(s, ".", ","))
End Function
