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
