Attribute VB_Name = "Guides"
Option Explicit

Sub GuidesFromBoundingBox()
    'Call ActivePage.GuidesLayer.CreateGuide( _
        ActiveSelectionRange.BoundingBox.Left, _
        ActiveSelectionRange.BoundingBox.Top, _
        ActiveSelectionRange.BoundingBox.Left, _
        ActiveSelectionRange.BoundingBox.Bottom)

    'Call ActivePage.GuidesLayer.CreateGuide( _
        ActiveSelectionRange.BoundingBox.Left, _
        ActiveSelectionRange.BoundingBox.Top, _
        ActiveSelectionRange.BoundingBox.Right, _
        ActiveSelectionRange.BoundingBox.Top)

    'Call ActivePage.GuidesLayer.CreateGuide( _
        ActiveSelectionRange.BoundingBox.Right, _
        ActiveSelectionRange.BoundingBox.Top, _
        ActiveSelectionRange.BoundingBox.Right, _
        ActiveSelectionRange.BoundingBox.Bottom)

    'Call ActivePage.GuidesLayer.CreateGuide( _
        ActiveSelectionRange.BoundingBox.Left, _
        ActiveSelectionRange.BoundingBox.Bottom, _
        ActiveSelectionRange.BoundingBox.Right, _
        ActiveSelectionRange.BoundingBox.Bottom)
    'activeselectionrange sa pravd. zmeni po vytvoreni vodiacej ciary
    Call ActivePage.GuidesLayer.CreateGuideAngle( _
        ActiveSelectionRange.BoundingBox.Left, _
        ActiveSelectionRange.BoundingBox.Top, _
        0)
    Call ActivePage.GuidesLayer.CreateGuideAngle( _
        ActiveSelectionRange.BoundingBox.Left, _
        ActiveSelectionRange.BoundingBox.Top, _
        90)
    Call ActivePage.GuidesLayer.CreateGuideAngle( _
        ActiveSelectionRange.BoundingBox.Right, _
        ActiveSelectionRange.BoundingBox.Bottom, _
        0)
    Call ActivePage.GuidesLayer.CreateGuideAngle( _
        ActiveSelectionRange.BoundingBox.Right, _
        ActiveSelectionRange.BoundingBox.Bottom, _
        90)
    
End Sub
