Attribute VB_Name = "CircleFromArc"
Option Explicit

'TODO: make function that returns center of circle from three points
Sub CircleFromArc()
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double
    Dim r As Double
    Dim Cx As Double, Cy As Double
    Dim cp As Double 'vektorovy sucin
    'ActiveSelection.Curve.Nodes
    ActiveSelectionRange(1).Curve.Nodes(1).GetPosition x1, y1
    ActiveSelectionRange(1).Curve.Nodes(2).GetPosition x2, y2
    ActiveSelectionRange(1).Curve.Nodes(3).GetPosition x3, y3
    'http://en.wikipedia.org/wiki/Radius
    'r = Sqr((Sq(x2 - x1) + Sq(y2 - y1)) * (Sq(x2 - x3) + Sq(y2 - y3)) * (Sq(x3 - x1) + Sq(y3 - y1))) / (2 * Abs(x1 * y2 + x2 * y3 + x3 * y1 - x1 * y3 - x2 * y1 - x3 * y2))
    'http://en.wikipedia.org/wiki/Circumscribed_circle#Cartesian_coordinates
    cp = 2 * (x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2))
    Cx = ((Sq(x1) + Sq(y1)) * (y2 - y3) + (Sq(x2) + Sq(y2)) * (y3 - y1) + (Sq(x3) + Sq(y3)) * (y1 - y2)) / cp
    Cy = ((Sq(x1) + Sq(y1)) * (x3 - x2) + (Sq(x2) + Sq(y2)) * (x1 - x3) + (Sq(x3) + Sq(y3)) * (x2 - x1)) / cp
    r = Sqr(Sq(x1 - Cx) + Sq(y1 - Cy)) 'precision?
    Call ActiveLayer.CreateEllipse2(Cx, Cy, r)
End Sub

Private Function Sq(x As Double) As Double
    Sq = x * x
End Function
