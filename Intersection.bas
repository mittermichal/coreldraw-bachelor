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
        txt = txt + "[" + CStr(point.offset * 25.4) + " " + CStr(point.Offset2 * 25.4) + "]"
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

Sub Intersections()
 Dim x As Double, y As Double
 Dim asr As ShapeRange, sr As ShapeRange
 Dim sh As Shape, sh1 As Shape, sh2 As Shape
 Dim sp1 As SubPath, sp2 As SubPath
 Dim cs As CrossPoints
 Dim cp As CrossPoint
 Dim i As Integer, j As Integer, k As Integer, l As Integer, c As Long
 Dim cps As Long
 Dim sw As StopWatch
 
 
Set asr = ActiveSelectionRange
 If asr.Count = 0 Then
    Exit Sub
 End If
 
Set sw = New StopWatch
 sw.StartTimer
 myOptimize False, True

 

 
 'Dim offset As Double
 'offset = StrToDbl(GetSetting("CorelDrawBachelor", "Intersections", "offset", "0.5"))
 
 ActiveDocument.unit = cdrMillimeter
 ActiveDocument.BeginCommandGroup "Intersections"
 Set sr = ActiveSelection.Duplicate.UngroupAllEx
 'Set asr = sh.Shapes.All.UngroupAllEx
 'Set sr = asr.Shapes.All.BreakApartEx
 'ActiveSelection.Shapes.All.
 'sr.Shapes.All.ConvertToCurves
 
 
 
 Dim bsr As ShapeRange
 Dim absr As ShapeRange
 Dim tmp_sr As ShapeRange
 'Dim n As Long
 'n = sr.Count
 'For i = 1 To sr.Count
 'While i < 0
    
 'Wend
 'sr.BreakApart
 'Next i
 'Set sr = ActiveSelectionRange
 
 If False Then
 For i = 1 To sr.Count
    'Set absr = sr(i).BreakApartEx
    'sr.AddRange absr
    'sr.Remove i
    'Dim n As Long, k As Long
    'k = sr(i).DisplayCurve.SubPaths.Count
    'sr(i).BreakApartEx
    'For n = 1 To k
     '   Set sh = sr(i).Previous
    '    sr.add sh
    'Next n
    

    Dim tmp_sh As Shape
    Set tmp_sr = CreateShapeRange
    If sr(i).DisplayCurve.SubPaths.Count > 1 Then
        Set tmp_sh = sr(i).Duplicate(1, 1)
        With sr(i).BreakApartEx
            If .Count > 1 Then
                'tmp_sh.Delete
                'MsgBox "count > 1"
                'Set sr(i) = Nothing
            End If
            tmp_sr.AddRange .All
        End With
        
    End If
    sr.AddRange tmp_sr
 Next
 End If
 
 
 'MsgBox CStr(sr.Count)
 'sr.AddRange bsr
 Dim progress As Long
 progress = 0
 Status.BeginProgress "Dvojice"
 c = 0
 cps = 0
 For i = 1 To sr.Count
    Set sh1 = sr.Shapes(i)
    'If sh1.DisplayCurve.SubPaths.Count <> 1 Then
        'sh1.CreateSelection
        'MsgBox (CStr(sh1.Curve.SubPaths.Count))
        'sh1.CreateSelection
    'End If


    For j = i + 1 To sr.Count  'not i + 1 because i need to check subpaths within 1 shape
        Status.progress = CLng(100 * progress / (sr.Count * (sr.Count) / 2))
        Set sh2 = sr.Shapes(j)
        progress = progress + 1
        'sh1.CreateSelection
        'sh2.AddToSelection
        If Not sh1.BoundingBox.Intersect(sh2.BoundingBox).IsEmpty Then
            For k = 1 To sh1.DisplayCurve.SubPaths.Count
                For l = 1 To sh2.DisplayCurve.SubPaths.Count
                    If Not sh1.DisplayCurve.SubPaths(k).BoundingBox.Intersect(sh2.DisplayCurve.SubPaths(l).BoundingBox).IsEmpty Then
                        Set cs = sh1.DisplayCurve.SubPaths(k).GetIntersections(sh2.DisplayCurve.SubPaths(l))
                        cps = cps + cs.Count
                        'For Each cp In cs
                        '    ActiveLayer.CreateEllipse2 cp.PositionX, cp.PositionY, 0.1
                        'Next cp
                    End If
                    c = c + 1
                Next l
            Next k
        End If
    Next j
    
    Set cs = ShapeIntersections(sh1)
    'cps = cps + cs.Count
    'sh1.Curve.SubPaths.First
 Next i
 MsgBox CStr(c) & " CPS: " & CStr(cps)
1000:
 Status.EndProgress
 sr.Delete
 'asr.Delete
 ActiveDocument.EndCommandGroup
 asr.CreateSelection
 myOptimize True, False
 MsgBox sw.EndSeconds
End Sub

Sub BreakTextLine()
Dim s As Shape
Dim n&, l&
    Set s = ActiveShape
    If s Is Nothing Then Exit Sub
    If s.Type <> cdrTextShape Then Exit Sub
    l = s.Text.Story.Lines.Count
    s.BreakApart
    For n = 1 To l - 1
        Set s = s.Previous
        s.AddToSelection
    Next n
End Sub

Sub Test()
 Dim s As Shape
 Dim n As Long
 n = 0
 For Each s In ActivePage.Shapes
  If s.Locked Then n = n + 1
 Next s
 MsgBox "There are " & n & " shapes locked on the current page"
End Sub

Private Function ShapeIntersections(sh As Shape) As CrossPoints
    Dim i As Long, j As Long
    Dim ret As Collection, cps As CrossPoints, cp As CrossPoint
    Dim sr As ShapeRange
    Set sr = CreateShapeRange
    Set ret = New Collection
    Set sr = sh.BreakApartEx
    For i = 1 To sr.Count
        For j = i + 1 To sr.Count
            If Not sr(i).DisplayCurve.SubPaths.First.BoundingBox.Intersect(sr().DisplayCurve.SubPaths.First.BoundingBox).IsEmpty Then
                
                Set cps = sh.DisplayCurve.SubPaths(i).GetIntersections(sh.DisplayCurve.SubPaths(j))
                For Each cp In cps
                    ret.add cp
                Next
            End If
        Next j
    Next i
    Set ShapeIntersections = CVar(ret)
End Function
