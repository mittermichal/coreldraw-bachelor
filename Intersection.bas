Attribute VB_Name = "Intersection"
Option Explicit

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
 myOptimize True, True

 

 
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
 Dim cps_array() As CrossPoint
 ReDim cps_array(0)
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
                        If UBound(cps_array) > 0 Then
                            ReDim Preserve cps_array(1 To cps + cs.Count)
                        ElseIf cs.Count > 0 Then
                            ReDim cps_array(1 To cs.Count)
                        End If
                        
                        Dim m As Long
                        For m = 1 To cs.Count
                            Set cps_array(cps + m) = cs(m)
                        Next
                        
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
    'ShapeIntersections sh1
    Dim tmp_array() As CrossPoint
    tmp_array = ShapeIntersections(sh1)
    If UBound(cps_array) > 0 Then
        ReDim Preserve cps_array(1 To cps + UBound(tmp_array))
    ElseIf UBound(tmp_array) > 0 Then
        ReDim cps_array(1 To UBound(tmp_array))
    End If
    
    For m = 1 To UBound(tmp_array)
        Set cps_array(cps + m) = tmp_array(m)
    Next
    
    cps = cps + UBound(tmp_array)
    
    'sh1.Curve.SubPaths.First
 Next i
 
 Dim lr As Layer
 If UBound(cps_array) > 0 Then
 Set lr = GetLayerOrCreate("InterSections")
 For i = 1 To UBound(cps_array)
    'a = s.GetTangentAt(offset, cdrRelativeSegmentOffset)
    Set sh = lr.CreateEllipse2(cps_array(i).PositionX, cps_array(i).PositionY, 0.5)
    sh.Outline.Color = CreateRGBColor(255, 0, 0)
 Next i
 End If
 
1000:
 Status.EndProgress
 sr.Delete 'preskocit chyby
 'asr.Delete
 ActiveDocument.EndCommandGroup
 'asr.CreateSelection
 myOptimize True, False
 MsgBox sw.EndSeconds & CStr(c) & " CPS: " & CStr(cps)
End Sub


Private Function ShapeIntersections(sh As Shape) As CrossPoint()
    Dim i As Long, j As Long
    Dim c As Long
    Dim ret As Collection, cps As CrossPoints, cp As CrossPoint
    Dim sr As ShapeRange
    Dim rets() As CrossPoint
    ReDim rets(0)
    Set sr = CreateShapeRange
    Set ret = New Collection
    Set sr = sh.Duplicate.BreakApartEx
    
    c = 0
    For i = 1 To sr.Count
        For j = i + 1 To sr.Count
            'If Not sr(i).DisplayCurve.SubPaths.First.BoundingBox.Intersect(sr(j).DisplayCurve.SubPaths.First.BoundingBox).IsEmpty Then
                
                Set cps = sr(i).DisplayCurve.SubPaths.First.GetIntersections(sr(j).DisplayCurve.SubPaths.First)

                If UBound(rets) > 0 Then
                    ReDim Preserve rets(1 To c + cps.Count)
                ElseIf cps.Count > 0 Then
                    ReDim rets(1 To cps.Count)
                End If
                Dim k As Long
                For k = 1 To cps.Count
                    Set rets(c + k) = cps(k)
                Next
                c = c + cps.Count
            'End If
        Next j
    Next i
    sr.Delete
    ShapeIntersections = rets
End Function
