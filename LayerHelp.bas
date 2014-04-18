Attribute VB_Name = "LayerHelp"
Option Explicit

Function GetLayerOrCreate(layerName As String) As Layer
    Dim l As Layer
    On Error Resume Next
    Set l = ActivePage.Layers(layerName)
    If l Is Nothing Then
        Set GetLayerOrCreate = ActiveDocument.ActivePage.CreateLayer(layerName)
    Else
        Set GetLayerOrCreate = l
    End If
End Function
