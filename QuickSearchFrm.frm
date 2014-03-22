VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QuickSearchFrm 
   Caption         =   "Quick Search"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14010
   OleObjectBlob   =   "QuickSearchFrm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QuickSearchFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    myOptimize chkOptim.Value, True
    Query TextBox1.Text, select_add.Value, select_sub.Value
    myOptimize chkOptim.Value, False
End Sub

Private Sub TextBox1_Change()
    'If (ActiveShape Is Nothing) Then
    '    Exit Sub
    'End If
    On Error GoTo ErrHandler
    Dim s As String
    Dim h$
    'Err.Clear
    s = ActiveShape.Evaluate(TextBox1.Text)
    h = Trim(TextBox1.Text)
    Dim aH() As String
    If (h <> "") Then
        aH = Split(h, " ")
        h = Trim(aH(UBound(aH)))
    End If
    h = h + "help"
    On Error Resume Next
    h = ActiveShape.Evaluate(h)
    's = CStr((ActivePage.Shapes.FindShapes(Query:="" + TextBox1.Text).Count))
    'On Error GoTo ErrHandler:
    Label1.Caption = s
    Label2.Caption = h
    lblCount.Caption = CStr((ActivePage.Shapes.FindShapes(Query:="" + TextBox1.Text).Count))
Exit Sub
ErrHandler:
    Label3.Caption = CStr(Err.Number) + " " + Err.Description + " " + Err.Source
    Err.Clear
    Resume Next
End Sub

Private Sub a()
    Math.Abs (2)
    'Dim list(20) As String
    'list(0) = "Name"
    'list(1) = "Fill"
    'list(2) = "Width"
    'list(3) = "Height"
    'For Index = 0 To list.GetUpperBound(0)
    '    ComboBox1.AddItem list(Index)
    'Next
    
End Sub


Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'MsgBox CStr(KeyAscii)
    If (KeyCode.Value = vbKeyUp) Then
        MsgBox "Up"
    End If
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'MsgBox CStr(KeyAscii)
    If (KeyAscii.Value = vbKeyUp) Then
        MsgBox "Up"
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim s As String
    Dim obj As Object
    Label1.Caption = "a"
End Sub

Sub eval()
    Dim Doc As Document
    Set Doc = ActiveDocument
    MsgBox Application.Evaluate("vba.test.Module1.doc.FileName")
    'MsgBox Application.Evaluate("vba.Application.ActiveDocument.ActiveLayer.CreateEllipse(0, 0, 10, 10, 90, 90, false)")
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  'MsgBox CStr(KeyAscii.Value)
End Sub
