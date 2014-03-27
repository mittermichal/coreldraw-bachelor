VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GuidesFrm 
   Caption         =   "Nastavenia"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2640
   OleObjectBlob   =   "GuidesFrm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GuidesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtOffset_Change()
    'MsgBox CStr(CDbl(txtOffset.Value))
    'If IsNumeric(txtOffset.Value) Then
        SaveSetting "CorelDrawBachelor", "Guides", "tangentOffset", txtOffset.Value
    'End If
    GuidesFrm.Label1.Caption = txtOffset.Value
End Sub

Private Sub UserForm_Initialize()
    txtOffset.Value = GetSetting("CorelDrawBachelor", "Guides", "tangentOffset", "0.5")
End Sub
