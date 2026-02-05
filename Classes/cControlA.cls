VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cControlA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Txt As VB.TextBox
Attribute Txt.VB_VarHelpID = -1

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And (Shift And vbCtrlMask) <> 0 Then
        Txt.SelStart = 0
        Txt.SelLength = Len(Txt.Text)
        KeyCode = 0
    End If
End Sub

