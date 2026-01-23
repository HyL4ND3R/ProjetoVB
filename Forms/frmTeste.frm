VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTeste 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox mskTeste 
      Height          =   405
      Left            =   690
      TabIndex        =   0
      Top             =   930
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   714
      _Version        =   393216
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "frmTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    mskTeste.Mask = "99999,99"
    mskTeste.Text = "00000,00"
End Sub
