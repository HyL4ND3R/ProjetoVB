VERSION 5.00
Begin VB.Form frmTestes 
   Caption         =   "Testes"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   10095
   Begin VB.CommandButton cmdString 
      Caption         =   "Ação"
      Height          =   375
      Left            =   1530
      TabIndex        =   3
      Top             =   900
      Width           =   1305
   End
   Begin VB.TextBox txtString3 
      Height          =   375
      Left            =   210
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1440
      Width           =   4005
   End
   Begin VB.TextBox txtString2 
      Height          =   345
      Left            =   2220
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   330
      Width           =   1995
   End
   Begin VB.TextBox txtString1 
      Height          =   345
      Left            =   210
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   330
      Width           =   1995
   End
End
Attribute VB_Name = "frmTestes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdString_Click()

    'Declaração
    Dim string1 As String
    Dim string2 As String
    'Pegar os Dados
    string1 = txtString1.Text
    string2 = txtString2.Text
    'Ação
    txtString3 = string1 & string2
    
End Sub

