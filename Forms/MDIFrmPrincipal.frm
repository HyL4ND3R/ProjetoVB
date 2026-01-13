VERSION 5.00
Begin VB.MDIForm MDIFrmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Principal"
   ClientHeight    =   8610
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   18090
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuTeste 
      Caption         =   "Teste"
   End
   Begin VB.Menu mnuOperadores 
      Caption         =   "Operadores"
   End
End
Attribute VB_Name = "MDIFrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuOperadores_Click()
    frmOperador.Show
End Sub

Private Sub mnuTeste_Click()
    frmTestes.Show
End Sub
