VERSION 5.00
Begin VB.MDIForm MDIFrmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Principal"
   ClientHeight    =   5955
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10605
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
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
Private Sub mnuTeste_Click()
    frmTestes.Show
End Sub
