VERSION 5.00
Begin VB.MDIForm MDIFrmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Principal"
   ClientHeight    =   8610
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   18090
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuTeste 
      Caption         =   "Teste"
   End
   Begin VB.Menu mnuOperadores 
      Caption         =   "Operadores"
   End
   Begin VB.Menu mnuClientes 
      Caption         =   "Clientes"
   End
   Begin VB.Menu mnuProdutos 
      Caption         =   "Produtos"
   End
   Begin VB.Menu mnuPedidos 
      Caption         =   "Pedidos"
   End
   Begin VB.Menu mnuRelPedidos 
      Caption         =   "Relatório Pedidos"
   End
End
Attribute VB_Name = "MDIFrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuClientes_Click()
    frmCliente.Show
End Sub

Private Sub mnuOperadores_Click()
    frmOperador.Show
End Sub

Private Sub mnuPedidos_Click()
    frmPedido.Show
End Sub

Private Sub mnuProdutos_Click()
    frmProduto.Show
End Sub

Private Sub mnuRelPedidos_Click()
    frmRelPedidos.Show
End Sub

Private Sub mnuTeste_Click()
    frmTeste.Show
End Sub
