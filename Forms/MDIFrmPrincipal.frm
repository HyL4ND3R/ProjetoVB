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
      Visible         =   0   'False
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
    If frmCliente.Visible Then
        frmCliente.SetFocus
    Else
        frmCliente.Show
    End If
End Sub

Private Sub mnuOperadores_Click()
    If frmOperador.Visible Then
        frmOperador.SetFocus
    Else
        frmOperador.Show
    End If
End Sub

Private Sub mnuPedidos_Click()
    If frmPedido.Visible Then
        frmPedido.SetFocus
    Else
        frmPedido.Show
    End If
End Sub

Private Sub mnuProdutos_Click()
    If frmProduto.Visible Then
        frmProduto.SetFocus
    Else
        frmProduto.Show
    End If
End Sub

Private Sub mnuRelPedidos_Click()
    If frmRelPedidos.Visible Then
        frmRelPedidos.SetFocus
    Else
        frmRelPedidos.Show
    End If
End Sub

Private Sub mnuTeste_Click()
    frmTeste.Show
End Sub
