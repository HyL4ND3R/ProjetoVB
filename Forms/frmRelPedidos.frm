VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRelPedidos 
   Caption         =   "Relatório De Pedidos"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   6885
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Visualizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2565
      TabIndex        =   12
      Top             =   2970
      Width           =   1275
   End
   Begin VB.CommandButton cmdListaProduto 
      DisabledPicture =   "frmRelPedidos.frx":0000
      DownPicture     =   "frmRelPedidos.frx":05E2
      Height          =   375
      Left            =   2730
      Picture         =   "frmRelPedidos.frx":0BC4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2070
      Width           =   525
   End
   Begin VB.TextBox txtNomeProduto 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3270
      TabIndex        =   10
      Top             =   2070
      Width           =   3150
   End
   Begin VB.TextBox txtCodProduto 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1890
      TabIndex        =   9
      Top             =   2070
      Width           =   855
   End
   Begin VB.TextBox txtNomeCliente 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3240
      TabIndex        =   8
      Top             =   1575
      Width           =   3150
   End
   Begin VB.TextBox txtCodCliente 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1890
      TabIndex        =   7
      Top             =   1575
      Width           =   825
   End
   Begin VB.CommandButton cmdListaCliente 
      DisabledPicture =   "frmRelPedidos.frx":11A6
      DownPicture     =   "frmRelPedidos.frx":1788
      Height          =   375
      Left            =   2715
      Picture         =   "frmRelPedidos.frx":1D6A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1575
      Width           =   525
   End
   Begin MSComCtl2.DTPicker dtpDataInicial 
      Height          =   375
      Left            =   1890
      TabIndex        =   0
      Top             =   450
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   62455809
      CurrentDate     =   46051
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1890
      TabIndex        =   1
      Top             =   945
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   62455809
      CurrentDate     =   46051
   End
   Begin VB.Label lblCliente 
      Alignment       =   1  'Right Justify
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   5
      Top             =   1575
      Width           =   1320
   End
   Begin VB.Label lblProduto 
      Alignment       =   1  'Right Justify
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   4
      Top             =   2070
      Width           =   1320
   End
   Begin VB.Label lblDataFinal 
      Alignment       =   1  'Right Justify
      Caption         =   "Data Final:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   3
      Top             =   945
      Width           =   1320
   End
   Begin VB.Label lblDataInicial 
      Alignment       =   1  'Right Justify
      Caption         =   "Data Inicial:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   2
      Top             =   450
      Width           =   1320
   End
End
Attribute VB_Name = "frmRelPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    
End Sub

Private Sub cmdVisualizar_Click()
    
    If (Not ValidaCampos) Then Exit Sub
    
    Dim rpt As New arRelatorioPedidos
    Dim Sql As String
    
    'Define a Conexão com o Banco
    rpt.dcRelPedidos.ConnectionString = Conn
    
    Sql = "select Pedido.Codigo As Pedido, Cliente.Nome As Cliente, FORMAT(pedido.Data,'dd/MM/yyyy') As DataPedido, " & _
            "isNull(Pedido.QtdeTotal,0) As QtdeTotal, IsNull(Pedido.ValorTotal,0) As ValorTotal, " & _
            "PedidoItem.ProdutoCodigo As ProdutoCod,  PedidoItem.Descricao As Produto, " & _
            "PedidoItem.Quantidade As ProdutoQtde, PedidoItem.ValorUn As ProdutoValorUn, " & _
            "PedidoItem.ValorTotal As ProdutoValorTotal " & _
            "From pedido " & _
            "Inner join Cliente on Pedido.ClienteCodigo = Cliente.Codigo " & _
            "Left join PedidoItem  on PedidoItem.ControlePedido = Pedido.Controle " & _
            "Order by Pedido.Codigo, PedidoItem.Item"
    
    'Define a string que vai ser executada no banco
    rpt.dcRelPedidos.Source = Sql
    
    rpt.Run
    
    rpt.Show vbModal

End Sub

Private Function ValidaCampos() As Boolean
    
End Function
