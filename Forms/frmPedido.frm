VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPedido 
   Caption         =   "Pedido"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11130
   ScaleWidth      =   20895
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancelarItem 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6000
      TabIndex        =   26
      Top             =   2280
      Width           =   1155
   End
   Begin VB.CommandButton cmdExcluirItem 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4860
      TabIndex        =   25
      Top             =   2280
      Width           =   1155
   End
   Begin VB.CommandButton cmdAlterarItem 
      Caption         =   "Alterar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   24
      Top             =   2280
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalvarItem 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2580
      TabIndex        =   23
      Top             =   2280
      Width           =   1155
   End
   Begin VB.CommandButton cmdNovoItem 
      Caption         =   "Novo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1440
      TabIndex        =   22
      Top             =   2280
      Width           =   1155
   End
   Begin VB.TextBox txtTotalItem 
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
      Left            =   1410
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtValorUn 
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
      Left            =   1410
      TabIndex        =   7
      Top             =   3630
      Width           =   1215
   End
   Begin VB.TextBox txtQtde 
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
      Left            =   1410
      TabIndex        =   6
      Top             =   3180
      Width           =   1215
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
      Left            =   1410
      TabIndex        =   4
      Top             =   2730
      Width           =   1215
   End
   Begin VB.TextBox txtDescricao 
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
      Left            =   3150
      TabIndex        =   5
      Top             =   2730
      Width           =   4005
   End
   Begin VB.CommandButton cmdListaProduto 
      DisabledPicture =   "frmPedido.frx":0000
      DownPicture     =   "frmPedido.frx":05E2
      Height          =   375
      Left            =   2610
      Picture         =   "frmPedido.frx":0BC4
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2730
      Width           =   525
   End
   Begin VB.TextBox txtValorTotal 
      Enabled         =   0   'False
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
      Left            =   3990
      TabIndex        =   15
      Top             =   6060
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grdItensPedido 
      Height          =   2805
      Left            =   2700
      TabIndex        =   14
      Top             =   3180
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   4948
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdListaCliente 
      DisabledPicture =   "frmPedido.frx":11A6
      DownPicture     =   "frmPedido.frx":1788
      Height          =   375
      Left            =   2610
      Picture         =   "frmPedido.frx":1D6A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1380
      Width           =   525
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1380
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker dtpDataPedido 
      Height          =   375
      Left            =   5550
      TabIndex        =   3
      Top             =   900
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   150011905
      CurrentDate     =   36526
      MaxDate         =   73415
      MinDate         =   36526
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   900
      Width           =   1005
   End
   Begin VB.CommandButton cmdListaPedido 
      DisabledPicture =   "frmPedido.frx":234C
      DownPicture     =   "frmPedido.frx":292E
      Height          =   375
      Left            =   2610
      Picture         =   "frmPedido.frx":2F10
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   900
      Width           =   525
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
      Left            =   3180
      TabIndex        =   0
      Top             =   1380
      Width           =   4005
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   20895
      _ExtentX        =   36856
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alterar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excluir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "desfazer"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "primeiro"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "proximo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ultimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "frmPedido.frx":34F2
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   20280
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":41CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":4EA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":5B80
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":685A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":7534
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":820E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":87EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":94C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":A19E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
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
      Left            =   450
      TabIndex        =   21
      Top             =   4080
      Width           =   915
   End
   Begin VB.Label lblValorUn 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Un:"
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
      Left            =   450
      TabIndex        =   20
      Top             =   3630
      Width           =   915
   End
   Begin VB.Label lblQtde 
      Alignment       =   1  'Right Justify
      Caption         =   "Qtde:"
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
      Left            =   450
      TabIndex        =   19
      Top             =   3180
      Width           =   915
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
      Left            =   450
      TabIndex        =   17
      Top             =   2730
      Width           =   915
   End
   Begin VB.Label lblValorTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2700
      TabIndex        =   16
      Top             =   6060
      Width           =   1245
   End
   Begin VB.Label lblCodigo 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   450
      TabIndex        =   11
      Top             =   900
      Width           =   1065
   End
   Begin VB.Label lblCliente 
      Alignment       =   1  'Right Justify
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   450
      TabIndex        =   10
      Top             =   1380
      Width           =   1065
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ModoAtualPedido As eModoFormulario
Private ModoAtualItens As eModoFormulario
Dim pedido As cPedido
Dim pedidoItem As cPedidoItem
Dim ControlePedido As Long
Dim ControlePedidoItem As Long

Private Sub Form_Load()
    
    Set pedido = New cPedido
    
    AjustarColunasGridItens
    CarregarPedidos

    If Not rsPedido.EOF Then 'Se não esta no fim da lista
        rsPedido.MoveLast 'Move para o final
        PreencherCampos
    Else
        dtpDataPedido.Format = dtpCustom 'Definir uma mascara para poder zerar o campo
        dtpDataPedido.CustomFormat = " " 'Zerando a data para não ficar preenchida
    End If

    modoConsultaPedido
    
End Sub

Private Sub cmdCancelarItem_Click()
    cancelarItem
End Sub

Private Sub cmdNovoItem_Click()
    modoInclusaoItem
End Sub

Private Sub cmdAlterarItem_Click()
    modoAlteracaoItem
End Sub

Private Sub cmdSalvarItem_Click()
    Set pedidoItem = New cPedidoItem
    Dim codigoAtual As Long

    If ModoAtualPedido = mfAlteracao Then
        pedidoItem.Controle = VerificaNull(ControlePedidoItem, 0)
        pedidoItem.ControlePedido = VerificaNull(ControlePedido, 0)
        pedidoItem.Item = 0
        pedidoItem.ProdutoCodigo = CLng(txtCodProduto.Text)
        pedidoItem.Descricao = txtDescricao.Text
        pedidoItem.Qtde = CDbl(txtQtde.Text)
        pedidoItem.ValorUn = CDbl(txtValorUn.Text)
        pedidoItem.ValorTotal = CDbl(txtQtde.Text) * CDbl(txtValorUn.Text)
        If (Not AlterarItemPedido(pedido)) Then
            MsgBox "Erro ao Alterar o Registro!"
            Exit Sub
        End If
    Else
        pedidoItem.ControlePedido = VerificaNull(ControlePedido, 0)
        pedidoItem.Item = 0
        pedidoItem.ProdutoCodigo = CLng(txtCodProduto.Text)
        pedidoItem.Descricao = txtDescricao.Text
        pedidoItem.Qtde = CDbl(txtQtde.Text)
        pedidoItem.ValorUn = CDbl(txtValorUn.Text)
        pedidoItem.ValorTotal = CDbl(txtQtde.Text) * CDbl(txtValorUn.Text)
        If (Not InserirItemPedido(pedidoItem)) Then
            MsgBox "Erro ao Inserir o Registro!"
            Exit Sub
        End If
    End If
    
    PreencherItensPedido
    modoInclusaoItem
    
End Sub

Private Sub modoInclusaoPedido()
'--------------TOOLBAR------------------------
    Toolbar1.Buttons("novo").Enabled = False 'Habilitar/Desabilitar botão da toolbar
    Toolbar1.Buttons("salvar").Enabled = True
    Toolbar1.Buttons("alterar").Enabled = False
    Toolbar1.Buttons("excluir").Enabled = False
    Toolbar1.Buttons("desfazer").Enabled = True
    Toolbar1.Buttons("primeiro").Enabled = False
    Toolbar1.Buttons("anterior").Enabled = False
    Toolbar1.Buttons("proximo").Enabled = False
    Toolbar1.Buttons("ultimo").Enabled = False
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = False 'Habilitar/Desabilitar txt
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaPedido.Enabled = False 'Habilitar/Desabilitar commandButton
    txtCodCliente.Enabled = True
    txtCodCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    cmdListaCliente.Enabled = True 'Habilitar/Desabilitar commandButton
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    dtpDataPedido.Enabled = True
'--------------BOTÕES ITENS------------------------
    cmdNovoItem.Enabled = False
    cmdSalvarItem.Enabled = False
    cmdAlterarItem.Enabled = False
    cmdExcluirItem.Enabled = False
    cmdCancelarItem.Enabled = False
'--------------CAMPOS ITENS------------------------
    txtCodProduto.Enabled = False
    txtCodProduto.BackColor = &H8000000F
    cmdListaProduto.Enabled = False
    txtDescricao.Enabled = False
    txtDescricao.BackColor = &H8000000F
    txtQtde.Enabled = False
    txtQtde.BackColor = &H8000000F
    txtValorUn.Enabled = False
    txtValorUn.BackColor = &H8000000F
    txtTotalItem.Enabled = False
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualPedido = mfInclusao
End Sub

Private Sub modoAlteracaoPedido()
'--------------TOOLBAR------------------------
    Toolbar1.Buttons("novo").Enabled = False 'Habilitar/Desabilitar botão da toolbar
    Toolbar1.Buttons("salvar").Enabled = True
    Toolbar1.Buttons("alterar").Enabled = False
    Toolbar1.Buttons("excluir").Enabled = False
    Toolbar1.Buttons("desfazer").Enabled = True
    Toolbar1.Buttons("primeiro").Enabled = False
    Toolbar1.Buttons("anterior").Enabled = False
    Toolbar1.Buttons("proximo").Enabled = False
    Toolbar1.Buttons("ultimo").Enabled = False
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = False 'Habilitar/Desabilitar txt
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaPedido.Enabled = False 'Habilitar/Desabilitar commandButton
    txtCodCliente.Enabled = True
    txtCodCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    cmdListaCliente.Enabled = True 'Habilitar/Desabilitar commandButton
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    dtpDataPedido.Enabled = True
'--------------BOTÕES ITENS------------------------
    cmdNovoItem.Enabled = False
    cmdSalvarItem.Enabled = False
    cmdAlterarItem.Enabled = False
    cmdExcluirItem.Enabled = False
    cmdCancelarItem.Enabled = False
'--------------CAMPOS ITENS------------------------
    txtCodProduto.Enabled = False
    txtCodProduto.BackColor = &H8000000F
    cmdListaProduto.Enabled = False
    txtDescricao.Enabled = False
    txtDescricao.BackColor = &H8000000F
    txtQtde.Enabled = False
    txtQtde.BackColor = &H8000000F
    txtValorUn.Enabled = False
    txtValorUn.BackColor = &H8000000F
    txtTotalItem.Enabled = False
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualPedido = mfAlteracao
End Sub

Private Sub modoConsultaPedido()
'--------------TOOLBAR------------------------
    Toolbar1.Buttons("novo").Enabled = True
    Toolbar1.Buttons("salvar").Enabled = False
    Toolbar1.Buttons("excluir").Enabled = True
    Toolbar1.Buttons("alterar").Enabled = True
    Toolbar1.Buttons("desfazer").Enabled = False
    Toolbar1.Buttons("primeiro").Enabled = True
    Toolbar1.Buttons("anterior").Enabled = True
    Toolbar1.Buttons("proximo").Enabled = True
    Toolbar1.Buttons("ultimo").Enabled = True
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = True
    txtCodigo.BackColor = vbWindowBackground
    cmdListaPedido.Enabled = True
    txtCodCliente.Enabled = False
    txtCodCliente.BackColor = &H8000000F
    cmdListaCliente.Enabled = False
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = &H8000000F
    dtpDataPedido.Enabled = False
'--------------BOTÕES ITENS------------------------
    cmdNovoItem.Enabled = True
    cmdSalvarItem.Enabled = False
    cmdAlterarItem.Enabled = True
    cmdExcluirItem.Enabled = True
    cmdCancelarItem.Enabled = False
'--------------CAMPOS ITENS------------------------
    txtCodProduto.Enabled = False
    txtCodProduto.BackColor = &H8000000F
    cmdListaProduto.Enabled = False
    txtDescricao.Enabled = False
    txtDescricao.BackColor = &H8000000F
    txtQtde.Enabled = False
    txtQtde.BackColor = &H8000000F
    txtValorUn.Enabled = False
    txtValorUn.BackColor = &H8000000F
    txtTotalItem.Enabled = False
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualPedido = mfConsulta
End Sub

Private Sub modoInclusaoItem()
'--------------BOTÕES ITENS------------------------
    cmdNovoItem.Enabled = False
    cmdSalvarItem.Enabled = True
    cmdAlterarItem.Enabled = False
    cmdExcluirItem.Enabled = False
    cmdCancelarItem.Enabled = True
'--------------CAMPOS ITENS------------------------
    txtCodProduto.Enabled = True
    txtCodProduto.Text = ""
    txtCodProduto.BackColor = vbWindowBackground
    cmdListaProduto.Enabled = True
    txtDescricao.Enabled = True
    txtDescricao.Text = ""
    txtDescricao.BackColor = vbWindowBackground
    txtQtde.Enabled = True
    txtQtde.Text = ""
    txtQtde.BackColor = vbWindowBackground
    txtValorUn.Enabled = True
    txtValorUn.Text = ""
    txtValorUn.BackColor = vbWindowBackground
    txtTotalItem.Enabled = False
    txtTotalItem.Text = ""
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualItens = mfInclusao
    
End Sub

Private Sub modoAlteracaoItem()
'--------------BOTÕES ITENS------------------------
    cmdNovoItem.Enabled = False
    cmdSalvarItem.Enabled = True
    cmdAlterarItem.Enabled = False
    cmdExcluirItem.Enabled = False
    cmdCancelarItem.Enabled = True
'--------------CAMPOS ITENS------------------------
    txtCodProduto.Enabled = True
    txtCodProduto.Text = ""
    txtCodProduto.BackColor = vbWindowBackground
    cmdListaProduto.Enabled = True
    txtDescricao.Enabled = True
    txtDescricao.Text = ""
    txtDescricao.BackColor = vbWindowBackground
    txtQtde.Enabled = True
    txtQtde.Text = ""
    txtQtde.BackColor = vbWindowBackground
    txtValorUn.Enabled = True
    txtValorUn.Text = ""
    txtValorUn.BackColor = vbWindowBackground
    txtTotalItem.Enabled = False
    txtTotalItem.Text = ""
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualItens = mfAlteracao
    
End Sub

Private Sub cancelarItem()
'--------------BOTÕES ITENS------------------------
    cmdNovoItem.Enabled = True
    cmdSalvarItem.Enabled = False
    cmdAlterarItem.Enabled = True
    cmdExcluirItem.Enabled = True
    cmdCancelarItem.Enabled = False
'--------------CAMPOS ITENS------------------------
    txtCodProduto.Enabled = False
    txtCodProduto.BackColor = &H8000000F
    cmdListaProduto.Enabled = False
    txtDescricao.Enabled = False
    txtDescricao.BackColor = &H8000000F
    txtQtde.Enabled = False
    txtQtde.BackColor = &H8000000F
    txtValorUn.Enabled = False
    txtValorUn.BackColor = &H8000000F
    txtTotalItem.Enabled = False
    txtTotalItem.Text = ""
    txtTotalItem.BackColor = &H8000000F
    
    PreencherCamposItem
End Sub



Private Sub PreencherCampos()

    If rsPedido.EOF Or rsPedido.BOF Then 'Se a lista não tem registros pula fora da Sub
        dtpDataPedido.Format = dtpCustom 'Definir uma mascara para poder zerar o campo
        dtpDataPedido.CustomFormat = " " 'Zerando a data para não ficar preenchida
        Exit Sub
    End If

    ControlePedido = rsPedido!Controle
    txtCodigo.Text = rsPedido!Codigo 'Atribuição de valor do RecordSet para o TextBox
    txtCodCliente.Text = rsPedido!ClienteCodigo
    txtNomeCliente.Text = rsPedido!ClienteNome
    dtpDataPedido.Format = dtpCustom
    dtpDataPedido.CustomFormat = "dd/MM/yyyy" 'Redefindindo a mascara caso não tinha registros na tela antes
    dtpDataPedido.Value = rsPedido!DataPedido
    txtValorTotal.Text = IIf(IsNull(rsPedido!ValorTotal), 0, rsPedido!ValorTotal)
    
    PreencherItensPedido
    
End Sub


Private Sub PreencherClientePedido()
    If rsCliente.EOF Or rsCliente.BOF Then Exit Sub
    
    txtCodCliente.Text = rsCliente!Codigo
    txtNomeCliente.Text = rsCliente!Nome
    
End Sub

Private Sub PreencherItensPedido()
    
    ' Limpa o grid (mantém só o cabeçalho)
    grdItensPedido.Rows = 1
    
    CarregarItensPedido (ControlePedido)
    If rsPedidoItem.EOF Or rsPedidoItem.BOF Then Exit Sub
    
    Dim linha As Long
    
    If Not IsNumeric(txtCodigo.Text) Then
        MsgBox "Código inválido.", vbExclamation
        txtCodigo.SetFocus
        Exit Sub
    End If
    
    linha = 1

    Do While Not rsPedidoItem.EOF
        grdItensPedido.Rows = grdItensPedido.Rows + 1

        grdItensPedido.TextMatrix(linha, 0) = rsPedidoItem!Controle
        grdItensPedido.TextMatrix(linha, 1) = rsPedidoItem!ProdutoCodigo
        grdItensPedido.TextMatrix(linha, 2) = rsPedidoItem!Descricao
        grdItensPedido.TextMatrix(linha, 3) = rsPedidoItem!Quantidade
        grdItensPedido.TextMatrix(linha, 4) = Format(rsPedidoItem!ValorUn, "0.00")
        grdItensPedido.TextMatrix(linha, 5) = Format(rsPedidoItem!ValorTotal, "0.00")

        linha = linha + 1
        rsPedidoItem.MoveNext
    Loop
    
    If grdItensPedido.Rows > 1 Then
        grdItensPedido.Row = 1
        PreencherCamposItem
    End If

End Sub

Private Sub AjustarColunasGridItens()
    With grdItensPedido
        .Rows = 1
        .Cols = 6
        
        .TextMatrix(0, 0) = "Controle"
        .TextMatrix(0, 1) = "Código"
        .TextMatrix(0, 2) = "Produto"
        .TextMatrix(0, 3) = "Qtde"
        .TextMatrix(0, 4) = "Vlr Unit"
        .TextMatrix(0, 5) = "Total"

        .ColWidth(0) = 0
        .ColWidth(1) = 1000
        .ColWidth(2) = 3800
        .ColWidth(3) = 1000
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
    End With
    
    grdItensPedido.SelectionMode = flexSelectionByRow
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
'-------------NOVO-------------------------------------------------------------------
        Case "novo"
            
            BuscarProximoCodPedido
            txtCodigo.Text = rsProximoCodigo!Codigo
            txtCodCliente.Text = ""
            txtNomeCliente.Text = ""
            dtpDataPedido.Format = dtpCustom
            dtpDataPedido.CustomFormat = "dd/MM/yyyy"
            dtpDataPedido.Value = Date
            txtValorTotal.Text = ""
            
            modoInclusaoPedido

'-------------SALVAR-----------------------------------------------------------------
        Case "salvar"
            Dim codigoAtual As Long

            If ModoAtualPedido = mfAlteracao Then
                pedido.Controle = VerificaNull(ControlePedido, 0)
                pedido.Codigo = CLng(txtCodigo.Text) 'Conversão de Texto para Long
                pedido.ClienteCodigo = CLng(txtCodCliente.Text)
                pedido.DataPedido = dtpDataPedido.Value
                If (Not AlterarPedido(pedido)) Then
                    MsgBox "Erro ao Alterar o Registro!"
                    Exit Sub
                End If
            Else
                pedido.Controle = VerificaNull(ControlePedido, 0)
                pedido.Codigo = CLng(txtCodigo.Text) 'Conversão de Texto para Long
                pedido.ClienteCodigo = CLng(txtCodCliente.Text)
                pedido.DataPedido = dtpDataPedido.Value
                If (Not InserirPedido(pedido)) Then
                    MsgBox "Erro ao Inserir o Registro!"
                    Exit Sub
                End If
            End If
            
            CarregarPedidos
            
            If ModoAtualPedido = mfAlteracao Then
                rsPedido.Find "Codigo = " & codigoAtual
            Else
                If Not rsPedido.EOF Then rsPedido.MoveLast
            End If
            
            PreencherCampos
            modoConsultaPedido

'-------------ALTERACAO--------------------------------------------------------------
        Case "alterar"
        
            If rsPedido.EOF Or rsPedido.BOF Then Exit Sub
            
            PreencherCampos
            modoAlteracaoPedido

'-------------EXCLUIR----------------------------------------------------------------
        Case "excluir"
            
            If rsPedido.EOF Or rsPedido.BOF Then Exit Sub
            
            'Mensagem de confirmação, se clicar no Não, cai fora da sub
            If MsgBox("Deseja realmente excluir este Registro?", _
                      vbQuestion + vbYesNo, _
                      "Confirmação") = vbNo Then Exit Sub

                  
            Conn.Execute "DELETE FROM Pedido WHERE Controle = " & ControlePedido
        
            CarregarPedidos
        
            If Not rsPedido.EOF Then
                rsPedido.Find "Codigo > " & codigoExcluir
                If rsPedido.EOF Then rsPedido.MoveLast
            End If
        
            PreencherCampos
            modoConsultaPedido

'-------------DESFAZER
        Case "desfazer"
            txtCodigo.Text = ""
            modoConsultaPedido
            PreencherCampos

'-------------PRIMEIRO
        Case "primeiro"
            rsPedido.MoveFirst
            PreencherCampos

'-------------ANTERIOR
        Case "anterior"
            If Not rsPedido.BOF Then rsPedido.MovePrevious
            If rsPedido.BOF Then rsPedido.MoveFirst
            PreencherCampos

'-------------PROXIMO
        Case "proximo"
            If Not rsPedido.EOF Then rsPedido.MoveNext
            If rsPedido.EOF Then rsPedido.MoveLast
            PreencherCampos

'-------------ULTIMO
        Case "ultimo"
            rsPedido.MoveLast
            PreencherCampos
        
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer) 'No Unload do formulario fecha o recordset
    If Not rsPedido Is Nothing Then 'Se ele não for nada (se existir)
        If rsPedido.State = adStateOpen Then rsPedido.Close 'Se esta aberto, fecha
        Set rsPedido = Nothing 'Seta como nada
    End If
    If Not rsPedidoItem Is Nothing Then 'Se ele não for nada (se existir)
        If rsPedidoItem.State = adStateOpen Then rsPedidoItem.Close 'Se esta aberto, fecha
        Set rsPedidoItem = Nothing 'Seta como nada
    End If
    If Not rsProximoCodigo Is Nothing Then 'Se ele não for nada (se existir)
        If rsProximoCodigo.State = adStateOpen Then rsProximoCodigo.Close 'Se esta aberto, fecha
        Set rsProximoCodigo = Nothing 'Seta como nada
    End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    
    If ModoAtualPedido = mfConsulta Then
        If KeyAscii = vbKeyReturn Then 'KeyCode do Enter
            KeyAscii = 0   ' evita o bip
            
            Dim codigoBusca As Long
    
            If Trim(txtCodigo.Text) = "" Then Exit Sub
            If Not IsNumeric(txtCodigo.Text) Then
                MsgBox "Código inválido.", vbExclamation
                txtCodigo.SetFocus
                Exit Sub
            End If
        
            codigoBusca = CLng(txtCodigo.Text)
            
            If BuscarRS(rsPedido, "Codigo", codigoBusca) Then
                PreencherCampos
            Else
                MsgBox "Não encontrado"
            End If
        End If
    End If
    
End Sub

Private Sub cmdListaPedido_Click()
    Dim f As New frmPesquisaPedido

    f.Show vbModal

    If f.CodigoSelecionado > 0 Then
        If BuscarRS(rsPedido, "Codigo", f.CodigoSelecionado) Then
            PreencherCampos
        End If
    End If

    Unload f
End Sub

Private Sub cmdListaCliente_Click()
    Dim f As New frmPesquisaCliente

    f.Show vbModal

    If f.CodigoSelecionado > 0 Then
        If BuscarRS(rsCliente, "Codigo", f.CodigoSelecionado) Then
            PreencherClientePedido
        End If
    End If

    Unload f
End Sub

Private Sub cmdListaProduto_Click()
    Dim f As New frmPesquisaProduto

    f.Show vbModal

    If f.CodigoSelecionado > 0 Then
        If BuscarRS(rsProduto, "Codigo", f.CodigoSelecionado) Then
            PreencherCamposItemInclusao
        End If
    End If

    Unload f
End Sub

Private Sub PreencherCamposItemInclusao()
    txtCodProduto.Text = rsProduto!Codigo
    txtDescricao.Text = rsProduto!Nome
    txtQtde.Text = 1
    txtValorUn.Text = rsProduto!valor
End Sub


Private Sub PreencherCamposItem() 'Evento de preenchimentos dos campos do item para ser chamado manualmente
    
    ControlePedidoItem = 0
    txtCodProduto.Text = ""
    txtDescricao.Text = ""
    txtQtde.Text = ""
    txtValorUn.Text = ""
    txtTotalItem.Text = ""
    
    ' Ignora cabeçalho ou grid vazio
    If grdItensPedido.Rows <= 1 Then Exit Sub
    If grdItensPedido.Row < 1 Then Exit Sub
    
    ControlePedidoItem = grdItensPedido.TextMatrix(grdItensPedido.Row, 0)
    txtCodProduto.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 1)
    txtDescricao.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 2)
    txtQtde.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 3)
    txtValorUn.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 4)
    txtTotalItem.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 5)
    
End Sub

Private Sub grdItensPedido_RowColChange() 'Quando muda de coluna ou de linha atualiza os campos do produto
    PreencherCamposItem
End Sub
