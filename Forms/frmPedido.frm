VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
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
   Begin MSMask.MaskEdBox mskDataPedido 
      Height          =   375
      Left            =   5730
      TabIndex        =   26
      Top             =   900
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   7
      Text            =   "0,00"
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
      TabIndex        =   6
      Text            =   "0,00"
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
      TabIndex        =   5
      Text            =   "0,00"
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
      TabIndex        =   3
      Top             =   2730
      Width           =   1215
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
      Left            =   3150
      TabIndex        =   4
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
      TabIndex        =   17
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
      TabIndex        =   14
      Text            =   "0,00"
      Top             =   6060
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grdItensPedido 
      Height          =   2805
      Left            =   2700
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   8
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   20895
      _ExtentX        =   36856
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "visualizar"
            ImageIndex      =   10
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
            NumListImages   =   10
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
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":A770
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5010
      TabIndex        =   27
      Top             =   930
      Width           =   705
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   10
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
      TabIndex        =   9
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
    InicializarCamposNumericosItens
    InicializarCamposNumericosItensPedido

    If Not rsPedido.EOF Then 'Se não esta no fim da lista
        rsPedido.MoveLast 'Move para o final
        PreencherCampos
    Else
        mskDataPedido.Mask = ""
        mskDataPedido.Text = ""
        mskDataPedido.Mask = "99/99/9999"
    End If

    ModoConsultaPedido
    
End Sub

Private Sub cmdExcluirItem_Click()
            
    If grdItensPedido.RowSel < 1 Then Exit Sub
    
    'Mensagem de confirmação, se clicar no Não, cai fora da sub
    If MsgBox("Deseja realmente excluir este Registro?", _
              vbQuestion + vbYesNo, _
              "Confirmação") = vbNo Then Exit Sub

    If Not IsNull(ControlePedidoItem) Then
        Conn.Execute "DELETE FROM PedidoItem WHERE Controle = " & ControlePedidoItem
    Else
        MsgBox "Controle não encontrado", vbOKOnly
    End If
    
    If (Not CalculaTotaisPedido(ControlePedido)) Then
        MsgBox "Erro ao calcular os totais do pedido!", vbOKOnly
    End If
    
    CarregarPedidos
    CarregarItensPedido CLng(txtCodigo.Text)
    PreencherItensPedido
    CancelarItem

End Sub

Private Sub cmdCancelarItem_Click()
    CancelarItem
End Sub

Private Sub cmdNovoItem_Click()
    ModoInclusaoItem
End Sub

Private Sub cmdAlterarItem_Click()
    If grdItensPedido.RowSel < 1 Then Exit Sub
    ModoAlteracaoItem
End Sub

Private Sub cmdSalvarItem_Click()
    Set pedidoItem = New cPedidoItem
    Dim CodigoAtual As Long

    'Validar os campos antes de tentar inserir
    If Not ValidaCamposItem Then Exit Sub

    'Inserção usando o Objeto cPedidoItem para passar os dados
    If ModoAtualItens = mfAlteracao Then
        pedidoItem.Controle = VerificaNull(ControlePedidoItem, 0)
        pedidoItem.ControlePedido = VerificaNull(ControlePedido, 0)
        BuscaProximoCodItemPedido (ControlePedido) 'Buscar o próximo codigo de Item
        pedidoItem.Item = CLng(rsPedidoItemCod!Item)
        pedidoItem.ProdutoCodigo = CLng(txtCodProduto.Text)
        pedidoItem.Descricao = txtNomeProduto.Text
        pedidoItem.Qtde = CDbl(txtQtde.Text)
        pedidoItem.ValorUn = CDbl(txtValorUn.Text)
        pedidoItem.ValorTotal = CDbl(txtQtde.Text) * CDbl(txtValorUn.Text)
        If (Not AlterarItemPedido(pedidoItem)) Then
            MsgBox "Erro ao Alterar o Registro!"
            Exit Sub
        End If
    Else
        pedidoItem.ControlePedido = VerificaNull(ControlePedido, 0)
        BuscaProximoCodItemPedido (ControlePedido) 'Buscar o próximo codigo de Item
        pedidoItem.Item = CLng(rsPedidoItemCod!Item)
        pedidoItem.ProdutoCodigo = CLng(txtCodProduto.Text)
        pedidoItem.Descricao = txtNomeProduto.Text
        pedidoItem.Qtde = CDbl(txtQtde.Text)
        pedidoItem.ValorUn = CDbl(txtValorUn.Text)
        pedidoItem.ValorTotal = CDbl(txtQtde.Text) * CDbl(txtValorUn.Text)
        If (Not InserirItemPedido(pedidoItem)) Then
            MsgBox "Erro ao Inserir o Registro!"
            Exit Sub
        End If
    End If
    
    If (Not CalculaTotaisPedido(ControlePedido)) Then
        MsgBox "Erro ao calcular os totais do pedido!", vbOKOnly
    End If
    
    CarregarPedidos
    PreencherItensPedido
    ModoInclusaoItem
    
End Sub

Private Sub SalvarPedido()
    If ModoAtualPedido = mfAlteracao Then
        pedido.Controle = VerificaNull(ControlePedido, 0)
        pedido.Codigo = CLng(txtCodigo.Text) 'Conversão de Texto para Long
        pedido.ClienteCodigo = CLng(txtCodCliente.Text)
        pedido.DataPedido = Format(mskDataPedido.Text, Date)
        If (Not AlterarPedido(pedido)) Then
            MsgBox "Erro ao Alterar o Registro!"
            Exit Sub
        End If
    Else
        pedido.Controle = VerificaNull(ControlePedido, 0)
        pedido.Codigo = CLng(txtCodigo.Text) 'Conversão de Texto para Long
        pedido.ClienteCodigo = CLng(txtCodCliente.Text)
        pedido.DataPedido = Format(mskDataPedido.Text, Date)
        If (Not InserirPedido(pedido)) Then
            MsgBox "Erro ao Inserir o Registro!"
            Exit Sub
        End If
    End If
End Sub

Private Sub ModoInclusaoPedido()
'--------------TOOLBAR------------------------
    Toolbar.Buttons("novo").Enabled = False 'Habilitar/Desabilitar botão da toolbar
    Toolbar.Buttons("salvar").Enabled = True
    Toolbar.Buttons("alterar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = False
    Toolbar.Buttons("desfazer").Enabled = True
    Toolbar.Buttons("primeiro").Enabled = False
    Toolbar.Buttons("anterior").Enabled = False
    Toolbar.Buttons("proximo").Enabled = False
    Toolbar.Buttons("ultimo").Enabled = False
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = False 'Habilitar/Desabilitar txt
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaPedido.Enabled = False 'Habilitar/Desabilitar commandButton
    txtCodCliente.Enabled = True
    txtCodCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    cmdListaCliente.Enabled = True 'Habilitar/Desabilitar commandButton
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    mskDataPedido.Enabled = True
    mskDataPedido.BackColor = vbWindowBackground
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
    txtNomeProduto.Enabled = False
    txtNomeProduto.BackColor = &H8000000F
    txtQtde.Enabled = False
    txtQtde.BackColor = &H8000000F
    txtValorUn.Enabled = False
    txtValorUn.BackColor = &H8000000F
    txtTotalItem.Enabled = False
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualPedido = mfInclusao
    txtCodCliente.SetFocus
End Sub

Private Sub ModoAlteracaoPedido()
'--------------TOOLBAR------------------------
    Toolbar.Buttons("novo").Enabled = False 'Habilitar/Desabilitar botão da toolbar
    Toolbar.Buttons("salvar").Enabled = True
    Toolbar.Buttons("alterar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = False
    Toolbar.Buttons("desfazer").Enabled = True
    Toolbar.Buttons("primeiro").Enabled = False
    Toolbar.Buttons("anterior").Enabled = False
    Toolbar.Buttons("proximo").Enabled = False
    Toolbar.Buttons("ultimo").Enabled = False
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = False 'Habilitar/Desabilitar txt
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaPedido.Enabled = False 'Habilitar/Desabilitar commandButton
    txtCodCliente.Enabled = True
    txtCodCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    cmdListaCliente.Enabled = True 'Habilitar/Desabilitar commandButton
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    mskDataPedido.Enabled = True
    mskDataPedido.BackColor = vbWindowBackground
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
    txtNomeProduto.Enabled = False
    txtNomeProduto.BackColor = &H8000000F
    txtQtde.Enabled = False
    txtQtde.BackColor = &H8000000F
    txtValorUn.Enabled = False
    txtValorUn.BackColor = &H8000000F
    txtTotalItem.Enabled = False
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualPedido = mfAlteracao
End Sub

Private Sub ModoConsultaPedido()
'--------------TOOLBAR------------------------
    Toolbar.Buttons("novo").Enabled = True
    Toolbar.Buttons("salvar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = True
    Toolbar.Buttons("alterar").Enabled = True
    Toolbar.Buttons("desfazer").Enabled = False
    Toolbar.Buttons("primeiro").Enabled = True
    Toolbar.Buttons("anterior").Enabled = True
    Toolbar.Buttons("proximo").Enabled = True
    Toolbar.Buttons("ultimo").Enabled = True
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = True
    txtCodigo.BackColor = vbWindowBackground
    cmdListaPedido.Enabled = True
    txtCodCliente.Enabled = False
    txtCodCliente.BackColor = &H8000000F
    cmdListaCliente.Enabled = False
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = &H8000000F
    mskDataPedido.Enabled = False
    mskDataPedido.BackColor = &H8000000F
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
    txtNomeProduto.Enabled = False
    txtNomeProduto.BackColor = &H8000000F
    txtQtde.Enabled = False
    txtQtde.BackColor = &H8000000F
    txtValorUn.Enabled = False
    txtValorUn.BackColor = &H8000000F
    txtTotalItem.Enabled = False
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualPedido = mfConsulta
End Sub

Private Sub ModoInclusaoItem()
'--------------TOOLBAR------------------------
    Toolbar.Buttons("novo").Enabled = False
    Toolbar.Buttons("salvar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = False
    Toolbar.Buttons("alterar").Enabled = False
    Toolbar.Buttons("desfazer").Enabled = False
    Toolbar.Buttons("primeiro").Enabled = False
    Toolbar.Buttons("anterior").Enabled = False
    Toolbar.Buttons("proximo").Enabled = False
    Toolbar.Buttons("ultimo").Enabled = False
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = False
    txtCodigo.BackColor = &H8000000F
    cmdListaPedido.Enabled = False
    txtCodCliente.Enabled = False
    txtCodCliente.BackColor = &H8000000F
    cmdListaCliente.Enabled = False
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = &H8000000F
    mskDataPedido.Enabled = False
    mskDataPedido.BackColor = &H8000000F
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
    txtNomeProduto.Enabled = True
    txtNomeProduto.Text = ""
    txtNomeProduto.BackColor = vbWindowBackground
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
    InicializarCamposNumericosItens
    txtCodProduto.SetFocus
    
End Sub

Private Sub ModoAlteracaoItem()
'--------------TOOLBAR------------------------
    Toolbar.Buttons("novo").Enabled = False
    Toolbar.Buttons("salvar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = False
    Toolbar.Buttons("alterar").Enabled = False
    Toolbar.Buttons("desfazer").Enabled = False
    Toolbar.Buttons("primeiro").Enabled = False
    Toolbar.Buttons("anterior").Enabled = False
    Toolbar.Buttons("proximo").Enabled = False
    Toolbar.Buttons("ultimo").Enabled = False
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = False
    txtCodigo.BackColor = &H8000000F
    cmdListaPedido.Enabled = False
    txtCodCliente.Enabled = False
    txtCodCliente.BackColor = &H8000000F
    cmdListaCliente.Enabled = False
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = &H8000000F
    mskDataPedido.Enabled = False
    mskDataPedido.BackColor = &H8000000F
'--------------BOTÕES ITENS------------------------
    cmdNovoItem.Enabled = False
    cmdSalvarItem.Enabled = True
    cmdAlterarItem.Enabled = False
    cmdExcluirItem.Enabled = False
    cmdCancelarItem.Enabled = True
'--------------CAMPOS ITENS------------------------
    txtCodProduto.Enabled = True
    txtCodProduto.BackColor = vbWindowBackground
    cmdListaProduto.Enabled = True
    txtNomeProduto.Enabled = True
    txtNomeProduto.BackColor = vbWindowBackground
    txtQtde.Enabled = True
    txtQtde.BackColor = vbWindowBackground
    txtValorUn.Enabled = True
    txtValorUn.BackColor = vbWindowBackground
    txtTotalItem.Enabled = False
    txtTotalItem.BackColor = &H8000000F
    
    ModoAtualItens = mfAlteracao
    txtCodProduto.SetFocus
    
End Sub

Private Sub CancelarItem()
'--------------TOOLBAR------------------------
    Toolbar.Buttons("novo").Enabled = True
    Toolbar.Buttons("salvar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = True
    Toolbar.Buttons("alterar").Enabled = True
    Toolbar.Buttons("desfazer").Enabled = False
    Toolbar.Buttons("primeiro").Enabled = True
    Toolbar.Buttons("anterior").Enabled = True
    Toolbar.Buttons("proximo").Enabled = True
    Toolbar.Buttons("ultimo").Enabled = True
'--------------CAMPOS PEDIDO------------------------
    txtCodigo.Enabled = True
    txtCodigo.BackColor = vbWindowBackground
    cmdListaPedido.Enabled = True
    txtCodCliente.Enabled = False
    txtCodCliente.BackColor = &H8000000F
    cmdListaCliente.Enabled = False
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = &H8000000F
    mskDataPedido.Enabled = False
    mskDataPedido.BackColor = &H8000000F
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
    txtNomeProduto.Enabled = False
    txtNomeProduto.BackColor = &H8000000F
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
    
    'Se tiver vazio limpa os campos
    If rsPedido.EOF Or rsPedido.BOF = True Then 'Se a lista não tem registros pula fora da Sub
        txtCodigo.Text = ""
        txtCodCliente.Text = ""
        txtNomeCliente.Text = ""
        mskDataPedido.Mask = ""
        mskDataPedido.Text = ""
        mskDataPedido.Mask = "99/99/9999"
        Exit Sub
    End If

    ControlePedido = rsPedido!Controle
    txtCodigo.Text = rsPedido!Codigo 'Atribuição de valor do RecordSet para o TextBox
    txtCodCliente.Text = rsPedido!ClienteCodigo
    txtNomeCliente.Text = rsPedido!ClienteNome
    
    mskDataPedido.Mask = ""
    mskDataPedido.Text = ""
    mskDataPedido.Mask = "99/99/9999"
    mskDataPedido.Text = Format(rsPedido!DataPedido, "dd/MM/yyyy")
    
    txtValorTotal.Text = Format((IIf(IsNull(rsPedido!ValorTotal), 0, rsPedido!ValorTotal)), "0.00")
    
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
    LimpaCamposItens
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
    
    If IsNumeric(txtCodigo.Text) Then
        BuscarRS rsPedido, "Codigo", txtCodigo.Text
        txtValorTotal.Text = Format(rsPedido!ValorTotal, "0.00")
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

'---------------------Case Toolbar---------------------------------------------------
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
'-------------NOVO-------------------------------------------------------------------
        Case "novo"
            
            BuscarProximoCodPedido
            txtCodigo.Text = rsProximoCodigo!Codigo
            txtCodCliente.Text = ""
            txtNomeCliente.Text = ""
            
            mskDataPedido.Mask = ""
            mskDataPedido.Text = ""
            mskDataPedido.Mask = "99/99/9999"
            mskDataPedido.Text = Format(Date, "dd/MM/yyyy")
            
            txtValorTotal.Text = ""
            
            ModoInclusaoPedido

'-------------SALVAR-----------------------------------------------------------------
        Case "salvar"
            Dim CodigoAtual As Long
            
            'Validar os campos antes de tentar inserir
            If Not ValidaCamposPedido Then Exit Sub

            If ModoAtualPedido = mfAlteracao Then
                pedido.Controle = VerificaNull(ControlePedido, 0)
                pedido.Codigo = CLng(txtCodigo.Text) 'Conversão de Texto para Long
                pedido.ClienteCodigo = CLng(txtCodCliente.Text)
                pedido.DataPedido = mskDataPedido.Text
                If (Not AlterarPedido(pedido)) Then
                    MsgBox "Erro ao Alterar o Registro!"
                    Exit Sub
                End If
            Else
                pedido.Controle = VerificaNull(ControlePedido, 0)
                pedido.Codigo = CLng(txtCodigo.Text) 'Conversão de Texto para Long
                pedido.ClienteCodigo = CLng(txtCodCliente.Text)
                pedido.DataPedido = mskDataPedido.Text
                If (Not InserirPedido(pedido)) Then
                    MsgBox "Erro ao Inserir o Registro!"
                    Exit Sub
                End If
            End If
            
            CarregarPedidos
            
            If ModoAtualPedido = mfAlteracao Then
                rsPedido.Find "Codigo = " & CodigoAtual
            Else
                If Not rsPedido.EOF Then rsPedido.MoveLast
            End If
            
            PreencherCampos
            ModoConsultaPedido

'-------------ALTERACAO--------------------------------------------------------------
        Case "alterar"
        
            If rsPedido.EOF Or rsPedido.BOF Then Exit Sub
            
            PreencherCampos
            ModoAlteracaoPedido

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
            ModoConsultaPedido

'-------------DESFAZER
        Case "desfazer"
            txtCodigo.Text = ""
            ModoConsultaPedido
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
            
'-------------VISUALIZAR
        Case "visualizar"
            Dim rpt As New arRelatorioPedidos
            Dim Sql As String
            
            'Define a Conexão com o Banco
            rpt.dcRelPedidos.ConnectionString = Conn
            
            Sql = "select Pedido.Codigo As Pedido, Cliente.Nome As Cliente, pedido.Data As DataPedido, " & _
                    "Pedido.QtdeTotal As QtdeTotal, Pedido.ValorTotal As ValorTotal, " & _
                    "PedidoItem.ProdutoCodigo As ProdutoCod,  PedidoItem.Descricao As Produto, " & _
                    "PedidoItem.Quantidade As ProdutoQtde, PedidoItem.ValorUn As ProdutoValorUn, " & _
                    "PedidoItem.ValorTotal As ProdutoValorTotal " & _
                    "From pedido " & _
                    "Inner join Cliente on Pedido.ClienteCodigo = Cliente.Codigo " & _
                    "Left join PedidoItem  on PedidoItem.ControlePedido = Pedido.Controle " & _
                    "Order by PedidoItem.Item"
            
            'Define a string que vai ser executada no banco
            rpt.dcRelPedidos.Source = Sql
            
            rpt.Run
            
            rpt.Show vbModal
        
    End Select

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
    txtNomeProduto.Text = rsProduto!Nome
    txtQtde.Text = 1
    txtValorUn.Text = rsProduto!Valor
End Sub

Private Sub LimpaCamposItens()
    ControlePedidoItem = 0
    txtCodProduto.Text = ""
    txtNomeProduto.Text = ""
    txtQtde.Text = ""
    txtValorUn.Text = ""
    txtTotalItem.Text = ""
End Sub

Private Sub PreencherCamposItem() 'Evento de preenchimentos dos campos do item para ser chamado manualmente
    
    LimpaCamposItens
    
    ' Ignora cabeçalho ou grid vazio
    If grdItensPedido.Rows <= 1 Then Exit Sub
    If grdItensPedido.Row < 1 Then Exit Sub
    
    ControlePedidoItem = grdItensPedido.TextMatrix(grdItensPedido.Row, 0)
    txtCodProduto.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 1)
    txtNomeProduto.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 2)
    txtQtde.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 3)
    txtValorUn.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 4)
    txtTotalItem.Text = grdItensPedido.TextMatrix(grdItensPedido.Row, 5)
    
End Sub

'Função para atualizar os campos do produto com base no item selecionado no Grid
Private Sub grdItensPedido_RowColChange() 'Quando muda de coluna ou de linha atualiza os campos do produto
    PreencherCamposItem
End Sub

'Função para validar os campos ao salvar o pedido
Private Function ValidaCamposPedido()
    
    If Not IsNumeric(txtCodigo.Text) Then
        MsgBox "Codigo Pedido inválido"
        txtCodigo.SetFocus
        ValidaCamposPedido = False
        Exit Function
    End If
    
    If Not IsNumeric(txtCodCliente.Text) Then
        MsgBox "Cliente inválido"
        txtCodCliente.SetFocus
        ValidaCamposPedido = False
        Exit Function
    End If
    
    If Not IsDate(mskDataPedido.Text) Then
        MsgBox "Data Inválida"
        mskDataPedido.SetFocus
        ValidaCamposPedido = False
        Exit Function
    End If
    
    ValidaCamposPedido = True
    
End Function

'Função para validar os campos ao salvar o item
Private Function ValidaCamposItem()
    
    If Not IsNumeric(txtCodProduto.Text) Then
        MsgBox "Codigo Produto inválido"
        txtCodProduto.SetFocus
        ValidaCamposItem = False
        Exit Function
    ElseIf CLng(txtCodProduto.Text) = 0 Then
        MsgBox "Codigo Produto inválido"
        txtCodProduto.SetFocus
        ValidaCamposItem = False
        Exit Function
    End If
    
    If Not IsNumeric(txtQtde.Text) Then
        MsgBox "Quantidade inválida"
        txtQtde.SetFocus
        ValidaCamposItem = False
        Exit Function
    ElseIf CLng(txtQtde.Text) = 0 Then
        MsgBox "Quantidade inválida"
        txtQtde.SetFocus
        ValidaCamposItem = False
        Exit Function
    End If
    
    If Not IsNumeric(txtValorUn.Text) Then
        MsgBox "Valor Inválido inválido"
        txtValorUn.SetFocus
        ValidaCamposItem = False
        Exit Function
    ElseIf CLng(txtValorUn.Text) = 0 Then
        MsgBox "Valor Inválido inválido"
        txtValorUn.SetFocus
        ValidaCamposItem = False
        Exit Function
    End If
    
    ValidaCamposItem = True
    
End Function
'Ajustar os campos de valor e quantidade para sempre começar com valor
Private Sub InicializarCamposNumericosItens()
    txtQtde.Text = IIf(ModoAtualItens = mfInclusao, "1", "0")
    txtValorUn.Text = "0,00"
    txtTotalItem.Text = "0,00"
End Sub

Private Sub InicializarCamposNumericosItensPedido()
    txtValorTotal.Text = "0,00"
End Sub

'----------------------------AJUSTES CAMPOS DO CORPO DO PEDIDO ----------------------------------

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
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
    
        ' Só números
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then 'Se a tecla for enter
        
        If Not IsNumeric(txtCodCliente.Text) Then 'Validação de Numérico
            MsgBox "Código Inválido", vbOKOnly 'Aviso de código invalido
            txtCodCliente.SetFocus 'Volta para o campo CodCliente
            Exit Sub 'Sai da Sub
        End If 'Se não
        
        BuscarClientePorCodigo CLng(txtCodCliente.Text) 'Busca o Cliente pelo Codigo
        
        If Not rsClienteCod.BOF Or Not rsClienteCod.EOF Then 'Se a lista não esta vazia
            txtCodCliente.Text = rsClienteCod!Codigo 'Atribui o Codigo ao Campo
            txtNomeCliente.Text = rsClienteCod!Nome 'Atribui o Nome ao Campo
        Else 'Se a Lista esta vazia
            MsgBox "Código não Encontrado", vbOKOnly 'Mensagem de aviso
            txtCodCliente.SetFocus 'Volta para o campo CodCliente
            Exit Sub 'Sai da sub
        End If
        
        AvancarComEnterKD KeyAscii, mskDataPedido 'Se tudo deu certo, avança para o próximo campo
    End If
    
    ' Só números
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
        
End Sub

Private Sub mskDataPedido_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub

    If KeyAscii = vbKeyReturn Then 'Verifica se a tecla digitada é enter
        
        KeyAscii = 0 'Limpa o Teclado
        
        If MsgBox("Confirma Dados?", _
            vbQuestion + vbYesNo, _
            "Confirmação") = vbNo Then 'Faz a pergunta, se não confirmar pula fora
            txtCodCliente.SetFocus
            Exit Sub
        End If
        
        SalvarPedido
    End If
End Sub

'----------------------------AJUSTES CAMPOS DOS ITENS----------------------------------

Private Sub txtCodProduto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then 'Se a tecla for enter
        
        If Not IsNumeric(txtCodProduto.Text) Then 'Validação de Numérico
            MsgBox "Código Inválido", vbOKOnly 'Aviso de código invalido
            txtCodProduto.SetFocus 'Volta para o campo CodProduto
            Exit Sub 'Sai da Sub
        End If 'Se não
        
        BuscarProdutoPorCodigo CLng(txtCodProduto.Text) 'Busca o Produto pelo Codigo
        
        If Not rsProdutoCod.BOF Or Not rsProdutoCod.EOF Then 'Se a lista não esta vazia
            txtCodProduto.Text = rsProdutoCod!Codigo 'Atribui o Codigo ao Campo
            txtNomeProduto.Text = rsProdutoCod!Nome 'Atribui o Nome ao Campo
        Else 'Se a Lista esta vazia
            MsgBox "Código não Encontrado", vbOKOnly 'Mensagem de aviso
            txtCodProduto.SetFocus 'Volta para o campo CodProduto
            Exit Sub 'Sai da sub
        End If
        
        AvancarComEnterKD KeyAscii, txtQtde 'Se tudo deu certo, avança para o próximo campo
    End If
    
    ' Só números
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub txtQtde_KeyPress(KeyAscii As Integer)

    ' Permite Backspace
    If KeyAscii = vbKeyBack Then Exit Sub

    ' ENTER
    If KeyAscii = vbKeyReturn Then
        
        AvancarComEnterKD KeyAscii, txtValorUn
        
        ' Valida Código Produto
        If Trim(txtCodProduto.Text) = "" Or Not IsNumeric(txtCodProduto.Text) Then
            MsgBox "Código Produto Inválido", vbOKOnly
            txtCodProduto.SetFocus
            Exit Sub
        End If
        
        ' Valida Quantidade
        If Trim(txtQtde.Text) = "" Or Not IsNumeric(txtQtde.Text) Then
            MsgBox "Quantidade Inválida", vbOKOnly
            txtQtde.SetFocus
            Exit Sub
        End If
        
        ' Busca Produto
        BuscarProdutoPorCodigo CLng(txtCodProduto.Text)
        
        If Not (rsProdutoCod.BOF And rsProdutoCod.EOF) Then
            
            ' ?? Só altera o ValorUnitário se estiver vazio ou zerado
            Dim valorZeradoOuVazio As Boolean
            valorZeradoOuVazio = False
            
            If Trim(txtValorUn.Text) = "" Then
                valorZeradoOuVazio = True
            ElseIf IsNumeric(txtValorUn.Text) Then
                If CDbl(txtValorUn.Text) = 0 Then
                    valorZeradoOuVazio = True
                End If
            End If
            
            If valorZeradoOuVazio Then
                txtValorUn.Text = Format(rsProdutoCod!Valor, "0.00")
            End If
            
            ' Calcula Total
            If IsNumeric(txtValorUn.Text) Then
                If IsNumeric(txtQtde.Text) Then
                    txtTotalItem.Text = Format( _
                        CDbl(txtValorUn.Text) * CDbl(txtQtde.Text), "0.00")
                Else
                    MsgBox "Quantidade inválida", vbOKOnly
                    txtQtde.SetFocus
                End If
            Else
                MsgBox "Valor inválido", vbOKOnly
                txtValorUn.SetFocus
            End If
            
        Else
            MsgBox "Código não Encontrado", vbOKOnly
            txtCodProduto.SetFocus
            Exit Sub
        End If
        
    End If

    ' Só números e vírgula
    If InStr("0123456789,", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If

    ' Só uma vírgula
    If Chr(KeyAscii) = "," And InStr(txtQtde.Text, ",") > 0 Then
        KeyAscii = 0
    End If

End Sub

'Ao entrar no campo seleciona todo o texto automaticamente
Private Sub txtQtde_GotFocus()
    txtQtde.SelStart = 0
    txtQtde.SelLength = Len(txtQtde.Text)
End Sub

'Ao sair do campo formata o conteudo para decimal
Private Sub txtQtde_LostFocus()
    txtQtde.Text = FormataDecimal(txtQtde.Text)
End Sub


Private Sub txtValorUn_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then Exit Sub

    ' Calcula Total
    If IsNumeric(txtValorUn.Text) Then
        If IsNumeric(txtQtde.Text) Then
            txtTotalItem.Text = Format( _
                CDbl(txtValorUn.Text) * CDbl(txtQtde.Text), "0.00")
        Else
            MsgBox "Quantidade inválida", vbOKOnly
            txtQtde.SetFocus
        End If
    Else
        MsgBox "Valor inválido", vbOKOnly
        txtValorUn.SetFocus
    End If
    
    If KeyAscii = vbKeyReturn Then 'Verifica se a tecla digitada é enter
        
        KeyAscii = 0 'Limpa o Teclado
        
        If MsgBox("Confirma Dados?", _
            vbQuestion + vbYesNo, _
            "Confirmação") = vbNo Then 'Faz a pergunta, se não confirmar pula fora
            txtCodProduto.SetFocus
            Exit Sub
        End If
        
        cmdSalvarItem_Click
        
    End If

    ' Só números e vírgula
    If InStr("0123456789,", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If

    ' Só uma vírgula
    If Chr(KeyAscii) = "," And InStr(txtValorUn.Text, ",") > 0 Then
        KeyAscii = 0
    End If

End Sub

'Ao entrar no campo já seleciona todo o conteudo automaticamente
Private Sub txtValorUn_GotFocus()
    txtValorUn.SelStart = 0
    txtValorUn.SelLength = Len(txtValorUn.Text)
End Sub

'Ao mudar algo no valor do campo já atualiza o valor do total item
Private Sub txtValorUn_Change()
    If IsNumeric(txtValorUn.Text) Then
        If IsNumeric(txtQtde.Text) Then
            txtTotalItem.Text = CDbl(txtValorUn.Text) * CDbl(txtQtde.Text)
            FormataDecimal txtTotalItem.Text
        End If
    End If
End Sub

'Ao Sair do Campo Formata ele para decimal
Private Sub txtValorUn_LostFocus()
    txtValorUn.Text = FormataDecimal(txtValorUn.Text)
End Sub

Private Sub txtValorTotal_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then Exit Sub

    ' Só números e vírgula
    If InStr("0123456789,", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If

    ' Só uma vírgula
    If Chr(KeyAscii) = "," And InStr(txtValorTotal.Text, ",") > 0 Then
        KeyAscii = 0
    End If

End Sub

'------------------------------UNLOAD DO FORMULARIO----------------------------------------------

'No Unload do formulario fecha os recordset's
Private Sub Form_Unload(Cancel As Integer)
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
    If Not rsClienteCod Is Nothing Then 'Se ele não for nada (se existir)
        If rsClienteCod.State = adStateOpen Then rsClienteCod.Close 'Se esta aberto, fecha
        Set rsClienteCod = Nothing 'Seta como nada
    End If
    If Not rsPedidoItemCod Is Nothing Then 'Se ele não for nada (se existir)
        If rsPedidoItemCod.State = adStateOpen Then rsPedidoItemCod.Close 'Se esta aberto, fecha
        Set rsPedidoItemCod = Nothing 'Seta como nada
    End If
End Sub
