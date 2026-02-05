VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
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
      Left            =   3240
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
      Width           =   825
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
      Top             =   1590
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
      Height          =   405
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
      Format          =   131334145
      CurrentDate     =   46051
   End
   Begin MSComCtl2.DTPicker dtpDataFinal 
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
      Format          =   131334145
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
Private CtrlATextos As Collection

Private Sub Form_Load()
    
    'Ajuste para todos os campos aceitarem ControlA
    Dim c As Control
    Dim h As cControlA
    
    Set CtrlATextos = New Collection
    
    For Each c In Me.Controls
        If TypeOf c Is TextBox Then
            Set h = New cControlA
            Set h.Txt = c
            CtrlATextos.Add h
        End If
    Next

    dtpDataInicial.Value = Date
    dtpDataFinal.Value = Date
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
            "Where Pedido.Data Between '" & Format(dtpDataInicial.Value, "yyyy-MM-dd") & "' and '" & Format(dtpDataFinal.Value, "yyyy-MM-dd") & "'"
            
    If Trim(txtCodCliente.Text <> "") Then
        Sql = Sql & " And Pedido.ClienteCodigo = " & txtCodCliente.Text
    End If
    
    If Trim(txtCodProduto.Text <> "") Then
        Sql = Sql & " And PedidoItem.ProdutoCodigo = " & txtCodProduto.Text
    End If
              
    Sql = Sql & " Order by Pedido.Codigo, PedidoItem.Item"
    
    'Define a string que vai ser executada no banco
    rpt.dcRelPedidos.Source = Sql
    
    If rpt.dcRelPedidos.Recordset.BOF Or rpt.dcRelPedidos.Recordset.EOF Then
        MsgBox "Nenhum Registro Encontrado!", vbOKOnly
        Exit Sub
    End If
    
    rpt.Run

    rpt.Show vbModal

End Sub

Private Sub dtpDataInicial_KeyDown(KeyAscii As Integer, Shift As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        If Not IsDate(dtpDataInicial.Value) Then
            MsgBox "Data inválida", vbOKOnly
            dtpDataInicial.SetFocus
            Exit Sub
        End If
        dtpDataFinal.SetFocus
    End If
    
End Sub

Private Sub dtpDataFinal_KeyDown(KeyAscii As Integer, Shift As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        If Not IsDate(dtpDataFinal.Value) Then
            MsgBox "Data inválida", vbOKOnly
            dtpDataFinal.SetFocus
            Exit Sub
        End If
        txtCodCliente.SetFocus
    End If
    
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then 'Se a tecla for enter
        
        If Trim(txtCodCliente.Text <> "") Then
            If Not IsNumeric(txtCodCliente.Text) Then 'Validação de Numérico
                MsgBox "Código Inválido", vbOKOnly 'Aviso de código invalido
                txtCodCliente.SetFocus 'Volta para o campo CodProduto
                Exit Sub 'Sai da Sub
            End If 'Se não
            
            BuscarClientePorCodigo CLng(txtCodCliente.Text) 'Busca o Cliente pelo Codigo
            
            If Not rsClienteCod.BOF Or Not rsClienteCod.EOF Then 'Se a lista não esta vazia
                txtCodCliente.Text = rsClienteCod!codigo 'Atribui o Codigo ao Campo
                txtNomeCliente.Text = rsClienteCod!Nome 'Atribui o Nome ao Campo
            Else 'Se a Lista esta vazia
                MsgBox "Código não Encontrado", vbOKOnly 'Mensagem de aviso
                txtCodCliente.SetFocus 'Volta para o campo CodProduto
                Exit Sub 'Sai da sub
            End If
            
        End If
        
        txtCodProduto.SetFocus 'Se tudo deu certo, avança para o próximo campo
        
    End If
    
    ' Só números
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub txtCodProduto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then 'Se a tecla for enter
        
        If Trim(txtCodProduto.Text <> "") Then
        
            If Not IsNumeric(txtCodProduto.Text) Then 'Validação de Numérico
                MsgBox "Código Inválido", vbOKOnly 'Aviso de código invalido
                txtCodProduto.SetFocus 'Volta para o campo CodProduto
                Exit Sub 'Sai da Sub
            End If 'Se não
            
            BuscarProdutoPorCodigo CLng(txtCodProduto.Text) 'Busca o Produto pelo Codigo
            
            If Not rsProdutoCod.BOF Or Not rsProdutoCod.EOF Then 'Se a lista não esta vazia
                txtCodProduto.Text = rsProdutoCod!codigo 'Atribui o Codigo ao Campo
                txtNomeProduto.Text = rsProdutoCod!Nome 'Atribui o Nome ao Campo
            Else 'Se a Lista esta vazia
                MsgBox "Código não Encontrado", vbOKOnly 'Mensagem de aviso
                txtCodProduto.SetFocus 'Volta para o campo CodProduto
                Exit Sub 'Sai da sub
            End If
            
        End If
        
        cmdVisualizar.SetFocus 'Se tudo deu certo, avança para o próximo campo
    End If
      
    ' Só números
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub cmdVisualizar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdVisualizar_Click
    End If
End Sub

Private Sub cmdListaCliente_Click()
    Dim f As New frmPesquisaCliente

    f.Show vbModal

    If f.CodigoSelecionado > 0 Then
        If BuscarRS(rsCliente, "Codigo", f.CodigoSelecionado) Then
            PreencherCliente
        End If
    End If

    Unload f
End Sub

Private Sub cmdListaProduto_Click()
    Dim f As New frmPesquisaProduto

    f.Show vbModal

    If f.CodigoSelecionado > 0 Then
        If BuscarRS(rsProduto, "Codigo", f.CodigoSelecionado) Then
            PreencherProduto
        End If
    End If

    Unload f
End Sub

Private Sub PreencherCliente()
    txtCodCliente.Text = rsCliente!codigo
    txtNomeCliente.Text = rsCliente!Nome
End Sub

Private Sub PreencherProduto()
    txtCodProduto.Text = rsProduto!codigo
    txtNomeProduto.Text = rsProduto!Nome
End Sub

Private Function ValidaCampos() As Boolean
        
    If Not IsDate(dtpDataInicial) Then
        MsgBox "Data Inicial Inválida!"
        dtpDataInicial.SetFocus
        ValidaCamposPedido = False
        Exit Function
    End If
    
    If Not IsDate(dtpDataFinal.Value) Then
        MsgBox "Data Final Inválida!"
        dtpDataFinal.SetFocus
        ValidaCamposPedido = False
        Exit Function
    End If
    
    If Trim(txtCodCliente.Text <> "") Then
        If Not IsNumeric(txtCodCliente.Text) Then
            MsgBox "Cliente inválido!"
            txtCodCliente.Text = ""
            txtNomeCliente.Text = ""
            txtCodCliente.SetFocus
            ValidaCamposPedido = False
            Exit Function
        End If
    End If
    
    If Trim(txtCodProduto.Text <> "") Then
        If Not IsNumeric(txtCodProduto.Text) Then
            MsgBox "Produto inválido!"
            txtCodProduto.Text = ""
            txtNomeProduto.Text = ""
            txtCodProduto.SetFocus
            ValidaCamposPedido = False
            Exit Function
        End If
    End If
    
    ValidaCampos = True
    
End Function
