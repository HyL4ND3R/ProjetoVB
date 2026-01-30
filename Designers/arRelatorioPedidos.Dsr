VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arRelatorioPedidos 
   Caption         =   "Impressao"
   ClientHeight    =   15615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   28560
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   50377
   _ExtentY        =   27543
   SectionData     =   "arRelatorioPedidos.dsx":0000
End
Attribute VB_Name = "arRelatorioPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fItem As Integer
Private CodPedido As Integer
Private rs As ADODB.Recordset

Private Sub ActiveReport_ReportStart()
    
    Set rs = Me.dcRelPedidos.Recordset
    CodPedido = 0
    primeiroItem = True
    fItem = 0
    
    'Ajustando formatação dos campos
    fldQtdeTotal.OutputFormat = "#,##0.00"
    fldValorTotal.OutputFormat = "#,##0.00"
    fldProdutoQtde.OutputFormat = "#,##0.00"
    fldProdutoValorUn.OutputFormat = "#,##0.00"
    fldProdutoValorTotal.OutputFormat = "#,##0.00"
    
End Sub

Private Sub Detail_Format()

    If (CodPedido = CLng(rs!pedido)) Then
        'Ocultado os campos para somente mostrar o item
        fldPedido.Visible = False
        fldCliente.Visible = False
        fldData.Visible = False
        fldQtdeTotal.Visible = False
        fldValorTotal.Visible = False
        lblProdutoCod.Visible = False
        lblProduto.Visible = False
        lblProdutoQtde.Visible = False
        lblProdutoValorUn.Visible = False
        lblProdutoValorTotal.Visible = False
        'Puxando o item para cima para ficarem todos colados um abaixo do outro
        fldProdutoCod.Top = 0
        fldProduto.Top = 0
        fldProdutoQtde.Top = 0
        fldProdutoValorUn.Top = 0
        fldProdutoValorTotal.Top = 0
        'Ajustando Tamanho do Detail para ficar pequeno e só mostrar a linha do item
        Detail.Height = 284
    Else
        'Atualizando a variavel de controle do Cod Pedido
        CodPedido = CLng(rs!pedido)
        'Incrementando a variavel de controle do fundo
        fItem = fItem + 1
        'Mostrando novamente os campos pois é um pedido novo
        fldPedido.Visible = True
        fldCliente.Visible = True
        fldData.Visible = True
        fldQtdeTotal.Visible = True
        fldValorTotal.Visible = True
        lblProdutoCod.Visible = True
        lblProduto.Visible = True
        lblProdutoQtde.Visible = True
        lblProdutoValorUn.Visible = True
        lblProdutoValorTotal.Visible = True
        'Puxando os itens para baixo para mostrar os dados do pedido
        fldProdutoCod.Top = 750
        fldProduto.Top = 750
        fldProdutoQtde.Top = 750
        fldProdutoValorUn.Top = 750
        fldProdutoValorTotal.Top = 750
        'Ajustando tamanho do Detail para ficar grande e mostrar tudo
        Detail.Height = 870
    End If
    
        'Alterar entre fundo branco e cinza
    If fItem Mod 2 = 0 Then
        Detail.BackColor = &HFFFFFF
    Else
        Detail.BackColor = &HE0E0E0
    End If
    
    
    If (IsNull(rs!ProdutoCod)) Then
        lblProdutoCod.Visible = False
        lblProduto.Visible = False
        lblProdutoQtde.Visible = False
        lblProdutoValorUn.Visible = False
        lblProdutoValorTotal.Visible = False
        'Ajustado tamanho do Detail para ficar pequeno e só mostrar a linha do pedido
        Detail.Height = 284
        
    End If

    
End Sub

Private Sub PageFooter_Format()
    
    fldOperador = "Operador: " & rsOperadorLogado!Codigo & " - " & rsOperadorLogado!Nome & ", Emitido em: " & _
                    Format(Date, "dd/MM/yyyy") & " as " & Left(Time, 5)
    fldPagina = "Página: " & Format(pageNumber, "###000")
    
End Sub

