VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProduto 
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   17805
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkInativo 
      Caption         =   "Inativo"
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
      Left            =   1950
      TabIndex        =   8
      Top             =   3000
      Width           =   1875
   End
   Begin VB.TextBox txtValor 
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
      Left            =   1950
      TabIndex        =   7
      Top             =   2340
      Width           =   1695
   End
   Begin VB.TextBox txtNome 
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
      Left            =   1950
      TabIndex        =   5
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CommandButton cmdListaProduto 
      DisabledPicture =   "frmProduto.frx":0000
      DownPicture     =   "frmProduto.frx":05E2
      Height          =   375
      Left            =   3150
      Picture         =   "frmProduto.frx":0BC4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1020
      Width           =   525
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
      Left            =   1950
      TabIndex        =   2
      Top             =   1020
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17805
      _ExtentX        =   31406
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
      MouseIcon       =   "frmProduto.frx":11A6
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   17190
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
               Picture         =   "frmProduto.frx":1E80
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduto.frx":2B5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduto.frx":3834
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduto.frx":450E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduto.frx":51E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduto.frx":5EC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduto.frx":649E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduto.frx":7178
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduto.frx":7E52
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblValor 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor:"
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
      Left            =   390
      TabIndex        =   6
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label lblNome 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
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
      Left            =   390
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
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
      Left            =   390
      TabIndex        =   1
      Top             =   1020
      Width           =   1455
   End
End
Attribute VB_Name = "frmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ModoAtual As eModoFormulario
Private CodigoAtual As Long

Private Sub Form_Load()

    txtCodigo.MaxLength = 9
    txtNome.MaxLength = 200
    txtValor.MaxLength = 10
    
    CarregarProdutos

    If Not rsProduto.EOF Then 'Se não esta no fim da lista
        rsProduto.MoveLast 'Move para o final
        PreencherCampos
    End If

    modoConsulta
    
End Sub

Private Sub modoInclusao()
    Toolbar.Buttons("novo").Enabled = False 'Habilitar/Desabilitar botão da toolbar
    Toolbar.Buttons("salvar").Enabled = True
    Toolbar.Buttons("alterar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = False
    Toolbar.Buttons("desfazer").Enabled = True
    Toolbar.Buttons("primeiro").Enabled = False
    Toolbar.Buttons("anterior").Enabled = False
    Toolbar.Buttons("proximo").Enabled = False
    Toolbar.Buttons("ultimo").Enabled = False
    txtCodigo.Enabled = False 'Habilitar/Desabilitar txt
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaProduto.Enabled = False 'Habilitar/Desabilitar commandButton
    txtNome.Enabled = True
    txtNome.BackColor = vbWindowBackground 'cor branca padrão do sistema
    txtValor.Enabled = True
    txtValor.BackColor = vbWindowBackground 'cor branca padrão do sistema
    chkInativo.Enabled = True
    ModoAtual = mfInclusao
    txtValor.Text = "0,00"
End Sub

Private Sub modoAlteracao()
    Toolbar.Buttons("novo").Enabled = False 'Habilitar/Desabilitar botão da toolbar
    Toolbar.Buttons("salvar").Enabled = True
    Toolbar.Buttons("alterar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = False
    Toolbar.Buttons("desfazer").Enabled = True
    Toolbar.Buttons("primeiro").Enabled = False
    Toolbar.Buttons("anterior").Enabled = False
    Toolbar.Buttons("proximo").Enabled = False
    Toolbar.Buttons("ultimo").Enabled = False
    txtCodigo.Enabled = False 'Habilitar/Desabilitar txt
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaProduto.Enabled = False 'Habilitar/Desabilitar commandButton
    txtNome.Enabled = True
    txtNome.BackColor = vbWindowBackground 'cor branca padrão do sistema
    txtValor.Enabled = True
    txtValor.BackColor = vbWindowBackground 'cor branca padrão do sistema
    chkInativo.Enabled = True
    ModoAtual = mfAlteracao
End Sub

Private Sub modoConsulta()
    Toolbar.Buttons("novo").Enabled = True
    Toolbar.Buttons("salvar").Enabled = False
    Toolbar.Buttons("excluir").Enabled = True
    Toolbar.Buttons("alterar").Enabled = True
    Toolbar.Buttons("desfazer").Enabled = False
    Toolbar.Buttons("primeiro").Enabled = True
    Toolbar.Buttons("anterior").Enabled = True
    Toolbar.Buttons("proximo").Enabled = True
    Toolbar.Buttons("ultimo").Enabled = True
    txtCodigo.Enabled = True
    txtCodigo.BackColor = vbWindowBackground&
    cmdListaProduto.Enabled = True
    txtNome.Enabled = False
    txtNome.BackColor = &H8000000F
    txtValor.Enabled = False
    txtValor.BackColor = &H8000000F
    chkInativo.Enabled = False
    ModoAtual = mfConsulta
End Sub


Private Sub PreencherCampos()

    If rsProduto.EOF Or rsProduto.BOF Then Exit Sub 'Se a lista não tem registros pula fora da Sub

    txtCodigo.Text = rsProduto!codigo 'Atribuição de valor do RecordSet para o TextBox
    txtNome.Text = rsProduto!Nome
    txtValor.Text = Format(rsProduto!Valor, "0.00")
    chkInativo.Value = IIf(rsProduto!Inativo = 1, vbChecked, vbUnchecked)

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    
'-------------NOVO
        Case "novo"
            txtCodigo.Text = ""
            txtNome.Text = ""
            txtValor.Text = ""
            chkInativo.Value = vbUnchecked
            modoInclusao

'-------------SALVAR
        Case "salvar"
            
            If Not ValidaCampos Then Exit Sub
            
            If Not SalvarProduto Then
                MsgBox "Erro ao salvar o Produto!", vbOKOnly
                Exit Sub
            End If
            
            CarregarProdutos
            
            If ModoAtual = mfAlteracao Then
                rsProduto.Find "Codigo = " & CodigoAtual
            Else
                If Not rsProduto.EOF Then rsProduto.MoveLast
            End If
            
            PreencherCampos
            modoConsulta

'-------------ALTERACAO
        Case "alterar"
        
            If rsProduto.EOF Or rsProduto.BOF Then Exit Sub
            
            PreencherCampos
            modoAlteracao

'-------------EXCLUIR
        Case "excluir"
            
            If rsProduto.EOF Or rsProduto.BOF Then Exit Sub
            
            'Mensagem de confirmação, se clicar no Não, cai fora da sub
            If MsgBox("Deseja realmente excluir este Produto?", _
                      vbQuestion + vbYesNo, _
                      "Confirmação") = vbNo Then Exit Sub

            Dim codigoExcluir As Long
            codigoExcluir = CLng(txtCodigo.Text)
        
            If Not ExcluirProduto(codigoExcluir) Then
                MsgBox "Erro ao Excluir o Produto!", vbInformation
                Exit Sub
            End If
        
            CarregarProdutos
        
            If Not rsProduto.EOF Then
                rsProduto.Find "Codigo > " & codigoExcluir
                If rsProduto.EOF Then rsProduto.MoveLast
            End If
            
            PreencherCampos
            modoConsulta

'-------------DESFAZER
        Case "desfazer"
            modoConsulta
            PreencherCampos

'-------------PRIMEIRO
        Case "primeiro"
            rsProduto.MoveFirst
            PreencherCampos

'-------------ANTERIOR
        Case "anterior"
            If Not rsProduto.BOF Then rsProduto.MovePrevious
            If rsProduto.BOF Then rsProduto.MoveFirst
            PreencherCampos

'-------------PROXIMO
        Case "proximo"
            If Not rsProduto.EOF Then rsProduto.MoveNext
            If rsProduto.EOF Then rsProduto.MoveLast
            PreencherCampos

'-------------ULTIMO
        Case "ultimo"
            rsProduto.MoveLast
            PreencherCampos
        
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer) 'No Unload do formulario fecha o recordset
    If Not rsProduto Is Nothing Then 'Se ele não for nada (se existir)
        If rsProduto.State = adStateOpen Then rsProduto.Close 'Se esta aberto, fecha
        Set rsProduto = Nothing 'Seta como nada
    End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    
    If ModoAtual = mfConsulta Then
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
            
            If BuscarRS(rsProduto, "Codigo", codigoBusca) Then
                PreencherCampos
            Else
                MsgBox "Não encontrado"
            End If
        End If
    End If
    
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtValor.SetFocus
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

    ' Permite Backspace
    If KeyAscii = vbKeyBack Then Exit Sub

    ' ENTER
    If KeyAscii = vbKeyReturn Then
    
        KeyAscii = 0
                
        If Not ValidaCampos Then Exit Sub
            
        If MsgBox("Confirma Dados?", _
            vbQuestion + vbYesNo, _
            "Confirmação") = vbNo Then 'Faz a pergunta, se não confirmar pula fora
            txtNome.SetFocus
            Exit Sub
        End If
            
        If Not SalvarProduto Then
            MsgBox "Erro ao salvar o Produto!", vbOKOnly
            Exit Sub
        End If
        
        CarregarProdutos
        
        If ModoAtual = mfAlteracao Then
            rsProduto.Find "Codigo = " & CodigoAtual
        Else
            If Not rsProduto.EOF Then rsProduto.MoveLast
        End If
        
        PreencherCampos
        modoConsulta
                
    End If

    ' Só números e vírgula
    If InStr("0123456789,", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If

    ' Só uma vírgula
    If Chr(KeyAscii) = "," And InStr(txtValor.Text, ",") > 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtValor_GotFocus()
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor.Text)
End Sub

Private Sub cmdListaProduto_Click()
    Dim f As New frmPesquisaProduto

    f.Show vbModal

    If f.CodigoSelecionado > 0 Then
        If BuscarRS(rsProduto, "Codigo", f.CodigoSelecionado) Then
            PreencherCampos
        End If
    End If

    Unload f
End Sub

Private Function SalvarProduto() As Boolean
    On Error GoTo Erro
    
    Dim produto As cProduto
    Set produto = New cProduto
    
    If ModoAtual = mfAlteracao Then
        produto.codigo = CLng(txtCodigo.Text)
        produto.Nome = txtNome.Text
        produto.Valor = CDbl(txtValor.Text)
        produto.Inativo = IIf(chkInativo.Value = vbChecked, 1, 0)
        If Not AlterarProduto(produto) Then
            MsgBox "Erro ao Alterar o Produto!", vbOKOnly
            Exit Function
        End If
    Else
        produto.Nome = txtNome.Text
        produto.Valor = CDbl(txtValor.Text)
        produto.Inativo = IIf(chkInativo.Value = vbChecked, 1, 0)
        If Not InserirProduto(produto) Then
            MsgBox "Erro ao Inserir o Produto!", vbOKOnly
            Exit Function
        End If
    End If
   
    SalvarProduto = True
    Exit Function
    
Erro:
    SalvarProduto = False
    
End Function


Private Function ValidaCampos() As Boolean
    
    If Trim(txtNome.Text = "") Then
        MsgBox "Nome Inválido"
        txtNome.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    If txtValor.Text = "" Or Not IsNumeric(txtValor.Text) Then
        MsgBox "Valor Inválido"
        txtValor.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    ValidaCampos = True
    
End Function

