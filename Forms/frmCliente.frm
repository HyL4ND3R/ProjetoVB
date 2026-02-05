VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCliente 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   8100
   ClientLeft      =   285
   ClientTop       =   630
   ClientWidth     =   18510
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   18510
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox mskDocumento 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
      _ExtentX        =   4471
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
      PromptChar      =   "_"
   End
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
      Left            =   2160
      TabIndex        =   6
      Top             =   3840
      Width           =   1875
   End
   Begin VB.TextBox txtTelefone 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton cmdListaCliente 
      DisabledPicture =   "frmCliente.frx":0000
      DownPicture     =   "frmCliente.frx":05E2
      Height          =   375
      Left            =   3360
      Picture         =   "frmCliente.frx":0BC4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   525
   End
   Begin VB.ComboBox cboTipoDocumento 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18510
      _ExtentX        =   32650
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "visualizar"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "frmCliente.frx":11A6
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   17880
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
               Picture         =   "frmCliente.frx":1E80
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":2B5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":3834
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":450E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":51E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":5EC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":649E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":7178
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":7E52
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCliente.frx":8424
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblTelefone 
      Alignment       =   1  'Right Justify
      Caption         =   "Telefone:"
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
      Left            =   600
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblDocumento 
      Alignment       =   1  'Right Justify
      Caption         =   "Documento:"
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
      Left            =   600
      TabIndex        =   9
      Top             =   2520
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
      Left            =   600
      TabIndex        =   8
      Top             =   1800
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
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ModoAtual As eModoFormulario
Private TipoDocumento As eTipoDocumentoCliente
Private CodigoAtual As Long
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
    
    txtCodigo.MaxLength = 9
    txtNome.MaxLength = 200
    txtTelefone.MaxLength = 17
    
    popularComboTipoDocumento
    CarregarClientes

    If Not rsCliente.EOF Then 'Se não esta no fim da lista
        rsCliente.MoveLast 'Move para o final
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
    cmdListaCliente.Enabled = False 'Habilitar/Desabilitar commandButton
    txtNome.Enabled = True
    txtNome.BackColor = vbWindowBackground 'cor branca padrão do sistema
    cboTipoDocumento.Enabled = True
    mskDocumento.Enabled = True
    mskDocumento.BackColor = vbWindowBackground 'cor branca padrão do sistema
    txtTelefone.Enabled = True
    txtTelefone.BackColor = vbWindowBackground 'cor branca padrão do sistema
    chkInativo.Enabled = True
    ModoAtual = mfInclusao
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
    cmdListaCliente.Enabled = False 'Habilitar/Desabilitar commandButton
    txtNome.Enabled = True
    txtNome.BackColor = vbWindowBackground 'cor branca padrão do sistema
    cboTipoDocumento.Enabled = True
    mskDocumento.Enabled = True
    mskDocumento.BackColor = vbWindowBackground 'cor branca padrão do sistema
    txtTelefone.Enabled = True
    txtTelefone.BackColor = vbWindowBackground 'cor branca padrão do sistema
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
    cmdListaCliente.Enabled = True
    txtNome.Enabled = False
    txtNome.BackColor = &H8000000F
    cboTipoDocumento.Enabled = False
    mskDocumento.Enabled = False
    mskDocumento.BackColor = &H8000000F
    txtTelefone.Enabled = False
    txtTelefone.BackColor = &H8000000F
    chkInativo.Enabled = False
    ModoAtual = mfConsulta
End Sub


Private Sub PreencherCampos()

    If rsCliente.EOF Or rsCliente.BOF Then Exit Sub 'Se a lista não tem registros pula fora da Sub

    txtCodigo.Text = rsCliente!codigo 'Atribuição de valor do RecordSet para o TextBox
    txtNome.Text = rsCliente!Nome
    
    Dim TipoDocumentoBanco As Integer
    TipoDocumentoBanco = rsCliente!TipoDocumento
    Dim i As Integer
    For i = 0 To cboTipoDocumento.ListCount - 1 'Atribuindo o valor ao combo com base no valor do banco
        If cboTipoDocumento.ItemData(i) = TipoDocumentoBanco Then
            cboTipoDocumento.ListIndex = i
            Exit For
        End If
    Next
    
    mskDocumento.Text = rsCliente!Documento
    txtTelefone.Text = rsCliente!Telefone
    chkInativo.Value = IIf(rsCliente!Inativo = 1, vbChecked, vbUnchecked)

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    
'-------------NOVO
        Case "novo"
            txtCodigo.Text = ""
            txtNome.Text = ""
            cboTipoDocumento.ListIndex = tdcCPF
            'mskDocumento.Text = ""
            mskDocumento.SelStart = 0
            mskDocumento.SelLength = Len(mskDocumento.Text)
            mskDocumento.SelText = ""
            txtTelefone.Text = ""
            chkInativo.Value = vbUnchecked
            modoInclusao

'-------------SALVAR
        Case "salvar"

            If Not (ValidaCampos) Then Exit Sub
            
            If Not SalvarCliente Then
                MsgBox "Erro ao Salvar o Cliente!", vbOKOnly
                Exit Sub
            End If
            
            CarregarClientes
            
            If ModoAtual = mfAlteracao Then
                rsCliente.Find "Codigo = " & CodigoAtual
            Else
                If Not rsCliente.EOF Then rsCliente.MoveLast
            End If
            
            PreencherCampos
            modoConsulta

'-------------ALTERACAO
        Case "alterar"
        
            If rsCliente.EOF Or rsCliente.BOF Then Exit Sub
            
            PreencherCampos
            modoAlteracao

'-------------EXCLUIR
        Case "excluir"
            
            If rsCliente.EOF Or rsCliente.BOF Then Exit Sub
            
            If Not IsNumeric(txtCodigo.Text) Then
                MsgBox "Código Inválido", vbInformation
                Exit Sub
            End If
            
            'Mensagem de confirmação, se clicar no Não, cai fora da sub
            If MsgBox("Deseja realmente excluir este Cliente?", _
                      vbQuestion + vbYesNo, _
                      "Confirmação") = vbNo Then Exit Sub
            
            Dim codigoExcluir As Long
            codigoExcluir = CLng(txtCodigo.Text)
            
            If Not ExcluirCliente(codigoExcluir) Then
                MsgBox "Erro ao Excluir o Cliente!", vbInformation
                Exit Sub
            End If
            
            CarregarClientes
        
            If Not rsCliente.EOF Then
                rsCliente.Find "Codigo > " & codigoExcluir
                If rsCliente.EOF Then rsCliente.MoveLast
            End If
            
            PreencherCampos
            modoConsulta

'-------------DESFAZER
        Case "desfazer"
            modoConsulta
            PreencherCampos

'-------------PRIMEIRO
        Case "primeiro"
            rsCliente.MoveFirst
            PreencherCampos

'-------------ANTERIOR
        Case "anterior"
            If Not rsCliente.BOF Then rsCliente.MovePrevious
            If rsCliente.BOF Then rsCliente.MoveFirst
            PreencherCampos

'-------------PROXIMO
        Case "proximo"
            If Not rsCliente.EOF Then rsCliente.MoveNext
            If rsCliente.EOF Then rsCliente.MoveLast
            PreencherCampos

'-------------ULTIMO
        Case "ultimo"
            rsCliente.MoveLast
            PreencherCampos
            
'-------------VISUALIZAR
        Case "visualizar"
            Dim rpt As New arImpressaoClientes
            Dim Sql As String
            
            'Define a Conexão com o Banco
            rpt.dcImpPedido.ConnectionString = Conn
            
            Sql = "SELECT Codigo, Nome, " & _
            "Case TipoDocumento When 0 Then 'CPF' When 1 Then 'CNPJ' ELSE 'Outros' End as TipoDocumento, " & _
            "Documento, Telefone, " & _
            "Case Inativo When 0 Then 'Não' When 1 Then 'Sim' Else 'ERRO' End as Inativo " & _
            "FROM Cliente ORDER BY Codigo"
            
            'Define a string que vai ser executada no banco
            rpt.dcImpPedido.Source = Sql
            
            rpt.Run
            
            rpt.Show vbModal
        
    End Select

End Sub

'Função para salvar o Cliente
Private Function SalvarCliente() As Boolean
    On Error GoTo Erro
    
    Dim cliente As cCliente
    Set cliente = New cCliente
    
    If ModoAtual = mfAlteracao Then
        cliente.codigo = CLng(txtCodigo.Text)
        cliente.Nome = txtNome.Text
        cliente.TipoDocumento = cboTipoDocumento.ListIndex
        cliente.Documento = mskDocumento.Text
        cliente.Telefone = txtTelefone.Text
        cliente.Inativo = IIf(chkInativo.Value = vbChecked, 1, 0)
        If Not AlterarCliente(cliente) Then
            MsgBox "Erro ao Alterar o Cliente!", vbOKOnly
            Exit Function
        End If
    Else
        cliente.Nome = txtNome.Text
        cliente.TipoDocumento = cboTipoDocumento.ListIndex
        cliente.Documento = mskDocumento.Text
        cliente.Telefone = txtTelefone.Text
        cliente.Inativo = IIf(chkInativo.Value = vbChecked, 1, 0)
        If Not InserirCliente(cliente) Then
            MsgBox "Erro ao Inserir o Cliente!", vbOKOnly
            Exit Function
        End If
    End If
    
    SalvarCliente = True
    Exit Function
    
Erro:
    SalvarCliente = False

End Function

Private Sub Form_Unload(Cancel As Integer) 'No Unload do formulario fecha o recordset
    If Not rsCliente Is Nothing Then 'Se ele não for nada (se existir)
        If rsCliente.State = adStateOpen Then rsCliente.Close 'Se esta aberto, fecha
        Set rsCliente = Nothing 'Seta como nada
    End If
End Sub

'Clique do txtCodigo
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
            
            If BuscarRS(rsCliente, "Codigo", codigoBusca) Then
                PreencherCampos
            Else
                MsgBox "Não encontrado"
            End If
        End If
    End If
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        cmdListaCliente_Click
    End If
End Sub

'Clique do txtNome
Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'Verifica se é enter
        KeyAscii = 0 'Cancela o Enter, sem beep do windws
        cboTipoDocumento.SetFocus
    End If
    
End Sub

Private Sub cboTipoDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'Verifica se é enter
        KeyAscii = 0 'Cancela o Enter, sem beep do windws
        mskDocumento.SetFocus
    End If
End Sub

'Atribuir a máscara ao documento com base no Tipo Documento
Private Sub cboTipoDocumento_Click()

    ' Remove a máscara primeiro para não dar erro
    mskDocumento.Mask = ""
    mskDocumento.Text = ""

    ' Depois define a mascara novamente
    Select Case cboTipoDocumento.ListIndex
        
        Case tdcCPF
            mskDocumento.Mask = "999.999.999-99"

        Case tdcCNPJ
            mskDocumento.Mask = "99.999.999/9999-99"
            
        Case tdcOutro
            mskDocumento.Mask = ""
            
    End Select

End Sub

Private Sub mskDocumento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then 'Verifica se é enter
        KeyAscii = 0 'Cancela o Enter, sem beep do windws
        txtTelefone.SetFocus
    End If
    
    ' Somente numeros para CNPJ e CPF
    If cboTipoDocumento.ListIndex <> tdcOutro Then
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub txtTelefone_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then 'Verifica se a tecla digitada é enter
    
        KeyAscii = 0 'Limpa o Teclado
        
        If Not (ValidaCampos) Then Exit Sub
        
        If MsgBox("Confirma Dados?", _
            vbQuestion + vbYesNo, _
            "Confirmação") = vbNo Then 'Faz a pergunta, se não confirmar pula fora
            txtNome.SetFocus
            Exit Sub
        End If
        
        If Not SalvarCliente Then
            MsgBox "Erro ao Salvar o Cliente!", vbOKOnly
            Exit Sub
        End If
        
        CarregarClientes
        
        If ModoAtual = mfAlteracao Then
            rsCliente.Find "Codigo = " & CodigoAtual
        Else
            If Not rsCliente.EOF Then rsCliente.MoveLast
        End If
        
        PreencherCampos
        modoConsulta
    
    End If

End Sub

Private Sub cmdListaCliente_Click()
    Dim f As New frmPesquisaCliente

    f.Show vbModal

    If f.CodigoSelecionado > 0 Then
        If BuscarRS(rsCliente, "Codigo", f.CodigoSelecionado) Then
            PreencherCampos
        End If
    End If

    Unload f
End Sub

Private Sub popularComboTipoDocumento() 'Sub para popular o ComboBox
    cboTipoDocumento.Clear 'Limpa o Conteudo do Combo

    cboTipoDocumento.AddItem "CPF" 'Atribui a nomenclatura ao indice
    cboTipoDocumento.ItemData(cboTipoDocumento.NewIndex) = tdcCPF 'Atribui o ItemData(Valor Real do Enum no Banco) ao Indice
    cboTipoDocumento.AddItem "CNPJ"
    cboTipoDocumento.ItemData(cboTipoDocumento.NewIndex) = tdcCNPJ
    cboTipoDocumento.AddItem "Outro"
    cboTipoDocumento.ItemData(cboTipoDocumento.NewIndex) = tdcOutro
    
End Sub


'Função para validar campos
Private Function ValidaCampos()
    
    If ModoAtual = mfAlteracao Then
        If Not IsNumeric(txtCodigo.Text) Then
            txtNome.SetFocus
            ValidaCampos = False
            Exit Function
        End If
    End If
    
    If Trim(txtNome.Text = "") Then
        MsgBox "Nome Inválido"
        txtNome.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    If cboTipoDocumento.ListIndex = -1 Then
        MsgBox "Tipo Documento inválido"
        cboTipoDocumento.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    Select Case cboTipoDocumento.ListIndex
        
        Case tdcCPF
            If mskDocumento.Text = "" Or _
            mskDocumento.Text = "___.___.___-__" Or _
            Len(mskDocumento.Text) <> 14 Then
                MsgBox "Documento inválido"
                mskDocumento.SetFocus
                ValidaCampos = False
                Exit Function
            End If
        
        Case tdcCNPJ
            If mskDocumento.Text = "" Or _
            mskDocumento.Text = "__.___.___/____-__" Or _
            Len(mskDocumento.Text) <> 18 Then
                MsgBox "Documento inválido"
                mskDocumento.SetFocus
                ValidaCampos = False
                Exit Function
            End If
        
        Case tdcOutro
            If mskDocumento.Text = "" Then
                MsgBox "Documento inválido"
                mskDocumento.SetFocus
                ValidaCampos = False
                Exit Function
            End If
    
    End Select
    
    If Len(txtTelefone.Text) < 9 Or Len(txtTelefone.Text) > 15 Then
        MsgBox "Telefone inválido"
        txtTelefone.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    ValidaCampos = True
    
End Function
