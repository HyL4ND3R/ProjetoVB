VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   7
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
      TabIndex        =   6
      Text            =   "Combo1"
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
      TabIndex        =   4
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
   Begin MSComctlLib.Toolbar Toolbar1 
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
            NumListImages   =   9
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
      TabIndex        =   8
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
      TabIndex        =   5
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
      TabIndex        =   3
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
      TabIndex        =   2
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

Private Sub Form_Load()
    
    popularComboTipoDocumento
    CarregarClientes

    If Not rsCliente.EOF Then 'Se não esta no fim da lista
        rsCliente.MoveLast 'Move para o final
        PreencherCampos
    End If

    modoConsulta
    
End Sub

Private Sub modoInclusao()
    Toolbar1.Buttons("novo").Enabled = False 'Habilitar/Desabilitar botão da toolbar
    Toolbar1.Buttons("salvar").Enabled = True
    Toolbar1.Buttons("alterar").Enabled = False
    Toolbar1.Buttons("excluir").Enabled = False
    Toolbar1.Buttons("desfazer").Enabled = True
    Toolbar1.Buttons("primeiro").Enabled = False
    Toolbar1.Buttons("anterior").Enabled = False
    Toolbar1.Buttons("proximo").Enabled = False
    Toolbar1.Buttons("ultimo").Enabled = False
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
    Toolbar1.Buttons("novo").Enabled = False 'Habilitar/Desabilitar botão da toolbar
    Toolbar1.Buttons("salvar").Enabled = True
    Toolbar1.Buttons("alterar").Enabled = False
    Toolbar1.Buttons("excluir").Enabled = False
    Toolbar1.Buttons("desfazer").Enabled = True
    Toolbar1.Buttons("primeiro").Enabled = False
    Toolbar1.Buttons("anterior").Enabled = False
    Toolbar1.Buttons("proximo").Enabled = False
    Toolbar1.Buttons("ultimo").Enabled = False
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
    Toolbar1.Buttons("novo").Enabled = True
    Toolbar1.Buttons("salvar").Enabled = False
    Toolbar1.Buttons("excluir").Enabled = True
    Toolbar1.Buttons("alterar").Enabled = True
    Toolbar1.Buttons("desfazer").Enabled = False
    Toolbar1.Buttons("primeiro").Enabled = True
    Toolbar1.Buttons("anterior").Enabled = True
    Toolbar1.Buttons("proximo").Enabled = True
    Toolbar1.Buttons("ultimo").Enabled = True
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

    txtCodigo.Text = rsCliente!Codigo 'Atribuição de valor do RecordSet para o TextBox
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

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
            Dim sql As String
            Dim codigoAtual As Long

            If ModoAtual = mfAlteracao Then
                codigoAtual = CLng(txtCodigo.Text) 'Conversão de Texto para Long
                sql = "UPDATE Cliente set Nome = " & "'" & txtNome.Text & "', " & _
                    "TipoDocumento = " & "'" & cboTipoDocumento.ItemData(cboTipoDocumento.ListIndex) & "', " & _
                    "Documento = '" & mskDocumento.Text & "', " & _
                    "Telefone = '" & txtTelefone.Text & "', " & _
                    "Inativo = " & IIf(chkInativo.Value = vbChecked, 1, 0) & _
                    "WHERE Codigo = " & txtCodigo.Text
            Else
                sql = "INSERT INTO Cliente (Nome, TipoDocumento, Documento, Telefone, Inativo) VALUES (" & _
                    "'" & txtNome.Text & "', " & _
                    "" & cboTipoDocumento.ItemData(cboTipoDocumento.ListIndex) & ", " & _
                    "'" & mskDocumento.Text & "', " & _
                    "'" & txtTelefone.Text & "', " & _
                    IIf(chkInativo.Value = vbChecked, 1, 0) & ")"
            End If

            Conn.Execute sql
            
            CarregarClientes
            
            If ModoAtual = mfAlteracao Then
                rsCliente.Find "Codigo = " & codigoAtual
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
            
            'Mensagem de confirmação, se clicar no Não, cai fora da sub
            If MsgBox("Deseja realmente excluir este Cliente?", _
                      vbQuestion + vbYesNo, _
                      "Confirmação") = vbNo Then Exit Sub

            Dim codigoExcluir As Long
            codigoExcluir = CLng(txtCodigo.Text)
        
            Conn.Execute "DELETE FROM Cliente WHERE Codigo = " & codigoExcluir
        
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
        
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer) 'No Unload do formulario fecha o recordset
    If Not rsCliente Is Nothing Then 'Se ele não for nada (se existir)
        If rsCliente.State = adStateOpen Then rsCliente.Close 'Se esta aberto, fecha
        Set rsCliente = Nothing 'Seta como nada
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
            
            If BuscarRS(rsCliente, "Codigo", codigoBusca) Then
                PreencherCampos
            Else
                MsgBox "Não encontrado"
            End If
        End If
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

Private Sub cboTipoDocumento_Click()

    ' Remove a máscara primeiro para não dar erro
    mskDocumento.Mask = ""
    mskDocumento.Text = ""

    ' Depois define a mascara novamente
    Select Case cboTipoDocumento.ListIndex
        
        Case tdcCPF
            mskDocumento.Mask = "###.###.###-##"

        Case tdcCNPJ
            mskDocumento.Mask = "##.###.###/####-##"
            
        Case tdcOutro
            mskDocumento.Mask = ""
            
    End Select

End Sub
