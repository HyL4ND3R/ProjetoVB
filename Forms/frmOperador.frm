VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmOperador 
   Caption         =   "Cadastro de Operadores"
   ClientHeight    =   7965
   ClientLeft      =   285
   ClientTop       =   630
   ClientWidth     =   18075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   18075
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   17550
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
            Picture         =   "frmOperador.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador.frx":19B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador.frx":268E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador.frx":3368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador.frx":4042
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador.frx":461E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador.frx":52F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador.frx":5FD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   18075
      _ExtentX        =   31882
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
      MouseIcon       =   "frmOperador.frx":65A4
   End
   Begin VB.TextBox txtSenha 
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
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1770
      Width           =   1965
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1260
      Width           =   4005
   End
   Begin VB.CommandButton cmdListaOperador 
      DisabledPicture =   "frmOperador.frx":727E
      DownPicture     =   "frmOperador.frx":7860
      Height          =   375
      Left            =   2730
      Picture         =   "frmOperador.frx":7E42
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   750
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
      Left            =   1680
      TabIndex        =   1
      Top             =   750
      Width           =   1005
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
      Left            =   1680
      TabIndex        =   5
      Top             =   2790
      Width           =   1335
   End
   Begin VB.CheckBox chkAdm 
      Caption         =   "Administrador"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1875
   End
   Begin VB.Label lblSenha 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
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
      Left            =   570
      TabIndex        =   7
      Top             =   1770
      Width           =   1065
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
      Left            =   570
      TabIndex        =   6
      Top             =   1260
      Width           =   1065
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
      Left            =   570
      TabIndex        =   0
      Top             =   750
      Width           =   1065
   End
End
Attribute VB_Name = "frmOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ModoAtual As eModoFormulario
Private CodigoAtual As Long

Private Sub Form_Load()
    
    CarregarOperadores

    If Not rsOperador.EOF Then 'Se não esta no fim da lista
        rsOperador.MoveLast 'Move para o final
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
    cmdListaOperador.Enabled = False 'Habilitar/Desabilitar commandButton
    txtNome.Enabled = True
    txtNome.BackColor = vbWindowBackground 'cor branca padrão do sistema
    txtSenha.Enabled = True
    txtSenha.BackColor = vbWindowBackground 'cor branca padrão do sistema
    chkAdm.Enabled = True
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
    cmdListaOperador.Enabled = False 'Habilitar/Desabilitar commandButton
    txtNome.Enabled = True
    txtNome.BackColor = vbWindowBackground 'cor branca padrão do sistema
    txtSenha.Enabled = True
    txtSenha.BackColor = vbWindowBackground 'cor branca padrão do sistema
    chkAdm.Enabled = True
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
    txtCodigo.BackColor = vbWindowBackground
    cmdListaOperador.Enabled = True
    txtNome.Enabled = False
    txtNome.BackColor = &H8000000F
    txtSenha.Enabled = False
    txtSenha.BackColor = &H8000000F
    chkAdm.Enabled = False
    chkInativo.Enabled = False
    ModoAtual = mfConsulta
End Sub


Private Sub PreencherCampos()

    If rsOperador.EOF Or rsOperador.BOF Then Exit Sub 'Se a lista não tem registros pula fora da Sub

    txtCodigo.Text = rsOperador!Codigo 'Atribuição de valor do RecordSet para o TextBox
    txtNome.Text = rsOperador!Nome
    txtSenha.Text = rsOperador!Senha
    chkAdm.Value = IIf(rsOperador!Admin = 1, vbChecked, vbUnchecked) 'Atribuição de valor do Recorset para o CheckBox
    chkInativo.Value = IIf(rsOperador!Inativo = 1, vbChecked, vbUnchecked)

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
'-------------NOVO
        Case "novo"
            txtCodigo.Text = ""
            txtNome.Text = ""
            txtSenha.Text = ""
            chkAdm.Value = vbUnchecked 'Atribuição de Marcado/Desmarcado
            chkInativo.Value = vbUnchecked
            modoInclusao

'-------------SALVAR
        Case "salvar"
            
            If Not ValidaCampos Then Exit Sub
            
            If Not SalvarOperador Then
                MsgBox "Erro ao Salvar o Operador!", vbOKOnly
                Exit Sub
            End If
            
            CarregarOperadores
            
            If ModoAtual = mfAlteracao Then
                rsOperador.Find "Codigo = " & CodigoAtual
            Else
                If Not rsOperador.EOF Then rsOperador.MoveLast
            End If
            
            PreencherCampos
            modoConsulta

'-------------ALTERACAO
        Case "alterar"
        
            If rsOperador.EOF Or rsOperador.BOF Then Exit Sub
            
            PreencherCampos
            modoAlteracao

'-------------EXCLUIR
        Case "excluir"
            
            If rsOperador.EOF Or rsOperador.BOF Then Exit Sub
            
            'Mensagem de confirmação, se clicar no Não, cai fora da sub
            If MsgBox("Deseja realmente excluir este operador?", _
                      vbQuestion + vbYesNo, _
                      "Confirmação") = vbNo Then Exit Sub

            Dim codigoExcluir As Long
            codigoExcluir = CLng(txtCodigo.Text)
        
            Conn.Execute "DELETE FROM Operador WHERE Codigo = " & codigoExcluir
        
            CarregarOperadores
        
            If Not rsOperador.EOF Then
                rsOperador.Find "Codigo > " & codigoExcluir
                If rsOperador.EOF Then rsOperador.MoveLast
            End If
        
            PreencherCampos
            modoConsulta

'-------------DESFAZER
        Case "desfazer"
            modoConsulta
            PreencherCampos

'-------------PRIMEIRO
        Case "primeiro"
            rsOperador.MoveFirst
            PreencherCampos

'-------------ANTERIOR
        Case "anterior"
            If Not rsOperador.BOF Then rsOperador.MovePrevious
            If rsOperador.BOF Then rsOperador.MoveFirst
            PreencherCampos

'-------------PROXIMO
        Case "proximo"
            If Not rsOperador.EOF Then rsOperador.MoveNext
            If rsOperador.EOF Then rsOperador.MoveLast
            PreencherCampos

'-------------ULTIMO
        Case "ultimo"
            rsOperador.MoveLast
            PreencherCampos
        
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer) 'No Unload do formulario fecha o recordset
    If Not rsOperador Is Nothing Then 'Se ele não for nada (se existir)
        If rsOperador.State = adStateOpen Then rsOperador.Close 'Se esta aberto, fecha
        Set rsOperador = Nothing 'Seta como nada
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
            
            If BuscarRS(rsOperador, "Codigo", codigoBusca) Then
                PreencherCampos
            Else
                MsgBox "Não encontrado"
            End If
        End If
    End If
    
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'Verifica se é enter
        KeyAscii = 0 'Cancela o Enter, sem beep do windws
        txtSenha.SetFocus
    End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'Verifica se a tecla digitada é enter
    
        KeyAscii = 0 'Limpa o Teclado
        
        If Not (ValidaCampos) Then Exit Sub
        
        If MsgBox("Confirma Dados?", _
            vbQuestion + vbYesNo, _
            "Confirmação") = vbNo Then 'Faz a pergunta, se não confirmar pula fora
            txtNome.SetFocus
            Exit Sub
        End If
        
        If Not SalvarOperador Then
            MsgBox "Erro ao Salvar o Operador!", vbOKOnly
            Exit Sub
        End If
        
        CarregarOperadores
        
        If ModoAtual = mfAlteracao Then
            rsOperador.Find "Codigo = " & CodigoAtual
        Else
            If Not rsOperador.EOF Then rsOperador.MoveLast
        End If
        
        PreencherCampos
        modoConsulta
    
    End If
End Sub

Private Sub cmdListaOperador_Click()
    Dim f As New frmPesquisaOperador

    f.Show vbModal

    If f.CodigoSelecionado > 0 Then
        If BuscarRS(rsOperador, "Codigo", f.CodigoSelecionado) Then
            PreencherCampos
        End If
    End If

    Unload f
End Sub

Private Function SalvarOperador() As Boolean
    On Error GoTo Erro
    
    Dim Sql As String
    
    If ModoAtual = mfAlteracao Then
        CodigoAtual = CLng(txtCodigo.Text) 'Conversão de Texto para Long
        Sql = "UPDATE Operador set Nome = " & "'" & txtNome.Text & "', " & _
            "Senha = " & "'" & txtSenha.Text & "', " & _
            "Admin = " & IIf(chkAdm.Value = vbChecked, 1, 0) & ", " & _
            "Inativo = " & IIf(chkInativo.Value = vbChecked, 1, 0) & " " & _
            "WHERE Codigo = " & txtCodigo.Text
    Else
        Sql = "INSERT INTO Operador (Nome, Senha, Admin, Inativo) VALUES (" & _
            "'" & txtNome.Text & "', " & _
            "'" & txtSenha.Text & "', " & _
            IIf(chkAdm.Value = vbChecked, 1, 0) & ", " & _
            IIf(chkInativo.Value = vbChecked, 1, 0) & ")"
    End If

    Conn.Execute Sql
    SalvarOperador = True
    Exit Function
    
Erro:
    SalvarOperador = False
    
End Function


Private Function ValidaCampos() As Boolean
    
    If Trim(txtNome.Text = "") Then
        MsgBox "Nome Inválido"
        txtNome.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    If Trim(txtSenha.Text = "") Then
        MsgBox "Senha Inválida"
        txtNome.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    ValidaCampos = True
    
End Function
