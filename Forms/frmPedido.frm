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
      Height          =   390
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   1065
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
      Height          =   390
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   1065
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
      Height          =   390
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   1065
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
      Left            =   1710
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grdItensPedido 
      Height          =   2805
      Left            =   420
      TabIndex        =   9
      Top             =   2700
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   4948
      _Version        =   393216
   End
   Begin VB.CommandButton cmdListaCliente 
      DisabledPicture =   "frmPedido.frx":0000
      DownPicture     =   "frmPedido.frx":05E2
      Height          =   375
      Left            =   2610
      Picture         =   "frmPedido.frx":0BC4
      Style           =   1  'Graphical
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   1380
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker dtpDataPedido 
      Height          =   375
      Left            =   5550
      TabIndex        =   6
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
      Format          =   149684225
      CurrentDate     =   46036
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
      TabIndex        =   2
      Top             =   900
      Width           =   1005
   End
   Begin VB.CommandButton cmdListaPedido 
      DisabledPicture =   "frmPedido.frx":11A6
      DownPicture     =   "frmPedido.frx":1788
      Height          =   375
      Left            =   2610
      Picture         =   "frmPedido.frx":1D6A
      Style           =   1  'Graphical
      TabIndex        =   1
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
      TabIndex        =   5
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
      MouseIcon       =   "frmPedido.frx":234C
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
               Picture         =   "frmPedido.frx":3026
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":3D00
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":49DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":56B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":638E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":7068
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":7644
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":831E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPedido.frx":8FF8
               Key             =   ""
            EndProperty
         EndProperty
      End
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
      Left            =   420
      TabIndex        =   12
      Top             =   5640
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   1380
      Width           =   1065
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ModoAtual As eModoFormulario
Dim Pedido As cPedido

Private Sub Form_Load()
    
    Set Pedido = New cPedido
    
    AjustarColunasGridItens
    CarregarPedidos

    If Not rsPedido.EOF Then 'Se não esta no fim da lista
        rsPedido.MoveLast 'Move para o final
        PreencherCampos
    End If

    modoConsulta
    
End Sub 'PAREI AQUI------------------------------------------

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
    
    cmdNovoItem.Enabled = False
    cmdAlterarItem.Enabled = False
    cmdExcluirItem.Enabled = False
    
    txtCodigo.Enabled = False 'Habilitar/Desabilitar txt
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaPedido.Enabled = False 'Habilitar/Desabilitar commandButton
    txtCodCliente.Enabled = True
    txtCodCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    cmdListaCliente.Enabled = True 'Habilitar/Desabilitar commandButton
    txtNomeCliente.Enabled = True
    txtNomeCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    dtpDataPedido.Enabled = True
    
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
    
    cmdNovoItem.Enabled = False
    cmdAlterarItem.Enabled = False
    cmdExcluirItem.Enabled = False
    
    txtCodigo.Enabled = False 'Habilitar/Desabilitar txt
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaPedido.Enabled = False 'Habilitar/Desabilitar commandButton
    txtCodCliente.Enabled = True
    txtCodCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    cmdListaCliente.Enabled = True 'Habilitar/Desabilitar commandButton
    txtNomeCliente.Enabled = True
    txtNomeCliente.BackColor = vbWindowBackground 'cor branca padrão do sistema
    dtpDataPedido.Enabled = True
    
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
    
    cmdNovoItem.Enabled = True
    cmdAlterarItem.Enabled = True
    cmdExcluirItem.Enabled = True
    
    txtCodigo.Enabled = True
    txtCodigo.BackColor = vbWindowBackground
    cmdListaPedido.Enabled = True
    txtCodCliente.Enabled = False
    txtCodCliente.BackColor = &H8000000F
    cmdListaCliente.Enabled = False
    txtNomeCliente.Enabled = False
    txtNomeCliente.BackColor = &H8000000F
    dtpDataPedido.Enabled = False
    
    ModoAtual = mfConsulta
End Sub


Private Sub PreencherCampos()

    If rsPedido.EOF Or rsPedido.BOF Then Exit Sub 'Se a lista não tem registros pula fora da Sub

    txtCodigo.Text = rsPedido!Codigo 'Atribuição de valor do RecordSet para o TextBox
    txtCodCliente.Text = rsPedido!CodigoCliente
    txtNomeCliente.Text = rsPedido!ClienteNome
    dtpDataPedido.Value = rsPedido!DataPedido
    txtValorTotal.Text = rsOperador!ValorTotal
    
    PreencherItensPedido
    
End Sub

Private Sub PreencherItensPedido()
    If rsPedidoItem.EOF Or rsPedidoItem.BOF Then Exit Sub
    
    Dim linha As Long
    
    If Not IsNumeric(txtCodigo.Text) Then
        MsgBox "Código inválido.", vbExclamation
        txtCodigo.SetFocus
        Exit Sub
    End If
    
    CarregarItensPedido (CLng(txtCodigo.Text))
    
    linha = 1

    Do While Not rs.EOF
        grdItens.Rows = grdItens.Rows + 1

        grdItens.TextMatrix(linha, 0) = rs!Codigo
        grdItens.TextMatrix(linha, 1) = rs!Produto
        grdItens.TextMatrix(linha, 2) = rs!Quantidade
        grdItens.TextMatrix(linha, 3) = Format(rs!ValorUn, "0.00")
        grdItens.TextMatrix(linha, 4) = Format(rs!ValorTotal, "0.00")

        linha = linha + 1
        rs.MoveNext
    Loop
    
    If grdItensPedido.Rows > 1 Then
        grdItensPedido.Row = 1
    End If

End Sub

Private Sub AjustarColunasGridItens()
    With grdItensPedido
        .Rows = 1
        .Cols = 5

        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Produto"
        .TextMatrix(0, 2) = "Qtde"
        .TextMatrix(0, 3) = "Vlr Unit"
        .TextMatrix(0, 4) = "Total"

        .ColWidth(0) = 800
        .ColWidth(1) = 3000
        .ColWidth(2) = 800
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
'-------------NOVO
        Case "novo"
            
            txtCodigo.Text = ""
            txtCodCliente.Text = ""
            txtNomeCliente.Text = ""
            dtpDataPedido.Value = Date
            txtValorTotal.Text = ""
            
            modoInclusao

'-------------SALVAR
        Case "salvar"
            Dim sql As String
            Dim codigoAtual As Long

            If ModoAtual = mfAlteracao Then
                Pedido.Codigo = CLng(txtCodigo.Text) 'Conversão de Texto para Long
                Pedido.ClienteCodigo = CLng(txtCodCliente.Text)
                Pedido.DataPedido = dtpDataPedido.Value
                'Parei AQUI -------------------------------------------------------------
                'Criar uma sub no recordset para salvar o pedido e mandar o objeto pedido para ele
            Else
                sql = "INSERT INTO Operador (Nome, Senha, Admin, Inativo) VALUES (" & _
                    "'" & txtNome.Text & "', " & _
                    "'" & txtSenha.Text & "', " & _
                    IIf(chkAdm.Value = vbChecked, 1, 0) & ", " & _
                    IIf(chkInativo.Value = vbChecked, 1, 0) & ")"
            End If

            Conn.Execute sql
            
            CarregarOperadores
            
            If ModoAtual = mfAlteracao Then
                rsOperador.Find "Codigo = " & codigoAtual
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
        
            If Not rsCliente.EOF Then
                rsCliente.Find "Codigo > " & codigoExcluir
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
    If Not rsPedido Is Nothing Then 'Se ele não for nada (se existir)
        If rsPedido.State = adStateOpen Then rsPedido.Close 'Se esta aberto, fecha
        Set rsPedido = Nothing 'Seta como nada
    End If
    If Not rsPedidoItem Is Nothing Then 'Se ele não for nada (se existir)
        If rsPedidoItem.State = adStateOpen Then rsPedidoItem.Close 'Se esta aberto, fecha
        Set rsPedidoItem = Nothing 'Seta como nada
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

