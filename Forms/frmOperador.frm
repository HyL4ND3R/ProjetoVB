VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmOperador 
   Caption         =   "Cadastro de Operadores"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   18120
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
      Height          =   630
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   18120
      _ExtentX        =   31962
      _ExtentY        =   1111
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
      Left            =   1680
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   1260
      Width           =   4005
   End
   Begin VB.CommandButton cmdListaOperador 
      Caption         =   "..."
      Height          =   375
      Left            =   2730
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   4
      Top             =   1770
      Width           =   1065
   End
   Begin VB.Label frmNome 
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
      TabIndex        =   3
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

Private Sub Form_Load()
    
    CarregarOperadores

    If Not RsOperador.EOF Then
        RsOperador.MoveLast
        PreencherCampos
    End If

    ModoConsulta
    
End Sub

Private Sub ModoAlteracao()
    Toolbar1.Buttons("novo").Enabled = False
    Toolbar1.Buttons("salvar").Enabled = True
    Toolbar1.Buttons("alterar").Enabled = False
    Toolbar1.Buttons("excluir").Enabled = False
    Toolbar1.Buttons("desfazer").Enabled = True
    Toolbar1.Buttons("primeiro").Enabled = False
    Toolbar1.Buttons("anterior").Enabled = False
    Toolbar1.Buttons("proximo").Enabled = False
    Toolbar1.Buttons("ultimo").Enabled = False
    txtCodigo.Enabled = False
    txtCodigo.BackColor = &H8000000F 'cor cinza padrão do sistema
    cmdListaOperador.Enabled = False
    txtNome.Enabled = True
    txtNome.BackColor = vbWindowBackground
    txtSenha.Enabled = True
    txtSenha.BackColor = vbWindowBackground
    chkAdm.Enabled = True
    chkInativo.Enabled = True
End Sub

Private Sub ModoConsulta()
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
    txtNome.BackColor = &H8000000F   ' cor cinza padrão do sistema
    txtSenha.Enabled = False
    txtSenha.BackColor = &H8000000F   ' cor cinza padrão do sistema
    chkAdm.Enabled = False
    chkInativo.Enabled = False
End Sub

Public Sub CarregarOperadores()

    Set RsOperador = New ADODB.Recordset

    RsOperador.CursorLocation = adUseClient
    RsOperador.Open _
        "SELECT Codigo, Nome, Senha, Admin, Inativo FROM Operador ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Private Sub PreencherCampos()

    If RsOperador.EOF Or RsOperador.BOF Then Exit Sub

    txtCodigo.Text = RsOperador!Codigo
    txtNome.Text = RsOperador!Nome
    txtSenha.Text = RsOperador!Senha
    chkAdm.Value = IIf(RsOperador!Admin = 1, vbChecked, vbUnchecked)
    chkInativo.Value = IIf(RsOperador!Inativo = 1, vbChecked, vbUnchecked)

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "novo"
            txtCodigo.Text = ""
            txtNome.Text = ""
            txtSenha.Text = ""
            chkAdm.Value = vbUnchecked
            chkInativo.Value = vbUnchecked
            ModoAlteracao

        Case "salvar"
            Dim sql As String
            Dim codigoAtual As Long
            Dim alteracao As Boolean
            alteracao = (Trim(txtCodigo.Text) <> "")

            If alteracao Then
                codigoAtual = CLng(txtCodigo.Text)
                sql = "UPDATE Operador set Nome = " & "'" & txtNome.Text & "', " & _
                    "Senha = " & "'" & txtSenha.Text & "', " & _
                    "Admin = " & IIf(chkAdm.Value = vbChecked, 1, 0) & ", " & _
                    "Inativo = " & IIf(chkInativo.Value = vbChecked, 1, 0) & _
                    "WHERE Codigo = " & txtCodigo.Text
            Else
                sql = "INSERT INTO Operador (Nome, Senha, Admin, Inativo) VALUES (" & _
                    "'" & txtNome.Text & "', " & _
                    "'" & txtSenha.Text & "', " & _
                    IIf(chkAdm.Value = vbChecked, 1, 0) & ", " & _
                    IIf(chkInativo.Value = vbChecked, 1, 0) & ")"
            End If

            Conn.Execute sql
            
            CarregarOperadores
            
            If alteracao Then
                RsOperador.Find "Codigo = " & codigoAtual
            Else
                If Not RsOperador.EOF Then RsOperador.MoveFirst
            End If
            
            PreencherCampos
            ModoConsulta

        Case "alterar"
        
            If RsOperador.EOF Or RsOperador.BOF Then Exit Sub
            
            PreencherCampos
            ModoAlteracao

        Case "excluir"
            
            If RsOperador.EOF Or RsOperador.BOF Then Exit Sub

            If MsgBox("Deseja realmente excluir este operador?", _
                      vbQuestion + vbYesNo, _
                      "Confirmação") = vbNo Then Exit Sub
        
            Dim codigoExcluir As Long
            codigoExcluir = CLng(txtCodigo.Text)
        
            Conn.Execute "DELETE FROM Operador WHERE Codigo = " & codigoExcluir
        
            CarregarOperadores
        
            ' Reposicionar após excluir
            If Not RsOperador.EOF Then
                RsOperador.Find "Codigo > " & codigoExcluir
                If RsOperador.EOF Then RsOperador.MoveLast
            End If
        
            PreencherCampos
            ModoConsulta

        Case "desfazer"
            ModoConsulta
            PreencherCampos

        Case "primeiro"
            RsOperador.MoveFirst
            PreencherCampos
        
        Case "anterior"
            If Not RsOperador.BOF Then RsOperador.MovePrevious
            If RsOperador.BOF Then RsOperador.MoveFirst
            PreencherCampos
        
        Case "proximo"
            If Not RsOperador.EOF Then RsOperador.MoveNext
            If RsOperador.EOF Then RsOperador.MoveLast
            PreencherCampos
        
        Case "ultimo"
            RsOperador.MoveLast
            PreencherCampos
        
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not RsOperador Is Nothing Then
        If RsOperador.State = adStateOpen Then RsOperador.Close
        Set RsOperador = Nothing
    End If
End Sub
