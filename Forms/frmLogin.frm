VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   3090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Entrar"
      Height          =   375
      Left            =   990
      TabIndex        =   3
      Top             =   2370
      Width           =   1000
   End
   Begin VB.TextBox txtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   810
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   1425
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   810
      TabIndex        =   0
      Top             =   1110
      Width           =   1455
   End
   Begin VB.Label lblConfigBanco 
      Height          =   315
      Left            =   420
      TabIndex        =   6
      Top             =   330
      Width           =   375
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2595
   End
   Begin VB.Label lblSenha 
      Caption         =   "Senha:"
      Height          =   225
      Left            =   300
      TabIndex        =   4
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label lblCodigo 
      Caption         =   "Codigo:"
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   1170
      Width           =   555
   End
End
Attribute VB_Name = "frmLogin"
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

    txtCodigo.MaxLength = 5
    txtSenha.MaxLength = 15
End Sub
 
 Private Sub cmdLogin_Click()
 
    If Not ValidaCampos Then Exit Sub
 
    If Not AbrirConexao Then
        
        MsgBox "Erro ao conectar no banco de dados.", vbCritical
        frmConfigBanco.Show vbModal
        
        If Not AbrirConexao Then
            MsgBox "Conexão não configurada. O sistema será fechado."
            End
        End If
        
    End If
    
    If Not rsOperadorLogado Is Nothing Then 'Se ele não for nada (se existir)
        If rsOperadorLogado.State = adStateOpen Then rsOperadorLogado.Close 'Se esta aberto, fecha
        Set rsOperadorLogado = Nothing 'Seta como nada
    End If

    Set rsOperadorLogado = New ADODB.Recordset
    rsOperadorLogado.CursorLocation = adUseClient
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "SELECT * FROM Operador " & _
            "WHERE Codigo = ? " & _
            "AND Senha = ? " & _
            "AND Inativo = 0"
    End With
    
    cmd.Parameters.Append cmd.CreateParameter(, adBigInt, adParamInput, , CLng(txtCodigo.Text))
    cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 255, txtSenha.Text)
    
    Set rsOperadorLogado = cmd.Execute

    If Not rsOperadorLogado.EOF Then
        Unload Me
        Load MDIFrmPrincipal
        MDIFrmPrincipal.Show
    Else
        MsgBox "Usuário ou senha inválidos.", vbCritical
        txtCodigo.SetFocus
    End If

    'rsOperadorLogado.Close
    'Set rsOperadorLogado = Nothing
End Sub

Private Sub lblConfigBanco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        frmConfigBanco.Show vbModal
    End If
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then Exit Sub 'BackSpace
    
    If (KeyAscii = vbKeyReturn) Then 'Enter
        KeyAscii = 0
        txtSenha.SetFocus
    End If
    
    ' Só números
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If

End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        cmdLogin.SetFocus
    End If
    
End Sub

Private Function ValidaCampos() As Boolean
    
    If Not IsNumeric(txtCodigo.Text) Then
        MsgBox "Código Inválido!", vbInformation
        txtCodigo.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    If Trim(txtSenha.Text) = "" Then
        MsgBox "Senha Inválida!", vbInformation
        txtSenha.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    ValidaCampos = True
    
End Function
