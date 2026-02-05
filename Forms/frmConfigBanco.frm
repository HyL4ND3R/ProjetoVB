VERSION 5.00
Begin VB.Form frmConfigBanco 
   Caption         =   "Config Banco"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTestar 
      Caption         =   "Testar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2745
      TabIndex        =   10
      Top             =   3060
      Width           =   1185
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1395
      TabIndex        =   9
      Top             =   3060
      Width           =   1185
   End
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1395
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1395
      TabIndex        =   6
      Top             =   2025
      Width           =   2535
   End
   Begin VB.TextBox txtBanco 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1395
      TabIndex        =   5
      Top             =   1530
      Width           =   2535
   End
   Begin VB.TextBox txtServidor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1395
      TabIndex        =   4
      Top             =   1035
      Width           =   2535
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Config Banco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      TabIndex        =   8
      Top             =   225
      Width           =   2490
   End
   Begin VB.Label lblSenha 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label lblUsuario 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   2
      Top             =   2025
      Width           =   915
   End
   Begin VB.Label lblBanco 
      Alignment       =   1  'Right Justify
      Caption         =   "Banco:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   1
      Top             =   1530
      Width           =   915
   End
   Begin VB.Label lblServidor 
      Alignment       =   1  'Right Justify
      Caption         =   "Servidor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   1035
      Width           =   915
   End
End
Attribute VB_Name = "frmConfigBanco"
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

    txtBanco.MaxLength = 15
    txtSenha.MaxLength = 30
    txtServidor.MaxLength = 50
    txtUsuario.MaxLength = 10

    Dim arqINI As String

    arqINI = App.Path & "\config.ini"

    ' Se o arquivo não existir, sai sem erro
    If Dir(arqINI) = "" Then
        MsgBox "Arquivo .ini não encontrado, criando um novo", vbExclamation
        Exit Sub
    End If
    
    txtServidor.Text = LerINI("BANCO", "Servidor", arqINI)
    txtBanco.Text = LerINI("BANCO", "Banco", arqINI)
    txtUsuario.Text = LerINI("BANCO", "Usuario", arqINI)
    txtSenha.Text = LerINI("BANCO", "Senha", arqINI)

End Sub

Private Sub cmdTestar_Click()
    Dim cn As New ADODB.Connection
    On Error GoTo Erro

    'Definindo TimeOut menor para caso de erro
    cn.ConnectionTimeout = 2 'segundos (ex: 3, 5, 10)

    cn.ConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=" & txtServidor.Text & ";" & _
        "Initial Catalog=" & txtBanco.Text & ";" & _
        "User ID=" & txtUsuario.Text & ";" & _
        "Password=" & txtSenha.Text & ";"

    cn.Open
    MsgBox "Conexão realizada com sucesso!", vbInformation
    cn.Close
    Exit Sub

Erro:
    MsgBox "Falha ao conectar: " & Err.Description, vbCritical
End Sub

Private Sub cmdSalvar_Click()
    Dim arqINI As String
    arqINI = App.Path & "\config.ini"

    GravarINI "BANCO", "Servidor", txtServidor.Text, arqINI
    GravarINI "BANCO", "Banco", txtBanco.Text, arqINI
    GravarINI "BANCO", "Usuario", txtUsuario.Text, arqINI
    GravarINI "BANCO", "Senha", txtSenha.Text, arqINI

    MsgBox "Configurações salvas!", vbInformation
    Unload Me
End Sub

Private Sub txtServidor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtBanco.SetFocus
    End If
End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtUsuario.SetFocus
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtSenha.SetFocus
    End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdSalvar.SetFocus
    End If
End Sub
