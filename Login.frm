VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Entrar"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   1000
   End
   Begin VB.TextBox txtSenha 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
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
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   2115
   End
   Begin VB.Label lblSenha 
      Caption         =   "Senha:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblCodigo 
      Caption         =   "Codigo:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Conn As ADODB.Connection

Public Sub Conectar()
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=JEAN-PC;" & _
        "Initial Catalog=PROJETOVB;" & _
        "User ID=sa;" & _
        "Password=sae;"
    Conn.Open
End Sub


Private Sub cmdLogin_Click()
    Dim rs As ADODB.Recordset
    Dim sql As String

    If Trim(txtUsuario.Text) = "" Or Trim(txtSenha.Text) = "" Then
        MsgBox "Informe usuário e senha.", vbExclamation
        Exit Sub
    End If

    sql = "SELECT * FROM Operador " & _
          "WHERE Codigo = '" & txtCodigo.Text & "' " & _
          "AND Senha = '" & txtSenha.Text & "' " & _
          "AND Inativo = 0"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        MsgBox "Login realizado com sucesso!", vbInformation
        Unload Me
        frmPrincipal.Show
    Else
        MsgBox "Usuário ou senha inválidos.", vbCritical
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Conectar
End Sub

