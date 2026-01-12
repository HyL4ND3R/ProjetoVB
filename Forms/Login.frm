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
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Entrar"
      Height          =   375
      Left            =   990
      TabIndex        =   4
      Top             =   2370
      Width           =   1000
   End
   Begin VB.TextBox txtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   810
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "123"
      Top             =   1680
      Width           =   1425
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   810
      TabIndex        =   0
      Text            =   "1"
      Top             =   1110
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
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2595
   End
   Begin VB.Label lblSenha 
      Caption         =   "Senha:"
      Height          =   225
      Left            =   300
      TabIndex        =   3
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label lblCodigo 
      Caption         =   "Codigo:"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   1170
      Width           =   555
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

    If Trim(txtCodigo.Text) = "" Or Trim(txtSenha.Text) = "" Then
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
        Unload Me
        Load MDIFrmPrincipal
        MDIFrmPrincipal.Show
    Else
        MsgBox "Usuário ou senha inválidos.", vbCritical
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Conectar
End Sub

