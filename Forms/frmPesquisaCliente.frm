VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPesquisaCliente 
   Caption         =   "Pesquisa de Clientes"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Height          =   465
      Left            =   6240
      TabIndex        =   2
      Top             =   5100
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   7560
      TabIndex        =   1
      Top             =   5100
      Width           =   1305
   End
   Begin MSFlexGridLib.MSFlexGrid grdCliente 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPesquisaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoSelecionado As Long

Private Sub Form_Load()
    CarregarClientes

    With grdCliente
        .Rows = 1
        .Cols = 4
        .TextMatrix(0, 0) = "Código"
        .ColWidth(0) = 1500
        .TextMatrix(0, 1) = "Nome"
        .ColWidth(1) = 3000
        .TextMatrix(0, 2) = "Documento"
        .ColWidth(2) = 2000
        .TextMatrix(0, 3) = "Telefone"
        .ColWidth(3) = 2000
        
        Do While Not rsCliente.EOF
            .AddItem _
                rsCliente!Codigo & vbTab & _
                rsCliente!Nome & vbTab & _
                rsCliente!Documento & vbTab & _
                rsCliente!Telefone
            rsCliente.MoveNext
        Loop
    End With
End Sub

Private Sub grdCliente_DblClick()
    Selecionar
End Sub

Private Sub cmdSelecionar_Click()
    Selecionar
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Selecionar()
    If grdCliente.Row = 0 Then Exit Sub

    CodigoSelecionado = CLng(grdCliente.TextMatrix(grdCliente.Row, 0))
    Unload Me
End Sub


