VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPesquisaProduto 
   Caption         =   "Pesquisar Produtos"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Height          =   465
      Left            =   6030
      TabIndex        =   2
      Top             =   4740
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   7350
      TabIndex        =   1
      Top             =   4740
      Width           =   1305
   End
   Begin MSFlexGridLib.MSFlexGrid grdProduto 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   8281
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
Attribute VB_Name = "frmPesquisaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoSelecionado As Long


Private Sub Form_Load()
    CarregarProdutos

    With grdProduto
        .Rows = 1
        .Cols = 4
        .TextMatrix(0, 0) = "Código"
        .ColWidth(0) = 1500
        .TextMatrix(0, 1) = "Nome"
        .ColWidth(1) = 3000
        .TextMatrix(0, 2) = "Valor"
        .ColWidth(2) = 2000
        .TextMatrix(0, 3) = "Inativo"
        .ColWidth(3) = 2000
        
        Do While Not rsProduto.EOF
            .AddItem _
                rsProduto!Codigo & vbTab & _
                rsProduto!Nome & vbTab & _
                rsProduto!Valor & vbTab & _
                IIf(rsProduto!Inativo = 1, "Sim", "Não")
            rsProduto.MoveNext
        Loop
    End With
End Sub

Private Sub grdProduto_DblClick()
    Selecionar
End Sub

Private Sub cmdSelecionar_Click()
    Selecionar
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Selecionar()
    If grdProduto.Row = 0 Then Exit Sub

    CodigoSelecionado = CLng(grdProduto.TextMatrix(grdProduto.Row, 0))
    Unload Me
End Sub

Private Sub grdProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
        Exit Sub
    End If
End Sub
