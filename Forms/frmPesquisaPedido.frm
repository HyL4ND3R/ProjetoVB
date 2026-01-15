VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPesquisaPedido 
   Caption         =   "Pesquisar Pedidos"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   7350
      TabIndex        =   1
      Top             =   4740
      Width           =   1305
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Height          =   465
      Left            =   6030
      TabIndex        =   0
      Top             =   4740
      Width           =   1305
   End
   Begin MSFlexGridLib.MSFlexGrid grdPedido 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
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
Attribute VB_Name = "frmPesquisaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoSelecionado As Long


Private Sub Form_Load()
    CarregarPedidos

        With grdPedido
            .Rows = 1
            .Cols = 4
            
            .TextMatrix(0, 0) = "Pedido"
            .TextMatrix(0, 1) = "Cliente"
            .TextMatrix(0, 2) = "Data"
            .TextMatrix(0, 3) = "Valor"
            
            .ColWidth(0) = 800
            .ColWidth(1) = 3000
            .ColWidth(2) = 800
            .ColWidth(3) = 1200
        End With
        
        Do While Not rsPedido.EOF
            .AddItem _
                rsPedido!Codigo & vbTab & _
                rsPedido!ClienteNome & vbTab & _
                rsPedido!DataPedido & vbTab & _
                rsPedido!ValorTotal
            rsPedido.MoveNext
        Loop
    End With
End Sub

Private Sub grdPedido_DblClick()
    Selecionar
End Sub

Private Sub cmdSelecionar_Click()
    Selecionar
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Selecionar()
    If grdPedido.Row = 0 Then Exit Sub

    CodigoSelecionado = CLng(grdPedido.TextMatrix(grdPedido.Row, 0))
    Unload Me
End Sub






