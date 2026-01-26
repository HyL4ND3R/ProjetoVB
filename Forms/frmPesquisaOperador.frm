VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPesquisaOperador 
   Caption         =   "Pesquisa Operadores"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Height          =   465
      Left            =   3270
      TabIndex        =   2
      Top             =   4740
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   4590
      TabIndex        =   1
      Top             =   4740
      Width           =   1305
   End
   Begin MSFlexGridLib.MSFlexGrid grdOperador 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      _ExtentX        =   10504
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
Attribute VB_Name = "frmPesquisaOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoSelecionado As Long

Private Sub Form_Load()
    CarregarOperadores

    With grdOperador
        .Rows = 1
        .Cols = 3
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Nome"
        .TextMatrix(0, 2) = "Inativo"

        Do While Not rsOperador.EOF
            .AddItem _
                rsOperador!Codigo & vbTab & _
                rsOperador!Nome & vbTab & _
                IIf(rsOperador!Inativo = 1, "Sim", "Não")
            rsOperador.MoveNext
        Loop
    End With
End Sub

Private Sub grdOperador_DblClick()
    Selecionar
End Sub

Private Sub cmdSelecionar_Click()
    Selecionar
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Selecionar()
    If grdOperador.Row = 0 Then Exit Sub

    CodigoSelecionado = CLng(grdOperador.TextMatrix(grdOperador.Row, 0))
    Unload Me
End Sub
