VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPesquisaOperador 
   Caption         =   "Pesquisa Operadores"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grdOperador 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   8281
      _Version        =   393216
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

Private Sub Selecionar()
    If grdOperador.Row = 0 Then Exit Sub

    CodigoSelecionado = CLng(grdOperador.TextMatrix(grdOperador.Row, 0))
    Unload Me
End Sub

Private Sub MSFlexGrid1_Click()

End Sub
