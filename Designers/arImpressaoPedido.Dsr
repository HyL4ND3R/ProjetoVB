VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arImpressaoPedido 
   Caption         =   "Impressao"
   ClientHeight    =   15615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   28560
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   50377
   _ExtentY        =   27543
   SectionData     =   "arImpressaoPedido.dsx":0000
End
Attribute VB_Name = "arImpressaoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fItem As Integer

Private Sub ActiveReport_ReportStart()
    
    fItem = 0
    
    'Ajustando formatação dos campos
    fldQtdeTotal.OutputFormat = "#,##0.00"
    fldValorTotal.OutputFormat = "#,##0.00"
    fldProdutoQtde.OutputFormat = "#,##0.00"
    fldProdutoValorUn.OutputFormat = "#,##0.00"
    fldProdutoValorTotal.OutputFormat = "#,##0.00"
    
End Sub

Private Sub Detail_Format()

    'Alterar entre fundo branco e cinza
    If fItem Mod 2 = 0 Then
        Detail.BackColor = &HFFFFFF
    Else
        Detail.BackColor = &HE0E0E0
    End If
    fItem = fItem + 1
    
End Sub

Private Sub PageFooter_Format()
    
    fldOperador = "Operador: " & rsOperadorLogado!Codigo & " - " & rsOperadorLogado!Nome & ", Emitido em: " & _
                    Format(Date, "dd/MM/yyyy") & " as " & Left(Time, 5)
    fldPagina = "Página: " & Format(pageNumber, "###000")
    
End Sub



