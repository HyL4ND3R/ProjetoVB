VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arImpressaoClientes 
   Caption         =   "Impressao"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18270
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   32226
   _ExtentY        =   14473
   SectionData     =   "arImpressaoClientes.dsx":0000
End
Attribute VB_Name = "arImpressaoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fItem As Integer

Private Sub ActiveReport_ReportStart()
    
    fItem = 0
    
End Sub

Private Sub Detail_Format()

    'Fundo
    If fItem Mod 2 = 0 Then
        Detail.BackColor = &HFFFFFF
    Else
        Detail.BackColor = &HE0E0E0
    End If
    fItem = fItem + 1

End Sub

Private Sub PageFooter_Format()
    
    fldOperador = "Operador: " & rsOperador!Codigo & " - " & rsOperador!Nome
    fldPagina = "Página: " & Format(pageNumber, "###000")
    
End Sub
