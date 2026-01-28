VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arRelatorioPedidos 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18300
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   32279
   _ExtentY        =   14261
   SectionData     =   "arRelatorioPedidos.dsx":0000
End
Attribute VB_Name = "arRelatorioPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fItem As Integer
Private codPedido As Integer
Private rs As ADODB.Recordset

Private Sub ActiveReport_ReportStart()
    
    Set rs = Me.dcRelPedidos.Recordset
    codPedido = 0
    primeiroItem = True
    fItem = 0
    
End Sub

Private Sub Detail_Format()

    'Alterar entre fundo branco e cinza
    If fItem Mod 2 = 0 Then
        Detail.BackColor = &HFFFFFF
    Else
        Detail.BackColor = &HE0E0E0
    End If
    fItem = fItem + 1

    If (codPedido = CLng(rs!pedido)) Then
        'fazer aqui a ocultação dos campos
    Else
        codPedido = CLng(rs!pedido)
    End If
    
    
    

End Sub

Private Sub PageFooter_Format()
    
    fldOperador = "Operador: " & rsOperadorLogado!Codigo & " - " & rsOperadorLogado!Nome & ", Emitido em: " & _
                    Format(Date, "dd/MM/yyyy") & " as " & Format(Date, "hh:mm")
    fldPagina = "Página: " & Format(pageNumber, "###000")
    
End Sub

