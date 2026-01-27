VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptClientes 
   Caption         =   "Impressao"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18270
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   32226
   _ExtentY        =   14473
   SectionData     =   "rptClientes.dsx":0000
End
Attribute VB_Name = "rptClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsDadosImpressao As ADODB.Recordset


Private Sub ActiveReport_ReportStart()
    rsDadosImpressao.MoveFirst
    
    If rsDadosImpressao.eof Then
        Exit Sub
    End If

    txtCodigo.Text = rsDadosImpressao!Codigo
    txtNome.Text = rsDadosImpressao!Nome

    rsDadosImpressao.MoveNext
    
End Sub
