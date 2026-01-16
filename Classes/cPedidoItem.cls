VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPedidoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- Atributos ---

Private mControle As Long
Private mControlePedido As Long
Private mItem As Long
Private mProdutoCodigo As Long
Private mDescricao As String
Private mQtde As Currency
Private mValorUn As Currency
Private mValorTotal As Currency

' --- Propriedades ---

Public Property Get Controle() As Long
    Controle = mControle
End Property

Public Property Let Controle(ByVal v As Long)
    mControle = v
End Property

Public Property Get ControlePedido() As Long
    ControlePedido = mControlePedido
End Property

Public Property Let ControlePedido(ByVal v As Long)
    mControlePedido = v
End Property

Public Property Get Item() As Long
    Item = mItem
End Property

Public Property Let Item(ByVal v As Long)
    mItem = v
End Property

Public Property Get ProdutoCodigo() As Long
    ProdutoCodigo = mProdutoCodigo
End Property

Public Property Let ProdutoCodigo(ByVal v As Long)
    mProdutoCodigo = v
End Property

Public Property Get Descricao() As String
    Descricao = mDescricao
End Property

Public Property Let Descricao(ByVal v As String)
    mDescricao = v
End Property

Public Property Get Qtde() As Currency
    Qtde = mQtde
End Property

Public Property Let Qtde(ByVal v As Currency)
    mQtde = v
End Property

Public Property Get ValorUn() As Currency
    ValorUn = mValorUn
End Property

Public Property Let ValorUn(ByVal v As Currency)
    mValorUn = v
End Property

Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property

Public Property Let ValorTotal(ByVal v As Currency)
    mValorTotal = v
End Property
