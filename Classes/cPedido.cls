VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cPedido.cls

Option Explicit

Private mCodigo As Long
Private mClienteCodigo As Long
Private mDataPedido As Date
Private mValorTotal As Currency

' --- Propriedades ---

Public Property Get Codigo() As Long
    Codigo = mCodigo
End Property

Public Property Let Codigo(ByVal v As Long)
    mCodigo = v
End Property

Public Property Get ClienteCodigo() As Long
    ClienteCodigo = mClienteCodigo
End Property

Public Property Let ClienteCodigo(ByVal v As Long)
    mClienteCodigo = v
End Property

Public Property Get DataPedido() As Date
    DataPedido = mDataPedido
End Property

Public Property Let DataPedido(ByVal v As Date)
    mDataPedido = v
End Property

Public Property Get ValorTotal() As Currency
    Total = mTotal
End Property

Public Property Let ValorTotal(ByVal v As Currency)
    mTotal = v
End Property

