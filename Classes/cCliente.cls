VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCodigo As Long
Private mNome As String
Private mTipoDocumento As eTipoDocumentoCliente
Private mDocumento As String
Private mTelefone As String
Private mInativo As Integer

' --- Propriedades ---

Public Property Get codigo() As Long
    codigo = mCodigo
End Property

Public Property Let codigo(ByVal v As Long)
    mCodigo = v
End Property

Public Property Get Nome() As String
    Nome = mNome
End Property

Public Property Let Nome(ByVal v As String)
    mNome = v
End Property

Public Property Get TipoDocumento() As eTipoDocumentoCliente
    TipoDocumento = mTipoDocumento
End Property

Public Property Let TipoDocumento(ByVal v As eTipoDocumentoCliente)
    mTipoDocumento = v
End Property

Public Property Get Documento() As String
    Documento = mDocumento
End Property

Public Property Let Documento(ByVal v As String)
    mDocumento = v
End Property

Public Property Get Telefone() As String
    Telefone = mTelefone
End Property

Public Property Let Telefone(ByVal v As String)
    mTelefone = v
End Property

Public Property Get Inativo() As Integer
    Inativo = mInativo
End Property

Public Property Let Inativo(ByVal v As Integer)
    mInativo = v
End Property
