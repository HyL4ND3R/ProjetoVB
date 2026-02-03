VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCodigo As Long
Private mNome As String
Private mSenha As String
Private mAdmin As Integer
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

Public Property Get Senha() As String
    Senha = mSenha
End Property

Public Property Let Senha(ByVal v As String)
    mSenha = v
End Property

Public Property Get Admin() As Integer
    Admin = mAdmin
End Property

Public Property Let Admin(ByVal v As Integer)
    mAdmin = v
End Property

Public Property Get Inativo() As Integer
    Inativo = mInativo
End Property

Public Property Let Inativo(ByVal v As Integer)
    mInativo = v
End Property
