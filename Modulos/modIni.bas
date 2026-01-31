Attribute VB_Name = "modIni"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long

Public Function LerINI(Secao As String, Chave As String, Arquivo As String) As String
    Dim Retorno As String * 255
    GetPrivateProfileString Secao, Chave, "", Retorno, 255, Arquivo
    LerINI = Left$(Retorno, InStr(Retorno, vbNullChar) - 1)
End Function

Public Sub GravarINI(Secao As String, Chave As String, Valor As String, Arquivo As String)
    WritePrivateProfileString Secao, Chave, Valor, Arquivo
End Sub

