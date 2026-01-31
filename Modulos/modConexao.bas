Attribute VB_Name = "modConexao"
Public Conn As ADODB.Connection

Public Function AbrirConexao() As Boolean
    Dim arqINI As String
    Dim servidor As String
    Dim banco As String
    Dim usuario As String
    Dim senha As String

    On Error GoTo ErroConexao

    arqINI = App.Path & "\config.ini"

    servidor = LerINI("BANCO", "Servidor", arqINI)
    banco = LerINI("BANCO", "Banco", arqINI)
    usuario = LerINI("BANCO", "Usuario", arqINI)
    senha = LerINI("BANCO", "Senha", arqINI)

    Set Conn = New ADODB.Connection
    
    'Definindo TimeOut menor para caso de erro
    Conn.ConnectionTimeout = 5 'segundos (ex: 3, 5, 10)
    Conn.CommandTimeout = 5
    
    Conn.ConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=" & servidor & ";" & _
        "Initial Catalog=" & banco & ";" & _
        "User ID=" & usuario & ";" & _
        "Password=" & senha & ";" & _
        "Persist Security Info=True;"

    Conn.Open
    AbrirConexao = True
    Exit Function

ErroConexao:
    AbrirConexao = False
End Function


Public Sub FecharConexao()
    If Not Conn Is Nothing Then
        If Conn.State = adStateOpen Then Conn.Close
        Set Conn = Nothing
    End If
End Sub

