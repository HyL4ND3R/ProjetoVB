Attribute VB_Name = "modRecordset"
Public rsOperador As ADODB.Recordset
Public rsCliente As ADODB.Recordset
Public rsProduto As ADODB.Recordset

Public Sub CarregarOperadores()

    Set rsOperador = New ADODB.Recordset

    rsOperador.CursorLocation = adUseClient
    rsOperador.Open _
        "SELECT Codigo, Nome, Senha, Admin, Inativo FROM Operador ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub CarregarClientes()

    Set rsCliente = New ADODB.Recordset

    rsCliente.CursorLocation = adUseClient
    rsCliente.Open _
        "SELECT Codigo, Nome, TipoDocumento, Documento, Telefone, Inativo FROM Cliente ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub CarregarProdutos()

    Set rsProduto = New ADODB.Recordset

    rsProduto.CursorLocation = adUseClient
    rsProduto.Open _
        "SELECT Codigo, Nome, Valor, Inativo FROM Produto ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Function BuscarRS(rs As ADODB.Recordset, _
                         ByVal campo As String, _
                         ByVal valor As Variant) As Boolean

    rs.MoveFirst
    rs.Find campo & " = " & valor
    BuscarRS = Not rs.EOF

End Function


