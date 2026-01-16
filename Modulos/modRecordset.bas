Attribute VB_Name = "modRecordset"
Public rsOperador As ADODB.Recordset
Public rsCliente As ADODB.Recordset
Public rsProduto As ADODB.Recordset
Public rsPedido As ADODB.Recordset
Public rsProximoCodigo As ADODB.Recordset
Public rsPedidoItem As ADODB.Recordset

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

Public Sub CarregarPedidos()

    Set rsPedido = New ADODB.Recordset

    rsPedido.CursorLocation = adUseClient
    rsPedido.Open _
        "Select Pedido.Controle, Pedido.Codigo, Pedido.ClienteCodigo, Cliente.Nome ClienteNome, Pedido.Data DataPedido, Pedido.ValorTotal  " & _
            "From Pedido " & _
            "inner join Cliente on Cliente.Codigo = Pedido.ClienteCodigo " & _
            "Order By Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub BuscarProximoCodPedido()
    Set rsProximoCodigo = New ADODB.Recordset
    
    rsProximoCodigo.CursorLocation = adUseClient
    rsProximoCodigo.Open _
        "select isnull(max(Codigo),0) + 1 as Codigo from Pedido", _
        Conn, adOpenStatic, adLockReadOnly
End Sub

Public Sub CarregarItensPedido(codigoPedido As Long)
    Set rsPedidoItem = New ADODB.Recordset
    
    rsPedidoItem.CursorLocation = adUseClient
    rsPedidoItem.Open _
        "Select Controle, ControlePedido, Item, ProdutoCodigo, Descricao, Quantidade, ValorUn, ValorTotal " & _
            "From PedidoItem " & _
            "Where ControlePedido = " & codigoPedido & _
            "Order By Item", _
        Conn, adOpenStatic, adLockReadOnly
    
End Sub

Public Function InserirPedido(pedido As cPedido)
    On Error GoTo Erro
    
    Dim sql As String
    
    sql = "Insert into Pedido (Codigo, ClienteCodigo, Data) Values (" & _
            pedido.Codigo & ", " & _
            pedido.ClienteCodigo & ", " & _
            "'" & Format(pedido.DataPedido, "yyyy-MM-dd") & "') "
    
    Conn.Execute sql
    InserirPedido = True
    Exit Function
    
Erro:
    InserirPedido = False
End Function

Public Function AlterarPedido(pedido As cPedido) As Boolean
    On Error GoTo Erro

    Dim sql As String

    sql = "UPDATE Pedido SET " & _
          "Codigo = " & pedido.Codigo & ", " & _
          "ClienteCodigo = " & pedido.ClienteCodigo & ", " & _
          "Data = '" & Format(pedido.DataPedido, "yyyy-MM-dd") & "' " & _
          "WHERE Controle = " & pedido.Controle

    Conn.Execute sql

    AlterarPedido = True
    Exit Function

Erro:
    AlterarPedido = False
End Function

Public Function InserirItemPedido(itemPedido As cPedidoItem) As Boolean
    On Error GoTo Erro
    
    Dim sql As String
    
    sql = "Insert into PedidoItem (ControlePedido, Item, ProdutoCodigo, Descricao, Quantidade, ValorUn, ValorTotal) " & _
            "Values (" & itemPedido.ControlePedido & ", " & _
            itemPedido.Item & ", " & _
            itemPedido.ProdutoCodigo & ", " & _
            "'" & itemPedido.Descricao & "', " & _
            itemPedido.Qtde & ", " & _
            itemPedido.ValorUn & ", " & _
            itemPedido.ValorTotal & ")"
    
    Conn.Execute sql
    
    InserirItemPedido = True
    Exit Function
    
Erro:
    InserirItemPedido = False
End Function

Public Function AlterarItemPedido(itemPedido As cPedidoItem) As Boolean
    On Error GoTo Erro
    
    Dim sql As String
    
    sql = "UPDATE PedidoItem SET " & _
            "Item = " & itemPedido.Item & ", " & _
            "ProdutoCodigo = " & itemPedido.ProdutoCodigo & ", " & _
            "Descricao = '" & itemPedido.Descricao & "', " & _
            "Quantidade = " & itemPedido.Qtde & ", " & _
            "ValorUn = " & itemPedido.ValorUn & ", " & _
            "ValorTotal = " & itemPedido.ValorTotal & _
            "Where Controle = " & itemPedido.Controle
    
    Conn.Execute sql
    
    AlterarItemPedido = True
    Exit Function
    
Erro:
    AlterarItemPedido = False
End Function

Public Function BuscarRS(rs As ADODB.Recordset, _
                         ByVal campo As String, _
                         ByVal valor As Variant) As Boolean

    rs.MoveFirst
    rs.Find campo & " = " & valor
    BuscarRS = Not rs.EOF

End Function


