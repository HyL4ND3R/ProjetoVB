Attribute VB_Name = "modRecordset"
Public rsOperadorLogado As ADODB.Recordset 'Usado para o operadorLogado
Public rsOperador As ADODB.Recordset 'Tabela Operador
Public rsCliente As ADODB.Recordset 'Tabela Cliente
Public rsClienteCod As ADODB.Recordset 'Buscar Cliente especifico
Public rsProduto As ADODB.Recordset 'Tabela Produto
Public rsProdutoCod As ADODB.Recordset 'Buscar Produto especifico
Public rsPedido As ADODB.Recordset 'Tabela Pedido
Public rsPedidoCod As ADODB.Recordset 'Buscar Pedido especifico
Public rsProximoCodigo As ADODB.Recordset 'Buscar próximo codigo de pedido
Public rsPedidoItem As ADODB.Recordset 'Tabela PedidoItem
Public rsPedidoItemCod As ADODB.Recordset 'Próximo cod de ITEM do pedido
'-------------OPERADORES------------------------------------------------------------------------------------------------------
Public Sub CarregarOperadores()

    Set rsOperador = New ADODB.Recordset

    rsOperador.CursorLocation = adUseClient
    rsOperador.Open _
        "SELECT Codigo, Nome, Senha, Admin, Inativo FROM Operador ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub
'-------------CLIENTES------------------------------------------------------------------------------------------------------
Public Sub CarregarClientes()

    Set rsCliente = New ADODB.Recordset

    rsCliente.CursorLocation = adUseClient
    rsCliente.Open _
        "SELECT Codigo, Nome, TipoDocumento, " & _
        "Case TipoDocumento When 0 Then 'CPF' When 1 Then 'CNPJ' ELSE 'Outros' End as TipoDocumentoExtenso, " & _
        "Documento, Telefone, Inativo FROM Cliente ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub BuscarClientePorCodigo(CodCliente As Integer)
    
    Set rsClienteCod = New ADODB.Recordset
    
    rsClienteCod.CursorLocation = adUseClient
    rsClienteCod.Open _
        "SELECT * FROM Cliente WHERE Codigo = " & CodCliente, _
        Conn, adOpenStatic, adLockReadOnly

End Sub
'-------------PRODUTOS------------------------------------------------------------------------------------------------------
Public Sub CarregarProdutos()

    Set rsProduto = New ADODB.Recordset

    rsProduto.CursorLocation = adUseClient
    rsProduto.Open _
        "SELECT Codigo, Nome, Valor, Inativo FROM Produto ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub BuscarProdutoPorCodigo(CodCliente As Integer)
    
    Set rsProdutoCod = New ADODB.Recordset
    
    rsProdutoCod.CursorLocation = adUseClient
    rsProdutoCod.Open _
        "SELECT * FROM Produto WHERE Codigo = " & CodCliente, _
        Conn, adOpenStatic, adLockReadOnly

End Sub
'-------------PEDIDOS------------------------------------------------------------------------------------------------------
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

Public Sub BuscarPedidoPorCodigo(CodPedido As Long)
    Set rsPedidoCod = New ADODB.Recordset
    
    rsPedidoCod.CursorLocation = adUseClient
    rsPedidoCod.Open _
        "select Pedido.Controle, Pedido.Codigo, Pedido.ClienteCodigo, Cliente.Nome As Cliente, Pedido.Data, " & _
            "ISNULL(Pedido.QtdeTotal, 0) as QtdeTotal, ISNULL(Pedido.ValorTotal, 0) as ValorTotal " & _
            "From pedido " & _
            "Inner join Cliente on Pedido.ClienteCodigo = Cliente.Codigo " & _
            "Where Pedido.Codigo = " & CodPedido, _
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
    
    Dim Sql As String
    
    Sql = "Insert into Pedido (Codigo, ClienteCodigo, Data) Values (" & _
            pedido.Codigo & ", " & _
            pedido.ClienteCodigo & ", " & _
            "'" & Format(pedido.DataPedido, "yyyy-MM-dd") & "') "
    
    Conn.Execute Sql
    InserirPedido = True
    Exit Function
    
Erro:
    InserirPedido = False
End Function

Public Function AlterarPedido(pedido As cPedido) As Boolean
    On Error GoTo Erro

    Dim Sql As String

    Sql = "UPDATE Pedido SET " & _
          "Codigo = " & pedido.Codigo & ", " & _
          "ClienteCodigo = " & pedido.ClienteCodigo & ", " & _
          "Data = '" & Format(pedido.DataPedido, "yyyy-MM-dd") & "' " & _
          "WHERE Controle = " & pedido.Controle

    Conn.Execute Sql

    AlterarPedido = True
    Exit Function

Erro:
    AlterarPedido = False
End Function

Public Function BuscaProximoCodItemPedido(ByVal ControlePedido As Long) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    Set rsPedidoItemCod = New ADODB.Recordset

    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "SELECT ISNULL(Max(item),0) + 1 as Item " & _
            "FROM PedidoItem WHERE ControlePedido = ?"

        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , ControlePedido)
    End With

    rsPedidoItemCod.CursorLocation = adUseClient
    rsPedidoItemCod.Open cmd, , adOpenStatic, adLockReadOnly

    Set BuscarItensPedido = rsPedidoItemCod
    
    BuscaProximoCodItemPedido = True
    Exit Function
Erro:
    BuscaProximoCodItemPedido = False

End Function

'Função para inserir item no pedido usando ADO Command e alterando os parametros
Public Function InserirItemPedido(itemPedido As cPedidoItem) As Boolean
    On Error GoTo Erro

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "INSERT INTO PedidoItem " & _
            "(ControlePedido, Item, ProdutoCodigo, Descricao, Quantidade, ValorUn, ValorTotal) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?)"

        ' --- Alterando Parâmetros ---
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , itemPedido.ControlePedido)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , itemPedido.Item)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , itemPedido.ProdutoCodigo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, itemPedido.Descricao)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , itemPedido.Qtde)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , itemPedido.ValorUn)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , itemPedido.ValorTotal)

        .Execute
    End With

    InserirItemPedido = True
    Exit Function

Erro:
    InserirItemPedido = False
End Function


Public Function AlterarItemPedidoxxxx(itemPedido As cPedidoItem) As Boolean
    On Error GoTo Erro
    
    Dim Sql As String
    
    Sql = "UPDATE PedidoItem SET " & _
            "Item = " & itemPedido.Item & ", " & _
            "ProdutoCodigo = " & itemPedido.ProdutoCodigo & ", " & _
            "Descricao = '" & itemPedido.Descricao & "', " & _
            "Quantidade = " & itemPedido.Qtde & ", " & _
            "ValorUn = " & itemPedido.ValorUn & ", " & _
            "ValorTotal = " & itemPedido.ValorTotal & _
            "Where Controle = " & itemPedido.Controle
    
    Conn.Execute Sql
    
    AlterarItemPedidoxxxx = True
    Exit Function
    
Erro:
    AlterarItemPedidoxxxx = False
End Function

Public Function AlterarItemPedido(itemPedido As cPedidoItem) As Boolean
    On Error GoTo Erro

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "UPDATE PedidoItem SET " & _
            "Item = ?, " & _
            "ProdutoCodigo = ?, " & _
            "Descricao = ?, " & _
            "Quantidade = ?, " & _
            "ValorUn = ?, " & _
            "ValorTotal = ? " & _
            "Where Controle = ? "

        ' --- Parametros ---
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , itemPedido.Item)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , itemPedido.ProdutoCodigo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, itemPedido.Descricao)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , itemPedido.Qtde)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , itemPedido.ValorUn)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , itemPedido.ValorTotal)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , itemPedido.Controle)

        .Execute
    End With

    AlterarItemPedido = True
    Exit Function

Erro:
    AlterarItemPedido = False
End Function

'-------------FUNÇÕES GENERICAS------------------------------------------------------------------------------------------------------
Public Function BuscarRS(rs As ADODB.Recordset, _
                         ByVal campo As String, _
                         ByVal Valor As Variant) As Boolean

    rs.MoveFirst
    rs.Find campo & " = " & Valor
    BuscarRS = Not rs.EOF

End Function


