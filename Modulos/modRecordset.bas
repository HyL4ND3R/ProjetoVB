Attribute VB_Name = "modRecordset"
Option Explicit

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

Public Function InserirOperador(operador As cOperador) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "INSERT INTO Operador (Nome, Senha, Admin, Inativo) Values (?,?,?,?)"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, operador.Nome)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, operador.Senha)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , operador.Admin)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , operador.Inativo)
        
        .Execute
    End With
    
    InserirOperador = True
    Exit Function
    
Erro:
    InserirOperador = False
End Function

Public Function AlterarOperador(operador As cOperador) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "UPDATE Operador set Nome = ?," & _
            "Senha = ?," & _
            "Admin = ?," & _
            "Inativo = ? " & _
            "WHERE Codigo = ?"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, operador.Nome)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, operador.Senha)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , operador.Admin)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , operador.Inativo)
        .Parameters.Append .CreateParameter(, adBigInt, adParamInput, , operador.codigo)
        
        .Execute
    End With
    
    AlterarOperador = True
    Exit Function
    
Erro:
    AlterarOperador = False
End Function

Public Function ExcluirOperador(codigo As Long) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "DELETE FROM Operador WHERE Codigo = ?"
        
        .Parameters.Append .CreateParameter(, adBigInt, adParamInput, , codigo)
        
        .Execute
    End With
    
    ExcluirOperador = True
    Exit Function
    
Erro:
    ExcluirOperador = False
End Function

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

Public Sub CarregarClientesAtivos()

    Set rsCliente = New ADODB.Recordset

    rsCliente.CursorLocation = adUseClient
    rsCliente.Open _
        "SELECT Codigo, Nome, TipoDocumento, " & _
        "Case TipoDocumento When 0 Then 'CPF' When 1 Then 'CNPJ' ELSE 'Outros' End as TipoDocumentoExtenso, " & _
        "Documento, Telefone, Inativo FROM Cliente " & _
        "WHERE Inativo = 0 " & _
        "ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub BuscarClientePorCodigo(CodCliente As Integer)
    
    Set rsClienteCod = New ADODB.Recordset
    
    rsClienteCod.CursorLocation = adUseClient
    rsClienteCod.Open _
        "SELECT * FROM Cliente WHERE Codigo = " & CodCliente, _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub BuscarClienteAtivoPorCodigo(CodCliente As Integer)
    
    Set rsClienteCod = New ADODB.Recordset
    
    rsClienteCod.CursorLocation = adUseClient
    rsClienteCod.Open _
        "SELECT * FROM Cliente WHERE Codigo = " & CodCliente, _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Function InserirCliente(cliente As cCliente) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
                "INSERT INTO Cliente (Nome, TipoDocumento, Documento, Telefone, Inativo) " & _
                "VALUES (?,?,?,?,?)"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, cliente.Nome)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , cliente.TipoDocumento)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, cliente.Documento)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, cliente.Telefone)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , cliente.Inativo)
        
        .Execute
    End With
    
    InserirCliente = True
    Exit Function
    
Erro:
    InserirCliente = False
End Function

Public Function AlterarCliente(cliente As cCliente) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
                "UPDATE Cliente set Nome = ?, " & _
                "TipoDocumento = ?, " & _
                "Documento = ?, " & _
                "Telefone = ?, " & _
                "Inativo = ? " & _
                "WHERE Codigo = ? "
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, cliente.Nome)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , cliente.TipoDocumento)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, cliente.Documento)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, cliente.Telefone)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , cliente.Inativo)
        .Parameters.Append .CreateParameter(, adBigInt, adParamInput, , cliente.codigo)
        
        .Execute
    End With
    
    AlterarCliente = True
    Exit Function
    
Erro:
    AlterarCliente = False
End Function

Public Function ExcluirCliente(codigo As Long) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "DELETE FROM Cliente WHERE Codigo = ?"
        
        .Parameters.Append .CreateParameter(, adBigInt, adParamInput, , codigo)
        
        .Execute
    End With
    
    ExcluirCliente = True
    Exit Function
    
Erro:
    ExcluirCliente = False
End Function

'-------------PRODUTOS------------------------------------------------------------------------------------------------------
Public Sub CarregarProdutos()

    Set rsProduto = New ADODB.Recordset

    rsProduto.CursorLocation = adUseClient
    rsProduto.Open _
        "SELECT Codigo, Nome, Valor, Inativo FROM Produto ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub CarregarProdutosAtivos()

    Set rsProduto = New ADODB.Recordset

    rsProduto.CursorLocation = adUseClient
    rsProduto.Open _
        "SELECT Codigo, Nome, Valor, Inativo FROM Produto " & _
        "WHERE Inativo = 0 " & _
        "ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub BuscarProdutoPorCodigo(codProduto As Integer)
    
    Set rsProdutoCod = New ADODB.Recordset
    
    rsProdutoCod.CursorLocation = adUseClient
    rsProdutoCod.Open _
        "SELECT * FROM Produto WHERE Codigo = " & codProduto, _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Sub BuscarProdutoAtivoPorCodigo(CodCliente As Integer)
    
    Set rsProdutoCod = New ADODB.Recordset
    
    rsProdutoCod.CursorLocation = adUseClient
    rsProdutoCod.Open _
        "SELECT * FROM Produto WHERE Codigo = " & CodCliente & ", " & _
        "AND Inativo = 0", _
        Conn, adOpenStatic, adLockReadOnly

End Sub

Public Function InserirProduto(produto As cProduto) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
                "INSERT INTO Produto (Nome, Valor, Inativo) " & _
                "VALUES (?,?,?)"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, produto.Nome)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , produto.Valor)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , produto.Inativo)
        
        .Execute
    End With
    
    InserirProduto = True
    Exit Function
    
Erro:
    InserirProduto = False
End Function

Public Function AlterarProduto(produto As cProduto) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
                "UPDATE Produto SET Nome = ?, " & _
                "Valor = ?, " & _
                "Inativo = ? " & _
                "WHERE Codigo = ?"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, produto.Nome)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , produto.Valor)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , produto.Inativo)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , produto.codigo)
        
        .Execute
    End With
    
    AlterarProduto = True
    Exit Function
    
Erro:
    AlterarProduto = False
End Function

Public Function ExcluirProduto(codigo As Long) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "DELETE FROM Produto WHERE Codigo = ?"
        
        .Parameters.Append .CreateParameter(, adBigInt, adParamInput, , codigo)
        
        .Execute
    End With
    
    ExcluirProduto = True
    Exit Function
    
Erro:
    ExcluirProduto = False
End Function

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
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "Insert into Pedido (Codigo, ClienteCodigo, Data) Values (?,?,?)"
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , pedido.codigo)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , pedido.ClienteCodigo)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , pedido.DataPedido)
        .Execute
    End With
    
    InserirPedido = True
    Exit Function
    
Erro:
    InserirPedido = False
End Function

Public Function AlterarPedido(pedido As cPedido) As Boolean
    On Error GoTo Erro

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "UPDATE Pedido SET " & _
            "Codigo = ? ," & _
            "ClienteCodigo = ? ," & _
            "Data = ? " & _
            "WHERE Controle = ?"
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , pedido.codigo)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , pedido.ClienteCodigo)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , pedido.DataPedido)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , pedido.Controle)
        .Execute
    End With

    AlterarPedido = True
    Exit Function

Erro:
    AlterarPedido = False
End Function

Public Function ExcluirPedido(codigo As Long) As Boolean
    On Error GoTo Erro
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = _
            "DELETE FROM Pedido WHERE Codigo = ?"
        
        .Parameters.Append .CreateParameter(, adBigInt, adParamInput, , codigo)
        
        .Execute
    End With
    
    ExcluirPedido = True
    Exit Function
    
Erro:
    ExcluirPedido = False
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


