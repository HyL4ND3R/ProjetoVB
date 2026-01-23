Attribute VB_Name = "modStoreProcedures"
Public Function CalculaTotaisPedido(ControlePedido As Long) As Boolean
    On Error GoTo Erro
    
    Dim Sql As String
    
    Sql = "EXEC sp_RecalcularTotaisPedido " & ControlePedido
    Conn.Execute Sql
    
    CalculaTotaisPedido = True
    Exit Function
    
Erro:
    CalculaTotaisPedido = False
End Function
