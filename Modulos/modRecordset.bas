Attribute VB_Name = "modRecordset"
Public rsOperador As ADODB.Recordset

Public Sub CarregarOperadores()

    Set rsOperador = New ADODB.Recordset

    rsOperador.CursorLocation = adUseClient
    rsOperador.Open _
        "SELECT Codigo, Nome, Senha, Admin, Inativo FROM Operador ORDER BY Codigo", _
        Conn, adOpenStatic, adLockReadOnly

End Sub
