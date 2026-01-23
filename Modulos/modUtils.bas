Attribute VB_Name = "modUtils"
'Função para verificar null e voltar um valor padrão
Public Function VerificaNull(ByVal v As Variant, Optional ByVal padrao As Variant) As Variant
    If IsObject(v) Then
        If v Is Nothing Then
            VerificaNull = padrao
        Else
            VerificaNull = v
        End If
    Else
        If IsNull(v) Then
            VerificaNull = padrao
        Else
            VerificaNull = v
        End If
    End If
End Function


'Função para formatar numero para decimal
Public Function FormataDecimal(ByVal Valor As String) As String
    If Trim(Valor) = "" Then
        FormataDecimal = "0,00"
    ElseIf IsNumeric(Valor) Then
        FormataDecimal = Format(CDbl(Valor), "0.00")
    Else
        FormataDecimal = "0,00"
    End If
End Function

'Função para fazer o replace da virgula por um ponto, para inserir no banco
Public Function NumeroSQL(ByVal Valor As Variant) As String
    If IsNumeric(Valor) Then
        NumeroSQL = Replace(CStr(Valor), ",", ".")
    Else
        NumeroSQL = "0"
    End If
End Function
