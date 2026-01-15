Attribute VB_Name = "modUtils"
'Função para verificar null e voltar um valor padrão
Function VerificaNull(ByVal v As Variant, Optional ByVal padrao As Variant) As Variant
    If IsNull(v) Then ' Se for null
        VerificaNull = padrao 'Retorna o Valor Padrão
    Else
        VerificaNull = v ' Se não, retorna o valor
    End If
End Function

'Função para validar Numero (só para lembrar)
Function ValidaNumeroSoLembrar()
    Dim controle As Long

    If Trim(txtControle.Text) = "" Then
        MsgBox "Informe o controle"
        Exit Function   'Exit Sub
    End If
    
    If Not IsNumeric(txtControle.Text) Then
        MsgBox "Controle deve ser numérico"
        Exit Function
    End If
    
    controle = CLng(txtControle.Text)
End Function
