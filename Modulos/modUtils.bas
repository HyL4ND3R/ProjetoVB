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

'Função para validar Numero (só para lembrar)
Function ValidaNumeroSoLembrar()
    Dim Controle As Long

    If Trim(txtControle.Text) = "" Then
        MsgBox "Informe o controle"
        Exit Function   'Exit Sub
    End If
    
    If Not IsNumeric(txtControle.Text) Then
        MsgBox "Controle deve ser numérico"
        Exit Function
    End If
    
    Controle = CLng(txtControle.Text)
End Function
