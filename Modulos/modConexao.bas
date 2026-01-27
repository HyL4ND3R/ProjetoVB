Attribute VB_Name = "modConexao"
Public Conn As ADODB.Connection

Public Sub AbrirConexao()
    If Conn Is Nothing Then
        Set Conn = New ADODB.Connection
        'Tenq ter o Persist pra funfa o AR
        Conn.ConnectionString = _
            "Provider=SQLOLEDB;" & _
            "Data Source=localhost;" & _
            "Persist Security Info=True;" & _
            "Initial Catalog=PROJETOVB;" & _
            "User ID=sa;" & _
            "Password=sae;"
        Conn.Open
    End If
End Sub

Public Sub FecharConexao()
    If Not Conn Is Nothing Then
        If Conn.State = adStateOpen Then Conn.Close
        Set Conn = Nothing
    End If
End Sub
