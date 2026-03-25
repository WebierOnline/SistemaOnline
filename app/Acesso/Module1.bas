Attribute VB_Name = "Module1"
'Public con As ADODB.Connection




Public Sub Conectar()
Dim sSQL As String
Dim rsUsuario As ADODB.Recordset
Dim rsAcesso As ADODB.Recordset
'Dim rsAcesso As ADODB.Recordset
'Dim rsUsuario As ADODB.Recordset

sSQL = "SELECT * FROM acesso ORDER BY codigo"
Set rsAcesso = dbData.OpenRecordset(sSQL)

sSQL = "SELECT usuario.* FROM usuario ORDER BY codigo"
Set rsUsuario = dbData.OpenRecordset(sSQL)

'antigo=========
'Set con = New ADODB.Connection
'Set rsAcesso = New ADODB.Recordset
'Set rsUsuario = New ADODB.Recordset

'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'         "Data Source=" & App.path & "\Banco.mdb;"



'rsAcesso.Open "SELECT * FROM acesso ORDER BY codigo;", con, adOpenKeyset, adLockOptimistic
'rsUsuario.Open "SELECT usuario.* FROM usuario ORDER BY codigo;", con, adOpenKeyset, adLockOptimistic
End Sub

Public Sub Desconectar()
'rsAcesso.Close
'rsUsuario.Close
'con.Close

'Set rsAcesso = Nothing
'Set rsUsuario = Nothing
'Set con = Nothing
End Sub
