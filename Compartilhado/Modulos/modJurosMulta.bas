Attribute VB_Name = "modJurosMulta"
Option Explicit
'Multa e juros por atraso da tabela configuracao (Configuração Geral > Juros): MULTA (%), JUROS_MES (%), JUROS_DIA (%).
'Textos usados nos comprovantes de venda a prazo (REL_Pedido_*) e na consulta de parcelas.

Private Function JM_Valor(ByVal sNome As String) As Currency
Dim v As Variant

On Error Resume Next
v = SQLExecutaRetorno("SELECT TOP 1 config_valor AS r FROM configuracao WHERE config_nome = '" & sNome & "'", "r", "0")
If IsNumeric(v) Then JM_Valor = CCur(v)
If JM_Valor < 0 Then JM_Valor = 0
Err.Clear
End Function

Public Function JM_Multa() As Currency
JM_Multa = JM_Valor("MULTA")
End Function

Public Function JM_JurosDia() As Currency
JM_JurosDia = JM_Valor("JUROS_DIA")
End Function

Public Function JM_JurosMes() As Currency
JM_JurosMes = JM_Valor("JUROS_MES")
End Function

'"*Após o vencimento cobrar multa de 2,00% e 0,0333% de juros ao dia." ("" se não houver multa nem juros)
Public Function JM_TextoAviso() As String
Dim cMulta As Currency, cDia As Currency

cMulta = JM_Multa()
cDia = JM_JurosDia()
If cMulta > 0 And cDia > 0 Then
    JM_TextoAviso = "*Após o vencimento cobrar multa de " & Format$(cMulta, "0.00") & "% e " & Format$(cDia, "0.00##") & "% de juros ao dia."
ElseIf cMulta > 0 Then
    JM_TextoAviso = "*Após o vencimento cobrar multa de " & Format$(cMulta, "0.00") & "%."
ElseIf cDia > 0 Then
    JM_TextoAviso = "*Após o vencimento cobrar " & Format$(cDia, "0.00##") & "% de juros ao dia."
End If
End Function

'"MULTA: 2,00% | JUROS MENSAL: 1,00% | JUROS DIÁRIO: 0,0333%"
Public Function JM_TextoConfig() As String
JM_TextoConfig = "MULTA: " & Format$(JM_Multa(), "0.00") & "% | JUROS MENSAL: " & Format$(JM_JurosMes(), "0.00") & _
                 "% | JUROS DIÁRIO: " & Format$(JM_JurosDia(), "0.00##") & "%"
End Function
