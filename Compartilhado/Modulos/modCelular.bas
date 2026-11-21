Attribute VB_Name = "modCelular"
Option Explicit
'Validação de celular brasileiro (cadastro de cliente e envios pelo WhatsApp).
'Válido: 11 dígitos (DDD + 9 dígitos começando com 9), DDD existente, sem máscara incompleta e sem número repetido.
'Aceita o 55 do país na frente (13 dígitos).

Private Const DDDS_VALIDOS As String = "|11|12|13|14|15|16|17|18|19|21|22|24|27|28|31|32|33|34|35|37|38|" & _
    "41|42|43|44|45|46|47|48|49|51|53|54|55|61|62|63|64|65|66|67|68|69|71|73|74|75|77|79|" & _
    "81|82|83|84|85|86|87|88|89|91|92|93|94|95|96|97|98|99|"

'True se o celular é válido; senão sMotivo diz por quê (ex.: "falta um dígito")
Public Function CelularValido(ByVal sNumero As String, Optional ByRef sMotivo As String) As Boolean
Dim d As String, i As Long, ch As String

sMotivo = ""
For i = 1 To Len(sNumero)
    ch = Mid$(sNumero, i, 1)
    If ch Like "#" Then d = d & ch
Next
If Len(d) = 13 And Left$(d, 2) = "55" Then d = Mid$(d, 3)

If Len(d) = 0 Then
    sMotivo = "sem número"
ElseIf InStr(sNumero, "_") > 0 And Len(d) < 11 Then
    sMotivo = "número incompleto"
ElseIf Len(d) = 10 Then
    If Mid$(d, 3, 1) Like "[2-5]" Then
        sMotivo = "telefone fixo, sem WhatsApp"
    Else
        sMotivo = "falta um dígito (celular tem 9 dígitos depois do DDD)"
    End If
ElseIf Len(d) < 10 Then
    sMotivo = "número incompleto"
ElseIf Len(d) > 11 Then
    sMotivo = "dígitos a mais"
ElseIf InStr(DDDS_VALIDOS, "|" & Left$(d, 2) & "|") = 0 Then
    sMotivo = "o DDD " & Left$(d, 2) & " não existe"
ElseIf Mid$(d, 3, 1) <> "9" Then
    sMotivo = "o celular deve começar com 9 depois do DDD"
ElseIf Mid$(d, 3) = String$(9, Mid$(d, 4, 1)) Then
    sMotivo = "número inválido"
Else
    CelularValido = True
End If
End Function
