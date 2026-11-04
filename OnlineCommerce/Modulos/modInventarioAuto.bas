Attribute VB_Name = "modInventarioAuto"
Option Explicit
'Inventario MENSAL automatico (Inventario_Cadastro / TbInventarios + TbInventariosItens).
'No primeiro login do mes gera o inventario do MES ANTERIOR (tipo "Mensal"), assim o estoque
'de 31/12 que a contabilidade pede em fevereiro ja esta salvo desde o comeco de janeiro.
'Obs: o saldo gravado e o ESTOQUE_FISCAL do momento da geracao.

Private Sub INV_Log(ByVal sMsg As String)
Dim nLog As Integer
On Error Resume Next
nLog = FreeFile
Open appPathApp & "InventarioAuto_log.txt" For Append As #nLog
Print #nLog, Format(Now, "dd/mm/yyyy hh:nn:ss") & "  " & sMsg
Close #nLog
End Sub

'Chamada logo apos o login (Senha.frm). Silenciosa: erro vai para InventarioAuto_log.txt.
Public Sub GerarInventarioMensalAuto(ByVal vCodUsuario As Long)
Dim dMes As Date
Dim nMes As Integer, nAno As Integer
Dim sErro As String
Const FILIAL As Integer = 1

On Error GoTo deuErro

'primeiro login do mes (qualquer dia da semana, inclusive feriado): se o mes anterior ainda nao foi gerado, gera.
'Loja fechada no dia 1 = fica pendente e gera no proximo dia em que alguem logar.
If vCodUsuario <= 0 Then Exit Sub

dMes = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
nMes = Month(dMes)
nAno = Year(dMes)

'ja gerado (por login anterior, outra estacao ou manualmente)?
If SQLExecutaRetorno("SELECT COUNT(*) r FROM TbInventarios WHERE (PeriodoInventarioTipo = 'Mensal') AND (AnoInventario = '" & nAno & "') AND (PeriodoInventario = '" & nMes & "') AND (IdFilial = " & FILIAL & ")", "r", 0) > 0 Then Exit Sub

sErro = SQLExecuta("EXEC InventarioGerar " & nAno & ", " & vCodUsuario & ", 'Mensal', '" & nMes & "', NULL, " & FILIAL)
If sErro <> "" Then
   INV_Log "Inventario mensal " & Format(dMes, "mm/yyyy") & " NAO gerado: " & sErro
Else
   INV_Log "Inventario mensal " & Format(dMes, "mm/yyyy") & " gerado (usuario " & vCodUsuario & ")."
End If
Exit Sub

deuErro:
INV_Log "Erro no inventario mensal automatico: " & Err.Description
Err.Clear
End Sub

'Chamada logo apos o login (Senha.frm): recalcula Produtos.ESTOQUE_FISCAL = saldo fiscal de hoje
'(inventario de abertura + notas). So roda quando ja existe inventario fechado (a abertura) - sem base
'o calculo daria saldos negativos. Silenciosa: erro vai para InventarioAuto_log.txt.
Public Sub RecalcularEstoqueFiscalAuto()
Dim sErro As String

On Error GoTo deuErro

If SQLExecutaRetorno("SELECT COUNT(*) r FROM TbInventarios WHERE (Fechado = 1) AND (DataInventario IS NOT NULL) AND (CAST(DataInventario AS DATE) <= CAST(GETDATE() AS DATE))", "r", 0) = 0 Then Exit Sub

sErro = SQLExecuta("EXEC EstoqueFiscalRecalcular")
If sErro <> "" Then INV_Log "Estoque fiscal NAO recalculado: " & sErro
Exit Sub

deuErro:
INV_Log "Erro no recalculo do estoque fiscal: " & Err.Description
Err.Clear
End Sub
