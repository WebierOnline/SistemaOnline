Attribute VB_Name = "modCobranca"
Option Explicit
'---------------------------------------------------------------------------
' Cobrança por WhatsApp (Financeiro > Cobrança por WhatsApp). Script 184 (CreateTables_CobrancaWhatsApp.sql).
'   CobrancaConfig: horário, limite por dia, intervalo entre mensagens, chave Pix.
'   CobrancaRegra : régua (antes do vencimento / no dia / após o vencimento), com o modelo da mensagem.
'   CobrancaEnvio + CobrancaEnvioParcela: histórico e controle de repetição.
' Uma mensagem por cliente por dia, com todas as parcelas em aberto dele (a regra é a da parcela mais atrasada).
' Parcelas: tabela parcelas (STATUS 0 = em aberto, DATA = vencimento), cliente pelo pedido;
' saldo = VALOR - parcelas_haver (sem juros). Envio: ZapOnlineEnviar (LicencaOnline.bas).
'---------------------------------------------------------------------------

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const COB_ANTES As Integer = 1
Public Const COB_NO_DIA As Integer = 2
Public Const COB_DEPOIS As Integer = 3
Private Const COB_MAX_LINHAS As Long = 10

'Envio em andamento: a tela liga para interromper a fila entre uma mensagem e outra
Public Cob_Parar As Boolean
'Intervalo mínimo entre mensagens (segundos): abaixo disso o WhatsApp tende a bloquear o número
Public Const COB_INTERVALO_MIN As Long = 30

Public Type CobItem
    CodCliente As Long
    Nome As String
    Celular As String
    CodRegra As Long
    Regra As String
    Mensagem As String       'modelo da regra (com as tags)
    Parcelas As String       '"|cod|cod|" das parcelas que dispararam a regra (vão para o histórico)
    Lista As String          'linhas das parcelas em aberto ({LISTA_PARCELAS}), separadas por vbLf
    QtdParcelas As Long
    Total As Currency
    DiasAtraso As Long
    Vencimento As Date       'vencimento mais antigo da lista
    Motivo As String         '"" = pode enviar; senão, por que não vai
End Type

Public Type CobConfig
    HoraInicio As String
    HoraFim As String
    LimiteDia As Long
    IntervaloMin As Long
    IntervaloMax As Long
    EnviarSabado As Boolean
    EnviarDomingo As Boolean
    ChavePix As String
    Automatico As Boolean
End Type

'Tabelas do script 184 existem?
Public Function Cob_BancoPronto() As Boolean
On Error GoTo Falha
Cob_BancoPronto = (Val(CStr(SQLExecutaRetorno("SELECT CASE WHEN OBJECT_ID('dbo.CobrancaRegra', 'U') IS NULL " & _
    "OR OBJECT_ID('dbo.CobrancaEnvioParcela', 'U') IS NULL OR COL_LENGTH('dbo.cliente', 'NaoCobrarWhatsApp') IS NULL " & _
    "THEN 0 ELSE 1 END AS r", "r", 0))) = 1)
Exit Function
Falha:
Err.Clear
Cob_BancoPronto = False
End Function

Public Sub Cob_LerConfig(ByRef c As CobConfig)
Dim r As ADODB.Recordset

c.HoraInicio = "08:00": c.HoraFim = "18:00"
c.LimiteDia = 50: c.IntervaloMin = 30: c.IntervaloMax = 60
c.EnviarSabado = False: c.EnviarDomingo = False
c.ChavePix = "": c.Automatico = False
On Error GoTo Fim
Set r = dbData.OpenRecordset("SELECT * FROM CobrancaConfig WHERE Id = 1")
If Not r.EOF Then
    c.HoraInicio = Trim$("" & r("HoraInicio"))
    c.HoraFim = Trim$("" & r("HoraFim"))
    c.LimiteDia = Val("" & r("LimiteDia"))
    c.IntervaloMin = Val("" & r("IntervaloMin"))
    c.IntervaloMax = Val("" & r("IntervaloMax"))
    c.EnviarSabado = CBool(r("EnviarSabado"))
    c.EnviarDomingo = CBool(r("EnviarDomingo"))
    c.ChavePix = Trim$("" & r("ChavePix"))
    c.Automatico = CBool(r("Automatico"))
End If
If c.IntervaloMin < COB_INTERVALO_MIN Then c.IntervaloMin = COB_INTERVALO_MIN
If c.IntervaloMax < c.IntervaloMin Then c.IntervaloMax = c.IntervaloMin
r.Close
Fim:
Err.Clear
End Sub

Public Function Cob_SalvarConfig(ByRef c As CobConfig) As Boolean
On Error GoTo Falha
dbData.Execute "UPDATE CobrancaConfig SET HoraInicio = '" & Cob_Hora(c.HoraInicio) & "', HoraFim = '" & Cob_Hora(c.HoraFim) & "', " & _
    "LimiteDia = " & c.LimiteDia & ", IntervaloMin = " & c.IntervaloMin & ", IntervaloMax = " & c.IntervaloMax & ", " & _
    "EnviarSabado = " & IIf(c.EnviarSabado, 1, 0) & ", EnviarDomingo = " & IIf(c.EnviarDomingo, 1, 0) & ", " & _
    "ChavePix = " & IIf(Len(c.ChavePix) = 0, "NULL", Cob_Q(c.ChavePix)) & " WHERE Id = 1"
Cob_SalvarConfig = True
Exit Function
Falha:
Err.Clear
End Function

'"8:5" -> "08:05"; inválida -> "08:00"
Public Function Cob_Hora(ByVal s As String) As String
Dim p As Long, h As Long, m As Long

s = Trim$(s)
p = InStr(s, ":")
If p = 0 Then h = Val(s) Else h = Val(Left$(s, p - 1)): m = Val(Mid$(s, p + 1))
If h < 0 Or h > 23 Or m < 0 Or m > 59 Then h = 8: m = 0
Cob_Hora = Format$(h, "00") & ":" & Format$(m, "00")
End Function

'Cria a régua padrão na primeira vez (o usuário ajusta os textos na tela da régua)
Public Sub Cob_GarantirRegrasPadrao()
Dim L As String, sAberto As String, sFim As String

On Error GoTo Fim
If Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) AS n FROM CobrancaRegra", "n", 0))) > 0 Then Exit Sub

L = vbCrLf
sFim = "Caso o pagamento já tenha sido efetuado, por favor, desconsidere esta mensagem. " & _
       "Se precisar da chave Pix ou de outra forma de pagamento, é só chamar! {CORACAO}" & L & L & "*{EMPRESA}*"

Cob_InserirRegra "Aviso 1 dia antes do vencimento", False, COB_ANTES, 1, 0, 1, _
    "Olá, *{PRIMEIRO_NOME}*! Tudo bem? {ACENO}" & L & L & _
    "Passando para lembrar que amanhã vence:" & L & "{LISTA_PARCELAS}" & L & L & _
    "Total: *R$ {VALOR_TOTAL}* {DOC}" & L & L & sFim
Cob_InserirRegra "No dia do vencimento", True, COB_NO_DIA, 0, 0, 1, _
    "Olá, *{PRIMEIRO_NOME}*! Tudo bem? {ACENO}" & L & L & _
    "Passando apenas para lembrar que vence hoje:" & L & "{LISTA_PARCELAS}" & L & L & _
    "Total: *R$ {VALOR_TOTAL}* {DOC}" & L & L & sFim
sAberto = "Olá, *{PRIMEIRO_NOME}*! Tudo bem? {ACENO}" & L & L & _
    "Passando apenas para lembrar que consta em aberto:" & L & "{LISTA_PARCELAS}" & L & L & _
    "Total: *R$ {VALOR_TOTAL}* {DOC}" & L & L & sFim
Cob_InserirRegra "3 dias após o vencimento", True, COB_DEPOIS, 3, 0, 1, sAberto
Cob_InserirRegra "10 dias após o vencimento (repete a cada 7 dias)", True, COB_DEPOIS, 10, 7, 3, _
    "Olá, *{NOME_CLIENTE}*." & L & L & _
    "Não identificamos o pagamento de:" & L & "{LISTA_PARCELAS}" & L & L & _
    "Total em aberto: *R$ {VALOR_TOTAL}* ({DIAS_ATRASO} dias em atraso)." & L & L & _
    "Pedimos que regularize o quanto antes ou fale com a gente para combinarmos a melhor forma de pagamento. " & _
    "Se o pagamento já foi feito, por favor, desconsidere esta mensagem." & L & L & "*{EMPRESA}*"
Fim:
Err.Clear
End Sub

Private Sub Cob_InserirRegra(ByVal sDesc As String, ByVal bAtiva As Boolean, ByVal nSit As Integer, ByVal nDias As Long, _
                             ByVal nRepetir As Long, ByVal nMax As Long, ByVal sMsg As String)
dbData.Execute "INSERT INTO CobrancaRegra (Descricao, Ativa, Situacao, Dias, RepetirDias, MaxEnvios, Mensagem) VALUES (" & _
    Cob_Q(sDesc) & ", " & IIf(bAtiva, 1, 0) & ", " & nSit & ", " & nDias & ", " & nRepetir & ", " & nMax & ", " & Cob_Q(sMsg) & ")"
End Sub

'Texto da situação para as telas
Public Function Cob_SituacaoTexto(ByVal nSit As Integer, ByVal nDias As Long) As String
Select Case nSit
    Case COB_ANTES: Cob_SituacaoTexto = nDias & " dia(s) antes do vencimento"
    Case COB_NO_DIA: Cob_SituacaoTexto = "No dia do vencimento"
    Case Else: Cob_SituacaoTexto = nDias & " dia(s) após o vencimento"
End Select
End Function

'---------------------------------------------------------------------------
' Quem recebe cobrança em dHoje. Devolve a quantidade (Itens(1..n)); inclui os que não podem receber, com o Motivo.
' bJuros = soma multa + juros do atraso (Configuração Geral: MULTA, JUROS_DIA e TIPO_JUROS) no valor de cada parcela vencida.
'---------------------------------------------------------------------------
Public Function Cob_Pendentes(ByVal dHoje As Date, ByRef Itens() As CobItem, Optional ByVal bJuros As Boolean = False) As Long
Dim rs As ADODB.Recordset
Dim nR As Long, i As Long, n As Long
Dim rId() As Long, rDesc() As String, rSit() As Integer, rDias() As Long, rRep() As Long, rMax() As Long, rMsg() As String
Dim maxAntes As Long
Dim hist As New Collection, hoje As New Collection
Dim v As Variant, sHoje As String
Dim cliAtual As Long, cli As Long
Dim dias As Long, ir As Long, melhor As Long, pesoMelhor As Long, peso As Long
Dim dVenc As Date, saldo As Currency, codParc As Long, bDispara As Boolean
Dim linhas As Long, extras As Long
Dim it As CobItem, naoCobrar As Boolean, sMotivoCel As String
Dim cJurosDia As Currency, nTipoJuros As Integer, cValor As Currency, cMulta As Currency, cBase As Currency

dHoje = DateValue(dHoje)
sHoje = Format$(dHoje, "yyyymmdd")
ReDim Itens(0)
If bJuros Then Cob_JurosConfig cJurosDia, nTipoJuros, cMulta
On Error GoTo Falha

'Regras ativas
Set rs = dbData.OpenRecordset("SELECT Id, Descricao, Situacao, Dias, RepetirDias, MaxEnvios, Mensagem FROM CobrancaRegra WHERE Ativa = 1 ORDER BY Id")
Do While Not rs.EOF
    nR = nR + 1
    ReDim Preserve rId(1 To nR), rDesc(1 To nR), rSit(1 To nR), rDias(1 To nR), rRep(1 To nR), rMax(1 To nR), rMsg(1 To nR)
    rId(nR) = rs("Id"): rDesc(nR) = "" & rs("Descricao"): rSit(nR) = rs("Situacao")
    rDias(nR) = Val("" & rs("Dias")): rRep(nR) = Val("" & rs("RepetirDias")): rMax(nR) = Val("" & rs("MaxEnvios"))
    If rMax(nR) < 1 Then rMax(nR) = 1
    rMsg(nR) = "" & rs("Mensagem")
    If rSit(nR) = COB_ANTES And rDias(nR) > maxAntes Then maxAntes = rDias(nR)
    rs.MoveNext
Loop
rs.Close
If nR = 0 Then Exit Function

'Envios já feitos por parcela e regra (só das parcelas ainda em aberto): "quantidade|último dia"
Set rs = dbData.OpenRecordset("SELECT ep.CodParcela, ep.CodRegra, COUNT(*) AS n, MAX(e.DataHora) AS ult " & _
    "FROM CobrancaEnvioParcela ep INNER JOIN CobrancaEnvio e ON e.Id = ep.CodEnvio " & _
    "WHERE e.Situacao = 1 AND ep.CodParcela IN (SELECT CODIGO FROM parcelas WHERE STATUS = 0) " & _
    "GROUP BY ep.CodParcela, ep.CodRegra")
Do While Not rs.EOF
    hist.Add rs("n") & "|" & Format$(rs("ult"), "yyyymmdd"), "p" & rs("CodParcela") & "r" & rs("CodRegra")
    rs.MoveNext
Loop
rs.Close

'Quem já recebeu cobrança hoje (uma mensagem por cliente por dia)
Set rs = dbData.OpenRecordset("SELECT DISTINCT CodCliente FROM CobrancaEnvio WHERE Situacao = 1 AND DataHora >= '" & sHoje & "'")
Do While Not rs.EOF
    hoje.Add 1, "c" & rs("CodCliente")
    rs.MoveNext
Loop
rs.Close

'Parcelas em aberto até o maior "dias antes" das regras
Set rs = dbData.OpenRecordset("SELECT p.CODIGO, p.COD_PEDIDO, ISNULL(p.COD_OS, 0) AS COD_OS, ISNULL(p.NUMERO, 0) AS NUMERO, p.DATA, " & _
    "(SELECT COUNT(*) FROM parcelas x WHERE x.COD_PEDIDO = p.COD_PEDIDO) AS QTD, " & _
    "p.VALOR AS VALOR_PARC, p.VALOR - ISNULL((SELECT SUM(h.valor_haver) FROM parcelas_haver h WHERE h.cod_parcela = p.CODIGO), 0) AS SALDO, " & _
    "c.CODIGO AS COD_CLIENTE, c.NOME, ISNULL(c.NaoCobrarWhatsApp, 0) AS NAO_COBRAR " & _
    "FROM parcelas p INNER JOIN pedidos pe ON pe.COD_PEDIDO = p.COD_PEDIDO INNER JOIN cliente c ON c.CODIGO = pe.COD_CLIENTE " & _
    "WHERE p.STATUS = 0 AND p.DATA < '" & Format$(dHoje + maxAntes + 1, "yyyymmdd") & "' " & _
    "ORDER BY c.CODIGO, p.DATA, p.CODIGO")

cliAtual = -1
Do While Not rs.EOF
    cli = rs("COD_CLIENTE")
    If cli <> cliAtual Then
        If cliAtual <> -1 Then GoSub Fechar
        cliAtual = cli
        it.CodCliente = cli
        it.Nome = Trim$("" & rs("NOME"))
        it.Celular = "": it.CodRegra = 0: it.Regra = "": it.Mensagem = "": it.Parcelas = "|"
        it.Lista = "": it.QtdParcelas = 0: it.Total = 0: it.DiasAtraso = 0: it.Vencimento = 0: it.Motivo = ""
        naoCobrar = CBool(rs("NAO_COBRAR"))
        pesoMelhor = -1: melhor = 0: linhas = 0: extras = 0
    End If

    If IsNull(rs("SALDO")) Then saldo = 0 Else saldo = CCur(rs("SALDO"))
    If saldo > 0.004 Then
        codParc = rs("CODIGO")
        dVenc = DateValue(rs("DATA"))
        dias = DateDiff("d", dVenc, dHoje)

        'Regra da parcela: após o vencimento = a de maior "dias" já alcançada; no dia; antes = dias exatos
        ir = 0
        If dias >= 1 Then
            For i = 1 To nR
                If rSit(i) = COB_DEPOIS And rDias(i) >= 1 And rDias(i) <= dias Then
                    If ir = 0 Then
                        ir = i
                    ElseIf rDias(i) > rDias(ir) Then
                        ir = i
                    End If
                End If
            Next
        ElseIf dias = 0 Then
            For i = 1 To nR
                If rSit(i) = COB_NO_DIA Then ir = i: Exit For
            Next
        Else
            For i = 1 To nR
                If rSit(i) = COB_ANTES And rDias(i) = -dias Then ir = i: Exit For
            Next
        End If

        bDispara = False
        If ir > 0 Then
            v = Empty
            On Error Resume Next
            v = hist("p" & codParc & "r" & rId(ir))
            On Error GoTo Falha
            If IsEmpty(v) Then
                bDispara = True
            ElseIf rRep(ir) > 0 Then
                bDispara = (Val(Split(v, "|")(0)) < rMax(ir)) And _
                           (DateDiff("d", DateSerial(Left$(Split(v, "|")(1), 4), Mid$(Split(v, "|")(1), 5, 2), Right$(Split(v, "|")(1), 2)), dHoje) >= rRep(ir))
            End If
        End If

        If bDispara Then
            it.Parcelas = it.Parcelas & codParc & "|"
            Select Case rSit(ir)
                Case COB_DEPOIS: peso = 2000 + rDias(ir)
                Case COB_NO_DIA: peso = 1000
                Case Else: peso = 500 - rDias(ir)
            End Select
            If peso > pesoMelhor Then pesoMelhor = peso: melhor = ir
        End If

        'Na lista: tudo o que já venceu (ou vence hoje) e as parcelas que dispararam o aviso antes do vencimento
        'Valor cobrado: saldo + multa (uma vez) + juros ao dia, sobre o saldo ou sobre o valor da parcela (conforme a configuração)
        cValor = saldo
        If bJuros And dias > 0 And (cJurosDia > 0 Or cMulta > 0) Then
            If nTipoJuros = 1 Then cBase = CCur(rs("VALOR_PARC")) Else cBase = saldo
            cValor = saldo + Round(cBase * (cMulta + cJurosDia * dias) / 100, 2)
        End If

        If dias >= 0 Or bDispara Then
            it.QtdParcelas = it.QtdParcelas + 1
            it.Total = it.Total + cValor
            If dias > it.DiasAtraso Then it.DiasAtraso = dias
            If it.Vencimento = 0 Or dVenc < it.Vencimento Then it.Vencimento = dVenc
            If linhas < COB_MAX_LINHAS Then
                it.Lista = it.Lista & IIf(Len(it.Lista) > 0, vbLf, "") & Cob_LinhaParcela(rs("COD_PEDIDO"), rs("COD_OS"), rs("NUMERO"), rs("QTD"), dVenc, cValor, dias)
                linhas = linhas + 1
            Else
                extras = extras + 1
            End If
        End If
    End If
    rs.MoveNext
Loop
rs.Close
If cliAtual <> -1 Then GoSub Fechar
Cob_Pendentes = n
Exit Function

Fechar:
    If melhor > 0 Then
        If extras > 0 Then it.Lista = it.Lista & vbLf & "• e mais " & extras & " parcela(s)"
        it.CodRegra = rId(melhor)
        it.Regra = rDesc(melhor)
        it.Mensagem = rMsg(melhor)
        it.Celular = PDFZap_CelularCliente(it.CodCliente)
        If naoCobrar Then
            it.Motivo = "Cliente marcado para não receber cobrança"
        ElseIf Cob_Existe(hoje, "c" & it.CodCliente) Then
            it.Motivo = "Já recebeu cobrança hoje"
        ElseIf Len(it.Celular) = 0 Then
            it.Motivo = "Sem celular no cadastro"
        ElseIf Not CelularValido(it.Celular, sMotivoCel) Then
            it.Motivo = "Celular inválido (" & sMotivoCel & ")"
        End If
        n = n + 1
        ReDim Preserve Itens(n)
        Itens(n) = it
    End If
    Return

Falha:
MsgBox "Não foi possível montar a lista de cobranças." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Cobrança por WhatsApp"
Err.Clear
Cob_Pendentes = n
End Function

'Juros ao dia (%), tipo (0 = sobre o saldo restante, 1 = sobre o valor da parcela) e multa (%) da Configuração Geral
Public Sub Cob_JurosConfig(ByRef cJurosDia As Currency, ByRef nTipo As Integer, Optional ByRef cMulta As Currency)
Dim oCfg As ConfigItem

cJurosDia = 0: nTipo = 0: cMulta = 0
On Error Resume Next
Set oCfg = sysConfig("JUROS_DIA")
cJurosDia = CCur(oCfg.Value)
Set oCfg = sysConfig("TIPO_JUROS")
nTipo = Val(CStr(oCfg.Value))
Set oCfg = Nothing
Set oCfg = sysConfig("MULTA")
If Not oCfg Is Nothing Then cMulta = CCur(oCfg.Value)
Set oCfg = Nothing
If cJurosDia < 0 Then cJurosDia = 0
If cMulta < 0 Then cMulta = 0
Err.Clear
End Sub

Private Function Cob_Existe(ByVal c As Collection, ByVal chave As String) As Boolean
Dim v As Variant
On Error Resume Next
v = c(chave)
Cob_Existe = (Err.Number = 0)
Err.Clear
End Function

'"• *Venda nº 1381* (parcela 1/3) – venc. 10/11/2026 – R$ 150,00 – 5 dia(s) em atraso"
Private Function Cob_LinhaParcela(ByVal lPedido As Long, ByVal lOS As Long, ByVal nNumero As Long, ByVal nQtd As Long, _
                                  ByVal dVenc As Date, ByVal cSaldo As Currency, ByVal nDias As Long) As String
Dim s As String

If lOS > 0 Then s = "*OS nº " & lOS & "*" Else s = "*Venda nº " & lPedido & "*"
If nQtd > 1 And nNumero > 0 Then s = s & " (parcela " & nNumero & "/" & nQtd & ")"
s = "• " & s & " – venc. " & Format$(dVenc, "dd/mm/yyyy") & " – R$ " & Format$(cSaldo, "#,##0.00")
If nDias > 0 Then s = s & " – " & nDias & " dia(s) em atraso"
Cob_LinhaParcela = s
End Function

'---------------------------------------------------------------------------
' Mensagem: troca as tags. bPrevia = para a tela (o TextBox do VB6 não mostra emoji: vira [nome]).
'---------------------------------------------------------------------------
Public Function Cob_Texto(ByRef it As CobItem, ByRef cfg As CobConfig, ByVal bPrevia As Boolean) As String
Dim s As String, i As Long, tags() As String, desc() As String, emo() As String

s = Replace(it.Mensagem, vbCrLf, vbLf)
s = Replace(s, "{PRIMEIRO_NOME}", Cob_PrimeiroNome(it.Nome), , , vbTextCompare)
s = Replace(s, "{NOME_CLIENTE}", StrConv(it.Nome, vbProperCase), , , vbTextCompare)
s = Replace(s, "{EMPRESA}", PDFZap_NomeEmpresa(), , , vbTextCompare)
s = Replace(s, "{LISTA_PARCELAS}", it.Lista, , , vbTextCompare)
s = Replace(s, "{VALOR_TOTAL}", Format$(it.Total, "#,##0.00"), , , vbTextCompare)
s = Replace(s, "{QTD_PARCELAS}", CStr(it.QtdParcelas), , , vbTextCompare)
s = Replace(s, "{DIAS_ATRASO}", CStr(it.DiasAtraso), , , vbTextCompare)
s = Replace(s, "{VENCIMENTO}", IIf(it.Vencimento > 0, Format$(it.Vencimento, "dd/mm/yyyy"), ""), , , vbTextCompare)
s = Replace(s, "{CHAVE_PIX}", cfg.ChavePix, , , vbTextCompare)

Cob_ListaTags tags, desc, emo
For i = 0 To UBound(tags)
    If Len(emo(i)) > 0 Then
        If bPrevia Then
            s = Replace(s, tags(i), "[" & LCase$(Mid$(tags(i), 2, Len(tags(i)) - 2)) & "]", , , vbTextCompare)
        Else
            s = Replace(s, tags(i), Cob_Emoji(emo(i)), , , vbTextCompare)
        End If
    End If
Next
If bPrevia Then s = Replace(s, vbLf, vbCrLf)
Cob_Texto = s
End Function

'Tags para inserir no modelo: tags(i), descrição para a lista, e o emoji (códigos hex separados por "+"; "" = dado)
Public Sub Cob_ListaTags(ByRef tags() As String, ByRef desc() As String, ByRef emo() As String)
Const LISTA As String = _
    "{PRIMEIRO_NOME}|Primeiro nome do cliente|;" & _
    "{NOME_CLIENTE}|Nome completo do cliente|;" & _
    "{LISTA_PARCELAS}|Lista das parcelas em aberto (venda/OS, vencimento, valor)|;" & _
    "{VALOR_TOTAL}|Total em aberto|;" & _
    "{QTD_PARCELAS}|Quantidade de parcelas|;" & _
    "{DIAS_ATRASO}|Dias em atraso (da parcela mais antiga)|;" & _
    "{VENCIMENTO}|Vencimento mais antigo|;" & _
    "{CHAVE_PIX}|Chave Pix da empresa (configuração)|;" & _
    "{EMPRESA}|Nome da empresa|;" & _
    "{ACENO}|Emoji: mão acenando|1F44B;" & _
    "{SORRISO}|Emoji: sorriso|1F60A;" & _
    "{DOC}|Emoji: documento|1F4C4;" & _
    "{CALENDARIO}|Emoji: calendário|1F4C5;" & _
    "{RELOGIO}|Emoji: relógio|23F0;" & _
    "{DINHEIRO}|Emoji: saco de dinheiro|1F4B0;" & _
    "{OK}|Emoji: confirmado|2705;" & _
    "{ATENCAO}|Emoji: atenção|26A0+FE0F;" & _
    "{CORACAO}|Emoji: coração laranja|1F9E1"
Dim lst() As String, p() As String, i As Long

lst = Split(LISTA, ";")
ReDim tags(UBound(lst)): ReDim desc(UBound(lst)): ReDim emo(UBound(lst))
For i = 0 To UBound(lst)
    p = Split(lst(i), "|")
    tags(i) = p(0): desc(i) = p(1): emo(i) = p(2)
Next
End Sub

'"1F44B" / "26A0+FE0F" -> caracteres (acima de FFFF vira par UTF-16, pelo PDFZap_Emoji)
Private Function Cob_Emoji(ByVal sCodigos As String) As String
Dim p() As String, i As Long, n As Long

p = Split(sCodigos, "+")
For i = 0 To UBound(p)
    n = CLng("&H" & p(i))
    If n > &HFFFF& Then
        Cob_Emoji = Cob_Emoji & PDFZap_Emoji(n)
    Else
        Cob_Emoji = Cob_Emoji & ChrW$(n)
    End If
Next
End Function

Private Function Cob_PrimeiroNome(ByVal s As String) As String
s = Trim$(s)
If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
Cob_PrimeiroNome = StrConv(s, vbProperCase)
End Function

'---------------------------------------------------------------------------
' Envio e histórico
'---------------------------------------------------------------------------
Public Function Cob_Enviar(ByRef it As CobItem, ByRef cfg As CobConfig, ByVal bManual As Boolean, ByRef sErro As String) As Boolean
sErro = ""
Cob_Enviar = ZapOnlineEnviar(it.Celular, Cob_Texto(it, cfg, False), "", sErro)
Cob_Registrar it, Cob_Enviar, sErro, bManual
End Function

Private Sub Cob_Registrar(ByRef it As CobItem, ByVal bOk As Boolean, ByVal sErro As String, ByVal bManual As Boolean)
Dim r As ADODB.Recordset, lId As Long, p() As String, i As Long, sUsuario As String

On Error GoTo Fim
sUsuario = ""
On Error Resume Next
sUsuario = Trim$(Tela_Principal.txtCodFuncionario.Text)
On Error GoTo Fim
Set r = dbData.OpenRecordset("INSERT INTO CobrancaEnvio (CodCliente, CodRegra, Celular, ValorTotal, Situacao, Erro, Manual, Usuario) " & _
    "OUTPUT INSERTED.Id VALUES (" & it.CodCliente & ", " & it.CodRegra & ", '" & Replace(Left$(it.Celular, 20), "'", "") & "', " & _
    FSQL(it.Total, 2) & ", " & IIf(bOk, 1, 2) & ", " & IIf(bOk, "NULL", Cob_Q(Left$(sErro, 300))) & ", " & IIf(bManual, 1, 0) & ", " & _
    Cob_Q(Left$(sUsuario, 50)) & ")")
If Not r.EOF Then lId = r(0)
r.Close
If Not bOk Or lId = 0 Then Exit Sub
p = Split(it.Parcelas, "|")
For i = 0 To UBound(p)
    If Val(p(i)) > 0 Then
        dbData.Execute "INSERT INTO CobrancaEnvioParcela (CodEnvio, CodParcela, CodRegra) VALUES (" & lId & ", " & Val(p(i)) & ", " & it.CodRegra & ")"
    End If
Next
Exit Sub
Fim:
Err.Clear
End Sub

Public Function Cob_EnviadosHoje() As Long
Cob_EnviadosHoje = Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) AS n FROM CobrancaEnvio WHERE Situacao = 1 AND DataHora >= '" & _
    Format$(Date, "yyyymmdd") & "'", "n", 0)))
End Function

'Marca/desmarca o cliente para não receber cobrança pelo WhatsApp
Public Function Cob_NaoCobrar(ByVal lCodCliente As Long, ByVal bNaoCobrar As Boolean) As Boolean
On Error GoTo Falha
dbData.Execute "UPDATE cliente SET NaoCobrarWhatsApp = " & IIf(bNaoCobrar, 1, 0) & " WHERE CODIGO = " & lCodCliente
Cob_NaoCobrar = True
Exit Function
Falha:
Err.Clear
End Function

'Espera nSegundos entre uma mensagem e outra, sem travar a tela. False se o usuário mandou parar.
Public Function Cob_Esperar(ByVal nSegundos As Long) As Boolean
Dim fim As Date

fim = DateAdd("s", nSegundos, Now)
Do
    If Cob_Parar Then Exit Function
    DoEvents
    Sleep 50
Loop While Now < fim
Cob_Esperar = Not Cob_Parar
End Function

'Intervalo aleatório entre o mínimo e o máximo da configuração
Public Function Cob_Intervalo(ByRef cfg As CobConfig) As Long
Dim a As Long, b As Long

a = cfg.IntervaloMin: b = cfg.IntervaloMax
If a < COB_INTERVALO_MIN Then a = COB_INTERVALO_MIN
If b < a Then b = a
Randomize
Cob_Intervalo = a + Int(Rnd * (b - a + 1))
End Function

Public Function Cob_Q(ByVal s As String) As String
Cob_Q = "N'" & Replace(s, "'", "''") & "'"
End Function
