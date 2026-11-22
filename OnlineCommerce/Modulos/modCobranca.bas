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
'Máximo de parcelas a vencer listadas na mensagem (regra com "listar as parcelas a vencer")
Private Const COB_MAX_FUTURAS As Long = 6

'Filtro de Cob_Pendentes: HOJE = só quem tem mensagem para sair hoje (envio automático);
'A_VENCER / VENCENDO / VENCIDO = todos os clientes daquela situação (a tela), com a previsão do próximo envio
Public Const COB_FILTRO_HOJE As Integer = 0
Public Const COB_FILTRO_A_VENCER As Integer = 1
Public Const COB_FILTRO_VENCENDO As Integer = 2
Public Const COB_FILTRO_VENCIDO As Integer = 3
'"A vencer" mostra as parcelas que vencem até esta quantidade de dias (ou até o maior "dias antes" da régua, se for maior)
Public Const COB_DIAS_A_VENCER As Long = 7

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
' nFiltro: COB_FILTRO_HOJE (só quem recebe hoje) ou A_VENCER / VENCENDO / VENCIDO (o cliente entra pela parcela mais
' atrasada dele; quem não tem envio hoje vem com CodRegra = 0, a regra prevista e o Motivo "Sem envio hoje (próximo em ...)").
' bJuros = soma multa + juros do atraso (Configuração Geral: MULTA, JUROS_DIA e TIPO_JUROS) no valor de cada parcela vencida.
'---------------------------------------------------------------------------
Public Function Cob_Pendentes(ByVal dHoje As Date, ByRef Itens() As CobItem, Optional ByVal bJuros As Boolean = False, _
                              Optional ByVal nFiltro As Integer = COB_FILTRO_HOJE) As Long
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
Dim it As CobItem, naoCobrar As Boolean, sMotivoCel As String, sLinha As String
Dim cJurosDia As Currency, nTipoJuros As Integer, cValor As Currency, cMulta As Currency, cBase As Currency
Dim nJanela As Long, nCat As Integer, dProx As Date, iProx As Long, dP As Date, vH As Variant, bListar As Boolean
Dim qF As Long, tF As Currency, vF As Date, linF As String, linhasF As Long, extrasF As Long
Dim gAtr As String, gHoje As String, gBreve As String, rListar() As Boolean, rr As Long

dHoje = DateValue(dHoje)
sHoje = Format$(dHoje, "yyyymmdd")
ReDim Itens(0)
If bJuros Then Cob_JurosConfig cJurosDia, nTipoJuros, cMulta
On Error GoTo Falha

'Regras ativas
Set rs = dbData.OpenRecordset("SELECT Id, Descricao, Situacao, Dias, RepetirDias, MaxEnvios, Mensagem, " & _
    IIf(Cob_TemListarAVencer(), "ListarAVencer", "1 AS ListarAVencer") & " FROM CobrancaRegra WHERE Ativa = 1 ORDER BY Id")
Do While Not rs.EOF
    nR = nR + 1
    ReDim Preserve rId(1 To nR), rDesc(1 To nR), rSit(1 To nR), rDias(1 To nR), rRep(1 To nR), rMax(1 To nR), rMsg(1 To nR), rListar(1 To nR)
    rId(nR) = rs("Id"): rDesc(nR) = "" & rs("Descricao"): rSit(nR) = rs("Situacao")
    rDias(nR) = Val("" & rs("Dias")): rRep(nR) = Val("" & rs("RepetirDias")): rMax(nR) = Val("" & rs("MaxEnvios"))
    If rMax(nR) < 1 Then rMax(nR) = 1
    rMsg(nR) = "" & rs("Mensagem")
    rListar(nR) = CBool(rs("ListarAVencer"))
    If rSit(nR) = COB_ANTES And rDias(nR) > maxAntes Then maxAntes = rDias(nR)
    rs.MoveNext
Loop
rs.Close
If nR = 0 Then Exit Function
nJanela = maxAntes
If nFiltro <> COB_FILTRO_HOJE And nJanela < COB_DIAS_A_VENCER Then nJanela = COB_DIAS_A_VENCER

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

'Todas as parcelas em aberto (as que ainda vão vencer entram na lista "Próximas parcelas" quando a regra pede)
Set rs = dbData.OpenRecordset("SELECT p.CODIGO, p.COD_PEDIDO, ISNULL(p.COD_OS, 0) AS COD_OS, ISNULL(p.NUMERO, 0) AS NUMERO, p.DATA, " & _
    "(SELECT COUNT(*) FROM parcelas x WHERE x.COD_PEDIDO = p.COD_PEDIDO) AS QTD, " & _
    "p.VALOR AS VALOR_PARC, p.VALOR - ISNULL((SELECT SUM(h.valor_haver) FROM parcelas_haver h WHERE h.cod_parcela = p.CODIGO), 0) AS SALDO, " & _
    "c.CODIGO AS COD_CLIENTE, c.NOME, ISNULL(c.NaoCobrarWhatsApp, 0) AS NAO_COBRAR " & _
    "FROM parcelas p INNER JOIN pedidos pe ON pe.COD_PEDIDO = p.COD_PEDIDO INNER JOIN cliente c ON c.CODIGO = pe.COD_CLIENTE " & _
    "WHERE p.STATUS = 0 " & _
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
        nCat = 0: dProx = 0: iProx = 0: qF = 0: tF = 0: vF = 0: linF = "": linhasF = 0: extrasF = 0
        gAtr = "": gHoje = "": gBreve = ""
    End If

    If IsNull(rs("SALDO")) Then saldo = 0 Else saldo = CCur(rs("SALDO"))
    If saldo > 0.004 Then
        codParc = rs("CODIGO")
        dVenc = DateValue(rs("DATA"))
        dias = DateDiff("d", dVenc, dHoje)

        'Situação do cliente = a da parcela mais atrasada (vencido > vencendo hoje > a vencer)
        If dias >= 1 Then
            nCat = COB_FILTRO_VENCIDO
        ElseIf dias = 0 Then
            If nCat < COB_FILTRO_VENCENDO Then nCat = COB_FILTRO_VENCENDO
        ElseIf -dias <= nJanela Then
            If nCat < COB_FILTRO_A_VENCER Then nCat = COB_FILTRO_A_VENCER
        End If

        'Próximo envio previsto desta parcela (depois de hoje), só para a tela
        If nFiltro <> COB_FILTRO_HOJE And -dias <= nJanela Then
            For i = 1 To nR
                dP = 0
                vH = Empty
                On Error Resume Next
                vH = hist("p" & codParc & "r" & rId(i))
                On Error GoTo Falha
                Select Case rSit(i)
                    Case COB_ANTES: If IsEmpty(vH) Then dP = dVenc - rDias(i)
                    Case COB_NO_DIA: If IsEmpty(vH) Then dP = dVenc
                    Case Else
                        If IsEmpty(vH) Then
                            dP = dVenc + rDias(i)
                        ElseIf rRep(i) > 0 Then
                            If Val(Split(vH, "|")(0)) < rMax(i) Then dP = Cob_DataHist(vH) + rRep(i)
                        End If
                End Select
                If dP > dHoje Then
                    If dProx = 0 Or dP < dProx Then dProx = dP: iProx = i
                End If
            Next
        End If

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
                           (DateDiff("d", Cob_DataHist(v), dHoje) >= rRep(ir))
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
                sLinha = Cob_LinhaParcela(rs("COD_PEDIDO"), rs("COD_OS"), rs("NUMERO"), rs("QTD"), dVenc, cValor, dias)
                If dias > 0 Then
                    gAtr = gAtr & IIf(Len(gAtr) > 0, vbLf, "") & sLinha
                ElseIf dias = 0 Then
                    gHoje = gHoje & IIf(Len(gHoje) > 0, vbLf, "") & sLinha
                Else
                    gBreve = gBreve & IIf(Len(gBreve) > 0, vbLf, "") & sLinha
                End If
                linhas = linhas + 1
            Else
                extras = extras + 1
            End If
        Else
            'Ainda vai vencer, sem aviso hoje: lista "Próximas parcelas" da mensagem; as da janela "A vencer" contam
            'para a tela quando o cliente não tem nada que já venceu
            If -dias <= nJanela Then
                qF = qF + 1
                tF = tF + cValor
                If vF = 0 Or dVenc < vF Then vF = dVenc
            End If
            If linhasF < COB_MAX_FUTURAS Then
                linF = linF & IIf(Len(linF) > 0, vbLf, "") & Cob_LinhaParcela(rs("COD_PEDIDO"), rs("COD_OS"), rs("NUMERO"), rs("QTD"), dVenc, cValor, dias)
                linhasF = linhasF + 1
            Else
                extrasF = extrasF + 1
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
    If nFiltro = COB_FILTRO_HOJE Then
        bListar = (melhor > 0)
    Else
        bListar = (nCat = nFiltro)
    End If
    If bListar Then
        'Regra que vale para o texto: a de hoje ou, sem envio hoje, a prevista
        If melhor > 0 Then rr = melhor Else rr = iProx
        If it.QtdParcelas = 0 Then
            'Só parcelas a vencer, sem aviso hoje
            it.QtdParcelas = qF: it.Total = tF: it.Vencimento = vF
            it.Lista = Cob_MontarLista("", "", linF, "", 0, extrasF)
        ElseIf rr > 0 Then
            If rListar(rr) Then
                it.Lista = Cob_MontarLista(gAtr, gHoje, gBreve, linF, extras, extrasF)
            Else
                it.Lista = Cob_MontarLista(gAtr, gHoje, gBreve, "", extras, 0)
            End If
        Else
            it.Lista = Cob_MontarLista(gAtr, gHoje, gBreve, "", extras, 0)
        End If
        If melhor > 0 Then
            it.CodRegra = rId(melhor)
            it.Regra = rDesc(melhor)
            it.Mensagem = rMsg(melhor)
        ElseIf iProx > 0 Then
            it.Regra = rDesc(iProx)      'regra prevista (CodRegra fica 0: não envia hoje)
            it.Mensagem = rMsg(iProx)
        End If
        it.Celular = PDFZap_CelularCliente(it.CodCliente)
        If naoCobrar Then
            it.Motivo = "Cliente marcado para não receber cobrança"
        ElseIf Cob_Existe(hoje, "c" & it.CodCliente) Then
            it.Motivo = "Já recebeu cobrança hoje"
        ElseIf Len(it.Celular) = 0 Then
            it.Motivo = "Sem celular no cadastro"
        ElseIf Not CelularValido(it.Celular, sMotivoCel) Then
            it.Motivo = "Celular inválido (" & sMotivoCel & ")"
        ElseIf melhor = 0 Then
            If iProx > 0 Then
                it.Motivo = "Sem envio hoje (próximo em " & Format$(dProx, "dd/mm/yyyy") & ")"
            Else
                it.Motivo = "Sem envio previsto (nenhuma regra ativa alcança este cliente)"
            End If
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

'Coluna CobrancaRegra.ListarAVencer (script 189) existe?
Public Function Cob_TemListarAVencer() As Boolean
Static nTem As Integer   '0 = não verificado, 1 = sim, 2 = não

If nTem = 0 Then
    On Error Resume Next
    nTem = 2
    If Val(CStr(SQLExecutaRetorno("SELECT CASE WHEN COL_LENGTH('dbo.CobrancaRegra', 'ListarAVencer') IS NULL THEN 0 ELSE 1 END AS r", "r", 0))) = 1 Then nTem = 1
    Err.Clear
End If
Cob_TemListarAVencer = (nTem = 1)
End Function

'{LISTA_PARCELAS}: com um grupo só sai a lista simples (como antes); com mais de um, separa em
'"Em atraso", "Vence hoje", "Vence em breve" e "Próximas parcelas (a vencer)", para o cliente não achar
'que pagando o atrasado quitou tudo. Linhas separadas por vbLf; os emojis vão como tags (Cob_Texto converte).
Public Function Cob_MontarLista(ByVal sAtraso As String, ByVal sHoje As String, ByVal sBreve As String, ByVal sFuturas As String, _
                                ByVal nExtras As Long, ByVal nExtrasFuturas As Long) As String
Dim nGrupos As Long, s As String

If nExtras > 0 Then
    'Parcelas além do limite: entram no último grupo com conteúdo (as listas vêm em ordem de vencimento)
    If Len(sBreve) > 0 Then
        sBreve = sBreve & vbLf & "• e mais " & nExtras & " parcela(s)"
    ElseIf Len(sHoje) > 0 Then
        sHoje = sHoje & vbLf & "• e mais " & nExtras & " parcela(s)"
    Else
        sAtraso = sAtraso & vbLf & "• e mais " & nExtras & " parcela(s)"
    End If
End If
If nExtrasFuturas > 0 And Len(sFuturas) > 0 Then sFuturas = sFuturas & vbLf & "• e mais " & nExtrasFuturas & " parcela(s) a vencer"

nGrupos = IIf(Len(sAtraso) > 0, 1, 0) + IIf(Len(sHoje) > 0, 1, 0) + IIf(Len(sBreve) > 0, 1, 0) + IIf(Len(sFuturas) > 0, 1, 0)
If nGrupos <= 1 Then
    Cob_MontarLista = sAtraso & sHoje & sBreve & sFuturas
    Exit Function
End If
If Len(sAtraso) > 0 Then s = s & IIf(Len(s) > 0, vbLf & vbLf, "") & "{VERMELHO} *Em atraso:*" & vbLf & sAtraso
If Len(sHoje) > 0 Then s = s & IIf(Len(s) > 0, vbLf & vbLf, "") & "{RELOGIO} *Vence hoje:*" & vbLf & sHoje
If Len(sBreve) > 0 Then s = s & IIf(Len(s) > 0, vbLf & vbLf, "") & "{RELOGIO} *Vence em breve:*" & vbLf & sBreve
If Len(sFuturas) > 0 Then s = s & IIf(Len(s) > 0, vbLf & vbLf, "") & "{CALENDARIO} *Próximas parcelas (a vencer):*" & vbLf & sFuturas
Cob_MontarLista = s
End Function

'Data gravada no controle de envios ("quantidade|yyyymmdd")
Private Function Cob_DataHist(ByVal v As Variant) As Date
Dim s As String

s = Split(v, "|")(1)
Cob_DataHist = DateSerial(Left$(s, 4), Mid$(s, 5, 2), Right$(s, 2))
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
    "{VERMELHO}|Emoji: círculo vermelho|1F534;" & _
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
