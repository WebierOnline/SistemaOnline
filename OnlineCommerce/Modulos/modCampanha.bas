Attribute VB_Name = "modCampanha"
Option Explicit
'---------------------------------------------------------------------------
' Promoções e avisos por WhatsApp (Financeiro > Promoções por WhatsApp). Script 191 (CreateTables_CampanhaWhatsApp.sql).
'   Campanha + CampanhaDestinatario: a tela (frmCampanhaZap) só grava a campanha e a lista de clientes;
'   quem envia é a tarefa do servidor (OnlineCommerce.exe /cobrancazap -> Camp_ProcessarFila), aos poucos:
'   limite próprio por dia (CobrancaConfig.PromoLimiteDia, padrão 40), depois das cobranças (prioridade),
'   no horário/dias da Régua de cobrança, intervalo aleatório, uma promoção por cliente por dia e nunca no
'   mesmo dia de uma cobrança. Texto com variações {Olá|Oi} e rodapé de opt-out opcional (Campanha.RodapeSair, script 192).
'   Só no servidor: o arquivo (imagem, vídeo ou PDF) é copiado para <pasta do sistema>\Campanhas\<nome da campanha>\
'   (Campanha.ArquivoNome guarda esse caminho relativo) e a pasta é apagada ao cancelar ou excluir a campanha.
'   cliente.NaoReceberPromocao: quem não quer promoções (a cobrança continua).
'---------------------------------------------------------------------------

Public Const CAMP_SO_TEXTO As Integer = 0
Public Const CAMP_IMAGEM As Integer = 1
Public Const CAMP_VIDEO As Integer = 2
Public Const CAMP_SO_IMAGEM As Integer = 3
Public Const CAMP_PDF As Integer = 4

Public Const CAMP_NA_FILA As Integer = 1
Public Const CAMP_PAUSADA As Integer = 2
Public Const CAMP_CONCLUIDA As Integer = 3
Public Const CAMP_CANCELADA As Integer = 4

Public Const CAMP_DEST_AGUARDANDO As Integer = 0
Public Const CAMP_DEST_ENVIADA As Integer = 1
Public Const CAMP_DEST_ERRO As Integer = 2
Public Const CAMP_DEST_PULADA As Integer = 3

Public Const CAMP_LIMITE_PADRAO As Long = 40
Public Const CAMP_RODAPE As String = "_Para não receber mais promoções, responda SAIR._"
Public Const CAMP_MAX_IMAGEM As Long = 5242880      '5 MB
Public Const CAMP_MAX_ARQUIVO As Long = 15728640    '15 MB (vídeo e PDF; limite do servidor)

'Filtros do público (todos os marcados valem juntos)
Public Type CampFiltro
    Aniver As Integer          '0 = não filtra, 1 = do mês, 2 = de hoje
    ComprouDias As Long        'compraram nos últimos X dias (também limita produto/categoria)
    SumidoDias As Long         'já compraram, mas não nos últimos X dias
    Produto As String          'compraram produto com este texto na descrição
    Categoria As String        'compraram produto desta categoria
    Cidade As String
    Bairro As String
    GastouAcima As Currency    'gastaram acima disso nos últimos 12 meses
    ExcluirDevedores As Boolean
    SoCelularValido As Boolean  'Todos (somente com celular válido): quem não tem celular válido nem aparece
End Type

Public Type CampCliente
    CodCliente As Long
    Nome As String
    Celular As String
    Cidade As String
    UltimaCompra As Date       '0 = nunca
    Motivo As String           '"" = pode receber
End Type

'Venda que conta como compra do cliente
Private Const SQL_VENDA As String = "p.STATUS_PEDIDO = 1 AND ISNULL(p.CANCELADO, 0) = 0 AND p.TIPO_PEDIDO IN ('VENDA', 'OFICINA', 'ALUGUEL')"

'Tabelas do script 191 existem?
Public Function Camp_BancoPronto() As Boolean
On Error Resume Next
Camp_BancoPronto = (Val(CStr(SQLExecutaRetorno("SELECT CASE WHEN OBJECT_ID('dbo.Campanha', 'U') IS NULL " & _
    "OR OBJECT_ID('dbo.CampanhaDestinatario', 'U') IS NULL OR COL_LENGTH('dbo.cliente', 'NaoReceberPromocao') IS NULL " & _
    "OR COL_LENGTH('dbo.CobrancaConfig', 'PromoLimiteDia') IS NULL THEN 0 ELSE 1 END AS r", "r", 0))) = 1)
Err.Clear
End Function

'Coluna Campanha.RodapeSair (script 192) existe?
Public Function Camp_TemRodape() As Boolean
Static nTem As Integer   '0 = não verificado, 1 = sim, 2 = não

If nTem = 0 Then
    On Error Resume Next
    nTem = 2
    If Val(CStr(SQLExecutaRetorno("SELECT CASE WHEN COL_LENGTH('dbo.Campanha', 'RodapeSair') IS NULL THEN 0 ELSE 1 END AS r", "r", 0))) = 1 Then nTem = 1
    Err.Clear
End If
Camp_TemRodape = (nTem = 1)
End Function

'---------------------------------------------------------------------------
' Público
'---------------------------------------------------------------------------
Public Function Camp_Buscar(ByRef f As CampFiltro, ByRef Cli() As CampCliente) As Long
Dim rs As ADODB.Recordset, s As String, w As String, sPeriodo As String, n As Long, sMotivo As String
Dim sHoje As String

ReDim Cli(0)
sHoje = "CONVERT(datetime, CONVERT(varchar(8), GETDATE(), 112))"
If f.ComprouDias > 0 Then sPeriodo = " AND p.DATA_COMPRA >= DATEADD(day, -" & f.ComprouDias & ", " & sHoje & ")"

w = "WHERE c.STATUS = 1"
Select Case f.Aniver
    Case 1: w = w & " AND MONTH(c.Data_de_Nascimento) = MONTH(GETDATE())"
    Case 2: w = w & " AND MONTH(c.Data_de_Nascimento) = MONTH(GETDATE()) AND DAY(c.Data_de_Nascimento) = DAY(GETDATE())"
End Select
If Len(f.Produto) > 0 Or Len(f.Categoria) > 0 Then
    s = "EXISTS (SELECT 1 FROM pedidos p INNER JOIN pedidos_itens i ON i.COD_PEDIDO = p.COD_PEDIDO " & _
        "INNER JOIN produtos pr ON pr.CODIGO = i.COD_PRODUTO WHERE p.COD_CLIENTE = c.CODIGO AND " & SQL_VENDA & _
        " AND ISNULL(i.cancelado, 0) = 0" & sPeriodo
    If Len(f.Produto) > 0 Then s = s & " AND pr.DESCRICAO LIKE " & Camp_Like(f.Produto)
    If Len(f.Categoria) > 0 Then s = s & " AND pr.CATEGORIA = " & Cob_Q(f.Categoria)
    w = w & " AND " & s & ")"
ElseIf f.ComprouDias > 0 Then
    w = w & " AND EXISTS (SELECT 1 FROM pedidos p WHERE p.COD_CLIENTE = c.CODIGO AND " & SQL_VENDA & sPeriodo & ")"
End If
If f.SumidoDias > 0 Then
    w = w & " AND EXISTS (SELECT 1 FROM pedidos p WHERE p.COD_CLIENTE = c.CODIGO AND " & SQL_VENDA & ")" & _
            " AND NOT EXISTS (SELECT 1 FROM pedidos p WHERE p.COD_CLIENTE = c.CODIGO AND " & SQL_VENDA & _
            " AND p.DATA_COMPRA >= DATEADD(day, -" & f.SumidoDias & ", " & sHoje & "))"
End If
If Len(f.Cidade) > 0 Then w = w & " AND c.Cidade = " & Cob_Q(f.Cidade)
If Len(f.Bairro) > 0 Then w = w & " AND c.Bairro = " & Cob_Q(f.Bairro)
If f.GastouAcima > 0 Then
    w = w & " AND (SELECT SUM(p.TOTAL) FROM pedidos p WHERE p.COD_CLIENTE = c.CODIGO AND " & SQL_VENDA & _
            " AND p.DATA_COMPRA >= DATEADD(month, -12, GETDATE())) > " & Replace(CStr(f.GastouAcima), ",", ".")
End If
If f.SoCelularValido Then w = w & " AND LEN(LTRIM(RTRIM(ISNULL(c.CELULAR, '')))) >= 8"
If f.ExcluirDevedores Then
    w = w & " AND NOT EXISTS (SELECT 1 FROM parcelas pa INNER JOIN pedidos pe ON pe.COD_PEDIDO = pa.COD_PEDIDO " & _
            "WHERE pe.COD_CLIENTE = c.CODIGO AND pa.STATUS = 0 AND pa.DATA < " & sHoje & ")"
End If

s = "SELECT c.CODIGO, c.Nome, c.CELULAR, c.Cidade, ISNULL(c.NaoReceberPromocao, 0) AS NAO, " & _
    "(SELECT MAX(p.DATA_COMPRA) FROM pedidos p WHERE p.COD_CLIENTE = c.CODIGO AND " & SQL_VENDA & ") AS ULT " & _
    "FROM cliente c " & w & " ORDER BY c.Nome"
On Error GoTo Falha
Set rs = dbData.OpenRecordset(s)
Do While Not rs.EOF
    n = n + 1
    ReDim Preserve Cli(n)
    With Cli(n)
        .CodCliente = rs("CODIGO")
        .Nome = Trim$("" & rs("Nome"))
        .Celular = Trim$("" & rs("CELULAR"))
        .Cidade = Trim$("" & rs("Cidade"))
        If IsNull(rs("ULT")) Then .UltimaCompra = 0 Else .UltimaCompra = DateValue(rs("ULT"))
        If CBool(rs("NAO")) Then
            .Motivo = "Não quer promoções"
        ElseIf Not CelularValido(.Celular, sMotivo) Then
            If sMotivo = "sem número" Then .Motivo = "Sem celular" Else .Motivo = "Celular inválido (" & sMotivo & ")"
        End If
    End With
    If f.SoCelularValido And Left$(Cli(n).Motivo, 3) <> "Não" And Len(Cli(n).Motivo) > 0 Then n = n - 1: ReDim Preserve Cli(n)   'sem celular válido: fora da lista
    rs.MoveNext
Loop
rs.Close
Camp_Buscar = n
Exit Function
Falha:
MsgBox "Não foi possível buscar os clientes." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Promoções por WhatsApp"
Err.Clear
Camp_Buscar = n
End Function

'LIKE com o texto escapado (% _ [ viram literais)
Private Function Camp_Like(ByVal s As String) As String
s = Replace(s, "[", "[[]")
s = Replace(s, "%", "[%]")
s = Replace(s, "_", "[_]")
Camp_Like = Cob_Q("%" & s & "%")
End Function

'---------------------------------------------------------------------------
' Texto da mensagem
'---------------------------------------------------------------------------
'Variações {Olá|Oi|Bom dia} sorteadas por cliente (bPrimeira = sempre a primeira, para a prévia), tags do cliente,
'emojis e, com bRodape, o rodapé de opt-out. bPrevia: emojis como [nome] (TextBox).
Public Function Camp_Texto(ByVal sMsg As String, ByVal sNomeCliente As String, ByVal bPrimeira As Boolean, ByVal bPrevia As Boolean, _
                           Optional ByVal bRodape As Boolean = False) As String
Dim s As String, i As Long, tags() As String, desc() As String, emo() As String

s = Camp_Variacoes(Replace(sMsg, vbCrLf, vbLf), bPrimeira)
s = Replace(s, "{PRIMEIRO_NOME}", Cob_PrimeiroNome(sNomeCliente), , , vbTextCompare)
s = Replace(s, "{NOME_CLIENTE}", StrConv(sNomeCliente, vbProperCase), , , vbTextCompare)
s = Replace(s, "{EMPRESA}", PDFZap_NomeEmpresa(), , , vbTextCompare)
s = Trim$(s)
Do While Right$(s, 1) = vbLf
    s = Left$(s, Len(s) - 1)
Loop
If bRodape Then s = s & IIf(Len(s) > 0, vbLf & vbLf, "") & CAMP_RODAPE

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
Camp_Texto = s
End Function

'Troca cada {a|b|c} por uma das opções (as tags, sem "|", ficam como estão)
Public Function Camp_Variacoes(ByVal s As String, ByVal bPrimeira As Boolean) As String
Dim pFim As Long, pIni As Long, sGrupo As String, op() As String, nIni As Long

Randomize
nIni = 1
Do
    pFim = InStr(nIni, s, "}")
    If pFim = 0 Then Exit Do
    pIni = InStrRev(s, "{", pFim)
    If pIni = 0 Then Exit Do
    sGrupo = Mid$(s, pIni + 1, pFim - pIni - 1)
    If InStr(sGrupo, "|") > 0 Then
        op = Split(sGrupo, "|")
        If bPrimeira Then sGrupo = op(0) Else sGrupo = op(Int(Rnd * (UBound(op) + 1)))
        s = Left$(s, pIni - 1) & sGrupo & Mid$(s, pFim + 1)
        nIni = pIni
    Else
        nIni = pFim + 1        'tag ({PRIMEIRO_NOME}...)
    End If
Loop
Camp_Variacoes = s
End Function

'Tags para a tela: dados do cliente + emojis da cobrança
Public Sub Camp_ListaTags(ByRef tags() As String, ByRef desc() As String)
Dim t() As String, d() As String, e() As String, i As Long, n As Long

Cob_ListaTags t, d, e
ReDim tags(0 To 3): ReDim desc(0 To 3)
tags(0) = "{PRIMEIRO_NOME}": desc(0) = "Primeiro nome do cliente"
tags(1) = "{NOME_CLIENTE}": desc(1) = "Nome completo do cliente"
tags(2) = "{EMPRESA}": desc(2) = "Nome da empresa"
tags(3) = "{Olá|Oi|Bom dia}": desc(3) = "Variação: cada cliente recebe uma das opções (evita bloqueio)"
n = 3
For i = 0 To UBound(t)
    If Len(e(i)) > 0 Then
        n = n + 1
        ReDim Preserve tags(0 To n): ReDim Preserve desc(0 To n)
        tags(n) = t(i): desc(n) = d(i)
    End If
Next
End Sub

'Prévia em HTML (balão do WhatsApp) para o WebBrowser: emojis de verdade (entidades HTML), *negrito* _itálico_ ~riscado~
Public Function Camp_PreviaHtml(ByVal sMsg As String, ByVal sNomeCliente As String, ByVal nTipo As Integer, ByVal sArquivo As String, _
                                Optional ByVal bRodape As Boolean = False) As String
Dim s As String, i As Long, tags() As String, desc() As String, emo() As String, p() As String, j As Long, sEnt As String
Dim sMidia As String

s = Camp_Texto(sMsg, sNomeCliente, True, True, bRodape)      'emojis como [nome]
s = Replace(s, vbCrLf, vbLf)
s = Replace(s, "&", "&amp;")
s = Replace(s, "<", "&lt;")
s = Replace(s, ">", "&gt;")
Cob_ListaTags tags, desc, emo
For i = 0 To UBound(tags)
    If Len(emo(i)) > 0 Then
        p = Split(emo(i), "+")
        sEnt = ""
        For j = 0 To UBound(p)
            sEnt = sEnt & "&#x" & p(j) & ";"
        Next
        s = Replace(s, "[" & LCase$(Mid$(tags(i), 2, Len(tags(i)) - 2)) & "]", "<span class=""e"">" & sEnt & "</span>")
    End If
Next
s = Camp_Marcar(s, "*", "b")
s = Camp_Marcar(s, "_", "i")
s = Camp_Marcar(s, "~", "s")
s = Replace(s, vbLf, "<br>")

Select Case nTipo
    Case CAMP_IMAGEM, CAMP_SO_IMAGEM
        If Len(sArquivo) > 0 Then
            sMidia = "<img src=""file:///" & Replace(sArquivo, "\", "/") & """ style=""width:100%;border-radius:6px;margin-bottom:4px"">"
        Else
            sMidia = "<div class=""m"">[imagem]</div>"
        End If
    Case CAMP_VIDEO
        sMidia = "<div class=""m"">&#x1F3AC; " & IIf(Len(sArquivo) > 0, Mid$(sArquivo, InStrRev(sArquivo, "\") + 1), "vídeo") & "</div>"
    Case CAMP_PDF
        sMidia = "<div class=""m"">&#x1F4C4; " & IIf(Len(sArquivo) > 0, Mid$(sArquivo, InStrRev(sArquivo, "\") + 1), "documento.pdf") & "</div>"
End Select

Camp_PreviaHtml = "<html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge""><meta charset=""windows-1252"">" & _
    "<style>body{margin:0;padding:10px;background:#E5DDD5;font-family:'Segoe UI',Arial;font-size:13px}" & _
    ".b{background:#DCF8C6;border-radius:8px;padding:6px 8px 4px 8px;margin-left:30px;box-shadow:0 1px 1px #aaa;word-wrap:break-word}" & _
    ".m{background:#cfe9ba;border-radius:6px;padding:14px 8px;margin-bottom:4px;color:#333}" & _
    ".h{text-align:right;color:#777;font-size:11px;margin-top:2px}.e{font-family:'Segoe UI Emoji'}</style></head>" & _
    "<body><div class=""b"">" & sMidia & s & "<div class=""h"">" & Format$(Now, "hh:mm") & "</div></div></body></html>"
End Function

'*texto* -> <b>texto</b> (só quando abre e fecha na mesma linha)
Private Function Camp_Marcar(ByVal s As String, ByVal sMarca As String, ByVal sTag As String) As String
Dim p1 As Long, p2 As Long, pLf As Long, nIni As Long

nIni = 1
Do
    p1 = InStr(nIni, s, sMarca)
    If p1 = 0 Then Exit Do
    p2 = InStr(p1 + 1, s, sMarca)
    If p2 = 0 Then Exit Do
    pLf = InStr(p1 + 1, s, vbLf)
    If (pLf > 0 And pLf < p2) Or p2 = p1 + 1 Then
        nIni = p1 + 1
    Else
        s = Left$(s, p1 - 1) & "<" & sTag & ">" & Mid$(s, p1 + 1, p2 - p1 - 1) & "</" & sTag & ">" & Mid$(s, p2 + 1)
        nIni = p2 + Len(sTag) * 2 + 3
    End If
Loop
Camp_Marcar = s
End Function

'---------------------------------------------------------------------------
' Gravação
'---------------------------------------------------------------------------
'Grava a campanha e os destinatários (na fila do servidor). Devolve o Id (0 = erro, sErro com o motivo).
Public Function Camp_Salvar(ByVal sNome As String, ByVal sMsg As String, ByVal nTipo As Integer, ByVal sArquivo As String, _
                            ByVal vAgendada As Variant, ByVal sPublico As String, ByRef Cods() As Long, ByRef Cels() As String, _
                            ByVal nQtd As Long, ByRef sErro As String, Optional ByVal bRodape As Boolean = False) As Long
Dim r As ADODB.Recordset, lId As Long, i As Long, s As String, nLote As Long, bTrans As Boolean, sUsuario As String, sAgenda As String
Dim sPastaCriada As String

On Error Resume Next
sUsuario = Trim$(Tela_Principal.txtCodFuncionario.Text)
Err.Clear
If IsDate(vAgendada) Then sAgenda = "'" & Format$(vAgendada, "yyyymmdd hh:nn") & "'" Else sAgenda = "NULL"   'IIf avaliaria Format$(Null)
On Error GoTo Falha
dbData.Execute "BEGIN TRANSACTION"
bTrans = True
Set r = dbData.OpenRecordset("INSERT INTO Campanha (Nome, Mensagem, TipoMidia, ArquivoNome, Publico, AgendadaPara, Situacao, Usuario) " & _
    "OUTPUT INSERTED.Id VALUES (" & Cob_Q(Left$(sNome, 100)) & ", " & Cob_Q(Left$(sMsg, 2000)) & ", " & nTipo & ", " & _
    "NULL, " & _
    Cob_Q(Left$(sPublico, 1000)) & ", " & sAgenda & ", " & _
    CAMP_NA_FILA & ", " & Cob_Q(Left$(sUsuario, 50)) & ")")
If Not r.EOF Then lId = r(0)
r.Close
If lId = 0 Then Err.Raise vbObjectError + 1, , "O banco não devolveu o número da campanha."

If bRodape And Camp_TemRodape() Then dbData.Execute "UPDATE Campanha SET RodapeSair = 1 WHERE Id = " & lId
If Len(sArquivo) > 0 Then
    s = Camp_CopiarArquivo(sArquivo, sNome, sPastaCriada)
    dbData.Execute "UPDATE Campanha SET ArquivoNome = " & Cob_Q(s) & " WHERE Id = " & lId
    s = ""
End If

'Destinatários em lotes de 500 linhas por INSERT
For i = 1 To nQtd
    s = s & IIf(nLote > 0, ", ", "") & "(" & lId & ", " & Cods(i) & ", " & Cob_Q(Left$(Cels(i), 20)) & ")"
    nLote = nLote + 1
    If nLote = 500 Or i = nQtd Then
        dbData.Execute "INSERT INTO CampanhaDestinatario (CodCampanha, CodCliente, Celular) VALUES " & s
        s = "": nLote = 0
    End If
Next
dbData.Execute "COMMIT TRANSACTION"
Camp_Salvar = lId
Exit Function
Falha:
sErro = Err.Description
Err.Clear
On Error Resume Next
If bTrans Then dbData.Execute "ROLLBACK TRANSACTION"
If Len(sPastaCriada) > 0 Then Camp_ApagarPasta sPastaCriada
Err.Clear
End Function

'---------------------------------------------------------------------------
' Arquivo da campanha: <pasta do sistema>\Campanhas\<nome da campanha>\<arquivo> (só no servidor)
'---------------------------------------------------------------------------
Private Function Camp_PastaSistema() As String
Camp_PastaSistema = appPathApp
If Right$(Camp_PastaSistema, 1) <> "\" Then Camp_PastaSistema = Camp_PastaSistema & "\"
End Function

'Nome da campanha sem os caracteres que o Windows não aceita em nome de pasta
Private Function Camp_NomePasta(ByVal sNome As String) As String
Dim i As Long, c As String, s As String

For i = 1 To Len(sNome)
    c = Mid$(sNome, i, 1)
    If InStr("\/:*?""<>|", c) > 0 Or Asc(c) < 32 Then c = " "
    s = s & c
Next
s = Trim$(s)
Do While InStr(s, "  ") > 0
    s = Replace(s, "  ", " ")
Loop
If Len(s) > 60 Then s = Trim$(Left$(s, 60))
Do While Right$(s, 1) = "."      'o Windows não aceita pasta terminada em ponto
    s = Trim$(Left$(s, Len(s) - 1))
Loop
Select Case UCase$(s)
    Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "LPT1", "LPT2", "LPT3": s = s & "_"
End Select
If Len(s) = 0 Then s = "Campanha"
Camp_NomePasta = s
End Function

'Copia o arquivo para Campanhas\<nome da campanha>\ (outra campanha com o mesmo nome: "<nome> (2)").
'Devolve o caminho relativo à pasta do sistema (vai para Campanha.ArquivoNome). Erro sobe para Camp_Salvar.
Private Function Camp_CopiarArquivo(ByVal sArquivo As String, ByVal sNomeCampanha As String, ByRef sPastaCriada As String) As String
Dim sRaiz As String, sPasta As String, sBase As String, sArq As String, n As Long

sRaiz = Camp_PastaSistema()
If Len(Dir$(sRaiz & "Campanhas", vbDirectory)) = 0 Then MkDir sRaiz & "Campanhas"
sBase = "Campanhas\" & Camp_NomePasta(sNomeCampanha)
sPasta = sBase
n = 1
Do While Len(Dir$(sRaiz & sPasta, vbDirectory)) > 0
    n = n + 1
    sPasta = sBase & " (" & n & ")"
Loop
MkDir sRaiz & sPasta
sPastaCriada = sRaiz & sPasta
sArq = Mid$(sArquivo, InStrRev(sArquivo, "\") + 1)
If Len(sPasta) + 1 + Len(sArq) > 200 And InStrRev(sArq, ".") > 0 Then sArq = "arquivo" & Mid$(sArq, InStrRev(sArq, "."))
FileCopy sArquivo, sRaiz & sPasta & "\" & sArq
SetAttr sRaiz & sPasta & "\" & sArq, vbNormal      'cópia de arquivo somente leitura não impediria apagar depois
Camp_CopiarArquivo = sPasta & "\" & sArq
End Function

'Caminho completo do arquivo da campanha no servidor ("" se não tiver ou se não estiver mais na pasta)
Public Function Camp_CaminhoArquivo(ByVal lCampanha As Long) As String
Dim s As String

s = Camp_ArquivoRelativo(lCampanha)
If Not Camp_RelativoValido(s) Then Exit Function
s = Camp_PastaSistema() & s
On Error Resume Next
If Len(Dir$(s)) > 0 Then Camp_CaminhoArquivo = s
Err.Clear
End Function

Private Function Camp_ArquivoRelativo(ByVal lCampanha As Long) As String
Dim r As ADODB.Recordset

On Error Resume Next
Set r = dbData.OpenRecordset("SELECT ArquivoNome FROM Campanha WHERE Id = " & lCampanha)
If Not r.EOF Then Camp_ArquivoRelativo = "" & r("ArquivoNome")
r.Close
Err.Clear
End Function

'Só Campanhas\<pasta>\<arquivo> (nunca apaga nada fora da pasta Campanhas)
Private Function Camp_RelativoValido(ByVal s As String) As Boolean
If InStr(1, s, "Campanhas\", vbTextCompare) <> 1 Then Exit Function
If InStr(s, "..") > 0 Or InStr(s, ":") > 0 Then Exit Function
Camp_RelativoValido = (UBound(Split(s, "\")) = 2)
End Function

'Apaga a pasta da campanha (ao cancelar ou excluir). Arquivo em uso no momento: fica e não dá erro.
Public Sub Camp_ApagarArquivo(ByVal lCampanha As Long)
Dim s As String

s = Camp_ArquivoRelativo(lCampanha)
If Not Camp_RelativoValido(s) Then Exit Sub
Camp_ApagarPasta Camp_PastaSistema() & Left$(s, InStrRev(s, "\") - 1)
End Sub

Private Sub Camp_ApagarPasta(ByVal sPasta As String)
Dim sArq As String, Nomes() As String, n As Long, i As Long

On Error Resume Next
If InStr(1, sPasta, Camp_PastaSistema() & "Campanhas\", vbTextCompare) <> 1 Then Exit Sub
If Len(Dir$(sPasta, vbDirectory)) = 0 Then Exit Sub
ReDim Nomes(0)
sArq = Dir$(sPasta & "\*.*", vbNormal + vbHidden + vbReadOnly + vbSystem)
Do While Len(sArq) > 0
    n = n + 1
    ReDim Preserve Nomes(n)
    Nomes(n) = sArq
    sArq = ""
    sArq = Dir$
Loop
For i = 1 To n
    SetAttr sPasta & "\" & Nomes(i), vbNormal
    Kill sPasta & "\" & Nomes(i)
Next
RmDir sPasta
Err.Clear
End Sub

'Exclui a campanha (histórico dos clientes junto) e apaga a pasta do arquivo
Public Function Camp_Excluir(ByVal lCampanha As Long, ByRef sErro As String) As Boolean
Dim s As String, bTrans As Boolean

s = Camp_ArquivoRelativo(lCampanha)
On Error GoTo Falha
dbData.Execute "BEGIN TRANSACTION"
bTrans = True
dbData.Execute "DELETE FROM CampanhaDestinatario WHERE CodCampanha = " & lCampanha
dbData.Execute "DELETE FROM Campanha WHERE Id = " & lCampanha
dbData.Execute "COMMIT TRANSACTION"
bTrans = False
If Camp_RelativoValido(s) Then Camp_ApagarPasta Camp_PastaSistema() & Left$(s, InStrRev(s, "\") - 1)
Camp_Excluir = True
Exit Function
Falha:
sErro = Err.Description
Err.Clear
On Error Resume Next
If bTrans Then dbData.Execute "ROLLBACK TRANSACTION"
Err.Clear
End Function

'---------------------------------------------------------------------------
' Limite e situação
'---------------------------------------------------------------------------
Public Function Camp_LimiteDia() As Long
On Error Resume Next
Camp_LimiteDia = Val(CStr(SQLExecutaRetorno("SELECT TOP 1 PromoLimiteDia AS r FROM CobrancaConfig WHERE Id = 1", "r", CAMP_LIMITE_PADRAO)))
If Camp_LimiteDia < 1 Then Camp_LimiteDia = CAMP_LIMITE_PADRAO
Err.Clear
End Function

Public Sub Camp_SalvarLimite(ByVal n As Long)
If n < 1 Then Exit Sub
On Error Resume Next
dbData.Execute "UPDATE CobrancaConfig SET PromoLimiteDia = " & n & " WHERE Id = 1"
Err.Clear
End Sub

Public Function Camp_EnviadasHoje() As Long
On Error Resume Next
Camp_EnviadasHoje = Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) AS n FROM CampanhaDestinatario WHERE Situacao = " & CAMP_DEST_ENVIADA & _
    " AND DataHora >= '" & Format$(Date, "yyyymmdd") & "'", "n", 0)))
Err.Clear
End Function

Public Function Camp_SituacaoTexto(ByVal n As Integer) As String
Select Case n
    Case CAMP_NA_FILA: Camp_SituacaoTexto = "Enviando"
    Case CAMP_PAUSADA: Camp_SituacaoTexto = "Pausada"
    Case CAMP_CONCLUIDA: Camp_SituacaoTexto = "Concluída"
    Case CAMP_CANCELADA: Camp_SituacaoTexto = "Cancelada"
    Case Else: Camp_SituacaoTexto = "?"
End Select
End Function

Public Function Camp_TipoTexto(ByVal n As Integer) As String
Select Case n
    Case CAMP_IMAGEM: Camp_TipoTexto = "Texto + imagem"
    Case CAMP_VIDEO: Camp_TipoTexto = "Texto + vídeo"
    Case CAMP_SO_IMAGEM: Camp_TipoTexto = "Só imagem"
    Case CAMP_PDF: Camp_TipoTexto = "Texto + PDF"
    Case Else: Camp_TipoTexto = "Só texto"
End Select
End Function

'Pausar / retomar / cancelar (o servidor obedece antes da próxima mensagem)
Public Function Camp_MudarSituacao(ByVal lCampanha As Long, ByVal nSituacao As Integer) As Boolean
On Error GoTo Falha
dbData.Execute "UPDATE Campanha SET Situacao = " & nSituacao & IIf(nSituacao = CAMP_CANCELADA, ", ConcluidaEm = GETDATE()", "") & _
    " WHERE Id = " & lCampanha & " AND Situacao IN (" & CAMP_NA_FILA & ", " & CAMP_PAUSADA & ")"
If nSituacao = CAMP_CANCELADA Then
    dbData.Execute "UPDATE CampanhaDestinatario SET Situacao = " & CAMP_DEST_PULADA & ", Erro = N'Campanha cancelada' " & _
        "WHERE CodCampanha = " & lCampanha & " AND Situacao = " & CAMP_DEST_AGUARDANDO
    Camp_ApagarArquivo lCampanha
End If
Camp_MudarSituacao = True
Exit Function
Falha:
Err.Clear
End Function

'Alguma campanha com mensagem para sair agora? (sem o script 191: False)
Public Function Camp_TemPendente() As Boolean
On Error Resume Next
Camp_TemPendente = (Val(CStr(SQLExecutaRetorno("SELECT CASE WHEN EXISTS (SELECT 1 FROM CampanhaDestinatario d INNER JOIN Campanha ca " & _
    "ON ca.Id = d.CodCampanha WHERE d.Situacao = " & CAMP_DEST_AGUARDANDO & " AND ca.Situacao = " & CAMP_NA_FILA & _
    " AND (ca.AgendadaPara IS NULL OR ca.AgendadaPara <= GETDATE())) THEN 1 ELSE 0 END AS r", "r", 0))) = 1)
Err.Clear
End Function

'---------------------------------------------------------------------------
' Envio (servidor, chamado por ExecutarCobrancaZapAuto depois das cobranças)
'---------------------------------------------------------------------------
Public Sub Camp_ProcessarFila(ByRef cfg As CobConfig)
Dim r As ADODB.Recordset, nRestam As Long, nOk As Long, nErro As Long, nPuladas As Long, nSeguidos As Long
Dim lDest As Long, lCamp As Long, lCampArq As Long, sArq As String, sCel As String, sNome As String, sMsg As String
Dim nTipo As Integer, sTexto As String, sErro As String, sMotivo As String, sSit As String, sNum As String, sNomeZap As String
Dim sHoje As String, sParou As String, bRodape As Boolean

On Error GoTo Falha
If Not Camp_BancoPronto() Then Exit Sub
nRestam = Camp_LimiteDia() - Camp_EnviadasHoje()
If nRestam <= 0 Then Exit Sub
If Not Camp_TemPendente() Then Exit Sub

sSit = ZapOnlineSituacao(sNum, sNomeZap, sMsg)
If sSit <> "conectado" Then
    If sSit = "" Then sErro = sMsg Else sErro = "O WhatsApp da empresa não está conectado (Sistema > WhatsApp)."
    Cob_Log "Promoções não enviadas: " & Replace(sErro, vbCrLf, " ")
    Exit Sub
End If

sHoje = Format$(Date, "yyyymmdd")
Cob_Log "Promoções: início (limite restante hoje: " & nRestam & ")."
Do While nOk + nErro < nRestam
    If Not Cob_DentroDoHorario(cfg) Then sParou = "fim do horário de envio": Exit Do
    'Próximo da fila: campanha mais antiga primeiro; uma promoção por cliente por dia e nunca no dia de uma cobrança
    Set r = dbData.OpenRecordset("SELECT TOP 1 d.Id, d.CodCliente, d.Celular, c.Nome, ISNULL(c.NaoReceberPromocao, 0) AS NAO, " & _
        "ca.Id AS CAMP, ca.Mensagem, ca.TipoMidia, " & IIf(Camp_TemRodape(), "ca.RodapeSair", "0 AS RodapeSair") & " FROM CampanhaDestinatario d INNER JOIN Campanha ca ON ca.Id = d.CodCampanha " & _
        "INNER JOIN cliente c ON c.CODIGO = d.CodCliente WHERE d.Situacao = " & CAMP_DEST_AGUARDANDO & " AND ca.Situacao = " & CAMP_NA_FILA & _
        " AND (ca.AgendadaPara IS NULL OR ca.AgendadaPara <= GETDATE()) " & _
        "AND d.CodCliente NOT IN (SELECT CodCliente FROM CampanhaDestinatario WHERE Situacao = " & CAMP_DEST_ENVIADA & " AND DataHora >= '" & sHoje & "') " & _
        "AND d.CodCliente NOT IN (SELECT CodCliente FROM CobrancaEnvio WHERE Situacao = 1 AND DataHora >= '" & sHoje & "') " & _
        "ORDER BY ca.Id, d.Id")
    If r.EOF Then r.Close: Exit Do
    lDest = r("Id"): lCamp = r("CAMP"): nTipo = Val("" & r("TipoMidia"))
    sCel = Trim$("" & r("Celular")): sNome = Trim$("" & r("Nome")): sMsg = "" & r("Mensagem")
    bRodape = CBool(r("RodapeSair"))
    sMotivo = ""
    If CBool(r("NAO")) Then sMotivo = "Cliente não quer promoções"
    r.Close
    If Len(sMotivo) = 0 Then
        If Not CelularValido(sCel, sMotivo) Then sMotivo = "Celular inválido (" & sMotivo & ")" Else sMotivo = ""
    End If

    If Len(sMotivo) > 0 Then
        Camp_MarcarDestinatario lDest, CAMP_DEST_PULADA, sMotivo
        nPuladas = nPuladas + 1
    Else
        If nTipo <> CAMP_SO_TEXTO And lCamp <> lCampArq Then
            sArq = Camp_CaminhoArquivo(lCamp)
            lCampArq = lCamp
        End If
        If nTipo <> CAMP_SO_TEXTO And Len(sArq) = 0 Then
            Camp_MarcarDestinatario lDest, CAMP_DEST_ERRO, "Arquivo da campanha não encontrado na pasta Campanhas do servidor"
            nErro = nErro + 1
        Else
            'intervalo antes de toda promoção (inclusive a primeira: as cobranças podem ter acabado de sair)
            Cob_Esperar Cob_Intervalo(cfg)
            'A campanha ainda está ativa? (pausada/cancelada pela tela durante a espera)
            If Val(CStr(SQLExecutaRetorno("SELECT Situacao AS r FROM Campanha WHERE Id = " & lCamp, "r", 0))) = CAMP_NA_FILA Then
                sTexto = Camp_Texto(sMsg, sNome, False, False, bRodape)
                If ZapOnlineEnviar(sCel, sTexto, IIf(nTipo = CAMP_SO_TEXTO, "", sArq), sErro) Then
                    Camp_MarcarDestinatario lDest, CAMP_DEST_ENVIADA, ""
                    nOk = nOk + 1
                    nSeguidos = 0
                Else
                    sErro = Replace(sErro, vbCrLf, " ")
                    Camp_MarcarDestinatario lDest, CAMP_DEST_ERRO, sErro
                    nErro = nErro + 1
                    Cob_Log "Promoção - erro: " & sNome & " (" & sCel & "): " & sErro
                    If ZapOnlineForaDoPlano() Then sParou = sErro: Exit Do
                    'número sem WhatsApp é problema do cliente, não do envio
                    If InStr(1, sErro, "não tem WhatsApp", vbTextCompare) = 0 Then nSeguidos = nSeguidos + 1
                    If nSeguidos >= 3 Then sParou = "3 erros seguidos (último: " & sErro & ")": Exit Do
                End If
            End If
        End If
    End If
    Camp_FecharConcluidas
Loop
Camp_FecharConcluidas
If nOk + nErro + nPuladas > 0 Then
    Cob_Log "Promoções: " & nOk & " enviada(s)" & IIf(nErro > 0, ", " & nErro & " com erro", "") & _
            IIf(nPuladas > 0, ", " & nPuladas & " pulada(s)", "") & IIf(Len(sParou) > 0, " (parou: " & sParou & ")", "") & "."
End If
Exit Sub
Falha:
Cob_Log "Promoções - erro: " & Err.Description
Err.Clear
End Sub

Private Sub Camp_MarcarDestinatario(ByVal lDest As Long, ByVal nSit As Integer, ByVal sErro As String)
On Error Resume Next
dbData.Execute "UPDATE CampanhaDestinatario SET Situacao = " & nSit & ", DataHora = GETDATE(), Erro = " & _
    IIf(Len(sErro) = 0, "NULL", Cob_Q(Left$(sErro, 300))) & " WHERE Id = " & lDest
Err.Clear
End Sub

'Campanha sem ninguém aguardando = concluída
Private Sub Camp_FecharConcluidas()
On Error Resume Next
dbData.Execute "UPDATE Campanha SET Situacao = " & CAMP_CONCLUIDA & ", ConcluidaEm = GETDATE() WHERE Situacao = " & CAMP_NA_FILA & _
    " AND NOT EXISTS (SELECT 1 FROM CampanhaDestinatario d WHERE d.CodCampanha = Campanha.Id AND d.Situacao = " & CAMP_DEST_AGUARDANDO & ")"
Err.Clear
End Sub
