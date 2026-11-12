Attribute VB_Name = "modMonitorNFe"
Option Explicit
'Monitor de NOTAS DESTINADAS (manifestação do destinatário), sem tela: OnlineCommerce.exe /monitornfe, pela tarefa
'agendada "OnlineCommerce_MonitorNFe" (de hora em hora), criada SÓ no servidor pelo Tela_Principal.
'1) Consulta a SEFAZ (DistribuicaoDFe) a partir de empresa.UltimoNSU - o mesmo NSU da tela Notas_Destinadas_Consulta.
'   Espera 1 hora desde a última consulta (empresa.DFeUltimaConsultaData, também gravada pela tela), como pede a SEFAZ
'   para "nenhum documento" (137) e "consumo indevido" (656); enquanto houver documento novo continua na mesma execução.
'2) Grava/atualiza TbNotasDestinadasManifestacao do mesmo jeito que a tela.
'3) Envia a CIÊNCIA DA OPERAÇÃO para toda nota autorizada ainda não manifestada: a SEFAZ libera o XML completo (procNFe)
'   numa próxima consulta; a DLL grava em nfe\arquivos\DFe e aqui é copiado para nfe\arquivos\ConfRecebto, onde a tela
'   procura para Importar/DANFe. Ciência recusada fica anotada em TextoCorrecao ("MonitorNFe: ...") e não é repetida.
'4) Com configuracao Importar_XML = 1 (padrão) importa as notas de compra na Entrada de Estoque (ver MON_ImportarNotas).
'Desligar numa empresa: configuracao MONITOR_NFE = 0. Log em MonitorNFe_log.txt na pasta do sistema.

Private Const MON_MINUTOS_ESPERA As Long = 60
Private Const MON_MAX_CONSULTAS As Long = 20     'lotes de até 50 documentos por execução
Private Const MON_MAX_CIENCIAS As Long = 50      'ciências por execução
Private Const MON_MAX_IMPORTACOES As Long = 30   'importações automáticas na Entrada de Estoque por execução

Private mDirXML As String
Private mCNPJ As String

Private Sub MON_Log(ByVal sMsg As String)
Dim nLog As Integer
On Error Resume Next
nLog = FreeFile
Open appPathApp & "MonitorNFe_log.txt" For Append As #nLog
Print #nLog, Format(Now, "dd/mm/yyyy hh:nn:ss") & "  " & sMsg
Close #nLog
End Sub

Public Sub ExecutarMonitorNFeAuto()
Dim sistNFe As snfe.Util
On Error GoTo deuErro

If Not EhServidorLocal() Then Exit Sub
If CStr(SQLExecutaRetorno("SELECT TOP 1 config_valor AS r FROM configuracao WHERE config_nome = 'MONITOR_NFE'", "r", "1")) = "0" Then Exit Sub
If Vazio(SQLExecutaRetorno("SELECT TOP 1 CertificadoDigital AS r FROM empresa", "r", "")) Then Exit Sub

mCNPJ = MON_SoDigitos(SQLExecutaRetorno("SELECT TOP 1 CNPJ AS r FROM empresa", "r", ""))
If Len(mCNPJ) <> 14 And Len(mCNPJ) <> 11 Then
   MON_Log "CNPJ da empresa inválido - monitor não executado."
   Exit Sub
End If
mDirXML = SQLExecutaRetorno("SELECT TOP 1 DiretorioXML AS r FROM empresa", "r", "")
If Vazio(mDirXML) Then mDirXML = appPathApp
If Right$(mDirXML, 1) <> "\" Then mDirXML = mDirXML & "\"

If Not ConfiguraDLLNFeNFCe(55, "1", sistNFe, True) Then
   MON_Log "Não foi possível configurar a DLL da NF-e."
   Exit Sub
End If
sistNFe.exibirAvisos = False
sistNFe.certificadoAvisaVencimento = False

MON_Consultar sistNFe
MON_EnviarCiencias sistNFe
Set sistNFe = Nothing
MON_ImportarNotas

Exit Sub

deuErro:
MON_Log "Erro: " & Err.Description
Err.Clear
Set sistNFe = Nothing
End Sub

'---------- 1) consulta (DistribuicaoDFe) ----------

Private Sub MON_Consultar(sistNFe As snfe.Util)
Dim vMin As Variant, n As Long, sNSU As String, sUlt As String, sMax As String
Dim bOk As Boolean, retWS As Long, qtReg As Long, retXML As String, ret As String
Dim cSt As Long, sMotivo As String

vMin = SQLExecutaRetorno("SELECT TOP 1 DATEDIFF(mi, DFeUltimaConsultaData, GETDATE()) AS r FROM empresa", "r", Null)
If Not IsNull(vMin) Then
   If CLng(vMin) >= 0 And CLng(vMin) < MON_MINUTOS_ESPERA Then Exit Sub
End If

sNSU = CStr(SQLExecutaRetorno("SELECT TOP 1 ISNULL(UltimoNSU, 0) AS r FROM empresa", "r", 0))
For n = 1 To MON_MAX_CONSULTAS
   retXML = "": ret = "": qtReg = 0: retWS = 0
   bOk = sistNFe.DistribuicaoDFe(LPad(CStr(sNSU), 15, "0"), retWS, qtReg, retXML, ret)
   If Not bOk Then
      MON_GravarConsulta ""
      MON_Log "Consulta não realizada: " & retXML
      Exit For
   End If
   cSt = sistNFe.retDistDFeInt.cStat
   sMotivo = sistNFe.retDistDFeInt.xMotivo
   sUlt = Format$(sistNFe.retDistDFeInt.ultNSU, "0")
   sMax = Format$(sistNFe.retDistDFeInt.maxNSU, "0")
   Select Case cSt
   Case 138
      MON_GravarDocumentos sistNFe, qtReg
      MON_GravarConsulta sUlt
      MON_Log qtReg & " documento(s) recebido(s) - NSU " & sUlt & " de " & sMax & "."
      If CDbl(sUlt) >= CDbl(sMax) Then Exit For
      sNSU = sUlt
   Case 137
      If CDbl(sUlt) > CDbl(sNSU) Then MON_GravarConsulta sUlt Else MON_GravarConsulta ""
      Exit For
   Case Else
      MON_GravarConsulta ""
      MON_Log "SEFAZ: " & cSt & " - " & sMotivo
      Exit For
   End Select
Next
End Sub

Private Sub MON_GravarConsulta(ByVal sUltNSU As String)
If Vazio(sUltNSU) Then
   SQLExecuta "UPDATE empresa SET DFeUltimaConsultaData = GETDATE(), DFeUltimaConsultaHora = GETDATE()"
Else
   SQLExecuta "UPDATE empresa SET UltimoNSU = " & sUltNSU & ", DFeUltimaConsultaData = GETDATE(), DFeUltimaConsultaHora = GETDATE()"
End If
End Sub

Private Sub MON_GravarDocumentos(sistNFe As snfe.Util, ByVal qtReg As Long)
Dim vTipo As Variant, vChave As Variant, vCNPJ As Variant, vNome As Variant, vEmi As Variant, vValor As Variant
Dim vTpNF As Variant, vSit As Variant, vConf As Variant
Dim i As Long, sChave As String, sSQL As String, e As String

If qtReg <= 0 Then Exit Sub
vTipo = sistNFe.retornoTipo
vChave = sistNFe.retornochNFe
vCNPJ = sistNFe.retornoCNPJ
vNome = sistNFe.retornoxNome
vEmi = sistNFe.retornodhEmi
vValor = sistNFe.retornovNF
vTpNF = sistNFe.retornotpNF
vSit = sistNFe.retornocSitNFe
vConf = sistNFe.retornocSitConf

For i = 0 To qtReg - 1
   sChave = MON_Txt(vChave(i))
   If Len(sChave) = 44 Then
      Select Case MON_Txt(vTipo(i))
      Case "resNFe"
         If Not MON_Existe(sChave) Then
            e = SQLExecuta(MON_SqlInserir("resNFe", sChave, vCNPJ(i), vNome(i), vEmi(i), vValor(i), vTpNF(i), vSit(i), vConf(i), False))
            If Not Vazio(e) Then MON_Log "Nota " & sChave & " não gravada: " & e
         End If
      Case "procNFe"
         If MON_Existe(sChave) Then
            sSQL = "UPDATE TbNotasDestinadasManifestacao SET Enviada = 1, SituacaoConfirmacao = 1, " & _
                   "DataHoraProcotolo = '" & MON_Sql(Left$(MON_Txt(vEmi(i)), 25)) & "' WHERE ChaveAcesso = '" & sChave & "'"
         Else
            sSQL = MON_SqlInserir("procNFe", sChave, vCNPJ(i), vNome(i), vEmi(i), vValor(i), vTpNF(i), 1, vConf(i), True)
         End If
         e = SQLExecuta(sSQL)
         If Not Vazio(e) Then MON_Log "Nota " & sChave & " não atualizada: " & e
         MON_CopiarXML sChave
      Case "procEventoNFe"
         'tpEvento em retornotpNF, descrição em retornoxNome, protocolo em retornocSitNFe, data em retornodhEmi
         sSQL = MON_SqlEvento(sChave, vTpNF(i), vNome(i), vSit(i), vEmi(i))
         If Len(sSQL) > 0 Then SQLExecuta sSQL
      End Select
   End If
Next
End Sub

'Texto gravado em TipoEvento para cada manifestação do destinatário (mesmo formato da tela Notas_Destinadas_Consulta).
'Aceita o código da tela ("0".."3") ou o tpEvento da SEFAZ (210200, 210210, 210220, 210240).
Public Function MON_TextoManifestacao(ByVal Tipo As String) As String
Select Case Tipo
Case "0", "210200": MON_TextoManifestacao = "0 - Confirmação da Operação"
Case "1", "210210": MON_TextoManifestacao = "1 - Ciência da Operação"
Case "2", "210220": MON_TextoManifestacao = "2 - Desconhecimento da Operação"
Case "3", "210240": MON_TextoManifestacao = "3 - Registro da Operação não Realizada"
End Select
End Function

'UPDATE para um evento recebido na distribuição (procEventoNFe); "" quando o evento não muda a nota.
' - 110111 (cancelamento pelo emitente): marca a nota como cancelada.
' - 210200/210220/210240 (manifestação definitiva): grava no formato da tela; não sobrepõe um cancelamento.
' - 210210 (ciência): só grava se a nota ainda não tem evento definitivo.
' - Demais (carta de correção etc.): não mexem na situação, para não apagar a manifestação já registrada.
Public Function MON_SqlEvento(ByVal Chave As String, ByVal TpEvento As Variant, ByVal xEvento As Variant, _
                              ByVal nProt As Variant, ByVal dhEvento As Variant) As String
Dim sTp As String, sFim As String
sTp = Left$(MON_Txt(TpEvento), 6)
sFim = "EventoProtocolo = " & MON_Inteiro(nProt) & ", EventoDataHora = '" & MON_Sql(Left$(MON_Txt(dhEvento), 25)) & "' " & _
       "WHERE ChaveAcesso = '" & MON_Sql(Chave) & "'"
Select Case sTp
Case "110111"
   MON_SqlEvento = "UPDATE TbNotasDestinadasManifestacao SET TipoEventoCodigo = '110111', " & _
                   "TipoEvento = '" & MON_Sql(Left$(MON_Txt(xEvento), 38)) & "', " & sFim
Case "210200", "210220", "210240"
   MON_SqlEvento = "UPDATE TbNotasDestinadasManifestacao SET " & _
                   "TipoEvento = CASE WHEN ISNULL(TipoEventoCodigo, '') = '110111' THEN TipoEvento ELSE '" & MON_TextoManifestacao(sTp) & "' END, " & _
                   "TipoEventoCodigo = CASE WHEN ISNULL(TipoEventoCodigo, '') = '110111' THEN TipoEventoCodigo ELSE '" & sTp & "' END, " & _
                   "Enviada = 1, " & sFim
Case "210210"
   MON_SqlEvento = "UPDATE TbNotasDestinadasManifestacao SET TipoEventoCodigo = '210210', " & _
                   "TipoEvento = '" & MON_TextoManifestacao(sTp) & "', " & sFim & _
                   " AND ISNULL(TipoEventoCodigo, '') NOT IN ('110111', '210200', '210220', '210240')" & _
                   " AND ISNULL(TipoEvento, '') NOT LIKE '[023] - %'"
End Select
End Function

'INSERT com todas as colunas obrigatórias (não depende dos DEFAULTs da tabela)
Private Function MON_SqlInserir(ByVal sTipo As String, ByVal sChave As String, ByVal vCNPJ As Variant, ByVal vNome As Variant, _
   ByVal vEmi As Variant, ByVal vValor As Variant, ByVal vTpNF As Variant, ByVal vSit As Variant, ByVal vProt As Variant, _
   ByVal bCompleta As Boolean) As String
MON_SqlInserir = "INSERT INTO TbNotasDestinadasManifestacao (CodigoConsulta, Data, TipoRetornoWS, TipoNotaWS, CNPJFornecedor, " & _
   "RazaoSocial, ChaveAcesso, DataEmissao, ValorNota, SituacaoNota, SituacaoConfirmacao, nSeqEvento, TextoCorrecao, TipoEvento, " & _
   "Justificativa, NumeroProtocolo, DataHoraProcotolo, Enviada, Enviar, XMLAutorizado, ImportadoXML, IdFilial, DataInclusao, " & _
   "TipoEventoCodigo, EventoProtocolo, EventoDataHora) VALUES (0, " & FdtSQL(Date) & ", '" & sTipo & "', " & _
   MON_Inteiro(vTpNF) & ", '" & MON_Sql(Left$(MON_SoDigitos(vCNPJ), 18)) & "', '" & MON_Sql(Left$(RemoveAcento(MON_Txt(vNome)), 60)) & "', '" & _
   sChave & "', " & MON_DataSQL(vEmi) & ", " & FSQL(MON_Valor(vValor), 2) & ", " & MON_Inteiro(vSit) & ", " & _
   IIf(bCompleta, "1", "0") & ", 0, '', '', '', " & MON_Inteiro(vProt) & ", '', " & IIf(bCompleta, "1", "0") & ", 0, '', 0, 1, GETDATE(), '', 0, '')"
End Function

Private Function MON_Existe(ByVal sChave As String) As Boolean
MON_Existe = (CLng(SQLExecutaRetorno("SELECT COUNT(*) AS r FROM TbNotasDestinadasManifestacao WHERE ChaveAcesso = '" & sChave & "'", "r", 0)) > 0)
End Function

'A DLL grava o procNFe em <DiretorioXML>\nfe\arquivos\DFe; a tela procura em nfe\arquivos\ConfRecebto
Private Sub MON_CopiarXML(ByVal sChave As String)
Dim sOrigem As String, sDestino As String
On Error GoTo deuErro
sOrigem = mDirXML & "nfe\arquivos\DFe\" & sChave & "-procNFe.xml"
If Existe(sOrigem) = 0 Then Exit Sub
sDestino = mDirXML & "nfe\arquivos\ConfRecebto"
If Len(Dir$(sDestino, vbDirectory)) = 0 Then MkDir sDestino
FileCopy sOrigem, sDestino & "\" & sChave & "-procNFe.xml"
Exit Sub
deuErro:
MON_Log "XML " & sChave & " não copiado: " & Err.Description
Err.Clear
End Sub

'---------- 3) ciência da operação ----------

Private Sub MON_EnviarCiencias(sistNFe As snfe.Util)
Dim r As ADODB.Recordset, aId() As Long, aChave() As String, qt As Long, i As Long
Dim dh As String, ret As String, arq As String, cSt As Long, cStEv As Long, sMotivo As String
Dim sProt As String, sDhReg As String, nOk As Long, sSQL As String

sSQL = "SELECT TOP " & MON_MAX_CIENCIAS & " IdNFConsulta, ChaveAcesso FROM TbNotasDestinadasManifestacao " & _
       "WHERE Enviada = 0 AND SituacaoNota = 1 AND LEN(ChaveAcesso) = 44 " & _
       "AND ISNULL(TipoEventoCodigo, '') <> '110111' AND LEFT(ISNULL(TipoEventoCodigo, ''), 2) <> '21' " & _
       "AND ISNULL(TipoEvento, '') NOT LIKE '[0-3] - %' " & _
       "AND ISNULL(CAST(TextoCorrecao AS varchar(12)), '') <> 'MonitorNFe: ' " & _
       "ORDER BY IdNFConsulta"
Set r = dbData.OpenRecordset(sSQL)
qt = 0
Do While Not r.EOF
   ReDim Preserve aId(qt): ReDim Preserve aChave(qt)
   aId(qt) = r("IdNFConsulta")
   aChave(qt) = Trim$("" & r("ChaveAcesso"))
   qt = qt + 1
   r.MoveNext
Loop
r.Close
Set r = Nothing
If qt = 0 Then Exit Sub

For i = 0 To qt - 1
   On Error GoTo erroEnvio
   dh = Format$(Date, "yyyy-mm-dd") & "T" & Format$(Now, "hh:mm:ss") & IIf(Vazio(UTC), "-03:00", UTC)
   ret = "": arq = ""
   Call sistNFe.ManifestacaoDestinatario("1", MON_IdLote(i), mCNPJ, aChave(i), dh, "", 1, ret, arq)
   cSt = sistNFe.retEnvEvento.cStat
   sMotivo = "" & sistNFe.retEnvEvento.xMotivo
   cStEv = 0: sProt = "": sDhReg = ""
   If cSt = 128 Then
      cStEv = sistNFe.retEnvEvento.retEvento.infEvento.cStat
      sMotivo = "" & sistNFe.retEnvEvento.retEvento.infEvento.xMotivo
      sProt = "" & sistNFe.retEnvEvento.retEvento.infEvento.nProt
      sDhReg = "" & sistNFe.retEnvEvento.retEvento.infEvento.dhRegEvento
   End If
   On Error GoTo 0

   If cStEv = 135 Or cStEv = 136 Or cStEv = 573 Then
      '573 = duplicidade de evento: a ciência já estava registrada na SEFAZ
      SQLExecuta "UPDATE TbNotasDestinadasManifestacao SET TipoEvento = '1 - Ciência da Operação', nSeqEvento = 1, " & _
                 "Data = GETDATE(), Enviada = 0, Enviar = 0, NumeroProtocolo = " & MON_Inteiro(sProt) & ", " & _
                 "DataHoraProcotolo = '" & MON_Sql(Left$(sDhReg, 25)) & "' WHERE IdNFConsulta = " & aId(i)
      nOk = nOk + 1
   ElseIf cSt = 0 Or cSt >= 9000 Then
      'comunicação/licença/certificado: tenta de novo na próxima execução
      MON_Log "Ciência interrompida: " & cSt & " - " & sMotivo
      Exit For
   Else
      SQLExecuta "UPDATE TbNotasDestinadasManifestacao SET TextoCorrecao = 'MonitorNFe: ciência recusada (" & _
                 IIf(cStEv <> 0, cStEv, cSt) & " - " & MON_Sql(Left$(sMotivo, 150)) & ") em " & Format$(Now, "dd/mm/yyyy") & "' " & _
                 "WHERE IdNFConsulta = " & aId(i)
      MON_Log "Ciência recusada para " & aChave(i) & ": " & IIf(cStEv <> 0, cStEv, cSt) & " - " & sMotivo
   End If
Next
If nOk > 0 Then MON_Log "Ciência da operação enviada para " & nOk & " nota(s)."
Exit Sub

erroEnvio:
MON_Log "Ciência interrompida: " & Err.Description
Err.Clear
If nOk > 0 Then MON_Log "Ciência da operação enviada para " & nOk & " nota(s)."
End Sub

'---------- 4) importação automática na Entrada de Estoque (configuracao Importar_XML = 1) ----------
'Importa com a mesma rotina da tela (Entrada_Estoque.ImportarXML, carregada sem aparecer, em modo silencioso): cadastra
'o fornecedor, vincula os produtos, lança as duplicatas no a_pagar e arquiva o XML em EntradasXML. Fica "Em Digitação":
'o estoque só muda quando o usuário confere e adiciona os itens. Só importa notas de compra emitidas a partir de
'IMPORTAR_XML_DESDE (gravada na primeira execução / quando a opção é ligada), para não puxar o histórico.
'ImportacaoStatus: 0 pendente, 1 importada (ou já lançada), 2 erro, 3 não importar (não é compra, outro destinatário...).

Private Sub MON_ImportarNotas()
Dim rs As ADODB.Recordset, sDesde As String, nOk As Long, nErro As Long, nIgn As Long
On Error GoTo deuErro

If CStr(SQLExecutaRetorno("SELECT TOP 1 config_valor AS r FROM configuracao WHERE config_nome = 'Importar_XML'", "r", "1")) <> "1" Then Exit Sub

sDesde = MON_Txt(SQLExecutaRetorno("SELECT TOP 1 config_valor AS r FROM configuracao WHERE config_nome = 'IMPORTAR_XML_DESDE'", "r", ""))
If Not IsDate(sDesde) Then
   sDesde = Format$(Date, "yyyy-mm-dd")
   SQLExecuta "IF NOT EXISTS (SELECT 1 FROM configuracao WHERE config_nome = 'IMPORTAR_XML_DESDE') " & _
              "INSERT INTO configuracao (config_nome, config_valor) VALUES ('IMPORTAR_XML_DESDE', '')"
   SQLExecuta "UPDATE configuracao SET config_valor = '" & sDesde & "' WHERE config_nome = 'IMPORTAR_XML_DESDE'"
   MON_Log "Importação automática ligada: notas emitidas a partir de " & Format$(Date, "dd/mm/yyyy") & "."
End If

Set rs = dbData.OpenRecordset("SELECT TOP " & MON_MAX_IMPORTACOES & " IdNFConsulta, ChaveAcesso " & _
       "FROM TbNotasDestinadasManifestacao " & _
       "WHERE ImportacaoStatus = 0 AND Enviada = 1 AND SituacaoNota = 1 AND LEN(ChaveAcesso) = 44 " & _
       "AND ISNULL(TipoEventoCodigo, '') NOT IN ('110111', '210220', '210240') " & _
       "AND ISNULL(TipoEvento, '') NOT LIKE '[23] - %' " & _
       "AND DataEmissao >= '" & Format$(CDate(sDesde), "yyyymmdd") & "' " & _
       "ORDER BY IdNFConsulta")
If rs Is Nothing Then Exit Sub
Do While Not rs.EOF
   Select Case MON_ImportarNota(CLng(rs("IdNFConsulta")), MON_Txt(rs("ChaveAcesso")))
   Case 1: nOk = nOk + 1
   Case 2: nErro = nErro + 1
   Case 3: nIgn = nIgn + 1
   End Select
   rs.MoveNext
Loop
rs.Close
Set rs = Nothing
If nOk + nErro + nIgn > 0 Then MON_Log "Importação automática: " & nOk & " importada(s), " & nErro & " com erro, " & nIgn & " não importada(s) (não é compra)."
Exit Sub

deuErro:
MON_Log "Erro na importação automática: " & Err.Description
Err.Clear
End Sub

'Importa uma nota. Retorna o status gravado (1, 2 ou 3) ou 0 quando ficou pendente (XML ainda não está na pasta).
Private Function MON_ImportarNota(ByVal vID As Long, ByVal sChave As String) As Long
Dim sArq As String, sMotivo As String, sErro As String, lCod As Long, bOk As Boolean
On Error GoTo deuErro

sArq = MON_ArquivoXML(sChave)
If Len(sArq) = 0 Then Exit Function

lCod = MON_EntradaDaChave(sChave)
If lCod > 0 Then
   MON_MarcarImportacao vID, 1, lCod, "Já lançada na Entrada de Estoque."
   MON_ImportarNota = 1
   Exit Function
End If

sMotivo = MON_MotivoNaoImportar(sArq)
If Len(sMotivo) > 0 Then
   MON_MarcarImportacao vID, 3, 0, sMotivo
   MON_ImportarNota = 3
   Exit Function
End If

Load Entrada_Estoque
Entrada_Estoque.ImportacaoSilenciosa = True
Entrada_Estoque.UltimoErroImportacao = ""
bOk = Entrada_Estoque.ImportarXML(sArq)
If bOk Then Entrada_Estoque.CriarEntradaEstoque Entrada_Estoque.CodigoNotaAtual
sErro = Entrada_Estoque.UltimoErroImportacao
Unload Entrada_Estoque

lCod = MON_EntradaDaChave(sChave)
If lCod > 0 Then
   MON_MarcarImportacao vID, 1, lCod, sErro
   MON_Log "Nota " & sChave & " importada na entrada " & lCod & IIf(Len(sErro) > 0, " (aviso: " & sErro & ")", "") & "."
   MON_ImportarNota = 1
Else
   If Len(sErro) = 0 Then sErro = "A importação não gravou a entrada."
   MON_MarcarImportacao vID, 2, 0, sErro
   MON_Log "Nota " & sChave & " não importada: " & sErro
   MON_ImportarNota = 2
End If
Exit Function

deuErro:
sErro = Err.Description
Err.Clear
On Error Resume Next
Unload Entrada_Estoque
MON_MarcarImportacao vID, 2, 0, sErro
MON_Log "Nota " & sChave & " não importada: " & sErro
MON_ImportarNota = 2
End Function

'XML completo (procNFe) da nota: ConfRecebto (cópia do monitor/tela) ou DFe (onde a DLL grava). "" se ainda não chegou.
Private Function MON_ArquivoXML(ByVal sChave As String) As String
Dim s As String
s = mDirXML & "nfe\arquivos\ConfRecebto\" & sChave & "-procNFe.xml"
If Existe(s) = 0 Then s = mDirXML & "nfe\arquivos\DFe\" & sChave & "-procNFe.xml"
If Existe(s) = 0 Then s = ""
MON_ArquivoXML = s
End Function

'Só importa nota de compra: emitida como saída pelo fornecedor, finalidade normal, para o CNPJ da empresa e com todos os
'CFOP de venda (x101-x123, x401-x405, x651-x656). Remessa, comodato, conserto, transferência, devolução etc. ficam
'para a importação manual. Retorna o motivo de não importar ("" = pode importar).
Private Function MON_MotivoNaoImportar(ByVal sArq As String) As String
Dim xml As New MSXML2.DOMDocument30, nl As MSXML2.IXMLDOMNodeList, i As Long, sCfop As String, nCfop As Long
Dim sDest As String

xml.async = False
If Not xml.Load(sArq) Then
   MON_MotivoNaoImportar = "XML inválido."
   Exit Function
End If
If MON_NoTexto(xml, "nfeProc/NFe/infNFe/ide/tpNF") <> "1" Then
   MON_MotivoNaoImportar = "Não é nota de saída do fornecedor (tpNF " & MON_NoTexto(xml, "nfeProc/NFe/infNFe/ide/tpNF") & ")."
   Exit Function
End If
If MON_NoTexto(xml, "nfeProc/NFe/infNFe/ide/finNFe") <> "1" Then
   MON_MotivoNaoImportar = "Finalidade da nota não é normal (finNFe " & MON_NoTexto(xml, "nfeProc/NFe/infNFe/ide/finNFe") & ")."
   Exit Function
End If
sDest = MON_SoDigitos(MON_NoTexto(xml, "nfeProc/NFe/infNFe/dest/CNPJ") & MON_NoTexto(xml, "nfeProc/NFe/infNFe/dest/CPF"))
If sDest <> mCNPJ Then
   MON_MotivoNaoImportar = "A empresa não é a destinatária da nota (transporte ou autorização de XML)."
   Exit Function
End If
Set nl = xml.selectNodes("nfeProc/NFe/infNFe/det/prod/CFOP")
If nl.Length = 0 Then
   MON_MotivoNaoImportar = "Nota sem itens."
   Exit Function
End If
For i = 0 To nl.Length - 1
   sCfop = Trim$(nl.Item(i).Text)
   nCfop = Val(Mid$(sCfop, 2))
   If Len(sCfop) <> 4 Or (Left$(sCfop, 1) <> "5" And Left$(sCfop, 1) <> "6") Or _
      Not ((nCfop >= 101 And nCfop <= 123) Or (nCfop >= 401 And nCfop <= 405) Or (nCfop >= 651 And nCfop <= 656)) Then
      MON_MotivoNaoImportar = "CFOP " & sCfop & " não é de compra (importe pela tela, se for o caso)."
      Exit Function
   End If
Next
End Function

Private Function MON_NoTexto(xml As MSXML2.DOMDocument30, ByVal sCaminho As String) As String
Dim no As MSXML2.IXMLDOMNode
Set no = xml.selectSingleNode(sCaminho)
If Not no Is Nothing Then MON_NoTexto = Trim$(no.Text)
End Function

Private Sub MON_MarcarImportacao(ByVal vID As Long, ByVal nStatus As Long, ByVal lCodNota As Long, ByVal sMotivo As String)
SQLExecuta "UPDATE TbNotasDestinadasManifestacao SET ImportacaoStatus = " & nStatus & ", " & _
           IIf(nStatus = 1, "ImportadoXML = 1, ", "") & _
           "CodigoNotaEntrada = " & lCodNota & ", ImportacaoData = GETDATE(), " & _
           "ImportacaoMotivo = '" & MON_Sql(Left$(sMotivo, 250)) & "' WHERE IdNFConsulta = " & vID
End Sub

Public Function MON_EntradaDaChave(ByVal sChave As String) As Long
MON_EntradaDaChave = Val(CStr(SQLExecutaRetorno("SELECT TOP 1 CodigoNota AS r FROM EntradaEstoque WHERE ChavedeAcesso = '" & _
                     MON_Sql(sChave) & "' ORDER BY CodigoNota DESC", "r", 0)))
End Function

'---------- exclusão / recusa: contas a pagar e entrada lançadas pela importação do XML ----------

'Duplicatas que a importação do XML lançou no a_pagar para a entrada (REF = CodigoNota, descrição "NF-e 000123 - DUP: ...").
Private Function MON_FiltroContas(ByVal lCodNota As Long) As String
MON_FiltroContas = "REF = '" & lCodNota & "' AND DESCRICAO LIKE 'NF-e " & Format$(lCodNota, "000000") & " - DUP:%'"
End Function

'Exclui as contas ainda em aberto; as já pagas ficam (Pagas retorna quantas). Retorna quantas foram excluídas.
Public Function MON_ExcluirContasEntrada(ByVal lCodNota As Long, Optional ByRef Pagas As Long) As Long
If lCodNota <= 0 Then Exit Function
Pagas = Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) AS r FROM a_pagar WHERE " & MON_FiltroContas(lCodNota) & _
        " AND ISNULL(STATUS, '') NOT LIKE '%PAGAR%'", "r", 0)))
MON_ExcluirContasEntrada = Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) AS r FROM a_pagar WHERE " & MON_FiltroContas(lCodNota) & _
        " AND ISNULL(STATUS, '') LIKE '%PAGAR%'", "r", 0)))
If MON_ExcluirContasEntrada > 0 Then
   SQLExecuta "DELETE FROM a_pagar WHERE " & MON_FiltroContas(lCodNota) & " AND ISNULL(STATUS, '') LIKE '%PAGAR%'"
End If
End Function

'Nota recusada (Desconhecimento / Operação não Realizada): exclui as contas a pagar em aberto e, se nenhum produto foi
'adicionado ao estoque, a entrada inteira. Retorna o texto do que foi feito ("" = a nota não estava lançada).
Public Function MON_ExcluirEntradaRecusada(ByVal sChave As String) As String
Dim lCod As Long, lCodEntrada As Long, nContas As Long, nPagas As Long, nAdic As Long, s As String
On Error GoTo deuErro

lCod = MON_EntradaDaChave(sChave)
If lCod = 0 Then Exit Function

nContas = MON_ExcluirContasEntrada(lCod, nPagas)
nAdic = Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) AS r FROM EntradaEstoqueItens WHERE CodigoNota = " & lCod & " AND Adicionada = 1", "r", 0)))
If nAdic > 0 Then
   s = "A entrada " & lCod & " já tem produtos adicionados ao estoque e NÃO foi excluída: faça a nota de devolução ao fornecedor."
Else
   lCodEntrada = Val(CStr(SQLExecutaRetorno("SELECT ISNULL(Cod_entrada, 0) AS r FROM EntradaEstoque WHERE CodigoNota = " & lCod, "r", 0)))
   SQLExecuta "DELETE FROM EntradaEstoqueItens WHERE CodigoNota = " & lCod
   SQLExecuta "DELETE FROM EntradaEstoque WHERE CodigoNota = " & lCod
   If lCodEntrada > 0 Then SQLExecuta "DELETE FROM produtos_entrada WHERE CODIGO = " & lCodEntrada
   SQLExecuta "UPDATE TbNotasDestinadasManifestacao SET CodigoNotaEntrada = 0 WHERE ChaveAcesso = '" & MON_Sql(sChave) & "'"
   s = "A entrada " & lCod & " (em digitação) foi excluída."
End If
s = s & " " & nContas & " conta(s) a pagar excluída(s)."
If nPagas > 0 Then s = s & " ATENÇÃO: " & nPagas & " conta(s) já paga(s) não foram excluídas."
MON_ExcluirEntradaRecusada = s
Exit Function

deuErro:
MON_ExcluirEntradaRecusada = "Erro ao excluir a entrada/contas da nota recusada: " & Err.Description
Err.Clear
End Function

Private Function MON_IdLote(ByVal i As Long) As Long
MON_IdLote = CLng(Format$(Now, "hhnnss")) * 100 + (i Mod 100)
End Function

'---------- utilitários ----------

Public Function MON_Txt(ByVal v As Variant) As String
On Error Resume Next
If IsNull(v) Or IsEmpty(v) Then MON_Txt = "" Else MON_Txt = Trim$(CStr(v))
End Function

Public Function MON_Sql(ByVal s As String) As String
MON_Sql = Replace(s, "'", "''")
End Function

Private Function MON_SoDigitos(ByVal v As Variant) As String
Dim s As String, i As Long, c As String
s = MON_Txt(v)
For i = 1 To Len(s)
   c = Mid$(s, i, 1)
   If c >= "0" And c <= "9" Then MON_SoDigitos = MON_SoDigitos & c
Next
End Function

Public Function MON_Inteiro(ByVal v As Variant) As String
MON_Inteiro = MON_SoDigitos(v)
If Len(MON_Inteiro) = 0 Then MON_Inteiro = "0"
End Function

'Valor vem do .NET no formato da máquina ("1234,56" ou "1234.56"); sem separador de milhar
Private Function MON_Valor(ByVal v As Variant) As Double
MON_Valor = Val(Replace(MON_Txt(v), ",", "."))
End Function

'Data vem "dd/mm/aaaa" (resNFe) ou "dd/mm/aaaa hh:mm:ss" (procNFe)
Private Function MON_DataSQL(ByVal v As Variant) As String
Dim s As String
s = MON_Txt(v)
If Len(s) >= 10 And Mid$(s, 3, 1) = "/" And Mid$(s, 6, 1) = "/" Then
   MON_DataSQL = "'" & Mid$(s, 7, 4) & Mid$(s, 4, 2) & Left$(s, 2) & "'"   'aaaammdd: igual em qualquer idioma do login
ElseIf IsDate(s) Then
   MON_DataSQL = FdtSQL(CDate(s))
Else
   MON_DataSQL = "NULL"
End If
End Function
