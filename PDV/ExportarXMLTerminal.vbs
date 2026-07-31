Option Explicit

' ============================================================================
' ExportarXMLTerminal.vbs
' Automatiza, em cada terminal PDV, a exportacao mensal de XML fiscal:
'  1) Garante que a unidade Z: esta mapeada para a pasta ExportarXML do
'     servidor (compartilhada por ExportarXMLServidor.vbs).
'  2) Gera o .rar do mes anterior, localmente, usando o nome desta maquina
'     (config.ini [DADOS_MAQUINA] maquina=), se ainda nao gerado neste mes
'     (marcador local). O .rar tem 3 pastas dentro:
'       Enviados     - XMLs das notas autorizadas (procNFe\AAAAMM)
'       Cancelados   - XMLs de evento de cancelamento das notas EMITIDAS
'                      neste mes (buscado por chave de acesso, nao pela data
'                      do arquivo - uma nota emitida no ultimo dia do mes pode
'                      ser cancelada ja no mes seguinte, dentro das 24h)
'       Inutilizados - XMLs de inutilizacao de numeracao deste mes (agrupado
'                      pela propria data do evento, ja que nao ha nota emitida
'                      de referencia nesse caso)
'       Entradas     - XMLs de notas de fornecedor importadas via cmdImportarXML
'                      (Entrada_Estoque.frm) - o OnlineCommerce.exe existe nos
'                      terminais tambem (raramente usado, mas pode acontecer de
'                      um cliente dar entrada de nota por um terminal), entao
'                      essa pasta e verificada aqui tambem, igual no servidor
'  3) Copia esse .rar para Z:\, se ainda nao copiado.
' Roda todo dia, via Agendador de Tarefas do Windows (tarefa
' "OnlinePDV_ExportarXML"), registrada automaticamente pelo proprio
' OnlinePDV.exe no Form_Load do PDV.frm, somente nos terminais (nao roda no
' servidor - la quem cuida disso e o ExportarXMLServidor.vbs).
' ============================================================================

Dim fso, oShell
Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

Dim scriptDir
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

Dim logPath
logPath = scriptDir & "\ExportarXMLTerminal_log.txt"

On Error Resume Next

RegistrarLog "----- Inicio da execucao -----"

' --- 1) Le o config.ini: ip do servidor + nome desta maquina ---
Dim vIP, vHost, vBarra, vComputerName, vNomeMaquina

vIP = LerIni(scriptDir & "\config.ini", "IP_MAQUINA", "ip")
If vIP = "" Then vIP = "localhost\SQLEXPRESS2008"

vBarra = InStr(vIP, "\")
If vBarra > 0 Then
   vHost = Left(vIP, vBarra - 1)
Else
   vHost = vIP
End If

vComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

If vHost = "." Or vHost = "127.0.0.1" Or LCase(vHost) = "localhost" Or LCase(vHost) = "(local)" Or UCase(vHost) = UCase(vComputerName) Then
   RegistrarLog "Esta maquina e o servidor - exportacao de terminal nao se aplica aqui."
   WScript.Quit
End If

vNomeMaquina = LerIni(scriptDir & "\config.ini", "DADOS_MAQUINA", "maquina")
If vNomeMaquina = "" Then vNomeMaquina = vComputerName

' --- 2) Garante que Z: esta mapeada para \\<servidor>\ExportarXML ---
Dim vCaminhoRede
vCaminhoRede = "\\" & vHost & "\ExportarXML"
GarantirMapeamento vCaminhoRede

' --- 3) Conecta no banco de dados ---
Dim conn
Set conn = CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER={Sql Server};SERVER=" & vIP & _
   ";uid=sa;pwd=190106web;DATABASE=cyber_base;Connect Timeout=600;TRUSTED_CONNECTION=NO"
conn.Open

If Err.Number <> 0 Then
   RegistrarLog "Falha ao conectar no banco de dados: " & Err.Description
   WScript.Quit
End If

' --- 4) Le os dados da empresa (CNPJ + DiretorioXML local deste terminal) ---
Dim rs
Set rs = conn.Execute("SELECT TOP 1 DiretorioXML, CNPJ FROM empresa ORDER BY fantasia")

If Err.Number <> 0 Then
   RegistrarLog "Falha ao consultar a tabela empresa: " & Err.Description
   conn.Close
   WScript.Quit
End If

If rs.EOF Then
   RegistrarLog "Nenhuma empresa cadastrada."
   conn.Close
   WScript.Quit
End If

Dim vDiretorioXML, vCNPJ
vDiretorioXML = rs("DiretorioXML").Value
vCNPJ = rs("CNPJ").Value
rs.Close

If Right(vDiretorioXML, 1) = "\" Then vDiretorioXML = Left(vDiretorioXML, Len(vDiretorioXML) - 1)

Dim vCaminhoExport
vCaminhoExport = vDiretorioXML & "\ExportarXML"

If Not fso.FolderExists(vCaminhoExport) Then
   fso.CreateFolder(vCaminhoExport)
   If Err.Number <> 0 Then
      RegistrarLog "Falha ao criar a pasta " & vCaminhoExport & ": " & Err.Description
      WScript.Quit
   End If
End If

' --- 5) Calcula o mes-alvo (mes anterior) ---
Dim vMesAlvo, vAnoAlvo, vMesNomeAlvo, vAnoMes
vMesAlvo = Month(DateAdd("m", -1, Date))
vAnoAlvo = Year(DateAdd("m", -1, Date))
vMesNomeAlvo = NomeMes(vMesAlvo)
vAnoMes = CStr(vAnoAlvo) & Right("0" & vMesAlvo, 2)

' --- 6) Gera o .rar do mes anterior (nome de maquina = config.ini), se ainda nao feito ---
Dim vMarcadorGeracao, vCnpjLimpo, vNomeArquivoRar, vCaminhoRar
vMarcadorGeracao = vCaminhoExport & "\.marker_" & vNomeMaquina & "_" & vAnoMes & ".txt"
vCnpjLimpo = SoNumeros(vCNPJ)
vNomeArquivoRar = vCnpjLimpo & "_" & vMesNomeAlvo & vAnoAlvo & "_" & vNomeMaquina & ".rar"
vCaminhoRar = vCaminhoExport & "\" & vNomeArquivoRar

If fso.FileExists(vMarcadorGeracao) Then
   RegistrarLog "Exportacao deste terminal referente a " & vMesNomeAlvo & "/" & vAnoAlvo & " ja foi realizada neste mes."
Else
   Dim vDiretorioOrigem, vPastaTemp
   vDiretorioOrigem = vDiretorioXML & "\nfe\arquivos\procNFe\" & vAnoMes
   vPastaTemp = vCaminhoExport & "\" & vAnoMes

   If fso.FolderExists(vPastaTemp) Then fso.DeleteFolder vPastaTemp, True
   fso.CreateFolder vPastaTemp

   ' Enviados: XMLs das notas autorizadas (ja organizados por mes pela propria DLL)
   If fso.FolderExists(vDiretorioOrigem) Then
      fso.CopyFolder vDiretorioOrigem, vPastaTemp & "\Enviados", True
   Else
      RegistrarLog "Nao existe a pasta de XMLs autorizados do periodo (" & vDiretorioOrigem & ")."
   End If

   ' Cancelados: XMLs de evento de cancelamento das notas EMITIDAS neste mes (mesmo que o
   ' cancelamento em si tenha sido feito ja no mes seguinte, dentro das 24h permitidas) -
   ' por isso a busca e por chave de acesso vinda do banco, e nao pela data do arquivo.
   Dim vChaves, vTotalChaves
   ColetarChavesCanceladas conn, vMesAlvo, vAnoAlvo, vChaves, vTotalChaves
   CopiarArquivosPorChave vDiretorioXML & "\nfe\arquivos\procEventoNFe", vPastaTemp & "\Cancelados", vChaves, vTotalChaves

   ' Inutilizados: nao existe "nota emitida" de referencia aqui (numero nunca foi
   ' transmitido), entao agrupa pela propria data do evento de inutilizacao.
   CopiarArquivosPorData vDiretorioXML & "\nfe\arquivos\Inutilizacao", vPastaTemp & "\Inutilizados", vMesAlvo, vAnoAlvo

   ' Entradas: XMLs de notas de fornecedor importadas via Entrada_Estoque.frm (cmdImportarXML) -
   ' o OnlineCommerce.exe existe nos terminais tambem (raramente usado), entao verifica aqui igual
   ' no servidor. Ja arquivadas por mes de DataEmissao no momento da importacao - so copia a pasta.
   Dim vDiretorioEntradas
   vDiretorioEntradas = vDiretorioXML & "\EntradasXML\" & vAnoMes
   If fso.FolderExists(vDiretorioEntradas) Then
      fso.CopyFolder vDiretorioEntradas, vPastaTemp & "\Entradas", True
   End If

   If fso.GetFolder(vPastaTemp).SubFolders.Count = 0 Then
      RegistrarLog "Nada para compactar neste periodo (nenhum Enviado/Cancelado/Inutilizado/Entrada encontrado)."
      fso.DeleteFolder vPastaTemp, True
   Else
      If CompactarPasta(vPastaTemp, vCaminhoRar) Then
         RegistrarLog "Arquivo gerado com sucesso: " & vCaminhoRar
         GravarMarcador vMarcadorGeracao
      Else
         RegistrarLog "Falha ao gerar o .rar deste terminal (" & vCaminhoRar & ")."
      End If
      fso.DeleteFolder vPastaTemp, True
   End If
End If

conn.Close

' --- 7) Copia o .rar gerado para a unidade Z: (servidor), se ainda nao copiado ---
Dim vMarcadorCopia, vDestinoZ
vMarcadorCopia = vCaminhoExport & "\.marker_copiado_" & vAnoMes & ".txt"
vDestinoZ = "Z:\" & vNomeArquivoRar

If fso.FileExists(vMarcadorCopia) Then
   RegistrarLog "Copia para o servidor referente a " & vMesNomeAlvo & "/" & vAnoAlvo & " ja foi realizada."
ElseIf Not fso.FileExists(vCaminhoRar) Then
   RegistrarLog "Arquivo .rar deste terminal ainda nao existe - copia para o servidor adiada."
Else
   fso.CopyFile vCaminhoRar, vDestinoZ, True
   If Err.Number = 0 Then
      RegistrarLog "Arquivo copiado com sucesso para " & vDestinoZ
      GravarMarcador vMarcadorCopia
   Else
      RegistrarLog "Falha ao copiar o arquivo para " & vDestinoZ & ": " & Err.Description & " (tenta de novo na proxima execucao diaria)"
      Err.Clear
   End If
End If

RegistrarLog "----- Fim da execucao -----"
WScript.Quit

' ============================================================================
' Funcoes auxiliares
' ============================================================================

Sub RegistrarLog(msg)
   Dim f
   Set f = fso.OpenTextFile(logPath, 8, True) '8 = ForAppending, True = cria se nao existir
   f.WriteLine Now & " - " & msg
   f.Close
End Sub

Function LerIni(caminho, secao, chave)
   Dim f, linha, secaoAtual, pos
   LerIni = ""
   If Not fso.FileExists(caminho) Then Exit Function

   Set f = fso.OpenTextFile(caminho, 1) '1 = ForReading
   secaoAtual = ""
   Do Until f.AtEndOfStream
      linha = Trim(f.ReadLine)
      If Left(linha, 1) = "[" And Right(linha, 1) = "]" Then
         secaoAtual = Mid(linha, 2, Len(linha) - 2)
      ElseIf LCase(secaoAtual) = LCase(secao) Then
         pos = InStr(linha, "=")
         If pos > 0 Then
            If LCase(Trim(Left(linha, pos - 1))) = LCase(chave) Then
               LerIni = Trim(Mid(linha, pos + 1))
               f.Close
               Exit Function
            End If
         End If
      End If
   Loop
   f.Close
End Function

Function SoNumeros(texto)
   Dim i, c, ret
   ret = ""
   For i = 1 To Len(texto)
      c = Mid(texto, i, 1)
      If InStr(".-/ ", c) = 0 Then ret = ret & c
   Next
   SoNumeros = ret
End Function

Function NomeMes(numero)
   Dim nomes(12)
   nomes(1) = "Janeiro"
   nomes(2) = "Fevereiro"
   nomes(3) = "Mar" & Chr(231) & "o"
   nomes(4) = "Abril"
   nomes(5) = "Maio"
   nomes(6) = "Junho"
   nomes(7) = "Julho"
   nomes(8) = "Agosto"
   nomes(9) = "Setembro"
   nomes(10) = "Outubro"
   nomes(11) = "Novembro"
   nomes(12) = "Dezembro"
   NomeMes = nomes(numero)
End Function

Sub GravarMarcador(caminho)
   Dim f
   Set f = fso.OpenTextFile(caminho, 2, True) '2 = ForWriting, True = cria se nao existir
   f.WriteLine "Gerado em: " & Now
   f.Close
End Sub

Sub GarantirMapeamento(caminhoRede)
   Dim resultado
   ' consulta se Z: ja esta mapeada (exit code 0 = existe)
   resultado = oShell.Run("cmd /c net use Z: >nul 2>&1", 0, True)
   If resultado = 0 Then
      RegistrarLog "Unidade Z: ja esta mapeada."
   Else
      resultado = oShell.Run("cmd /c net use Z: " & Chr(34) & caminhoRede & Chr(34) & " /PERSISTENT:YES >nul 2>&1", 0, True)
      If resultado = 0 Then
         RegistrarLog "Unidade Z: mapeada com sucesso para " & caminhoRede
      Else
         RegistrarLog "Falha ao mapear a unidade Z: para " & caminhoRede & " (codigo " & resultado & "). Verifique permissoes/rede."
      End If
   End If
End Sub

Function LocalizarCompressor()
   Dim ret, vWinRar, vWinZip
   vWinRar = ""
   vWinZip = ""

   ret = oShell.RegRead("HKEY_CLASSES_ROOT\WinRAR\shell\open\command\")
   If Err.Number = 0 And ret <> "" Then vWinRar = Left(ret, InStrRev(ret, " "))
   Err.Clear

   If vWinRar = "" Then
      ret = oShell.RegRead("HKEY_CLASSES_ROOT\WinZip\shell\open\command\")
      If Err.Number = 0 And ret <> "" Then vWinZip = Left(ret, InStrRev(ret, " "))
      Err.Clear
   End If

   If vWinRar <> "" Then
      LocalizarCompressor = vWinRar
   Else
      LocalizarCompressor = vWinZip
   End If
End Function

Function CompactarPasta(pastaOrigem, arquivoDestino)
   Dim vCompressor, vCmd, vTentativas
   Dim vTamanhoAnterior, vEstavel, vTentativasTamanho

   CompactarPasta = False

   vCompressor = LocalizarCompressor()
   If vCompressor = "" Then
      RegistrarLog "Nenhum compactador (WinRAR/WinZip) encontrado nesta maquina."
      Exit Function
   End If

   If fso.FileExists(arquivoDestino) Then fso.DeleteFile(arquivoDestino)
   Err.Clear

   vCmd = vCompressor & "a -ep1 " & Chr(34) & arquivoDestino & Chr(34) & " " & Chr(34) & pastaOrigem & Chr(34)
   oShell.Run vCmd, 0, False

   vTentativas = 0
   Do While Not fso.FileExists(arquivoDestino)
      WScript.Sleep 1000
      vTentativas = vTentativas + 1
      If vTentativas > 120 Then
         RegistrarLog "Tempo esgotado aguardando a compactacao (2 minutos)."
         Exit Function
      End If
   Loop

   vTamanhoAnterior = -1
   vEstavel = 0
   vTentativasTamanho = 0
   Do While vEstavel < 3
      WScript.Sleep 1000
      If fso.GetFile(arquivoDestino).Size = vTamanhoAnterior Then
         vEstavel = vEstavel + 1
      Else
         vTamanhoAnterior = fso.GetFile(arquivoDestino).Size
         vEstavel = 0
      End If
      vTentativasTamanho = vTentativasTamanho + 1
      If vTentativasTamanho > 300 Then Exit Do
   Loop

   CompactarPasta = True
End Function

' Busca no banco as chaves de acesso de NFe/NFCe EMITIDAS no mes-alvo e ja
' canceladas - usado para localizar os XMLs de evento de cancelamento
' correspondentes, independente da data em que o cancelamento em si ocorreu.
Sub ColetarChavesCanceladas(conexao, vMesAlvo, vAnoAlvo, chaves, totalChaves)
   Dim rsC, sql

   totalChaves = 0
   ReDim chaves(500)

   sql = "SELECT ChavedeAcesso FROM NotaFiscal WHERE Cancelada = 1 AND MONTH(DataEmissao) = " & vMesAlvo & _
         " AND YEAR(DataEmissao) = " & vAnoAlvo & " AND ChavedeAcesso IS NOT NULL AND ChavedeAcesso <> ''"
   Set rsC = conexao.Execute(sql)
   If Err.Number = 0 Then
      Do While Not rsC.EOF
         chaves(totalChaves) = Trim(rsC("ChavedeAcesso").Value)
         totalChaves = totalChaves + 1
         rsC.MoveNext
      Loop
   End If
   Err.Clear

   sql = "SELECT NFCeChaveAcesso FROM TbNFCe WHERE NFCeCancelada = 1 AND MONTH(DataEmissao) = " & vMesAlvo & _
         " AND YEAR(DataEmissao) = " & vAnoAlvo & " AND NFCeChaveAcesso IS NOT NULL AND NFCeChaveAcesso <> ''"
   Set rsC = conexao.Execute(sql)
   If Err.Number = 0 Then
      Do While Not rsC.EOF
         chaves(totalChaves) = Trim(rsC("NFCeChaveAcesso").Value)
         totalChaves = totalChaves + 1
         rsC.MoveNext
      Loop
   End If
   Err.Clear

   If totalChaves = 0 Then
      ReDim chaves(-1)
   Else
      ReDim Preserve chaves(totalChaves - 1)
   End If
End Sub

' Copia de pastaOrigem pra pastaDestino todo arquivo cujo nome contenha
' alguma das chaves informadas (os XMLs de evento trazem a chave de 44
' digitos embutida no meio do nome do arquivo).
Sub CopiarArquivosPorChave(pastaOrigem, pastaDestino, chaves, totalChaves)
   If totalChaves = 0 Then Exit Sub
   If Not fso.FolderExists(pastaOrigem) Then Exit Sub

   Dim arqs, arq, i, vCopiados
   vCopiados = 0
   Set arqs = fso.GetFolder(pastaOrigem).Files
   For Each arq In arqs
      For i = 0 To totalChaves - 1
         If InStr(arq.Name, chaves(i)) > 0 Then
            If Not fso.FolderExists(pastaDestino) Then fso.CreateFolder pastaDestino
            fso.CopyFile arq.Path, pastaDestino & "\" & arq.Name, True
            vCopiados = vCopiados + 1
            Exit For
         End If
      Next
   Next
   If vCopiados > 0 Then RegistrarLog "Cancelados: " & vCopiados & " arquivo(s) copiado(s) de " & pastaOrigem
End Sub

' Copia de pastaOrigem pra pastaDestino todo arquivo cuja data de modificacao
' caia no mes/ano informado (usado pra Inutilizados, que nao tem uma nota
' emitida de referencia pra buscar por chave).
Sub CopiarArquivosPorData(pastaOrigem, pastaDestino, vMesAlvo, vAnoAlvo)
   If Not fso.FolderExists(pastaOrigem) Then Exit Sub

   Dim arqs, arq, vCopiados
   vCopiados = 0
   Set arqs = fso.GetFolder(pastaOrigem).Files
   For Each arq In arqs
      If Month(arq.DateLastModified) = vMesAlvo And Year(arq.DateLastModified) = vAnoAlvo Then
         If Not fso.FolderExists(pastaDestino) Then fso.CreateFolder pastaDestino
         fso.CopyFile arq.Path, pastaDestino & "\" & arq.Name, True
         vCopiados = vCopiados + 1
      End If
   Next
   If vCopiados > 0 Then RegistrarLog "Inutilizados: " & vCopiados & " arquivo(s) copiado(s) de " & pastaOrigem
End Sub
