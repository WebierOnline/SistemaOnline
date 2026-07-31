Option Explicit

' ============================================================================
' ExportarXMLServidor.vbs
' Automatiza, na maquina servidor, a exportacao mensal de XML/PDF fiscal:
'  1) Garante que a pasta ExportarXML esta compartilhada na rede.
'  2) Gera o .rar do mes anterior (nome de maquina = SERVIDOR), se ainda nao
'     gerado neste mes (marcador local). O .rar tem 3 pastas dentro:
'       Enviados     - XMLs das notas autorizadas (procNFe\AAAAMM)
'       Cancelados   - XMLs de evento de cancelamento das notas EMITIDAS
'                      neste mes (buscado por chave de acesso, nao pela data
'                      do arquivo - uma nota emitida no ultimo dia do mes pode
'                      ser cancelada ja no mes seguinte, dentro das 24h)
'       Inutilizados - XMLs de inutilizacao de numeracao deste mes (agrupado
'                      pela propria data do evento, ja que nao ha nota emitida
'                      de referencia nesse caso)
'       Entradas     - XMLs de notas de fornecedor importadas via cmdImportarXML
'                      (Entrada_Estoque.frm), ja organizados por mes em
'                      EntradasXML\AAAAMM no momento da importacao
'  3) A partir do dia 2 de cada mes, se ainda nao enviado, junta todos os
'     .rar do mes (SERVIDOR + terminais que ja copiaram o deles) e manda um
'     unico email com todos em anexo para o contador.
' Roda todo dia, via Agendador de Tarefas do Windows (tarefa
' "OnlineCommerce_ExportarXMLServidor"), registrada automaticamente pelo
' proprio OnlineCommerce.exe no MDIForm_Load, somente na maquina servidor.
' ============================================================================

Dim fso, oShell
Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

Dim scriptDir
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

Dim logPath
logPath = scriptDir & "\ExportarXMLServidor_log.txt"

On Error Resume Next

RegistrarLog "----- Inicio da execucao -----"

' --- 1) Le o config.ini e verifica se esta maquina e o servidor ---
Dim vIP, vHost, vBarra, vComputerName

vIP = LerIni(scriptDir & "\config.ini", "IP_MAQUINA", "ip")
If vIP = "" Then vIP = "localhost\SQLEXPRESS2008"

vBarra = InStr(vIP, "\")
If vBarra > 0 Then
   vHost = Left(vIP, vBarra - 1)
Else
   vHost = vIP
End If

vComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

If Not (vHost = "." Or vHost = "127.0.0.1" Or LCase(vHost) = "localhost" Or LCase(vHost) = "(local)" Or UCase(vHost) = UCase(vComputerName)) Then
   RegistrarLog "Esta maquina nao e o servidor (ip=" & vIP & "). Exportacao automatica nao executada."
   WScript.Quit
End If

' --- 2) Conecta no banco de dados ---
Dim conn
Set conn = CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER={Sql Server};SERVER=" & vIP & _
   ";uid=sa;pwd=190106web;DATABASE=cyber_base;Connect Timeout=600;TRUSTED_CONNECTION=NO"
conn.Open

If Err.Number <> 0 Then
   RegistrarLog "Falha ao conectar no banco de dados: " & Err.Description
   WScript.Quit
End If

' --- 3) Le os dados da empresa ---
Dim rs
Set rs = conn.Execute("SELECT TOP 1 DiretorioXML, CNPJ, Razao, Fantasia, Estado, AmbienteNF, CertificadoDigital, " & _
   "NFCeIDToken, NFCeCSC, LicencaDLL, Email, caminho FROM empresa ORDER BY fantasia")

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

Dim vDiretorioXML, vCNPJ, vRazao, vFantasia, vEstado, vAmbienteNF, vCertificadoDigital
Dim vNFCeIDToken, vNFCeCSC, vLicencaDLL, vEmailEmpresa, vCaminhoDANFe

vDiretorioXML = rs("DiretorioXML").Value
vCNPJ = rs("CNPJ").Value
vRazao = NzStr(rs("Razao").Value, "OnLine Info")
vFantasia = NzStr(rs("Fantasia").Value, vRazao)
vEstado = NzStr(rs("Estado").Value, "PI")
vAmbienteNF = NzStr(rs("AmbienteNF").Value, "2")
vCertificadoDigital = NzStr(rs("CertificadoDigital").Value, "")
vNFCeIDToken = NzStr(rs("NFCeIDToken").Value, "1")
vNFCeCSC = NzStr(rs("NFCeCSC").Value, "")
vLicencaDLL = NzStr(rs("LicencaDLL").Value, "")
vEmailEmpresa = NzStr(rs("Email").Value, "financeiroonlineinfo@gmail.com")
vCaminhoDANFe = NzStr(rs("caminho").Value, "")
rs.Close

If Right(vDiretorioXML, 1) = "\" Then vDiretorioXML = Left(vDiretorioXML, Len(vDiretorioXML) - 1)

Dim vCaminhoExport
vCaminhoExport = vDiretorioXML & "\ExportarXML"

If Not fso.FolderExists(vCaminhoExport) Then
   fso.CreateFolder(vCaminhoExport)
   If Err.Number <> 0 Then
      RegistrarLog "Falha ao criar a pasta " & vCaminhoExport & ": " & Err.Description
      conn.Close
      WScript.Quit
   End If
End If

' --- 4) Garante o compartilhamento de rede da pasta ExportarXML ---
GarantirCompartilhamento vCaminhoExport

' --- 5) Calcula o mes-alvo (mes anterior) ---
Dim vMesAlvo, vAnoAlvo, vMesNomeAlvo, vAnoMes
vMesAlvo = Month(DateAdd("m", -1, Date))
vAnoAlvo = Year(DateAdd("m", -1, Date))
vMesNomeAlvo = NomeMes(vMesAlvo)
vAnoMes = CStr(vAnoAlvo) & Right("0" & vMesAlvo, 2)

' --- 6) Gera o .rar do mes anterior (nome de maquina = SERVIDOR), se ainda nao feito ---
Dim vMarcadorServidor, vCnpjLimpo, vNomeArquivoRar, vCaminhoRar
vMarcadorServidor = vCaminhoExport & "\.marker_servidor_" & vAnoMes & ".txt"
vCnpjLimpo = SoNumeros(vCNPJ)
vNomeArquivoRar = vCnpjLimpo & "_" & vMesNomeAlvo & vAnoAlvo & "_SERVIDOR.rar"
vCaminhoRar = vCaminhoExport & "\" & vNomeArquivoRar

If fso.FileExists(vMarcadorServidor) Then
   RegistrarLog "Exportacao do SERVIDOR referente a " & vMesNomeAlvo & "/" & vAnoAlvo & " ja foi realizada neste mes. Nada a fazer."
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

   ' Entradas: XMLs de notas de fornecedor importadas via Entrada_Estoque.frm (cmdImportarXML),
   ' ja arquivadas por mes de DataEmissao no momento da importacao - so precisa copiar a pasta.
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
         GravarMarcador vMarcadorServidor
      Else
         RegistrarLog "Falha ao gerar o .rar do SERVIDOR (" & vCaminhoRar & ")."
      End If
      fso.DeleteFolder vPastaTemp, True
   End If
End If

' --- 7) A partir do dia 2, envia para o contador (uma unica vez por mes) ---
Dim vMarcadorEnvio
vMarcadorEnvio = vCaminhoExport & "\.marker_envio_contador_" & vAnoMes & ".txt"

If Day(Date) < 2 Then
   RegistrarLog "Ainda nao chegou o dia 2 do mes - envio ao contador aguarda mais um dia de folga pros terminais."
ElseIf fso.FileExists(vMarcadorEnvio) Then
   RegistrarLog "Envio ao contador referente a " & vMesNomeAlvo & "/" & vAnoAlvo & " ja foi realizado. Nada a fazer."
Else
   Dim vEmailContador
   vEmailContador = ""
   Dim rsCont
   Set rsCont = conn.Execute("SELECT TOP 1 Email FROM TbContabilista")
   If Err.Number = 0 And Not rsCont.EOF Then
      vEmailContador = NzStr(rsCont("Email").Value, "")
   End If
   Err.Clear

   If vEmailContador = "" Then
      RegistrarLog "Nenhum email de contador cadastrado em TbContabilista. Envio nao realizado."
   Else
      Dim vPadrao, vArquivo, vAnexos(), vTotalAnexos
      vPadrao = "_" & vMesNomeAlvo & vAnoAlvo & "_"
      vTotalAnexos = 0
      ReDim vAnexos(100)

      Dim vArqs, vArq
      Set vArqs = fso.GetFolder(vCaminhoExport).Files
      For Each vArq In vArqs
         If LCase(fso.GetExtensionName(vArq.Name)) = "rar" Then
            If InStr(vArq.Name, vPadrao) > 0 Then
               vAnexos(vTotalAnexos) = vArq.Path
               vTotalAnexos = vTotalAnexos + 1
            End If
         End If
      Next

      If vTotalAnexos = 0 Then
         RegistrarLog "Nenhum arquivo .rar encontrado para " & vMesNomeAlvo & "/" & vAnoAlvo & ". Envio nao realizado."
      Else
         ReDim Preserve vAnexos(vTotalAnexos - 1)

         RegistrarLog "Enviando " & vTotalAnexos & " arquivo(s) para o contador (" & vEmailContador & ")..."

         If EnviarEmailContador(vAnexos, vTotalAnexos, vMesAlvo, vAnoAlvo) Then
            RegistrarLog "Email enviado com sucesso para o contador."
            GravarMarcador vMarcadorEnvio
         Else
            RegistrarLog "Falha ao enviar o email para o contador (ver EnviarEmailContador_log.txt) - tenta de novo na proxima execucao diaria"
         End If
      End If
   End If
End If

conn.Close
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

Function NzStr(valor, padrao)
   If IsNull(valor) Then
      NzStr = padrao
   ElseIf Trim(CStr(valor)) = "" Then
      NzStr = padrao
   Else
      NzStr = valor
   End If
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

Sub GarantirCompartilhamento(pasta)
   Dim resultado
   ' consulta se o compartilhamento ja existe (exit code 0 = existe)
   resultado = oShell.Run("cmd /c net share ExportarXML >nul 2>&1", 0, True)
   If resultado = 0 Then
      RegistrarLog "Pasta ExportarXML ja esta compartilhada na rede."
   Else
      resultado = oShell.Run("cmd /c net share ExportarXML=" & Chr(34) & pasta & Chr(34) & " /GRANT:Everyone,FULL >nul 2>&1", 0, True)
      If resultado = 0 Then
         RegistrarLog "Pasta ExportarXML compartilhada na rede com sucesso (" & pasta & ")."
      Else
         RegistrarLog "Falha ao compartilhar a pasta ExportarXML na rede (codigo " & resultado & "). Verifique permissoes/rede."
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

' O EmailEnviar da DLL snfe.Util exige um array ByRef fortemente tipado (String())
' como parametro - o VBScript so tem arrays de Variant e nao suporta isso via
' CreateObject tardio (lanca "Tipos incompativeis"). Testado e confirmado: o
' PowerShell CONSEGUE (usando [string[]] + [ref]), entao delega-se essa etapa
' pra um script PowerShell separado (EnviarEmailContador.ps1, mesma pasta).
' O envio e assincrono do lado do servidor SMTP do fornecedor - o retorno True
' so confirma que a chamada foi aceita, o email pode demorar alguns minutos
' pra chegar de fato (confirmado em teste real).
Function EnviarEmailContador(anexos, totalAnexos, vMesNum, vAnoNum)
   Dim vAnexosStr, i, vComandoPS, vExitCode, vShellPS

   EnviarEmailContador = False

   vAnexosStr = anexos(0)
   For i = 1 To totalAnexos - 1
      vAnexosStr = vAnexosStr & "|" & anexos(i)
   Next

   Set vShellPS = CreateObject("WScript.Shell")

   vComandoPS = Chr(34) & "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" & Chr(34) & _
      " -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & scriptDir & "\EnviarEmailContador.ps1" & Chr(34) & _
      " -Anexos " & Chr(34) & vAnexosStr & Chr(34) & " -Mes " & vMesNum & " -Ano " & vAnoNum

   vExitCode = vShellPS.Run(vComandoPS, 0, True)

   EnviarEmailContador = (vExitCode = 0)
End Function
