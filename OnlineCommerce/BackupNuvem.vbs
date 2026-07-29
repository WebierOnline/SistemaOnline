Option Explicit

' ============================================================================
' BackupNuvem.vbs
' Envia automaticamente para o Google Drive o .rar de backup ja gerado
' localmente (Menu Sistema > Backup, dentro do OnlineCommerce.exe).
' Roda de forma independente do OnlineCommerce.exe, via Agendador de Tarefas
' do Windows (tarefa "OnlineCommerce_BackupNuvem", registrada automaticamente
' pelo proprio OnlineCommerce.exe no MDIForm_Load, somente na maquina servidor).
' ============================================================================

Dim fso, oShell
Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

Dim scriptDir
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

Dim logPath
logPath = scriptDir & "\BackupNuvem_log.txt"

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

If Not (vHost = "." Or LCase(vHost) = "localhost" Or LCase(vHost) = "(local)" Or UCase(vHost) = UCase(vComputerName)) Then
   RegistrarLog "Esta maquina nao e o servidor (ip=" & vIP & "). Backup automatico nao executado."
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
Set rs = conn.Execute("SELECT TOP 1 DiretorioXML, CNPJ, BackupDataHora FROM empresa ORDER BY fantasia")

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

Dim vDiretorioXML, vCNPJ, vBackupDataHora
vDiretorioXML = rs("DiretorioXML").Value
vCNPJ = rs("CNPJ").Value
If Right(vDiretorioXML, 1) <> "\" Then vDiretorioXML = vDiretorioXML & "\"

If Not IsNull(rs("BackupDataHora").Value) Then
   vBackupDataHora = rs("BackupDataHora").Value
End If
rs.Close

' --- 4) Ja foi feito backup hoje? ---
If Not IsEmpty(vBackupDataHora) Then
   If DateValue(vBackupDataHora) = DateValue(Now) Then
      RegistrarLog "Backup na nuvem ja foi realizado hoje (" & vBackupDataHora & "). Nada a fazer."
      conn.Close
      WScript.Quit
   End If
End If

' --- 5) Monta o caminho do .rar (mesmo padrao do cmdLogon no VB6) ---
Dim vCNPJLimpo, i, c, vNomeArquivo, vCaminhoBK, vCaminhoArquivo

vCNPJLimpo = ""
For i = 1 To Len(vCNPJ)
   c = Mid(vCNPJ, i, 1)
   If InStr(".-/ ", c) = 0 Then vCNPJLimpo = vCNPJLimpo & c
Next

vNomeArquivo = vCNPJLimpo & ".rar"
vCaminhoBK = vDiretorioXML & "backup"
vCaminhoArquivo = vCaminhoBK & "\" & vNomeArquivo

If Not fso.FileExists(vCaminhoArquivo) Then
   RegistrarLog "Nenhum backup local encontrado para enviar (" & vCaminhoArquivo & ")."
   conn.Close
   WScript.Quit
End If

' --- 6) Envia para o Google Drive (Service Account + Shared Drive "Backups") ---
Dim uploader, vFid

Set uploader = CreateObject("GoogleDriveUploader.Uploader")
If uploader Is Nothing Then
   RegistrarLog "Componente GoogleDriveUploader.Uploader nao esta registrado nesta maquina."
   conn.Close
   WScript.Quit
End If

uploader.UseInternalDialog = True
uploader.ApplicationName = "OnlineInfo"
uploader.CredentialsPath = scriptDir & "\NFE\backupclientes-488021-067a35f7c315.json"
uploader.AuthMode = "ServiceAccount"
uploader.SharedDriveId = "0ADpoMDm94fJRUk9PVA"

vFid = uploader.UploadFileToFolderName(vCaminhoArquivo, "BACKUP")

If Err.Number <> 0 Then
   RegistrarLog "Falha ao enviar o backup para a nuvem: " & Err.Description
   conn.Close
   WScript.Quit
End If

' --- 7) Atualiza a data/hora do ultimo backup ---
Dim vAgora, vAgoraSQL
vAgora = Now
vAgoraSQL = "'" & Year(vAgora) & "-" & Right("0" & Month(vAgora), 2) & "-" & Right("0" & Day(vAgora), 2) & _
   " " & Right("0" & Hour(vAgora), 2) & ":" & Right("0" & Minute(vAgora), 2) & ":" & Right("0" & Second(vAgora), 2) & "'"

conn.Execute "UPDATE empresa SET BackupDataHora = " & vAgoraSQL

conn.Close

RegistrarLog "Backup enviado para a nuvem com sucesso (" & vCaminhoArquivo & ")."
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
