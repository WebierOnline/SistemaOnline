Attribute VB_Name = "modSaidaPDF"
Option Explicit
'---------------------------------------------------------------------------
' Saida de documentos da OS por PDF/e-mail (tela frmSaidaDocumento).
' A impressora do e-mail vem do config.ini:
'     [IMPRESSORA_PDF_EMAIL]
'     impressora=Impressora PDF Email
' - Driver Bullzip: o destino do arquivo e configurado pelo COM do Bullzip (sem janela).
' - Driver Microsoft Print To PDF (Windows 10/11): a impressora e ligada a uma porta
'   de arquivo (PowerShell, exige administrador na 1a vez) e o PDF sai sem Salvar Como.
'---------------------------------------------------------------------------

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public PDFEmail_UltimoErro As String

'Nome da impressora do e-mail no config.ini ("" se nao configurada)
Public Function PDFEmail_Impressora() As String
Dim o As Ini

Set o = New Ini
o.Arquivo = appPathApp & "config.ini"
PDFEmail_Impressora = Trim$(o.LerTexto("IMPRESSORA_PDF_EMAIL", "impressora"))
Set o = Nothing
End Function

'Caminho fixo do PDF gerado (um por vez)
Public Function PDFEmail_Arquivo() As String
Dim sBase As String

sBase = Environ$("ProgramData")
If sBase = "" Then sBase = Environ$("ALLUSERSPROFILE")
If sBase = "" Then sBase = "C:"
On Error Resume Next
MkDir sBase & "\OnlineCommerce"
MkDir sBase & "\OnlineCommerce\PDF"
On Error GoTo 0
PDFEmail_Arquivo = sBase & "\OnlineCommerce\PDF\email.pdf"
End Function

'Nome do driver da impressora (Printer.DriverName do VB6 so devolve "winspool")
Private Function PDFEmail_NomeDriver(ByVal sImp As String) As String
Dim sh As Object

On Error Resume Next
Set sh = CreateObject("WScript.Shell")
PDFEmail_NomeDriver = CStr(sh.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers\" & Replace(sImp, "\", ",") & "\Printer Driver"))
Err.Clear
Set sh = Nothing
End Function

'Prepara a impressora para gravar sem janela em sArq. False + sMsg se nao der.
Public Function PDFEmail_Preparar(ByVal sImp As String, ByVal sArq As String, ByRef sMsg As String) As Boolean
Dim Prt As Printer
Dim bAchou As Boolean
Dim sDriver As String
Dim o As Object
Dim sPS As String
Dim sh As Object
Dim nRet As Long

sMsg = ""
If sImp = "" Then
    sMsg = "A impressora de PDF para e-mail n" & Chr$(227) & "o est" & Chr$(225) & " configurada." & vbCrLf & vbCrLf & _
           "No config.ini inclua:" & vbCrLf & "[IMPRESSORA_PDF_EMAIL]" & vbCrLf & "impressora=Impressora PDF Email"
    Exit Function
End If
If LCase$(sImp) = "impressora pdf" Then
    sMsg = "A impressora do e-mail n" & Chr$(227) & "o pode ser a ""Impressora PDF"" (ela abre o Salvar Como)." & vbCrLf & "Use outro nome no config.ini."
    Exit Function
End If

For Each Prt In Printers
    If Prt.DeviceName = sImp Then
        bAchou = True
        sDriver = PDFEmail_NomeDriver(sImp)
        Exit For
    End If
Next

On Error Resume Next
Kill sArq
On Error GoTo 0

If bAchou And InStr(1, sDriver, "Bullzip", vbTextCompare) > 0 Then
    On Error GoTo erroBullzip
    Set o = CreateObject("Bullzip.PDFPrinterSettings")
    o.SetPrinterName sImp
    o.SetValue "Output", sArq
    o.SetValue "ShowSettings", "never"
    o.SetValue "ShowSaveAS", "never"
    o.SetValue "ShowProgress", "no"
    o.SetValue "ShowProgressFinished", "no"
    o.SetValue "ShowPDF", "no"
    o.SetValue "ConfirmOverwrite", "no"
    o.WriteSettings True
    Set o = Nothing
    PDFEmail_Preparar = True
    Exit Function
End If

If bAchou And InStr(1, sDriver, "Microsoft Print To PDF", vbTextCompare) = 0 Then
    sMsg = "A impressora """ & sImp & """ usa o driver """ & sDriver & """, que n" & Chr$(227) & "o salva sozinha." & vbCrLf & "Use o Bullzip PDF Printer ou a Microsoft Print to PDF."
    Exit Function
End If

'Microsoft Print To PDF: cria/ajusta a impressora para a porta de arquivo
PDFEmail_Preparar = PDFEmail_ConfigurarMS(sImp, sArq, sMsg)
Exit Function

erroBullzip:
sMsg = "Falha ao configurar o Bullzip PDF Printer: " & Err.Description
Err.Clear
Exit Function

erroPS:
sMsg = "Falha ao configurar a impressora de PDF: " & Err.Description
Err.Clear
End Function

'Liga a impressora (driver Microsoft Print To PDF) a uma porta de arquivo via PowerShell.
'Tenta sem elevacao; se falhar, pede a elevacao (UAC) uma vez.
Private Function PDFEmail_ConfigurarMS(ByVal sImp As String, ByVal sArq As String, ByRef sMsg As String) As Boolean
Dim sPs1 As String
Dim sRes As String
Dim sh As Object
Dim nF As Integer
Dim sLinha As String
Dim sTexto As String
Dim t As Single
Dim sResp As String

On Error GoTo deuErro
sPs1 = Environ$("TEMP") & "\oc_pdfprinter.ps1"
sRes = Environ$("TEMP") & "\oc_pdfprinter.res"

nF = FreeFile
Open sPs1 For Output As #nF
Print #nF, "$ErrorActionPreference='Stop'"
Print #nF, "$res='" & Replace(sRes, "'", "''") & "'"
Print #nF, "try {"
Print #nF, "  $n='" & Replace(sImp, "'", "''") & "'; $p='" & Replace(sArq, "'", "''") & "'"
Print #nF, "  $pr = Get-Printer -Name $n -ErrorAction SilentlyContinue"
Print #nF, "  if (-not ($pr -and $pr.PortName -eq $p)) {"
Print #nF, "    if (-not (Get-PrinterPort -Name $p -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $p }"
Print #nF, "    if ($pr) { Set-Printer -Name $n -PortName $p } else { Add-Printer -Name $n -DriverName 'Microsoft Print To PDF' -PortName $p }"
Print #nF, "  }"
Print #nF, "  Set-Content -Path $res -Value 'ok'"
Print #nF, "} catch { Set-Content -Path $res -Value ('erro: ' + $_.Exception.Message) }"
Close #nF

Set sh = CreateObject("WScript.Shell")

'1a tentativa: sem elevacao
On Error Resume Next
Kill sRes
On Error GoTo deuErro
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & sPs1 & """", 0, True
sResp = PDFEmail_LerArquivo(sRes)

'2a tentativa: pede elevacao (UAC)
If Left$(sResp, 2) <> "ok" Then
    On Error Resume Next
    Kill sRes
    On Error GoTo deuErro
    CreateObject("Shell.Application").ShellExecute "powershell.exe", "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & sPs1 & """", "", "runas", 0
    t = Timer
    Do While Abs(Timer - t) < 40
        DoEvents
        Sleep 300
        sResp = PDFEmail_LerArquivo(sRes)
        If sResp <> "" Then Exit Do
    Loop
End If

On Error Resume Next
Kill sPs1
Kill sRes
On Error GoTo 0

If Left$(sResp, 2) = "ok" Then
    PDFEmail_ConfigurarMS = True
Else
    sMsg = "N" & Chr$(227) & "o foi poss" & Chr$(237) & "vel preparar a impressora """ & sImp & """." & vbCrLf & vbCrLf
    If sResp = "" Then
        sMsg = sMsg & "A permiss" & Chr$(227) & "o de administrador n" & Chr$(227) & "o foi concedida (ou a configura" & Chr$(231) & Chr$(227) & "o n" & Chr$(227) & "o respondeu)."
    Else
        sMsg = sMsg & sResp
    End If
End If
Exit Function

deuErro:
sMsg = "Falha ao configurar a impressora de PDF: " & Err.Description
Err.Clear
End Function

Private Function PDFEmail_LerArquivo(ByVal sArq As String) As String
Dim nF As Integer
Dim s As String

On Error Resume Next
If Dir$(sArq) = "" Then Exit Function
nF = FreeFile
Open sArq For Input As #nF
If Err.Number = 0 Then
    Line Input #nF, s
    Close #nF
    PDFEmail_LerArquivo = Trim$(s)
End If
Err.Clear
End Function

'Espera o PDF ficar pronto (existe, tamanho estavel e destravado)
Public Function PDFEmail_Aguardar(ByVal sArq As String, ByVal nSegundos As Long) As Boolean
Dim t As Single
Dim nTam As Long, nTamAnt As Long, nEst As Long
Dim nF As Integer

t = Timer
Do While Abs(Timer - t) < nSegundos
    DoEvents
    Sleep 300
    If Dir$(sArq) <> "" Then
        On Error Resume Next
        nTam = FileLen(sArq)
        On Error GoTo 0
        If nTam > 0 And nTam = nTamAnt Then
            nEst = nEst + 1
        Else
            nEst = 0
        End If
        nTamAnt = nTam
        If nEst >= 3 Then
            On Error Resume Next
            nF = FreeFile
            Err.Clear
            Open sArq For Binary Access Read Lock Read Write As #nF
            If Err.Number = 0 Then
                Close #nF
                On Error GoTo 0
                PDFEmail_Aguardar = True
                Exit Function
            End If
            Err.Clear
            On Error GoTo 0
        End If
    End If
Loop
End Function

'E-mail do cliente no cadastro ("" se nao tiver)
Public Function PDFEmail_EmailCliente(ByVal lCodCliente As Long) As String
PDFEmail_EmailCliente = Trim$(CStr(SQLExecutaRetorno("SELECT correio_eletronico FROM cliente WHERE codigo = " & lCodCliente, "correio_eletronico", "")))
End Function

'Pergunta o e-mail (ja preenchido com o do cadastro). "" = cancelou.
Public Function PDFEmail_PedirEmail(ByVal sSugestao As String) As String
Dim s As String

Do
    s = Trim$(InputBox("Informe o e-mail para enviar o PDF:", "Enviar por e-mail", sSugestao))
    If s = "" Then Exit Function
    If InStr(s, "@") > 1 And InStr(InStr(s, "@"), s, ".") > InStr(s, "@") + 1 And InStr(s, " ") = 0 Then
        PDFEmail_PedirEmail = s
        Exit Function
    End If
    MsgBox "E-mail inv" & Chr$(225) & "lido. Verifique e digite novamente.", vbExclamation, "Enviar por e-mail"
    sSugestao = s
Loop
End Function

'Envia o PDF pela DLL da NFe (mesmo caminho do Resumo Fiscal)
Public Function PDFEmail_Enviar(ByVal sEmailPara As String, ByVal sAnexo As String, ByVal sAssunto As String, ByVal sTexto As String) As Boolean
Dim sistNFe As snfe.Util
Dim pathAnexo() As String
Dim emailCC() As String
Dim corpoEmail As String
Dim vRet As Variant

On Error GoTo deuErro
PDFEmail_UltimoErro = ""
Set sistNFe = New snfe.Util
Call ConfiguraDLLNFeNFCe(55, "1", sistNFe)

ReDim emailCC(0)
emailCC(0) = sEmailPara
ReDim pathAnexo(0)
pathAnexo(0) = sAnexo

corpoEmail = sTexto & "<br><br>Atenciosamente, <br><br>" & _
             Trim$(CStr(SQLExecutaRetorno("SELECT TOP 1 RAZAO FROM empresa", "RAZAO", "")))

vRet = sistNFe.EmailEnviar(sEmailPara, sAssunto, corpoEmail, pathAnexo, emailCC)
If vRet Then
    PDFEmail_Enviar = True
Else
    PDFEmail_UltimoErro = "A DLL de e-mail retornou falha no envio."
End If
Set sistNFe = Nothing
Exit Function

deuErro:
PDFEmail_UltimoErro = Err.Description
Err.Clear
Set sistNFe = Nothing
End Function

'Grava o e-mail no cadastro do cliente (coluna tem 30 posicoes). True se atualizou.
Public Function PDFEmail_AtualizarCliente(ByVal lCodCliente As Long, ByVal sEmail As String) As Boolean
If lCodCliente <= 0 Then Exit Function
If Len(sEmail) > 30 Then Exit Function
On Error GoTo deuErro
dbData.Execute "UPDATE cliente SET correio_eletronico = '" & Replace(sEmail, "'", "''") & "' WHERE codigo = " & lCodCliente
PDFEmail_AtualizarCliente = True
Exit Function
deuErro:
Err.Clear
End Function

'---------------------------------------------------------------------------
' WhatsApp: o PDF sai pela mesma impressora do e-mail ([IMPRESSORA_PDF_EMAIL]) e vai pelo WhatsApp da
' empresa. Celular, nome e emoji ficam em modZapEnvio.bas (Compartilhado).
'---------------------------------------------------------------------------

'Mensagem do WhatsApp para o documento da OS (modelos aprovados pelo usuário em 2026-11-18).
'nTipo: 1 = Orçamento, 2 = Pedido (serviço concluído), 3 = Entrada, 4 = Garantia. sTipoOS = vTipoOS da tela.
'*texto* = negrito no WhatsApp. Emojis montados por PDFZap_Emoji (o .bas é Windows-1252).
Public Function PDFZap_Mensagem(ByVal nTipo As Integer, ByVal lOS As Long, ByVal sNome As String, _
                                ByVal sEmpresa As String, ByVal sTipoOS As String) As String
Dim L As String, sOla As String, sEmp As String, sAssin As String, sOS As String
Dim sEntrada As String, sFimEntrada As String, sVeic As String

L = vbLf
sOla = "Olá" & IIf(sNome <> "", ", *" & sNome & "*", "") & "!"
sEmp = IIf(sEmpresa <> "", "*" & sEmpresa & "*", "")
sOS = "*Ordem de Serviço nº " & lOS & "*"

Select Case sTipoOS
    Case "Automóveis"
        sVeic = PDFZap_Emoji(&H1F697)                                   'carro
        sEntrada = "Confirmamos a entrada do seu veículo em nossa oficina!"
        sFimEntrada = "Qualquer novidade sobre o diagnóstico ou andamento, avisaremos você por aqui. Conte com a gente! "
    Case "Motocicletas"
        sVeic = PDFZap_Emoji(&H1F3CD) & ChrW$(&HFE0F)                   'moto
        sEntrada = "Confirmamos a entrada da sua moto em nossa oficina!"
        sFimEntrada = "Qualquer novidade sobre o diagnóstico ou andamento, avisaremos você por aqui. Conte com a gente! "
    Case "Informática", "Celular", "Climatização"
        sEntrada = "Confirmamos a entrada do seu equipamento!"
        sFimEntrada = "Qualquer novidade sobre o diagnóstico ou andamento, avisaremos você por aqui. Conte com a gente! "
    Case Else
        sEntrada = "Confirmamos a abertura do seu atendimento!"
        sFimEntrada = "Qualquer novidade sobre o andamento, avisaremos você por aqui. Conte com a gente! "
End Select
sAssin = sEmp
If sAssin <> "" And sVeic <> "" Then sAssin = sAssin & " " & PDFZap_Emoji(&H1F527) & sVeic       'chave inglesa + veículo

Select Case nTipo
    Case 1  'Orçamento: foco em aprovação
        PDFZap_Mensagem = sOla & " Tudo bem? " & PDFZap_Emoji(&H1F44B) & L & L & _
            "Segue em anexo o orçamento detalhado referente à sua " & sOS & ". " & PDFZap_Emoji(&H1F4C4) & L & L & _
            "Dê uma olhada nos itens e valores. Caso tenha qualquer dúvida ou já queira autorizar o serviço para agilizar, é só me responder por aqui!" & _
            IIf(sAssin <> "", L & L & "Atenciosamente," & L & sAssin, "")
    Case 3  'Entrada: foco em transparência e segurança
        PDFZap_Mensagem = sOla & IIf(sEmp <> "", " Seja bem-vindo(a) à " & sEmp & ".", "") & " " & PDFZap_Emoji(&H1F44B) & L & L & _
            sEntrada & " Segue em anexo o comprovante de abertura da " & sOS & " com os detalhes iniciais. " & PDFZap_Emoji(&H1F4CB) & L & L & _
            sFimEntrada & IIf(sVeic <> "", sVeic, "") & ChrW$(&H2728)
    Case 2  'Pedido / serviço concluído: foco em pós-venda
        PDFZap_Mensagem = sOla & " Tudo certo? " & PDFZap_Emoji(&H1F60A) & L & L & _
            "Muito obrigado pela confiança em nosso trabalho! Segue em anexo o resumo da sua " & sOS & " para o seu controle. " & _
            PDFZap_Emoji(&H1F4C4) & ChrW$(&H2705) & L & L & _
            "Se precisar de qualquer suporte ou revisão futura, estamos sempre à disposição. Até a próxima!" & _
            IIf(sAssin <> "", L & L & sAssin, "")
    Case Else  'Garantia
        PDFZap_Mensagem = sOla & " Tudo certo? " & PDFZap_Emoji(&H1F60A) & L & L & _
            "Segue em anexo o *termo de garantia* da sua " & sOS & ". Guarde para a sua segurança. " & _
            PDFZap_Emoji(&H1F6E1) & ChrW$(&HFE0F) & L & L & _
            "Se precisar de qualquer suporte dentro da garantia, é só chamar por aqui. Estamos à disposição!" & _
            IIf(sAssin <> "", L & L & sAssin, "")
End Select
End Function

'Mensagem do WhatsApp da NFS-e (tela frmNFSe), no estilo dos modelos da OS.
Public Function PDFZap_MensagemNFSe(ByVal sNome As String, ByVal sNumero As String, ByVal sValor As String, _
                                    ByVal bCancelada As Boolean, ByVal sEmpresa As String) As String
Dim L As String

L = vbLf
PDFZap_MensagemNFSe = "Olá" & IIf(sNome <> "", ", *" & sNome & "*", "") & "! Tudo bem? " & PDFZap_Emoji(&H1F44B) & L & L & _
    "Segue em anexo a *Nota Fiscal de Serviço eletrônica (NFS-e) nº " & sNumero & "*, no valor de *R$ " & sValor & "*. " & _
    PDFZap_Emoji(&H1F4C4) & L & L & _
    IIf(bCancelada, ChrW$(&H26A0) & ChrW$(&HFE0F) & " *Atenção: esta NFS-e foi CANCELADA.*" & L & L, "") & _
    "A autenticidade pode ser conferida pelo QR Code do PDF ou no portal nacional da NFS-e (www.nfse.gov.br)." & _
    IIf(sEmpresa <> "", L & L & "Atenciosamente," & L & "*" & sEmpresa & "*", "")
End Function
