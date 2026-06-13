# -*- coding: utf-8 -*-
# Unifica cmdEnviarXML_Click (XML) + cmdEnviarPDF_Click (PDF) em um único botão:
# - Localiza XML pelo grid (col 8 = chave, col 11 = data protocolo)
# - Gera PDF via DANFeImprimir se não existir
# - Valida e-mail (vazio + sem '@')
# - Envia XML + PDF em um e-mail via nova sub EnviaEmailXMLPDF
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

patches = []

# ── Patch 1: Substituir cmdEnviarXML_Click pelo botão unificado ──────────────
OLD_CMD = (
    'Private Sub cmdEnviarXML_Click()\r\n'
    'If GridNotas.rows <= 1 Then\r\n'
    '    MsgBox "N\xe3o existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"\r\n'
    '    Exit Sub\r\n'
    'End If\r\n'
    '\r\n'
    'Dim NomeEmp As String, emailDestino As String, i As Integer, ComandoSQL As String\r\n'
    '\r\n'
    "'parte de encontrar o arquivo\r\n"
    'sSQL = "SELECT DiretorioXML, fantasia FROM Empresa"\r\n'
    'Set rEmpresa = dbData.OpenRecordset(sSQL)\r\n'
    '\r\n'
    'If Not rEmpresa.EOF Then\r\n'
    '    dirXML = IIf(Right(rEmpresa!DiretorioXML, 1) = "\\", rEmpresa!DiretorioXML, rEmpresa!DiretorioXML & "\\")\r\n'
    'End If\r\n'
    '\r\n'
    'IdNFProd = Val(GridNotas.TextMatrix(GridNotas.Row, 1))\r\n'
    '\r\n'
    'sSQL = "SELECT ChaveDEAcesso, DataEmissao FROM NotaFiscal WHERE CodigoNota = " & IdNFProd\r\n'
    'NFeChaveAcesso = SQLExecutaRetorno(sSQL, "ChaveDEAcesso", "")\r\n'
    'NFeDataHoraEnvio = SQLExecutaRetorno(sSQL, "DataEmissao", "")\r\n'
    '\r\n'
    'xCaminhoXML = dirXML & "nfe\\arquivos\\procNFe\\" & NFeChaveAcesso & "-procNFe.xml"\r\n'
    'anoEmes = dirXML & "nfe\\arquivos\\procNFe\\" & Format(NFeDataHoraEnvio, "yyyymm") & "\\"\r\n'
    '\r\n'
    'If Not Existe(anoEmes) Then MsgBox "N\xe3o existe a pasta referente ao m\xeas selecionado!", vbInformation, "Aviso do Sistema": Exit Sub\r\n'
    '\r\n'
    'If Not Existe(xCaminhoXML) Then xCaminhoXML = anoEmes & NFeChaveAcesso & "-procNFe.xml"\r\n'
    "'verifica se o arquivo existe\r\n"
    'If Not Existe(xCaminhoXML) Then MsgBox "N\xe3o existe o arquivo XML dessa venda nesse computador!", vbInformation, "Aviso do Sistema": Exit Sub\r\n'
    '\r\n'
    "'envio do arquivo\r\n"
    'emailDestino = InputBox("Informe o e-mail do destinat\xe1rio", "Envio de Email", "")\r\n'
    '\r\n'
    'If Not Vazio(emailDestino) Then\r\n'
    '   Call EnviaEmail(emailDestino, xCaminhoXML)\r\n'
    '   DoEvents\r\n'
    'End If\r\n'
    'End Sub'
)

NEW_CMD = (
    'Private Sub cmdEnviarXML_Click()\r\n'
    'If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub\r\n'
    'Dim vChave As String, dhDataProt As String, vDataProt As Date\r\n'
    'vChave = GridNotas.TextMatrix(GridNotas.Row, 8)\r\n'
    'dhDataProt = GridNotas.TextMatrix(GridNotas.Row, 11)\r\n'
    'If Len(dhDataProt) >= 10 Then vDataProt = CDate(Left(dhDataProt, 10))\r\n'
    '\r\n'
    'Dim emailDestino As String\r\n'
    'sSQL = "SELECT DiretorioXML, fantasia FROM Empresa"\r\n'
    'Set rEmpresa = dbData.OpenRecordset(sSQL)\r\n'
    '\r\n'
    'On Error GoTo deuErro\r\n'
    'Dim sistNFe As snfe.Util\r\n'
    'Set sistNFe = New snfe.Util\r\n'
    '\r\n'
    'dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", App.Path)\r\n'
    'xCaminhoXML = dirXML & "\\nfe\\arquivos\\procNFe\\" & Format(vDataProt, "yyyymm") & "\\" & vChave & "-procNFe.xml"\r\n'
    'xCaminhoPDF = dirXML & "\\nfe\\arquivos\\PDF\\NFe" & vChave & ".pdf"\r\n'
    '\r\n'
    'If Not Existe(xCaminhoXML) Then consultaNFe vChave, True\r\n'
    'If Not Existe(xCaminhoXML) Then\r\n'
    '    Set sistNFe = Nothing\r\n'
    '    MsgBox "N\xe3o foi encontrado o arquivo XML desta NF-e!", vbExclamation, "Aviso"\r\n'
    '    Exit Sub\r\n'
    'End If\r\n'
    '\r\n'
    'If Not Existe(xCaminhoPDF) Then\r\n'
    '    iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)\r\n'
    '    Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 0, xCancelada, True, "", True, False, False, False, True)\r\n'
    'End If\r\n'
    '\r\n'
    'On Error GoTo 0\r\n'
    'Set sistNFe = Nothing\r\n'
    '\r\n'
    'emailDestino = InputBox("Informe o e-mail do destinat\xe1rio", "Envio de Email", "")\r\n'
    'If Vazio(emailDestino) Then Exit Sub\r\n'
    'If InStr(1, emailDestino, "@") = 0 Then\r\n'
    '    MsgBox "E-mail inv\xe1lido! O endere\xe7o n\xe3o cont\xe9m \'@\'.", vbExclamation, "Aviso"\r\n'
    '    Exit Sub\r\n'
    'End If\r\n'
    '\r\n'
    'If Existe(xCaminhoPDF) Then\r\n'
    '    Call EnviaEmailXMLPDF(emailDestino, xCaminhoXML, xCaminhoPDF)\r\n'
    'Else\r\n'
    '    Call EnviaEmail(emailDestino, xCaminhoXML)\r\n'
    'End If\r\n'
    'DoEvents\r\n'
    'Exit Sub\r\n'
    '\r\n'
    'deuErro:\r\n'
    '    If InStr(1, Err.Description, "Exception") > 0 Then\r\n'
    '       iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)\r\n'
    '       Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 1, xCancelada, True, "", True, False, False, False, True)\r\n'
    '    Else\r\n'
    '       MsgBox Err.Description, vbInformation\r\n'
    '    End If\r\n'
    '    Err.Clear\r\n'
    '    Set sistNFe = Nothing\r\n'
    'End Sub'
)

patches.append(('cmdEnviarXML_unificado', OLD_CMD, NEW_CMD))

# ── Patch 2: Inserir EnviaEmailXMLPDF entre EnviaEmailPDF e EnviaEmail ────────
NEW_XMLPDF_SUB = (
    'Private Sub EnviaEmailXMLPDF(EmailPara As String, Anexo1 As String, Anexo2 As String)\r\n'
    'Dim emailDest As String, pathAnexo() As String, NomeRemetente As String, corpoEmail As String, emailCC() As String\r\n'
    'Dim sistNFe As snfe.Util\r\n'
    '\r\n'
    'On Error GoTo deuErro\r\n'
    '\r\n'
    'Set sistNFe = New snfe.Util\r\n'
    '\r\n'
    'emailDest = EmailPara\r\n'
    '\r\n'
    'ReDim emailCC(0)\r\n'
    'emailCC(0) = EmailPara\r\n'
    '\r\n'
    'If Vazio(emailDest) Then Exit Sub\r\n'
    '\r\n'
    'iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)\r\n'
    '\r\n'
    'If Not Vazio(Anexo2) And Existe(Anexo2) Then\r\n'
    '    ReDim pathAnexo(1)\r\n'
    '    pathAnexo(0) = Anexo1\r\n'
    '    pathAnexo(1) = Anexo2\r\n'
    'Else\r\n'
    '    ReDim pathAnexo(0)\r\n'
    '    pathAnexo(0) = Anexo1\r\n'
    'End If\r\n'
    '\r\n'
    'NomeRemetente = SQLExecutaRetorno("SELECT Fantasia FROM empresa", "Fantasia", rEmpresa!Fantasia)\r\n'
    'corpoEmail = "Seguem em anexo os arquivos XML e PDF da NFe emitida. " & _\r\n'
    '             "<br><br>" & _\r\n'
    '             "Atenciosamente, " & _\r\n'
    '             "<br><br>" & _\r\n'
    '             "#nome_emitente#"\r\n'
    'corpoEmail = Substitui(corpoEmail, "#nome_emitente#", SQLExecutaRetorno("SELECT RAZAO FROM empresa", "RAZAO"), SO_UM)\r\n'
    '\r\n'
    'If (emailDest <> Empty) Then\r\n'
    '   Screen.MousePointer = vbHourglass\r\n'
    '   iRetorno = sistNFe.EmailEnviar(emailDest, "Arquivo XML e PDF referente a NFe emitida ", corpoEmail, pathAnexo, emailCC)\r\n'
    '   Screen.MousePointer = vbDefault\r\n'
    'End If\r\n'
    '\r\n'
    'If iRetorno Then MsgBox "Email enviado com sucesso!", vbInformation + vbOKOnly, "EMAIL OK!"\r\n'
    '\r\n'
    'Set sistNFe = Nothing\r\n'
    '\r\n'
    'Exit Sub\r\n'
    'deuErro:\r\n'
    '    MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO: Envio Email"\r\n'
    '    Err.Clear\r\n'
    '    Set sistNFe = Nothing\r\n'
    'End Sub\r\n'
)

patches.append(('inserir_EnviaEmailXMLPDF',
    '\r\nEnd Sub\r\nPrivate Sub EnviaEmail(EmailPara As String, Anexo1 As String)\r\n',
    '\r\nEnd Sub\r\n\r\n' + NEW_XMLPDF_SUB + '\r\nPrivate Sub EnviaEmail(EmailPara As String, Anexo1 As String)\r\n'
))

# ── Verificar e aplicar ───────────────────────────────────────────────────────
all_ok = True
for name, old, new in patches:
    c = text.count(old)
    status = 'OK' if c == 1 else f'ERRO count={c}'
    print(f'{name}: {status}')
    if c != 1:
        all_ok = False

if all_ok:
    for _, old, new in patches:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('\nOK — cmdEnviarXML unificado (XML + PDF) + EnviaEmailXMLPDF adicionado')
else:
    print('\nABORTADO')
