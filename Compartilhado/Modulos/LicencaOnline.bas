Attribute VB_Name = "LicencaOnline"
'=============================================================================
' LICENÇA ONLINE - mesmo servidor e mesmas regras do Online Food.
'
' O servidor de licenças emite uma licença assinada (ECDSA P-256). Aqui ela só
' é conferida com a chave pública e guardada em od.Licenca: sem internet, vale
' a última licença guardada. Pagamento pelo Pix ("Pagar agora") libera sozinho.
' Códigos de liberação offline (8 letras) são gerados no painel de licenças.
'
' Usa somente o que já existe no Windows (nada para instalar ou registrar):
'   - MSXML2.ServerXMLHTTP / MSXML2.DOMDocument por CreateObject (sem referência)
'   - bcrypt.dll (SHA-256, HMAC, ECDSA, números aleatórios) - Windows 7 ou mais novo
'
' Usado pelo OnlineCommerce, PDV e Ordem de Serviço (mesma licença para todos).
'=============================================================================
Option Explicit

Private Const URL_PADRAO As String = "https://licencas.onlinecommerce.com.br"
Private Const DIAS_AVISO As Long = 3
'Primeira instalação sem conseguir falar com o servidor: não bloqueia por estes dias.
Private Const DIAS_INSTALACAO As Long = 3
Private Const MAX_ERROS_CODIGO As Long = 5
Private Const SUPORTE As String = "(89) 9 9427-5280"

'Chave pública do servidor de licenças (X e Y da curva P-256, em hexadecimal).
Private Const CHAVE_PUBLICA_XY As String = "0B4DB14F451429ACA533D32E7F39A874F838A810AB270A94E6950A494219F0AB" & _
                                           "AB8B7E4A8E18669310C2231454CEDE4FE3D5C3B501ED5D47701EAE09B4C10D5A"
Private Const ALFABETO_CODIGO As String = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"

Public Enum LicEstado
   licLiberada = 0
   licAviso = 1
   licBloqueada = 2
   licSemLicenca = 3
   licCnpjDiferente = 4
End Enum

Private Type LicDados
   Cnpj As String
   Cliente As String
   Plano As String
   Vencimento As Date
   DataBloqueio As Date
   ChaveOffline As String
   Abertas As String          'ex.: "09/2026, 10/2026"
   ValorAberto As Currency
End Type

'Resultado da última avaliação (usado pela tela Licenca_Bloqueio e pela barra de status).
Public LicEstadoAtual As LicEstado
Public LicBloqueiaEm As Date
Public LicDiasParaBloqueio As Long
Public LicUltimaMensagem As String
'Segundos que a chamada atual ao servidor já está esperando (a tela mostra enquanto aguarda).
Public LicSegundosEspera As Long
'Detalhe da última falha de comunicação (mostrado junto da mensagem de erro, para o suporte).
Public LicErroHttp As String

Private mDados As LicDados
Private mTemDados As Boolean
Private mEmInstalacao As Boolean
Private mTabelasOk As Boolean
Private mErrosCodigo As Long
Private mErrosAte As Date

Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As Long) As Long
Private Declare Function BCryptImportKeyPair Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, ByRef phKey As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptVerifySignature Lib "bcrypt.dll" (ByVal hKey As Long, ByVal pPaddingInfo As Long, ByVal pbHash As Long, ByVal cbHash As Long, ByVal pbSignature As Long, ByVal cbSignature As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt.dll" (ByVal hKey As Long) As Long
Private Declare Function BCryptGenRandom Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal pbBuffer As Long, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
Private Declare Function LicMultiByteToWideChar Lib "kernel32" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

'=============================================================================
' ENTRADA PRINCIPAL - chamar no login. True = pode continuar; False = fechar o sistema.
'=============================================================================
Public Function VerificarLicencaOnline(Optional ByVal CodUsuario As Long = 0) As Boolean
   On Error GoTo Falha
   Dim msg As String

   VerificarLicencaOnline = True
   If Not LicGarantirTabelas() Then Exit Function

   LicSincronizar msg
   Select Case LicAvaliar()
      Case licLiberada
         'segue normalmente
      Case licAviso
         LicMostrarTela True, CodUsuario
      Case Else
         VerificarLicencaOnline = LicMostrarTela(False, CodUsuario)
   End Select
   Exit Function

Falha:
   'Erro inesperado aqui nunca trava o cliente.
   Debug.Print "VerificarLicencaOnline: " & Err.Number & " " & Err.Description
   VerificarLicencaOnline = True
End Function

Private Function LicMostrarTela(ByVal ModoAviso As Boolean, ByVal CodUsuario As Long) As Boolean
   Load Licenca_Bloqueio
   Licenca_Bloqueio.Preparar ModoAviso, CodUsuario
   Licenca_Bloqueio.Show vbModal
   LicMostrarTela = Licenca_Bloqueio.pDesbloqueado
   Unload Licenca_Bloqueio
End Function

'Botão "Licença" do PDV: pagar pelo Pix a mensalidade que ainda vai vencer, antes do bloqueio.
Public Sub LicPagarMensalidade(Optional ByVal CodUsuario As Long = 0)
   On Error GoTo Falha
   If Not LicGarantirTabelas() Then Exit Sub
   LicAvaliar
   Load Licenca_Bloqueio
   Licenca_Bloqueio.PrepararPagamento CodUsuario
   Licenca_Bloqueio.Show vbModal
   Unload Licenca_Bloqueio
   Exit Sub
Falha:
   MsgBox "Não foi possível abrir o pagamento da mensalidade: " & Err.Description, vbExclamation, "Licença do sistema"
End Sub

'Mensagem da tela no modo "pagar mensalidade" (sem aviso nem bloqueio).
Public Function LicMensagemPagamento() As String
   Dim s As String
   s = "Pague a sua mensalidade pelo Pix, mesmo antes do vencimento." & vbCrLf & _
       "Clique em PAGAR AGORA: aparece o QR Code com o mês/ano e o valor que está sendo pago." & vbCrLf & _
       "Assim que o pagamento for confirmado, a licença é atualizada sozinha."
   If LicBloqueiaEm > 0 Then s = s & vbCrLf & vbCrLf & "Licença atual válida até " & Format$(LicBloqueiaEm, "dd/mm/yyyy") & "."
   s = s & LicTextoAbertas()
   LicMensagemPagamento = s
End Function

'Texto para a barra de status da tela principal.
Public Function LicTextoStatus() As String
   If LicBloqueiaEm > 0 Then LicTextoStatus = "SUA LICENÇA VENCE EM: " & Format$(LicBloqueiaEm, "dd/mm/yyyy")
End Function

'Mensagem da tela de aviso/bloqueio conforme a última avaliação.
Public Function LicMensagemTela() As String
   Dim s As String
   Select Case LicEstadoAtual
      Case licAviso
         If LicDiasParaBloqueio >= 1 Then
            s = "Sua licença será bloqueada em " & LicDiasParaBloqueio & IIf(LicDiasParaBloqueio = 1, " dia", " dias") & _
                " (" & Format$(LicBloqueiaEm, "dd/mm/yyyy") & ")."
         Else
            s = "Mensalidade vencida. O sistema será bloqueado no próximo dia útil."
         End If
         s = s & LicTextoAbertas() & vbCrLf & vbCrLf & _
             "Pague agora pelo Pix (botão PAGAR AGORA): a liberação é automática."
      Case licBloqueada
         s = "O sistema está bloqueado por falta de pagamento." & LicTextoAbertas() & vbCrLf & vbCrLf & _
             "Pague pelo Pix (botão PAGAR AGORA): assim que o pagamento for confirmado, o sistema libera sozinho." & vbCrLf & _
             "Se já pagou de outra forma, peça o código de liberação ao financeiro: " & SUPORTE & "."
      Case licCnpjDiferente
         s = "A licença guardada é de outro CNPJ. Confira o CNPJ no cadastro da empresa e clique em VERIFICAR NOVAMENTE." & vbCrLf & _
             "Suporte: " & SUPORTE & "."
      Case licSemLicenca
         s = "Este computador ainda não tem a licença do sistema." & vbCrLf & _
             "Conecte à internet e clique em VERIFICAR NOVAMENTE. Suporte: " & SUPORTE & "."
   End Select
   If Len(LicUltimaMensagem) > 0 And LicEstadoAtual <> licLiberada Then s = s & vbCrLf & vbCrLf & "(" & LicUltimaMensagem & ")"
   LicMensagemTela = s
End Function

Private Function LicTextoAbertas() As String
   If mTemDados And Len(mDados.Abertas) > 0 Then
      LicTextoAbertas = vbCrLf & "Em aberto: " & mDados.Abertas & " - R$ " & Format$(mDados.ValorAberto, "#,##0.00") & "."
   End If
End Function

'=============================================================================
' AVALIAÇÃO (sem rede) - mesmas regras do AvaliadorLicenca do Online Food
'=============================================================================
Public Function LicAvaliar() As LicEstado
   Dim r As ADODB.Recordset
   Dim agora As Date, hoje As Date, bloqueiaEm As Date
   Dim token As String, liberadoAte As Variant, criadoEm As Variant
   Dim podeHoje As Boolean, estado As LicEstado, relogioAtrasado As Boolean

   LicBloqueiaEm = 0
   LicDiasParaBloqueio = 0
   mEmInstalacao = False

   Set r = dbData.OpenRecordset("SELECT Token, UltimaDataVista, LiberadoAte, AtualizadoEm FROM od.Licenca WHERE Id = 1;")
   agora = Now
   If Not r.EOF Then
      token = "" & r("Token")
      liberadoAte = r("LiberadoAte")
      criadoEm = r("AtualizadoEm")
      'Relógio atrasado não "volta no tempo": usa a maior data/hora já vista.
      If Not IsNull(r("UltimaDataVista")) Then
         If r("UltimaDataVista") > agora Then
            agora = r("UltimaDataVista")
            relogioAtrasado = True
         End If
      End If
   Else
      liberadoAte = Null
      criadoEm = Null
   End If
   r.Close
   Set r = Nothing
   If Not relogioAtrasado Then
      dbData.Execute "UPDATE od.Licenca SET UltimaDataVista = '" & LicSqlData(agora) & "' WHERE Id = 1;"
   End If

   hoje = DateValue(agora)
   podeHoje = LicPodeBloquear(hoje)
   mTemDados = LicLerToken(token, mDados)

   If Not mTemDados Then
      estado = IIf(podeHoje, licSemLicenca, licLiberada)
      'Acabou de instalar e ainda não conseguiu falar com o servidor: alguns dias de tolerância.
      If estado = licSemLicenca And Len(token) = 0 And Not IsNull(criadoEm) Then
         If agora < DateAdd("d", DIAS_INSTALACAO, criadoEm) Then
            estado = licLiberada
            mEmInstalacao = True
         End If
      End If
   ElseIf mDados.Cnpj <> LicCnpjEmpresa() Then
      estado = IIf(podeHoje, licCnpjDiferente, licLiberada)
   Else
      bloqueiaEm = mDados.DataBloqueio
      If Not IsNull(liberadoAte) Then
         If liberadoAte > bloqueiaEm Then bloqueiaEm = liberadoAte
      End If
      LicBloqueiaEm = bloqueiaEm
      If agora >= bloqueiaEm Then
         'Vencida, mas hoje é fim de semana/feriado: não bloqueia, só avisa.
         estado = IIf(podeHoje, licBloqueada, licAviso)
      Else
         LicDiasParaBloqueio = DateValue(bloqueiaEm) - hoje
         If LicDiasParaBloqueio >= 1 And LicDiasParaBloqueio <= DIAS_AVISO Then
            estado = licAviso
         Else
            estado = licLiberada
         End If
      End If
   End If

   LicEstadoAtual = estado
   LicAvaliar = estado
End Function

'=============================================================================
' SERVIDOR DE LICENÇAS
'=============================================================================
'Busca a licença atual no servidor e guarda. Na primeira vez, envia o pré-cadastro da empresa.
Public Function LicSincronizar(ByRef Msg As String) As Boolean
   On Error GoTo Falha
   Dim cnpj As String, chave As String, base As String, st As Long, resp As String

   Msg = ""
   If Not LicConexao(cnpj, chave, base, Msg) Then GoTo Fim

   If Not LicHttp("GET", base & "/token", "", chave, st, resp) Then
      Msg = LicSemConexao(): GoTo Fim
   End If
   If st = 401 Then
      'Chave ainda desconhecida no servidor: envia o pré-cadastro e tenta de novo.
      If Not LicEnviarPreCadastro(base, chave, Msg) Then GoTo Fim
      If Not LicHttp("GET", base & "/token", "", chave, st, resp) Then
         Msg = LicSemConexao(): GoTo Fim
      End If
   End If
   If st = 202 Or st = 401 Then
      Msg = "Instalação aguardando liberação do suporte.": GoTo Fim
   End If
   If st <> 200 Then
      Msg = "O servidor de licenças não respondeu (" & st & ").": GoTo Fim
   End If

   LicSincronizar = LicInstalarToken(resp, Msg)
   If LicSincronizar Then dbData.Execute "UPDATE od.Licenca SET UltimaVerificacao = GETDATE() WHERE Id = 1;"
Fim:
   LicUltimaMensagem = IIf(LicSincronizar, "", Msg)
   Exit Function
Falha:
   Msg = "Falha ao consultar a licença: " & Err.Description
   Resume Fim
End Function

Private Function LicEnviarPreCadastro(ByVal Base As String, ByVal Chave As String, ByRef Msg As String) As Boolean
   Dim r As ADODB.Recordset, corpo As String, st As Long, resp As String

   Set r = dbData.OpenRecordset("SELECT TOP 1 FANTASIA, RAZAO, CELULAR, EMAIL, CIDADE, ESTADO FROM empresa;")
   If r.EOF Then
      Msg = "Cadastre a empresa antes de liberar o sistema."
      Exit Function
   End If
   corpo = "{""fantasia"":" & LicJsonStr("" & r("FANTASIA")) & _
           ",""razao"":" & LicJsonStr("" & r("RAZAO")) & _
           ",""celular"":" & LicJsonStr("" & r("CELULAR")) & _
           ",""email"":" & LicJsonStr("" & r("EMAIL")) & _
           ",""cidade"":" & LicJsonStr("" & r("CIDADE")) & _
           ",""estado"":" & LicJsonStr("" & r("ESTADO")) & "}"
   r.Close
   Set r = Nothing

   If Not LicHttp("POST", Base & "/instalacao", corpo, Chave, st, resp) Then
      Msg = LicSemConexao()
   ElseIf st = 200 Or st = 202 Then
      LicEnviarPreCadastro = True
   ElseIf st = 400 Then
      Msg = LicJsonTexto(resp, "mensagem")
      If Len(Msg) = 0 Then Msg = "Dados da empresa recusados pelo servidor de licenças."
   ElseIf st = 429 Then
      Msg = "Muitas tentativas. Tente novamente mais tarde."
   Else
      Msg = "O servidor de licenças não respondeu (" & st & ")."
   End If
End Function

'"Pagar agora": pede o Pix das mensalidades em aberto. ArquivoImagem = BMP do QR Code numa pasta temporária.
Public Function LicCriarPix(ByRef TxId As String, ByRef CopiaECola As String, ByRef ArquivoImagem As String, _
                            ByRef Descricao As String, ByRef Msg As String) As Boolean
   On Error GoTo Falha
   Dim cnpj As String, chave As String, base As String, st As Long, resp As String
   Dim img As String, p As Long, bytes() As Byte, f As Integer, comp As String

   If Not LicConexao(cnpj, chave, base, Msg) Then Exit Function
   If Not LicHttp("POST", base & "/pix?formato=bmp", "", chave, st, resp) Then
      Msg = LicSemConexao() & " Use um código de liberação."
      Exit Function
   End If
   If st = 404 Then
      Msg = "Não há mensalidade em aberto. Clique em VERIFICAR NOVAMENTE."
      Exit Function
   ElseIf st = 401 Then
      Msg = "Esta instalação ainda aguarda a liberação do suporte: " & SUPORTE & "."
      Exit Function
   ElseIf st = 429 Then
      Msg = "Muitas tentativas. Aguarde alguns minutos e tente de novo."
      Exit Function
   ElseIf st <> 200 Then
      Msg = "Não foi possível gerar o Pix agora. Tente novamente em instantes."
      Exit Function
   End If

   TxId = LicJsonTexto(resp, "txId")
   CopiaECola = LicJsonTexto(resp, "copiaECola")
   'Mês/ano que está sendo pago, ex.: "Mensalidade 10/2026 - R$ 90,00" ou "Mensalidades 08 e 09/2026 - R$ 180,00".
   comp = LicJsonTexto(resp, "competencia")
   Descricao = IIf(InStr(comp, ",") > 0 Or InStr(comp, " e ") > 0, "Mensalidades ", "Mensalidade ") & comp & _
               "  -  R$ " & Format$(Val(LicJsonValor(resp, "valor")), "#,##0.00")

   ArquivoImagem = ""
   img = LicJsonTexto(resp, "imagem")
   p = InStr(img, "base64,")
   If p > 0 Then
      If LicBase64(Mid$(img, p + 7), bytes) Then
         ArquivoImagem = Environ$("TEMP") & "\licenca_pix.bmp"
         On Error Resume Next
         Kill ArquivoImagem
         On Error GoTo Falha
         f = FreeFile
         Open ArquivoImagem For Binary Access Write As #f
         Put #f, , bytes
         Close #f
      End If
   End If
   LicCriarPix = (Len(TxId) > 0)
   If Not LicCriarPix Then Msg = "Resposta inválida do servidor de licenças."
   Exit Function
Falha:
   Msg = "Falha ao gerar o Pix: " & Err.Description
End Function

'Consulta se o Pix foi pago; se sim, guarda a licença nova. True = pago e licença atualizada.
Public Function LicVerificarPix(ByVal TxId As String) As Boolean
   On Error GoTo Falha
   Dim cnpj As String, chave As String, base As String, st As Long, resp As String, msg As String

   If Not LicConexao(cnpj, chave, base, msg) Then Exit Function
   If Not LicHttp("GET", base & "/pix/" & TxId, "", chave, st, resp) Then Exit Function
   If st <> 200 Or LicJsonValor(resp, "pago") <> "true" Then Exit Function
   If LicInstalarToken(LicJsonTexto(resp, "token"), msg) Then
      dbData.Execute "UPDATE od.Licenca SET UltimaVerificacao = GETDATE() WHERE Id = 1;"
      LicUltimaMensagem = ""
      LicVerificarPix = True
   End If
   Exit Function
Falha:
   LicVerificarPix = False
End Function

Private Function LicConexao(ByRef Cnpj As String, ByRef Chave As String, ByRef Base As String, ByRef Msg As String) As Boolean
   Dim url As String

   Cnpj = LicCnpjEmpresa()
   If Len(Cnpj) <> 11 And Len(Cnpj) <> 14 Then
      Msg = "Cadastre a empresa com um CNPJ/CPF válido."
      Exit Function
   End If
   url = LicConfig("LicencaServidorUrl")
   If Len(url) = 0 Then url = URL_PADRAO
   If Right$(url, 1) = "/" Then url = Left$(url, Len(url) - 1)

   Chave = LicConfig("LicencaChaveCliente")
   If Len(Chave) = 0 Then
      Chave = LicNovaChave()
      If Len(Chave) = 0 Then
         Msg = "Não foi possível gerar a chave desta instalação."
         Exit Function
      End If
      LicGravarConfig "LicencaChaveCliente", Chave
   End If

   Base = url & "/api/licencas/" & Cnpj
   LicConexao = True
End Function

Private Function LicSemConexao() As String
   LicSemConexao = "Sem conexão com o servidor de licenças" & IIf(Len(LicErroHttp) > 0, " (" & LicErroHttp & ")", "") & "."
End Function

'Tenta sem travar a tela; se essa forma falhar (não por demora do servidor), repete do jeito simples.
Private Function LicHttp(ByVal Metodo As String, ByVal Url As String, ByVal Corpo As String, ByVal Chave As String, _
                         ByRef Status As Long, ByRef Resposta As String) As Boolean
   Dim porTempo As Boolean

   LicErroHttp = ""
   LicHttp = LicHttpEnviar(Metodo, Url, Corpo, Chave, True, Status, Resposta, porTempo)
   If Not LicHttp And Not porTempo Then
      LicHttp = LicHttpEnviar(Metodo, Url, Corpo, Chave, False, Status, Resposta, porTempo)
   End If
End Function

Private Function LicHttpEnviar(ByVal Metodo As String, ByVal Url As String, ByVal Corpo As String, ByVal Chave As String, _
                               ByVal Assincrono As Boolean, ByRef Status As Long, ByRef Resposta As String, _
                               ByRef PorTempo As Boolean) As Boolean
   On Error GoTo Falha
   Dim x As Object, inicio As Single, decorrido As Single, etapa As String

   PorTempo = False
   etapa = "criar"
   Set x = LicCriarHttp()
   If x Is Nothing Then LicErroHttp = "MSXML2.ServerXMLHTTP indisponível": Exit Function
   x.setTimeouts 3000, 4000, 8000, 15000
   etapa = "abrir"
   'Assíncrono: enquanto espera, a tela continua respondendo (não parece travada).
   x.Open Metodo, Url, Assincrono
   x.setRequestHeader "X-Chave-Cliente", Chave
   etapa = "enviar"
   If Len(Corpo) > 0 Then
      x.setRequestHeader "Content-Type", "application/json"
      x.send Corpo
   Else
      x.send ""
   End If
   If Assincrono Then
      etapa = "aguardar"
      inicio = Timer
      Do Until x.waitForResponse(1)
         DoEvents
         decorrido = Timer - inicio
         If decorrido < 0 Then decorrido = decorrido + 86400   'passou da meia-noite
         LicSegundosEspera = Int(decorrido)
         If decorrido > 30 Then
            x.abort
            PorTempo = True
            LicErroHttp = "o servidor não respondeu em 30 segundos"
            LicSegundosEspera = 0
            Exit Function
         End If
      Loop
   End If
   etapa = "ler"
   LicSegundosEspera = 0
   Status = x.Status
   Resposta = x.responseText
   LicHttpEnviar = True
   Exit Function
Falha:
   LicErroHttp = IIf(Assincrono, "assíncrono", "simples") & "/" & etapa & ": " & Err.Number & " " & Err.Description
   Debug.Print "LicHttp " & LicErroHttp
   LicSegundosEspera = 0
   LicHttpEnviar = False
End Function

Private Function LicCriarHttp() As Object
   On Error Resume Next
   Set LicCriarHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
   If LicCriarHttp Is Nothing Then Set LicCriarHttp = CreateObject("MSXML2.ServerXMLHTTP")
End Function

'Confere a licença e grava em od.Licenca (só se a assinatura confere e o CNPJ é desta empresa).
Private Function LicInstalarToken(ByVal Token As String, ByRef Msg As String) As Boolean
   Dim d As LicDados

   Token = Trim$(Token)
   If Not LicLerToken(Token, d) Then
      Msg = "Licença inválida recebida do servidor."
      Exit Function
   End If
   If d.Cnpj <> LicCnpjEmpresa() Then
      Msg = "A licença recebida é de outro CNPJ."
      Exit Function
   End If
   dbData.Execute "UPDATE od.Licenca SET Token = '" & Replace(Token, "'", "") & "', AtualizadoEm = GETDATE() WHERE Id = 1;"
   Msg = "Licença atualizada. Próximo bloqueio: " & Format$(d.DataBloqueio, "dd/mm/yyyy") & "."
   LicInstalarToken = True
End Function

'=============================================================================
' CÓDIGO DE LIBERAÇÃO OFFLINE (gerado no painel de licenças)
'=============================================================================
Public Function LicAplicarCodigo(ByVal Codigo As String, ByVal CodUsuario As Long, ByRef Msg As String) As Boolean
   Dim r As ADODB.Recordset, d As LicDados, cnpj As String, normalizado As String
   Dim tipo As Long, dataCodigo As Date, agora As Date, hoje As Date, novoLimite As Date
   Dim temporarioUsadoPara As Variant

   Set r = dbData.OpenRecordset("SELECT Token, UltimaDataVista, TemporarioUsadoPara FROM od.Licenca WHERE Id = 1;")
   If r.EOF Then Msg = "Nenhuma licença instalada.": Exit Function
   agora = Now
   If Not IsNull(r("UltimaDataVista")) Then
      If r("UltimaDataVista") > agora Then agora = r("UltimaDataVista")
   End If
   temporarioUsadoPara = r("TemporarioUsadoPara")
   If Not LicLerToken("" & r("Token"), d) Then
      Msg = "Nenhuma licença instalada. Conecte à internet e clique em VERIFICAR NOVAMENTE."
      Exit Function
   End If
   r.Close
   hoje = DateValue(agora)

   If mErrosCodigo >= MAX_ERROS_CODIGO And Now < mErrosAte Then
      Msg = "Muitos códigos inválidos. Aguarde alguns minutos."
      Exit Function
   End If

   cnpj = LicCnpjEmpresa()
   normalizado = LicNormalizarCodigo(Codigo)
   If Not LicLerCodigo(normalizado, cnpj, d.ChaveOffline, tipo, dataCodigo) Then
      If Now >= mErrosAte Then mErrosCodigo = 0
      mErrosCodigo = mErrosCodigo + 1
      mErrosAte = DateAdd("n", 5, Now)
      Msg = "Código inválido."
      Exit Function
   End If
   mErrosCodigo = 0

   Set r = dbData.OpenRecordset("SELECT Codigo FROM od.LicencaCodigosUsados WHERE Codigo = '" & normalizado & "';")
   If Not r.EOF Then Msg = "Este código já foi usado.": Exit Function
   r.Close
   Set r = Nothing

   If tipo = 2 Then
      'Temporário: 24h a partir de agora, um por mensalidade.
      If dataCodigo > hoje Or dataCodigo < hoje - 2 Then
         Msg = "Código temporário vencido. Peça um novo.": Exit Function
      End If
      If Not IsNull(temporarioUsadoPara) Then
         If DateValue(temporarioUsadoPara) = d.DataBloqueio Then
            Msg = "O código temporário desta mensalidade já foi usado.": Exit Function
         End If
      End If
      novoLimite = DateAdd("h", 24, agora)
      dbData.Execute "UPDATE od.Licenca SET TemporarioUsadoPara = '" & Format$(d.DataBloqueio, "yyyymmdd") & "' WHERE Id = 1;"
      Msg = "Liberado até " & Format$(novoLimite, "dd/mm hh:nn") & "."
   Else
      'Pago: liberado até a data do código.
      novoLimite = dataCodigo
      If novoLimite <= agora Then Msg = "Código de liberação vencido.": Exit Function
      Msg = "Liberado. Próximo bloqueio: " & Format$(dataCodigo, "dd/mm/yyyy") & "."
   End If

   dbData.Execute "UPDATE od.Licenca SET LiberadoAte = '" & LicSqlData(novoLimite) & "' " & _
                  "WHERE Id = 1 AND (LiberadoAte IS NULL OR LiberadoAte < '" & LicSqlData(novoLimite) & "');"
   dbData.Execute "INSERT INTO od.LicencaCodigosUsados (Codigo, Tipo, UsadoEm, CodUsuario) VALUES ('" & _
                  normalizado & "', " & tipo & ", GETDATE(), " & IIf(CodUsuario > 0, CStr(CodUsuario), "NULL") & ");"
   LicAplicarCodigo = True
End Function

Private Function LicNormalizarCodigo(ByVal Codigo As String) As String
   Dim i As Long, ch As String, s As String
   Codigo = UCase$(Codigo)
   For i = 1 To Len(Codigo)
      ch = Mid$(Codigo, i, 1)
      If ch Like "[A-Z0-9]" Then
         If ch = "O" Then ch = "0"
         If ch = "I" Or ch = "L" Then ch = "1"
         s = s & ch
      End If
   Next
   LicNormalizarCodigo = s
End Function

'8 caracteres = 40 bits: tipo (2) + dias desde 01/01/2024 (14) + HMAC-SHA256 truncado (24).
'Contas em Double (exato até 2^53) porque o VB6 não tem inteiro de 64 bits.
Private Function LicLerCodigo(ByVal Normalizado As String, ByVal Cnpj As String, ByVal ChaveOffline As String, _
                              ByRef Tipo As Long, ByRef DataCodigo As Date) As Boolean
   Dim v As Double, resto As Double, mac As Double, i As Long, idx As Long, dias As Long
   Dim chave() As Byte, mensagem() As Byte, h() As Byte

   If Len(Normalizado) <> 8 Then Exit Function
   For i = 1 To 8
      idx = InStr(ALFABETO_CODIGO, Mid$(Normalizado, i, 1)) - 1
      If idx < 0 Then Exit Function
      v = v * 32# + idx
   Next
   Tipo = Int(v / 274877906944#)              '2^38
   resto = v - Tipo * 274877906944#
   dias = Int(resto / 16777216#)              '2^24
   mac = resto - dias * 16777216#
   If Tipo <> 1 And Tipo <> 2 Then Exit Function

   If Not LicBase64(ChaveOffline, chave) Then Exit Function
   mensagem = StrConv("OD-LIB|" & Cnpj & "|" & Tipo & "|" & dias, vbFromUnicode)
   If Not LicCalcularHash(mensagem, chave, True, h) Then Exit Function
   If mac <> CDbl(h(0)) * 65536# + CDbl(h(1)) * 256# + CDbl(h(2)) Then Exit Function

   DataCodigo = DateSerial(2024, 1, 1) + dias
   LicLerCodigo = True
End Function

'=============================================================================
' LICENÇA ASSINADA: base64url(json) + "." + base64url(assinatura ECDSA P-256/SHA-256)
'=============================================================================
Private Function LicLerToken(ByVal Token As String, ByRef D As LicDados) As Boolean
   On Error GoTo Falha
   Dim p As Long, corpo() As Byte, assinatura() As Byte, json As String

   Token = Trim$(Token)
   p = InStr(Token, ".")
   If p < 2 Then Exit Function
   If InStr(p + 1, Token, ".") > 0 Then Exit Function
   If Not LicBase64Url(Left$(Token, p - 1), corpo) Then Exit Function
   If Not LicBase64Url(Mid$(Token, p + 1), assinatura) Then Exit Function
   If Not LicVerificarAssinatura(corpo, assinatura) Then Exit Function

   json = LicUtf8(corpo)
   D.Cnpj = LicSoDigitos(LicJsonTexto(json, "cnpj"))
   D.Cliente = LicJsonTexto(json, "cliente")
   D.Plano = LicJsonTexto(json, "plano")
   D.Vencimento = LicJsonData(json, "vencimento")
   D.DataBloqueio = LicJsonData(json, "dataBloqueio")
   D.ChaveOffline = LicJsonTexto(json, "chaveOffline")
   LicLerAbertas json, D
   LicLerToken = (Len(D.Cnpj) > 0 And Len(D.ChaveOffline) > 0 And D.DataBloqueio > 0)
   Exit Function
Falha:
   LicLerToken = False
End Function

'"emAberto":[{"competencia":"2026-09-01","valor":10.0}, ...]
Private Sub LicLerAbertas(ByVal Json As String, ByRef D As LicDados)
   Dim ini As Long, fim As Long, lista As String, p As Long, item As String, prox As Long, comp As Date

   D.Abertas = ""
   D.ValorAberto = 0
   ini = InStr(Json, """emAberto"":[")
   If ini = 0 Then Exit Sub
   ini = ini + Len("""emAberto"":[")
   fim = InStr(ini, Json, "]")
   If fim = 0 Then Exit Sub
   lista = Mid$(Json, ini, fim - ini)

   p = InStr(lista, "{")
   Do While p > 0
      prox = InStr(p + 1, lista, "}")
      If prox = 0 Then Exit Do
      item = Mid$(lista, p, prox - p + 1)
      comp = LicJsonData(item, "competencia")
      If comp > 0 Then
         D.Abertas = D.Abertas & IIf(Len(D.Abertas) > 0, ", ", "") & Format$(comp, "mm/yyyy")
         D.ValorAberto = D.ValorAberto + CCur(Val(LicJsonValor(item, "valor")))
      End If
      p = InStr(prox + 1, lista, "{")
   Loop
End Sub

Private Function LicVerificarAssinatura(Corpo() As Byte, Assinatura() As Byte) As Boolean
   On Error GoTo Falha
   Dim hAlg As Long, hKey As Long, alg As String, tipoBlob As String
   Dim blob() As Byte, hash() As Byte, nada() As Byte

   If LicTamanho(Assinatura) <> 64 Then Exit Function
   If Not LicCalcularHash(Corpo, nada, False, hash) Then Exit Function
   blob = LicBlobChavePublica()

   alg = "ECDSA_P256"
   tipoBlob = "ECCPUBLICBLOB"
   If BCryptOpenAlgorithmProvider(hAlg, StrPtr(alg), 0, 0) <> 0 Then Exit Function
   If BCryptImportKeyPair(hAlg, 0, StrPtr(tipoBlob), hKey, VarPtr(blob(0)), UBound(blob) + 1, 0) = 0 Then
      LicVerificarAssinatura = (BCryptVerifySignature(hKey, 0, VarPtr(hash(0)), 32, VarPtr(Assinatura(0)), 64, 0) = 0)
      BCryptDestroyKey hKey
   End If
   BCryptCloseAlgorithmProvider hAlg, 0
   Exit Function
Falha:
   LicVerificarAssinatura = False
End Function

'BCRYPT_ECCKEY_BLOB: "ECS1" (chave pública P-256) + tamanho 32 + X + Y.
Private Function LicBlobChavePublica() As Byte()
   Dim b(71) As Byte, i As Long
   b(0) = &H45: b(1) = &H43: b(2) = &H53: b(3) = &H31
   b(4) = 32
   For i = 0 To 63
      b(8 + i) = CByte("&H" & Mid$(CHAVE_PUBLICA_XY, i * 2 + 1, 2))
   Next
   LicBlobChavePublica = b
End Function

'SHA-256 (ComChave = False) ou HMAC-SHA256 (ComChave = True).
Private Function LicCalcularHash(Dados() As Byte, Chave() As Byte, ByVal ComChave As Boolean, Saida() As Byte) As Boolean
   Dim hAlg As Long, hHash As Long, alg As String, n As Long, pChave As Long, nChave As Long

   alg = "SHA256"
   If BCryptOpenAlgorithmProvider(hAlg, StrPtr(alg), 0, IIf(ComChave, 8&, 0&)) <> 0 Then Exit Function
   If ComChave Then
      nChave = LicTamanho(Chave)
      If nChave > 0 Then pChave = VarPtr(Chave(0))
   End If
   If BCryptCreateHash(hAlg, hHash, 0, 0, pChave, nChave, 0) = 0 Then
      n = LicTamanho(Dados)
      If n > 0 Then BCryptHashData hHash, VarPtr(Dados(0)), n, 0
      ReDim Saida(31)
      LicCalcularHash = (BCryptFinishHash(hHash, VarPtr(Saida(0)), 32, 0) = 0)
      BCryptDestroyHash hHash
   End If
   BCryptCloseAlgorithmProvider hAlg, 0
End Function

'Chave desta instalação: 16 bytes aleatórios em hexadecimal (32 caracteres).
Private Function LicNovaChave() As String
   On Error GoTo Falha
   Dim b(15) As Byte, i As Long, s As String
   If BCryptGenRandom(0, VarPtr(b(0)), 16, 2) <> 0 Then Exit Function   '2 = gerador do sistema
   For i = 0 To 15
      s = s & LCase$(Right$("0" & Hex$(b(i)), 2))
   Next
   LicNovaChave = s
   Exit Function
Falha:
   LicNovaChave = ""
End Function

'=============================================================================
' CALENDÁRIO: nunca bloqueia em sábado, domingo ou feriado nacional
'=============================================================================
Private Function LicPodeBloquear(ByVal Dia As Date) As Boolean
   Select Case Weekday(Dia)
      Case vbSaturday, vbSunday
         LicPodeBloquear = False
      Case Else
         LicPodeBloquear = Not LicEhFeriado(Dia)
   End Select
End Function

Private Function LicEhFeriado(ByVal Dia As Date) As Boolean
   Dim md As String, pascoa As Date
   md = Format$(Dia, "mm-dd")
   Select Case md
      Case "01-01", "04-21", "05-01", "09-07", "10-12", "11-02", "11-15", "12-25"
         LicEhFeriado = True: Exit Function
      Case "11-20"   'Consciência Negra (Lei 14.759/2023)
         If Year(Dia) >= 2024 Then LicEhFeriado = True: Exit Function
   End Select
   pascoa = LicPascoa(Year(Dia))
   LicEhFeriado = (Dia = pascoa - 48 Or Dia = pascoa - 47 Or Dia = pascoa - 2 Or Dia = pascoa + 60)
End Function

'Algoritmo de Meeus/Jones/Butcher (calendário gregoriano).
Private Function LicPascoa(ByVal Ano As Long) As Date
   Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long
   Dim h As Long, i As Long, k As Long, l As Long, m As Long
   a = Ano Mod 19: b = Ano \ 100: c = Ano Mod 100: d = b \ 4: e = b Mod 4
   f = (b + 8) \ 25: g = (b - f + 1) \ 3: h = (19 * a + b - d - g + 15) Mod 30
   i = c \ 4: k = c Mod 4: l = (32 + 2 * e + 2 * i - h - k) Mod 7
   m = (a + 11 * h + 22 * l) \ 451
   LicPascoa = DateSerial(Ano, (h + l - 7 * m + 114) \ 31, ((h + l - 7 * m + 114) Mod 31) + 1)
End Function

'=============================================================================
' BANCO
'=============================================================================
'Cria as tabelas da licença se ainda não existem (mesmas do Online Food; script 117 faz o mesmo).
Private Function LicGarantirTabelas() As Boolean
   On Error GoTo Falha
   Dim r As ADODB.Recordset, existe As Boolean

   If mTabelasOk Then LicGarantirTabelas = True: Exit Function
   Set r = dbData.OpenRecordset("SELECT CASE WHEN OBJECT_ID('od.Licenca') IS NOT NULL AND OBJECT_ID('od.LicencaCodigosUsados') IS NOT NULL " & _
                                "AND OBJECT_ID('od.Configuracoes') IS NOT NULL THEN 1 ELSE 0 END AS Existe;")
   existe = (r("Existe") = 1)
   r.Close
   Set r = Nothing

   If Not existe Then
      dbData.Execute "IF SCHEMA_ID('od') IS NULL EXEC('CREATE SCHEMA od');"
      dbData.Execute "IF OBJECT_ID('od.Configuracoes') IS NULL CREATE TABLE od.Configuracoes (" & _
                     "Chave varchar(50) NOT NULL CONSTRAINT PK_Configuracoes PRIMARY KEY, Valor nvarchar(200) NULL);"
      dbData.Execute "IF OBJECT_ID('od.Licenca') IS NULL CREATE TABLE od.Licenca (" & _
                     "Id tinyint NOT NULL CONSTRAINT PK_Licenca PRIMARY KEY CONSTRAINT CK_Licenca_Unica CHECK (Id = 1), " & _
                     "Token varchar(max) NULL, " & _
                     "AtualizadoEm datetime NOT NULL CONSTRAINT DF_Licenca_Atualizado DEFAULT (GETDATE()), " & _
                     "UltimaVerificacao datetime NULL, UltimaDataVista datetime NULL, LiberadoAte datetime NULL, " & _
                     "TemporarioUsadoPara date NULL);"
      dbData.Execute "IF OBJECT_ID('od.LicencaCodigosUsados') IS NULL CREATE TABLE od.LicencaCodigosUsados (" & _
                     "Codigo char(8) NOT NULL CONSTRAINT PK_LicencaCodigosUsados PRIMARY KEY, Tipo tinyint NOT NULL, " & _
                     "UsadoEm datetime NOT NULL CONSTRAINT DF_LicCod_UsadoEm DEFAULT (GETDATE()), CodUsuario int NULL);"
   End If
   dbData.Execute "IF NOT EXISTS (SELECT 1 FROM od.Licenca WHERE Id = 1) INSERT INTO od.Licenca (Id) VALUES (1);"

   mTabelasOk = True
   LicGarantirTabelas = True
   Exit Function
Falha:
   Debug.Print "LicGarantirTabelas: " & Err.Number & " " & Err.Description
   LicGarantirTabelas = False
End Function

'Usado ao zerar a base para um cliente novo: a licença e a chave desta instalação não vão junto.
Public Sub LicZerar()
   dbData.Execute "IF OBJECT_ID('od.Licenca') IS NOT NULL UPDATE od.Licenca SET Token = NULL, AtualizadoEm = GETDATE(), " & _
                  "UltimaVerificacao = NULL, UltimaDataVista = NULL, LiberadoAte = NULL, TemporarioUsadoPara = NULL;"
   dbData.Execute "IF OBJECT_ID('od.LicencaCodigosUsados') IS NOT NULL DELETE FROM od.LicencaCodigosUsados;"
   dbData.Execute "IF OBJECT_ID('od.Configuracoes') IS NOT NULL DELETE FROM od.Configuracoes WHERE Chave = 'LicencaChaveCliente';"
End Sub

Private Function LicCnpjEmpresa() As String
   Dim r As ADODB.Recordset
   Set r = dbData.OpenRecordset("SELECT TOP 1 CNPJ FROM empresa;")
   If Not r.EOF Then LicCnpjEmpresa = LicSoDigitos("" & r("CNPJ"))
   r.Close
   Set r = Nothing
End Function

Private Function LicConfig(ByVal Chave As String) As String
   Dim r As ADODB.Recordset
   Set r = dbData.OpenRecordset("SELECT Valor FROM od.Configuracoes WHERE Chave = '" & Chave & "';")
   If Not r.EOF Then LicConfig = Trim$("" & r("Valor"))
   r.Close
   Set r = Nothing
End Function

Private Sub LicGravarConfig(ByVal Chave As String, ByVal Valor As String)
   Valor = Replace(Valor, "'", "''")
   dbData.Execute "IF EXISTS (SELECT 1 FROM od.Configuracoes WHERE Chave = '" & Chave & "') " & _
                  "UPDATE od.Configuracoes SET Valor = N'" & Valor & "' WHERE Chave = '" & Chave & "' " & _
                  "ELSE INSERT INTO od.Configuracoes (Chave, Valor) VALUES ('" & Chave & "', N'" & Valor & "');"
End Sub

'Data/hora no formato que o SQL Server entende em qualquer idioma.
Private Function LicSqlData(ByVal D As Date) As String
   LicSqlData = Format$(D, "yyyymmdd hh:nn:ss")
End Function

'=============================================================================
' AUXILIARES: JSON simples, base64, UTF-8
'=============================================================================
'Valor de um campo de texto do JSON ("nome":"valor"), com os escapes resolvidos.
Private Function LicJsonTexto(ByVal Json As String, ByVal Nome As String) As String
   Dim p As Long, i As Long, ch As String, s As String

   p = InStr(Json, """" & Nome & """:")
   If p = 0 Then Exit Function
   p = p + Len(Nome) + 3
   Do While Mid$(Json, p, 1) = " "
      p = p + 1
   Loop
   If Mid$(Json, p, 1) <> """" Then Exit Function

   i = p + 1
   Do While i <= Len(Json)
      ch = Mid$(Json, i, 1)
      If ch = """" Then Exit Do
      If ch = "\" Then
         i = i + 1
         ch = Mid$(Json, i, 1)
         Select Case ch
            Case "n": ch = vbLf
            Case "r": ch = vbCr
            Case "t": ch = vbTab
            Case "b": ch = Chr$(8)
            Case "f": ch = Chr$(12)
            Case "u"
               ch = ChrW$(CLng("&H" & Mid$(Json, i + 1, 4)))
               i = i + 4
         End Select
      End If
      s = s & ch
      i = i + 1
   Loop
   LicJsonTexto = s
End Function

'Valor de um campo não texto do JSON (número, true/false/null), como texto.
Private Function LicJsonValor(ByVal Json As String, ByVal Nome As String) As String
   Dim p As Long, i As Long, ch As String

   p = InStr(Json, """" & Nome & """:")
   If p = 0 Then Exit Function
   p = p + Len(Nome) + 3
   For i = p To Len(Json)
      ch = Mid$(Json, i, 1)
      If ch = "," Or ch = "}" Or ch = "]" Then Exit For
   Next
   LicJsonValor = Trim$(Mid$(Json, p, i - p))
End Function

'"aaaa-mm-dd" (ou "aaaa-mm-ddThh:mm:ss") -> Date. 0 se vazio/inválido.
Private Function LicJsonData(ByVal Json As String, ByVal Nome As String) As Date
   Dim s As String
   s = LicJsonTexto(Json, Nome)
   If Len(s) < 10 Then Exit Function
   If Not IsNumeric(Left$(s, 4)) Then Exit Function
   LicJsonData = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 6, 2)), Val(Mid$(s, 9, 2)))
End Function

'Texto -> string JSON. Acentos vão como \uXXXX (independe da codificação da conexão).
Private Function LicJsonStr(ByVal Texto As String) As String
   Dim i As Long, ch As String, c As Long, s As String
   Texto = Trim$(Texto)
   For i = 1 To Len(Texto)
      ch = Mid$(Texto, i, 1)
      c = AscW(ch) And &HFFFF&
      Select Case True
         Case ch = """": s = s & "\"""
         Case ch = "\": s = s & "\\"
         Case c < 32 Or c > 126: s = s & "\u" & Right$("000" & Hex$(c), 4)
         Case Else: s = s & ch
      End Select
   Next
   LicJsonStr = """" & s & """"
End Function

Private Function LicBase64Url(ByVal Texto As String, ByRef Saida() As Byte) As Boolean
   Texto = Replace(Replace(Texto, "-", "+"), "_", "/")
   Select Case Len(Texto) Mod 4
      Case 2: Texto = Texto & "=="
      Case 3: Texto = Texto & "="
      Case 1: Exit Function
   End Select
   LicBase64Url = LicBase64(Texto, Saida)
End Function

Private Function LicBase64(ByVal Texto As String, ByRef Saida() As Byte) As Boolean
   On Error GoTo Falha
   Dim doc As Object, el As Object
   Set doc = CreateObject("MSXML2.DOMDocument")
   Set el = doc.createElement("b64")
   el.DataType = "bin.base64"
   el.Text = Texto
   Saida = el.nodeTypedValue
   LicBase64 = (LicTamanho(Saida) > 0)
   Exit Function
Falha:
   LicBase64 = False
End Function

Private Function LicUtf8(B() As Byte) As String
   Dim n As Long, s As String
   n = LicTamanho(B)
   If n = 0 Then Exit Function
   s = String$(n, vbNullChar)
   n = LicMultiByteToWideChar(65001, 0, VarPtr(B(0)), LicTamanho(B), StrPtr(s), n)
   LicUtf8 = Left$(s, n)
End Function

Private Function LicTamanho(B() As Byte) As Long
   On Error GoTo Vazio
   LicTamanho = UBound(B) - LBound(B) + 1
   Exit Function
Vazio:
   LicTamanho = 0
End Function

Private Function LicSoDigitos(ByVal Texto As String) As String
   Dim i As Long, ch As String, s As String
   For i = 1 To Len(Texto)
      ch = Mid$(Texto, i, 1)
      If ch Like "#" Then s = s & ch
   Next
   LicSoDigitos = s
End Function
