Attribute VB_Name = "NFe_DLL"
'usado no Projeto OnlineCommerce
'* Sistema...: Módulo NFe/NFCe
'* Empresa...: EkklesiaSoft Tecnologia em Sistemas
'* Módulo....: NFe_DLL
'* Função....: Módulo de funções da Nota Fiscal Eletrônica e Nota Fiscal Consumidor Eletrônica
'* CopyRight.: (C)2015 EkklesiaSoft Tecnologia em Sistemas
'* Criação...: EkklesiaSoft Tecnologia em Sistemas
'* Data......: 16/01/2014 07:49:46
'* * * * * * *

Option Explicit                                   'requer variáveis explicitamente declaradas

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

Public NFeXML As String, NFeValidate As String, NFeChaveAcesso As String, NFeChaveAcessoAdicional As String, NFecNF As String, NFeMotivo As String, NFeResposta As String, NFeNumeroRecibo As String, NFeNumeroProtocolo As String, NFeDataHora As String, NFeDataHoraEnvio As String
Public vsCertificado As String, msgResultado As String, nfeRetorno As String, iRetorno As Long, XMLOK As Boolean, cStat As Long, cStat2 As Long
Public cabMsg As String, DadosMsg As String, msgRetWS As String, Proxy As String, UsuarioProxy As String, SenhaProxy As String
Public xCaminhoXML As String, xCaminhoXMLAuxiliar As String, xCaminhoTXT As String, xCaminhoPDF As String, dirXML As String, vsSQL As String, retXML As String
Public vsNumeroNota As Variant, mensagemAlerta As String, mensagemErro As String   'tirei o UTC As String
Public retornoTipo() As String, retornoNSU() As String, retornodhEmi() As String, retornocSitConf() As String, retornochNFe() As String
Public retornoCNPJ() As String, retornocSitNFe() As String, retornodhRecbto() As String, retornodigVal() As String, retornoIE() As String
Public retornotpNF() As String, retornovNF() As String, retornoxNome() As String, retindCont As String, retultNSU As String, retCons As String
Dim nroRecibo As String, nroProtocolo As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function FormExists(ByVal pstrFormName As String) As Form
Dim frmForm As Form
   For Each frmForm In Forms
      If frmForm.Name = pstrFormName Then
         Set FormExists = frmForm
         Set frmForm = Nothing
         Exit For
      End If
   Next frmForm
End Function
Public Function ConfiguraDLLNFeNFCe(Modelo As Integer, TipoEmissao As String, ByRef objNFeNFCe As snfe.Util) As Boolean
Dim ComandoSQL As String
   
   On Error GoTo deuErro
   
   Set objNFeNFCe = New snfe.Util
   
   ComandoSQL = "SELECT CNPJ, Razao, Cidade, Estado, CodigoIBGE, CRT, AmbienteNF, DiretorioXML, CertificadoDigital, " & _
                "NFCeIDToken, NFCeCSC, LicencaDLL, Email, caminho, NFCeOffline " & _
                "FROM Empresa"
   'MsgBox ComandoSQL
   If Modelo = 65 Then
      If SQLExecutaRetorno(ComandoSQL, "NFCeOffline", False) Then
         TipoEmissao = "9 - Contingência off-line da NFC-e"
      ElseIf Left(TipoEmissao, 1) = "9" Then
         TipoEmissao = "1 - Normal"
      End If
   End If
   
   dirXML = SQLExecutaRetorno(ComandoSQL, "DiretorioXML", App.path)
   dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
   iRetorno = objNFeNFCe.ConfigurarDLL("", _
                                       SQLExecutaRetorno(ComandoSQL, "CertificadoDigital", ""), _
                                       "", _
                                       1, _
                                       dirXML & "nfe\", _
                                       dirXML & "nfe\schemas", _
                                       SQLExecutaRetorno(ComandoSQL, "Estado", "PI"), _
                                       Modelo, _
                                       "1", _
                                       "30000", _
                                       Left(SQLExecutaRetorno(ComandoSQL, "AmbienteNF", 2), 1), _
                                       Left(TipoEmissao, 1), _
                                       4, _
                                       SQLExecutaRetorno(ComandoSQL, "CNPJ", ""), _
                                       LPad(SQLExecutaRetorno(ComandoSQL, "NFCeIDToken", "1"), 6, "0"), _
                                       SQLExecutaRetorno(ComandoSQL, "NFCeCSC", ""), _
                                       CNPJSoftHouse, _
                                       "OnLine Info", False, _
                                       SQLExecutaRetorno(ComandoSQL, "LicencaDLL", ""))
                                       
                                          
   iRetorno = objNFeNFCe.ConfigurarEmail("mail.ekklesiasoft.com.br", 587, 100000, True, "dev@ekklesiasoft.com.br", "Ekk29639780", True, SQLExecutaRetorno(ComandoSQL, "Email", "financeiroonlineinfo@gmail.com"), SQLExecutaRetorno(ComandoSQL, "Razao", "OnLine Info"))
   
   iRetorno = objNFeNFCe.ConfigurarDANFe(SQLExecutaRetorno(ComandoSQL, "caminho", ""), True, False, True, False)
   
   objNFeNFCe.certificadoAvisaVencimento = False
   objNFeNFCe.certificadoDiasAviso = 10
   
   'coloco false para não exibir msg
   objNFeNFCe.exibirAvisos = True
   
   iRetorno = objNFeNFCe.CarregarConfiguracoes
   
   ConfiguraDLLNFeNFCe = True
   
   Exit Function
                                       
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
   Err.Clear
End Function

Public Function TransmitirNFe(ByVal NumeroNota As Variant, ByVal SerieNF As Variant, Optional PodeEnviar As Boolean = False) As Boolean  'Função que monta o arquivo XML e faz o envio para a Receita
Dim txtNumerado As String, Retorno As String, vsNFe As String, dhEmi As String, dhProtocolo As String
Dim Parametros As New ADODB.Recordset, Destinatario As New ADODB.Recordset, Produtos As New ADODB.Recordset
Dim NFe As New ADODB.Recordset, NFeItens As New ADODB.Recordset, NFeParcelas As New ADODB.Recordset, NFeOBS As New ADODB.Recordset, NFeAutorizados As New ADODB.Recordset, NFeReferenciadas As New ADODB.Recordset
Dim NFeMedicamentos As New ADODB.Recordset, NFeArmamento As New ADODB.Recordset, NFeCombustivel As New ADODB.Recordset, NFeVeiculos As New ADODB.Recordset
Dim n As Integer, i As Long, destIE As String
Dim vBCCBSIBS As Double, pIBSUF As Double, vIBSUF As Double, pIBSMun As Double, vIBSMun As Double, vIBS As Double, pCBS As Double, vCBS As Double
Dim TotvBCCBSIBS As Double, TotvIBSUF As Double, TotvIBSMun As Double, TotvIBS As Double, TotvCBS As Double
Dim vsXML As String, XMLAuxiliar As String, XMLAuxiliarParcelas As String
Dim msgErro As String, qterro As Long, Prod_DetEspecifico As String
Dim vlTrib As Double, vCredICMSSN As Double
Dim CNPJContavil As String
Dim xCNPJ As String, xCPF As String
Dim yCNPJ As String, yCPF As String
Screen.MousePointer = vbHourglass

''On Error GoTo DeuErro

 'instancia o componente
 Dim sistNFe As snfe.Util
 Set sistNFe = New snfe.Util
 
 vsSQL = "SELECT * FROM Empresa"
 RsOpen Parametros, vsSQL

 vsSQL = "SELECT * " & _
         "FROM NotaFiscal " & _
         "WHERE CodigoNota = " & NumeroNota
 RsOpen NFe, vsSQL

 vlTrib = 0
 NFeMotivo = ""

 If NFe.RecordCount > 0 Then
    NFe.MoveFirst
    iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
    sistNFe.exibirAvisos = False
    If Not Vazio(NFe!ChavedeAcesso) Then
       iRetorno = sistNFe.ConsultarProtocolo(NFe!ChavedeAcesso)
       NFeChaveAcesso = NFe!ChavedeAcesso
       cStat = sistNFe.retConsulta.cStat
       NFeMotivo = sistNFe.retConsulta.xMotivo
       NFeDataHora = ""
       NFeNumeroProtocolo = ""
       If cStat = 100 Or cStat = 101 Or cStat = 110 Then
          Sleep 10000
          NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
          NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
          GoTo buscaNFe
       Else
          If Not Vazio(NFe!NumeroRecibo) Then
             iRetorno = sistNFe.ConsultarReciboDeEnvio(NFe!NumeroRecibo)
             cStat = sistNFe.retConsRec.cStat
             NFeMotivo = sistNFe.retConsRec.xMotivo
             If cStat = 104 Then
                cStat2 = sistNFe.retConsRec.protNFe.infProt.cStat
                'NFeMotivo = sistNFe.retConsRec.xMotivo
                If cStat2 = 100 Or cStat2 = 101 Or cStat = 110 Or cStat = 150 Then GoTo buscaNFe
             End If
          End If
       End If
    End If
    
    If NFe!TipoCliente = "FORNECEDOR" Then
        vsSQL = "SELECT * FROM fornecedor WHERE CODIGO = " & NFe!CodigoCorrentista  '##MUDEI AQUI##
    Else
        vsSQL = "SELECT * FROM cliente WHERE CODIGO = " & NFe!CodigoCorrentista  '##MUDEI AQUI##
    End If
    
    RsOpen Destinatario, vsSQL
    
    If Destinatario.RecordCount = 0 Then
       NFeMotivo = "CLIENTE/DESTINATÁRIO NÃO ENCONTRADO!"
       GoTo Caifora
    End If
    
    'EMITIENTE COM CPF
    If Len(Parametros!CNPJ) = 18 Then
      yCNPJ = Trim(Parametros!CNPJ)                                   ' CNPJ do destinatario sem máscara de formatação
      yCPF = ""
    Else
      yCPF = Trim(Parametros!CNPJ)                                   ' CPF do destinatario, uso exclusivo do Fisco
      yCNPJ = ""
    End If
    
    iRetorno = sistNFe.IncluirNF(mensagemAlerta, mensagemErro)
    
    '===================grupo de identificação do emitente (grupo B do Manual de integração - páginas 90)=======================
    'iRetorno = sistNFe.GerarEmitente(RemoveAcento(Parametros!Razao), RemoveAcento(Parametros!Fantasia), Parametros!CNPJ, "", Parametros!IE, "", "", "", Left(Parametros!CRT, 1), RemoveAcento(Parametros!Endereco), Parametros!Numero, "", Parametros!bairro, Parametros!CodigoIBGE, RemoveAcento(Parametros!Cidade), Parametros!Estado, Parametros!CEP, 1058, "BRASIL", Parametros!Telefone, mensagemAlerta, mensagemErro)
    iRetorno = sistNFe.GerarEmitente(RemoveAcento(Parametros!Razao), RemoveAcento(Parametros!Fantasia), yCNPJ, yCPF, Parametros!IE, "", "", "", Left(Parametros!CRT, 1), RemoveAcento(Parametros!Endereco), Parametros!Numero, "", Parametros!bairro, Parametros!CodigoIBGE, RemoveAcento(Parametros!Cidade), Parametros!Estado, Parametros!CEP, 1058, "BRASIL", Parametros!telefone, mensagemAlerta, mensagemErro)

    '======= grupo de identificação da NF-e - grupo B do Manual de integração - páginas 86 a 89
    Dim dhContingencia As String, justContingencia As String
    If Left(NFe!FormatoEmissaoNFe, 1) <> "1" Then
       dhContingencia = NFe!ContingenciaDataHora                                        'v2.03 - dhCont  AAAA-MM-DDTHH:MM:SS
       justContingencia = NFe!ContingenciaJustificativa                                 'v2.03 - xJust Justificativa da entrada em contingência
    End If

    If Not Vazio(Destinatario!IE) Then destIE = Destinatario!IE
    If Vazio(destIE) Then NFe!IndicadorIEDestinatario = "9"
    Dim indFinal As Integer
    indFinal = 0
    If Len(Destinatario!CPF) = 18 And Len(destIE) = 0 Then   '18 é cnpj
        indFinal = 1
    ElseIf Len(Destinatario!CPF) = 14 And Len(destIE) = 0 Then   '14 é cpf
        indFinal = 1
    ElseIf Len(destIE) > 0 Then
        indFinal = 0
    End If                          'Indica operação com Consumidor final
    
    If NFe!ConsumidorFinal Then
       indFinal = 1
    Else
       indFinal = 0
    End If

    dhEmi = Format(NFe!DataEmissao, "yyyy/mm/dd")

    iRetorno = sistNFe.GeraIdentificacao(NFe!cCodigoNota, NFe!NaturezaOperacao, 55, NFe!SerieNF, NFe!NumeroNota, Format(NFe!DataEmissao, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ss") & UTC, Format(NFe!DataSaida, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ss") & UTC, NFe!TipoDocumento, 0, NFe!IdentificadorDestino, Parametros!CodigoIBGE, 1, Left(NFe!FormatoEmissaoNFe, 1), Left(NFe!FinalidadeEmissaoNFe, 1), indFinal, 1, Left$("ONLINE COMMERCE - v." & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision), 20), dhContingencia, justContingencia, 0, 0, "", "", mensagemAlerta, mensagemErro)
   
    If NFe!ChavedeAcessoAdicional <> "" Then
       iRetorno = sistNFe.GerarNotasReferenciadas("NFe", NFe!ChavedeAcessoAdicional, 0, "", "", "", "", "", "", mensagemAlerta, mensagemErro)
    End If

    vsSQL = "SELECT * " & _
            "FROM NotaFiscalReferenciada " & _
            "WHERE CodigoNota = " & NumeroNota
    RsOpen NFeReferenciadas, vsSQL
    
    If NFeReferenciadas.RecordCount > 0 Then
       If NFeReferenciadas!ProdutorRural Then          'NFe Referenciada -> NF de Produtor referenciada
          iRetorno = sistNFe.GerarNotasReferenciadas("NFP", NFeReferenciadas!NumeroNF, NFeReferenciadas!SerieNFRef, NFeReferenciadas!ModeloNF, NFeReferenciadas!CodigoUF, NFeReferenciadas!AnoMesEmissaoNFe, "", NFeReferenciadas!CNPJ_CPF, NFeReferenciadas!InscricaoEstadual, mensagemAlerta, mensagemErro)
       ElseIf NFeReferenciadas!ModeloNF = "55" Or NFeReferenciadas!ModeloNF = "65" Then   'NFe Referenciada -> NFe Complementar, Devolução, Retorno
          iRetorno = sistNFe.GerarNotasReferenciadas("NFe", NFeReferenciadas!ChavedeAcesso, 0, "", "", "", "", "", "", mensagemAlerta, mensagemErro)
       ElseIf NFeReferenciadas!ModeloNF = "57" Then    'NFe Referenciada -> CTe
          iRetorno = sistNFe.GerarNotasReferenciadas("CTe", NFeReferenciadas!ChaveCTe, 0, "", "", "", "", "", "", mensagemAlerta, mensagemErro)
       ElseIf NFeReferenciadas!CupomFiscal Then        'NFe Referenciada -> ECF
          iRetorno = sistNFe.GerarNotasReferenciadas("ECF", NFeReferenciadas!nCOO, NFeReferenciadas!nECF, "", "", "", "", "", "", mensagemAlerta, mensagemErro)
       End If
    End If
   
    '================grupo de identificação do destinatario (grupo E do Manual de integração - páginas 92)=======================
    Dim xRazaoSocial As String, xTelefone As String
    'CLIENTE COM CPF
    If Len(Destinatario!CPF) = 18 Then
      xCNPJ = Trim(Destinatario!CPF)                                   ' CNPJ do destinatario sem máscara de formatação
      xCPF = ""
    Else
      xCPF = Trim(Destinatario!CPF)                                   ' CPF do destinatario, uso exclusivo do Fisco
      xCNPJ = ""
    End If

    If NFe!TipoCliente = "FORNECEDOR" Then
        xRazaoSocial = RemoveAcento(Trim(Left(Destinatario!Razao, 60)))           ' Razão social do destinatario, evitar caracteres acentuados e &
        xTelefone = Trim(Retira(ValidateNull(Destinatario!telefone), "()-. ", UM_A_UM))     ' número do telefone sem máscara
    Else
        xRazaoSocial = RemoveAcento(Trim(Left(Destinatario!Nome, 60)))           ' Razão social do destinatario, evitar caracteres acentuados e &
        xTelefone = Trim(Retira(ValidateNull(Destinatario!Telefone1), "()-. ", UM_A_UM))   ' número do telefone sem máscara
    End If
      
If Parametros!AmbienteNF = 2 Then
    xCNPJ = "99999999000191"
    xRazaoSocial = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
    destIE = "55931575"
    Destinatario!CodigoIBGE = 2927408
    Destinatario!Cidade = "SALVADOR"
    Destinatario!Estado = "BA"
    NFe!IndicadorIEDestinatario = "1"
End If
    
    iRetorno = sistNFe.GerarDestinatario(4, xRazaoSocial, xCNPJ, xCPF, "", destIE, NFe!InscricaoMunicipal, Left$(NFe!IndicadorIEDestinatario, 1), "", RemoveAcento(Destinatario!Endereco), Destinatario!Numero, RemoveAcento(ValidateNull(Destinatario!Ponto_de_referencia)), RemoveAcento(Destinatario!bairro), Destinatario!CodigoIBGE, RemoveAcento(Destinatario!Cidade), Destinatario!Estado, Retira(Destinatario!CEP, ".- ", UM_A_UM), 1058, "BRASIL", xTelefone, ValidateNull(Destinatario!Correio_eletronico), mensagemAlerta, mensagemErro)
    
    'Grupo de identificação do Local de RETIRADA
    'Informar apenas quando for diferente do endereço do remetente.
    'dest(29) a dest(36)
    'dest(29) = CNPJ ou CPF
    'dest(30) = xLgr
    'dest(31) = nro
    'dest(32) = xCpl
    'dest(33) = xBairro
    'dest(34) = cMun
    'dest(35) = xMun
    'dest(36) = UF
        
    'Grupo de identificação do Local de ENTREGA
    'Informar apenas quando for diferente do endereço do remetente.
    'dest(37) a dest(44)
    'dest(37) = CNPJ ou CPF
    'dest(38) = xLgr
    'dest(39) = nro
    'dest(40) = xCpl
    'dest(41) = xBairro
    'dest(42) = cMun
    'dest(43) = xMun
    'dest(44) = UF
    
    
'    If IsNull(Parametros!CNPJcontabil) = False Then
'        vsSQL = "SELECT * FROM NotaFiscalAutorizados WHERE CodigoNota = " & NumeroNota
'        RsOpen NFeAutorizados, vsSQL
       
'        If NFeAutorizados.BOF Then
'            CNPJContavil = ValidateNull(Parametros!CNPJcontabil)
'            CNPJContavil = Replace(CNPJContavil, ".", "")
'            CNPJContavil = Replace(CNPJContavil, "/", "")
'            CNPJContavil = Replace(CNPJContavil, "-", "")
             
'            vsSQL = "INSERT INTO NotaFiscalAutorizados  (CodigoNota, CNPJCPF) Values " & _
'                   "(" & NumeroNota & ", '" & CNPJContavil & "')"
'            vgDb.Execute vsSQL
'        End If
    
'        For i = 1 To NFeAutorizados.RecordCount
'           xCNPJ = ""
'           xCPF = ""
'           'If Len(NFeAutorizados!CNPJCPF) = 14 Then
'              xCNPJ = NFeAutorizados!CNPJCPF
'           'Else
'              'xCPF = NFeAutorizados!CNPJCPF
'           'End If
'
'           iRetorno = sistNFe.GerarAutorizadosXML(xCNPJ, xCPF, mensagemAlerta, mensagemErro)
'        Next
'    End If
        
    vsSQL = "SELECT NotaFiscalItens.* " & _
            "FROM NotaFiscalItens " & _
            "WHERE CodigoNota = " & NumeroNota & " " & _
            "ORDER BY Item"

    RsOpen NFeItens, vsSQL
    
    n = NFeItens.RecordCount

    TotvBCCBSIBS = 0
    TotvIBSUF = 0
    TotvIBSMun = 0
    TotvIBS = 0
    TotvCBS = 0

    For i = 1 To n
        vsSQL = "SELECT * FROM produtos WHERE CODIGO = " & NFeItens!CodigoProduto
        RsOpen Produtos, vsSQL
        
        If Produtos.RecordCount = 0 Then
           NFeMotivo = "CADASTRO DO PRODUTO NÃO ENCONTRADO!" & vbNewLine & vbNewLine & "PRODUTO: " & NFeItens!NomeProduto
           GoTo Caifora
        End If
        
        '================grupo de detalhe do produto (grupo I01 do Manual de integração - páginas 95)=======================
        Dim infAdiProd As String
        If NFeItens!ValorTributos > 0 Then
           infAdiProd = RemoveAcento(Trim(Left(NFeItens!InformacoesAdicionaisProduto, 400))) & " - Valor Aproximado dos Tributos R$ " & Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)        ' informações adicionais do produto
           vlTrib = vlTrib + NFeItens!ValorTributos
        Else
           infAdiProd = RemoveAcento(Trim(Left(NFeItens!InformacoesAdicionaisProduto, 500)))     ' informações adicionais do produto
        End If
        
        If NFeItens!TipoProduto = "Combustível" Then infAdiProd = "ICMS monofasico sobre combustiveis cobrado anteriormente conforme Convenio ICMS 199/2022. " + infAdiProd
        
        infAdiProd = Trim(infAdiProd)
        
        iRetorno = sistNFe.GerarItens(i, NFeItens!CodigoProduto, RemoveAcento(NFeItens!NomeProduto), Produtos!NCM, "", "", IIf(Vazio(Produtos!EAN), "SEM GTIN", Produtos!EAN), IIf(Vazio(Produtos!EAN), "SEM GTIN", Produtos!EAN), _
                                      NFeItens!CFOP, NFeItens!QuantidadeComercial, NFeItens!ValorUnitarioComercializacao, NFeItens!UnidadeComercial, NFeItens!QuantidadeComercial, NFeItens!ValorUnitarioComercializacao, NFeItens!UnidadeComercial, Round(NFeItens!QuantidadeComercial * NFeItens!ValorUnitarioComercializacao, 2), NFeItens!ValorFrete, NFeItens!ValorDesconto, _
                                      NFeItens!ValorOutros, NFeItens!ValorSeguro, "", "", 0, "", "", "", "", "", 1, infAdiProd, 0, "", 0, mensagemAlerta, mensagemErro)

'ESSE PONTO DO DEPOSITO DE GÁS

        Select Case NFeItens!TipoProduto     'Veículo|Medicamento|Armamento|Combustível
          Case "Armamento"
'            vsSQL = "SELECT * " & _
'                    "FROM NotaFiscalItensArmamento " & _
'                    "WHERE (NumeroNota = " & NumeroNota & " AND SerieNF = " & SerieNF & ") AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
'            RsOpen NFeArmamento, vsSQL
'            If NFeArmamento.RecordCount > 0 Then
'              Do While Not NFeArmamento.EOF
'                'Prod_DetEspecifico = Prod_DetEspecifico + objNFeUtil.arma(NFeArmamento!tpArma, NFeArmamento!nSerie, NFeArmamento!nCano, NFeArmamento!ArmDescricao)
'                NFeArmamento.MoveNext
'              Loop
'            End If
'
'
          Case "Combustível"
                vsSQL = "SELECT Cod_Produto, CODIF, cProdANP, descricaoANP, pGLP, pGNi, pGNn, pMixGN, ValorPartida " & _
                        "FROM Produtos_Gas " & _
                        "WHERE Cod_Produto = " & NFeItens!CodigoProduto
                RsOpen NFeCombustivel, vsSQL
                        
                If NFeCombustivel.RecordCount > 0 Then
                   iRetorno = sistNFe.GerarCombustivel(NFeCombustivel!CODIF, NFeCombustivel!cProdANP, NFeCombustivel!descricaoANP, NFeCombustivel!pGLP, NFeCombustivel!pGNi, NFeCombustivel!pGNn, NFeCombustivel!pMixGN, 0, Parametros!Estado, NFeCombustivel!ValorPartida, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, mensagemAlerta, mensagemErro)
                End If
'
'
          Case "Medicamento"
'            vsSQL = "SELECT * " & _
'                    "FROM NotaFiscalItensMedicamento " & _
'                    "WHERE (NumeroNota = " & NumeroNota & " AND SerieNF = " & SerieNF & ") AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
'            RsOpen NFeMedicamentos, vsSQL
'            If NFeMedicamentos.RecordCount > 0 Then
'              Do While Not NFeMedicamentos.EOF
'                PROD(i, 71) = IIf(Vazio(NFeMedicamentos!nLote), "0", NFeMedicamentos!nLote)
'                PROD(i, 72) = Format(NFeMedicamentos!QuantidadeLote, "#0.000")
'                PROD(i, 73) = IIf(IsNull(NFeMedicamentos!DataFabricacao), Format(DateAdd("yyyy", -1, Date), "mm/dd/yyyy"), Format(NFeMedicamentos!DataFabricacao, "yyyy-mm-dd"))
'                PROD(i, 74) = IIf(IsNull(NFeMedicamentos!DataValidade), Format(DateAdd("m", 6, NFeMedicamentos!DataValidade), "mm/dd/yyyy"), Format(NFeMedicamentos!DataValidade, "yyyy-mm-dd"))
'                PROD(i, 75) = Format(NFeMedicamentos!PMC, "#0.00")
'                NFeMedicamentos.MoveNext
'              Loop
'            End If
          Case "Veículo"
'            vsSQL = "SELECT * " & _
'                    "FROM NotaFiscalItensVeiculos " & _
'                    "WHERE (NumeroNota = " & NumeroNota & " AND SerieNF = " & SerieNF & ") AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
'            RsOpen NFeVeiculos, vsSQL
'            If NFeVeiculos.RecordCount > 0 Then
'              Do While Not NFeVeiculos.EOF
'                'Prod_Renavam = LPad(Retira(NFeVeiculos!RENAVAM, ".", UM_A_UM), 9, "0")
'                'Prod_DetEspecifico = Prod_DetEspecifico + objNFeUtil.veicProd(Left(NFeVeiculos!TipoOperacao, 1), NFeVeiculos!Chassi, NFeVeiculos!Cor, NFeVeiculos!DescricaoCor, NFeVeiculos!PotenciaMotor, NFeVeiculos!CM3, NFeVeiculos!VeicPesoLiquido, NFeVeiculos!VeicPesoBruto, NFeVeiculos!VeicSerie, NFeVeiculos!VeicTipoCombustivel, NFeVeiculos!VeicNumeroMotor, NFeVeiculos!CMKG, NFeVeiculos!DistanciaentreEixos, Prod_Renavam, NFeVeiculos!AnoMod, NFeVeiculos!AnoFab, NFeVeiculos!tpPintura, NFeVeiculos!tpVeiculo, NFeVeiculos!espVeiculo, NFeVeiculos!VIN, NFeVeiculos!condVeiculo, NFeVeiculos!cModelo)
'                'Prod_DetEspecifico = Prod_DetEspecifico + objNFeUtil.veicProd2G(Left(NFeVeiculos!TipoOperacao, 1), NFeVeiculos!Chassi, NFeVeiculos!Cor, NFeVeiculos!DescricaoCor, NFeVeiculos!PotenciaMotor, NFeVeiculos!CM3, NFeVeiculos!VeicPesoLiquido, NFeVeiculos!VeicPesoBruto, NFeVeiculos!VeicSerie, NFeVeiculos!VeicTipoCombustivel, NFeVeiculos!VeicNumeroMotor, NFeVeiculos!CMKG, NFeVeiculos!DistanciaentreEixos, NFeVeiculos!AnoMod, NFeVeiculos!AnoFab, NFeVeiculos!tpPintura, NFeVeiculos!tpVeiculo, NFeVeiculos!espVeiculo, NFeVeiculos!VIN, NFeVeiculos!condVeiculo, NFeVeiculos!cModelo, Left(NFeVeiculos!cCorDENATRAN, 1), NFeVeiculos!veicLotacao, Left(NFeVeiculos!veictpRestricao, 1))
'                NFeVeiculos.MoveNext
'              Loop
'              PROD(i, 79) = Prod_DetEspecifico
'            End If
        End Select

        'Valor aproximado total de tributos federais, estaduais e municipais
        'If NFeItens!ValorTributos > 0 Then prod(i, 81) = Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)
        
'FIM DO PONTO DO DEPOSITO DE GAS
        
        '=========dados do ICMS (grupo N01 do Manual de integração - páginas 100)=====================
        Dim sCSOSNCheck As String
        sCSOSNCheck = Right(Format(NFeItens!CST, "@"), 3)
        Dim dblPCredSN As Double
        dblPCredSN = 0
        vCredICMSSN = 0
        If Left(Parametros!CRT, 1) < 3 And (sCSOSNCheck = "101" Or sCSOSNCheck = "201") Then
            dblPCredSN = CDbl(Parametros!pCreditoICMSSimplesNacional)
            vCredICMSSN = IIf(IsNull(NFeItens!vCredICMSSN), 0, CDbl(NFeItens!vCredICMSSN))
            If vCredICMSSN = 0 And dblPCredSN > 0 Then ' fallback para notas antigas
                vCredICMSSN = Round(NFeItens!ValorTotalBruto * (dblPCredSN / 100), 2)
            End If
        End If
        Dim ICMSCST As String
        If Left(Parametros!CRT, 1) < 3 Then
           ICMSCST = NFeItens!CST
           If Len(NFeItens!CST) > 3 Then ICMSCST = Mid$(NFeItens!CST, 2)
        Else
           ICMSCST = NFeItens!CST
           If Len(NFeItens!CST) > 2 Then ICMSCST = Mid$(NFeItens!CST, 2)
           'If Len(NFeItens!CST) > 2 Then icmsCST = NFeItens!CST
        End If
        
        If NFeItens!TipoProduto <> "Combustível" Then
            'acrescentei um zero antes da mensagemAlerta: NFeItens!vBC, 0, 0, mensagemAlerta,
           iRetorno = sistNFe.GerarItensImpostoEstadual(NFeItens!ValorTributos, "0", ICMSCST, IIf(Vazio(NFeItens!modBC), 3, Left$(NFeItens!modBC, 1)), NFeItens!vBC, NFeItens!pICMS, NFeItens!vICMS, NFeItens!pRedBC, 0, 0, 0, _
                                                        IIf(Not Vazio(NFeItens!modBCST), Left(NFeItens!modBCST, 1), 5), NFeItens!pMVAST, NFeItens!pRedBCST, NFeItens!vBCST, NFeItens!pICMSST, NFeItens!vICMSST, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, dblPCredSN, vCredICMSSN, 0, NFeItens!vBC, 0, 0, mensagemAlerta, mensagemErro)
        
           If Left(Parametros!CRT, 1) = 3 Then
              vBCCBSIBS = NFeItens!ValorTotalBruto
              pIBSUF = 0.1
              vIBSUF = Round(NFeItens!ValorTotalBruto * (0.1 / 100), 2)
              pIBSMun = 0
              vIBSMun = 0
              vIBS = vIBSUF + vIBSMun
              pCBS = 0.9
              vCBS = Round(NFeItens!ValorTotalBruto * (0.9 / 100), 2)
              iRetorno = sistNFe.GerarItensImpostoIBSCBS("000", "000001", vBCCBSIBS, pIBSUF, 0, 0, 0, 0, 0, vIBSUF, pIBSMun, 0, 0, 0, 0, 0, vIBSMun, vIBS, pCBS, 0, 0, 0, 0, 0, vCBS, mensagemAlerta, mensagemErro)
           End If
        Else
           iRetorno = sistNFe.GerarItensImpostoEstadualMonofasico(NFeItens!ValorTributos, "0", "61", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, mensagemAlerta, mensagemErro)
           
           'iRetorno = sistNFe.GerarItensObservacao("CST61", "ICMS monofásico sobre combustíveis cobrado anteriormente conforme Convênio ICMS 199/2022;", "", "", mensagemAlerta, mensagemErro)
           
           If Left(Parametros!CRT, 1) = 3 Then
              vBCCBSIBS = NFeItens!ValorTotalBruto
              pIBSUF = 0.1
              vIBSUF = Round(NFeItens!ValorTotalBruto * (0.1 / 100), 2)
              pIBSMun = 0
              vIBSMun = 0
              vIBS = vIBSUF
              pCBS = 0.9
              vCBS = Round(NFeItens!ValorTotalBruto * (0.9 / 100), 2)
              iRetorno = sistNFe.GerarItensImpostoIBSCBS("000", "000001", vBCCBSIBS, pIBSUF, 0, 0, 0, 0, 0, vIBSUF, pIBSMun, 0, 0, 0, 0, 0, vIBSMun, vIBS, pCBS, 0, 0, 0, 0, 0, vCBS, mensagemAlerta, mensagemErro)
           End If
        End If
        
        TotvBCCBSIBS = TotvBCCBSIBS + vBCCBSIBS
        TotvIBSUF = TotvIBSUF + vIBSUF
        TotvIBSMun = TotvIBSMun + vIBSMun
        TotvIBS = TotvIBS + vIBS
        TotvCBS = TotvCBS + vCBS
        
        iRetorno = sistNFe.GerarItensImpostoFederal(IIf(Vazio(NFeItens!COFINSCST) Or IsNull(NFeItens!COFINSCST), "99", NFeItens!COFINSCST), NFeItens!COFINSvBC, NFeItens!COFINSpCOFINS, NFeItens!COFINSvCOFINS, 0, 0, _
                                                    IIf(Vazio(NFeItens!PISCST) Or IsNull(NFeItens!PISCST), "99", NFeItens!PISCST), NFeItens!PISvBC, NFeItens!PISpPIS, NFeItens!PISvPIS, 0, 0, _
                                                    IIf(Vazio(NFeItens!IPICST), "50", NFeItens!IPICST), NFeItens!IPIvBC, NFeItens!IPIpIPI, NFeItens!IPIvIPI, 0, 0, 999, "", "", "", 0, mensagemAlerta, mensagemErro)

        '   gera grupo do II - Importação
        If (NFeItens!IIvBC > 0) Then
           iRetorno = sistNFe.GerarItensImpostoII(NFeItens!IIvBC, NFeItens!IIvDespAdu, NFeItens!IIvII, NFeItens!IIvIOF, mensagemAlerta, mensagemErro)
        End If

        If NFe!IdentificadorDestino = 2 Then
            If NFe!ConsumidorFinal = True Then
                iRetorno = sistNFe.GerarItensImpostoUFDest(NFeItens!pICMSInter, NFeItens!pICMSInterPart, NFeItens!pICMSUFDest, NFeItens!pFCPUFDest, NFeItens!vBCFCPUFDest, NFeItens!vBCUFDest, NFeItens!vFCPUFDest, NFeItens!vICMSUFDest, NFeItens!vICMSUFRemet, mensagemAlerta, mensagemErro)
            End If
        End If
        
        iRetorno = sistNFe.GerarItensIncluir(mensagemAlerta, mensagemErro)
        
        NFeItens.MoveNext
    Next

    'atualização de total
    iRetorno = sistNFe.GerarTotalProdutos(NFe!BaseICMS, NFe!ValorICMS, NFe!BaseICMSST, NFe!ValorICMSST, NFe!ValorCOFINS, NFe!ValorPIS, NFe!ValorIPI, NFe!ValorDesconto, NFe!ValorSeguro, NFe!ValorFrete, NFe!ValorOutrasDespesas, 0, 0, 0, NFe!vFCPUFDest, 0, NFe!vICMSUFDest, NFe!vICMSUFRemet, _
                                          NFe!ValorImportacao, 0, NFe!ValorProdutos, NFe!ValorNota, vlTrib, 0, 0, 0, 0, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, mensagemAlerta, mensagemErro)

    If Left(Parametros!CRT, 1) = 3 Then
       iRetorno = sistNFe.GerarTotalIBSCBS(0, TotvBCCBSIBS, 0, 0, TotvIBSUF, 0, 0, TotvIBSMun, TotvIBS, 0, 0, 0, 0, TotvCBS, 0, 0, 0, 0, 0, 0, 0, 0, (TotvBCCBSIBS + TotvIBS + TotvCBS), mensagemAlerta, mensagemErro)
    End If
    
    '============dados do transportador
    xCNPJ = ""
    xCPF = ""
    If Not Vazio(NFe!TranspCNPJ_CPF) Then
        If Len(Retira(NFe!TranspCNPJ_CPF, ".-/", UM_A_UM)) = 14 Then
          xCNPJ = Trim(NFe!TranspCNPJ_CPF)                                         ' CNPJ da Transportadora sem mascara
        Else
          xCPF = Trim(NFe!TranspCNPJ_CPF)                                          ' CPF da Transportadora sem mascara
        End If
    End If
    iRetorno = sistNFe.GerarTransporte(IIf(Vazio(NFe!modFrete), 9, Left(NFe!modFrete, 1)), NFe!VolumeQuantidade, NFe!VolumeEspecie, NFe!VolumeMarca, NFe!VolumeNumeracao, NFe!VolumePesoBruto, NFe!VolumePesoLiquido, _
                                       xCNPJ, xCPF, NFe!TranspInscricaoEstadual, RemoveAcento(Trim(NFe!TranspNome)), RemoveAcento(Trim(NFe!TranspEndereco)), RemoveAcento(Trim(NFe!TranspMunicipio)), NFe!TranspUF, _
                                       NFe!TranspPlaca, NFe!TranspRNTC, NFe!TranspPlacaUF, mensagemAlerta, mensagemErro)
    'parcelas e pagamentos
    If Left(NFe!FinalidadeEmissaoNFe, 1) = 4 Then  'devolução
        iRetorno = sistNFe.GerarPagamentos(0, 90, 0, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)
    ElseIf Left(NFe!FinalidadeEmissaoNFe, 1) = 3 Then       'nota de ajuste
        iRetorno = sistNFe.GerarPagamentos(0, 90, 0, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)
    Else
        vsSQL = "SELECT * " & _
                "FROM NotaFiscalParcelas " & _
                "WHERE CodigoNota = " & NumeroNota
        RsOpen NFeParcelas, vsSQL
    
        'iRetorno = sistNFe.GerarCobranca(NFe!NumeroNota, 0, NFe!ValorNota, NFe!ValorNota, mensagemAlerta, mensagemErro)
        iRetorno = sistNFe.GerarCobranca(NFe!NumeroNota, NFe!ValorDesconto, NFe!ValorProdutos, NFe!ValorNota, mensagemAlerta, mensagemErro)
        
        If NFeParcelas.RecordCount > 0 Then     'se tem duplicatas à prazo
           For i = 0 To NFeParcelas.RecordCount - 1
               iRetorno = sistNFe.GerarCobrancaDuplicatas(LPad(i + 1, 3, "0"), NFeParcelas!Vencimento, NFeParcelas!ValorDocumento, mensagemAlerta, mensagemErro)
               iRetorno = sistNFe.GerarPagamentos(1, 15, NFeParcelas!ValorDocumento, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)
               NFeParcelas.MoveNext
           Next
        Else
           'iRetorno = sistNFe.GerarCobrancaDuplicatas("001", NFe!DataEmissao, NFe!ValorNota, mensagemAlerta, mensagemErro)
           Select Case Left(NFe!IndicadorFormaPagamento, 1)
               Case 1
                  'iRetorno = sistNFe.GerarPagamentos(1, 5, NFe!ValorNota, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)
                    Select Case Left(NFe!FormaPagamento, 2)
                       Case "01"
                           iRetorno = sistNFe.GerarPagamentos(1, 1, NFe!ValorNota, 0, 0, 0, "", "", "Dinheiro", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "02"
                           iRetorno = sistNFe.GerarPagamentos(1, 2, NFe!ValorNota, 0, 0, 0, "", "", "Cheque", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "03"
                           iRetorno = sistNFe.GerarPagamentos(1, 3, NFe!ValorNota, 0, 0, 0, "", "", "Cartão de Crédito", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "04"
                           iRetorno = sistNFe.GerarPagamentos(1, 4, NFe!ValorNota, 0, 0, 0, "", "", "Cartão de Crédito", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "05"
                           iRetorno = sistNFe.GerarPagamentos(1, 5, NFe!ValorNota, 0, 0, 0, "", "", "Crédito Loja", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "10"
                           iRetorno = sistNFe.GerarPagamentos(1, 10, NFe!ValorNota, 0, 0, 0, "", "", "Vale Alimentação", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "11"
                           iRetorno = sistNFe.GerarPagamentos(1, 11, NFe!ValorNota, 0, 0, 0, "", "", "Vale Refeição", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "12"
                           iRetorno = sistNFe.GerarPagamentos(1, 12, NFe!ValorNota, 0, 0, 0, "", "", "Vale Presente", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "13"
                           iRetorno = sistNFe.GerarPagamentos(1, 13, NFe!ValorNota, 0, 0, 0, "", "", "Vale Combustível", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "14"
                           iRetorno = sistNFe.GerarPagamentos(1, 14, NFe!ValorNota, 0, 0, 0, "", "", "Duplicata Mercantil", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "15"
                           iRetorno = sistNFe.GerarPagamentos(1, 15, NFe!ValorNota, 0, 0, 0, "", "", "Boleto Bancário", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "16"
                           iRetorno = sistNFe.GerarPagamentos(1, 16, NFe!ValorNota, 0, 0, 0, "", "", "Depósito Bancário", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "18"
                           iRetorno = sistNFe.GerarPagamentos(1, 18, NFe!ValorNota, 0, 0, 0, "", "", "Transferência bancária", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "19"
                           iRetorno = sistNFe.GerarPagamentos(1, 19, NFe!ValorNota, 0, 0, 0, "", "", "Programa de fidelidade", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "20"
                           iRetorno = sistNFe.GerarPagamentos(1, 20, NFe!ValorNota, 0, 0, 0, "", "", "Programa de fidelidade", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "99"
                           iRetorno = sistNFe.GerarPagamentos(2, 99, NFe!ValorNota, 0, 0, 0, "", "", "Outros", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case Else
                           iRetorno = sistNFe.GerarPagamentos(2, 90, NFe!ValorNota, 0, 0, 0, "", "", "Sem pagamento", "", "", "", "", "", mensagemAlerta, mensagemErro)
                    End Select
               Case 2
                  iRetorno = sistNFe.GerarPagamentos(2, 99, NFe!ValorNota, 0, 0, 0, "", "", "OUTROS", "", "", "", "", "", mensagemAlerta, mensagemErro)
               Case Else
                  'desativei essa linha para gerar as parcelas à vista com varias formas de pagamento
                  'iRetorno = sistNFe.GerarPagamentos(0, 1, NFe!ValorNota, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)
                    Select Case Left(NFe!FormaPagamento, 2)
                       Case "01"
                           iRetorno = sistNFe.GerarPagamentos(0, 1, NFe!ValorNota, 0, 0, 0, "", "", "Dinheiro", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "02"
                           iRetorno = sistNFe.GerarPagamentos(0, 2, NFe!ValorNota, 0, 0, 0, "", "", "Cheque", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "03"
                           iRetorno = sistNFe.GerarPagamentos(0, 3, NFe!ValorNota, 0, 0, 0, "", "", "Cartão de Crédito", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "04"
                           iRetorno = sistNFe.GerarPagamentos(0, 4, NFe!ValorNota, 0, 0, 0, "", "", "Cartão de Crédito", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "05"
                           iRetorno = sistNFe.GerarPagamentos(0, 5, NFe!ValorNota, 0, 0, 0, "", "", "Crédito Loja", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "10"
                           iRetorno = sistNFe.GerarPagamentos(0, 10, NFe!ValorNota, 0, 0, 0, "", "", "Vale Alimentação", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "11"
                           iRetorno = sistNFe.GerarPagamentos(0, 11, NFe!ValorNota, 0, 0, 0, "", "", "Vale Refeição", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "12"
                           iRetorno = sistNFe.GerarPagamentos(0, 12, NFe!ValorNota, 0, 0, 0, "", "", "Vale Presente", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "13"
                           iRetorno = sistNFe.GerarPagamentos(0, 13, NFe!ValorNota, 0, 0, 0, "", "", "Vale Combustível", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "14"
                           iRetorno = sistNFe.GerarPagamentos(0, 14, NFe!ValorNota, 0, 0, 0, "", "", "Duplicata Mercantil", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "15"
                           iRetorno = sistNFe.GerarPagamentos(0, 15, NFe!ValorNota, 0, 0, 0, "", "", "Boleto Bancário", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "16"
                           iRetorno = sistNFe.GerarPagamentos(0, 16, NFe!ValorNota, 0, 0, 0, "", "", "Depósito Bancário", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "18"
                           iRetorno = sistNFe.GerarPagamentos(0, 18, NFe!ValorNota, 0, 0, 0, "", "", "Transferência bancária", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "19"
                           iRetorno = sistNFe.GerarPagamentos(0, 19, NFe!ValorNota, 0, 0, 0, "", "", "Programa de fidelidade", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "20"
                           iRetorno = sistNFe.GerarPagamentos(0, 20, NFe!ValorNota, 0, 0, 0, "", "", "Programa de fidelidade", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case "99"
                           iRetorno = sistNFe.GerarPagamentos(2, 99, NFe!ValorNota, 0, 0, 0, "", "", "Outros", "", "", "", "", "", mensagemAlerta, mensagemErro)
                       Case Else
                           iRetorno = sistNFe.GerarPagamentos(2, 90, NFe!ValorNota, 0, 0, 0, "", "", "Sem pagamento", "", "", "", "", "", mensagemAlerta, mensagemErro)
                    End Select
           End Select
           




        End If
    End If
    '============= informações adcionais

    vsSQL = "SELECT ObservacoesNFe.Observacao " & _
            "FROM NotaFiscalObservacoes INNER JOIN ObservacoesNFe ON NotaFiscalObservacoes.CodigoObservacao = ObservacoesNFe.CodigoObservacao " & _
            "WHERE CodigoNota = " & NumeroNota

    RsOpen NFeOBS, vsSQL

    If NFeOBS.RecordCount > 0 Then
    Dim OBSNFe As String
    OBSNFe = ""
      Do While Not NFeOBS.EOF
        OBSNFe = OBSNFe + NFeOBS!OBSERVACAO + " // "
        NFeOBS.MoveNext
      Loop
    End If

    iRetorno = sistNFe.GerarInformacoesAdicionais(RemoveAcento(Trim(NFe!InformacoesComplementares)) & IIf(Vazio(OBSNFe), "", " // " & RemoveAcento(Trim(OBSNFe))), RemoveAcento(Trim(NFe!InformacoesAdicionais)), mensagemAlerta, mensagemErro)
    
End If

dirXML = Parametros!DiretorioXML
dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")

'gera a chave da nfe
Dim id_chave As String, numero_nfe_gerado As String
'pega o endereço do arquivo a ser gerado
If Not Existe(dirXML) Then MkDir dirXML
'If iRetorno = -1 Then GoTo NaoEnviou
iRetorno = sistNFe.GerarXML(numero_nfe_gerado, xCaminhoXML, True, xCaminhoXMLAuxiliar, mensagemAlerta, mensagemErro)
id_chave = numero_nfe_gerado
numero_nfe_gerado = Replace(numero_nfe_gerado, "NFe", "")
NFeChaveAcesso = numero_nfe_gerado

If Not Vazio(NFeChaveAcesso) Then
   vsSQL = "UPDATE NotaFiscal SET " & _
           "ChavedeAcesso = '" & NFeChaveAcesso & "' " & _
           "WHERE CodigoNota = " & NumeroNota
   vgDb.Execute vsSQL
End If

If Not PodeEnviar Then GoTo NaoPodeEnviar

iRetorno = sistNFe.EnviarNFe(NFe!NumeroNota, 1, False)

cStat = sistNFe.retEnvio.cStat
NFeMotivo = sistNFe.retEnvio.xMotivo

If cStat = 104 Then
   cStat = sistNFe.retEnvio.protNFe.infProt.cStat
   NFeMotivo = sistNFe.retEnvio.protNFe.infProt.xMotivo
End If

If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
   MsgBox "*** Aparentemente Ocorreram Erros na Recepção do Lote (nfeAutorizacao)***" & vbLf & NFeMotivo, vbExclamation, "PROCESSO INTERROMPIDO - NfeAutorizacao"
   GoTo Caifora
End If

If cStat = 103 Then NFeNumeroRecibo = sistNFe.retEnvio.infRec.nRec
NFeDataHora = sistNFe.retEnvio.dhRecbto

If cStat <> 103 Then
   If cStat = 104 Or cStat = 100 And Vazio(NFeNumeroRecibo) Then GoTo buscaNFe
   GoTo NaoEnviou
End If

vsSQL = "UPDATE NotaFiscal SET " & _
        "NumeroRecibo = " & NFeNumeroRecibo & " " & _
        "WHERE CodigoNota = " & NumeroNota
vgDb.Execute vsSQL

DoEvents

consultaNFe:
    On Error Resume Next
    iRetorno = sistNFe.ConsultarReciboDeEnvio(NFeNumeroRecibo)
   
    cStat = sistNFe.retConsRec.cStat
    NFeMotivo = sistNFe.retConsRec.xMotivo
    
    If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
       MsgBox NFeMotivo, vbExclamation, "Retorno Autorização"
       GoTo Caifora
    End If

    ' Testa erro 217-Rejeição: NF-e não consta na base de dados da SEFAZ <dhRecbto xmlns="http://www.portalfiscal.inf.br/nfe">2015-03-29T09:56:36-03:00
    If cStat = 217 Then
            Sleep 3000 ' Aguarda mais 3 segundos
            On Error Resume Next
            iRetorno = sistNFe.ConsultarReciboDeEnvio(NFeNumeroRecibo) 'refaz a consulta
            cStat = sistNFe.retConsRec.cStat
            NFeMotivo = sistNFe.retConsRec.xMotivo
            If cStat = 104 Then
               cStat2 = sistNFe.retConsRec.protNFe.infProt.cStat
               NFeMotivo = sistNFe.retConsRec.protNFe.infProt.xMotivo
            End If
            If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
               MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote (nfeRetAutorizacao)***" & vbLf & NFeMotivo, vbExclamation, "PROCESSO INTERROMPIDO - NfeRetAutorizacao"
               GoTo Caifora
            End If
    End If

    If cStat = 105 Or cStat = 217 Then
       i = i + 1
       If i > 5 And cStat = 105 Then
          msgResultado = "A Nota Fiscal ficou em PROCESSAMENTO." & vbNewLine & "Efetue a consulta do Recibo daqui 5 minutos novamente."
          MsgBox msgResultado, vbInformation + vbOKOnly, "Transmitir NFe"
          vsSQL = "UPDATE NotaFiscal SET EmProcessamento = 1 WHERE CodigoNota = " & NumeroNota
          vgDb.Execute vsSQL
          GoTo PodeSair
       End If
       Sleep 10000
       GoTo buscaNFe
    End If
    NFeNumeroProtocolo = sistNFe.retConsRec.protNFe.infProt.nProt
    NFeDataHora = sistNFe.retConsRec.protNFe.infProt.ProxyDhRecbto
    cStat2 = sistNFe.retConsRec.protNFe.infProt.cStat
    nfeRetorno = sistNFe.retConsRec.protNFe.infProt.xMotivo
    NFeValidate = sistNFe.retConsRec.protNFe.infProt.chNFe
    DoEvents
    
    If cStat2 <> 100 And cStat2 <> 101 And cStat2 <> 110 And cStat2 <> 150 Then
       NFeMotivo = ""
       MsgBox str$(cStat2) & " - " & nfeRetorno, vbExclamation, "Retorno Consulta Recibo"
       If cStat2 = 206 Then
          vsSQL = "UPDATE NotaFiscal SET Inutilizada = 1 WHERE CodigoNota = " & NumeroNota
          vgDb.Execute vsSQL
          GoTo PodeSair
       End If
       GoTo Caifora
    End If

buscaNFe:
   'Consulta Nfe
   iRetorno = sistNFe.ConsultarProtocolo(NFeChaveAcesso)
   
   cStat = sistNFe.retConsulta.cStat
   NFeMotivo = sistNFe.retConsulta.xMotivo
   
   If cStat = 217 Then
      Sleep 3000 ' aguarda mais 3 segundos
      iRetorno = sistNFe.ConsultarProtocolo(NFeChaveAcesso)
      cStat = sistNFe.retConsulta.cStat
      NFeMotivo = sistNFe.retConsulta.xMotivo
         
      If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
         MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote ***" & vbLf & NFeMotivo, vbExclamation, "PROCESSO INTERROMPIDO - NfeConsulta"
         GoTo Caifora
      End If
   End If

   If Not iRetorno Then
      NFeNumeroProtocolo = ""
      GoTo Caifora
   End If
    
   NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
   NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
   
   DoEvents

   If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo Caifora
  
   If cStat = 204 Or cStat = 539 Then
      NFeChaveAcesso = Mid(nfeRetorno, InStr(nfeRetorno, "chNFe:") + 6, 44)
      nroRecibo = Mid(nfeRetorno, InStr(nfeRetorno, "nRec:") + 5)
      nroRecibo = Left(nroRecibo, Len(nroRecibo) - 1)
      NFeNumeroRecibo = Left(NFeNumeroRecibo, 15 - Len(nroRecibo)) + nroRecibo
      If Vazio(NFeChaveAcesso) Or Len(NFeNumeroRecibo) < 15 Then
         NFeMotivo = nfeRetorno
         GoTo Caifora
      End If
      vsSQL = "UPDATE NotaFiscal SET " & _
              "ChavedeAcesso = '" & NFeChaveAcesso & "', " & _
              "NumeroRecibo = " & NFeNumeroRecibo & " " & _
              "WHERE CodigoNota = " & NumeroNota
      vgDb.Execute vsSQL
      GoTo buscaNFe
   ElseIf cStat <> 100 And cStat <> 301 Then
      NFeMotivo = str$(cStat) + " - " + NFeMotivo
      GoTo NaoEnviou
   ElseIf cStat = 105 And cStat = 217 Then
      GoTo consultaNFe
   ElseIf cStat = 100 Then
      nfeRetorno = "Nota Fiscal Eletronica Autorizado o Uso."
      msgResultado = "Chave NF-e.: " + NFeChaveAcesso & vbCrLf
      msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
      'msgResultado = msgResultado + "Recibo.: " + NFeNumeroRecibo & vbCrLf
      msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
      msgResultado = msgResultado + "Resposta da Fazenda.: " + str$(cStat) & " - " & NFeMotivo
        
      MsgBox msgResultado, vbInformation + vbOKOnly
       
      DoEvents
        
      iRetorno = sistNFe.ConsultarProtocolo(NFeChaveAcesso)
             
      vsSQL = "UPDATE NotaFiscal SET " & _
              "ChavedeAcesso = '" & NFeChaveAcesso & "', " & _
              "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
              "DataHoraProcotolo = '" & NFeDataHora & "' " & _
              "WHERE CodigoNota = " & NumeroNota
      vgDb.Execute vsSQL
      
      If SQLExecutaRetorno("SELECT ISNULL(COUNT(NumeroProtocolo), 0) r FROM NotaFiscalRecibos WHERE CodigoNota = " & NumeroNota & " AND NumeroProtocolo = " & NFeNumeroProtocolo, "r", 0) = 0 Then
         vsSQL = "INSERT INTO NotaFiscalRecibos (CodigoNota, NumeroProtocolo, DataHora) Values " & _
                 "(" & NumeroNota & ", " & NFeNumeroProtocolo & ", '" & NFeDataHora & "')"
         vgDb.Execute vsSQL
      End If
    End If

    'Gera PDF DANFE
    'NFeDataHora = NFe!DataEmissao
    dhProtocolo = Format(Left(NFeDataHora, 10), "yyyy/mm/dd")
    xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
    xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & Format(CDate(dhProtocolo), "yyyymm") & "\" & NFeChaveAcesso & "-procNFe.xml"            'Aqui Gera o DANFE
    If Existe(xCaminhoXML) Then Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 0, False, False, "", True, False, False, False, True) 'gera pdf

NaoPodeEnviar:
    Set sistNFe = Nothing
    Set Parametros = Nothing
    Set Destinatario = Nothing
    Set Produtos = Nothing
    Set NFe = Nothing
    Set NFeItens = Nothing
    Set NFeParcelas = Nothing
    Set NFeOBS = Nothing
    Set NFeArmamento = Nothing
    Set NFeCombustivel = Nothing
    Set NFeMedicamentos = Nothing
    Set NFeVeiculos = Nothing
    
    Screen.MousePointer = vbDefault
    TransmitirNFe = True
    'NFe_Completa.picAguarde.Visible = False    'desativei para funcionando pdv
    Exit Function

PodeSair:
    Set sistNFe = Nothing
    Set Parametros = Nothing
    Set Destinatario = Nothing
    Set Produtos = Nothing
    Set NFe = Nothing
    Set NFeItens = Nothing
    Set NFeParcelas = Nothing
    Set NFeOBS = Nothing
    Set NFeArmamento = Nothing
    Set NFeCombustivel = Nothing
    Set NFeMedicamentos = Nothing
    Set NFeVeiculos = Nothing
    
    Screen.MousePointer = vbDefault
    TransmitirNFe = True
    'NFe_Completa.picAguarde.Visible = False    'desativei para funcionando pdv
    Exit Function
Resume
NaoEnviou:
    Set sistNFe = Nothing
    Set Parametros = Nothing
    Set Destinatario = Nothing
    Set Produtos = Nothing
    Set NFe = Nothing
    Set NFeItens = Nothing
    Set NFeParcelas = Nothing
    Set NFeOBS = Nothing
    Set NFeArmamento = Nothing
    Set NFeCombustivel = Nothing
    Set NFeMedicamentos = Nothing
    Set NFeVeiculos = Nothing
    
    If PodeEnviar Then MsgBox NFeMotivo, vbCritical + vbOKOnly
    Screen.MousePointer = vbDefault
    TransmitirNFe = False
    'NFe_Completa.picAguarde.Visible = False    'desativei por causa do pdv
    Exit Function
    
    Resume
    
Caifora:
    If Not Vazio(NFeMotivo) Then MsgBox NFeMotivo, vbCritical + vbOKOnly
    
    Set sistNFe = Nothing
    Set Parametros = Nothing
    Set Destinatario = Nothing
    Set Produtos = Nothing
    Set NFe = Nothing
    Set NFeItens = Nothing
    Set NFeParcelas = Nothing
    Set NFeOBS = Nothing
    Set NFeArmamento = Nothing
    Set NFeCombustivel = Nothing
    Set NFeMedicamentos = Nothing
    Set NFeVeiculos = Nothing
    
    Screen.MousePointer = vbDefault
    TransmitirNFe = False
    'NFe_Completa.picAguarde.Visible = False    'desativei por causa do pdv
    Exit Function
    
    Resume
    
'DeuErro:
'    MsgBox Err.Description, vbCritical
'    Err.Clear
'    Screen.MousePointer = vbDefault
'    TransmitirNFe = False
End Function

Public Function TransmitirNFCe(ByVal NumeroNota As Variant, ByVal SerieNF As Variant, Optional PodeEnviar As Boolean = False, Optional ModeloNF As String = "65") As Boolean  'Função que monta o arquivo XML e faz o envio para a Receita
 Dim txtNumerado As String, Retorno As String, vsNFe As String, empUF As String, SQL As String
 Dim Parametros As New ADODB.Recordset
 Dim NFe As New ADODB.Recordset, NFeItens As New ADODB.Recordset, NFeParcelas As New ADODB.Recordset
 Dim NFeMedicamentos As New ADODB.Recordset, NFeArmamento As New ADODB.Recordset, NFeCombustivel As New ADODB.Recordset, NFeVeiculos As New ADODB.Recordset
 Dim NFeDeclaracaoImposto As New ADODB.Recordset, NFeAdicao As New ADODB.Recordset
          
 Dim n As Integer, i As Integer
 Dim vsXML As String, XMLAuxiliar As String, XMLAuxiliarParcelas As String
 Dim msgErro As String, qterro As Long, IdToken As String, Token As String
 Dim pDesconto As Double, pFrete As Double, pOutras As Double, pTributos As Double
 Dim vlPIS As Double, vlCOFINS As Double, vlTrib As Double, vlNF As Double
 Dim NFCeContingenciaOFF As Boolean
 
 'On Error GoTo TransmitirNFCe_Error
 
 On Error GoTo deuErro

 Dim sistNFCe As snfe.Util
 Set sistNFCe = New snfe.Util

 vlPIS = 0
 vlCOFINS = 0
 vlTrib = 0
 pFrete = 0
 pDesconto = 0
 pOutras = 0
 pTributos = 0
' vsSQL = "SELECT *, 0 AS COFINSAliquota, 0 AS PISAliquota FROM Empresa"
 vsSQL = "SELECT * FROM Empresa"
 RsOpen Parametros, vsSQL

 empUF = Parametros!Estado
 NFCeContingenciaOFF = Parametros!NFCeOffline

 If ModeloNF = "65" Then
    IdToken = LPad(Parametros!NFCeIDToken, 6, "0")
    Token = Parametros!NfceCsc
 End If

 Screen.MousePointer = vbHourglass

 vsSQL = "SELECT TbNFCe.*, Cidade.CodigoMunicipio CodigoIBGE " & _
         "FROM TbNFCe INNER JOIN Cidade ON TbNFCe.Municipio = Cidade.Nome AND TbNFCe.UF = Cidade.UF " & _
         "WHERE IdNFProd = " & NumeroNota
 RsOpen NFe, vsSQL
 'Set NFe = vgDb.OpenRecordset(vsSQL)

 If Not NFCeContingenciaOFF And PodeEnviar And Left(NFe!NFeTipoEmissao, 1) = 9 Then
    NFeChaveAcesso = NFe!NFCeChaveAcesso
    dirXML = Parametros!DiretorioXML
    dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
    'pega o endereço do arquivo a ser gerado
    If Not Existe(dirXML) Then MkDir dirXML
    
    xCaminhoXML = dirXML & "\nfe\arquivos\assinado\NFe" & NFeChaveAcesso & "-assinado.xml"
    xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
    iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFCe)
    DoEvents
    iRetorno = sistNFCe.CarregarXML(xCaminhoXML, True)
    
    GoTo SoEnviar
 End If

 If NFe.RecordCount > 0 Then
    NFe.MoveFirst
    iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFCe)
    sistNFCe.exibirAvisos = True
    '===================grupo de identificação do emitente (grupo B do Manual de integração - páginas 90)=======================
    iRetorno = sistNFCe.IncluirNF(mensagemAlerta, mensagemErro)
    
    '===================grupo de identificação do emitente (grupo B do Manual de integração - páginas 90)=======================
    iRetorno = sistNFCe.GerarEmitente(RemoveAcento(Parametros!Razao), RemoveAcento(Parametros!Fantasia), Parametros!CNPJ, "", Parametros!IE, "", "", "", Left(Parametros!CRT, 1), RemoveAcento(Parametros!Endereco), Parametros!Numero, "", Parametros!bairro, Parametros!CodigoIBGE, RemoveAcento(Parametros!Cidade), Parametros!Estado, Parametros!CEP, 1058, "BRASIL", Parametros!Celular, mensagemAlerta, mensagemErro)

    '======= grupo de identificação da NF-e - grupo B do Manual de integração - páginas 86 a 89
    Dim dhContingencia As String, justContingencia As String
    If Left(NFe!NFeTipoEmissao, 1) <> "1" Then
       dhContingencia = NFe!NFCeDataHoraContingencia & UTC 'v2.03 - dhCont  AAAA-MM-DDTHH:MM:SS
       justContingencia = NFe!NFCeJustificativaContingencia                                 'v2.03 - xJust Justificativa da entrada em contingência
    End If
    
    Dim indFinal As Integer
    indFinal = 0
    If Len(NFe!CPF_CNPJ) = 18 And Len(NFe!InscEst) = 0 Then   '18 é cnpj
        indFinal = 1
    ElseIf Len(NFe!CPF_CNPJ) = 14 And Len(NFe!InscEst) = 0 Then   '14 é cpf
        indFinal = 1
    ElseIf Len(NFe!InscEst) > 0 Then
        indFinal = 0
    End If                          'Indica operação com Consumidor final
    
    If NFe!NFeConsumidorFinal Then
       indFinal = 1
    Else
       indFinal = 0
    End If

    Dim dhEmiSaiEnt As String
    If Not IsNull(NFe!DataSaidaEntrada) Then dhEmiSaiEnt = Format(NFe!DataSaidaEntrada, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ss") & UTC
    iRetorno = sistNFCe.GeraIdentificacao(NFe!NFeCodigoNota, NFe!NaturezaOperacao, 65, CLng(NFe!SerieNF), NFe!NumeNota, Format(NFe!DataEmissao, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ss") & UTC, dhEmiSaiEnt, 1, 0, Left(NFe!NFeIdentificadorDestino, 1), Parametros!CodigoIBGE, 4, Left(NFe!NFeTipoEmissao, 1), 1, indFinal, Left(NFe!NFeIndicadorPresencaComprador, 1), Left$("ONLINE COMMERCE - v." & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision), 20), dhContingencia, justContingencia, 0, 0, "", "", mensagemAlerta, mensagemErro)
    
    '================grupo de identificação do destinatario (grupo E do Manual de integração - páginas 92)=======================
    Dim xRazaoSocial As String, xCNPJ As String, xCPF As String, xTelefone As String
    'CLIENTE COM CPF
    If Len(NFe!CPF_CNPJ) = 18 Then
      xCNPJ = Trim(NFe!CPF_CNPJ)                                   ' CNPJ do destinatario sem máscara de formatação
      xCPF = ""
    Else
      xCPF = Trim(NFe!CPF_CNPJ)                                    ' CPF do destinatario, uso exclusivo do Fisco 'aqui
      xCNPJ = ""
    End If
    xRazaoSocial = RemoveAcento(NFe!NomeRazSocial)
    
    If Left(Parametros!AmbienteNF, 1) = 2 Then
       indFinal = "1"
       xRazaoSocial = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
       NFe!InscEst = ""
       NFe!NFeIndicadorIEDestinatario = "9"
    End If
    
    If Len(NFe!CPF_CNPJ) > 0 Then
       iRetorno = sistNFCe.GerarDestinatario(4, xRazaoSocial, xCNPJ, xCPF, "", Retira(NFe!InscEst, ".,-/", UM_A_UM), "", Left$(NFe!NFeIndicadorIEDestinatario, 1), "", RemoveAcento(NFe!Endereco), NFe!Num, "", RemoveAcento(NFe!bairro), NFe!CodigoIBGE, RemoveAcento(NFe!Municipio), NFe!UF, Retira(NFe!CEP, ".- ", UM_A_UM), 1058, "BRASIL", NFe!Fone, "", mensagemAlerta, mensagemErro)
    End If
    'desabilitei pq o campo tipo_produto tá dando erro na consulta
    vsSQL = "SELECT TbNFCe_Itens.IdNFProd, TbNFCe_Itens.IdNFProd_Item, TbNFCe_Itens.CodProduto, TbNFCe_Itens.IdProduto, TbNFCe_Itens.DescricaoProduto, TbNFCe_Itens.ValorOutras, TbNFCe_Itens.TipoProduto, " & _
            "TbNFCe_Itens.CodBarras, TbNFCe_Itens.UN, TbNFCe_Itens.CFOP, TbNFCe_Itens.QtdeMov, TbNFCe_Itens.ValorUnit, TbNFCe_Itens.Desconto, TbNFCe_Itens.Aliq_Icms AS Aliquota, " & _
            "TbNFCe_Itens.Bc_Icms, TbNFCe_Itens.Bc_AliquotaReducao, TbNFCe_Itens.Vlr_Icms, TbNFCe_Itens.Aliq_IPI As AliqIPI, TbNFCe_Itens.Vlr_IPI As ValorIPI, TbNFCe_Itens.Valor_Frete As ValorFrete, " & _
            "TbNFCe_Itens.ICMSCST, TbNFCe_Itens.PISCST, TbNFCe_Itens.COFINSCST, TbNFCe_Itens.IPICST, TbNFCe_Itens.Codncm As NCM, TbNFCe_Itens.ProdInfAdicional, TbNFCe_Itens.ValorTributos, " & _
            "TbNFCe_Itens.BCSTRet, TbNFCe_Itens.ICMSSTRet, TbNFCe_Itens.BCImpostoImportacao, TbNFCe_Itens.DespesasAduaneiras, TbNFCe_Itens.ValorImpostoImportacao, TbNFCe_Itens.ValorIOF, TbNFCe_Itens.Aliq_PIS, TbNFCe_Itens.Aliq_COFINS, TbNFCe_Itens.vlr_COFINS, TbNFCe_Itens.vlr_PIS " & _
            "FROM TbNFCe_Itens " & _
            "WHERE TbNFCe_Itens.IdNFProd = " & NumeroNota & " " & _
            "ORDER BY TbNFCe_Itens.IdNFProd_Item"
        
        'desabilitei para ver sobre o campo que estava dando erro... sql abaixo tá funcionando para emissão todas menos gás
        'vsSQL = "SELECT TbNFCe_Itens.IdNFProd, TbNFCe_Itens.IdNFProd_Item, TbNFCe_Itens.CodProduto, TbNFCe_Itens.IdProduto, TbNFCe_Itens.DescricaoProduto, TbNFCe_Itens.ValorOutras, " & _
            "TbNFCe_Itens.CodBarras, TbNFCe_Itens.UN, TbNFCe_Itens.CFOP, TbNFCe_Itens.QtdeMov, TbNFCe_Itens.ValorUnit, TbNFCe_Itens.Desconto, TbNFCe_Itens.Aliq_Icms AS Aliquota, " & _
            "TbNFCe_Itens.Bc_Icms, TbNFCe_Itens.Bc_AliquotaReducao, TbNFCe_Itens.Vlr_Icms, TbNFCe_Itens.Aliq_IPI As AliqIPI, TbNFCe_Itens.Vlr_IPI As ValorIPI, TbNFCe_Itens.Valor_Frete As ValorFrete, " & _
            "TbNFCe_Itens.ICMSCST, TbNFCe_Itens.PISCST, TbNFCe_Itens.COFINSCST, TbNFCe_Itens.IPICST, TbNFCe_Itens.Codncm As NCM, TbNFCe_Itens.ProdInfAdicional, TbNFCe_Itens.ValorTributos, " & _
            "TbNFCe_Itens.BCSTRet, TbNFCe_Itens.ICMSSTRet, TbNFCe_Itens.BCImpostoImportacao, TbNFCe_Itens.DespesasAduaneiras, TbNFCe_Itens.ValorImpostoImportacao, TbNFCe_Itens.ValorIOF, TbNFCe_Itens.Aliq_PIS, TbNFCe_Itens.Aliq_COFINS, TbNFCe_Itens.vlr_COFINS, TbNFCe_Itens.vlr_PIS " & _
            "FROM TbNFCe_Itens " & _
            "WHERE TbNFCe_Itens.IdNFProd = " & NumeroNota & " " & _
            "ORDER BY TbNFCe_Itens.IdNFProd_Item"
            'Debug.Print vsSQL

    RsOpen NFeItens, vsSQL
    'Set NFeItens = vgDb.OpenRecordset(vsSQL)

    'parte do gás que estava desabilitada por erro
    Dim vGasCounter As Integer
    vGasCounter = 0
    If NFeItens!TipoProduto = "Combustível" Then    'coloquei pq Lider tava dando erro ao localizar isso, sem precisar
        vsSQL = "SELECT Cod_Produto, CODIF, cProdANP, descricaoANP, pGLP, pGNi, pGNn, pMixGN, ValorPartida " & _
                "FROM Produtos_Gas " & _
                "WHERE Cod_Produto = " & NFeItens!IDProduto
                'Debug.Print vsSQL
        ''vsSQL = "SELECT Cod_Produto, CODIF, cProdANP, descricaoANP, pGLP, pGNi, pGNn, pMixGN, ValorPartida " & _
                "FROM Produtos_Gas " & _
                "WHERE Cod_Produto = " & NFeItens!CodigoProduto   'desabiLITEI NO DIA DO GAS DA GLECIA
        RsOpen NFeCombustivel, vsSQL
        vGasCounter = 1
    End If
    
    
    n = NFeItens.RecordCount

    If NFe!Valor_NF_Prod > 0 Then
       pFrete = Format((NFe!Valor_Frete / NFe!Valor_NF_Prod) * 100, "######0.000000")
       pOutras = Format((NFe!OutrasDespesasAces / NFe!Valor_NF_Prod) * 100, "######0.000000")
    End If

    For i = 1 To NFeItens.RecordCount
        '================grupo de detalhe do produto (grupo I01 do Manual de integração - páginas 95)=======================
        Dim infAdProd As String, pAliqSN As Double, vCredSN As Double, vlCredICMSSN As Double
        If NFeItens!ValorTributos > 0 Then
           infAdProd = RemoveAcento(Trim(Left(NFeItens!ProdInfAdicional, 400))) & " - Valor Aproximado dos Tributos R$ " & Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)        ' informações adicionais do produto
           vlTrib = vlTrib + NFeItens!ValorTributos
        Else
           infAdProd = RemoveAcento(Trim(Left(NFeItens!ProdInfAdicional, 500)))     ' informações adicionais do produto
        End If
        
        'desativei pq nao encontrei o campo TipoProduto na tabela de NFCeItens
        'If NFeCombustivel.RecordCount > 0 Then infAdProd = "ICMS monofásico sobre combustíveis cobrado anteriormente conforme Convênio ICMS 199/2022. " + infAdProd
       
        infAdProd = Trim(infAdProd)
       
        iRetorno = sistNFCe.GerarItens(i, Trim$(NFeItens!IDProduto), RemoveAcento(NFeItens!DescricaoProduto), NFeItens!NCM, "", "", NFeItens!CodBarras, NFeItens!CodBarras, _
                                       NFeItens!CFOP, NFeItens!QtdeMov, NFeItens!ValorUnit, NFeItens!UN, NFeItens!QtdeMov, NFeItens!ValorUnit, NFeItens!UN, (NFeItens!QtdeMov * NFeItens!ValorUnit), NFeItens!ValorFrete, NFeItens!Desconto, NFeItens!ValorOutras, 0, "", "", 0, "", "", "", "", "", IIf(NFeItens!CFOP = 1603, 0, 1), infAdProd, 0, "", 0, mensagemAlerta, mensagemErro)
        
        '=========dados do ICMS (grupo N01 do Manual de integração - páginas 100)=====================
        'Parametros!
        If Left(Parametros!CRT, 1) = 1 Then
           If (Right(NFeItens!ICMSCST, 3) = "101" Or Right(NFeItens!ICMSCST, 3) = "201") Then
              'pAliqSN = Format(NFeItens!pCreditoICMSSimplesNacional, "#0.00")      ' <pCredSN> Simples Nacional "DESATIVEI para teste no parametro"
              pAliqSN = Format(Parametros!pCreditoICMSSimplesNacional, "#0.00")      ' <pCredSN>          Simples Nacional
              vCredSN = Format(((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!pCreditoICMSSimplesNacional / 100)), "#0.00") ' <vCredICMSSN>      Simples Nacional
              vlCredICMSSN = vlCredICMSSN + Round(((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!pCreditoICMSSimplesNacional / 100)), 2)
           Else
              pAliqSN = "0.00"                                          ' <pCredSN>          Simples Nacional
              vCredSN = "0.00"                                          ' <vCredICMSSN>      Simples Nacional
           End If
        End If
        
        'desativei esse IF pq nao encontrei o campo TipoProduto na tabela de NFCeItens
        If vGasCounter = 1 Then
            If NFeCombustivel.RecordCount = 0 Then
               iRetorno = sistNFCe.GerarItensImpostoEstadual(NFeItens!ValorTributos, Left(NFeItens!ICMSCST, 1), Right(NFeItens!ICMSCST, IIf(Left(Parametros!CRT, 1) = 1, 3, 2)), 3, NFeItens!Bc_Icms, NFeItens!Aliquota, NFeItens!Vlr_Icms, NFeItens!Bc_AliquotaReducao, _
                                                             0, 0, 0, 5, 0, 0, NFeItens!BCSTRet, 0, NFeItens!ICMSSTRet, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                                             pAliqSN, vCredSN, 0, 0, 0, 0, mensagemAlerta, mensagemErro)
            Else
               iRetorno = sistNFCe.GerarCombustivel(NFeCombustivel!CODIF, NFeCombustivel!cProdANP, NFeCombustivel!descricaoANP, NFeCombustivel!pGLP, NFeCombustivel!pGNi, NFeCombustivel!pGNn, NFeCombustivel!pMixGN, 0, Parametros!Estado, NFeCombustivel!ValorPartida, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, mensagemAlerta, mensagemErro)
               
               iRetorno = sistNFCe.GerarItensImpostoEstadualMonofasico(NFeItens!ValorTributos, "0", "61", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, mensagemAlerta, mensagemErro)
               ''iRetorno = sistNFCe.GerarItensObservacao("CST61", "ICMS monofásico sobre combustíveis cobrado anteriormente conforme Convênio ICMS 199/2022;", "", "", mensagemAlerta, mensagemErro)
               infAdProd = "ICMS monofásico sobre combustíveis cobrado anteriormente conforme Convênio ICMS 199/2022. " + infAdProd
            End If
        Else
            iRetorno = sistNFCe.GerarItensImpostoEstadual(NFeItens!ValorTributos, Left(NFeItens!ICMSCST, 1), Right(NFeItens!ICMSCST, IIf(Left(Parametros!CRT, 1) = 1, 3, 2)), 3, NFeItens!Bc_Icms, NFeItens!Aliquota, NFeItens!Vlr_Icms, NFeItens!Bc_AliquotaReducao, _
                                                             0, 0, 0, 5, 0, 0, NFeItens!BCSTRet, 0, NFeItens!ICMSSTRet, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                                             pAliqSN, vCredSN, 0, 0, 0, 0, mensagemAlerta, mensagemErro)
        End If

        'pis e cofins - Impostos federais
        Dim COFINSCST As String, PISCST As String
        
        COFINSCST = IIf(Vazio(NFeItens!COFINSCST) Or IsNull(NFeItens!COFINSCST), "07", NFeItens!COFINSCST)
        Select Case COFINSCST
           Case "04", "06", "07", "08", "09"
              vlCOFINS = vlCOFINS
           Case Else
              vlCOFINS = vlCOFINS + (NFeItens!QtdeMov * NFeItens!ValorUnit) * (NFeItens!Aliq_COFINS / 100)
            ' vlCOFINS = vlCOFINS + Round((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!COFINSAliquota / 100), 2)
        End Select
        
        PISCST = IIf(Vazio(NFeItens!PISCST) Or IsNull(NFeItens!PISCST), "07", NFeItens!PISCST)
        Select Case PISCST
           Case "04", "06", "07", "08", "09"
              vlPIS = vlPIS
           Case Else
              vlPIS = vlPIS + (NFeItens!QtdeMov * NFeItens!ValorUnit) * (NFeItens!Aliq_PIS / 100)
             'vlPIS = vlPIS + Round((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!PISAliquota / 100), 2)
        End Select
                                                    'pisCST As String, pisvBC As Double, pPIS As Double, vPIS As Double

        iRetorno = sistNFCe.GerarItensImpostoFederal(COFINSCST, NFeItens!Bc_Icms, NFeItens!Aliq_COFINS, NFeItens!vlr_COFINS, 0, 0, _
                                                     PISCST, NFeItens!Bc_Icms, NFeItens!Aliq_PIS, NFeItens!vlr_PIS, 0, 0, _
                                                     NFeItens!IPICST, (NFeItens!QtdeMov * NFeItens!ValorUnit), NFeItens!AliqIPI, NFeItens!ValorIPI, 0, 0, "999", "", "", "", 0, mensagemAlerta, mensagemErro)
        'iRetorno = sistNFCe.GerarItensImpostoFederal(cofinsCST, (NFeItens!QtdeMov * NFeItens!ValorUnit), Parametros!COFINSAliquota, Round((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!COFINSAliquota / 100), 2), 0, 0, _
                                                     pisCST, (NFeItens!QtdeMov * NFeItens!ValorUnit), Parametros!PISAliquota, Round((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!PISAliquota / 100), 2), 0, 0, _
                                                     NFeItens!IPICST, (NFeItens!QtdeMov * NFeItens!ValorUnit), NFeItens!AliqIPI, NFeItens!ValorIPI, 0, 0, "999", "", "", "", 0, mensagemAlerta, mensagemErro)

        iRetorno = sistNFCe.GerarItensIncluir(mensagemAlerta, mensagemErro)
        
        NFeItens.MoveNext
    Next
    
    ' atualização de total
     'vsSQL = "SELECT ISNULL(SUM(Bc_Icms), 0) AS vValorBC, ISNULL(SUM(Vlr_Icms), 0) AS vValorTotalICMS FROM TbNFCe_Itens WHERE (Aliq_Icms <> '0.00') and IdNFProd = " & NumeroNota
    'RsOpen Totais, vsSQL
    
    vlNF = (NFe!Valor_NF_Prod - NFe!DescontoPromocional) + (NFe!Valor_Frete + NFe!Valor_Seguro + NFe!Valor_IPI + NFe!OutrasDespesasAces + NFe!ValorImpostoImportacao + NFe!Valor_ICMS_Subst)
    iRetorno = sistNFCe.GerarTotalProdutos(NFe!BaseCalc_ICMS, NFe!Valor_ICMS, NFe!BaseCalc_ICSM_Subst, NFe!Valor_ICMS_Subst, vlCOFINS, vlPIS, NFe!Valor_IPI, NFe!DescontoPromocional, NFe!Valor_Seguro, NFe!Valor_Frete, NFe!OutrasDespesasAces, 0, 0, 0, 0, 0, 0, 0, NFe!ValorImpostoImportacao, 0, NFe!Valor_NF_Prod, vlNF, vlTrib, _
                                           0, 0, 0, 0, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, mensagemAlerta, mensagemErro)
    
    'Dim vBCICMS As Currency
    'Dim vVLRICMS As Currency
    'vBCICMS = Totais!vValorBC
    'vVLRICMS = Totais!vValorTotalICMS
    
    'vsSQL = "UPDATE TbNFCe SET " & _
         "BaseCalc_ICMS = " & Replace(CDbl(vBCICMS), ",", ".") & ", " & _
         "Valor_ICMS = " & Replace(CDbl(vVLRICMS), ",", ".") & " " & _
         "WHERE IdNFProd = " & NumeroNota
    'vgDb.Execute vsSQL
    
    '============dados do transportador
    If Len(Retira(NFe!CPF_CNPJ_Transp, ".-/", UM_A_UM)) > 11 Then
      xCNPJ = Trim(NFe!CPF_CNPJ_Transp)                                         ' CNPJ da Transportadora sem mascara
      xCPF = ""
    Else
      xCPF = Trim(NFe!CPF_CNPJ_Transp)                                          ' CPF da Transportadora sem mascara
      xCNPJ = ""
    End If
    
    iRetorno = sistNFCe.GerarTransporte(IIf(Vazio(NFe!Frete_Por_Conta), 0, Left(NFe!Frete_Por_Conta, 1)), NFe!Qtde_Trasnp, RemoveAcento(NFe!Especie_Transp), RemoveAcento(NFe!Marca_Trasnp), NFe!Num_Transp, NFe!PesoBruto_Transp, NFe!PesoLiq_Transp, _
                                        xCNPJ, xCPF, Retira(NFe!InscEst_Trasnp, ".,-/", UM_A_UM), RemoveAcento(NFe!NomeTrasnportador), RemoveAcento(Trim(NFe!Endereco_Transp)), RemoveAcento(Trim(NFe!Cidade_Transp)), NFe!UF_Mot_Transp, _
                                        Retira(Trim(NFe!Placa_Veiculo), "-", UM_A_UM), "", NFe!UF_Trasnportador, mensagemAlerta, mensagemErro)
    'parcelas
    vsSQL = "SELECT COUNT(IDParcela) as qt FROM TbNFCe_Faturas WHERE IdNFProd = " & NumeroNota & " GROUP BY IDParcela, TipoPgto, IdBandeira"

    Dim idparc As Integer, vTotalRecebido As Double, vTotalNF As Double, vTotalDinheiro As Double, vTotalOutras As Double, vTroco As Double
    idparc = 0
    If SQLExecutaRetorno(vsSQL, "qt", 0) >= 0 Then
       'iRetorno = sistNFCe.GerarCobranca(NFe!NumeNota, 0, vlNF, vlNF, mensagemAlerta, mensagemErro)
       vsSQL = "SELECT IDParcela, SUM(Valor) AS Valor, TipoPgto, IdBandeira " & _
               "FROM TbNFCe_Faturas " & _
               "WHERE idNFProd = " & NumeroNota & " " & _
               "GROUP BY IDParcela, TipoPgto, IdBandeira"
               
        'Debug.Print vsSQL
       'vTotalRecebido = SQLExecutaRetorno("SELECT recebido r FROM pedidos WHERE cod_pedido = " & NFe!Num_OS_VD_Origem, "r", 0)
       'vTotalDinheiro = SQLExecutaRetorno("SELECT SUM(Valor) r FROM TbNFCe_Faturas WHERE IdNFProd = " & NumeroNota & " AND TipoPgto = 'DH'", "r", 0)
       vTroco = SQLExecutaRetorno("SELECT troco r FROM pedidos WHERE cod_pedido = " & NFe!Num_OS_VD_Origem, "r", 0)
       RsOpen NFeParcelas, vsSQL
       'Set NFeParcelas = vgDb.OpenRecordset(vsSQL)
       vTotalNF = vlNF
       Do While Not NFeParcelas.EOF
          '01 - Dinheiro|02 - Cheque|03 - Cartão de Crédito|04 - Cartão de Débito|05 - Crédito Loja|10 - Vale Alimentação|11 - Vale Refeição|12 - Vale Presente|13 - Vale Combustível| 17 - PIX - 99 - Outros
          Select Case NFeParcelas!TipoPgto
                 Case "DH": iRetorno = sistNFCe.GerarPagamentos(0, 1, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)               'pag <tPag>
                 Case "CH": iRetorno = sistNFCe.GerarPagamentos(1, 2, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
                 Case "CC": iRetorno = sistNFCe.GerarPagamentos(0, 3, NFeParcelas!Valor, 0, NFeParcelas!IdBandeira, 2, "", Retira(Parametros!CNPJ, ".-/ ", UM_A_UM), "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
                 Case "CD": iRetorno = sistNFCe.GerarPagamentos(0, 4, NFeParcelas!Valor, 0, NFeParcelas!IdBandeira, 2, "", Retira(Parametros!CNPJ, ".-/ ", UM_A_UM), "", "", "", "", "", "", mensagemAlerta, mensagemErro)                     'pag <tPag>
                 Case "PM": iRetorno = sistNFCe.GerarPagamentos(1, 5, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
                 Case "FI": iRetorno = sistNFCe.GerarPagamentos(1, 14, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
                 Case "BL": iRetorno = sistNFCe.GerarPagamentos(1, 15, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
                 Case "DP": iRetorno = sistNFCe.GerarPagamentos(0, 16, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
                 Case "PX": iRetorno = sistNFCe.GerarPagamentos(0, 20, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
                 Case "TR": iRetorno = sistNFCe.GerarPagamentos(0, 18, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
                 Case Else: iRetorno = sistNFCe.GerarPagamentos(0, 99, NFeParcelas!Valor, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
          End Select
          idparc = idparc + 1
          NFeParcelas.MoveNext
       Loop
       If NFeParcelas.RecordCount = 0 Then
          iRetorno = sistNFCe.GerarPagamentos(0, 1, vlNF, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
       End If
    Else
       iRetorno = sistNFCe.GerarPagamentos(0, 1, vlNF, 0, 0, 0, "", "", "", "", "", "", "", "", mensagemAlerta, mensagemErro)                    'pag <tPag>
    End If

    '============= informações adcionais
    Dim obsAdic As String, obsCpl As String
    obsAdic = ""   'RemoveAcento(Trim(NFe!InformacoesAdicionais))
    obsCpl = RemoveAcento(Trim(NFe!Linha1))
    obsCpl = obsCpl & IIf(Not Vazio(NFe!Linha2), " " & RemoveAcento(Trim(NFe!Linha2)), "")
    obsCpl = obsCpl & IIf(Not Vazio(NFe!Linha3), " " & RemoveAcento(Trim(NFe!Linha3)), "")
    obsCpl = obsCpl & IIf(Not Vazio(NFe!Linha4), " " & RemoveAcento(Trim(NFe!Linha4)), "")
    obsCpl = obsCpl & IIf(Not Vazio(NFe!Linha5), " " & RemoveAcento(Trim(NFe!Linha5)), "")
    If vlTrib > 0 And vlNF > 0 Then
       vlNF = (NFe!Valor_NF_Prod - NFe!DescontoPromocional) + (NFe!Valor_Frete + NFe!Valor_Seguro + NFe!Valor_IPI + NFe!OutrasDespesasAces + NFe!ValorImpostoImportacao)
       pTributos = Format((vlTrib / vlNF) * 100, "#0.00")
       obsCpl = obsCpl & " - Valor Aproximado dos Tributos R$ " & FormatoDecimal(Format(vlTrib, "#0.00")) & " (" & FormatoDecimal(pTributos) & "%) (Conforme Lei Fed. 12.741/2012) Fonte: IBPT"
    End If
    iRetorno = sistNFCe.GerarInformacoesAdicionais(obsCpl, obsAdic, mensagemAlerta, mensagemErro)
Else
   MsgBox "Ocorreu um erro ao gerar a NFCe, verifique novamente os dados da NFCe!", vbCritical + vbOKOnly, "ERRO"
   GoTo Caifora
End If

'gera a chave da nfe
Dim id_chave As String
Dim numero_nfe_gerado As String
iRetorno = sistNFCe.GerarXML(numero_nfe_gerado, xCaminhoXML, False, xCaminhoXMLAuxiliar, mensagemAlerta, mensagemErro)
id_chave = numero_nfe_gerado
numero_nfe_gerado = Replace(numero_nfe_gerado, "NFe", "")
NFeChaveAcesso = numero_nfe_gerado

If Not Vazio(NFeChaveAcesso) Then
   vsSQL = "UPDATE TbNFCe SET " & _
           "NFCeChaveAcesso = '" & NFeChaveAcesso & "' " & _
           "WHERE IdNFProd = " & NumeroNota
   vgDb.Execute vsSQL
End If

If NFCeContingenciaOFF And Not PodeEnviar Then GoTo NaoEnviou
If Not PodeEnviar Then GoTo NaoEnviou

NFeResposta = ""

SoEnviar:

iRetorno = sistNFCe.EnviarNFe(NFe!NumeNota, 1, False)

NFeResposta = sistNFCe.retEnvio.protNFe.infProt.xMotivo

If Not iRetorno Then
   MsgBox "*** Aparentemente Ocorreram Erros na Recepção do Lote (nfeAutorizacao)***" & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeAutorizacao"
   GoTo Caifora
End If

NFeNumeroRecibo = ""
If Not iRetorno Then GoTo Caifora
cStat = sistNFCe.retEnvio.protNFe.infProt.cStat
NFeMotivo = sistNFCe.retEnvio.protNFe.infProt.xMotivo
If cStat = 103 Then NFeNumeroRecibo = sistNFCe.retEnvio.infRec.nRec  'Parse(NFeResposta, "#")
If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo Caifora

If cStat <> 103 Then
   If cStat = 104 Or cStat = 100 And Vazio(NFeNumeroRecibo) Then GoTo buscaNFe
   GoTo NaoEnviou
End If

vsSQL = "UPDATE TbNFCe SET " & _
        "NFCeRecibo = " & NFeNumeroRecibo & " " & _
        "WHERE IdNFProd = " & NumeroNota
vgDb.Execute vsSQL

DoEvents

consultaNFe:
   On Error Resume Next
   iRetorno = sistNFCe.ConsultarReciboDeEnvio(NFeNumeroRecibo)
   'On Error GoTo TransmitirNFCe_Error
   
   cStat = sistNFCe.retConsRec.cStat
   NFeMotivo = sistNFCe.retConsRec.xMotivo
   
   If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
      MsgBox NFeMotivo, vbExclamation, "Retorno Autorização"
      GoTo Caifora
   End If

   If cStat = 217 Then
      Sleep 3000 ' Aguarda mais 3 segundos
      On Error Resume Next
      iRetorno = sistNFCe.ConsultarReciboDeEnvio(NFeNumeroRecibo)
      On Error GoTo TransmitirNFCe_Error
      
      cStat = sistNFCe.retConsRec.cStat
      NFeMotivo = sistNFCe.retConsRec.xMotivo
      
      If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
         MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote (nfeRetAutorizacao)***" & vbLf & NFeMotivo, vbExclamation, "PROCESSO INTERROMPIDO - NfeRetAutorizacao"
         GoTo Caifora
      End If
   
   End If
   
   If cStat <> 105 Then
      cStat2 = sistNFCe.retConsRec.protNFe.infProt.cStat
      NFeValidate = sistNFCe.retConsRec.protNFe.infProt.xMotivo
   End If
 
   If InStr(NFeValidate, "Erro") > 0 Or InStr(NFeValidate, "Rejeicao") > 0 Or InStr(NFeValidate, "Rejeição") > 0 Then
      MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote (nfeRetAutorizacao)***" & vbLf & CStr(cStat2) & " - " & NFeValidate, vbExclamation, "PROCESSO INTERROMPIDO - NfeRetAutorizacao"
      NFeMotivo = CStr(cStat2) & " - " & NFeValidate
      GoTo Caifora
   End If

buscaNFe:

   'Consulta Nfe
   If Not IsNumeric(NFeChaveAcesso) Then GoTo Caifora
   
   iRetorno = sistNFCe.ConsultarProtocolo(NFeChaveAcesso)

   cStat = sistNFCe.retConsulta.cStat
   NFeMotivo = sistNFCe.retConsulta.xMotivo
   
   If cStat = 217 Then
      Sleep 3000 ' aguarda mais 3 segundos
      iRetorno = sistNFCe.ConsultarProtocolo(NFeChaveAcesso)
      
      cStat = sistNFCe.retConsulta.cStat
      NFeMotivo = sistNFCe.retConsulta.xMotivo

      If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
         MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote ***" & vbLf & NFeMotivo, vbExclamation, "PROCESSO INTERROMPIDO - NfeConsulta"
         GoTo Caifora
      End If
   End If

   If Not iRetorno Then
      NFeNumeroProtocolo = ""
      GoTo Caifora
   End If

   cStat = sistNFCe.retConsulta.cStat
   NFeMotivo = sistNFCe.retConsulta.xMotivo
   If cStat = 104 Or InStr(NFeMotivo, "Autorizado") > 0 Then
      NFeDataHora = sistNFCe.retConsulta.protNFe.infProt.ProxyDhRecbto
      NFeNumeroProtocolo = sistNFCe.retConsulta.protNFe.infProt.nProt
      cStat2 = sistNFCe.retConsulta.protNFe.infProt.cStat
   Else
      NFeDataHora = ""
      NFeNumeroProtocolo = ""
      cStat2 = 0
   End If
   
    If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo Caifora

    If cStat2 = 204 Or cStat2 = 539 Then
       NFeChaveAcesso = Mid(nfeRetorno, InStr(nfeRetorno, "chNFe:") + 6, 44)
       nroRecibo = Mid(nfeRetorno, InStr(nfeRetorno, "nRec:") + 5)
       nroRecibo = Left(nroRecibo, Len(nroRecibo) - 1)
       NFeNumeroRecibo = Left(NFeNumeroRecibo, 15 - Len(nroRecibo)) + nroRecibo
       If Vazio(NFeChaveAcesso) Or Len(NFeNumeroRecibo) < 15 Then
          NFeMotivo = nfeRetorno
          GoTo Caifora
       End If
       If NFeNumeroRecibo <> "" Then
          vsSQL = "UPDATE TbNFCe SET " & _
                  "NFCeChaveAcesso = '" & NFeChaveAcesso & "' " & _
                  "NFCeRecibo = " & NFeNumeroRecibo & " " & _
                  "WHERE IdNFProd = " & NumeroNota
          vgDb.Execute vsSQL
       End If
       GoTo buscaNFe
    ElseIf cStat2 <> 100 And cStat2 <> 301 Then
        NFeMotivo = nroRecibo + " - " + nfeRetorno
        GoTo NaoEnviou
    ElseIf cStat2 = 105 And cStat2 = 217 Then
        GoTo consultaNFe
    ElseIf cStat2 = 100 Then
        nfeRetorno = "Nota Fiscal de Consumidor Eletronica Autorizado o Uso."
        NFeDataHora = Format(Now, "dd/mm/yyyy h:mm:ss")
        msgResultado = "Chave NF-e.: " + NFeChaveAcesso & vbCrLf
        msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
        If NFeNumeroRecibo <> "" Then msgResultado = msgResultado + "Recibo.: " + NFeNumeroRecibo & vbCrLf
        msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
        msgResultado = msgResultado + "Resposta da Fazenda.: " + str$(cStat2) + " - " & nfeRetorno

        ' mensagem de emissao  MsgBox msgResultado, vbInformation + vbOKOnly

        On Error Resume Next
        NFeResposta = sistNFCe.ConsultarProtocolo(NFeChaveAcesso)

        vsSQL = "UPDATE TbNFCe SET " & _
                "NFCeEnviada = 1, " & _
                "NFCeChaveAcesso = '" & NFeChaveAcesso & "', " & _
                "NFCeProtocolo = " & NFeNumeroProtocolo & ", " & _
                "NFCeProtocoloDataHora = '" & NFeDataHora & "' " & _
                "WHERE IdNFProd = " & NumeroNota
        vgDb.Execute vsSQL
    End If

    'xCaminhoXML = dirXML & "\arquivos\procNFe\" & NFeChaveAcesso & "-procNFe.xml"
    
    Dim xmlPathPDF As String
    Dim anoEmes As String
    Dim Arquivo As String
    xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & Format(NFe!DataEmissao, "yyyymm") & "\" & NFeChaveAcesso & "-procNFe.xml"
    xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
    If Not Existe(xCaminhoXML) Then xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & NFeChaveAcesso & "-procNFe.xml"             '  Aqui Gera o DANFE

PodeSair:
Set sistNFCe = Nothing
Screen.MousePointer = vbDefault
TransmitirNFCe = True
Exit Function

Resume

NaoEnviou:
Set sistNFCe = Nothing

If Not NFCeContingenciaOFF And PodeEnviar Then MsgBox NFeMotivo, vbCritical + vbOKOnly
If NFCeContingenciaOFF And Not PodeEnviar Then
   TransmitirNFCe = True
   xCaminhoXML = xCaminhoXMLAuxiliar
   xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
End If
Screen.MousePointer = vbDefault
Exit Function
Resume

Caifora:
    If Not Vazio(NFeMotivo) Then MsgBox NFeMotivo, vbCritical + vbOKOnly
    If Not Vazio(NFeResposta) Then MsgBox NFeResposta, vbCritical + vbOKOnly
    
    Set sistNFCe = Nothing
    
    Screen.MousePointer = vbDefault
    TransmitirNFCe = False
    
    On Error GoTo 0
    Exit Function

Resume

deuErro:
'    If sistNFCe.certificadoVencimento <= Date Then
'
'    Else
'       'MsgBox Err.Description, vbCritical + vbOKOnly
'    End If

    If InStr(1, sistNFCe.xMotivo, "Erros na validação") > 0 Then
       MsgBox TrataErroValidacao(sistNFCe.xMotivo), vbExclamation + vbOKOnly, "ERRO VALIDAÇÃO XML"
    ElseIf Not Vazio(sistNFCe.xMotivo) Then
       MsgBox sistNFCe.xMotivo, vbExclamation + vbOKOnly, "ERRO"
    End If
    
    Set sistNFCe = Nothing
    
    Screen.MousePointer = vbDefault
    TransmitirNFCe = False
    
    Err.Clear
    Exit Function
   

TransmitirNFCe_Error:    'reativei dia 06/03/26 para ver
    'Screen.MousePointer = vbDefault
    ''MsgBox "ERRO AO TRANSMITIR A NFCe." & vbNewLine & "Confira os produtos e tente transferir novamente!", vbCritical, "Falha"
    ''MsgBox "Falha (" & Err.Description & ")" & vbNewLine & "Em TransmitirNFCe no Módulo NFe_DLL", vbCritical, "Falha"
    'MsgBox CStr(sistNFCe.cStat) & " - " & sistNFCe.xMotivo, vbCritical, "Falha"
    'Set sistNFCe = Nothing
    'Err.Clear
End Function

Public Function CancelaNFe(ChaveAcesso As Variant, Protocolo As Variant, Justificativa As Variant, GravaProtocolo As Boolean) As Boolean  'Função para envio do cancelamento da NFe
Dim IdLote As Long, dhEvento As String, CNPJ As String
Dim sistNFe As snfe.Util
   
   Screen.MousePointer = vbHourglass
   Set sistNFe = New snfe.Util
   
   iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)

   IdLote = Int(Rnd(918274 * 999) * 1000)
   dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
   CNPJ = SQLExecutaRetorno("SELECT CNPJ FROM Empresa", "CNPJ", "")
   
   iRetorno = sistNFe.CancelarNFe(CNPJ, IdLote, 1, ChaveAcesso, Protocolo, Justificativa, xCaminhoXML)
 
   If Not iRetorno Then GoTo Caifora
   
   cStat = sistNFe.retEnvEvento.cStat        'Parse(NFeResposta, "#")
   NFeMotivo = sistNFe.retEnvEvento.xMotivo  'Parse(NFeResposta, "#")
   If cStat = 128 Then
      cStat2 = sistNFe.retEnvEvento.retEvento.infEvento.cStat              'Parse(NFeResposta, "#")
      NFeValidate = sistNFe.retEnvEvento.retEvento.infEvento.xMotivo       'Parse(NFeResposta, "#")
      NFeNumeroProtocolo = sistNFe.retEnvEvento.retEvento.infEvento.nProt  'Parse(NFeResposta, "#")
      NFeDataHora = sistNFe.retEnvEvento.retEvento.infEvento.dhRegEvento   'Parse(NFeResposta, "#")
   Else
      cStat2 = 0
      NFeValidate = ""
      NFeNumeroProtocolo = ""
      NFeDataHora = ""
   End If
   
   If cStat2 = 135 Or cStat2 = 155 Then
      GoTo continua
   Else
      If cStat2 > 0 Then
         MsgBox str$(cStat2) & " - " & NFeValidate, vbInformation, "Cancelar NFe"
      Else
         MsgBox str$(cStat) & " - " & NFeMotivo, vbInformation, "Cancelar NFe"
      End If
      GoTo Caifora
   End If
   
continua:
   msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + str(cStat2) & " - " & NFeValidate
   
   MsgBox msgResultado, vbInformation + vbOKOnly, "Cancelar NFe"

   If GravaProtocolo Then
      vsSQL = "INSERT INTO NotaFiscalRecibos (CodigoNota, NumeroProtocolo, DataHora) Values " & _
              "(" & vsNumeroNota & ", " & NFeNumeroProtocolo & ", '" & NFeDataHora & "')"
      vgDb.Execute vsSQL, True
   End If

   Screen.MousePointer = vbDefault
   CancelaNFe = True
   Exit Function

Caifora:
   Set sistNFe = Nothing
   Screen.MousePointer = vbDefault
   CancelaNFe = False
End Function

Public Function CancelaNFCe(ChaveAcesso As Variant, Protocolo As Variant, Justificativa As Variant, GravaProtocolo As Boolean) As Boolean  'Função para envio do cancelamento da NFe
Dim IdLote As Long, dhEvento As String, CNPJ As String
Dim sistNFCe As snfe.Util
   
   Screen.MousePointer = vbHourglass
   
   Set sistNFCe = New snfe.Util
   
   iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFCe)

   If Vazio(UTC) Then UTC = "-03:00"
   IdLote = Int(Rnd(918274 * 999) * 1000)
   dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
   CNPJ = SQLExecutaRetorno("SELECT CNPJ FROM Empresa", "CNPJ", "")
   
   iRetorno = sistNFCe.CancelarNFe(CNPJ, IdLote, 1, ChaveAcesso, Protocolo, Justificativa, xCaminhoXML)
 
   If Not iRetorno Then GoTo Caifora
   
   cStat = sistNFCe.retEnvEvento.cStat        'Parse(NFeResposta, "#")
   NFeMotivo = sistNFCe.retEnvEvento.xMotivo  'Parse(NFeResposta, "#")
   If cStat = 128 Then
      cStat2 = sistNFCe.retEnvEvento.retEvento.infEvento.cStat              'Parse(NFeResposta, "#")
      NFeValidate = sistNFCe.retEnvEvento.retEvento.infEvento.xMotivo       'Parse(NFeResposta, "#")
      NFeNumeroProtocolo = sistNFCe.retEnvEvento.retEvento.infEvento.nProt  'Parse(NFeResposta, "#")
      NFeDataHora = sistNFCe.retEnvEvento.retEvento.infEvento.dhRegEvento   'Parse(NFeResposta, "#")
   Else
      cStat2 = 0
      NFeValidate = ""
      NFeNumeroProtocolo = ""
      NFeDataHora = ""
   End If
   
   If cStat2 = 135 Or cStat2 = 155 Then
      GoTo continua
   Else
      MsgBox CStr(cStat2) & " - " & NFeMotivo, vbInformation
     GoTo Caifora
   End If
   
continua:
   msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + str(cStat2) & " - " & NFeValidate
   
   'msgResultado = NFeResposta
   
   MsgBox msgResultado, vbInformation + vbOKOnly

   Screen.MousePointer = vbDefault
   CancelaNFCe = True
   Exit Function

Caifora:
   Set sistNFCe = Nothing
   Screen.MousePointer = vbDefault
   CancelaNFCe = False
End Function

Public Sub ConsultaRecibo(Recibo As Variant, ChaveAcesso As String, Optional FormatoEmissao As String, Optional MostraMsg As Boolean = True)
Dim i As Integer
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

  Screen.MousePointer = vbHourglass
   
  iRetorno = ConfiguraDLLNFeNFCe(55, FormatoEmissao, sistNFe)
  
  i = 0
  
buscaNFe:
  iRetorno = sistNFe.ConsultarReciboDeEnvio(Recibo)
   
  If Not iRetorno Then
     NFeValidate = "ERRO"
     NFeNumeroProtocolo = ""
     GoTo Caifora
  End If
   
  cStat = sistNFe.retConsRec.cStat
  NFeMotivo = sistNFe.retConsRec.xMotivo
  
  If cStat = 105 Or cStat = 217 Then
     i = i + 1
     If i > 5 And cStat = 105 Then
        msgResultado = "A Nota Fiscal ficou em PROCESSAMENTO." & vbNewLine & "Efetue a consulta do Recibo daqui 5 minutos novamente."
        NFeValidate = "NFe/NFCe PROCESSAMENTO"
        GoTo Caifora
     End If
     Sleep 10000
     GoTo buscaNFe
  ElseIf cStat = 106 Then
     msgResultado = str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe NÃO LOCALIZADA"
     iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
    
     If Not iRetorno Then
        NFeValidate = "ERRO"
        NFeNumeroProtocolo = ""
        GoTo Caifora
     End If
     cStat = sistNFe.retConsulta.cStat
     NFeMotivo = sistNFe.retConsulta.xMotivo
     If cStat = 613 Then
        NFeChaveAcesso = Mid(NFeMotivo, InStr(NFeMotivo, "Numerico da NF-e [") + 18, 44)
        iRetorno = sistNFe.ConsultarProtocolo(NFeChaveAcesso)
        cStat = sistNFe.retConsulta.cStat
        NFeMotivo = sistNFe.retConsulta.xMotivo
        
        If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
           NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
           NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
           ChaveAcesso = sistNFe.retConsulta.protNFe.infProt.chNFe
           NFeChaveAcesso = ChaveAcesso
           GoTo continuaConsulta
        Else
           GoTo Caifora
        End If
     ElseIf cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
           nroRecibo = sistNFe.retConsulta.cStat
           NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
           NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
           ChaveAcesso = sistNFe.retConsulta.protNFe.infProt.chNFe
           NFeChaveAcesso = ChaveAcesso
           GoTo continuaConsulta
     Else
        GoTo Caifora
     End If
  ElseIf cStat = 239 Then
     msgResultado = str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe COM PROBLEMAS"
     GoTo Caifora
  ElseIf cStat = 215 Then
     msgResultado = str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe COM PROBLEMAS"
     GoTo Caifora
  End If
  
  NFeNumeroProtocolo = sistNFe.retConsRec.protNFe.infProt.nProt
  NFeDataHora = sistNFe.retConsRec.protNFe.infProt.ProxyDhRecbto
  nroRecibo = sistNFe.retConsRec.protNFe.infProt.cStat
  nfeRetorno = sistNFe.retConsRec.protNFe.infProt.xMotivo
  NFeValidate = sistNFe.retConsRec.protNFe.infProt.chNFe
  
  If ChaveAcesso = "" Then
     ChaveAcesso = NFeValidate
     NFeChaveAcesso = ChaveAcesso
  End If
  If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then iRetorno = 1

  If nroRecibo = 204 Or nroRecibo = 539 Then
     ChaveAcesso = Mid(nfeRetorno, InStr(nfeRetorno, "chNFe:") + 6, 44)
     nroRecibo = Mid(nfeRetorno, InStr(nfeRetorno, "nRec:") + 5)
     nroRecibo = Left(nroRecibo, Len(nroRecibo) - 1)
     Recibo = Left(Recibo, 15 - Len(nroRecibo)) + nroRecibo
     vsSQL = "UPDATE NotaFiscal SET " & _
             "ChavedeAcesso = '" & ChaveAcesso & "', " & _
             "NumeroRecibo = " & Recibo & " " & _
             "WHERE CodigoNota = " & vsNumeroNota
     vgDb.Execute vsSQL
     GoTo buscaNFe
  ElseIf nroRecibo = 206 Then
     NFeValidate = "NFe INUTILIZADA"
     ChaveAcesso = sistNFe.retConsulta.protNFe.infProt.chNFe
     msgResultado = "Chave NF-e.: " + ChaveAcesso & vbCrLf
     msgResultado = msgResultado + "Recibo.: " + Recibo & vbCrLf
     msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
     msgResultado = msgResultado + "Resposta da Fazenda.: " + str(nroRecibo) + " - " & nfeRetorno
     vsSQL = "UPDATE NotaFiscal SET " & _
             "ChavedeAcesso = '" & ChaveAcesso & "', " & _
             "Enviada = 1, " & _
             "Inutilizada = 1, " & _
             "DataHoraProcotolo = '" & NFeDataHora & "' " & _
             "WHERE CodigoNota = " & vsNumeroNota
     vgDb.Execute vsSQL
     GoTo Caifora
  ElseIf nroRecibo <> 100 And nroRecibo <> 301 Then
     msgResultado = nroRecibo + " - " + nfeRetorno
     NFeValidate = "ERRO"
     GoTo Caifora
  ElseIf nroRecibo = 105 Then
     GoTo buscaNFe
  End If

  iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
  
  If Not iRetorno Then
     NFeValidate = "ERRO"
     NFeNumeroProtocolo = ""
     GoTo Caifora
  End If

  cStat = sistNFe.retConsulta.cStat
  NFeMotivo = sistNFe.retConsulta.xMotivo
  If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
     NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
     NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
  Else
     NFeDataHora = ""
     NFeNumeroProtocolo = ""
  End If

continuaConsulta:
  msgResultado = "Chave NF-e.: " + ChaveAcesso & vbCrLf
  msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
  msgResultado = msgResultado + "Recibo.: " + Recibo & vbCrLf
  msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
  msgResultado = msgResultado + "Resposta da Fazenda.: " + str(cStat) + " - " & NFeMotivo
  
  If nroRecibo = 100 Then
     NFeValidate = "NFe AUTORIZADA"
     NFeChaveAcesso = ChaveAcesso
     vsSQL = "UPDATE NotaFiscal SET " & _
             "ChavedeAcesso = '" & ChaveAcesso & "', " & _
             "Enviada = 1, " & _
             "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
             "DataHoraProcotolo = '" & NFeDataHora & "' " & _
             "WHERE CodigoNota = " & vsNumeroNota
     vgDb.Execute vsSQL
  ElseIf nroRecibo = 101 Then
     NFeValidate = "NFe CANCELADA"
     NFeChaveAcesso = ChaveAcesso
  ElseIf nroRecibo = 110 Then
     NFeValidate = "NFe DENEGADA"
     NFeChaveAcesso = ChaveAcesso
     vsSQL = "UPDATE NotaFiscal SET " & _
             "Denegada = 1, " & _
             "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
             "DataHoraProcotolo = '" & NFeDataHora & "' " & _
             "WHERE CodigoNota = " & vsNumeroNota
     vgDb.Execute vsSQL
  Else
     msgResultado = NFeMotivo
     NFeChaveAcesso = "0"
  End If
   
Caifora:
  MsgBox msgResultado, vbInformation + vbOKOnly, NFeValidate
  Set sistNFe = Nothing
  Screen.MousePointer = vbDefault
End Sub

Public Sub ConsultaReciboNFCe(Recibo As Variant, ChaveAcesso As String, Optional FormatoEmissao As String, Optional MostraMsg As Boolean = True)
Dim i As Integer
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

  Screen.MousePointer = vbHourglass
   
  iRetorno = ConfiguraDLLNFeNFCe(65, FormatoEmissao, sistNFe)
  
  i = 0
  dirXML = SQLExecutaRetorno("SELECT DiretorioXML from empresa", "DiretorioXML", "")
  If Vazio(dirXML) Then Exit Sub
  dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
  
buscaNFe:
  iRetorno = sistNFe.ConsultarReciboDeEnvio(Recibo)
   
  If Not iRetorno Then
     NFeValidate = "ERRO"
     NFeNumeroProtocolo = ""
     GoTo Caifora
  End If
   
  cStat = sistNFe.retConsRec.cStat
  NFeMotivo = sistNFe.retConsRec.xMotivo
  
  If cStat = 105 Or cStat = 217 Then
     i = i + 1
     If i > 5 And cStat = 105 Then
        msgResultado = "A Nota Fiscal ficou em PROCESSAMENTO." & vbNewLine & "Efetue a consulta do Recibo daqui 5 minutos novamente."
        NFeValidate = "NFe/NFCe PROCESSAMENTO"
        GoTo Caifora
     End If
     Sleep 10000
     GoTo buscaNFe
  ElseIf cStat = 106 Then
     msgResultado = str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe NÃO LOCALIZADA"
     iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
    
     If Not iRetorno Then
        NFeValidate = "ERRO"
        NFeNumeroProtocolo = ""
        GoTo Caifora
     End If
     cStat = sistNFe.retConsulta.cStat
     NFeMotivo = sistNFe.retConsulta.xMotivo
     If cStat = 613 Then
        NFeChaveAcesso = Mid(NFeMotivo, InStr(NFeMotivo, "Numerico da NF-e [") + 18, 44)
        iRetorno = sistNFe.ConsultarProtocolo(NFeChaveAcesso)
        cStat = sistNFe.retConsulta.cStat
        NFeMotivo = sistNFe.retConsulta.xMotivo
        
        If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
           NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
           NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
           ChaveAcesso = sistNFe.retConsulta.protNFe.infProt.chNFe
           GoTo continuaConsulta
        Else
           GoTo Caifora
        End If
     Else
        GoTo Caifora
     End If
  ElseIf cStat = 239 Then
     msgResultado = str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe COM PROBLEMAS"
     GoTo Caifora
  End If
  NFeNumeroProtocolo = sistNFe.retConsRec.protNFe.infProt.nProt
  NFeDataHora = sistNFe.retConsRec.protNFe.infProt.ProxyDhRecbto
  nroRecibo = sistNFe.retConsRec.protNFe.infProt.cStat
  nfeRetorno = sistNFe.retConsRec.protNFe.infProt.xMotivo
  NFeValidate = sistNFe.retConsRec.protNFe.infProt.chNFe
  
  If ChaveAcesso = "" Then ChaveAcesso = NFeValidate

  If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then iRetorno = 1

  If nroRecibo = 204 Or nroRecibo = 539 Then
     ChaveAcesso = Mid(nfeRetorno, InStr(nfeRetorno, "chNFe:") + 6, 44)
     nroRecibo = Mid(nfeRetorno, InStr(nfeRetorno, "nRec:") + 5)
     nroRecibo = Left(nroRecibo, Len(nroRecibo) - 1)
     Recibo = Left(Recibo, 15 - Len(nroRecibo)) + nroRecibo
     vsSQL = "UPDATE NotaFiscal SET " & _
             "ChavedeAcesso = '" & ChaveAcesso & "', " & _
             "NumeroRecibo = " & Recibo & " " & _
             "WHERE CodigoNota = " & vsNumeroNota
     vgDb.Execute vsSQL
     GoTo buscaNFe
  ElseIf nroRecibo <> 100 And nroRecibo <> 301 Then
     msgResultado = nroRecibo + " - " + nfeRetorno
     NFeValidate = "ERRO"
     GoTo Caifora
  ElseIf nroRecibo = 105 Then
     GoTo buscaNFe
  End If

  iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
  
  If Not iRetorno Then
     NFeValidate = "ERRO"
     NFeNumeroProtocolo = ""
     GoTo Caifora
  End If

  cStat = sistNFe.retConsulta.cStat
  NFeMotivo = sistNFe.retConsulta.xMotivo
  If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
     NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
     NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
  Else
     NFeDataHora = ""
     NFeNumeroProtocolo = ""
  End If

continuaConsulta:
  msgResultado = "Chave NF-e.: " + ChaveAcesso & vbCrLf
  msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
  msgResultado = msgResultado + "Recibo.: " + Recibo & vbCrLf
  msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
  msgResultado = msgResultado + "Resposta da Fazenda.: " + str$(cStat) + " - " & NFeMotivo
  If nroRecibo = 100 Then
     NFeValidate = "NFe AUTORIZADA"
     NFeChaveAcesso = ChaveAcesso
     vsSQL = "UPDATE NotaFiscal SET " & _
             "ChavedeAcesso = '" & ChaveAcesso & "', " & _
             "Enviada = 1, " & _
             "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
             "DataHoraProcotolo = '" & NFeDataHora & "' " & _
             "WHERE CodigoNota = " & vsNumeroNota
     vgDb.Execute vsSQL
  ElseIf nroRecibo = 101 Then
     NFeValidate = "NFe CANCELADA"
     NFeChaveAcesso = ChaveAcesso
  ElseIf nroRecibo = 110 Then
     NFeValidate = "NFe DENEGADA"
     NFeChaveAcesso = ChaveAcesso
     vsSQL = "UPDATE NotaFiscal SET " & _
             "Denegada = 1, " & _
             "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
             "DataHoraProcotolo = '" & NFeDataHora & "' " & _
             "WHERE CodigoNota = " & vsNumeroNota
     vgDb.Execute vsSQL
  Else
     msgResultado = NFeMotivo
     NFeChaveAcesso = "0"
  End If
   
Caifora:
  MsgBox msgResultado, vbInformation + vbOKOnly, NFeValidate
  Set sistNFe = Nothing
  Screen.MousePointer = vbDefault
End Sub

Public Sub consultaNFe(ChaveAcesso As Variant, Optional NaoMostraMSG As Boolean) 'Sub que faz a consulta da NFe na Receita adaptada para nfe 3.1 ass 668

   On Error GoTo deuErro

   Screen.MousePointer = vbHourglass

   Dim sistNFe As snfe.Util
   Set sistNFe = New snfe.Util
   
   iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)

   iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
   
   cStat = sistNFe.retConsulta.cStat
   NFeMotivo = sistNFe.retConsulta.xMotivo
   If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
      NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
      NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
   Else
      NFeDataHora = ""
      NFeNumeroProtocolo = ""
   End If

   NFeChaveAcesso = ChaveAcesso
   msgResultado = "Chave NF-e.: " & ChaveAcesso & vbCrLf
   msgResultado = msgResultado + "Protocolo.: " & NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data e Hora.: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + str$(cStat) & " - " & NFeMotivo
   
   If cStat = 100 Then
      NFeValidate = "NFe AUTORIZADA"
   ElseIf cStat = 101 Then
      NFeValidate = "NFe CANCELADA"
   ElseIf cStat = 110 Then
      NFeValidate = "NFe DENEGADA"
   Else
      msgResultado = NFeMotivo
   End If

   If Not NaoMostraMSG Then MsgBox msgResultado, vbInformation + vbOKOnly, NFeValidate

   GoTo Caifora
   
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
   Err.Clear
   cStat = 0
   NFeNumeroProtocolo = ""
   NFeDataHora = ""
   NFeMotivo = ""
   msgResultado = ""
   NFeValidate = ""

Caifora:
   Set sistNFe = Nothing
   Screen.MousePointer = vbDefault
End Sub

Public Sub consultaNFCe(ChaveAcesso As Variant, Optional NaoMostraMSG As Boolean) 'Sub que faz a consulta da NFe na Receita adaptada para nfe 3.1 ass 668

   On Error GoTo deuErro

   Screen.MousePointer = vbHourglass

   Dim sistNFe As snfe.Util
   Set sistNFe = New snfe.Util
   
    'Dim vCHAVE As String
    'vCHAVE = r!NFCeChaveAcesso
    If ChaveAcesso = Empty Then GoTo Caifora
   
   iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)

   iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
   
   cStat = sistNFe.retConsulta.cStat
   NFeMotivo = sistNFe.retConsulta.xMotivo
   
   Dim vStatusMotivo As Boolean
   vStatusMotivo = InStr(12, NFeMotivo, "indisponivel")
   
   If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
      If vStatusMotivo = False Then
        NFeDataHora = sistNFe.retConsulta.protNFe.infProt.ProxyDhRecbto
        NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
      Else
        NFeDataHora = ""
        NFeNumeroProtocolo = "0"
      End If
   Else
      NFeDataHora = ""
      NFeNumeroProtocolo = ""
   End If
   
   msgResultado = "Chave NF-e.: " & ChaveAcesso & vbCrLf
   msgResultado = msgResultado + "Protocolo.: " & NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data e Hora.: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + str$(cStat) & " - " & NFeMotivo
   
   If cStat = 100 Then
      NFeValidate = "NFe AUTORIZADA"
   ElseIf cStat = 101 Then
      NFeValidate = "NFe CANCELADA"
   ElseIf cStat = 110 Then
      NFeValidate = "NFe DENEGADA"
   Else
      msgResultado = NFeMotivo
   End If

   If Not NaoMostraMSG Then MsgBox msgResultado, vbInformation + vbOKOnly, NFeValidate

   GoTo Caifora
   
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
   Err.Clear
   cStat = 0
   NFeNumeroProtocolo = ""
   NFeDataHora = ""
   NFeMotivo = ""
   msgResultado = ""
   NFeValidate = ""

Caifora:
   Set sistNFe = Nothing
   Screen.MousePointer = vbDefault
End Sub

Public Sub ConsultaStatus(Optional ModeloNF As Integer = 55)  'Sub que consulta o Status do Serviço da Receita
Dim sistNFe As snfe.Util
   
   On Error GoTo deuErro
   
   Set sistNFe = New snfe.Util

   Screen.MousePointer = vbHourglass
   
   iRetorno = ConfiguraDLLNFeNFCe(ModeloNF, "1", sistNFe)

   sistNFe.exibirAvisos = False

   'NFe
   If ModeloNF = 55 Then
      iRetorno = sistNFe.ConsultarStatusServico
      NFeResposta = CStr(sistNFe.retStatusWS.cStat) + " - " + sistNFe.retStatusWS.xMotivo
    End If
   'NFCe
   If ModeloNF = 65 Then
      iRetorno = sistNFe.ConsultarStatusServico
      NFeResposta = CStr(sistNFe.retStatusWS.cStat) + " - " + sistNFe.retStatusWS.xMotivo
   End If

   MsgBox "CONSULTA DE STATUS DO WS" & vbNewLine & vbNewLine & NFeResposta, vbInformation + vbOKOnly

   Set sistNFe = Nothing

   Screen.MousePointer = vbDefault
   
   Exit Sub
   
deuErro:   'estava desabilitado, habilitei no dia que fui testar o açougue uniao
   MsgBox Err.Description, vbCritical
   Err.Clear
   Set sistNFe = Nothing

   Screen.MousePointer = vbDefault
End Sub

Public Function TransmitirCCe(ChaveAcesso As Variant, DATA As Variant, nProtocolo As Variant, SeqCorrecao As Variant, textoCorrecao As Variant) As Boolean  'Função para envio da carta de correção da NFe
Dim IdLote As Long, dhEvento As String, CNPJ As String
Dim dirXML As String
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

  Screen.MousePointer = vbHourglass
  
  IdLote = Int(Rnd(918274 * 999) * 1000)
  dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
  CNPJ = SQLExecutaRetorno("SELECT CNPJ FROM Empresa", "CNPJ", "")
  
  iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
  
  iRetorno = sistNFe.CartaCorrecao(CNPJ, IdLote, SeqCorrecao, ChaveAcesso, textoCorrecao, xCaminhoXML)
     
  If Not iRetorno Then
     NFeNumeroProtocolo = ""
     GoTo Caifora
  End If
   
   cStat = sistNFe.retEnvEvento.cStat
   NFeMotivo = sistNFe.retEnvEvento.xMotivo
   If cStat = 128 Then
     cStat2 = sistNFe.retEnvEvento.retEvento.infEvento.cStat
     NFeValidate = sistNFe.retEnvEvento.retEvento.infEvento.xMotivo
     NFeNumeroProtocolo = sistNFe.retEnvEvento.retEvento.infEvento.nProt
     NFeDataHora = sistNFe.retEnvEvento.retEvento.infEvento.dhRegEvento
   Else
     cStat2 = 0
     NFeValidate = ""
     NFeNumeroProtocolo = ""
     NFeDataHora = ""
   End If
   
   If cStat2 = 135 Or cStat2 = 155 Then
      GoTo continua
   Else
      MsgBox CStr(cStat2) & " - " & NFeMotivo, vbInformation
      GoTo Caifora
   End If
   
continua:
   msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + str(cStat) & " - " & NFeMotivo
   
   MsgBox msgResultado, vbInformation + vbOKOnly, "ENVIO CCe"
  
   Screen.MousePointer = vbDefault
   Set sistNFe = Nothing
   TransmitirCCe = True
   Exit Function

Caifora:
   Set sistNFe = Nothing
 
   Screen.MousePointer = vbDefault
   TransmitirCCe = False
End Function

'Fornecedor!CNPJCPF, ChaveAcesso, DataHora, TipoEvento, Justificativa
Public Function TransmitirManDest(CNPJ As Variant, ChaveAcesso As Variant, DATA As Variant, TipoEvento As Variant, Justificativa As Variant, Optional SemMSG As Boolean = False) As Boolean  'Função para envio da carta de correção da NFe
Dim IdLote As Long, dhEvento As String
Dim dirXML As String
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

  Screen.MousePointer = vbHourglass
  
  IdLote = Int(Rnd(918274 * 999) * 1000)
  dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
  
  iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
  
  iRetorno = sistNFe.ManifestacaoDestinatario(TipoEvento, IdLote, CNPJ, ChaveAcesso, dhEvento, Justificativa, 1, NFeResposta, xCaminhoXML)
  
   cStat = sistNFe.retEnvEvento.cStat
   NFeMotivo = sistNFe.retEnvEvento.xMotivo
   If cStat = 128 Then
      cStat2 = sistNFe.retEnvEvento.retEvento.infEvento.cStat
      NFeValidate = sistNFe.retEnvEvento.retEvento.infEvento.xMotivo
      NFeNumeroProtocolo = sistNFe.retEnvEvento.retEvento.infEvento.nProt
      NFeDataHora = sistNFe.retEnvEvento.retEvento.infEvento.dhRegEvento
   Else
     cStat2 = 0
     NFeValidate = ""
     NFeNumeroProtocolo = ""
     NFeDataHora = ""
   End If
  
  If cStat = 135 Or cStat2 = 135 Then
     GoTo continua
  Else
     MsgBox str(cStat2) & " - " & NFeValidate & vbNewLine & "CHAVE: " & ChaveAcesso, vbInformation, "ERRO"
     GoTo Caifora
  End If
         
continua:
  msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
  msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
  msgResultado = msgResultado + "Resposta da Fazenda.: " + str(cStat2) & " - " & NFeValidate
    
  If Not SemMSG Then MsgBox msgResultado, vbInformation + vbOKOnly, "Envio Manifestação do Destinatário"
  
  Screen.MousePointer = vbDefault
  Set sistNFe = Nothing
  TransmitirManDest = True
  Exit Function

Caifora:
  Set sistNFe = Nothing
 
  Screen.MousePointer = vbDefault
  TransmitirManDest = False
End Function

Public Function TransmitirConsultaNFDestinada(ultNSU As Variant) As Boolean
Dim iCont As Integer, qtRegistro As Long, i As Long
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util
Dim OBJDocumento As New MSXML2.DOMDocument30
Dim xChaveNF As String, xCNPJ As String, xNomeFornecedor As String, xDataEmissao As String, xValorNF As String, xSituacaoNF As String

  Screen.MousePointer = vbHourglass
  
  On Error GoTo deuErro
  
  iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
  
  iCont = 0
  nfeRetorno = ""
  
ConsultaNovamente:
  iRetorno = sistNFe.DistribuicaoDFe(LPad(ultNSU, 15, "0"), cStat, qtRegistro, retXML, msgRetWS)
  
  If Not iRetorno Then GoTo Caifora
  cStat = sistNFe.retDistDFeInt.cStat
  NFeMotivo = sistNFe.retDistDFeInt.xMotivo
  NFeDataHora = sistNFe.retDistDFeInt.dhResp
  retindCont = 0   'sistNFe.retDistDFeInt.maxNSU
  retultNSU = sistNFe.retDistDFeInt.ultNSU
  
  If cStat = 215 Then
     MsgBox NFeMotivo, vbInformation, "Consulta NF Destinada"
     GoTo Caifora
  ElseIf cStat = 656 Then
     MsgBox CLng(cStat) & " - " & NFeMotivo, vbExclamation, "Consulta NF Destinada"
     GoTo Caifora
  ElseIf cStat = 137 And iCont = 0 Then
     retindCont = 1
     If ultNSU = sistNFe.retDistDFeInt.maxNSU Then
        MsgBox CLng(cStat) & " - " & NFeMotivo, vbExclamation, "Consulta NF Destinada"
        GoTo Caifora
     End If
     ultNSU = sistNFe.retDistDFeInt.maxNSU
     iCont = 1
     GoTo ConsultaNovamente
  End If
  
  cStat2 = qtRegistro
  retornoTipo = sistNFe.retornoTipo
  retornoNSU = sistNFe.retornoNSU
  retornodhEmi = sistNFe.retornodhEmi
  retornocSitConf = sistNFe.retornocSitConf
  retornochNFe = sistNFe.retornochNFe
  retornoCNPJ = sistNFe.retornoCNPJ
  retornocSitNFe = sistNFe.retornocSitNFe
  retornodhRecbto = sistNFe.retornodhRecbto
  retornodigVal = sistNFe.retornodigVal
  retornoIE = sistNFe.retornoIE
  retornotpNF = sistNFe.retornotpNF
  retornovNF = sistNFe.retornovNF
  retornoxNome = sistNFe.retornoxNome

  'MsgBox NFeMotivo, vbInformation + vbOKOnly, "Consulta NF Destinada"
   
  If Vazio(msgRetWS) Then GoTo Caifora
  
  Screen.MousePointer = vbDefault
  Set sistNFe = Nothing
  Set OBJDocumento = Nothing
  TransmitirConsultaNFDestinada = True
  Exit Function

Caifora:
  Set sistNFe = Nothing
 
  Screen.MousePointer = vbDefault
  TransmitirConsultaNFDestinada = False
  Exit Function
  
deuErro:
  MsgBox Err.Description, vbInformation + vbOKOnly, "ERRO"
  Err.Clear
  Set sistNFe = Nothing
  Screen.MousePointer = vbDefault
  TransmitirConsultaNFDestinada = False
  
End Function



'Retorna fórmula direta campo CHAVEDEACESSO, tabela NOTAFISCAL
Public Sub GeraChavedeAcesso(NumeroNota As Variant, SerieNF As Variant, DataEmissao As Variant)
Dim sistNFe As snfe.Util
    Set sistNFe = New snfe.Util

    NFecNF = sistNFe.GetHashCode  'Deve retornar um número

    Set sistNFe = Nothing
End Sub

'Gera o CodigoNota para ser usado na Chave de Acesso
Public Function GeraCodigoNota() As Double
Dim sistNFe As snfe.Util
    Set sistNFe = New snfe.Util

    GeraCodigoNota = sistNFe.GetHashCode  'Deve retornar um número

    Set sistNFe = Nothing
End Function

'Converte a string para codificação UTF-8
'Este processo evita problemas de leitura via browser e principalmente no visualizador da RFB
Private Function UTF8_Encode(ByVal sStr As String)
    Dim l As Long, lChar As Integer, sUtf8 As String
    For l = 1 To Len(sStr)
        lChar = AscW(Mid(sStr, l, 1))
        If lChar < 128 Then
            sUtf8 = sUtf8 + Mid(sStr, l, 1)
        ElseIf ((lChar > 127) And (lChar < 2048)) Then
            sUtf8 = sUtf8 + Chr(((lChar \ 64) Or 192))
            sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
        Else
            sUtf8 = sUtf8 + Chr(((lChar \ 144) Or 234))
            sUtf8 = sUtf8 + Chr((((lChar \ 64) And 63) Or 128))
            sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
        End If
    Next l
    UTF8_Encode = sUtf8
End Function

Public Function TrataErroValidacao(ByVal mensagemErro As String) As String

    Dim Campo As String
    Dim Valor As String
    Dim MsgFinal As String
    
    ' Extrair nome do campo
    If InStr(mensagemErro, ":CFOP") > 0 Then
        Campo = "CFOP"
    End If
    
    ' Extrair valor informado
    Dim InicioValor As Long
    Dim FimValor As Long
    
    InicioValor = InStr(mensagemErro, "O valor '") + 9
    FimValor = InStr(InicioValor, mensagemErro, "' é inválido")
    
    If InicioValor > 9 And FimValor > 0 Then
        Valor = Mid(mensagemErro, InicioValor, FimValor - InicioValor)
    End If
    
    ' Montar mensagem amigável
    Select Case Campo
    
        Case "CFOP"
            MsgFinal = "CFOP inválido." & vbCrLf & _
                       "O código informado foi: " & Valor & vbCrLf & _
                       "Verifique se o CFOP está correto e permitido para esta operação."
        
        Case Else
            MsgFinal = "Erro na validação da NF-e." & vbCrLf & _
                       "Detalhes técnicos: " & mensagemErro
    End Select
    
    TrataErroValidacao = MsgFinal

End Function

Public Sub SaveKey(hKey As Long, strPath As String)
Dim r
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim DataType As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim r
Dim lValueType
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
    If GetString = "" Then
    GetString = ""
    End If
End Function

Public Function SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim keyhand As Long
Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    r = RegCloseKey(keyhand)
    If r = 0 Then
    SaveString = ""
    Else
    SaveString = ""
    End If
End Function

Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim LBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, LBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = LBuf
        End If
    End If
    r = RegCloseKey(keyhand)
    If GetDWord = "" Then
    GetDWord = ""
    End If
End Function

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
Dim lResult As Long
Dim keyhand As Long
Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
    If r = 0 Then
    SaveDword = ""
    Else
    SaveDword = ""
    End If
End Function

Public Function DownloadXML(chaveNFe As String) As Boolean
Dim Parametros As ADODB.Recordset
'instancia o componente
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

    vsSQL = "SELECT * FROM Empresa"
    RsOpen Parametros, vsSQL
    
    dirXML = Parametros!DiretorioXML
    
    iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
    
    iRetorno = sistNFe.DownloadXML(chaveNFe, dirXML, NFeResposta, xCaminhoXML)
       
    If Not iRetorno Then GoTo Caifora
   
    DownloadXML = True
    
    Set sistNFe = Nothing
    
    Exit Function
Caifora:
    DownloadXML = False
    xCaminhoXML = ""
    Set sistNFe = Nothing
End Function
