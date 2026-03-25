Attribute VB_Name = "NFe_DLL"
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
Public xCaminhoXML As String, xCaminhoXMLAuxiliar As String, xCaminhoTXT As String, xCaminhoPDF As String, dirXML As String, vsSQL As String
Public vsNumeroNota As Variant, UTC As String, mensagemAlerta As String, mensagemErro As String
Dim nroRecibo As String, nroProtocolo As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function ConfiguraDLLNFeNFCe(Modelo As Integer, TipoEmissao As String, ByRef objNFeNFCe As snfe.Util) As Boolean
Dim ComandoSQL As String
   
   On Error GoTo deuErro
   
   Set objNFeNFCe = New snfe.Util
   
   ComandoSQL = "SELECT CNPJ, Razao, Cidade, Estado, CodigoIBGE, CRT, AmbienteNF, DiretorioXML, CertificadoDigital, " & _
                "NFCeIDToken, NFCeCSC, LicencaDLL, Email, caminho " & _
                "FROM Empresa"
                
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
                                       "02.382.419/0001-80", _
                                       "OnLine Info", _
                                       SQLExecutaRetorno(ComandoSQL, "LicencaDLL", ""))
                                       
                                          
   iRetorno = objNFeNFCe.ConfigurarEmail("smtp.gmail.com", 587, 30000, True, "nferondosoft@gmail.com", "1234rondo", True, SQLExecutaRetorno(ComandoSQL, "Email", "ekklesiasoft@gmail.com"), SQLExecutaRetorno(ComandoSQL, "Razao", "OnLine Info"))
   
   iRetorno = objNFeNFCe.ConfigurarDANFe(SQLExecutaRetorno(ComandoSQL, "caminho", ""), True, False, True, False)
   
   iRetorno = objNFeNFCe.CarregarConfiguracoes
   
   ConfiguraDLLNFeNFCe = True
   
   Exit Function
                                       
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
   Err.Clear
End Function

Public Function TransmitirNFe(ByVal NumeroNota As Variant, ByVal SerieNF As Variant, Optional PodeEnviar As Boolean = False) As Boolean  'Função que monta o arquivo XML e faz o envio para a Receita
 Dim txtNumerado As String, Retorno As String, vsNFe As String, dhEmi As String
 Dim Parametros As New ADODB.Recordset, Destinatario As New ADODB.Recordset, Produtos As New ADODB.Recordset
 Dim NFe As New ADODB.Recordset, NFeItens As New ADODB.Recordset, NFeParcelas As New ADODB.Recordset, NFeOBS As New ADODB.Recordset, NFeAutorizados As New ADODB.Recordset, NFeReferenciadas As New ADODB.Recordset
 Dim NFeMedicamentos As New ADODB.Recordset, NFeArmamento As New ADODB.Recordset, NFeCombustivel As New ADODB.Recordset, NFeVeiculos As New ADODB.Recordset

 Dim n As Integer, i As Long, destIE As String
 Dim vsXML As String, XMLAuxiliar As String, XMLAuxiliarParcelas As String
 Dim msgerro As String, qterro As Long, Prod_DetEspecifico As String
 Dim vlTrib As Double, vCredICMSSN As Double

 Screen.MousePointer = vbHourglass

 On Error GoTo deuErro

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
    If Not Vazio(NFe!ChavedeAcesso) Then
       iRetorno = sistNFe.ConsultarProtocolo(NFe!ChavedeAcesso)
       NFeChaveAcesso = NFe!ChavedeAcesso
       cStat = sistNFe.retConsulta.cStat
       NFeMotivo = sistNFe.retConsulta.xMotivo
       NFeDataHora = ""
       NFeNumeroProtocolo = ""
       If cStat = 100 Or cStat = 101 Or cStat = 110 Then
          Sleep 10000
          NFeDataHora = sistNFe.retConsulta.protNFe.infProt.dhRecbto
          NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
          GoTo buscaNFe
       Else
          If Not Vazio(NFe!NumeroRecibo) Then
             iRetorno = sistNFe.ConsultarReciboDeEnvio(NFe!NumeroRecibo)
             cStat = sistNFe.retConsRec.cStat
             NFeMotivo = sistNFe.retConsRec.xMotivo
             If cStat = 104 Then
                cStat2 = sistNFe.retConsRec.protNFe.infProt.cStat
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
       GoTo caiFora
    End If
    
    iRetorno = sistNFe.IncluirNF(mensagemAlerta, mensagemErro)
    
    '===================grupo de identificação do emitente (grupo B do Manual de integração - páginas 90)=======================
    iRetorno = sistNFe.GerarEmitente(RemoveAcento(Parametros!RAZAO), RemoveAcento(Parametros!Fantasia), Parametros!CNPJ, "", Parametros!IE, "", "", "", Left(Parametros!CRT, 1), RemoveAcento(Parametros!ENDERECO), Parametros!Numero, "", Parametros!bairro, Parametros!CodigoIBGE, RemoveAcento(Parametros!Cidade), Parametros!Estado, Parametros!CEP, 1058, "BRASIL", Parametros!Telefone, mensagemAlerta, mensagemErro)
    
    '======= grupo de identificação da NF-e - grupo B do Manual de integração - páginas 86 a 89
    Dim dhContingencia As String, justContingencia As String
    If Left(NFe!FormatoEmissaoNFe, 1) <> "1" Then
       dhContingencia = Format(NFe!ContingenciaDataHora, "yyyy-mm-ddThh:mm:ss") & UTC 'v2.03 - dhCont  AAAA-MM-DDTHH:MM:SS
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

    iRetorno = sistNFe.GeraIdentificacao(NFe!cCodigoNota, NFe!NaturezaOperacao, 55, 1, NFe!NumeroNota, Format(NFe!DataEmissao, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ss") & UTC, Format(NFe!DataSaida, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ss") & UTC, NFe!TipoDocumento, 0, NFe!IdentificadorDestino, Parametros!CodigoIBGE, 1, Left(NFe!FormatoEmissaoNFe, 1), Left(NFe!FinalidadeEmissaoNFe, 1), indFinal, 1, Left$("ONLINE COMMERCE - v." & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision), 20), dhContingencia, justContingencia, mensagemAlerta, mensagemErro)
   
    If NFe!ChavedeAcessoAdicional <> "" Then
       iRetorno = sistNFe.GerarNotasReferenciadas("NFe", NFe!ChaveAcessoAdicional, 0, "", "", "", "", "", "", mensagemAlerta, mensagemErro)
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
    Dim xRazaoSocial As String, xCNPJ As String, xCPF As String, xTelefone As String
    'CLIENTE COM CPF
    If Len(Destinatario!CPF) = 18 Then
      xCNPJ = Trim(Destinatario!CPF)                                   ' CNPJ do destinatario sem máscara de formatação
      xCPF = ""
    Else
      xCPF = Trim(Destinatario!CPF)                                   ' CPF do destinatario, uso exclusivo do Fisco
      xCNPJ = ""
    End If

    If NFe!TipoCliente = "FORNECEDOR" Then
        xRazaoSocial = RemoveAcento(Trim(Left(Destinatario!RAZAO, 60)))           ' Razão social do destinatario, evitar caracteres acentuados e &
        xTelefone = Trim(Retira(Destinatario!Telefone, "()-. ", UM_A_UM))     ' número do telefone sem máscara
    Else
        xRazaoSocial = RemoveAcento(Trim(Left(Destinatario!nome, 60)))           ' Razão social do destinatario, evitar caracteres acentuados e &
        xTelefone = Trim(Retira(Destinatario!Telefone1, "()-. ", UM_A_UM))   ' número do telefone sem máscara
    End If
      
    If Parametros!AmbienteNF = 2 Then
       xCNPJ = "99999999000191"
       xRazaoSocial = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
       Destinatario!IE = ""
       NFe!IndicadorIEDestinatario = "9"
    End If
    
    iRetorno = sistNFe.GerarDestinatario(4, xRazaoSocial, xCNPJ, xCPF, "", destIE, NFe!InscricaoMunicipal, Left$(NFe!IndicadorIEDestinatario, 1), "", RemoveAcento(Destinatario!ENDERECO), Destinatario!Numero, RemoveAcento(Destinatario!Ponto_de_referencia), RemoveAcento(Destinatario!bairro), Destinatario!CodigoIBGE, RemoveAcento(Destinatario!Cidade), Destinatario!Estado, Retira(Destinatario!CEP, ".- ", UM_A_UM), 1058, "BRASIL", xTelefone, Destinatario!Correio_eletronico, mensagemAlerta, mensagemErro)
    
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
    
    vsSQL = "SELECT * FROM NotaFiscalAutorizados WHERE CodigoNota = " & NumeroNota
    RsOpen NFeAutorizados, vsSQL
    
    For i = 1 To NFeAutorizados.RecordCount
       xCNPJ = ""
       xCPF = ""
       If Len(NFeAutorizados!CNPJCPF) = 18 Then
          xCNPJ = NFeAutorizados!CNPJCPF
       Else
          xCPF = NFeAutorizados!CNPJCPF
       End If
       
       iRetorno = sistNFe.GerarAutorizadosXML(xCNPJ, xCPF, mensagemAlerta, mensagemErro)
    Next
        
    vsSQL = "SELECT NotaFiscalItens.* " & _
            "FROM NotaFiscalItens " & _
            "WHERE CodigoNota = " & NumeroNota & " " & _
            "ORDER BY Item"

    RsOpen NFeItens, vsSQL
    
    n = NFeItens.RecordCount

    For i = 1 To n
        vsSQL = "SELECT * FROM produtos WHERE CODIGO = " & NFeItens!CodigoProduto
        RsOpen Produtos, vsSQL
        If Produtos.RecordCount = 0 Then
           NFeMotivo = "CADASTRO DO PRODUTO NÃO ENCONTRADO!" & vbNewLine & vbNewLine & "PRODUTO: " & NFeItens!NomeProduto
           GoTo caiFora
        End If
        
        '================grupo de detalhe do produto (grupo I01 do Manual de integração - páginas 95)=======================
        Dim infAdiProd As String
        If NFeItens!ValorTributos > 0 Then
           infAdiProd = RemoveAcento(Trim(Left(NFeItens!InformacoesAdicionaisProduto, 400))) & " - Valor Aproximado dos Tributos R$ " & Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)        ' informações adicionais do produto
           vlTrib = vlTrib + NFeItens!ValorTributos
        Else
           infAdiProd = RemoveAcento(Trim(Left(NFeItens!InformacoesAdicionaisProduto, 500)))     ' informações adicionais do produto
        End If
        iRetorno = sistNFe.GerarItens(i, NFeItens!CodigoProduto, RemoveAcento(NFeItens!NomeProduto), Produtos!NCM, "", "", IIf(Vazio(Produtos!COD_BARRA), "SEM GTIN", Produtos!COD_BARRA), IIf(Vazio(Produtos!COD_BARRA), "SEM GTIN", Produtos!COD_BARRA), _
                                      NFeItens!CFOP, NFeItens!QuantidadeComercial, NFeItens!ValorUnitarioComercializacao, NFeItens!UnidadeComercial, NFeItens!QuantidadeComercial, NFeItens!ValorUnitarioComercializacao, NFeItens!UnidadeComercial, (NFeItens!QuantidadeComercial * NFeItens!ValorUnitarioComercializacao), NFeItens!ValorFrete, NFeItens!ValorDesconto, _
                                      0, NFeItens!ValorSeguro, "", "", 0, "", "", "", "", "", 1, infAdiProd, mensagemAlerta, mensagemErro)
        
'        '   gera grupo do destinatário
'        Select Case NFeItens!TipoProduto     'Veículo|Medicamento|Armamento|Combustível
'          Case "Armamento"
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
'          Case "Combustível"
'            vsSQL = "SELECT * " & _
'                    "FROM NotaFiscalItensCombustivel " & _
'                    "WHERE CodigoNota = " & NumeroNota & " AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
'            RsOpen NFeCombustivel, vsSQL
'            If NFeCombustivel.RecordCount > 0 Then
'               Do While Not NFeCombustivel.EOF
'                  prod(i, 93) = NFeCombustivel!cProdANP                                              'cProdANP
'                  If NFeCombustivel!cProdANP = "210203001" Then
'                     prod(i, 94) = Format(NFeCombustivel!pMixGN, "#0.0000")                          'pMixGN
'                  End If
'                  prod(i, 95) = NFeCombustivel!CODIF                                                 'CODIF
'                  prod(i, 96) = Format(NFeCombustivel!qTemp, "#0.0000")                              'qTemp
'                  prod(i, 97) = NFeCombustivel!UFCons                                                'UFCons
'                  'CIDE
'                  If NFeCombustivel!qBCProd > 0 Then
'                     prod(i, 98) = Format(NFeCombustivel!qBCProd, "#0.0000")                         'qBCProd
'                     prod(i, 99) = Format(NFeCombustivel!vAliqProd, "#0.0000")                       'vAliqProd
'                     prod(i, 100) = Format(NFeCombustivel!vCIDE, "#0.00")                            'vCIDE
'                  End If
'                  NFeCombustivel.MoveNext
'               Loop
'            End If
'          Case "Medicamento"
'            vsSQL = "SELECT * " & _
'                    "FROM NotaFiscalItensMedicamento " & _
'                    "WHERE (NumeroNota = " & NumeroNota & " AND SerieNF = " & SerieNF & ") AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
'            RsOpen NFeMedicamentos, vsSQL
'            If NFeMedicamentos.RecordCount > 0 Then
'              Do While Not NFeMedicamentos.EOF
'                prod(i, 71) = IIf(Vazio(NFeMedicamentos!nLote), "0", NFeMedicamentos!nLote)
'                prod(i, 72) = Format(NFeMedicamentos!QuantidadeLote, "#0.000")
'                prod(i, 73) = IIf(IsNull(NFeMedicamentos!DataFabricacao), Format(DateAdd("yyyy", -1, Date), "mm/dd/yyyy"), Format(NFeMedicamentos!DataFabricacao, "yyyy-mm-dd"))
'                prod(i, 74) = IIf(IsNull(NFeMedicamentos!DataValidade), Format(DateAdd("m", 6, NFeMedicamentos!DataValidade), "mm/dd/yyyy"), Format(NFeMedicamentos!DataValidade, "yyyy-mm-dd"))
'                prod(i, 75) = Format(NFeMedicamentos!PMC, "#0.00")
'                NFeMedicamentos.MoveNext
'              Loop
'            End If
'          Case "Veículo"
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
'              prod(i, 79) = Prod_DetEspecifico
'            End If
'        End Select

        'Valor aproximado total de tributos federais, estaduais e municipais
        'If NFeItens!ValorTributos > 0 Then prod(i, 81) = Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)
        '=========dados do ICMS (grupo N01 do Manual de integração - páginas 100)=====================
        vCredICMSSN = 0
        If Parametros!pCreditoICMSSimplesNacional > 0 Then vCredICMSSN = Round(NFeItens!ValorTotalBruto * (Parametros!pCreditoICMSSimplesNacional / 100), 2)
        iRetorno = sistNFe.GerarItensImpostoEstadual(NFeItens!ValorTributos, Left$(NFeItens!CST, 1), Mid$(NFeItens!CST, 2), IIf(Vazio(NFeItens!modBC), 3, Left$(NFeItens!modBC, 1)), NFeItens!vBC, NFeItens!pICMS, NFeItens!vICMS, NFeItens!pRedBC, 0, 0, 0, _
                                                     IIf(Not Vazio(NFeItens!modBCST), Left(NFeItens!modBCST, 1), 5), NFeItens!pMVAST, NFeItens!pRedBCST, NFeItens!vBCST, NFeItens!pICMSST, NFeItens!vICMSST, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Parametros!pCreditoICMSSimplesNacional, vCredICMSSN, mensagemAlerta, mensagemErro)
        

        iRetorno = sistNFe.GerarItensImpostoFederal(IIf(Vazio(NFeItens!COFINSCST) Or IsNull(NFeItens!COFINSCST), "99", NFeItens!COFINSCST), NFeItens!COFINSvBC, NFeItens!COFINSpCOFINS, NFeItens!COFINSvCOFINS, 0, 0, _
                                                    IIf(Vazio(NFeItens!PISCST) Or IsNull(NFeItens!PISCST), "99", NFeItens!PISCST), NFeItens!PISvBC, NFeItens!PISpPIS, NFeItens!PISvPIS, 0, 0, _
                                                    IIf(Vazio(NFeItens!IPICST), "53", NFeItens!IPICST), NFeItens!IPIvBC, NFeItens!IPIpIPI, NFeItens!IPIvIPI, 0, 0, 999, "", "", "", 0, mensagemAlerta, mensagemErro)

        '   gera grupo do II - Importação
        If (NFeItens!IIvBC > 0) Then
           iRetorno = sistNFe.GerarItensImpostoII(NFeItens!IIvBC, NFeItens!IIvDespAdu, NFeItens!IIvII, NFeItens!IIvIOF, mensagemAlerta, mensagemErro)
        End If

        iRetorno = sistNFe.GerarItensIncluir(mensagemAlerta, mensagemErro)
        
        NFeItens.MoveNext
    Next

    'atualização de total
    iRetorno = sistNFe.GerarTotalProdutos(NFe!BaseICMS, NFe!ValorICMS, NFe!BaseICMSST, NFe!ValorICMSST, NFe!ValorCOFINS, NFe!ValorPIS, NFe!ValorIPI, NFe!ValorDesconto, NFe!ValorSeguro, NFe!ValorFrete, NFe!ValorOutrasDespesas, 0, 0, 0, 0, 0, 0, 0, _
                                          NFe!ValorImportacao, 0, NFe!ValorProdutos, NFe!ValorNota, vlTrib, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, mensagemAlerta, mensagemErro)

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

    vsSQL = "SELECT * " & _
            "FROM NotaFiscalParcelas " & _
            "WHERE CodigoNota = " & NumeroNota
    RsOpen NFeParcelas, vsSQL

    iRetorno = sistNFe.GerarCobranca(NFe!NumeroNota, 0, NFe!ValorNota, NFe!ValorNota, mensagemAlerta, mensagemErro)
    If NFeParcelas.RecordCount > 0 Then
       For i = 0 To NFeParcelas.RecordCount - 1
           iRetorno = sistNFe.GerarCobrancaDuplicatas(NFeParcelas!Documento, NFeParcelas!Vencimento, NFeParcelas!ValorDocumento, mensagemAlerta, mensagemErro)
           iRetorno = sistNFe.GerarPagamentos(1, 15, NFeParcelas!ValorDocumento, 0, 0, 0, "", "", mensagemAlerta, mensagemErro)
           NFeParcelas.MoveNext
       Next
    Else
       iRetorno = sistNFe.GerarPagamentos(0, 1, NFe!ValorNota, 0, 0, 0, "", "", mensagemAlerta, mensagemErro)
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
iRetorno = sistNFe.GerarXML(numero_nfe_gerado, xCaminhoXML, mensagemAlerta, mensagemErro)
id_chave = numero_nfe_gerado
numero_nfe_gerado = Replace(numero_nfe_gerado, "NFe", "")
NFeChaveAcesso = numero_nfe_gerado

If Not Vazio(NFeChaveAcesso) Then
   vsSQL = "UPDATE NotaFiscal SET " & _
           "ChavedeAcesso = '" & NFeChaveAcesso & "' " & _
           "WHERE CodigoNota = " & NumeroNota
   vgDb.Execute vsSQL
End If

If Not PodeEnviar Then GoTo NaoEnviou

iRetorno = sistNFe.EnviarNFe(NFe!NumeroNota, False, False)

cStat = sistNFe.retEnvio.cStat
NFeMotivo = sistNFe.retEnvio.xMotivo

If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
   MsgBox "*** Aparentemente Ocorreram Erros na Recepção do Lote (nfeAutorizacao)***" & vbLf & NFeMotivo, vbExclamation, "PROCESSO INTERROMPIDO - NfeAutorizacao"
   GoTo caiFora
End If

If cStat = 103 Then NFeNumeroRecibo = sistNFe.retEnvio.infRec.nRec
NFeDataHora = sistNFe.retEnvio.dhRecbto

If cStat <> 103 Then
   GoTo NaoEnviou
End If

vsSQL = "UPDATE NotaFiscal SET " & _
        "NumeroRecibo = " & NFeNumeroRecibo & " " & _
        "WHERE CodigoNota = " & NumeroNota
vgDb.Execute vsSQL

DoEvents

consultaNFe:

    iRetorno = sistNFe.ConsultarReciboDeEnvio(NFeNumeroRecibo)
   
    cStat = sistNFe.retConsRec.cStat
    NFeMotivo = sistNFe.retConsRec.xMotivo
    
    If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
       MsgBox NFeMotivo, vbExclamation, "Retorno Autorização"
       GoTo caiFora
    End If

    ' Testa erro 217-Rejeição: NF-e não consta na base de dados da SEFAZ <dhRecbto xmlns="http://www.portalfiscal.inf.br/nfe">2015-03-29T09:56:36-03:00
    If cStat = 217 Then
            Sleep 3000 ' Aguarda mais 3 segundos
            iRetorno = sistNFe.ConsultarReciboDeEnvio(NFeNumeroRecibo) 'refaz a consulta
            cStat = sistNFe.retConsRec.cStat
            NFeMotivo = sistNFe.retConsRec.xMotivo
            If cStat = 104 Then
               cStat2 = sistNFe.retConsRec.protNFe.infProt.cStat
               NFeMotivo = sistNFe.retConsRec.protNFe.infProt.xMotivo
            End If
            If InStr(NFeMotivo, "Erro") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0 Or InStr(NFeMotivo, "Rejeição") > 0 Then
               MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote (nfeRetAutorizacao)***" & vbLf & NFeMotivo, vbExclamation, "PROCESSO INTERROMPIDO - NfeRetAutorizacao"
               GoTo caiFora
            End If
    End If

    If cStat = 105 Or cStat = 217 Then
       i = i + 1
       If i > 5 And cStat = 105 Then
          msgResultado = "A Nota Fiscal ficou em PROCESSAMENTO." & vbNewLine & "Efetue a consulta do Recibo daqui 5 minutos novamente."
          MsgBox msgResultado, vbInformation + vbOKOnly, "Transmitir NFe"
          GoTo PodeSair
       End If
       Sleep 10000
       GoTo buscaNFe
    End If
    NFeNumeroProtocolo = sistNFe.retConsRec.protNFe.infProt.nProt
    NFeDataHora = sistNFe.retConsRec.protNFe.infProt.dhRecbto
    cStat2 = sistNFe.retConsRec.protNFe.infProt.cStat
    nfeRetorno = sistNFe.retConsRec.protNFe.infProt.xMotivo
    NFeValidate = sistNFe.retConsRec.protNFe.infProt.chNFe
    DoEvents
    
    If cStat2 <> 100 And cStat2 <> 101 And cStat2 <> 110 And cStat2 <> 150 Then
       NFeMotivo = ""
       MsgBox Str$(cStat2) & " - " & nfeRetorno, vbExclamation, "Retorno Consulta Recibo"
       GoTo caiFora
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
         GoTo caiFora
      End If
   End If

   If Not iRetorno Then
      NFeNumeroProtocolo = ""
      GoTo caiFora
   End If
    
   NFeDataHora = sistNFe.retConsulta.protNFe.infProt.dhRecbto
   NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
   
   DoEvents

   If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo caiFora
  
   If cStat = 204 Or cStat = 539 Then
      NFeChaveAcesso = Mid(nfeRetorno, InStr(nfeRetorno, "chNFe:") + 6, 44)
      nroRecibo = Mid(nfeRetorno, InStr(nfeRetorno, "nRec:") + 5)
      nroRecibo = Left(nroRecibo, Len(nroRecibo) - 1)
      NFeNumeroRecibo = Left(NFeNumeroRecibo, 15 - Len(nroRecibo)) + nroRecibo
      If Vazio(NFeChaveAcesso) Or Len(NFeNumeroRecibo) < 15 Then
         NFeMotivo = nfeRetorno
         GoTo caiFora
      End If
      vsSQL = "UPDATE NotaFiscal SET " & _
              "ChavedeAcesso = '" & NFeChaveAcesso & "', " & _
              "NumeroRecibo = " & NFeNumeroRecibo & " " & _
              "WHERE CodigoNota = " & NumeroNota
      vgDb.Execute vsSQL
      GoTo buscaNFe
   ElseIf cStat <> 100 And cStat <> 301 Then
      NFeMotivo = Str$(cStat) + " - " + NFeMotivo
      GoTo NaoEnviou
   ElseIf cStat = 105 And cStat = 217 Then
      GoTo consultaNFe
   ElseIf cStat = 100 Then
      nfeRetorno = "Nota Fiscal Eletronica Autorizado o Uso."
      msgResultado = "Chave NF-e.: " + NFeChaveAcesso & vbCrLf
      msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
      msgResultado = msgResultado + "Recibo.: " + NFeNumeroRecibo & vbCrLf
      msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
      msgResultado = msgResultado + "Resposta da Fazenda.: " + Str$(cStat) & " - " & NFeMotivo
        
      MsgBox msgResultado, vbInformation + vbOKOnly
       
      DoEvents
        
      iRetorno = sistNFe.ConsultarProtocolo(NFeChaveAcesso)
             
      vsSQL = "UPDATE NotaFiscal SET " & _
              "ChavedeAcesso = '" & NFeChaveAcesso & "', " & _
              "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
              "DataHoraProcotolo = '" & NFeDataHora & "' " & _
              "WHERE CodigoNota = " & NumeroNota
      vgDb.Execute vsSQL
      
      vsSQL = "INSERT INTO NotaFiscalRecibos (CodigoNota, NumeroProtocolo, DataHora) Values " & _
              "(" & NumeroNota & ", " & NFeNumeroProtocolo & ", '" & NFeDataHora & "')"
      vgDb.Execute vsSQL
    
    End If

    'Gera PDF DANFE
    xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
    xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & Format(CDate(dhEmi), "yyyymm") & "\" & NFeChaveAcesso & "-procNFe.xml"            'Aqui Gera o DANFE
    Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 0, False, False, "")    'gera pdf

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
    Exit Function

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
    
    Exit Function
    
    Resume
    
caiFora:
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
    Exit Function
    
    Resume
    
deuErro:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Screen.MousePointer = vbDefault
    TransmitirNFe = False
End Function

Public Function TransmitirNFCe(ByVal NumeroNota As Variant, ByVal SerieNF As Variant, Optional PodeEnviar As Boolean = False, Optional ModeloNF As String = "65") As Boolean  'Função que monta o arquivo XML e faz o envio para a Receita
 Dim txtNumerado As String, Retorno As String, vsNFe As String, empUF As String, SQL As String
 Dim Parametros As New ADODB.Recordset
 Dim NFe As New ADODB.Recordset, NFeItens As New ADODB.Recordset, NFeParcelas As New ADODB.Recordset
 Dim NFeMedicamentos As New ADODB.Recordset, NFeArmamento As New ADODB.Recordset, NFeCombustivel As New ADODB.Recordset, NFeVeiculos As New ADODB.Recordset
 Dim NFeDeclaracaoImposto As New ADODB.Recordset, NFeAdicao As New ADODB.Recordset
          
 Dim n As Integer, i As Integer
 Dim vsXML As String, XMLAuxiliar As String, XMLAuxiliarParcelas As String
 Dim msgerro As String, qterro As Long, IdToken As String, Token As String
 Dim pDesconto As Double, pFrete As Double, pOutras As Double, pTributos As Double
 Dim vlPIS As Double, vlCOFINS As Double, vlTrib As Double, vlNF As Double
' On Error GoTo TransmitirNFCe_Error

' Dim sistNFCe As snfe.Util
' Set sistNFCe = New snfe.Util
'
' vlPIS = 0
' vlCOFINS = 0
' vlTrib = 0
' pFrete = 0
' pDesconto = 0
' pOutras = 0
' pTributos = 0
'
' vsSQL = "SELECT *, 0 AS COFINSAliquota, 0 AS PISAliquota FROM Empresa"
' RsOpen Parametros, vsSQL
'
' empUF = Parametros!Estado
'
' If ModeloNF = "65" Then
'    IdToken = LPad(Parametros!NFCeIDToken, 6, "0")
'    Token = Parametros!NFCeCSC
' End If
'
' Screen.MousePointer = vbHourglass
'
' If ModeloNF = "55" Then
'    'xVerProcNFe = SQLExecutaRetorno("SELECT VersaoLeiauteNFe As r FROM TbFiliais WHERE IdFilial = " & vgFilialNF, "r", "2.00")
'    'SaveString &H80000001, "nfe", "VerProc", xVerProcNFe
'    'vsSQL = "SELECT TbNotaFiscalProd.*, TbCidades.CodigoIBGE " & _
'    '        "FROM TbNotaFiscalProd INNER JOIN TbCidades ON TbNotaFiscalProd.Municipio = TbCidades.NomeCidade AND TbNotaFiscalProd.UF = TbCidades.UF " & _
'    '        "WHERE IdNFProd = " & NumeroNota
' Else
'    vsSQL = "SELECT TbNFCe.*, Cidade.CodigoMunicipio CodigoIBGE " & _
'            "FROM TbNFCe INNER JOIN Cidade ON TbNFCe.Municipio = Cidade.Nome AND TbNFCe.UF = Cidade.UF " & _
'            "WHERE IdNFProd = " & NumeroNota
' End If
' RsOpen NFe, vsSQL
' 'Set NFe = vgDb.OpenRecordset(vsSQL)
'
' If NFe.RecordCount > 0 Then
'    NFe.MoveFirst
'    '
'    '         criação dos grupos
'    '
'    '===================grupo de identificação do emitente (grupo B do Manual de integração - páginas 90)=======================
'    '
'    '        <>&" são caracteres reservados do XML e devem ser evitados ou substituídos
'    '        por &lt; &gy; &amp; &quot;
'    '
'    '        Vale ressaltar que as aplicações das UF devem mostrar DIAS &amp; DIAS TENTANDO S/A,
'    '        pois não entedem &amp; como &, assim talvez seja melhor substituir o & por e.
'    '
'    ReDim emit(15)
'    emit(0) = RemoveAcento(Trim(Parametros!RAZAO))                                               ' Razão social do emitente, evitar caracteres acentuados e &
'    emit(1) = RemoveAcento(Trim(Parametros!Fantasia))                                            ' Nome fantasia
'    emit(2) = RemoveAcento(Trim(Left(Parametros!ENDERECO, 60)))                                  ' logradouro
'    emit(3) = RemoveAcento(Trim(Parametros!Numero))                                         ' número, informar S/N quano inexistente para erro de Schema XML
'    emit(4) = ""  'RemoveAcento(Trim(Parametros!Complemento))                                    ' complemento do endereço, o conteúdo pode ser omitido
'    emit(5) = RemoveAcento(Trim(Parametros!bairro))                                              ' bairro
'    emit(6) = Trim(Parametros!CodigoIBGE)                                                        ' código do município (vide página 141 do manual), deve ser compatível com a UF
'    emit(7) = RemoveAcento(Trim(Left(Parametros!Cidade, 60)))                                    ' nome do município
'    emit(8) = Retira(Parametros!CEP, ".-/ ", UM_A_UM)                                            ' CEP - sem máscara
'    emit(9) = Retira(Parametros!Telefone, "().- ", UM_A_UM)                                      ' número do telefone sem máscara
'    emit(10) = Trim(Retira(Parametros!IE, ".,-/ ", UM_A_UM))                                     ' Inscrição Estadual do emitente sem máscara
'    emit(11) = Trim(Retira(Parametros!InscricaoMunicipal, ".,-/", UM_A_UM))                  ' Inscrição Municipal
'    If Not Vazio(emit(11)) Then emit(12) = "" 'Trim(Retira(Parametros!CNAEFiscal, ".,-/", UM_A_UM))  ' Código do CNAE
'    'emit(13) = Trim(Retira(Parametros!InscricaoEstadualSubsTributari, ".,-/", UM_A_UM))          ' Inscrição Estadual do ST
'    emit(14) = Parametros!CRT                                                                     ' <CRT> 1 – Simples Nacional; 2 – Simples Nacional – excesso de sublimite de receita bruta; 3 – Regime Normal
'
'    '
'    '======= grupo de identificação da NF-e - grupo B do Manual de integração - páginas 86 a 89
'    '
'    ReDim ide(29)
'    ide(0) = Left(Parametros!CodigoIBGE, 2)                        ' código da UF - tabela do IBGE: 35 - SP, 43 - RS, etc. (vide página 141 do manual)
'    ide(1) = NFe!NFeCodigoNota
'    ide(2) = RemoveAcento(NFe!NaturezaOperacao)                    ' natureza da operação
'    ide(3) = Left(NFe!NFeIndicadorFormaPagto, 1)                   ' Indicador da forma de pagamento  0 = Pagamento a vista / 1 = Pagamento a prazo / 2 = Outros
'    ide(4) = ModeloNF                                              ' modelo da nota fiscal eletronica
'    ide(5) = NFe!SerieNF                                           ' série única = 0
'    ide(6) = Val(NFe!NumeNota)                                     ' número da NF-e
'    ide(7) = Format(NFe!DataEmissao, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC       ' data de emissão
'    If Not IsNull(NFe!DataSaidaEntrada) Then ide(8) = Format(NFe!DataSaidaEntrada, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC  ' data em branco = 30/12/1899
'    ide(9) = IIf(NFe!TipoNF = "E", 0, 1)                           ' tipo do documento 0 - Entrada / 1 - Saida
'    ide(10) = Parametros!CodigoIBGE                                ' código do município do IBGE de ocorrência do FG do ICMS (vide página 141 do manual)
'    ide(11) = Left(NFe!NFeTipoEmissao, 1)                          ' forma de emissão da NF-e 1- normal, 2 - contingência FS, 3 - contingência SCAN, etc.
'    ide(12) = "1"  'Left(NFe!NFeFinalidadeEmissao, 1)                    ' finalidade da emissão da NF-e 1- NF-e normal
'    If Not Vazio(NFe!NFCeChaveAcessoReferenciada) Then
'       ide(13) = NFe!NFCeChaveAcessoReferenciada
'    End If
'    If ide(11) <> 1 Then
'       ide(15) = IIf(Vazio(NFe!NFCeDataHoraContingencia), Null, NFe!NFCeDataHoraContingencia) ' Data/Hora Contingencia
'       ide(16) = NFe!NFCeJustificativaContingencia                                            ' Justificativa Contingencia
'    End If
'    ide(18) = IIf(NFe!NFeConsumidorFinal, 1, 0)                          'Indica operação com Consumidor final
'    ide(17) = Left(NFe!NFeIdentificadorDestino, 1)                       'Identificador de local de destino da operação - 1 - Operação interna|2 - Operação interestadual|3 - Operação com exterior
'    ide(19) = Left(NFe!NFeIndicadorPresencaComprador, 1)                 'Indicador de presença do comprador no estabelecimento comercial no momento da operação - 0 - Não se aplica|1 - Operação presencial|2 - Operação não presencial, pela Internet|3 - Operação não presencial, Teleatendimento|4 - NFC-e em operação com entrega a domicílio|9 - Operação não presencial, outros
'    ide(27) = "4"
'    ide(28) = Left$("ONLINE COMMERCE - v." & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision), 20)
'    '
'    '================grupo de identificação do destinatario (grupo E do Manual de integração - páginas 92)=======================
'    '
'    ReDim dest(45)
'    If Len(NFe!CPF_CNPJ) = 0 Then dest(18) = "1"
'    If Len(Retira(NFe!CPF_CNPJ, ".,-/", UM_A_UM)) > 11 Then
'      dest(0) = Trim(NFe!CPF_CNPJ)                                        ' CNPJ do destinatario sem máscara de formatação
'      dest(0) = Retira(dest(0), ".-/", UM_A_UM)                           ' CNPJ do destinatario sem máscara de formatação
'    Else
'      dest(0) = Trim(NFe!CPF_CNPJ)                                        ' CPF do destinatario, uso exclusivo do Fisco
'      dest(0) = Retira(dest(0), ".-/", UM_A_UM)                           ' CPF do destinatario, uso exclusivo do Fisco
'    End If
'    dest(1) = RemoveAcento(Trim(Left(NFe!NomeRazSocial, 60)))             ' Razão social do destinatario, evitar caracteres acentuados e &
'    dest(2) = RemoveAcento(Trim(NFe!ENDERECO))                            ' logradouro
'    dest(3) = RemoveAcento(Trim(NFe!Num))                                 ' número, informar S/N quando inexistente para erro de Schema XML
'    dest(4) = ""                                                          ' complemento do endereço, o conteúdo pode ser omitido
'    dest(5) = RemoveAcento(Trim(NFe!bairro))                              ' bairro
'    dest(6) = Trim(NFe!CodigoIBGE)                                        ' código do município (vide página 141 do manual), deve ser compatível com a UF
'    dest(7) = RemoveAcento(Trim(NFe!Municipio))                           ' nome do município
'    dest(8) = Trim(NFe!uf)                                                ' sigla da UF
'    dest(9) = Retira(NFe!CEP, ".-/", UM_A_UM)                             ' CEP - sem máscara
'    dest(10) = NFe!CodigoPais                                             ' código do pais - deve fixo em 1058 - Brasil
'    dest(11) = RemoveAcento(NFe!NomePais)                                 ' nome do pais (Brasil ou BRASIL)
'    dest(12) = Trim(Retira(NFe!fone, "()-.", UM_A_UM))                    ' número do telefone sem máscara
'    dest(12) = Retira(dest(12), " ", UM_A_UM)                             ' número do telefone sem máscara
'    dest(12) = IIf(Left(dest(12), 1) = "0", Mid(dest(12), 2), dest(12))   ' número do telefone sem máscara
'    If Len(dest(12)) = 0 Then dest(12) = ""
'    dest(13) = Trim(Retira(NFe!InscEst, ".,-/", UM_A_UM))                 ' Inscrição Estadual do destinatario sem máscara
'    dest(14) = ""                                                         ' Inscrição SUFRAMA
'    dest(15) = ""                                                         ' Email
'    dest(16) = Left(NFe!NFeIndicadorIEDestinatario, 1)                    ' Indicador da IE do Destinatário - 1 - Contribuinte ICMS (informar a IE do destinatário)|2 - Contribuinte isento de Inscrição no cadastro de Contribuintes do ICMS|9 - Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS
'    dest(17) = ""                                                         ' Inscrição Municipal do Tomador do Serviço
'
'    If Left(Parametros!AmbienteNF, 1) = 2 Then
'       dest(18) = "1"
'       dest(1) = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
'       dest(13) = ""
'       dest(16) = "9"
'    End If
'
'    If NFe!CodigoPais <> 1058 Then dest(0) = ""
'
'    ReDim autXML(0)
'
'    autXML(0) = "" 'Retira(Nfe!CPF_CNPJ, ".-/", UM_A_UM)
'
'    If ModeloNF = "55" Then
''        vsSQL = "SELECT TbNotaFiscalProd_Itens.IdNFProd, TbNotaFiscalProd_Itens.IdNFProd_Item, TbNotaFiscalProd_Itens.CodProduto, TbNotaFiscalProd_Itens.IdProduto, TbNotaFiscalProd_Itens.DescricaoProduto, TbNotaFiscalProd_Itens.ValorOutras, " & _
''                "TbNotaFiscalProd_Itens.CodBarras, TbNotaFiscalProd_Itens.UN, TbNotaFiscalProd_Itens.CFOP, TbNotaFiscalProd_Itens.QtdeMov, TbNotaFiscalProd_Itens.ValorUnit, TbNotaFiscalProd_Itens.Desconto, TbNotaFiscalProd_Itens.Aliq_Icms AS Aliquota, " & _
''                "TbNotaFiscalProd_Itens.Bc_Icms, TbNotaFiscalProd_Itens.Bc_AliquotaReducao, TbNotaFiscalProd_Itens.Vlr_Icms,TbNotaFiscalProd_Itens.Aliq_IPI As AliqIPI, TbNotaFiscalProd_Itens.Vlr_IPI As ValorIPI, TbNotaFiscalProd_Itens.Valor_Frete As ValorFrete, " & _
''                "TbNotaFiscalProd_Itens.ICMSCST, TbNotaFiscalProd_Itens.PISCST, TbNotaFiscalProd_Itens.COFINSCST, TbNotaFiscalProd_Itens.IPICST, TbNotaFiscalProd_Itens.Codncm As NCM, TbNotaFiscalProd_Itens.ProdInfAdicional, TbNotaFiscalProd_Itens.ValorTributos, " & _
''                "TbNotaFiscalProd_Itens.BCSTRet, TbNotaFiscalProd_Itens.ICMSSTRet, TbNotaFiscalProd_Itens.BCImpostoImportacao, TbNotaFiscalProd_Itens.DespesasAduaneiras, TbNotaFiscalProd_Itens.ValorImpostoImportacao, TbNotaFiscalProd_Itens.ValorIOF " & _
''                "FROM TbNotaFiscalProd_Itens " & _
''                "WHERE TbNotaFiscalProd_Itens.IdNFProd = " & NumeroNota & " " & _
''                "ORDER BY TbNotaFiscalProd_Itens.IdNFProd, TbNotaFiscalProd_Itens.IdNFProd_Item"
'    Else
'        vsSQL = "SELECT TbNFCe_Itens.IdNFProd, TbNFCe_Itens.IdNFProd_Item, TbNFCe_Itens.CodProduto, TbNFCe_Itens.IdProduto, TbNFCe_Itens.DescricaoProduto, TbNFCe_Itens.ValorOutras, " & _
'                "TbNFCe_Itens.CodBarras, TbNFCe_Itens.UN, TbNFCe_Itens.CFOP, TbNFCe_Itens.QtdeMov, TbNFCe_Itens.ValorUnit, TbNFCe_Itens.Desconto, TbNFCe_Itens.Aliq_Icms AS Aliquota, " & _
'                "TbNFCe_Itens.Bc_Icms, TbNFCe_Itens.Bc_AliquotaReducao, TbNFCe_Itens.Vlr_Icms, TbNFCe_Itens.Aliq_IPI As AliqIPI, TbNFCe_Itens.Vlr_IPI As ValorIPI, TbNFCe_Itens.Valor_Frete As ValorFrete, " & _
'                "TbNFCe_Itens.ICMSCST, TbNFCe_Itens.PISCST, TbNFCe_Itens.COFINSCST, TbNFCe_Itens.IPICST, TbNFCe_Itens.Codncm As NCM, TbNFCe_Itens.ProdInfAdicional, TbNFCe_Itens.ValorTributos, " & _
'                "TbNFCe_Itens.BCSTRet, TbNFCe_Itens.ICMSSTRet, TbNFCe_Itens.BCImpostoImportacao, TbNFCe_Itens.DespesasAduaneiras, TbNFCe_Itens.ValorImpostoImportacao, TbNFCe_Itens.ValorIOF " & _
'                "FROM TbNFCe_Itens " & _
'                "WHERE TbNFCe_Itens.IdNFProd = " & NumeroNota & " " & _
'                "ORDER BY TbNFCe_Itens.IdNFProd_Item"
'    End If
'
'    RsOpen NFeItens, vsSQL
'    'Set NFeItens = vgDb.OpenRecordset(vsSQL)
'
'    n = NFeItens.RecordCount - 1
'
'    ReDim prod(n, 132)
'
'    'PDesconto = Format(Nfe!DescontoPromocional / (Nfe!Valor_NF_Prod - Nfe!DescontoPromocional), "######0.000000")
'
'    If NFe!Valor_NF_Prod > 0 Then
'       pFrete = Format((NFe!Valor_Frete / NFe!Valor_NF_Prod) * 100, "######0.000000")
'       pOutras = Format((NFe!OutrasDespesasAces / NFe!Valor_NF_Prod) * 100, "######0.000000")
'    End If
'
'    For i = 0 To n
'        '
'        '================grupo de detalhe do produto (grupo I01 do Manual de integração - páginas 95)=======================
'        '
'        prod(i, 0) = Trim(NFeItens!IdProduto)                                         ' código do produto
'        prod(i, 1) = IIf(Vazio(NFeItens!CodBarras), "", NFeItens!CodBarras)           ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN
'        prod(i, 2) = RemoveAcento(Trim(NFeItens!DescricaoProduto))                    ' código do produto, espaços em branco consecutivos ou no início ou fim do campo podem gerar erro de Schema XML, além de caracteres reservados do XML <>&"'
'        prod(i, 3) = NFeItens!NCM                                                     ' código NCM, pode ser omitido se não sujeito ao IPI
'        prod(i, 109) = "" '"AA1000;BB1001;CC1002;DD1003"                              '<NVE>
'        prod(i, 5) = Trim(Str(NFeItens!CFOP))                                         ' CFOP do operação, causa erro de XML se informado um código inexistente
'        prod(i, 6) = RemoveAcento(Trim(NFeItens!UN))                                  ' unidade de comercialização
'        prod(i, 7) = Format(NFeItens!QtdeMov, "######0.000")                          ' quantidade de comercialização
'        prod(i, 7) = Substitui(prod(i, 7), ",", ".", UM_A_UM)
'        prod(i, 8) = Format(NFeItens!ValorUnit, "######0.000")                        ' valor unitário de comercialização, campo de mera demonstração deve ser o resultado da divisão do vProd / qCom
'        prod(i, 8) = Substitui(prod(i, 8), ",", ".", UM_A_UM)
'        prod(i, 9) = Format(NFeItens!QtdeMov * NFeItens!ValorUnit, "######0.00")      ' valor do total do item
'        prod(i, 9) = Substitui(prod(i, 9), ",", ".", UM_A_UM)
'        prod(i, 10) = prod(i, 1)                                                      ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN, em geral é o mesmo código do EAN de comercialização
'        prod(i, 11) = RemoveAcento(Trim(NFeItens!UN))                                 ' unidade de tributação, na maioria dos casos é idêntico  ao vUnCom, pode diferente nos casos de produtos sujeitos a ST em que a unidade de pauta é diferente da unidade de comercialização
'                                                                                      ' Ex. unidade de comercialização = 1 pack de lata de cerveja => unidade de tributação = 1 lata (preço de pauta)
'        prod(i, 12) = Format(NFeItens!QtdeMov, "######0.000")                         ' quantidade de comercialização
'        prod(i, 12) = Substitui(prod(i, 12), ",", ".", UM_A_UM)
'        prod(i, 13) = Format(NFeItens!ValorUnit, "######0.000")                       ' valor unitário de tributação, campo de mera demonstração deve ser o resultado da divisão do vProd / qTrib
'        prod(i, 13) = Substitui(prod(i, 13), ",", ".", UM_A_UM)
'        prod(i, 14) = Format(NFeItens!ValorFrete, "######0.00")                       ' valor do frete, se cobrado do cliente deve ser rateado entre os itens de produto
'        prod(i, 14) = Substitui(prod(i, 14), ",", ".", UM_A_UM)
'        prod(i, 15) = Format(0, "######0.00")                                         ' valor do seguro, se cobrado do cliente deve ser rateado entre os itens de produto
'        prod(i, 15) = Substitui(prod(i, 15), ",", ".", UM_A_UM)
'        prod(i, 16) = Format(NFeItens!Desconto, "######0.00")                         ' valor do desconto concedido
'        prod(i, 16) = Substitui(prod(i, 16), ",", ".", UM_A_UM)
'        prod(i, 96) = Format(NFeItens!ValorOutras, "######0.00")                      ' valor das outras despesas
'        prod(i, 96) = Substitui(prod(i, 96), ",", ".", UM_A_UM)
'        If NFeItens!ValorTributos > 0 Then
'           prod(i, 53) = RemoveAcento(Trim(Left(NFeItens!ProdInfAdicional, 400))) & " - Valor Aproximado dos Tributos R$ " & Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)        ' informações adicionais do produto
'           vlTrib = vlTrib + NFeItens!ValorTributos
'        Else
'           prod(i, 53) = RemoveAcento(Trim(Left(NFeItens!ProdInfAdicional, 500)))     ' informações adicionais do produto
'        End If
'
'        If prod(i, 5) = "1603" Then
'           prod(i, 76) = 0                                                               ' Indica se o valor do item entra no valor total da NFe
'        Else
'           prod(i, 76) = 1                                                               ' Indica se o valor do item entra no valor total da NFe
'        End If
'
'        'Valor aproximado total de tributos federais, estaduais e municipais
'        If NFeItens!ValorTributos > 0 Then prod(i, 104) = Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)
'
'        '
'        '=========dados do ICMS (grupo N01 do Manual de integração - páginas 100)=====================
'        '
'        prod(i, 17) = Left(NFeItens!ICMSCST, 1)                         ' Tabela A - origem da mercadoria 0=nacional
'
'        prod(i, 18) = Right(NFeItens!ICMSCST, IIf(emit(14) = 1, 3, 2))   ' Tabela B - CST=00-tributação normal
'        prod(i, 19) = 3                                                  ' modalidade de determinação da BC = 3-valor da operação
'        prod(i, 20) = Format(NFeItens!Bc_Icms, "######0.00")             ' valor da BC do ICMS = vProd + vFrete + vSeguro
'        prod(i, 20) = Substitui(prod(i, 20), ",", ".", UM_A_UM)
'        prod(i, 21) = Format(NFeItens!Aliquota, "######0.00")            ' alíquota do ICMS
'        prod(i, 21) = Substitui(prod(i, 21), ",", ".", UM_A_UM)
'        prod(i, 22) = Format(NFeItens!Vlr_Icms, "######0.00")            ' valor do ICMS
'        prod(i, 22) = Substitui(prod(i, 22), ",", ".", UM_A_UM)
'        prod(i, 46) = "5"                                                ' modalidade de determinação da BC ICMS ST
'        prod(i, 47) = "" 'Format(0, "######0.00")                            ' percentual de valor de margem e valor adicionado
'        'prod(i, 47) = Substitui(prod(i, 47), ",", ".", UM_A_UM)
'        prod(i, 48) = "" 'Format(0, "######0.00")                            ' percentual de redução da BC do ICMS ST
'        'prod(i, 48) = Substitui(prod(i, 48), ",", ".", UM_A_UM)
'        prod(i, 49) = Format(NFeItens!BCSTRet, "######0.00")             ' BC do ICMS ST
'        prod(i, 49) = Substitui(prod(i, 49), ",", ".", UM_A_UM)
'        prod(i, 50) = Format(0, "######0.00")                            ' percentual do ICMSST
'        prod(i, 50) = Substitui(prod(i, 50), ",", ".", UM_A_UM)
'        prod(i, 51) = Format(NFeItens!ICMSSTRet, "######0.00")           ' valor do ICMS ST devido
'        prod(i, 51) = Substitui(prod(i, 51), ",", ".", UM_A_UM)
'        prod(i, 52) = Format(NFeItens!Bc_AliquotaReducao, "######0.00")  ' percentual de redução da BC
'        prod(i, 52) = Substitui(prod(i, 52), ",", ".", UM_A_UM)
'
'        If prod(i, 18) = "20" Or prod(i, 18) = "30" Or prod(i, 18) = "40" Or prod(i, 18) = "70" Or prod(i, 18) = "90" Then
'           'Prod(i, 88) = Format(0, "######0.00")                    ' vICMSDeson
'           'Prod(i, 88) = Substitui(Prod(i, 88), ",", ".", UM_A_UM)
'           prod(i, 85) = "9"                                        ' motDesICMS
'        End If
'
'        If emit(14) = "1" And (prod(i, 18) = "101" Or prod(i, 18) = "201") Then
'           prod(i, 17) = "0"                                                              ' Tabela A - origem da mercadoria 0=nacional
'           prod(i, 77) = Format(NFeItens!pCreditoICMSSimplesNacional, "#0.00")            ' <pCredSN>          Simples Nacional
'           prod(i, 77) = Substitui(prod(i, 77), ",", ".", UM_A_UM)
'           prod(i, 78) = Format((NFeItens!QtdeMov * NFeItens!ValorUnit) * (NFeItens!pCreditoICMSSimplesNacional / 100), "#0.00")   ' <vCredICMSSN>      Simples Nacional
'           prod(i, 78) = Substitui(prod(i, 78), ",", ".", UM_A_UM)
'        End If
'
'        'tag IPI
'        prod(i, 23) = NFeItens!IPICST                                                    'IPI <CST>
'        prod(i, 24) = Format(NFeItens!QtdeMov * NFeItens!ValorUnit, "######0.00")        'IPI <vBC>
'        prod(i, 24) = Substitui(prod(i, 24), ",", ".", UM_A_UM)                          'IPI <vBC>
'        prod(i, 25) = Format(NFeItens!AliqIPI, "######0.00")                             'IPI <pIPI>
'        prod(i, 25) = Substitui(prod(i, 25), ",", ".", UM_A_UM)                          'IPI <pIPI>
'        prod(i, 26) = Format(NFeItens!ValorIPI, "######0.00")                            'IPI <vIPI>
'        prod(i, 26) = Substitui(prod(i, 26), ",", ".", UM_A_UM)                          'IPI <vIPI>
'
'        '
'        '=========dados do PIS (grupo Q do Manual de Integração - páginas 110) =============
'        '
'        prod(i, 31) = IIf(Vazio(NFeItens!PISCST) Or IsNull(NFeItens!PISCST), "07", NFeItens!PISCST)
'        prod(i, 32) = Format(NFeItens!QtdeMov * NFeItens!ValorUnit, "######0.00")
'        prod(i, 32) = Substitui(prod(i, 32), ",", ".", UM_A_UM)
'        prod(i, 33) = Format(Parametros!PISAliquota, "###0.00")
'        prod(i, 33) = Substitui(prod(i, 33), ",", ".", UM_A_UM)
'        prod(i, 34) = Format(Round((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!PISAliquota / 100), 2), "###0.00")
'
'        Select Case prod(i, 31)
'           Case "04", "06", "07", "08", "09"
'              vlPIS = vlPIS
'           Case Else
'              vlPIS = vlPIS + prod(i, 34)
'        End Select
'
'        prod(i, 34) = Substitui(prod(i, 34), ",", ".", UM_A_UM)
'        prod(i, 45) = "0.00"
'
'        'tag PISST
'        prod(i, 54) = ""
'        prod(i, 55) = ""
'        prod(i, 56) = ""
'
'        '
'        '========dados do COFINS (grupo s do Manual de Integração - páginas 113) ============
'        '
'        prod(i, 35) = IIf(Vazio(NFeItens!COFINSCST) Or IsNull(NFeItens!COFINSCST), "07", NFeItens!COFINSCST)
'        prod(i, 36) = Format((NFeItens!QtdeMov * NFeItens!ValorUnit), "######0.00")
'        prod(i, 36) = Substitui(prod(i, 36), ",", ".", UM_A_UM)
'        prod(i, 37) = Format(Parametros!COFINSAliquota, "###0.00")
'        prod(i, 37) = Substitui(prod(i, 37), ",", ".", UM_A_UM)
'        prod(i, 38) = Format(Round((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!COFINSAliquota / 100), 2), "###0.00")
'
'        Select Case prod(i, 35)
'           Case "04", "06", "07", "08", "09"
'              vlCOFINS = vlCOFINS
'           Case Else
'              vlCOFINS = vlCOFINS + prod(i, 38)
'        End Select
'
'        prod(i, 38) = Substitui(prod(i, 38), ",", ".", UM_A_UM)
'        prod(i, 44) = "0.00"
'
'
'        'tag COFINSST
'        prod(i, 57) = ""
'        prod(i, 58) = ""
'        prod(i, 59) = ""
'
'        'Tag da Declaração de Importação
'        If ModeloNF = "55" Then
'            SQL = "SELECT IdNFProd_Item_Seq, DI_Numero, DI_Data, DI_UF_Desembarque, DI_Local_Desembarque, " & _
'                  "DI_Data_Desembarque, DI_Codigo_Exportador " & _
'                  "FROM TbNotaFiscalProd_Itens_DI " & _
'                  "WHERE IdNFProd = " & NFeItens!IdNFProd & " AND IdNFProd_Item = " & NFeItens!IdNFProd_Item & " " & _
'                  "ORDER BY IdNFProd_Item_Seq"
'            Set NFeDeclaracaoImposto = vgDb.OpenRecordset(SQL)
'            If NFeDeclaracaoImposto.RecordCount > 0 Then
'                prod(i, 60) = NFeDeclaracaoImposto!DI_Numero                                  'nDI
'                prod(i, 61) = Format(NFeDeclaracaoImposto!DI_Data, "yyyy-mm-dd")              'dDI
'                prod(i, 62) = NFeDeclaracaoImposto!DI_Local_Desembarque                       'xLocDesemb
'                prod(i, 63) = NFeDeclaracaoImposto!DI_UF_Desembarque                          'UFDesemb
'                prod(i, 64) = Format(NFeDeclaracaoImposto!DI_Data_Desembarque, "yyyy-mm-dd")  'dDesemb
'                prod(i, 65) = NFeDeclaracaoImposto!DI_Codigo_Exportador                       'cExportador
'
'                SQL = "SELECT IdNFProd_Item_Seq_Item, AD_Numero, AD_Fabricante, AD_Desconto " & _
'                      "FROM TbNotaFiscalProd_Itens_DI_ADI " & _
'                      "WHERE IdNFProd = " & NFeItens!IdNFProd & " And IdNFProd_Item = " & NFeItens!IdNFProd_Item & " And IdNFProd_Item_Seq = " & NFeDeclaracaoImposto!IdNFProd_Item_Seq & " " & _
'                      "ORDER BY IdNFProd_Item_Seq_Item"
'                Set NFeAdicao = vgDb.OpenRecordset(SQL)
'                If NFeAdicao.RecordCount > 0 Then
'                    prod(i, 66) = NFeAdicao!AD_Numero                        'adi: nAdicao
'                    prod(i, 67) = NFeAdicao!IdNFProd_Item_Seq_Item           'adi: nSeqAdic
'                    prod(i, 68) = NFeAdicao!AD_Fabricante                    'adi: cFabricante
'                    If NFeAdicao!AD_Desconto > 0 Then
'                       prod(i, 69) = Format(NFeAdicao!AD_Desconto, "#0.00")  'adi: vDescDI
'                       prod(i, 69) = Substitui(prod(i, 69), ",", ".", UM_A_UM)
'                    End If
'                    NFeAdicao.MoveNext
'                End If
'                NFeDeclaracaoImposto.MoveNext
'            End If
'        End If
'
'        NFeItens.MoveNext
'    Next
'    'MsgBox "Grupo de Tributos/COFINS " + COFINS, vbInformation, "Resultado"
'    '
'    '   atualização de total
'    '
'    ReDim tot(39)
'    tot(0) = Format(NFe!BaseCalc_ICMS, "######0.00")
'    tot(0) = Substitui(tot(0), ",", ".", UM_A_UM)
'    tot(1) = Format(NFe!Valor_ICMS, "######0.00")
'    tot(1) = Substitui(tot(1), ",", ".", UM_A_UM)
'    tot(2) = Format(NFe!BaseCalc_ICSM_Subst, "######0.00")
'    tot(2) = Substitui(tot(2), ",", ".", UM_A_UM)
'    tot(3) = Format(NFe!Valor_ICMS_Subst, "######0.00")
'    tot(3) = Substitui(tot(3), ",", ".", UM_A_UM)
'    tot(4) = Format(NFe!Valor_NF_Prod, "######0.00")
'    tot(4) = Substitui(tot(4), ",", ".", UM_A_UM)
'    tot(5) = Format(NFe!Valor_Frete, "######0.00")
'    tot(5) = Substitui(tot(5), ",", ".", UM_A_UM)
'    tot(6) = Format(NFe!Valor_Seguro, "######0.00")
'    tot(6) = Substitui(tot(6), ",", ".", UM_A_UM)
'    tot(7) = Format(NFe!DescontoPromocional, "######0.00")
'    tot(7) = Substitui(tot(7), ",", ".", UM_A_UM)
'    tot(8) = Format(NFe!ValorImpostoImportacao, "######0.00")
'    tot(8) = Substitui(tot(8), ",", ".", UM_A_UM)
'    tot(9) = Format(NFe!Valor_IPI, "######0.00")
'    tot(9) = Substitui(tot(9), ",", ".", UM_A_UM)
'    tot(10) = Format(vlPIS, "######0.00")
'    tot(10) = Substitui(tot(10), ",", ".", UM_A_UM)
'    tot(11) = Format(vlCOFINS, "######0.00")
'    tot(11) = Substitui(tot(11), ",", ".", UM_A_UM)
'    tot(12) = Format(NFe!OutrasDespesasAces, "######0.00")
'    tot(12) = Substitui(tot(12), ",", ".", UM_A_UM)
'    vlNF = (NFe!Valor_NF_Prod - NFe!DescontoPromocional) + (NFe!Valor_Frete + NFe!Valor_Seguro + NFe!Valor_IPI + NFe!OutrasDespesasAces + NFe!ValorImpostoImportacao + NFe!Valor_ICMS_Subst)
'    tot(13) = Format((NFe!Valor_NF_Prod - NFe!DescontoPromocional) + (NFe!Valor_Frete + NFe!Valor_Seguro + NFe!Valor_IPI + NFe!OutrasDespesasAces + NFe!ValorImpostoImportacao + NFe!Valor_ICMS_Subst), "######0.00")
'    tot(13) = Substitui(tot(13), ",", ".", UM_A_UM)
'    If vlTrib > 0 Then tot(19) = Substitui(Format(vlTrib, "#0.00"), ",", ".", UM_A_UM)
'    tot(20) = "0.00"    'ICMSTot <vICMSDesn>
'    tot(37) = "0.00"    'ICMSTot <vFCPUFDest>
'    tot(38) = "0.00"    'ICMSTot <vICMSUFDest>
'    tot(39) = "0.00"    'ICMSTot <vICMSUFRemet>
'
'    'grupo ISSQN
'    tot(14) = ""    'ISSQNtot <vServ>
'    tot(15) = ""    'ISSQNtot <vBC>
'    tot(16) = ""    'ISSQNtot <vISS>
'    tot(17) = ""    'ISSQNtot <vPIS>
'    tot(18) = ""    'ISSQNtot <vCOFINS>
'
'    '
'    '============dados do transportador
'    '
'    ReDim trp(16)
'    trp(0) = IIf(Vazio(NFe!Frete_Por_Conta), 0, Left(NFe!Frete_Por_Conta, 1))        ' responsabilidade do frete 0-emitente, 1-destinatário
'    If Not Vazio(NFe!NomeTrasnportador) Then
'        If Len(Retira(NFe!CPF_CNPJ_Transp, ".-/", UM_A_UM)) > 11 Then
'          trp(1) = Trim(NFe!CPF_CNPJ_Transp)                                         ' CNPJ da Transportadora sem mascara
'          trp(1) = Retira(trp(1), ".-/", UM_A_UM)                            ' CNPJ da Transportadora sem mascara
'        Else
'          trp(1) = Trim(NFe!CPF_CNPJ_Transp)                                          ' CPF da Transportadora sem mascara
'          trp(1) = Retira(trp(1), ".-/", UM_A_UM)                             ' CPF da Transportadora sem mascara
'        End If
'        trp(2) = RemoveAcento(Trim(NFe!NomeTrasnportador))
'        trp(3) = Trim(Retira(NFe!InscEst_Trasnp, ".,-/", UM_A_UM))                   ' Inscrição Estadual da Transportadora sem máscara
'        trp(4) = RemoveAcento(Trim(NFe!Endereco_Transp))
'        trp(5) = RemoveAcento(Trim(NFe!Cidade_Transp))
'        trp(6) = NFe!UF_Mot_Transp
'    End If
'
'    If Not Vazio(NFe!Placa_Veiculo) Then
'       trp(7) = Retira(Trim(NFe!Placa_Veiculo), "-", UM_A_UM)
'       trp(8) = NFe!UF_Trasnportador
'       trp(15) = ""
'    End If
'
'    '  ============== criação dos lacres do volume
'    '
'    If NFe!Qtde_Trasnp > 0 Then
'       trp(9) = NFe!Qtde_Trasnp                                         ' quantidade de volumes
'       trp(9) = Substitui(trp(9), ",", ".", UM_A_UM)
'       trp(10) = RemoveAcento(NFe!Especie_Transp)                       ' espécie dos volumes
'       trp(11) = RemoveAcento(NFe!Marca_Trasnp)                         ' marca dos volumes
'       trp(12) = NFe!Num_Transp                                         ' numeração dos volumes
'       trp(13) = Format(NFe!PesoLiq_Transp, "#0.000")                   ' peso líquido
'       trp(13) = Substitui(trp(13), ",", ".", UM_A_UM)
'       trp(14) = Format(NFe!PesoBruto_Transp, "#0.000")                 ' peso bruto
'       trp(14) = Substitui(trp(14), ",", ".", UM_A_UM)
'    End If
'
'    vsSQL = "SELECT Count(IDParcela) as qt FROM TbNFCe_Faturas WHERE idNFProd = " & NumeroNota
'    cob_numero_parcelas = SQLExecutaRetorno(vsSQL, "qt", 0) - 1
'
'    Dim idparc As Integer, vTotalRecebido As Double, vTotalNF As Double, vTotalDinheiro As Double, vTotalOutras As Double
'    idparc = 0
'    If cob_numero_parcelas >= 0 Then
'       ReDim cob(cob_numero_parcelas, 12)
'       cob(0, 0) = ide(6)
'       cob(0, 1) = tot(13)
'       cob(0, 2) = tot(13)
'       vsSQL = "SELECT IDParcela, Vencimento, Valor, TipoPgto, IdBandeira, CartaoNumeroAutorizacao " & _
'               "FROM TbNFCe_Faturas " & _
'               "WHERE idNFProd = " & NumeroNota
'       vTotalRecebido = SQLExecutaRetorno("SELECT SUM(Valor) r FROM TbNFCe_Faturas WHERE IdNFProd = " & NumeroNota, "r", 0)
'       vTotalDinheiro = SQLExecutaRetorno("SELECT SUM(Valor) r FROM TbNFCe_Faturas WHERE IdNFProd = " & NumeroNota & " AND TipoPgto = 'DH'", "r", 0)
'       vTotalOutras = SQLExecutaRetorno("SELECT SUM(Valor) r FROM TbNFCe_Faturas WHERE IdNFProd = " & NumeroNota & " AND TipoPgto <> 'DH'", "r", 0)
'       RsOpen NFeParcelas, vsSQL
'       'Set NFeParcelas = vgDb.OpenRecordset(vsSQL)
'       vTotalNF = vlNF
'       Do While Not NFeParcelas.EOF
'          '01 - Dinheiro|02 - Cheque|03 - Cartão de Crédito|04 - Cartão de Débito|05 - Crédito Loja|10 - Vale Alimentação|11 - Vale Refeição|12 - Vale Presente|13 - Vale Combustível|99 - Outros
'          Select Case NFeParcelas!TipoPgto
'                 Case "DH": cob(idparc, 6) = "01"                            'pag <tPag>
'                 Case "CH": cob(idparc, 6) = "02"                            'pag <tPag>
'                 Case "CC": cob(idparc, 6) = "03"                            'pag <tPag>
'                 Case "CD": cob(idparc, 6) = "04"                            'pag <tPag>
'                 Case "CT": cob(idparc, 6) = "05"                            'pag <tPag>
'                 Case Else: cob(idparc, 6) = "99"                                 'pag <tPag>
'          End Select
'          cob(idparc, 7) = Format(NFeParcelas!Valor, "######0.00")        'pag <vPag>
'          cob(idparc, 7) = Substitui(cob(idparc, 7), ",", ".", UM_A_UM)   'pag <vPag>
'          If cob(idparc, 6) = "03" Or cob(idparc, 6) = "04" Then
'             cob(idparc, 8) = Retira(Parametros!CNPJ, ".-/ ", UM_A_UM)    'card <CNPJ>  Informar o CNPJ da Credenciadora de cartão de crédito / débito
'             cob(idparc, 9) = Left(NFeParcelas!CartaoBandeira, 2)         'card <tBand> 01 - Visa|02 - Mastercard|03 - American Express|04 - Sorocred|99 - Outros
'             cob(idparc, 10) = Trim(NFeParcelas!CartaoNumeroAutorizacao)  'card <cAut>  Identifica o número da autorização da transação da operação com cartão de crédito e/ou débito
'             cob(idparc, 11) = "2"                       'card <tpIntegra> Tipo de Integração do processo de pagamento com o sistema de automação da empresa:
'                                                         '     1 - Pagamento integrado com o sistema de automação da empresa (Ex.: equipamento TEF, Comércio Eletrônico);
'                                                         '     2 - Pagamento não integrado com o sistema de automação da empresa (Ex.: equipamento POS);
'          End If
'          idparc = idparc + 1
'          NFeParcelas.MoveNext
'       Loop
'       If NFeParcelas.RecordCount = 0 Then
'          cob(0, 6) = "01"            'pag <tPag>
'          cob(0, 7) = tot(13)         'pag <vPag>
'       End If
'    Else
'       ReDim cob(0, 11)
'       cob_numero_parcelas = 0
'       cob(0, 6) = "01"            'pag <tPag>
'       cob(0, 7) = tot(13)         'pag <vPag>
'    End If
'
'    '
'    '============= informações adcionais
'    '
'    ReDim obs(10)
'    obs(0) = ""   'RemoveAcento(Trim(NFe!InformacoesAdicionais))
'    obs(1) = RemoveAcento(Trim(NFe!Linha1)) & " // " & RemoveAcento(Trim(NFe!Linha2)) & " // " & RemoveAcento(Trim(NFe!Linha3)) & " // " & RemoveAcento(Trim(NFe!Linha4)) & " // " & RemoveAcento(Trim(NFe!Linha5))
'    If vlTrib > 0 And vlNF > 0 Then
'       vlNF = (NFe!Valor_NF_Prod - NFe!DescontoPromocional) + (NFe!Valor_Frete + NFe!Valor_Seguro + NFe!Valor_IPI + NFe!OutrasDespesasAces + NFe!ValorImpostoImportacao)
'       pTributos = Format((vlTrib / vlNF) * 100, "#0.00")
'       obs(1) = obs(1) & " - Valor Aproximado dos Tributos R$ " & FormatoDecimal(Format(vlTrib, "#0.00")) & " (" & FormatoDecimal(pTributos) & "%) (Conforme Lei Fed. 12.741/2012) Fonte: IBPT"
'    End If
'
'    'tag exporta v2.03
'    obs(2) = ""      'UFEmbarq
'    obs(3) = ""      'xLocEmbarq
'
'    'tag compra v2.03
'    obs(4) = ""      'xNEmp
'    obs(5) = ""      'xPed
'    obs(6) = ""      'infCpl
'
'Else
'   GoTo caiFora
'End If
'
'dirXML = Parametros!DiretorioXML
'dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
'
''pega o endereço do arquivo a ser gerado
'If Not Existe(dirXML) Then MkDir dirXML
''gera a chave da nfe
'Dim id_chave As String
'Dim numero_nfe_gerado As String
'numero_nfe_gerado = sistNFCe.GeraXML(ide(), emit(), dest(), prod(), tot(), trp(), cob(), obs(), autXML(), False)
'id_chave = numero_nfe_gerado
'numero_nfe_gerado = Replace(numero_nfe_gerado, "NFe", "")
'NFeChaveAcesso = numero_nfe_gerado
'
'If Not Vazio(NFeChaveAcesso) Then
'   vsSQL = "UPDATE TbNFCe SET " & _
'           "NFCeChaveAcesso = '" & NFeChaveAcesso & "' " & _
'           "WHERE IdNFProd = " & NumeroNota
'   vgDb.Execute vsSQL
'End If
'
'xCaminhoXML = dirXML & "\nfe\arquivos\" & id_chave & ".xml"
'
'NFeResposta = sistNFCe.AssinarArquivoXML(xCaminhoXML, "infNFe")
'
'DoEvents
'
'If InStr(NFeResposta, "Erro") > 0 Then
'   NFeMotivo = NFeResposta
'   GoTo NaoEnviou
'End If
'
'xCaminhoXML = dirXML & "\nfe\arquivos\assinado\" & id_chave & "-assinado.xml"
'
'XMLOK = sistNFCe.ValidarArquivoXML(xCaminhoXML, False, NFeValidate) 'PL_007a.zip
'
'If Not XMLOK Then
'   NFeMotivo = "Erro na Validação do XML, falha no Schema" & vbNewLine & NFeValidate
'   GoTo caiFora
'End If
'
'xCaminhoXML = dirXML & "\nfe\arquivos\gerados\" & id_chave & ".xml"    'dirXML & "nfe\lotes\" & LPad(ide(6), 12, "0") & "-env-lot.xml"
'
'If Not PodeEnviar Then GoTo NaoEnviou
'
'NFeResposta = sistNFCe.NfeAutorizacao(xCaminhoXML, True)
'
'If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
'   MsgBox "*** Aparentemente Ocorreram Erros na Recepção do Lote (nfeAutorizacao)***" _
'   & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeAutorizacao"
'   GoTo caiFora
'End If
'
'NFeMotivo = Parse(NFeResposta, "#")
'If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo caiFora
'NFeNumeroRecibo = Parse(NFeResposta, "#")
'cStat = Parse(NFeResposta, "#")
'
'If cStat <> 103 Then
'   GoTo NaoEnviou
'End If
'
'vsSQL = "UPDATE TbNFCe SET " & _
'        "NFCeRecibo = " & NFeNumeroRecibo & " " & _
'        "WHERE IdNFProd = " & NumeroNota
'vgDb.Execute vsSQL
'
'DoEvents
'
'consultaNFe:
'
'   NFeResposta = sistNFCe.NfceRetAutorizacao(NFeNumeroRecibo, dirXML)
'
'   If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
'      MsgBox NFeResposta, vbExclamation, "Retorno Autorização"
'      GoTo caiFora
'   End If
'
'   If InStr(NFeResposta, "217") > 0 Then
'      Sleep 3000 ' Aguarda mais 3 segundos
'      NFeResposta = sistNFCe.NfceRetAutorizacao(NFeNumeroRecibo, dirXML) 'refaz a consulta
'      If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
'         MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote (nfeRetAutorizacao)***" _
'         & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeRetAutorizacao"
'         GoTo caiFora
'     End If
'   End If
'
'buscaNFe:
'
''Consulta Nfe
'   NFeResposta = sistNFCe.NfceConsulta(NFeChaveAcesso)
'
'   Dim NFeRespostaFinal As String
'   NFeRespostaFinal = NFeResposta
'
'   ' testa erro 217 Rejeição: NF-e não consta na base de dados da SEFAZ <dhRecbto xmlns="http://www.portalfiscal.inf.br/nfe">2015-03-29T09:56:36-03:00
'   If InStr(NFeResposta, "217") > 0 Then
'      Sleep 3000 ' aguarda mais 3 segundos
'      NFeResposta = sistNFCe.NfceConsulta(NFeChaveAcesso) ' consulta novamente
'
'      If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
'         MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote ***" _
'         & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeConsulta"
'         GoTo caiFora
'      End If
'   End If
'
'   If InStr(NFeResposta, "Erro 98") > 0 Then
'      NFeMotivo = Parse(NFeResposta, "#")
'      NFeNumeroProtocolo = ""
'      GoTo caiFora
'   End If
'
'   On Error Resume Next
'    cStat = Parse(NFeResposta, "#")
'    NFeMotivo = Parse(NFeResposta, "#")
'    NFeDataHora = Parse(NFeResposta, "#")
'    If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
'    If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
'    NFeNumeroProtocolo = Parse(NFeResposta, "#")
'    xCaminhoXML = NFeResposta
'    nroRecibo = cStat
'
'    If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo caiFora
'
'    If nroRecibo = 204 Or nroRecibo = 539 Then
'       NFeChaveAcesso = Mid(nfeRetorno, InStr(nfeRetorno, "chNFe:") + 6, 44)
'       nroRecibo = Mid(nfeRetorno, InStr(nfeRetorno, "nRec:") + 5)
'       nroRecibo = Left(nroRecibo, Len(nroRecibo) - 1)
'       NFeNumeroRecibo = Left(NFeNumeroRecibo, 15 - Len(nroRecibo)) + nroRecibo
'       If Vazio(NFeChaveAcesso) Or Len(NFeNumeroRecibo) < 15 Then
'          NFeMotivo = nfeRetorno
'          GoTo caiFora
'       End If
'       vsSQL = "UPDATE TbNFCe SET " & _
'               "NFCeChaveAcesso = '" & NFeChaveAcesso & "' " & _
'               "NFCeRecibo = " & NFeNumeroRecibo & " " & _
'               "WHERE IdNFProd = " & NumeroNota
'       vgDb.Execute vsSQL
'       GoTo buscaNFe
'    ElseIf nroRecibo <> 100 And nroRecibo <> 301 Then
'        NFeMotivo = nroRecibo + " - " + nfeRetorno
'        GoTo NaoEnviou
'    ElseIf nroRecibo = 105 And nroRecibo = 217 Then
'        GoTo consultaNFe
'    ElseIf nroRecibo = 100 Then
'        nfeRetorno = "Nota Fiscal de Consumidor Eletronica Autorizado o Uso."
'        NFeDataHora = Format(Now, "dd/mm/yyyy h:mm:ss")
'        msgResultado = "Chave NF-e.: " + NFeChaveAcesso & vbCrLf
'        msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
'        msgResultado = msgResultado + "Recibo.: " + NFeNumeroRecibo & vbCrLf
'        msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
'        msgResultado = msgResultado + "Resposta da Fazenda.: " + nroRecibo + " - " & nfeRetorno
'
'        ' mensagem de emissao  MsgBox msgResultado, vbInformation + vbOKOnly
'
'        On Error Resume Next
'        NFeResposta = sistNFCe.NfceConsulta(NFeChaveAcesso)
'
'        vsSQL = "UPDATE TbNFCe SET " & _
'                "NFCeEnviada = 1, " & _
'                "NFCeChaveAcesso = '" & NFeChaveAcesso & "', " & _
'                "NFCeProtocolo = " & NFeNumeroProtocolo & ", " & _
'                "NFCeProtocoloDataHora = '" & NFeDataHora & "' " & _
'                "WHERE IdNFProd = " & NumeroNota
'        vgDb.Execute vsSQL
'    End If
'
'    'xCaminhoXML = dirXML & "\arquivos\procNFe\" & id_chave & "-procNFe.xml"
'    Dim xmlPathPDF As String
'    Dim anoEmes As String
'    Dim Arquivo As String
'    anoEmes = dirXML & "\nfe\arquivos\procNFe\" & Mid(ide(7), 1, 4) & Mid(ide(7), 6, 2) & "\"
'    xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
'    If Not Existe(xCaminhoXML) Then xCaminhoXML = anoEmes & NFeChaveAcesso & "-procNFe.xml"             '  Aqui Gera o DANFE
'
'PodeSair:
'Set sistNFCe = Nothing
'Screen.MousePointer = vbDefault
'TransmitirNFCe = True
'Exit Function
'
'NaoEnviou:
'Set sistNFCe = Nothing
'
'If PodeEnviar Then MsgBox NFeMotivo, vbCritical + vbOKOnly
'Screen.MousePointer = vbDefault
'Exit Function
'Resume
'
'caiFora:
'If Not Vazio(NFeMotivo) Then MsgBox NFeMotivo, vbCritical + vbOKOnly
'If Not Vazio(NFeResposta) Then MsgBox NFeResposta, vbCritical + vbOKOnly
'
'Set sistNFCe = Nothing
'
'Screen.MousePointer = vbDefault
'TransmitirNFCe = False
'
'On Error GoTo 0
'Exit Function
'
'Resume
'TransmitirNFCe_Error:
'
'    Screen.MousePointer = vbDefault
'    MsgBox "Falha (" & Err.Description & ")" & vbNewLine & "Em TransmitirNFCe no Módulo NFe_DLL", vbCritical, "Falha"
'    Err.Clear
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
   
   iRetorno = sistNFe.CancelarNFe(CNPJ, IdLote, 1, ChaveAcesso, Protocolo, Justificativa)
 
   If Not iRetorno Then GoTo caiFora
   
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
         MsgBox Str$(cStat2) & " - " & NFeValidate, vbInformation, "Cancelar NFe"
      Else
         MsgBox Str$(cStat) & " - " & NFeMotivo, vbInformation, "Cancelar NFe"
      End If
      GoTo caiFora
   End If
   
continua:
   msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + Str(cStat2) & " - " & NFeValidate
   
   MsgBox msgResultado, vbInformation + vbOKOnly, "Cancelar NFe"

   If GravaProtocolo Then
      vsSQL = "INSERT INTO NotaFiscalRecibos (CodigoNota, NumeroProtocolo, DataHora) Values " & _
              "(" & vsNumeroNota & ", " & NFeNumeroProtocolo & ", '" & NFeDataHora & "')"
      vgDb.Execute vsSQL, True
   End If

   Screen.MousePointer = vbDefault
   CancelaNFe = True
   Exit Function

caiFora:
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
   
   iRetorno = sistNFCe.CancelarNFe(CNPJ, IdLote, 1, ChaveAcesso, Protocolo, Justificativa)
 
   If Not iRetorno Then GoTo caiFora
   
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
     GoTo caiFora
   End If
   
continua:
   msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + Str(cStat2) & " - " & NFeValidate
   
   'msgResultado = NFeResposta
   
   MsgBox msgResultado, vbInformation + vbOKOnly

   Screen.MousePointer = vbDefault
   CancelaNFCe = True
   Exit Function

caiFora:
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
     GoTo caiFora
  End If
   
  cStat = sistNFe.retConsRec.cStat
  NFeMotivo = sistNFe.retConsRec.xMotivo
  
  If cStat = 105 Or cStat = 217 Then
     i = i + 1
     If i > 5 And cStat = 105 Then
        msgResultado = "A Nota Fiscal ficou em PROCESSAMENTO." & vbNewLine & "Efetue a consulta do Recibo daqui 5 minutos novamente."
        NFeValidate = "NFe/NFCe PROCESSAMENTO"
        GoTo caiFora
     End If
     Sleep 10000
     GoTo buscaNFe
  ElseIf cStat = 106 Then
     msgResultado = Str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe NÃO LOCALIZADA"
     iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
    
     If Not iRetorno Then
        NFeValidate = "ERRO"
        NFeNumeroProtocolo = ""
        GoTo caiFora
     End If
     cStat = sistNFe.retConsulta.cStat
     NFeMotivo = sistNFe.retConsulta.xMotivo
     If cStat = 613 Then
        NFeChaveAcesso = Mid(NFeMotivo, InStr(NFeMotivo, "Numerico da NF-e [") + 18, 44)
        iRetorno = sistNFe.ConsultarProtocolo(NFeChaveAcesso)
        cStat = sistNFe.retConsulta.cStat
        NFeMotivo = sistNFe.retConsulta.xMotivo
        
        If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
           NFeDataHora = sistNFe.retConsulta.protNFe.infProt.dhRecbto
           NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
           ChaveAcesso = sistNFe.retConsulta.protNFe.infProt.chNFe
           GoTo continuaConsulta
        Else
           GoTo caiFora
        End If
     Else
        GoTo caiFora
     End If
  ElseIf cStat = 239 Then
     msgResultado = Str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe COM PROBLEMAS"
     GoTo caiFora
  End If
  NFeNumeroProtocolo = sistNFe.retConsRec.protNFe.infProt.nProt
  NFeDataHora = sistNFe.retConsRec.protNFe.infProt.dhRecbto
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
     GoTo caiFora
  ElseIf nroRecibo = 105 Then
     GoTo buscaNFe
  End If

  iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
  
  If Not iRetorno Then
     NFeValidate = "ERRO"
     NFeNumeroProtocolo = ""
     GoTo caiFora
  End If

  cStat = sistNFe.retConsulta.cStat
  NFeMotivo = sistNFe.retConsulta.xMotivo
  If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
     NFeDataHora = sistNFe.retConsulta.protNFe.infProt.dhRecbto
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
  msgResultado = msgResultado + "Resposta da Fazenda.: " + Str(cStat) + " - " & NFeMotivo
  
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
   
caiFora:
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
     GoTo caiFora
  End If
   
  cStat = sistNFe.retConsRec.cStat
  NFeMotivo = sistNFe.retConsRec.xMotivo
  
  If cStat = 105 Or cStat = 217 Then
     i = i + 1
     If i > 5 And cStat = 105 Then
        msgResultado = "A Nota Fiscal ficou em PROCESSAMENTO." & vbNewLine & "Efetue a consulta do Recibo daqui 5 minutos novamente."
        NFeValidate = "NFe/NFCe PROCESSAMENTO"
        GoTo caiFora
     End If
     Sleep 10000
     GoTo buscaNFe
  ElseIf cStat = 106 Then
     msgResultado = Str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe NÃO LOCALIZADA"
     iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
    
     If Not iRetorno Then
        NFeValidate = "ERRO"
        NFeNumeroProtocolo = ""
        GoTo caiFora
     End If
     cStat = sistNFe.retConsulta.cStat
     NFeMotivo = sistNFe.retConsulta.xMotivo
     If cStat = 613 Then
        NFeChaveAcesso = Mid(NFeMotivo, InStr(NFeMotivo, "Numerico da NF-e [") + 18, 44)
        iRetorno = sistNFe.ConsultarProtocolo(NFeChaveAcesso)
        cStat = sistNFe.retConsulta.cStat
        NFeMotivo = sistNFe.retConsulta.xMotivo
        
        If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
           NFeDataHora = sistNFe.retConsulta.protNFe.infProt.dhRecbto
           NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
           ChaveAcesso = sistNFe.retConsulta.protNFe.infProt.chNFe
           GoTo continuaConsulta
        Else
           GoTo caiFora
        End If
     Else
        GoTo caiFora
     End If
  ElseIf cStat = 239 Then
     msgResultado = Str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe COM PROBLEMAS"
     GoTo caiFora
  End If
  NFeNumeroProtocolo = sistNFe.retConsRec.protNFe.infProt.nProt
  NFeDataHora = sistNFe.retConsRec.protNFe.infProt.dhRecbto
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
     GoTo caiFora
  ElseIf nroRecibo = 105 Then
     GoTo buscaNFe
  End If

  iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
  
  If Not iRetorno Then
     NFeValidate = "ERRO"
     NFeNumeroProtocolo = ""
     GoTo caiFora
  End If

  cStat = sistNFe.retConsulta.cStat
  NFeMotivo = sistNFe.retConsulta.xMotivo
  If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
     NFeDataHora = sistNFe.retConsulta.protNFe.infProt.dhRecbto
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
  msgResultado = msgResultado + "Resposta da Fazenda.: " + Str$(cStat) + " - " & NFeMotivo
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
   
caiFora:
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
      NFeDataHora = sistNFe.retConsulta.protNFe.infProt.dhRecbto
      NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
   Else
      NFeDataHora = ""
      NFeNumeroProtocolo = ""
   End If

   msgResultado = "Chave NF-e.: " & ChaveAcesso & vbCrLf
   msgResultado = msgResultado + "Protocolo.: " & NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data e Hora.: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + Str$(cStat) & " - " & NFeMotivo
   
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

   GoTo caiFora
   
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
   Err.Clear
   cStat = 0
   NFeNumeroProtocolo = ""
   NFeDataHora = ""
   NFeMotivo = ""
   msgResultado = ""
   NFeValidate = ""

caiFora:
   Set sistNFe = Nothing
   Screen.MousePointer = vbDefault
End Sub

Public Sub consultaNFCe(ChaveAcesso As Variant, Optional NaoMostraMSG As Boolean) 'Sub que faz a consulta da NFe na Receita adaptada para nfe 3.1 ass 668

   On Error GoTo deuErro

   Screen.MousePointer = vbHourglass

   Dim sistNFe As snfe.Util
   Set sistNFe = New snfe.Util
   
   iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)

   iRetorno = sistNFe.ConsultarProtocolo(ChaveAcesso)
   
   cStat = sistNFe.retConsulta.cStat
   NFeMotivo = sistNFe.retConsulta.xMotivo
   If cStat = 100 Or cStat = 101 Or cStat = 110 Or cStat = 150 Then
      NFeDataHora = sistNFe.retConsulta.protNFe.infProt.dhRecbto
      NFeNumeroProtocolo = sistNFe.retConsulta.protNFe.infProt.nProt
   Else
      NFeDataHora = ""
      NFeNumeroProtocolo = ""
   End If
   
   msgResultado = "Chave NF-e.: " & ChaveAcesso & vbCrLf
   msgResultado = msgResultado + "Protocolo.: " & NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data e Hora.: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + Str$(cStat) & " - " & NFeMotivo
   
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

   GoTo caiFora
   
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
   Err.Clear
   cStat = 0
   NFeNumeroProtocolo = ""
   NFeDataHora = ""
   NFeMotivo = ""
   msgResultado = ""
   NFeValidate = ""

caiFora:
   Set sistNFe = Nothing
   Screen.MousePointer = vbDefault
End Sub

Public Sub ConsultaStatus(Optional ModeloNF As Integer = 55)  'Sub que consulta o Status do Serviço da Receita
Dim sistNFe As snfe.Util
   
   On Error GoTo deuErro
   
   Set sistNFe = New snfe.Util

   Screen.MousePointer = vbHourglass
   
   iRetorno = ConfiguraDLLNFeNFCe(ModeloNF, "1", sistNFe)

   'NFe
   If ModeloNF = 55 Then
      Call sistNFe.ConsultarStatusServico
      NFeResposta = CStr(sistNFe.retStatusWS.cStat) + " - " + sistNFe.retStatusWS.xMotivo
    End If
   'NFCe
   If ModeloNF = 65 Then
      Call sistNFe.ConsultarStatusServico
      NFeResposta = CStr(sistNFe.retStatusWS.cStat) + " - " + sistNFe.retStatusWS.xMotivo
   End If

   MsgBox "CONSULTA DE STATUS DO WS" & vbNewLine & vbNewLine & NFeResposta, vbInformation + vbOKOnly

   Set sistNFe = Nothing

   Screen.MousePointer = vbDefault
   
   Exit Sub
   
deuErro:
   MsgBox Err.Description, vbCritical
   Err.Clear
   Set sistNFe = Nothing

   Screen.MousePointer = vbDefault
End Sub

Public Function TransmitirCCe(ChaveAcesso As Variant, Data As Variant, nProtocolo As Variant, SeqCorrecao As Variant, TextoCorrecao As Variant) As Boolean  'Função para envio da carta de correção da NFe
Dim IdLote As Long, dhEvento As String, CNPJ As String
Dim dirXML As String
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

  Screen.MousePointer = vbHourglass
  
  IdLote = Int(Rnd(918274 * 999) * 1000)
  dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
  CNPJ = SQLExecutaRetorno("SELECT CNPJ FROM Empresa", "CNPJ", "")
  
  iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
  
  iRetorno = sistNFe.CartaCorrecao(CNPJ, IdLote, SeqCorrecao, ChaveAcesso, TextoCorrecao)
     
  If Not iRetorno Then
     NFeNumeroProtocolo = ""
     GoTo caiFora
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
      GoTo caiFora
   End If
   
continua:
   msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + Str(cStat) & " - " & NFeMotivo
   
   MsgBox msgResultado, vbInformation + vbOKOnly, "ENVIO CCe"
  
   Screen.MousePointer = vbDefault
   Set sistNFe = Nothing
   TransmitirCCe = True
   Exit Function

caiFora:
   Set sistNFe = Nothing
 
   Screen.MousePointer = vbDefault
   TransmitirCCe = False
End Function

'Fornecedor!CNPJCPF, ChaveAcesso, DataHora, TipoEvento, Justificativa
Public Function TransmitirManDest(CNPJ As Variant, ChaveAcesso As Variant, Data As Variant, TipoEvento As Variant, Justificativa As Variant) As Boolean  'Função para envio da carta de correção da NFe
Dim IdLote As Long, dhEvento As String
Dim dirXML As String
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

  Screen.MousePointer = vbHourglass
  
  IdLote = Int(Rnd(918274 * 999) * 1000)
  dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
  
  iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
  
  iRetorno = sistNFe.ManifestacaoDestinatario(TipoEvento, IdLote, CNPJ, ChaveAcesso, dhEvento, Justificativa, 1, NFeResposta)
  
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
  
  If cStat = 135 Then
     GoTo continua
  Else
     MsgBox cStat & " - " & NFeMotivo, vbInformation, "ERRO"
     GoTo caiFora
  End If
         
continua:
  msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
  msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
  msgResultado = msgResultado + "Resposta da Fazenda.: " + Str(cStat) & " - " & NFeMotivo
    
  MsgBox msgResultado, vbInformation + vbOKOnly, "Envio Manifestação do Destinatário"
  
  Screen.MousePointer = vbDefault
  Set sistNFe = Nothing
  TransmitirManDest = True
  Exit Function

caiFora:
  Set sistNFe = Nothing
 
  Screen.MousePointer = vbDefault
  TransmitirManDest = False
End Function

Public Function RemoveAcento(sString As String) As String

    Dim sRet As String

    sRet = sString

    sRet = Replace(sRet, "<", " ")
    sRet = Replace(sRet, ">", " ")
    sRet = Replace(sRet, "&", "E")
    sRet = Replace(sRet, "'", " ")

    sRet = Replace(sRet, "á", "a")
    sRet = Replace(sRet, "à", "a")
    sRet = Replace(sRet, "â", "a")
    sRet = Replace(sRet, "ã", "a")
    sRet = Replace(sRet, "ä", "a")

    sRet = Replace(sRet, "é", "e")
    sRet = Replace(sRet, "è", "e")
    sRet = Replace(sRet, "ê", "e")
    sRet = Replace(sRet, "ë", "e")

    sRet = Replace(sRet, "í", "i")
    sRet = Replace(sRet, "ì", "i")
    sRet = Replace(sRet, "î", "i")
    sRet = Replace(sRet, "ï", "i")

    sRet = Replace(sRet, "ó", "o")
    sRet = Replace(sRet, "ò", "o")
    sRet = Replace(sRet, "ô", "o")
    sRet = Replace(sRet, "õ", "o")
    sRet = Replace(sRet, "ö", "o")

    sRet = Replace(sRet, "ú", "u")
    sRet = Replace(sRet, "ù", "u")
    sRet = Replace(sRet, "û", "u")
    sRet = Replace(sRet, "ü", "u")

    sRet = Replace(sRet, "ç", "c")

    sRet = Replace(sRet, "Á", "A")
    sRet = Replace(sRet, "À", "A")
    sRet = Replace(sRet, "Â", "A")
    sRet = Replace(sRet, "Ã", "A")
    sRet = Replace(sRet, "Ä", "A")

    sRet = Replace(sRet, "É", "E")
    sRet = Replace(sRet, "È", "E")
    sRet = Replace(sRet, "Ê", "E")
    sRet = Replace(sRet, "Ë", "E")

    sRet = Replace(sRet, "Í", "I")
    sRet = Replace(sRet, "Ì", "I")
    sRet = Replace(sRet, "Î", "I")
    sRet = Replace(sRet, "Ï", "I")

    sRet = Replace(sRet, "Ó", "O")
    sRet = Replace(sRet, "Ò", "O")
    sRet = Replace(sRet, "Ô", "O")
    sRet = Replace(sRet, "Õ", "O")
    sRet = Replace(sRet, "Ö", "O")

    sRet = Replace(sRet, "Ú", "U")
    sRet = Replace(sRet, "Ù", "U")
    sRet = Replace(sRet, "Û", "U")
    sRet = Replace(sRet, "Ü", "U")

    sRet = Replace(sRet, "Ç", "C")

    sRet = Replace(sRet, "°", ".")
    sRet = Replace(sRet, "º", ".")
    sRet = Replace(sRet, "ª", ".")
    
    sRet = Replace(sRet, Chr(13), " ")
    sRet = Replace(sRet, Chr(10), " ")
    sRet = Replace(sRet, vbNewLine, " ")
    sRet = Replace(sRet, "  ", " ")
    
    sRet = Replace(sRet, "§", "INCISO(S)")
    
    sRet = LTrim(sRet)
    sRet = RTrim(sRet)

    RemoveAcento = UCase(sRet)

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

Public Function DownloadXML(ChaveNFe As String) As Boolean
'instancia o componente
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

    dirXML = GetString(&H80000001, "nfe", "PathPrincipal")
    If VbInDesign Then
       dirXML = "C:\nfe-app\nfe-app"
       SaveString &H80000001, "nfe", "PathPrincipal", dirXML
    End If
    'dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
    
    iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
    
    iRetorno = sistNFe.DownloadXML(ChaveNFe, dirXML, NFeResposta, xCaminhoXML)
       
    If Not iRetorno Then GoTo caiFora
   
    DownloadXML = True
    
    Set sistNFe = Nothing
    
    Exit Function
caiFora:
    DownloadXML = False
    xCaminhoXML = ""
    Set sistNFe = Nothing
End Function
