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

Public ide() As String
Public emit() As String
Public dest() As String
Public prod() As String
Public tot() As String
Public trp() As String
Public cob() As String
Public cob_numero_parcelas As Integer
Public pagList() As String
Public obs() As String
Public autXML() As String
Public NFeXML As String, NFeValidate As String, NFeChaveAcesso As String, NFeChaveAcessoAdicional As String, NFecNF As String, NFeMotivo As String, NFeResposta As String, NFeNumeroRecibo As String, NFeNumeroProtocolo As String, NFeDataHora As String, NFeDataHoraEnvio As String
Public vsCertificado As String, msgResultado As String, nfeRetorno As String, iRetorno As Long, XMLOK As Boolean, cStat As Long, cStat2 As Long
Public cabMsg As String, DadosMsg As String, msgRetWS As String, Proxy As String, UsuarioProxy As String, SenhaProxy As String
Public xCaminhoXML As String, xCaminhoXMLAuxiliar As String, xCaminhoTXT As String, xCaminhoPDF As String, dirXML As String, vsSQL As String
Public vsNumeroNota As Variant
Dim nroRecibo As String, nroProtocolo As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function TransmitirNFe(ByVal NumeroNota As Variant, ByVal SerieNF As Variant, Optional PodeEnviar As Boolean = False) As Boolean  'Função que monta o arquivo XML e faz o envio para a Receita
 Dim txtNumerado As String
 Dim Retorno As String
 Dim vsNFe As String
 Dim Parametros As New ADODB.Recordset, Destinatario As New ADODB.Recordset, Produtos As New ADODB.Recordset
 Dim NFe As New ADODB.Recordset, NFeItens As New ADODB.Recordset, NFeParcelas As New ADODB.Recordset, NFeOBS As New ADODB.Recordset, NFeAutorizados As New ADODB.Recordset, NFeReferenciadas As New ADODB.Recordset
 Dim NFeMedicamentos As New ADODB.Recordset, NFeArmamento As New ADODB.Recordset, NFeCombustivel As New ADODB.Recordset, NFeVeiculos As New ADODB.Recordset

 Dim n As Integer
 Dim i As Long
 Dim vsXML As String
 Dim XMLAuxiliar As String
 Dim XMLAuxiliarParcelas As String
 Dim msgerro As String
 Dim qterro As Long
 Dim Prod_DetEspecifico As String
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
    If Not Vazio(NFe!ChavedeAcesso) Then
       NFeResposta = sistNFe.NfeConsulta(NFe!ChavedeAcesso)
       cStat = Parse(NFeResposta, "#")
       NFeMotivo = Parse(NFeResposta, "#")
       NFeDataHora = Parse(NFeResposta, "#")
       If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
       If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
       NFeNumeroProtocolo = Parse(NFeResposta, "#")
       If cStat = 100 Or cStat = 101 Or cStat = 110 Then
          Sleep 10000
          GoTo buscaNFe
       Else
          NFeResposta = sistNFe.NfeRetAutorizacao(NFe!NumeroRecibo)
          cStat = Parse(NFeResposta, "#")
          NFeMotivo = Parse(NFeResposta, "#")
       End If
    End If
    vsSQL = "SELECT * FROM cliente WHERE CODIGO = " & NFe!CodigoCorrentista
    RsOpen Destinatario, vsSQL
    If Destinatario.RecordCount = 0 Then
       NFeMotivo = "CLIENTE/DESTINATÁRIO NÃO ENCONTRADO!"
       GoTo Caifora
    End If
    '
    '         criação dos grupos
    '
    '===================grupo de identificação do emitente (grupo B do Manual de integração - páginas 90)=======================
    '
    '        <>&" são caracteres reservados do XML e devem ser evitados ou substituídos
    '        por &lt; &gy; &amp; &quot;
    '
    '        Vale ressaltar que as aplicações das UF devem mostrar DIAS &amp; DIAS TENTANDO S/A,
    '        pois não entedem &amp; como &, assim talvez seja melhor substituir o & por e.
    '
    ReDim emit(15) 'ok
    emit(0) = RemoveAcento(Trim(Parametros!RAZAO))                                               ' Razão social do emitente, evitar caracteres acentuados e &
    emit(1) = RemoveAcento(Trim(Parametros!Fantasia))                                            ' Nome fantasia
    emit(2) = RemoveAcento(Trim(Left(Parametros!ENDERECO, 60)))                                  ' logradouro
    emit(3) = RemoveAcento(Trim(Parametros!Numero))                                         ' número, informar S/N quano inexistente para erro de Schema XML
    emit(4) = ""  'RemoveAcento(Trim(Parametros!Complemento))                                    ' complemento do endereço, o conteúdo pode ser omitido
    emit(5) = RemoveAcento(Trim(Parametros!Bairro))                                              ' bairro
    emit(6) = Trim(Parametros!CodigoIBGE)                                                        ' código do município (vide página 141 do manual), deve ser compatível com a UF
    emit(7) = RemoveAcento(Trim(Left(Parametros!Cidade, 60)))                                    ' nome do município
    emit(8) = Retira(Parametros!CEP, ".-/ ", UM_A_UM)                                            ' CEP - sem máscara
    emit(9) = Retira(Parametros!Telefone, "().- ", UM_A_UM)                                      ' número do telefone sem máscara
    emit(10) = Trim(Retira(Parametros!IE, ".,-/ ", UM_A_UM))                                     ' Inscrição Estadual do emitente sem máscara
    emit(11) = Trim(Retira(Parametros!InscricaoMunicipal, ".,-/", UM_A_UM))                  ' Inscrição Municipal
    If Not Vazio(emit(11)) Then emit(12) = "" 'Trim(Retira(Parametros!CNAEFiscal, ".,-/", UM_A_UM))  ' Código do CNAE
    'emit(13) = Trim(Retira(Parametros!InscricaoEstadualSubsTributari, ".,-/", UM_A_UM))          ' Inscrição Estadual do ST
    emit(14) = Parametros!CRT                                                                     ' <CRT> 1 – Simples Nacional; 2 – Simples Nacional – excesso de sublimite de receita bruta; 3 – Regime Normal

    '
    '======= grupo de identificação da NF-e - grupo B do Manual de integração - páginas 86 a 89
    '
    ReDim ide(29) 'ok
    ide(0) = Left(Parametros!CodigoIBGE, 2)                          ' código da UF - tabela do IBGE: 35 - SP, 43 - RS, etc. (vide página 141 do manual)
    ide(1) = NFe!cCodigoNota
    ide(2) = RemoveAcento(NFe!NaturezaOperacao)                      ' natureza da operação
    ide(3) = Left(NFe!IndicadorFormaPagamento, 1)                    ' Indicador da forma de pagamento  0 = Pagamento a vista / 1 = Pagamento a prazo / 2 = Outros
    ide(4) = "55"                                                    ' modelo da nota fiscal eletronica
    If Vazio(ide(4)) Then ide(4) = "55"
    ide(5) = "1"                                                     ' série única = 0
    ide(6) = Val(NFe!NumeroNota)                                     ' número da NF-e
    ide(7) = Format(NFe!DataEmissao, "yyyy-mm-dd") & "T" & Format(NFe!HoraSaida, "hh:mm:ss") & UTC     ' data de emissão
    ide(8) = Format(NFe!DataSaida, "yyyy-mm-dd") & "T" & Format(NFe!HoraSaida, "hh:mm:ss") & UTC       ' data em branco = 30/12/1899
    
    'nfe 3.1 ' tipo da nota fiscal 0-entrada/1-saída
    ide(9) = NFe!TipoDocumento                                       ' número da nota fiscal de saída
    '-----------------------------
    ide(10) = Parametros!CodigoIBGE                                  ' código do município do IBGE de ocorrência do FG do ICMS (vide página 141 do manual)
    
    'nfe 3.1 tipo emissao
    ide(11) = Left(NFe!FormatoEmissaoNFe, 1)                         ' forma de emissão da NF-e 1- normal, 2 - contingência FS, 3 - contingência SCAN, etc.
    Dim retorno_tpEmis As String
    retorno_tpEmis = SaveString(&H80000001, "nfe", "TipoEmissao", ide(11))
'----------------------------------------------------------------------------
    
    ide(12) = Left(NFe!FinalidadeEmissaoNFe, 1)                      ' finalidade da emissão da NF-e 1- NF-e normal
    If Vazio(ide(12)) Then ide(12) = "1"
    
    ide(13) = ""                                                     ' NF referenciada
    ide(14) = Format(NFe!HoraSaida, "hh:mm:ss")                      '<hSaiEnt> Formato “HH:MM:SS”
    
'---------------------------------------------------------------------------------------------------------------------------------------------
    If ide(11) <> "1" Then
       ide(15) = Format(NFe!ContingenciaDataHora, "yyyy-mm-ddThh:mm:ss") & UTC 'v2.03 - dhCont  AAAA-MM-DDTHH:MM:SS
       ide(16) = NFe!ContingenciaJustificativa                                 'v2.03 - xJust Justificativa da entrada em contingência
    End If
    
    
    ide(17) = "1"                           'Identificador de local de destino da operação
    
    If Len(Destinatario!CPF) = 18 And Len(Destinatario!IE) = 0 Then        'Indica operação com Consumidor final
        dest(16) = "2"  'indicador do tipo de IE
        ide(18) = "1"   'indicador do tipo de consumidor
    ElseIf Len(Destinatario!CPF) = 14 And Len(Destinatario!IE) = 0 Then
        dest(16) = "9"
        ide(18) = "1"
    ElseIf Len(Destinatario!IE) > 0 Then
        dest(16) = "1"
        ide(18) = "0"
    End If
    
    ide(18) = "1"
    ide(19) = "1"                           'Indicador de presença do comprador no estabelecimento comercial no momento da operação

    vsSQL = "SELECT * " & _
            "FROM NotaFiscalReferenciada " & _
            "WHERE CodigoNota = " & NumeroNota
    RsOpen NFeReferenciadas, vsSQL
    
    If NFeReferenciadas.RecordCount > 0 Then
       If NFeReferenciadas!ProdutorRural Then          'NFe Referenciada -> NF de Produtor referenciada
          'ide(13) = "NFP" & ";"
          'ide(21) = NFeReferenciadas!CodigoUF & ";"
          'ide(22) = NFeReferenciadas!AnoMesEmissaoNFe & ";"
          'ide(23) = Retira(NFeReferenciadas!CNPJ_CPF, ".-/", UM_A_UM) & ";"
          'ide(24) = Trim(Retira(NFeReferenciadas!InscricaoEstadual, ".,-/", UM_A_UM)) & ";"
          'ide(25) = NFeReferenciadas!ModeloNF & ";"
          'ide(26) = NFeReferenciadas!SerieNFRef & ";"
          'ide(20) = NFeReferenciadas!NumeroNF & ";"
          '------------------------------------------------------------
          'ide(13) = "NFP"
          'ide(21) = NFeReferenciadas!CodigoUF & ";"
          'ide(22) = NFeReferenciadas!AnoMesEmissaoNFe & ";"
          'ide(23) = Retira(NFeReferenciadas!CNPJ_CPF, ".-/", UM_A_UM) & ";"
          'ide(24) = NFeReferenciadas!ModeloNF & ";"
          'ide(25) = NFeReferenciadas!SerieNFRef & ";"
          
          'ide(13) = "NFP"
          'grupo de informação das NFP referenciadas. Atenção! No caso de mais de uma NFP referenciada, separe por ";" conforme ex abaixo. Todas as tags desse grupo são obrigatórias 1-1
          ide(27) = NFeReferenciadas!CodigoUF & ";"
          ide(28) = NFeReferenciadas!AnoMesEmissaoNFe & ";"
          ide(29) = Retira(NFeReferenciadas!CNPJ_CPF, ".-/", UM_A_UM) & ";"
          ide(30) = Trim(Retira(NFeReferenciadas!InscricaoEstadual, ".,-/", UM_A_UM)) & ";"
          ide(31) = NFeReferenciadas!ModeloNF & ";"
          ide(32) = NFeReferenciadas!SerieNFRef & ";"
          ide(33) = NFeReferenciadas!NumeroNF & ";"

          
       ElseIf NFeReferenciadas!ModeloNF = "55" Or NFeReferenciadas!ModeloNF = "65" Then   'NFe Referenciada -> NFe Complementar, Devolução, Retorno
          ide(13) = "NFe" & ";"
          ide(20) = NFeReferenciadas!ChavedeAcesso & ";"
          'ide(20) = "" & ";"
       ElseIf NFeReferenciadas!ModeloNF = "57" Then    'NFe Referenciada -> CTe
          ide(13) = "CTe" & ";"
          ide(20) = NFeReferenciadas!ChaveCTe & ";"
       ElseIf NFeReferenciadas!CupomFiscal Then        'NFe Referenciada -> ECF
          ide(13) = "ECF" & ";"
          ide(19) = NFeReferenciadas!nCOO & ";"
          ide(18) = NFeReferenciadas!nECF & ";"
       End If
    End If
    
   ide(27) = "1"
   ide(28) = Left$("ONLINE COMMERCE - v." & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision), 20)


'================grupo de identificação do destinatario (grupo E do Manual de integração - páginas 92)=======================
    '
    ReDim dest(45) 'ok
    If Len(Destinatario!CPF) = 0 Then dest(18) = "1"
    
    If Len(Destinatario!CPF) = 18 Then
      dest(0) = Trim(Destinatario!CPF)                                   ' CNPJ do destinatario sem máscara de formatação
      dest(0) = Retira(dest(0), ".-/", UM_A_UM)                          ' CNPJ do destinatario sem máscara de formatação
      dest(13) = Trim(Retira(Destinatario!IE, ".,-/", UM_A_UM))          ' Inscrição Estadual do destinatario sem máscara
      'dest(16) = "1"                                                     ' '<indIEDest>': Indicador da IE do destinatário:1 – Contribuinte ICMSpagamento à vista;2 – Contribuinte isento de inscrição;9 – Não Contribuinte
    Else
      dest(0) = Trim(Destinatario!CPF)                                   ' CPF do destinatario, uso exclusivo do Fisco
      dest(0) = Retira(dest(0), ".-/", UM_A_UM)                          ' CPF do destinatario, uso exclusivo do Fisco
      dest(13) = "" 'Trim(Retira(Destinatario!RG, ".,-/", UM_A_UM))          ' RG do destinatario sem máscara
      'dest(16) = "9"                                                     ' '<indIEDest>': Indicador da IE do destinatário:1 – Contribuinte ICMSpagamento à vista;2 – Contribuinte isento de inscrição;9 – Não Contribuinte
      'ide(18) = "1"
  End If
    dest(1) = RemoveAcento(Trim(Left(Destinatario!nome, 60)))            ' Razão social do destinatario, evitar caracteres acentuados e &
    dest(2) = RemoveAcento(Trim(Destinatario!ENDERECO))                  ' logradouro
    dest(3) = RemoveAcento(Trim(Destinatario!Numero))                    ' número, informar S/N quano inexistente para erro de Schema XML
    dest(4) = RemoveAcento(Trim(Destinatario!Ponto_de_referencia))       ' complemento do endereço, o conteúdo pode ser omitido
    dest(5) = RemoveAcento(Trim(Destinatario!Bairro))                    ' bairro
    dest(6) = Trim(Destinatario!CodigoIBGE)                              ' código do município (vide página 141 do manual), deve ser compatível com a UF
    dest(7) = RemoveAcento(Trim(Destinatario!Cidade))                    ' nome do município
    dest(8) = RemoveAcento(Trim(Destinatario!Estado))                    ' sigla da UF
    dest(9) = Retira(Destinatario!CEP, ".-/", UM_A_UM)                   ' CEP - sem máscara
    dest(10) = "1058"                                                    ' código do pais - deve fixo em 1058 - Brasil
    dest(11) = "BRASIL"                                                  ' nome do pais (Brasil ou BRASIL)
    dest(12) = Trim(Retira(Destinatario!Telefone1, "()-. ", UM_A_UM))    ' número do telefone sem máscara
    dest(14) = ""  'Trim(Retira(NFe!Suframa, ".,-/", UM_A_UM))           ' Inscrição SUFRAMA
    dest(15) = Trim(Destinatario!Correio_eletronico)                     ' Email
    dest(17) = Trim(NFe!InscricaoMunicipal)                         ' Inscrição Municipal do Tomador do Serviço
    
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

    If Parametros!AmbienteNF = 2 Then
       dest(0) = "99999999000191"
       dest(1) = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
       dest(13) = ""
       dest(16) = "9"
    End If
    
    vsSQL = "SELECT * FROM NotaFiscalAutorizados WHERE CodigoNota = " & NumeroNota
    RsOpen NFeAutorizados, vsSQL
    
    n = NFeAutorizados.RecordCount

    ReDim autXML(0)
    autXML(0) = ""
    
    For i = 0 To NFeAutorizados.RecordCount
       ' If i = 0 Then
       '    autXML(i, 0) = Retira(NFe!CNPJ_CPF, ".-/", UM_A_UM)
       ' Else
       '    autXML(i, 0) = Retira(NFeAutorizados!CNPJCPF, ".-/", UM_A_UM)
       ''    NFeAutorizados.MoveNext
       ' End If
    Next
        
    vsSQL = "SELECT NotaFiscalItens.* " & _
            "FROM NotaFiscalItens " & _
            "WHERE CodigoNota = " & NumeroNota & " " & _
            "ORDER BY Item"

    RsOpen NFeItens, vsSQL

    n = NFeItens.RecordCount - 1

    ReDim prod(n, 132)

    For i = 0 To n
        vsSQL = "SELECT * FROM produtos WHERE CODIGO = " & NFeItens!CodigoProduto
        RsOpen Produtos, vsSQL
        If Produtos.RecordCount = 0 Then
           NFeMotivo = "CADASTRO DO PRODUTO NÃO ENCONTRADO!" & vbNewLine & vbNewLine & "PRODUTO: " & NFeItens!NomeProduto
           GoTo Caifora
        End If
        '
        '================grupo de detalhe do produto (grupo I01 do Manual de integração - páginas 95)=======================
        '
        prod(i, 0) = RemoveAcento(Trim(NFeItens!CodigoProduto))                      ' código do produto
        prod(i, 1) = IIf(Vazio(Produtos!COD_BARRA), "", Produtos!COD_BARRA)          ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN
        prod(i, 2) = RemoveAcento(Trim(NFeItens!NomeProduto))                        ' código do produto, espaços em branco consecutivos ou no início ou fim do campo podem gerar erro de Schema XML, além de caracteres reservados do XML <>&"'
        prod(i, 3) = IIf(Vazio(Trim(Produtos!NCM)), "00000000", Trim(Produtos!NCM))  ' código NCM, pode ser omitido se não sujeito ao IPI
        prod(i, 82) = ""                                                             '<NVE>
        prod(i, 4) = ""                                                              ' ExTipi, especialização do código NCM, informar apenas se existir
        prod(i, 5) = Trim(Str(NFeItens!CFOP))                                        ' CFOP do operação, causa erro de XML se informado um código inexistente
        prod(i, 6) = RemoveAcento(Trim(NFeItens!UnidadeComercial))                   ' unidade de comercialização
        prod(i, 7) = Format(NFeItens!QuantidadeComercial, "######0.0000")            ' quantidade de comercialização
        prod(i, 7) = Substitui(prod(i, 7), ",", ".", UM_A_UM)
        prod(i, 8) = Format(NFeItens!ValorUnitarioComercializacao, "######0.0000")   ' valor unitário de comercialização, campo de mera demonstração deve ser o resultado da divisão do vProd / qCom
        prod(i, 8) = Substitui(prod(i, 8), ",", ".", UM_A_UM)
        prod(i, 9) = Format(NFeItens!ValorTotalBruto, "######0.00")                  ' valor do total do item
        prod(i, 9) = Substitui(prod(i, 9), ",", ".", UM_A_UM)
        prod(i, 10) = IIf(Vazio(Produtos!COD_BARRA), "", Produtos!COD_BARRA)         ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN, em geral é o mesmo código do EAN de comercialização
        prod(i, 11) = RemoveAcento(Trim(NFeItens!UnidadeComercial))                  ' unidade de tributação, na maioria dos casos é idêntico  ao vUnCom, pode diferente nos casos de produtos sujeitos a ST em que a unidade de pauta é diferente da unidade de comercialização
                                                                                     ' Ex. unidade de comercialização = 1 pack de lata de cerveja => unidade de tributação = 1 lata (preço de pauta)
        prod(i, 12) = Format(NFeItens!QuantidadeComercial, "######0.0000")           ' quantidade de comercialização
        prod(i, 12) = Substitui(prod(i, 12), ",", ".", UM_A_UM)
        prod(i, 13) = Format(NFeItens!ValorUnitarioComercializacao, "######0.0000")  ' valor unitário de tributação, campo de mera demonstração deve ser o resultado da divisão do vProd / qTrib
        prod(i, 13) = Substitui(prod(i, 13), ",", ".", UM_A_UM)
        prod(i, 14) = Format(NFeItens!ValorFrete, "######0.00")                      ' valor do frete, se cobrado do cliente deve ser rateado entre os itens de produto
        prod(i, 14) = Substitui(prod(i, 14), ",", ".", UM_A_UM)
        prod(i, 15) = Format(NFeItens!ValorSeguro, "######0.00")                     ' valor do seguro, se cobrado do cliente deve ser rateado entre os itens de produto
        prod(i, 15) = Substitui(prod(i, 15), ",", ".", UM_A_UM)
        prod(i, 16) = Format(NFeItens!ValorDesconto, "######0.00")                   ' valor do desconto concedido
        prod(i, 16) = Substitui(prod(i, 16), ",", ".", UM_A_UM)
        prod(i, 80) = "0.00"

        If NFeItens!ValorTributos > 0 Then
           prod(i, 53) = RemoveAcento(Trim(Left(NFeItens!InformacoesAdicionaisProduto, 400))) & " - Valor Aproximado dos Tributos R$ " & Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)        ' informações adicionais do produto
           vlTrib = vlTrib + NFeItens!ValorTributos
        Else
           prod(i, 53) = RemoveAcento(Trim(Left(NFeItens!InformacoesAdicionaisProduto, 500)))     ' informações adicionais do produto
        End If

        prod(i, 76) = "1"
        
        'Valor aproximado total de tributos federais, estaduais e municipais
        If NFeItens!ValorTributos > 0 Then prod(i, 81) = Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)
        
        '
        '   gera grupo do destinatário
        '
        Select Case NFeItens!TipoProduto     'Veículo|Medicamento|Armamento|Combustível
          Case "Armamento"
            vsSQL = "SELECT * " & _
                    "FROM NotaFiscalItensArmamento " & _
                    "WHERE (NumeroNota = " & NumeroNota & " AND SerieNF = " & SerieNF & ") AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
            RsOpen NFeArmamento, vsSQL
            If NFeArmamento.RecordCount > 0 Then
              Do While Not NFeArmamento.EOF
                'Prod_DetEspecifico = Prod_DetEspecifico + objNFeUtil.arma(NFeArmamento!tpArma, NFeArmamento!nSerie, NFeArmamento!nCano, NFeArmamento!ArmDescricao)
                NFeArmamento.MoveNext
              Loop
            End If
          Case "Combustível"
            vsSQL = "SELECT * " & _
                    "FROM NotaFiscalItensCombustivel " & _
                    "WHERE CodigoNota = " & NumeroNota & " AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
            RsOpen NFeCombustivel, vsSQL
            If NFeCombustivel.RecordCount > 0 Then
               Do While Not NFeCombustivel.EOF
                  prod(i, 93) = NFeCombustivel!cProdANP                                              'cProdANP
                  If NFeCombustivel!cProdANP = "210203001" Then
                     prod(i, 94) = Format(NFeCombustivel!pMixGN, "#0.0000")                          'pMixGN
                     prod(i, 94) = Substitui(prod(i, 94), ",", ".", UM_A_UM)
                  End If
                  prod(i, 95) = NFeCombustivel!CODIF                                                 'CODIF
                  prod(i, 96) = Format(NFeCombustivel!qTemp, "#0.0000")                              'qTemp
                  prod(i, 96) = Substitui(prod(i, 96), ",", ".", UM_A_UM)
                  prod(i, 97) = NFeCombustivel!UFCons                                                'UFCons
                  'CIDE
                  If NFeCombustivel!qBCProd > 0 Then
                     prod(i, 98) = Format(NFeCombustivel!qBCProd, "#0.0000")                         'qBCProd
                     prod(i, 98) = Substitui(prod(i, 98), ",", ".", UM_A_UM)
                     prod(i, 99) = Format(NFeCombustivel!vAliqProd, "#0.0000")                       'vAliqProd
                     prod(i, 99) = Substitui(prod(i, 99), ",", ".", UM_A_UM)
                     prod(i, 100) = Format(NFeCombustivel!vCIDE, "#0.00")                            'vCIDE
                     prod(i, 100) = Substitui(prod(i, 100), ",", ".", UM_A_UM)
                  End If
                  NFeCombustivel.MoveNext
               Loop
            End If
          Case "Medicamento"
            vsSQL = "SELECT * " & _
                    "FROM NotaFiscalItensMedicamento " & _
                    "WHERE (NumeroNota = " & NumeroNota & " AND SerieNF = " & SerieNF & ") AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
            RsOpen NFeMedicamentos, vsSQL
            If NFeMedicamentos.RecordCount > 0 Then
              Do While Not NFeMedicamentos.EOF
                prod(i, 71) = IIf(Vazio(NFeMedicamentos!nLote), "0", NFeMedicamentos!nLote)
                prod(i, 72) = Format(NFeMedicamentos!QuantidadeLote, "#0.000")
                prod(i, 72) = Substitui(prod(i, 72), ",", ".", UM_A_UM)
                prod(i, 73) = IIf(IsNull(NFeMedicamentos!DataFabricacao), Format(DateAdd("yyyy", -1, Date), "mm/dd/yyyy"), Format(NFeMedicamentos!DataFabricacao, "yyyy-mm-dd"))
                prod(i, 74) = IIf(IsNull(NFeMedicamentos!DataValidade), Format(DateAdd("m", 6, NFeMedicamentos!DataValidade), "mm/dd/yyyy"), Format(NFeMedicamentos!DataValidade, "yyyy-mm-dd"))
                prod(i, 75) = Format(NFeMedicamentos!PMC, "#0.00")
                prod(i, 75) = Substitui(prod(i, 75), ",", ".", UM_A_UM)
                NFeMedicamentos.MoveNext
              Loop
            End If
          Case "Veículo"
            vsSQL = "SELECT * " & _
                    "FROM NotaFiscalItensVeiculos " & _
                    "WHERE (NumeroNota = " & NumeroNota & " AND SerieNF = " & SerieNF & ") AND (CodigoProduto = '" & NFeItens!CodigoProduto & "') "
            RsOpen NFeVeiculos, vsSQL
            If NFeVeiculos.RecordCount > 0 Then
              Do While Not NFeVeiculos.EOF
                'Prod_Renavam = LPad(Retira(NFeVeiculos!RENAVAM, ".", UM_A_UM), 9, "0")
                'Prod_DetEspecifico = Prod_DetEspecifico + objNFeUtil.veicProd(Left(NFeVeiculos!TipoOperacao, 1), NFeVeiculos!Chassi, NFeVeiculos!Cor, NFeVeiculos!DescricaoCor, NFeVeiculos!PotenciaMotor, NFeVeiculos!CM3, NFeVeiculos!VeicPesoLiquido, NFeVeiculos!VeicPesoBruto, NFeVeiculos!VeicSerie, NFeVeiculos!VeicTipoCombustivel, NFeVeiculos!VeicNumeroMotor, NFeVeiculos!CMKG, NFeVeiculos!DistanciaentreEixos, Prod_Renavam, NFeVeiculos!AnoMod, NFeVeiculos!AnoFab, NFeVeiculos!tpPintura, NFeVeiculos!tpVeiculo, NFeVeiculos!espVeiculo, NFeVeiculos!VIN, NFeVeiculos!condVeiculo, NFeVeiculos!cModelo)
                'Prod_DetEspecifico = Prod_DetEspecifico + objNFeUtil.veicProd2G(Left(NFeVeiculos!TipoOperacao, 1), NFeVeiculos!Chassi, NFeVeiculos!Cor, NFeVeiculos!DescricaoCor, NFeVeiculos!PotenciaMotor, NFeVeiculos!CM3, NFeVeiculos!VeicPesoLiquido, NFeVeiculos!VeicPesoBruto, NFeVeiculos!VeicSerie, NFeVeiculos!VeicTipoCombustivel, NFeVeiculos!VeicNumeroMotor, NFeVeiculos!CMKG, NFeVeiculos!DistanciaentreEixos, NFeVeiculos!AnoMod, NFeVeiculos!AnoFab, NFeVeiculos!tpPintura, NFeVeiculos!tpVeiculo, NFeVeiculos!espVeiculo, NFeVeiculos!VIN, NFeVeiculos!condVeiculo, NFeVeiculos!cModelo, Left(NFeVeiculos!cCorDENATRAN, 1), NFeVeiculos!veicLotacao, Left(NFeVeiculos!veictpRestricao, 1))
                NFeVeiculos.MoveNext
              Loop
              prod(i, 79) = Prod_DetEspecifico
            End If
        End Select

        '
        '=========dados do ICMS (grupo N01 do Manual de integração - páginas 100)=====================
        '
        prod(i, 17) = Left(NFeItens!CST, 1)                           ' Tabela A - origem da mercadoria 0=nacional
        prod(i, 18) = Mid(NFeItens!CST, Len(NFeItens!CST) - 1)        ' Tabela B - CST=00-tributação normal
        If Vazio(NFeItens!modBC) Then
          prod(i, 19) = 3                                            ' modalidade de determinação da BC = 3-valor da operação
        Else
          prod(i, 19) = Left(NFeItens!modBC, 1)                      ' modalidade de determinação da BC = 3-valor da operação
        End If
        prod(i, 20) = Format(NFeItens!vBC, "######0.00")               ' valor da BC do ICMS = vProd + vFrete + vSeguro
        prod(i, 20) = Substitui(prod(i, 20), ",", ".", UM_A_UM)
        prod(i, 21) = Format(NFeItens!pICMS, "######0.00")           ' alíquota do ICMS
        prod(i, 21) = Substitui(prod(i, 21), ",", ".", UM_A_UM)
        prod(i, 22) = Format(NFeItens!vICMS, "######0.000")           ' valor do ICMS
        prod(i, 22) = Substitui(prod(i, 22), ",", ".", UM_A_UM)
        prod(i, 46) = IIf(Not Vazio(NFeItens!modBCST), Left(NFeItens!modBCST, 1), "5")
        prod(i, 47) = Format(NFeItens!pMVAST, "######0.00")         ' percentual de valor de margem e valor adicionado
        prod(i, 47) = Substitui(prod(i, 47), ",", ".", UM_A_UM)
        prod(i, 48) = Format(NFeItens!pRedBCST, "######0.00")       ' percentual de redução da BC do ICMS ST
        prod(i, 48) = Substitui(prod(i, 48), ",", ".", UM_A_UM)
        prod(i, 49) = Format(NFeItens!vBCST, "######0.00")          ' BC do ICMS ST
        prod(i, 49) = Substitui(prod(i, 49), ",", ".", UM_A_UM)
        prod(i, 50) = Format(NFeItens!pICMSST, "######0.00")        ' percentual do ICMSST
        prod(i, 50) = Substitui(prod(i, 50), ",", ".", UM_A_UM)
        prod(i, 51) = Format(NFeItens!vICMSST, "######0.00")        ' valor do ICMS ST devido
        prod(i, 51) = Substitui(prod(i, 51), ",", ".", UM_A_UM)
        prod(i, 52) = Format(NFeItens!pRedBC, "######0.00")         ' percentual de redução da BC
        prod(i, 52) = Substitui(prod(i, 52), ",", ".", UM_A_UM)

 '       If Prod(i, 18) = "20" Or Prod(i, 18) = "30" Or Prod(i, 18) = "40" Or Prod(i, 18) = "70" Or Prod(i, 18) = "90" Then
 '          Prod(i, 88) = Format(0, "######0.00")                    ' vICMSDeson
 '          Prod(i, 88) = Substitui(Prod(i, 88), ",", ".", UM_A_UM)
 '          Prod(i, 89) = "9"                                        ' motDesICMS
 '       End If
'        If prod(i, 18) = "20" Or prod(i, 18) = "30" Or prod(i, 18) = "40" Or prod(i, 18) = "70" Or prod(i, 18) = "90" Then
'           prod(i, 106) = Format(0, "######0.00")                    ' vICMSDeson
'           prod(i, 106) = Substitui(prod(i, 88), ",", ".", UM_A_UM)
'           prod(i, 85) = "9"                                        ' motDesICMS
'        End If

        If emit(14) = "1" Then
           vCredICMSSN = 0
           prod(i, 17) = "0"                                                      ' Tabela A - origem da mercadoria 0=nacional
           prod(i, 18) = NFeItens!CST                                             ' Tabela B - CST=00-tributação normal
           prod(i, 77) = Format(Parametros!pCreditoICMSSimplesNacional, "#0.00")  ' <pCredSN>          Simples Nacional
           prod(i, 77) = Substitui(prod(i, 77), ",", ".", UM_A_UM)
           If Parametros!pCreditoICMSSimplesNacional > 0 Then
              vCredICMSSN = Round(NFeItens!ValorTotalBruto * (Parametros!pCreditoICMSSimplesNacional / 100), 2)
              prod(i, 78) = Format(vCredICMSSN, "#0.00")                          ' <vCredICMSSN>      Simples Nacional
              prod(i, 78) = Substitui(prod(i, 78), ",", ".", UM_A_UM)
           Else
              prod(i, 78) = "0.00"                                                ' <vCredICMSSN>      Simples Nacional
           End If
        End If
        '
        '=========dados do PIS (grupo Q do Manual de Integração - páginas 110) =============
        '
        prod(i, 31) = IIf(Vazio(NFeItens!PISCST) Or IsNull(NFeItens!PISCST), "99", NFeItens!PISCST)
        prod(i, 32) = Format(NFeItens!PISvBC, "######0.00")
        prod(i, 32) = Substitui(prod(i, 32), ",", ".", UM_A_UM)
        prod(i, 33) = Format(NFeItens!PISpPIS, "####0.00")
        prod(i, 33) = Substitui(prod(i, 33), ",", ".", UM_A_UM)
        prod(i, 34) = Format(NFeItens!PISvPIS, "####0.00")
        prod(i, 34) = Substitui(prod(i, 34), ",", ".", UM_A_UM)
        prod(i, 45) = Format(NFeItens!PISvAliqProd, "####0.00")
        prod(i, 45) = Substitui(prod(i, 45), ",", ".", UM_A_UM)
        'tag PISST
        prod(i, 54) = ""       'vBC
        prod(i, 55) = ""       'pPIS
        prod(i, 56) = ""       'vPIS

        '
        '========dados do COFINS (grupo s do Manual de Integração - páginas 113) ============
        '
        prod(i, 35) = IIf(Vazio(NFeItens!COFINSCST) Or IsNull(NFeItens!COFINSCST), "99", NFeItens!COFINSCST)
        prod(i, 36) = Format(NFeItens!COFINSvBC, "######0.00")
        prod(i, 36) = Substitui(prod(i, 36), ",", ".", UM_A_UM)
        prod(i, 37) = Format(NFeItens!COFINSpCOFINS, "####0.00")
        prod(i, 37) = Substitui(prod(i, 37), ",", ".", UM_A_UM)
        prod(i, 38) = Format(NFeItens!COFINSvCOFINS, "####0.00")
        prod(i, 38) = Substitui(prod(i, 38), ",", ".", UM_A_UM)
        prod(i, 44) = Format(NFeItens!COFINSvAliqProd, "####0.00")
        prod(i, 44) = Substitui(prod(i, 44), ",", ".", UM_A_UM)
       'tag COFINSST
        prod(i, 57) = ""       'vBC
        prod(i, 58) = ""       'pCOFINS
        prod(i, 59) = ""       'vCOFINS

        '   gera grupo do IPI
        '
        If Not Vazio(NFeItens!IPICST) Then
           prod(i, 23) = NFeItens!IPICST
           prod(i, 24) = NFeItens!IPIvBC
           prod(i, 25) = NFeItens!IPIpIPI
           prod(i, 26) = NFeItens!IPIvIPI
        Else
           prod(i, 23) = "53"
        End If
        
        prod(i, 109) = "999"

        '   gera grupo do II - Importação
        '
        prod(i, 27) = FormatoDecimal(NFeItens!IIvBC)
        prod(i, 28) = FormatoDecimal(NFeItens!IIvDespAdu)
        prod(i, 29) = FormatoDecimal(NFeItens!IIvII)
        prod(i, 30) = FormatoDecimal(NFeItens!IIvIOF)

        'Tag da Declaração de Importação
        prod(i, 60) = ""      'nDI
        prod(i, 61) = ""      'dDI
        prod(i, 62) = ""      'xLocDesemb
        prod(i, 63) = ""      'UFDesemb
        prod(i, 64) = ""      'dDesemb
        prod(i, 65) = ""      'cExportador
        prod(i, 66) = ""      'adi: nAdicao
        prod(i, 67) = ""      'adi: nSeqAdic
        prod(i, 68) = ""      'adi: cFabricante
        prod(i, 69) = ""      'adi: vDescDI
        
        'tag ISSQN
        prod(i, 39) = "0.00"                              'ISSQN <vBC>
        prod(i, 40) = "0.00"                              'ISSQN <vAliq>
        prod(i, 41) = "0.00"                              'ISSQN <vISSQN>
        prod(i, 42) = ""                                  'ISSQN <cMunFG>
        prod(i, 43) = ""                                  'ISSQN <cListServ>
        prod(i, 70) = ""                                  'ISSQN: cSitTrib -  Código da tributação do ISSQN: N – NORMAL; R – RETIDA; S –SUBSTITUTA; I – ISENTA. (v.2.0)

        NFeItens.MoveNext
    Next

    '
    '   atualização de total
    '
    ReDim tot(40) 'ok
    tot(0) = Format(NFe!BaseICMS, "######0.00")
    tot(0) = Substitui(tot(0), ",", ".", UM_A_UM)
    tot(1) = Format(NFe!ValorICMS, "######0.00")
    tot(1) = Substitui(tot(1), ",", ".", UM_A_UM)
    tot(2) = Format(NFe!BaseICMSST, "######0.00")
    tot(2) = Substitui(tot(2), ",", ".", UM_A_UM)
    tot(3) = Format(NFe!ValorICMSST, "######0.00")
    tot(3) = Substitui(tot(3), ",", ".", UM_A_UM)
    tot(4) = Format(NFe!ValorProdutos, "######0.00")
    tot(4) = Substitui(tot(4), ",", ".", UM_A_UM)
    tot(5) = Format(NFe!ValorFrete, "######0.00")
    tot(5) = Substitui(tot(5), ",", ".", UM_A_UM)
    tot(6) = Format(NFe!ValorSeguro, "######0.00")
    tot(6) = Substitui(tot(6), ",", ".", UM_A_UM)
    tot(7) = Format(NFe!ValorDesconto, "######0.00")
    tot(7) = Substitui(tot(7), ",", ".", UM_A_UM)
    tot(8) = Format(NFe!ValorImportacao, "######0.00")
    tot(8) = Substitui(tot(8), ",", ".", UM_A_UM)
    tot(9) = Format(NFe!ValorIPI, "######0.00")
    tot(9) = Substitui(tot(9), ",", ".", UM_A_UM)
    tot(10) = Format(NFe!ValorPIS, "######0.00")
    tot(10) = Substitui(tot(10), ",", ".", UM_A_UM)
    tot(11) = Format(NFe!ValorCOFINS, "######0.00")
    tot(11) = Substitui(tot(11), ",", ".", UM_A_UM)
    tot(12) = Format(NFe!ValorOutrasDespesas, "######0.00")
    tot(12) = Substitui(tot(12), ",", ".", UM_A_UM)
    tot(13) = Format(NFe!ValorNota, "######0.00")
    tot(13) = Substitui(tot(13), ",", ".", UM_A_UM)
    tot(19) = "0.00"    'ICMSTot <vTotTrib>
    If vlTrib > 0 Then tot(19) = Substitui(Format(vlTrib, "#0.00"), ",", ".", UM_A_UM)
    
    tot(20) = "0.00"    'ICMSTot <vICMSDesn>
    tot(37) = "0.00"    'ICMSTot <vFCPUFDest>
    tot(38) = "0.00"    'ICMSTot <vICMSUFDest>
    tot(39) = "0.00"    'ICMSTot <vICMSUFRemet>
    
    'grupo ISSQN
    tot(14) = "" '"10.00"    'ISSQNtot <vServ>
    tot(15) = "" '"10.00"    'ISSQNtot <vBC>
    tot(16) = "" '"0.50"     'ISSQNtot <vISS>
    tot(17) = "" '"0.10"     'ISSQNtot <vPIS>
    tot(18) = "" '"50.00"    'ISSQNtot <vCOFINS>

    '
    '============dados do transportador
    '
    ReDim trp(16) 'ok
    trp(0) = IIf(Vazio(NFe!modFrete), 0, Left(NFe!modFrete, 1))       ' responsabilidade do frete 0-emitente, 1-destinatário
    If Not Vazio(NFe!TranspCNPJ_CPF) Then
        If Len(Retira(NFe!TranspCNPJ_CPF, ".-/", UM_A_UM)) = 14 Then
          trp(1) = Trim(NFe!TranspCNPJ_CPF)                                         ' CNPJ da Transportadora sem mascara
          trp(1) = Retira(trp(1), ".-/", UM_A_UM)                          ' CNPJ da Transportadora sem mascara
        Else
          trp(1) = Trim(NFe!TranspCNPJ_CPF)                                          ' CPF da Transportadora sem mascara
          trp(1) = Retira(trp(1), ".-/", UM_A_UM)                            ' CPF da Transportadora sem mascara
        End If
        trp(2) = RemoveAcento(Trim(NFe!TranspNome))
        trp(3) = Trim(Retira(NFe!TranspInscricaoEstadual, ".,-/", UM_A_UM))         ' Inscrição Estadual da Transportadora sem máscara
        trp(4) = RemoveAcento(Trim(NFe!TranspEndereco))
        trp(5) = RemoveAcento(Trim(NFe!TranspMunicipio))
        trp(6) = NFe!TranspUF
    End If

    If Not Vazio(NFe!TranspPlaca) Then
       trp(7) = Retira(Trim(NFe!TranspPlaca), "-", UM_A_UM)
       trp(8) = NFe!TranspPlacaUF
       trp(15) = NFe!TranspRNTC
    End If

    '  ============== criação dos lacres do volume
    '
    If NFe!VolumeQuantidade > 0 Then
       trp(9) = NFe!VolumeQuantidade                                  ' quantidade de volumes
       trp(9) = Substitui(trp(9), ",", ".", UM_A_UM) & ";"
       trp(10) = RemoveAcento(NFe!VolumeEspecie)                      ' espécie dos volumes
       trp(11) = RemoveAcento(NFe!VolumeMarca) & ";"                  ' marca dos volumes
       trp(12) = NFe!VolumeNumeracao                                  ' numeração dos volumes
       trp(13) = NFe!VolumePesoLiquido                                ' peso líquido
       trp(13) = Substitui(trp(13), ",", ".", UM_A_UM) & ";"
       trp(14) = NFe!VolumePesoBruto                                  ' peso bruto
       trp(14) = Substitui(trp(14), ",", ".", UM_A_UM) & ";"
    End If

    vsSQL = "SELECT * " & _
            "FROM NotaFiscalParcelas " & _
            "WHERE CodigoNota = " & NumeroNota

    RsOpen NFeParcelas, vsSQL

    ReDim cob(0, 12)
    If NFeParcelas.RecordCount > 0 Then
      ReDim cob(NFeParcelas.RecordCount - 1, 12)
      cob_numero_parcelas = NFeParcelas.RecordCount - 1
      If ide(4) = "55" Then
         For i = 0 To NFeParcelas.RecordCount - 1
           cob(i, 3) = NFeParcelas!Documento
           cob(i, 4) = Format(NFeParcelas!Vencimento, "yyyy-mm-dd")
           cob(i, 5) = Format(NFeParcelas!ValorDocumento, "####0.00")
           cob(i, 5) = Substitui(cob(i, 5), ",", ".", UM_A_UM)
           NFeParcelas.MoveNext
         Next
      ElseIf ide(4) = "65" Then
         For i = 0 To NFeParcelas.RecordCount - 1
           cob(i, 6) = "05"
           cob(i, 7) = tot(13)
           NFeParcelas.MoveNext
         Next
      End If
    Else
      cob(0, 3) = ide(6)
      cob(0, 4) = Format(Date, "yyyy-mm-dd")
      cob(0, 5) = tot(13)
    End If

    If Not Vazio(NFe!NumeroFatura) And ide(4) = "55" Then
      cob(0, 0) = ide(6)
      cob(0, 1) = tot(13)
      cob(0, 2) = tot(13)
    End If
    '
    '============= informações adcionais
    '
    ReDim obs(10)
    obs(0) = RemoveAcento(Trim(NFe!InformacoesAdicionais))

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

    obs(1) = RemoveAcento(Trim(NFe!InformacoesComplementares)) & IIf(Vazio(OBSNFe), "", " // " & RemoveAcento(Trim(OBSNFe)))

    'tag exporta v2.03
    obs(2) = ""      'UFEmbarq
    obs(3) = ""      'xLocEmbarq

    'tag compra v2.03
    obs(4) = ""      'xNEmp
    obs(5) = ""      'xPed
    obs(6) = ""      'infCpl

End If

dirXML = Parametros!DiretorioXML
dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")

'gera a chave da nfe
Dim id_chave As String, numero_nfe_gerado As String

'pega o endereço do arquivo a ser gerado
If Not Existe(dirXML) Then MkDir dirXML

numero_nfe_gerado = sistNFe.GeraXML(ide(), emit(), dest(), prod(), tot(), trp(), cob(), obs(), autXML(), False)
id_chave = numero_nfe_gerado
numero_nfe_gerado = Replace(numero_nfe_gerado, "NFe", "")
NFeChaveAcesso = numero_nfe_gerado

If Not Vazio(NFeChaveAcesso) Then
   vsSQL = "UPDATE NotaFiscal SET " & _
           "ChavedeAcesso = '" & NFeChaveAcesso & "' " & _
           "WHERE CodigoNota = " & NumeroNota
   vgDb.Execute vsSQL
End If

xCaminhoXML = dirXML & "nfe\arquivos\" & id_chave & ".xml"

NFeResposta = sistNFe.AssinarArquivoXML(xCaminhoXML, "infNFe")

'MsgBox NFeResposta, vbExclamation, "Assinatura"

If InStr(NFeResposta, "Erro") > 0 Then
   NFeMotivo = NFeResposta
   GoTo NaoEnviou
End If

xCaminhoXML = dirXML & "nfe\arquivos\assinado\" & id_chave & "-assinado.xml"

XMLOK = sistNFe.ValidarArquivoXML(xCaminhoXML, False, NFeValidate)

If Not XMLOK Then
   NFeMotivo = "Erro na Validação do XML, falha no Schema" & vbNewLine & NFeValidate
   GoTo Caifora
End If

xCaminhoXML = dirXML & "nfe\arquivos\gerados\" & id_chave & ".xml"

If Not PodeEnviar Then GoTo NaoEnviou

NFeResposta = sistNFe.NfeAutorizacao(xCaminhoXML, True)

If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
   MsgBox "*** Aparentemente Ocorreram Erros na Recepção do Lote (nfeAutorizacao)***" _
   & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeAutorizacao"
   NFeMotivo = NFeResposta
   GoTo Caifora
End If

If InStr(NFeResposta, "Erro") > 0 Then GoTo Caifora
NFeMotivo = Parse(NFeResposta, "#")
If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo Caifora
NFeNumeroRecibo = Parse(NFeResposta, "#")
cStat = Right(Parse(NFeResposta, "#"), 3)
NFeDataHora = Right(Parse(NFeResposta, "#"), 25)
If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)

If cStat <> 103 Then
   GoTo NaoEnviou
End If

vsSQL = "UPDATE NotaFiscal SET " & _
        "NumeroRecibo = " & NFeNumeroRecibo & " " & _
        "WHERE CodigoNota = " & NumeroNota
vgDb.Execute vsSQL

DoEvents

consultaNFe:

    NFeResposta = sistNFe.NfeRetAutorizacao(NFeNumeroRecibo)
   
    If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
       MsgBox NFeResposta, vbExclamation, "Retorno Autorização"
       NFeMotivo = NFeResposta
       GoTo Caifora
    End If
   

    ' Testa erro 217-Rejeição: NF-e não consta na base de dados da SEFAZ <dhRecbto xmlns="http://www.portalfiscal.inf.br/nfe">2015-03-29T09:56:36-03:00
    If InStr(NFeResposta, "217") > 0 Then
            Sleep 3000 ' Aguarda mais 3 segundos
            NFeResposta = sistNFe.NfeRetAutorizacao(NFeNumeroRecibo) 'refaz a consulta
            If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
               MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote (nfeRetAutorizacao)***" _
                       & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeRetAutorizacao"
               GoTo Caifora
            End If
    End If

    cStat = Parse(NFeResposta, "#")
    NFeMotivo = Parse(NFeResposta, "#")
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
    NFeNumeroProtocolo = Parse(NFeResposta, "#")
    NFeDataHora = Parse(NFeResposta, "#")
    If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
    If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
    nroRecibo = Parse(NFeResposta, "#")
    nfeRetorno = Parse(NFeResposta, "#")
    NFeValidate = Parse(NFeResposta, "#")
    DoEvents

buscaNFe:


   'Consulta Nfe
   NFeResposta = sistNFe.NfeConsulta(NFeChaveAcesso)
   
   Dim NFeRespostaFinal As String
   NFeRespostaFinal = NFeResposta
   
   
'        If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
'           MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote ***" _
'           & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO"
'           GoTo caiFora
'        End If
'------------------------------------------------
' testa erro 217 Rejeição: NF-e não consta na base de dados da SEFAZ <dhRecbto xmlns="http://www.portalfiscal.inf.br/nfe">2015-03-29T09:56:36-03:00
      If InStr(NFeResposta, "217") > 0 Then
         Sleep 3000 ' aguarda mais 3 segundos
         NFeResposta = sistNFe.NfeConsulta(NFeChaveAcesso) ' consulta novamente
         
         'MsgBox NFeResposta, vbExclamation, "Consulta Chave Acesso Seg.Tentativa"
         If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
            MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote ***" _
                     & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeConsulta"
            GoTo Caifora
         End If
     End If
'--------------------------------------------------

    If InStr(NFeResposta, "Erro 98") > 0 Then
       NFeMotivo = Parse(NFeResposta, "#")
       NFeNumeroProtocolo = ""
       GoTo Caifora
    End If
    
    On Error Resume Next
    cStat = Parse(NFeResposta, "#")
    NFeMotivo = Parse(NFeResposta, "#")
    NFeDataHora = Parse(NFeResposta, "#")
    If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
    If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
    NFeNumeroProtocolo = Parse(NFeResposta, "#")
    DoEvents

    If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo Caifora
    
    If nroRecibo = 204 Or nroRecibo = 539 Then
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
    ElseIf nroRecibo <> 100 And nroRecibo <> 301 Then
        NFeMotivo = nroRecibo + " - " + nfeRetorno
        GoTo NaoEnviou
    ElseIf nroRecibo = 105 And nroRecibo = 217 Then
        GoTo consultaNFe
    ElseIf nroRecibo = 100 Then
        nfeRetorno = "Nota Fiscal Eletronica Autorizado o Uso."
        msgResultado = "Chave NF-e.: " + NFeChaveAcesso & vbCrLf
        msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
        msgResultado = msgResultado + "Recibo.: " + NFeNumeroRecibo & vbCrLf
        msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
        msgResultado = msgResultado + "Resposta da Fazenda.: " + nroRecibo + " - " & nfeRetorno
        
        MsgBox msgResultado, vbInformation + vbOKOnly
        
        DoEvents
        
        On Error Resume Next
        NFeResposta = sistNFe.NfeConsulta(NFeChaveAcesso)
             
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
'''End If
' Gera PDF danfe

    Dim xmlPathPDF As String
    Dim anoEmes As String
    Dim Arquivo As String
    anoEmes = dirXML & "nfe\arquivos\procNFe\" & Mid(ide(7), 1, 4) & Mid(ide(7), 6, 2) & "\"
    xmlPathPDF = dirXML & "nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
    Arquivo = anoEmes & NFeChaveAcesso & "-procNFe.xml"             '  Aqui Gera o DANFE
    Call sistNFe.ImpNFe(Arquivo, False, "", True, xmlPathPDF) 'gera pdf

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
Caifora:

    MsgBox NFeMotivo, vbCritical + vbOKOnly
    
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
 On Error GoTo TransmitirNFCe_Error

 Dim sistNFCe As snfe.Util
 Set sistNFCe = New snfe.Util
 
 vlPIS = 0
 vlCOFINS = 0
 vlTrib = 0
 pFrete = 0
 pDesconto = 0
 pOutras = 0
 pTributos = 0
 
 vsSQL = "SELECT * FROM Empresa"
 RsOpen Parametros, vsSQL
 
 empUF = Parametros!UF
 
 If ModeloNF = "65" Then
    IdToken = LPad(Parametros!NFCeIDToken, 6, "0")
    Token = Parametros!NFCeToken
 End If
  
 Screen.MousePointer = vbHourglass
 
 If ModeloNF = "55" Then
    'xVerProcNFe = SQLExecutaRetorno("SELECT VersaoLeiauteNFe As r FROM TbFiliais WHERE IdFilial = " & vgFilialNF, "r", "2.00")
    'SaveString &H80000001, "nfe", "VerProc", xVerProcNFe
    'vsSQL = "SELECT TbNotaFiscalProd.*, TbCidades.CodigoIBGE " & _
    '        "FROM TbNotaFiscalProd INNER JOIN TbCidades ON TbNotaFiscalProd.Municipio = TbCidades.NomeCidade AND TbNotaFiscalProd.UF = TbCidades.UF " & _
    '        "WHERE IdNFProd = " & NumeroNota
 Else
    vsSQL = "SELECT TbNFCe.*, Cidades.CodigoIBGE " & _
            "FROM TbNFCe INNER JOIN Cidades ON TbNFCe.Municipio = Cidades.Municipio AND TbNFCe.UF = Cidades.UF " & _
            "WHERE IdNFProd = " & NumeroNota
 End If
 Set NFe = vgDb.OpenRecordset(vsSQL)
 
 If NFe.RecordCount > 0 Then
    NFe.MoveFirst
    '
    '         criação dos grupos
    '
    '===================grupo de identificação do emitente (grupo B do Manual de integração - páginas 90)=======================
    '
    '        <>&" são caracteres reservados do XML e devem ser evitados ou substituídos
    '        por &lt; &gy; &amp; &quot;
    '
    '        Vale ressaltar que as aplicações das UF devem mostrar DIAS &amp; DIAS TENTANDO S/A,
    '        pois não entedem &amp; como &, assim talvez seja melhor substituir o & por e.
    '
    ReDim emit(15)
    emit(0) = RemoveAcento(Trim(Parametros!RazaoSocial))        ' Razão social do emitente, evitar caracteres acentuados e &
    emit(1) = RemoveAcento(Trim(Parametros!Fantasia))           ' Nome fantasia
    emit(2) = RemoveAcento(Trim(Left(Parametros!ENDERECO, 60))) ' logradouro
    emit(3) = RemoveAcento(Trim(Parametros!Numero))             ' número, informar S/N quano inexistente para erro de Schema XML
    emit(4) = RemoveAcento(Trim(Parametros!Complemento))        ' complemento do endereço, o conteúdo pode ser omitido
    emit(5) = RemoveAcento(Trim(Parametros!Bairro))             ' bairro
    emit(6) = Parametros!CodigoIBGE                             ' código do município (vide página 141 do manual), deve ser compatível com a UF
    emit(7) = RemoveAcento(Trim(Left(Parametros!Cidade, 60)))   ' nome do município
    emit(8) = Retira(Parametros!CEP, ".-/", UM_A_UM)            ' CEP - sem máscara
    emit(9) = Retira(Parametros!Telefone, "().-", UM_A_UM)
    emit(9) = Retira(IIf(Left(emit(9), 1) = "0", Mid(emit(9), 2), emit(9)), " ", UM_A_UM)                ' número do telefone sem máscara
    If Len(emit(9)) = 1 Then emit(9) = ""
    emit(10) = Trim(Retira(Parametros!InscricaoEstadual, ".,-/", UM_A_UM))   ' Inscrição Estadual do emitente sem máscara
    emit(11) = ""                                                                         ' Inscrição Municipal
    If Not Vazio(emit(11)) Then emit(12) = ""                                             ' Código do CNAE
    emit(13) = ""                                                                       ' Inscrição Estadual do ST
    emit(14) = Left(NFe!CRT, 1)
    
    '
    '======= grupo de identificação da NF-e - grupo B do Manual de integração - páginas 86 a 89
    '
    ReDim ide(36)
    ide(0) = Left(Parametros!CodigoIBGE, 2)                        ' código da UF - tabela do IBGE: 35 - SP, 43 - RS, etc. (vide página 141 do manual)
    ide(1) = NFe!NFeCodigoNota
    ide(2) = RemoveAcento(NFe!NaturezaOperacao)                    ' natureza da operação
    ide(3) = Left(NFe!NFeIndicadorFormaPagto, 1)                   ' Indicador da forma de pagamento  0 = Pagamento a vista / 1 = Pagamento a prazo / 2 = Outros
    ide(4) = ModeloNF                                              ' modelo da nota fiscal eletronica
    ide(5) = NFe!SerieNF                                           ' série única = 0
    ide(6) = Val(NFe!NumeNota)                                     ' número da NF-e
    ide(7) = Format(NFe!DataEmissao, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC       ' data de emissão
    If Not IsNull(NFe!DataSaidaEntrada) Then ide(8) = Format(NFe!DataSaidaEntrada, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC  ' data em branco = 30/12/1899
    ide(9) = IIf(NFe!TipoNF = "E", 0, 1)                           ' tipo do documento 0 - Entrada / 1 - Saida
    ide(10) = Parametros!CodigoIBGE                                ' código do município do IBGE de ocorrência do FG do ICMS (vide página 141 do manual)
    ide(11) = Left(NFe!NFeTipoEmissao, 1)                          ' forma de emissão da NF-e 1- normal, 2 - contingência FS, 3 - contingência SCAN, etc.
    ide(12) = Left(NFe!NFeFinalidadeEmissao, 1)                    ' finalidade da emissão da NF-e 1- NF-e normal
    If Not Vazio(NFe!NFCeChaveAcessoReferenciada) Then
       ide(13) = NFe!NFCeChaveAcessoReferenciada
    End If
    If ide(11) <> 1 Then
       ide(15) = IIf(Vazio(NFe!NFCeDataHoraContingencia), Null, NFe!NFCeDataHoraContingencia) ' Data/Hora Contingencia
       ide(16) = NFe!NFCeJustificativaContingencia                                            ' Justificativa Contingencia
    End If
    If ModeloNF = "55" Then
       ide(14) = IIf(Len(Retira(NFe!CPF_CNPJ, ".,-/", UM_A_UM)) > 11, 0, 1) 'Indica operação com Consumidor final
       ide(34) = IIf(NFe!UF = empUF, "1", IIf(NFe!UF = "EX", "3", "2"))     'Identificador de local de destino da operação - 1 - Operação interna|2 - Operação interestadual|3 - Operação com exterior
       ide(35) = "0"                                                        'Indicador de presença do comprador no estabelecimento comercial no momento da operação - 0 - Não se aplica|1 - Operação presencial|2 - Operação não presencial, pela Internet|3 - Operação não presencial, Teleatendimento|4 - NFC-e em operação com entrega a domicílio|9 - Operação não presencial, outros
    Else
       ide(14) = IIf(NFe!NFeConsumidorFinal, 1, 0)                          'Indica operação com Consumidor final
       ide(34) = Left(NFe!NFeIdentificadorDestino, 1)                       'Identificador de local de destino da operação - 1 - Operação interna|2 - Operação interestadual|3 - Operação com exterior
       ide(35) = Left(NFe!NFeIndicadorPresencaComprador, 1)                 'Indicador de presença do comprador no estabelecimento comercial no momento da operação - 0 - Não se aplica|1 - Operação presencial|2 - Operação não presencial, pela Internet|3 - Operação não presencial, Teleatendimento|4 - NFC-e em operação com entrega a domicílio|9 - Operação não presencial, outros
    End If
    ide(36) = "v. " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)  '<verProc> (versão do aplicativo)
    '
    '================grupo de identificação do destinatario (grupo E do Manual de integração - páginas 92)=======================
    '
    ReDim dest(39)
    If Len(NFe!CPF_CNPJ) = 0 Then dest(18) = "1"
    If Len(Retira(NFe!CPF_CNPJ, ".,-/", UM_A_UM)) > 11 Then
      dest(0) = Trim(NFe!CPF_CNPJ)                                        ' CNPJ do destinatario sem máscara de formatação
      dest(0) = Retira(dest(0), ".-/", UM_A_UM)                           ' CNPJ do destinatario sem máscara de formatação
    Else
      dest(0) = Trim(NFe!CPF_CNPJ)                                        ' CPF do destinatario, uso exclusivo do Fisco
      dest(0) = Retira(dest(0), ".-/", UM_A_UM)                           ' CPF do destinatario, uso exclusivo do Fisco
    End If
    dest(1) = RemoveAcento(Trim(Left(NFe!NomeRazSocial, 60)))             ' Razão social do destinatario, evitar caracteres acentuados e &
    dest(2) = RemoveAcento(Trim(NFe!ENDERECO))                            ' logradouro
    dest(3) = RemoveAcento(Trim(NFe!Num))                                 ' número, informar S/N quando inexistente para erro de Schema XML
    dest(4) = ""                                                          ' complemento do endereço, o conteúdo pode ser omitido
    dest(5) = RemoveAcento(Trim(NFe!Bairro))                              ' bairro
    dest(6) = Trim(NFe!CodigoIBGE)                                        ' código do município (vide página 141 do manual), deve ser compatível com a UF
    dest(7) = RemoveAcento(Trim(NFe!Municipio))                           ' nome do município
    dest(8) = Trim(NFe!UF)                                                ' sigla da UF
    dest(9) = Retira(NFe!CEP, ".-/", UM_A_UM)                             ' CEP - sem máscara
    dest(10) = NFe!CodigoPais                                             ' código do pais - deve fixo em 1058 - Brasil
    dest(11) = RemoveAcento(NFe!NomePais)                                 ' nome do pais (Brasil ou BRASIL)
    dest(12) = Trim(Retira(NFe!fone, "()-.", UM_A_UM))                    ' número do telefone sem máscara
    dest(12) = Retira(dest(12), " ", UM_A_UM)                             ' número do telefone sem máscara
    dest(12) = IIf(Left(dest(12), 1) = "0", Mid(dest(12), 2), dest(12))   ' número do telefone sem máscara
    If Len(dest(12)) = 0 Then dest(12) = ""
    dest(13) = Trim(Retira(NFe!InscEst, ".,-/", UM_A_UM))                 ' Inscrição Estadual do destinatario sem máscara
    dest(14) = ""                                                         ' Inscrição SUFRAMA
    dest(15) = ""                                                         ' Email
    If ModeloNF = "55" Then
       dest(37) = IIf(Vazio(NFe!InscEst), "9", "1")                       ' Indicador da IE do Destinatário - 1 - Contribuinte ICMS (informar a IE do destinatário)|2 - Contribuinte isento de Inscrição no cadastro de Contribuintes do ICMS|9 - Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS
    Else
       dest(37) = Left(NFe!NFeIndicadorIEDestinatario, 1)                 ' Indicador da IE do Destinatário - 1 - Contribuinte ICMS (informar a IE do destinatário)|2 - Contribuinte isento de Inscrição no cadastro de Contribuintes do ICMS|9 - Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS
    End If
    dest(17) = ""                                                         ' Inscrição Municipal do Tomador do Serviço
    
    If Left(Parametros!IdentificacaoAmbiente, 1) = 2 Then
       If ModeloNF = "55" Then dest(0) = "99999999000191"
       dest(1) = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
       dest(13) = ""
       If ModeloNF = "55" Then dest(37) = "2"
       If ModeloNF = "65" Then dest(37) = "9"
    End If
    
    If NFe!CodigoPais <> 1058 Then dest(0) = ""
       
    ReDim autXML(1)
    
    autXML(0) = "" 'Retira(Nfe!CPF_CNPJ, ".-/", UM_A_UM)
    
    If ModeloNF = "55" Then
'        vsSQL = "SELECT TbNotaFiscalProd_Itens.IdNFProd, TbNotaFiscalProd_Itens.IdNFProd_Item, TbNotaFiscalProd_Itens.CodProduto, TbNotaFiscalProd_Itens.IdProduto, TbNotaFiscalProd_Itens.DescricaoProduto, TbNotaFiscalProd_Itens.ValorOutras, " & _
'                "TbNotaFiscalProd_Itens.CodBarras, TbNotaFiscalProd_Itens.UN, TbNotaFiscalProd_Itens.CFOP, TbNotaFiscalProd_Itens.QtdeMov, TbNotaFiscalProd_Itens.ValorUnit, TbNotaFiscalProd_Itens.Desconto, TbNotaFiscalProd_Itens.Aliq_Icms AS Aliquota, " & _
'                "TbNotaFiscalProd_Itens.Bc_Icms, TbNotaFiscalProd_Itens.Bc_AliquotaReducao, TbNotaFiscalProd_Itens.Vlr_Icms,TbNotaFiscalProd_Itens.Aliq_IPI As AliqIPI, TbNotaFiscalProd_Itens.Vlr_IPI As ValorIPI, TbNotaFiscalProd_Itens.Valor_Frete As ValorFrete, " & _
'                "TbNotaFiscalProd_Itens.ICMSCST, TbNotaFiscalProd_Itens.PISCST, TbNotaFiscalProd_Itens.COFINSCST, TbNotaFiscalProd_Itens.IPICST, TbNotaFiscalProd_Itens.Codncm As NCM, TbNotaFiscalProd_Itens.ProdInfAdicional, TbNotaFiscalProd_Itens.ValorTributos, " & _
'                "TbNotaFiscalProd_Itens.BCSTRet, TbNotaFiscalProd_Itens.ICMSSTRet, TbNotaFiscalProd_Itens.BCImpostoImportacao, TbNotaFiscalProd_Itens.DespesasAduaneiras, TbNotaFiscalProd_Itens.ValorImpostoImportacao, TbNotaFiscalProd_Itens.ValorIOF " & _
'                "FROM TbNotaFiscalProd_Itens " & _
'                "WHERE TbNotaFiscalProd_Itens.IdNFProd = " & NumeroNota & " " & _
'                "ORDER BY TbNotaFiscalProd_Itens.IdNFProd, TbNotaFiscalProd_Itens.IdNFProd_Item"
    Else
        vsSQL = "SELECT TbNFCe_Itens.IdNFProd, TbNFCe_Itens.IdNFProd_Item, TbNFCe_Itens.CodProduto, TbNFCe_Itens.IdProduto, TbNFCe_Itens.DescricaoProduto, TbNFCe_Itens.ValorOutras, " & _
                "TbNFCe_Itens.CodBarras, TbNFCe_Itens.UN, TbNFCe_Itens.CFOP, TbNFCe_Itens.QtdeMov, TbNFCe_Itens.ValorUnit, TbNFCe_Itens.Desconto, TbNFCe_Itens.Aliq_Icms AS Aliquota, " & _
                "TbNFCe_Itens.Bc_Icms, TbNFCe_Itens.Bc_AliquotaReducao, TbNFCe_Itens.Vlr_Icms, TbNFCe_Itens.Aliq_IPI As AliqIPI, TbNFCe_Itens.Vlr_IPI As ValorIPI, TbNFCe_Itens.Valor_Frete As ValorFrete, " & _
                "TbNFCe_Itens.ICMSCST, TbNFCe_Itens.PISCST, TbNFCe_Itens.COFINSCST, TbNFCe_Itens.IPICST, TbNFCe_Itens.Codncm As NCM, TbNFCe_Itens.ProdInfAdicional, TbNFCe_Itens.ValorTributos, " & _
                "TbNFCe_Itens.BCSTRet, TbNFCe_Itens.ICMSSTRet, TbNFCe_Itens.BCImpostoImportacao, TbNFCe_Itens.DespesasAduaneiras, TbNFCe_Itens.ValorImpostoImportacao, TbNFCe_Itens.ValorIOF " & _
                "FROM TbNFCe_Itens " & _
                "WHERE TbNFCe_Itens.IdNFProd = " & NumeroNota & " " & _
                "ORDER BY TbNFCe_Itens.IdNFProd, TbNFCe_Itens.IdNFProd_Item"
    End If
    
    Set NFeItens = vgDb.OpenRecordset(vsSQL)
    
    n = NFeItens.RecordCount - 1
    
    ReDim prod(n, 140)
    
    'PDesconto = Format(Nfe!DescontoPromocional / (Nfe!Valor_NF_Prod - Nfe!DescontoPromocional), "######0.000000")
    
    If NFe!Valor_NF_Prod > 0 Then
       pFrete = Format((NFe!Valor_Frete / NFe!Valor_NF_Prod) * 100, "######0.000000")
       pOutras = Format((NFe!OutrasDespesasAces / NFe!Valor_NF_Prod) * 100, "######0.000000")
    End If
    
    For i = 0 To n
        '
        '================grupo de detalhe do produto (grupo I01 do Manual de integração - páginas 95)=======================
        '
        prod(i, 0) = Trim(NFeItens!CodProduto)                                        ' código do produto
        prod(i, 1) = IIf(Vazio(NFeItens!CodBarras), "", NFeItens!CodBarras)           ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN
        prod(i, 2) = RemoveAcento(Trim(NFeItens!DescricaoProduto))                    ' código do produto, espaços em branco consecutivos ou no início ou fim do campo podem gerar erro de Schema XML, além de caracteres reservados do XML <>&"'
        prod(i, 3) = NFeItens!NCM                                                     ' código NCM, pode ser omitido se não sujeito ao IPI
        prod(i, 109) = "" '"AA1000;BB1001;CC1002;DD1003"                              '<NVE>
        prod(i, 5) = Trim(Str(NFeItens!CFOP))                                         ' CFOP do operação, causa erro de XML se informado um código inexistente
        prod(i, 6) = RemoveAcento(Trim(NFeItens!UN))                                  ' unidade de comercialização
        prod(i, 7) = Format(NFeItens!QtdeMov, "######0.000")                          ' quantidade de comercialização
        prod(i, 7) = Substitui(prod(i, 7), ",", ".", UM_A_UM)
        prod(i, 8) = Format(NFeItens!ValorUnit, "######0.000")                        ' valor unitário de comercialização, campo de mera demonstração deve ser o resultado da divisão do vProd / qCom
        prod(i, 8) = Substitui(prod(i, 8), ",", ".", UM_A_UM)
        prod(i, 9) = Format(NFeItens!QtdeMov * NFeItens!ValorUnit, "######0.00")      ' valor do total do item
        prod(i, 9) = Substitui(prod(i, 9), ",", ".", UM_A_UM)
        prod(i, 10) = prod(i, 1)                                                      ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN, em geral é o mesmo código do EAN de comercialização
        prod(i, 11) = RemoveAcento(Trim(NFeItens!UN))                                 ' unidade de tributação, na maioria dos casos é idêntico  ao vUnCom, pode diferente nos casos de produtos sujeitos a ST em que a unidade de pauta é diferente da unidade de comercialização
                                                                                      ' Ex. unidade de comercialização = 1 pack de lata de cerveja => unidade de tributação = 1 lata (preço de pauta)
        prod(i, 12) = Format(NFeItens!QtdeMov, "######0.000")                         ' quantidade de comercialização
        prod(i, 12) = Substitui(prod(i, 12), ",", ".", UM_A_UM)
        prod(i, 13) = Format(NFeItens!ValorUnit, "######0.000")                       ' valor unitário de tributação, campo de mera demonstração deve ser o resultado da divisão do vProd / qTrib
        prod(i, 13) = Substitui(prod(i, 13), ",", ".", UM_A_UM)
        prod(i, 14) = Format(NFeItens!ValorFrete, "######0.00")                       ' valor do frete, se cobrado do cliente deve ser rateado entre os itens de produto
        prod(i, 14) = Substitui(prod(i, 14), ",", ".", UM_A_UM)
        prod(i, 15) = Format(0, "######0.00")                                         ' valor do seguro, se cobrado do cliente deve ser rateado entre os itens de produto
        prod(i, 15) = Substitui(prod(i, 15), ",", ".", UM_A_UM)
        prod(i, 16) = Format(NFeItens!Desconto, "######0.00")                         ' valor do desconto concedido
        prod(i, 16) = Substitui(prod(i, 16), ",", ".", UM_A_UM)
        prod(i, 96) = Format(NFeItens!ValorOutras, "######0.00")                      ' valor das outras despesas
        prod(i, 96) = Substitui(prod(i, 96), ",", ".", UM_A_UM)
        If NFeItens!ValorTributos > 0 Then
           prod(i, 53) = RemoveAcento(Trim(Left(NFeItens!ProdInfAdicional, 400))) & " - Valor Aproximado dos Tributos R$ " & Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)        ' informações adicionais do produto
           vlTrib = vlTrib + NFeItens!ValorTributos
        Else
           prod(i, 53) = RemoveAcento(Trim(Left(NFeItens!ProdInfAdicional, 500)))     ' informações adicionais do produto
        End If
        
        If prod(i, 5) = "1603" Then
           prod(i, 76) = 0                                                               ' Indica se o valor do item entra no valor total da NFe
        Else
           prod(i, 76) = 1                                                               ' Indica se o valor do item entra no valor total da NFe
        End If
        
        'Valor aproximado total de tributos federais, estaduais e municipais
        If NFeItens!ValorTributos > 0 Then prod(i, 104) = Substitui(Format(NFeItens!ValorTributos, "#0.00"), ",", ".", UM_A_UM)
        
        '
        '=========dados do ICMS (grupo N01 do Manual de integração - páginas 100)=====================
        '
        prod(i, 17) = Left(NFeItens!ICMSCST, 1)                         ' Tabela A - origem da mercadoria 0=nacional
        
        prod(i, 18) = Right(NFeItens!ICMSCST, IIf(emit(14) = 1, 3, 2))   ' Tabela B - CST=00-tributação normal
        prod(i, 19) = 3                                                  ' modalidade de determinação da BC = 3-valor da operação
        prod(i, 20) = Format(NFeItens!Bc_Icms, "######0.00")             ' valor da BC do ICMS = vProd + vFrete + vSeguro
        prod(i, 20) = Substitui(prod(i, 20), ",", ".", UM_A_UM)
        prod(i, 21) = Format(NFeItens!Aliquota, "######0.00")            ' alíquota do ICMS
        prod(i, 21) = Substitui(prod(i, 21), ",", ".", UM_A_UM)
        prod(i, 22) = Format(NFeItens!Vlr_Icms, "######0.00")            ' valor do ICMS
        prod(i, 22) = Substitui(prod(i, 22), ",", ".", UM_A_UM)
        prod(i, 46) = "5"                                                ' modalidade de determinação da BC ICMS ST
        prod(i, 47) = "" 'Format(0, "######0.00")                            ' percentual de valor de margem e valor adicionado
        'prod(i, 47) = Substitui(prod(i, 47), ",", ".", UM_A_UM)
        prod(i, 48) = "" 'Format(0, "######0.00")                            ' percentual de redução da BC do ICMS ST
        'prod(i, 48) = Substitui(prod(i, 48), ",", ".", UM_A_UM)
        prod(i, 49) = Format(NFeItens!BCSTRet, "######0.00")             ' BC do ICMS ST
        prod(i, 49) = Substitui(prod(i, 49), ",", ".", UM_A_UM)
        prod(i, 50) = Format(0, "######0.00")                            ' percentual do ICMSST
        prod(i, 50) = Substitui(prod(i, 50), ",", ".", UM_A_UM)
        prod(i, 51) = Format(NFeItens!ICMSSTRet, "######0.00")           ' valor do ICMS ST devido
        prod(i, 51) = Substitui(prod(i, 51), ",", ".", UM_A_UM)
        prod(i, 52) = Format(NFeItens!Bc_AliquotaReducao, "######0.00")  ' percentual de redução da BC
        prod(i, 52) = Substitui(prod(i, 52), ",", ".", UM_A_UM)
        
        If prod(i, 18) = "20" Or prod(i, 18) = "30" Or prod(i, 18) = "40" Or prod(i, 18) = "70" Or prod(i, 18) = "90" Then
           'Prod(i, 88) = Format(0, "######0.00")                    ' vICMSDeson
           'Prod(i, 88) = Substitui(Prod(i, 88), ",", ".", UM_A_UM)
           prod(i, 85) = "9"                                        ' motDesICMS
        End If
        
        If emit(14) = "1" Then
           prod(i, 17) = "0"                                             ' Tabela A - origem da mercadoria 0=nacional
           prod(i, 77) = "0.00"                                          ' <pCredSN>          Simples Nacional
           prod(i, 78) = "0.00"                                          ' <vCredICMSSN>      Simples Nacional
        End If
        
        'tag IPI
        prod(i, 23) = NFeItens!IPICST                                                    'IPI <CST>
        prod(i, 24) = Format(NFeItens!QtdeMov * NFeItens!ValorUnit, "######0.00")        'IPI <vBC>
        prod(i, 24) = Substitui(prod(i, 24), ",", ".", UM_A_UM)                          'IPI <vBC>
        prod(i, 25) = Format(NFeItens!AliqIPI, "######0.00")                             'IPI <pIPI>
        prod(i, 25) = Substitui(prod(i, 25), ",", ".", UM_A_UM)                          'IPI <pIPI>
        prod(i, 26) = Format(NFeItens!ValorIPI, "######0.00")                            'IPI <vIPI>
        prod(i, 26) = Substitui(prod(i, 26), ",", ".", UM_A_UM)                          'IPI <vIPI>
        
        'tag II
        If ModeloNF = "55" Then
           prod(i, 27) = Format(NFeItens!BCImpostoImportacao, "#0.00")                      'II <vBC>
           prod(i, 27) = Substitui(prod(i, 27), ",", ".", UM_A_UM)
           prod(i, 28) = Format(NFeItens!DespesasAduaneiras, "#0.00")                       'II <vDespAdu>
           prod(i, 28) = Substitui(prod(i, 28), ",", ".", UM_A_UM)
           prod(i, 29) = Format(NFeItens!ValorImpostoImportacao, "#0.00")                   'II <vII>
           prod(i, 29) = Substitui(prod(i, 29), ",", ".", UM_A_UM)
           prod(i, 30) = Format(NFeItens!ValorIOF, "#0.00")                                 'II <vIOF>
           prod(i, 30) = Substitui(prod(i, 30), ",", ".", UM_A_UM)
        End If
        
        '
        '=========dados do PIS (grupo Q do Manual de Integração - páginas 110) =============
        '
        prod(i, 31) = IIf(Vazio(NFeItens!PISCST) Or IsNull(NFeItens!PISCST), "07", NFeItens!PISCST)
        prod(i, 32) = Format(NFeItens!QtdeMov * NFeItens!ValorUnit, "######0.00")
        prod(i, 32) = Substitui(prod(i, 32), ",", ".", UM_A_UM)
        prod(i, 33) = Format(Parametros!PISAliquota, "###0.00")
        prod(i, 33) = Substitui(prod(i, 33), ",", ".", UM_A_UM)
        prod(i, 34) = Format(Round((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!PISAliquota / 100), 2), "###0.00")
        
        Select Case prod(i, 31)
           Case "04", "06", "07", "08", "09"
              vlPIS = vlPIS
           Case Else
              vlPIS = vlPIS + prod(i, 34)
        End Select

        prod(i, 34) = Substitui(prod(i, 34), ",", ".", UM_A_UM)
        prod(i, 45) = "0.00"
         
        'tag PISST
        prod(i, 54) = ""
        prod(i, 55) = ""
        prod(i, 56) = ""
        
        '
        '========dados do COFINS (grupo s do Manual de Integração - páginas 113) ============
        '
        prod(i, 35) = IIf(Vazio(NFeItens!COFINSCST) Or IsNull(NFeItens!COFINSCST), "07", NFeItens!COFINSCST)
        prod(i, 36) = Format((NFeItens!QtdeMov * NFeItens!ValorUnit), "######0.00")
        prod(i, 36) = Substitui(prod(i, 36), ",", ".", UM_A_UM)
        prod(i, 37) = Format(Parametros!COFINSAliquota, "###0.00")
        prod(i, 37) = Substitui(prod(i, 37), ",", ".", UM_A_UM)
        prod(i, 38) = Format(Round((NFeItens!QtdeMov * NFeItens!ValorUnit) * (Parametros!COFINSAliquota / 100), 2), "###0.00")
        
        Select Case prod(i, 35)
           Case "04", "06", "07", "08", "09"
              vlCOFINS = vlCOFINS
           Case Else
              vlCOFINS = vlCOFINS + prod(i, 38)
        End Select

        prod(i, 38) = Substitui(prod(i, 38), ",", ".", UM_A_UM)
        prod(i, 44) = "0.00"
           
        
        'tag COFINSST
        prod(i, 57) = ""
        prod(i, 58) = ""
        prod(i, 59) = ""
        
        'Tag da Declaração de Importação
        If ModeloNF = "55" Then
            SQL = "SELECT IdNFProd_Item_Seq, DI_Numero, DI_Data, DI_UF_Desembarque, DI_Local_Desembarque, " & _
                  "DI_Data_Desembarque, DI_Codigo_Exportador " & _
                  "FROM TbNotaFiscalProd_Itens_DI " & _
                  "WHERE IdNFProd = " & NFeItens!IdNFProd & " AND IdNFProd_Item = " & NFeItens!IdNFProd_Item & " " & _
                  "ORDER BY IdNFProd_Item_Seq"
            Set NFeDeclaracaoImposto = vgDb.OpenRecordset(SQL)
            If NFeDeclaracaoImposto.RecordCount > 0 Then
                prod(i, 60) = NFeDeclaracaoImposto!DI_Numero                                  'nDI
                prod(i, 61) = Format(NFeDeclaracaoImposto!DI_Data, "yyyy-mm-dd")              'dDI
                prod(i, 62) = NFeDeclaracaoImposto!DI_Local_Desembarque                       'xLocDesemb
                prod(i, 63) = NFeDeclaracaoImposto!DI_UF_Desembarque                          'UFDesemb
                prod(i, 64) = Format(NFeDeclaracaoImposto!DI_Data_Desembarque, "yyyy-mm-dd")  'dDesemb
                prod(i, 65) = NFeDeclaracaoImposto!DI_Codigo_Exportador                       'cExportador
            
                SQL = "SELECT IdNFProd_Item_Seq_Item, AD_Numero, AD_Fabricante, AD_Desconto " & _
                      "FROM TbNotaFiscalProd_Itens_DI_ADI " & _
                      "WHERE IdNFProd = " & NFeItens!IdNFProd & " And IdNFProd_Item = " & NFeItens!IdNFProd_Item & " And IdNFProd_Item_Seq = " & NFeDeclaracaoImposto!IdNFProd_Item_Seq & " " & _
                      "ORDER BY IdNFProd_Item_Seq_Item"
                Set NFeAdicao = vgDb.OpenRecordset(SQL)
                If NFeAdicao.RecordCount > 0 Then
                    prod(i, 66) = NFeAdicao!AD_Numero                        'adi: nAdicao
                    prod(i, 67) = NFeAdicao!IdNFProd_Item_Seq_Item           'adi: nSeqAdic
                    prod(i, 68) = NFeAdicao!AD_Fabricante                    'adi: cFabricante
                    If NFeAdicao!AD_Desconto > 0 Then
                       prod(i, 69) = Format(NFeAdicao!AD_Desconto, "#0.00")  'adi: vDescDI
                       prod(i, 69) = Substitui(prod(i, 69), ",", ".", UM_A_UM)
                    End If
                    NFeAdicao.MoveNext
                End If
                NFeDeclaracaoImposto.MoveNext
            End If
        End If
        
        NFeItens.MoveNext
    Next
    'MsgBox "Grupo de Tributos/COFINS " + COFINS, vbInformation, "Resultado"
    '
    '   atualização de total
    '
    ReDim tot(35)
    tot(0) = Format(NFe!BaseCalc_ICMS, "######0.00")
    tot(0) = Substitui(tot(0), ",", ".", UM_A_UM)
    tot(1) = Format(NFe!Valor_ICMS, "######0.00")
    tot(1) = Substitui(tot(1), ",", ".", UM_A_UM)
    tot(2) = Format(NFe!BaseCalc_ICSM_Subst, "######0.00")
    tot(2) = Substitui(tot(2), ",", ".", UM_A_UM)
    tot(3) = Format(NFe!Valor_ICMS_Subst, "######0.00")
    tot(3) = Substitui(tot(3), ",", ".", UM_A_UM)
    tot(4) = Format(NFe!Valor_NF_Prod, "######0.00")
    tot(4) = Substitui(tot(4), ",", ".", UM_A_UM)
    tot(5) = Format(NFe!Valor_Frete, "######0.00")
    tot(5) = Substitui(tot(5), ",", ".", UM_A_UM)
    tot(6) = Format(NFe!Valor_Seguro, "######0.00")
    tot(6) = Substitui(tot(6), ",", ".", UM_A_UM)
    tot(7) = Format(NFe!DescontoPromocional, "######0.00")
    tot(7) = Substitui(tot(7), ",", ".", UM_A_UM)
    tot(8) = Format(NFe!ValorImpostoImportacao, "######0.00")
    tot(8) = Substitui(tot(8), ",", ".", UM_A_UM)
    tot(9) = Format(NFe!Valor_IPI, "######0.00")
    tot(9) = Substitui(tot(9), ",", ".", UM_A_UM)
    tot(10) = Format(vlPIS, "######0.00")
    tot(10) = Substitui(tot(10), ",", ".", UM_A_UM)
    tot(11) = Format(vlCOFINS, "######0.00")
    tot(11) = Substitui(tot(11), ",", ".", UM_A_UM)
    tot(12) = Format(NFe!OutrasDespesasAces, "######0.00")
    tot(12) = Substitui(tot(12), ",", ".", UM_A_UM)
    vlNF = (NFe!Valor_NF_Prod - NFe!DescontoPromocional) + (NFe!Valor_Frete + NFe!Valor_Seguro + NFe!Valor_IPI + NFe!OutrasDespesasAces + NFe!ValorImpostoImportacao + NFe!Valor_ICMS_Subst)
    tot(13) = Format((NFe!Valor_NF_Prod - NFe!DescontoPromocional) + (NFe!Valor_Frete + NFe!Valor_Seguro + NFe!Valor_IPI + NFe!OutrasDespesasAces + NFe!ValorImpostoImportacao + NFe!Valor_ICMS_Subst), "######0.00")
    tot(13) = Substitui(tot(13), ",", ".", UM_A_UM)
    If vlTrib > 0 Then tot(26) = Substitui(Format(vlTrib, "#0.00"), ",", ".", UM_A_UM)
    tot(27) = Format(0, "######0.00")
    tot(27) = Substitui(tot(27), ",", ".", UM_A_UM)
    
    'grupo ISSQN
    tot(14) = ""    'ISSQNtot <vServ>
    tot(15) = ""    'ISSQNtot <vBC>
    tot(16) = ""    'ISSQNtot <vISS>
    tot(17) = ""    'ISSQNtot <vPIS>
    tot(18) = ""    'ISSQNtot <vCOFINS>
    
    '
    '============dados do transportador
    '
    ReDim trp(28)
    trp(0) = IIf(Vazio(NFe!Frete_Por_Conta), 0, Left(NFe!Frete_Por_Conta, 1))        ' responsabilidade do frete 0-emitente, 1-destinatário
    If Not Vazio(NFe!NomeTrasnportador) Then
        If Len(Retira(NFe!CPF_CNPJ_Transp, ".-/", UM_A_UM)) > 11 Then
          trp(1) = Trim(NFe!CPF_CNPJ_Transp)                                         ' CNPJ da Transportadora sem mascara
          trp(1) = Retira(trp(1), ".-/", UM_A_UM)                            ' CNPJ da Transportadora sem mascara
        Else
          trp(1) = Trim(NFe!CPF_CNPJ_Transp)                                          ' CPF da Transportadora sem mascara
          trp(1) = Retira(trp(1), ".-/", UM_A_UM)                             ' CPF da Transportadora sem mascara
        End If
        trp(2) = RemoveAcento(Trim(NFe!NomeTrasnportador))
        trp(3) = Trim(Retira(NFe!InscEst_Trasnp, ".,-/", UM_A_UM))                   ' Inscrição Estadual da Transportadora sem máscara
        trp(4) = RemoveAcento(Trim(NFe!Endereco_Transp))
        trp(5) = RemoveAcento(Trim(NFe!Cidade_Transp))
        trp(6) = NFe!UF_Mot_Transp
    End If
    
    If Not Vazio(NFe!Placa_Veiculo) Then
       trp(7) = Retira(Trim(NFe!Placa_Veiculo), "-", UM_A_UM)
       trp(8) = NFe!UF_Trasnportador
       trp(15) = ""
    End If

    '  ============== criação dos lacres do volume
    '
    If NFe!Qtde_Trasnp > 0 Then
       trp(9) = NFe!Qtde_Trasnp                                         ' quantidade de volumes
       trp(9) = Substitui(trp(9), ",", ".", UM_A_UM)
       trp(10) = RemoveAcento(NFe!Especie_Transp)                       ' espécie dos volumes
       trp(11) = RemoveAcento(NFe!Marca_Trasnp)                         ' marca dos volumes
       trp(12) = NFe!Num_Transp                                         ' numeração dos volumes
       trp(13) = Format(NFe!PesoLiq_Transp, "#0.000")                   ' peso líquido
       trp(13) = Substitui(trp(13), ",", ".", UM_A_UM)
       trp(14) = Format(NFe!PesoBruto_Transp, "#0.000")                 ' peso bruto
       trp(14) = Substitui(trp(14), ",", ".", UM_A_UM)
    End If
    
    If ModeloNF = "55" Then
       vsSQL = "SELECT Count(IDParcela) as qt FROM TbNotaFiscalProd_Faturas WHERE idNFProd = " & NumeroNota
    Else
       vsSQL = "SELECT Count(IDParcela) as qt FROM TbNFCe_Faturas WHERE idNFProd = " & NumeroNota
    End If
    cob_numero_parcelas = SQLExecutaRetorno(vsSQL, "qt", 0) - 1
    
    Dim idparc As Integer, vTotalRecebido As Double, vTotalNF As Double, vTotalDinheiro As Double, vTotalOutras As Double
    idparc = 0
    If cob_numero_parcelas >= 0 Then
       ReDim cob(cob_numero_parcelas, 6)
       cob(0, 0) = ide(6)
       cob(0, 1) = tot(13)
       cob(0, 2) = tot(13)
       If ModeloNF = "55" Then
          vsSQL = "SELECT IDParcela, Vencimento, Valor FROM TbNotaFiscalProd_Faturas WHERE idNFProd = " & NumeroNota
       Else
          vsSQL = "SELECT IDParcela, Vencimento, Valor, TipoPgto, IdBandeira, CartaoNumeroAutorizacao " & _
                  "FROM TbNFCe_Faturas " & _
                  "WHERE idNFProd = " & NumeroNota
          vTotalRecebido = SQLExecutaRetorno("SELECT SUM(Valor) r FROM TbNFCe_Faturas WHERE IdNFProd = " & NumeroNota, "r", 0)
          vTotalDinheiro = SQLExecutaRetorno("SELECT SUM(Valor) r FROM TbNFCe_Faturas WHERE IdNFProd = " & NumeroNota & " AND TipoPgto = 'DH'", "r", 0)
          vTotalOutras = SQLExecutaRetorno("SELECT SUM(Valor) r FROM TbNFCe_Faturas WHERE IdNFProd = " & NumeroNota & " AND TipoPgto <> 'DH'", "r", 0)
       End If
       Set NFeParcelas = vgDb.OpenRecordset(vsSQL)
       vTotalNF = vlNF
       If ModeloNF = "55" Then
          Do While Not NFeParcelas.EOF
             cob(idparc, 3) = NFeParcelas!IDParcela
             cob(idparc, 4) = Format(NFeParcelas!Vencimento, "YYYY-MM-DD")
             cob(idparc, 5) = Format(NFeParcelas!Valor, "######0.00")
             cob(idparc, 5) = Substitui(cob(idparc, 5), ",", ".", UM_A_UM)
             idparc = idparc + 1
             NFeParcelas.MoveNext
          Loop
       Else
          ReDim pagList(4)
          Do While Not NFeParcelas.EOF
             '01 - Dinheiro|02 - Cheque|03 - Cartão de Crédito|04 - Cartão de Débito|05 - Crédito Loja|10 - Vale Alimentação|11 - Vale Refeição|12 - Vale Presente|13 - Vale Combustível|99 - Outros
             Select Case NFeParcelas!TipoPgto
                 Case "DH": pagList(0) = pagList(0) & "01;"                            'pag <tPag>
                 Case "CH": pagList(0) = pagList(0) & "02;"                            'pag <tPag>
                 Case "CC": pagList(0) = pagList(0) & "03;"                            'pag <tPag>
                 Case "CD": pagList(0) = pagList(0) & "04;"                            'pag <tPag>
                 Case "CT": pagList(0) = pagList(0) & "05;"                            'pag <tPag>
                 Case Else: pagList(0) = pagList(0) & "99;"                            'pag <tPag>
             End Select
             If (vTotalNF <> vTotalRecebido And NFeParcelas!TipoPgto = "DH") Then
                vTotalDinheiro = vTotalNF - vTotalOutras
                pagList(1) = pagList(1) & Substitui(Format(vTotalDinheiro, "######0.00"), ",", ".", UM_A_UM) & ";"  'pag <vPag>
             Else
                pagList(1) = pagList(1) & Substitui(Format(NFeParcelas!Valor, "######0.00"), ",", ".", UM_A_UM) & ";"  'pag <vPag>
             End If
             'If NFeParcelas!TipoPgto = "CC" Or NFeParcelas!TipoPgto = "CD" Then
                pagList(2) = pagList(2) & Retira(Parametros!CNPJ, ".-/ ", UM_A_UM) & ";"    'card <CNPJ>  Informar o CNPJ da Credenciadora de cartão de crédito / débito
                pagList(3) = pagList(3) & LPad(NFeParcelas!IdBandeira, 2, "0") & ";"       'card <tBand> 01 - Visa|02 - Mastercard|03 - American Express|04 - Sorocred|99 - Outros
                pagList(4) = pagList(4) & Trim(NFeParcelas!CartaoNumeroAutorizacao) & ";"   'card <cAut>  Identifica o número da autorização da transação da operação com cartão de crédito e/ou débito
             'End If
             idparc = idparc + 1
             NFeParcelas.MoveNext
          Loop
          If NFeParcelas.RecordCount = 0 Then
             pagList(0) = "01;"          'pag <tPag>
             pagList(1) = tot(13) & ";"  'pag <vPag>
          End If
       End If
    Else
       If ModeloNF = "65" Then
          ReDim pagList(4)
          cob_numero_parcelas = 0
          pagList(0) = "01;"          'pag <tPag>
          pagList(1) = tot(13) & ";"  'pag <vPag>
       End If
    End If
    
    '
    '============= informações adcionais
    '
    ReDim obs(12)
    obs(0) = ""   'RemoveAcento(Trim(NFe!InformacoesAdicionais))
    obs(1) = RemoveAcento(Trim(NFe!Linha1)) & " // " & RemoveAcento(Trim(NFe!Linha2)) & " // " & RemoveAcento(Trim(NFe!Linha3)) & " // " & RemoveAcento(Trim(NFe!Linha4)) & " // " & RemoveAcento(Trim(NFe!Linha5))
    If vlTrib > 0 And vlNF > 0 Then
       vlNF = (NFe!Valor_NF_Prod - NFe!DescontoPromocional) + (NFe!Valor_Frete + NFe!Valor_Seguro + NFe!Valor_IPI + NFe!OutrasDespesasAces + NFe!ValorImpostoImportacao)
       pTributos = Format((vlTrib / vlNF) * 100, "#0.00")
       obs(1) = obs(1) & " - Valor Aproximado dos Tributos R$ " & FormatoDecimal(Format(vlTrib, "#0.00")) & " (" & FormatoDecimal(pTributos) & "%) (Conforme Lei Fed. 12.741/2012) Fonte: IBPT"
    End If
    
    'tag exporta v2.03
    obs(2) = ""      'UFEmbarq
    obs(3) = ""      'xLocEmbarq
    
    'tag compra v2.03
    obs(4) = ""      'xNEmp
    obs(5) = ""      'xPed
    obs(6) = ""      'infCpl

Else
   GoTo Caifora
End If

dirXML = GetString(&H80000001, "nfce", "PathPrincipal")
If VbInDesign Then
   dirXML = "C:\nfce-app\"
   SaveString &H80000001, "nfce", "PathPrincipal", dirXML
End If
dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")


'pega o endereço do arquivo a ser gerado
If Not Existe(dirXML) Then MkDir dirXML
'gera a chave da nfe
Dim id_chave As String
Dim numero_nfe_gerado As String
numero_nfe_gerado = sistNFCe.GeraXML(ide(), emit(), dest(), prod(), tot(), trp(), cob(), obs(), autXML(), False)
id_chave = numero_nfe_gerado
numero_nfe_gerado = Replace(numero_nfe_gerado, "NFe", "")
NFeChaveAcesso = numero_nfe_gerado

If Not Vazio(NFeChaveAcesso) Then
   If ModeloNF = "55" Then
      vsSQL = "UPDATE TbNotaFiscalProd SET " & _
              "NFeChaveAcesso = '" & NFeChaveAcesso & "' " & _
              "WHERE IdNFProd = " & NumeroNota
   Else
      vsSQL = "UPDATE TbNFCe SET " & _
              "NFCeChaveAcesso = '" & NFeChaveAcesso & "' " & _
              "WHERE IdNFProd = " & NumeroNota
   End If
   vgDb.Execute vsSQL
End If

xCaminhoXML = dirXML & "nfe\arquivos\" & id_chave & ".xml"

NFeResposta = sistNFCe.AssinarArquivoXML(xCaminhoXML, "infNFe")

DoEvents

If InStr(NFeResposta, "Erro") > 0 Then
   NFeMotivo = NFeResposta
   GoTo NaoEnviou
End If
  
xCaminhoXML = dirXML & "nfe\arquivos\assinado\" & id_chave & "-assinado.xml"

XMLOK = sistNFCe.ValidarArquivoXML(xCaminhoXML, False, NFeValidate) 'PL_007a.zip

If Not XMLOK Then
   NFeMotivo = "Erro na Validação do XML, falha no Schema" & vbNewLine & NFeValidate
   GoTo Caifora
End If

xCaminhoXML = dirXML & "nfe\arquivos\gerados\" & id_chave & ".xml"    'dirXML & "nfe\lotes\" & LPad(ide(6), 12, "0") & "-env-lot.xml"

If Not PodeEnviar Then GoTo NaoEnviou

NFeResposta = sistNFCe.NfeAutorizacao(xCaminhoXML, True)

If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
   MsgBox "*** Aparentemente Ocorreram Erros na Recepção do Lote (nfeAutorizacao)***" _
   & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeAutorizacao"
   GoTo Caifora
End If

NFeMotivo = Parse(NFeResposta, "#")
If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo Caifora
NFeNumeroRecibo = Parse(NFeResposta, "#")
cStat = Right(Parse(NFeMotivo, "-"), 3)

If cStat <> 103 Then
   GoTo NaoEnviou
End If

vsSQL = "UPDATE TbNFCe SET " & _
        "NFCeRecibo = " & NFeNumeroRecibo & " " & _
        "WHERE IdNFProd = " & NumeroNota
vgDb.Execute vsSQL

DoEvents

consultaNFe:

   NFeResposta = sistNFCe.NfceRetAutorizacao(NFeNumeroRecibo, dirXML)
   
   If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
      MsgBox NFeResposta, vbExclamation, "Retorno Autorização"
      GoTo Caifora
   End If
   
   If InStr(NFeResposta, "217") > 0 Then
      Sleep 3000 ' Aguarda mais 3 segundos
      NFeResposta = sistNFCe.NfceRetAutorizacao(NFeNumeroRecibo, dirXML) 'refaz a consulta
      If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
         MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote (nfeRetAutorizacao)***" _
         & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeRetAutorizacao"
         GoTo Caifora
     End If
   End If

buscaNFe:

'Consulta Nfe
   NFeResposta = sistNFCe.NfceConsulta(NFeChaveAcesso)
   
   Dim NFeRespostaFinal As String
   NFeRespostaFinal = NFeResposta
      
   ' testa erro 217 Rejeição: NF-e não consta na base de dados da SEFAZ <dhRecbto xmlns="http://www.portalfiscal.inf.br/nfe">2015-03-29T09:56:36-03:00
   If InStr(NFeResposta, "217") > 0 Then
      Sleep 3000 ' aguarda mais 3 segundos
      NFeResposta = sistNFCe.NfceConsulta(NFeChaveAcesso) ' consulta novamente
         
      If InStr(NFeResposta, "Erro") > 0 Or InStr(NFeResposta, "Rejeicao") > 0 Or InStr(NFeResposta, "Rejeição") > 0 Then
         MsgBox "*** Aparentemente Ocorreram Erros na Consulta do Lote ***" _
         & vbLf & NFeResposta, vbExclamation, "PROCESSO INTERROMPIDO - NfeConsulta"
         GoTo Caifora
      End If
   End If

   If InStr(NFeResposta, "Erro 98") > 0 Then
      NFeMotivo = Parse(NFeResposta, "#")
      NFeNumeroProtocolo = ""
      GoTo Caifora
   End If
    
   On Error Resume Next
    cStat = Parse(NFeResposta, "-")
    NFeNumeroProtocolo = Right(NFeResposta, 15)
    nroRecibo = cStat

    If InStr(NFeMotivo, "Erro") > 0 Or (InStr(NFeMotivo, "Rejeição") > 0 Or InStr(NFeMotivo, "Rejeicao") > 0) Then GoTo Caifora
    
    If nroRecibo = 204 Or nroRecibo = 539 Then
       NFeChaveAcesso = Mid(nfeRetorno, InStr(nfeRetorno, "chNFe:") + 6, 44)
       nroRecibo = Mid(nfeRetorno, InStr(nfeRetorno, "nRec:") + 5)
       nroRecibo = Left(nroRecibo, Len(nroRecibo) - 1)
       NFeNumeroRecibo = Left(NFeNumeroRecibo, 15 - Len(nroRecibo)) + nroRecibo
       If Vazio(NFeChaveAcesso) Or Len(NFeNumeroRecibo) < 15 Then
          NFeMotivo = nfeRetorno
          GoTo Caifora
       End If
       vsSQL = "UPDATE TbNFCe SET " & _
               "NFCeChaveAcesso = '" & NFeChaveAcesso & "' " & _
               "NFCeRecibo = " & NFeNumeroRecibo & " " & _
               "WHERE IdNFProd = " & NumeroNota
       vgDb.Execute vsSQL
       GoTo buscaNFe
    ElseIf nroRecibo <> 100 And nroRecibo <> 301 Then
        NFeMotivo = nroRecibo + " - " + nfeRetorno
        GoTo NaoEnviou
    ElseIf nroRecibo = 105 And nroRecibo = 217 Then
        GoTo consultaNFe
    ElseIf nroRecibo = 100 Then
        nfeRetorno = "Nota Fiscal de Consumidor Eletronica Autorizado o Uso."
        NFeDataHora = Format(Now, "dd/mm/yyyy h:mm:ss")
        msgResultado = "Chave NF-e.: " + NFeChaveAcesso & vbCrLf
        msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
        msgResultado = msgResultado + "Recibo.: " + NFeNumeroRecibo & vbCrLf
        msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
        msgResultado = msgResultado + "Resposta da Fazenda.: " + nroRecibo + " - " & nfeRetorno
        
        ' mensagem de emissao  MsgBox msgResultado, vbInformation + vbOKOnly
                
        On Error Resume Next
        NFeResposta = sistNFCe.NfceConsulta(NFeChaveAcesso)
             
        vsSQL = "UPDATE TbNFCe SET " & _
                "NFCeChaveAcesso = '" & NFeChaveAcesso & "', " & _
                "NFCeProtocolo = " & NFeNumeroProtocolo & ", " & _
                "NFCeProtocoloDataHora = '" & NFeDataHora & "' " & _
                "WHERE IdNFProd = " & NumeroNota
        vgDb.Execute vsSQL
    End If

PodeSair:
Set sistNFCe = Nothing
Screen.MousePointer = vbDefault
TransmitirNFCe = True
Exit Function

NaoEnviou:
Set sistNFCe = Nothing

If PodeEnviar Then MsgBox NFeMotivo, vbCritical + vbOKOnly
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
TransmitirNFCe_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Falha (" & Err.Description & ")" & vbNewLine & "Em TransmitirNFCe no Módulo NFe_DLL", vbCritical, "Falha"
    Err.Clear
End Function

Public Function CancelaNFe(ChaveAcesso As Variant, Protocolo As Variant, Justificativa As Variant, GravaProtocolo As Boolean) As Boolean  'Função para envio do cancelamento da NFe
Dim IdLote As Long, dhEvento As String
Dim sistNFe As snfe.Util
   
   Screen.MousePointer = vbHourglass
   Set sistNFe = New snfe.Util

   IdLote = Int(Rnd(918274 * 999) * 1000)
   dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
   NFeResposta = sistNFe.NfeRecepcaoEvento("Cancelamento", IdLote, ChaveAcesso, dhEvento, Protocolo, Justificativa, "1")
   
   If InStr(NFeResposta, "Erro 92") > 0 Then
      NFeMotivo = Parse(NFeResposta, "#")
      NFeNumeroProtocolo = ""
      GoTo Caifora
   End If
   
   cStat = Parse(NFeResposta, "#")
   NFeMotivo = Parse(NFeResposta, "#")
   iRetorno = Parse(NFeResposta, "#")
   NFeValidate = Parse(NFeResposta, "#")
   NFeNumeroProtocolo = Parse(NFeResposta, "#")
   NFeDataHora = Parse(NFeResposta, "#")
   If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
   If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
   If iRetorno = 135 Or iRetorno = 155 Then
      GoTo continua
   Else
      If iRetorno > 0 Then
         MsgBox Str$(iRetorno) & " - " & NFeValidate, vbInformation, "Cancelar NFe"
      Else
         MsgBox Str$(cStat) & " - " & NFeMotivo, vbInformation, "Cancelar NFe"
      End If
      GoTo Caifora
   End If
   
continua:
   msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
   msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
   msgResultado = msgResultado + "Resposta da Fazenda.: " + Str(iRetorno) & " - " & NFeValidate
   
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
Dim IdLote As Long, dhEvento As String
Dim sistNFCe As snfe.Util
   
   Screen.MousePointer = vbHourglass
   
   Set sistNFCe = New snfe.Util

   'NFeResposta = sistnfe.NfeCancelamento(ChaveAcesso, Protocolo, Justificativa)

   IdLote = Int(Rnd(918274 * 999) * 1000)
   dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
   'NFeResposta = sistNFe.NfeRecepcaoEvento("Cancelamento", IdLote, ChaveAcesso, dhEvento, Protocolo, Justificativa)
   NFeResposta = sistNFCe.NfceRecepcaoEvento("Cancelamento", IdLote, ChaveAcesso, dhEvento, Protocolo, Justificativa, "1")
   
    If InStr(NFeResposta, "Erro 92") > 0 Or InStr(NFeResposta, "Falha") > 0 Or InStr(1, NFeResposta, "Erro") > 0 Then
       'NFeMotivo = Parse(NFeResposta, "#")
       'NFeNumeroProtocolo = ""
       MsgBox NFeResposta
       GoTo Caifora
    End If
   
   '((cStat == string.Empty) ? string.Empty : cStat) + "#" + ((xMotivo == string.Empty) ? string.Empty : xMotivo) + "#" + ((cStat2 == string.Empty) ? string.Empty : cStat2) + "#" + ((xMotivo2 == string.Empty) ? string.Empty : xMotivo2) + "#" + ((nProt == string.Empty) ? string.Empty : nProt) + "#" + ((dhRegEvento == string.Empty) ? string.Empty : dhRegEvento)
   'If InStr(1, NFeResposta, "Erro") > 0 Then GoTo caiFora
   'cStat = Parse(NFeResposta, "#")
   'NFeMotivo = Parse(NFeResposta, "#")
   'cStat2 = Parse(NFeResposta, "#")
   'NFeValidate = Parse(NFeResposta, "#")
   'NFeNumeroProtocolo = Parse(NFeResposta, "#")
   
   NFeNumeroProtocolo = Right(NFeResposta, 15)
   
   'NFeDataHora = Parse(NFeResposta, "#")
   
   NFeDataHora = Format(Now, "dd/mm/yyyy h:mm:ss")
   
   'If (NFeDataHora <> Empty) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
   'If cStat2 = 135 Or cStat2 = 155 Then
   '   GoTo continua
   'Else
   '   MsgBox CStr(cStat2) & " - " & NFeMotivo, vbInformation
    '  GoTo caiFora
   'End If
   
continua:
   'msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
   'msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
   'msgResultado = msgResultado + "Resposta da Fazenda.: " + Str(cStat2) & " - " & NFeValidate
   
   msgResultado = NFeResposta
   
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
   
  iRetorno = 0
  i = 0
  
buscaNFe:
  NFeResposta = sistNFe.NfeRetAutorizacao(Recibo)
   
  If InStr(NFeResposta, "Erro 98") > 0 Then
     msgResultado = Parse(NFeResposta, "#")
     NFeValidate = "ERRO"
     NFeNumeroProtocolo = ""
     GoTo Caifora
  End If
   
  On Error Resume Next
  cStat = Parse(NFeResposta, "#")
  NFeMotivo = Parse(NFeResposta, "#")
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
     msgResultado = Str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe NÃO LOCALIZADA"
     NFeResposta = sistNFe.NfeConsulta(ChaveAcesso)
    
     If InStr(NFeResposta, "Erro 93") > 0 Then
        msgResultado = Parse(NFeResposta, "#")
        NFeValidate = "ERRO"
        NFeNumeroProtocolo = ""
        GoTo Caifora
     End If
     cStat = Parse(NFeResposta, "#")
     NFeMotivo = Parse(NFeResposta, "#")
     If cStat = 613 Then
        NFeChaveAcesso = Mid(NFeMotivo, InStr(NFeMotivo, "Numerico da NF-e [") + 18, 44)
        NFeResposta = sistNFe.NfeConsulta(NFeChaveAcesso)
        cStat = Parse(NFeResposta, "#")
        NFeMotivo = Parse(NFeResposta, "#")
        If cStat = 100 Or cStat = 110 Then
           NFeDataHora = Parse(NFeResposta, "#")
           If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
           If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
           NFeNumeroProtocolo = Parse(NFeResposta, "#")
           ChaveAcesso = NFeChaveAcesso
           GoTo continuaConsulta
        Else
           GoTo Caifora
        End If
     Else
        GoTo Caifora
     End If
  ElseIf cStat = 239 Then
     msgResultado = Str$(cStat) + " - " + NFeMotivo
     NFeChaveAcesso = ChaveAcesso
     NFeValidate = "NFe/NFCe COM PROBLEMAS"
     GoTo Caifora
  End If
  NFeNumeroProtocolo = Parse(NFeResposta, "#")
  NFeDataHora = Parse(NFeResposta, "#")
  If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
  If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
  nroRecibo = Parse(NFeResposta, "#")
  nfeRetorno = Parse(NFeResposta, "#")
  NFeValidate = Parse(NFeResposta, "#")

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

  NFeResposta = sistNFe.NfeConsulta(ChaveAcesso)
  
  If InStr(NFeResposta, "Erro 93") > 0 Then
     msgResultado = Parse(NFeResposta, "#")
     NFeValidate = "ERRO"
     NFeNumeroProtocolo = ""
     GoTo Caifora
  End If

  nroRecibo = Parse(NFeResposta, "#")
  NFeMotivo = Parse(NFeResposta, "#")
  NFeDataHora = Parse(NFeResposta, "#")
  If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
  If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
  NFeNumeroProtocolo = Parse(NFeResposta, "#")

continuaConsulta:
  msgResultado = "Chave NF-e.: " + ChaveAcesso & vbCrLf
  msgResultado = msgResultado + "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
  msgResultado = msgResultado + "Recibo.: " + Recibo & vbCrLf
  msgResultado = msgResultado + "Data e Hora.: " + NFeDataHora & vbCrLf
  msgResultado = msgResultado + "Resposta da Fazenda.: " + nroRecibo + " - " & NFeMotivo
  If nroRecibo = 100 Then
     NFeValidate = "NFe AUTORIZADA"
     NFeChaveAcesso = ChaveAcesso
  ElseIf nroRecibo = 101 Then
     NFeValidate = "NFe CANCELADA"
     NFeChaveAcesso = ChaveAcesso
  ElseIf nroRecibo = 110 Then
     NFeValidate = "NFe DENEGADA"
     NFeChaveAcesso = ChaveAcesso
     vsSQL = "UPDATE NotaFiscal SET " & _
             "Denegada = 1, " & _
             "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
             "DataHoraProtocolo = '" & NFeDataHora & "' " & _
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

Public Sub consultaNFe(ChaveAcesso As Variant) 'Sub que faz a consulta da NFe na Receita adaptada para nfe 3.1 ass 668
Dim sistNFe As snfe.Util

   Screen.MousePointer = vbHourglass

   Set sistNFe = New snfe.Util

   NFeResposta = sistNFe.NfeConsulta(ChaveAcesso)
   
   cStat = Parse(NFeResposta, "#")
   NFeMotivo = Parse(NFeResposta, "#")
   NFeDataHora = Parse(NFeResposta, "#")
   If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "#"), 25)
   If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
   NFeNumeroProtocolo = Parse(NFeResposta, "#")

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

   MsgBox msgResultado, vbInformation + vbOKOnly, NFeValidate

   Set sistNFe = Nothing

   Screen.MousePointer = vbDefault
End Sub

Public Sub ConsultaStatus(Optional ModeloNF As Integer = 55)  'Sub que consulta o Status do Serviço da Receita
Dim sistNFe As snfe.Util
   
   On Error GoTo deuErro
   
   Set sistNFe = New snfe.Util

   Screen.MousePointer = vbHourglass

   'NFe
   If ModeloNF = 55 Then
      Call sistNFe.NfeStatusServico(False)
      NFeResposta = CStr(sistNFe.retStatusWS.cStat) + " - " + sistNFe.retStatusWS.xMotivo
    End If
   'NFCe
   If ModeloNF = 65 Then
      Call sistNFe.NfceStatusServico(False)
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
Dim IdLote As Long, dhEvento As String
Dim dirXML As String
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

  Screen.MousePointer = vbHourglass
  
  IdLote = Int(Rnd(918274 * 999) * 1000)
  dhEvento = Format$(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
  'NFeResposta = sistNFe.NfeRecepcaoEvento("CCe", IdLote, ChaveAcesso, dhEvento, SeqCorrecao, TextoCorrecao)
  NFeResposta = sistNFe.NfeRecepcaoEvento("CCe", IdLote, ChaveAcesso, dhEvento, nProtocolo, TextoCorrecao, SeqCorrecao)
     
    If InStr(NFeResposta, "Erro 92") > 0 Then
       NFeMotivo = Parse(NFeResposta, "#")
       NFeNumeroProtocolo = ""
       GoTo Caifora
    End If
   
   '((cStat == string.Empty) ? string.Empty : cStat) + "#" + ((xMotivo == string.Empty) ? string.Empty : xMotivo) + "#" + ((cStat2 == string.Empty) ? string.Empty : cStat2) + "#" + ((xMotivo2 == string.Empty) ? string.Empty : xMotivo2) + "#" + ((nProt == string.Empty) ? string.Empty : nProt) + "#" + ((dhRegEvento == string.Empty) ? string.Empty : dhRegEvento)
   If InStr(1, NFeResposta, "Erro") > 0 Then GoTo Caifora
   cStat = Parse(NFeResposta, "#")
   NFeMotivo = Parse(NFeResposta, "#")
   cStat2 = Parse(NFeResposta, "#")
   NFeValidate = Parse(NFeResposta, "#")
   NFeNumeroProtocolo = Parse(NFeResposta, "#")
   NFeDataHora = Parse(NFeResposta, "#")
   If (NFeDataHora <> Empty) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
   If cStat2 = 135 Or cStat2 = 155 Then
      GoTo continua
   Else
      MsgBox CStr(cStat2) & " - " & NFeMotivo, vbInformation
      GoTo Caifora
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

Caifora:
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
  'NFeResposta = sistNFe.NfeManDest(Left(TipoEvento, 1), IdLote, CNPJ, ChaveAcesso, dhEvento, Justificativa)
  NFeResposta = sistNFe.NfeManDest(TipoEvento, IdLote, CNPJ, ChaveAcesso, dhEvento, Justificativa, "1")
     
  'RetornaValorTag(ret, "retEvento/infEvento/xMotivo") + ((nProt == string.Empty) ? string.Empty : "#" + nProt) + ((dhRegEvento == string.Empty) ? string.Empty : "#" + dhRegEvento) + ((cStat == string.Empty) ? string.Empty : "#" + cStat);
  NFeMotivo = Parse(NFeResposta, "#")
  NFeNumeroProtocolo = Parse(NFeResposta, "#")
  NFeDataHora = Parse(NFeResposta, "#")
  If InStr(NFeDataHora, "dhRecbto") > 0 Then NFeDataHora = Right(Parse(NFeDataHora, "-"), 25)
  If Not Vazio(NFeDataHora) Then NFeDataHora = Format(Left(NFeDataHora, 10), "dd/mm/yyyy") & " às " & Mid(NFeDataHora, 12, 8)
  cStat = Parse(NFeResposta, "#")
  If cStat = 135 Then
     GoTo continua
  Else
     MsgBox cStat & " - " & NFeMotivo, vbInformation, "ERRO"
     GoTo Caifora
  End If
         
continua:
  msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
  msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
  msgResultado = msgResultado + "Resposta da Fazenda.: " + Str(cStat) & " - " & NFeMotivo
    
  MsgBox msgResultado, vbInformation + vbOKOnly, "Envio CCe"
  
  Screen.MousePointer = vbDefault
  Set sistNFe = Nothing
  TransmitirManDest = True
  Exit Function

Caifora:
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
    Dim L As Long, lChar As Integer, sUtf8 As String
    For L = 1 To Len(sStr)
        lChar = AscW(Mid(sStr, L, 1))
        If lChar < 128 Then
            sUtf8 = sUtf8 + Mid(sStr, L, 1)
        ElseIf ((lChar > 127) And (lChar < 2048)) Then
            sUtf8 = sUtf8 + Chr(((lChar \ 64) Or 192))
            sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
        Else
            sUtf8 = sUtf8 + Chr(((lChar \ 144) Or 234))
            sUtf8 = sUtf8 + Chr((((lChar \ 64) And 63) Or 128))
            sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
        End If
    Next L
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
    
    NFeResposta = sistNFe.DownloadNFe(ChaveNFe, "", dirXML)
    
    If InStr(NFeResposta, "Erro") > 0 Then GoTo Caifora
    
    xCaminhoXML = NFeResposta
    
    DownloadXML = True
    
    Set sistNFe = Nothing
    
    Exit Function
Caifora:
    DownloadXML = False
    xCaminhoXML = ""
    Set sistNFe = Nothing
End Function
