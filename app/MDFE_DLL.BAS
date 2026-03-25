Attribute VB_Name = "MDFe_DLL"
'* Sistema...: Módulo MDFe
'* Empresa...: EkklesiaSoft Tecnologia em Sistemas
'* Módulo....: modMDFe
'* Função....: Módulo de funções do Manifesto
'* CopyRight.: (C)2019 EkklesiaSoft Tecnologia em Sistemas
'* Criação...: EkklesiaSoft Tecnologia em Sistemas
'* Data......: 06/08/2019 12:40:46
'* * * * * * *

Option Explicit                                   'requer variáveis explicitamente declaradas

Public MDFeXMLOK As Boolean, MDFeXMLAssinado As Boolean, MDFeXMLValidado As Boolean
Dim msgAlerta As String, msgErro As String, iRetorno As Boolean
Dim URLQRCode As String, cResultado As Long, msgResultado As String
Public Function ConfiguraDLLMDFe(Filial As Integer, ByRef objMDFe As MDFeDLL.Util) As Boolean
Dim ComandoSQL As String
Dim sistMDFe As MDFeDLL.Util
   
   On Error GoTo deuErro
   
   Set sistMDFe = New MDFeDLL.Util
   
   ComandoSQL = "SELECT CNPJ, Razao, Cidade, Estado, CodigoIBGE, CRT, AmbienteNF, DiretorioXMLMDFe, CertificadoDigital, " & _
                "NFCeIDToken, NFCeCSC, LicencaDLLMDFe, Email, caminho " & _
                "FROM Empresa"
                
   dirXML = SQLExecutaRetorno(ComandoSQL, "DiretorioXMLMDFe", App.path)
   dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
   
   msgAlerta = ""
   msgErro = ""
   
   iRetorno = sistMDFe.ConfigurarDLL("", _
                                     SQLExecutaRetorno(ComandoSQL, "CertificadoDigital", ""), _
                                     "", _
                                     True, _
                                     dirXML & "mdfe", _
                                     dirXML & "mdfe\schemas", _
                                     SQLExecutaRetorno(ComandoSQL, "Estado", "PI"), _
                                     58, _
                                     "1", _
                                     "30000", _
                                     Left(SQLExecutaRetorno(ComandoSQL, "AmbienteNF", 2), 1), _
                                     "1", _
                                     3, _
                                     SQLExecutaRetorno(ComandoSQL, "CNPJ", ""), _
                                     CNPJSoftHouse, _
                                     "OnLine Info", _
                                     SQLExecutaRetorno(ComandoSQL, "LicencaDLLMDFe", ""))
   
   iRetorno = sistMDFe.ConfigurarEmail("smtp.gmail.com", 587, 30000, True, "nferondosoft@gmail.com", "1234rondo", True, SQLExecutaRetorno(ComandoSQL, "Email", "ekklesiasoft@gmail.com"), SQLExecutaRetorno(ComandoSQL, "Razao", "OnLine Info"))
   
   iRetorno = sistMDFe.ConfigurarDAMDFe(SQLExecutaRetorno(ComandoSQL, "caminho", ""), "")
   
   iRetorno = sistMDFe.CarregarConfiguracoes
   
   Set objMDFe = sistMDFe
   
   Set sistMDFe = Nothing
   
   ConfiguraDLLMDFe = True
   
   Exit Function
                                       
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO - ConfiguraDLLMDFe"
   Err.Clear
   Set sistMDFe = Nothing
End Function



'Tipo: 1 - Gerar / 2 - Assinar/Validar / 3 - Transmitir
Public Function TransmitirMDFe(CodigoMDFe As Variant, Filial As Integer, Tipo As Integer) As Boolean
  Dim n As Integer
  Dim CIOT() As String, Contratante() As String, ValePedagio() As String
  Dim cMunCarrega As String, xMunCarrega As String, chCTe As String, chNFe As String, chMdfe As String, CondutorCPF As String, CondutorNome As String
  Dim certNumSerie As String, autXML As String, nLacres As String, nAverbacao As String, ufPercurso As String, ChaveMDFe As String, LicencaDLLMDFe As String, MDFeRecibo As String
  Dim rsMDFe As New ADODB.Recordset, rsMDFeCarrega As New ADODB.Recordset, rsMDFeDesCarrega As New ADODB.Recordset, rsMDFeCondutorRodoviario As New ADODB.Recordset
  Dim rsMDFeAutoDownload As New ADODB.Recordset, rsMDFeLacres As New ADODB.Recordset, rsMDFeSeguros As New ADODB.Recordset, rsMDFeAverbacao As New ADODB.Recordset
  Dim rsMDFePercurso As New ADODB.Recordset, rsEmitente As New ADODB.Recordset, rsMDFeDesCarregaDocumentos As New ADODB.Recordset
  
  Dim MDFe As String, mdfeEmit As String, mdfeIde As String, mdfeInfMunCarrega As String, mdfeInfModal As String, mdfeInfDoc As String, mdfeSeg As String, mdfeTot As String, mdfeLacres As String, mdfeAutXML As String, mdfeInfAdic As String, mdfeInfResp As String, mdfeinfMDFeSupl As String
  Dim mdfeInfANTT As String, mdfeVeicTracao As String, mdfeVeicReboque As String, mdfeInfMunDescarga As String
  
  On Error GoTo deuErro

  Dim objMDFeDLL As MDFeDLL.Util, objMDFeUtil As MDFe_Util.Util
  Set objMDFeDLL = New MDFeDLL.Util
  Set objMDFeUtil = New MDFe_Util.Util
  
  iRetorno = ConfiguraDLLMDFe(Filial, objMDFeDLL)
  If Not iRetorno Then GoTo Caifora
 
  'Manifesto
  vsSQL = "SELECT [CodigoManifesto],[Serie],[nMDF],[cMDF],[DataEmissao],[HoraEmissao],[MDFeTipoEmitente],[TipoEmissao],[TipoTransportador],[ModalidadeMDFe],[UFCarregamento],[UFDescarregamento],[DataInicioViagem], [HoraInicioViagem] " & _
          ",[QuantidadeCTe],[QuantidadeNFe],[QuantidadeMDFe],[ValorTotalCarga],[CodigoUnidade],[PesoBrutoCarga],[InformacoesAdicionaisFisco],[InformacoesComplementares] " & _
          ",[RNTRC],[CodigoAgendamentoPorto],[VeiculoTracaoCodInternoVeiculo],[VeiculoTracaoRenavam],[VeiculoTracaoPlaca],[VeiculoTracaoUF],[VeiculoTracaoTara],[VeiculoTracaoCapKG],[VeiculoTracaoCapM3] " & _
          ",[VeiculoTerceiro],[VeiculoTerceiroTipoProprietari],[VeiculoTerceiroNome],[VeiculoTerceiroCNPJCPF],[VeiculoTerceiroIE],[VeiculoTerceiroRNTRC],[VeiculoTerceiroUF], [VeiculoTracaoTipoRodado], [VeiculoTracaoTipoCarroceria] " & _
          ",[IdFilial],[MDFeXMLGerado],[MDFeXMLAssinado],[MDFeXMLValidado],[ChaveMDFe],[MDFeRecibo],[MDFeProtoloco],[MDFeProtocoloDataHora],[MDFeEnviada] " & _
          ",[MDFeCancelada],[MDFeCanceladaProtocolo],[MDFeCanceladaProtocoloDataHora],[MDFeCanceladaJustificativa] " & _
          ",[MDFeEncerrada],[MDFeEncerradaProtocolo],[MDFeEncerradaProtocoloDataHora] " & _
          "FROM [TbMDFe] " & _
          "WHERE [CodigoManifesto] = " & CodigoMDFe
  
  RsOpen rsMDFe, vsSQL
  If rsMDFe.RecordCount = 0 Then GoTo Caifora
  rsMDFe.MoveFirst
     
  RsOpen rsEmitente, "SELECT * FROM Empresa"
     
  LicencaDLLMDFe = rsEmitente!LicencaDLLMDFe
     
  MDFeXMLOK = rsMDFe!MDFeXMLGerado
  MDFeXMLAssinado = rsMDFe!MDFeXMLAssinado
  MDFeXMLValidado = rsMDFe!MDFeXMLValidado
  
  If MDFeXMLOK Then
     MDFeXMLOK = (Existe(dirXML & "mdfe\arquivos\gerados\" & rsMDFe!ChaveMDFe & "-mdfe.xml") = True)
  End If
  
  If MDFeXMLOK And MDFeXMLAssinado Then
     MDFeXMLAssinado = (Existe(dirXML & "mdfe\arquivos\assinado\MDFe" & rsMDFe!ChaveMDFe & "-assinado.xml") = True)
  End If
  
  If MDFeXMLOK Then
     If MsgBox("O XML do MDFe n° " & CStr(rsMDFe!nMDF) & " já foi gerado anteriormente!" & vbNewLine & "Substituir o XML gerado por um novo?", vbQuestion + vbYesNo + vbDefaultButton2, "XML GERADO") = vbNo Then
        If MDFeXMLAssinado Then
           If MsgBox("O XML do MDFe n° " & CStr(rsMDFe!nMDF) & " já foi assinado anteriormente!" & vbNewLine & "Substituir a assinatura do XML gerado por uma nova?", vbQuestion + vbYesNo + vbDefaultButton2, "ASSINAR XML") = vbYes Then
              GoTo AssinarXML
           Else
              If MDFeXMLValidado Then
                 If MsgBox("O XML do MDFe n° " & CStr(rsMDFe!nMDF) & " já foi validado!" & vbNewLine & "Validar o arquivo XML novamente?", vbQuestion + vbYesNo + vbDefaultButton2, "VALIDAR XML") = vbYes Then
                    GoTo ValidarXML
                 Else
                    If MsgBox("Transmitir o MDFe n° " & CStr(rsMDFe!nMDF) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "TRANSMITIR XML") = vbYes Then
                       GoTo Transmitir
                    Else
                       GoTo PodeSair
                    End If
                 End If
              End If
           End If
        Else
           vsSQL = "UPDATE TbMDFe SET MDFeXMLAssinado = 0, MDFeXMLValidado = 0 WHERE CodigoManifesto = " & CodigoMDFe
           SQLExecuta vsSQL
           GoTo AssinarXML
        End If
     End If
  End If
  
  Select Case Tipo
     Case 2:
        If MDFeXMLOK And (Not MDFeXMLAssinado Or Not MDFeXMLValidado) Then
            GoTo AssinarXML
        End If
     Case 3:
     If MDFeXMLOK And MDFeXMLAssinado And MDFeXMLValidado Then
        GoTo Transmitir
     End If
  End Select
  
  'Municipio Carregamento
  vsSQL = "SELECT [CodigoManifesto],[CodigoCidade], [TbCidades].[NomeCidade], [TbCidades].[UF], [TbCidades].[CodigoIBGE] " & _
          "FROM [TbMDFeCarregamento] INNER JOIN [TbCidades] ON [TbMDFeCarregamento].[CodigoCidade] = [TbCidades].[IDCidade] " & _
          "WHERE [CodigoManifesto] = " & CodigoMDFe & " ORDER BY [IdCidadeCarregamento]"
  mdfeInfMunCarrega = ""
  RsOpen rsMDFeCarrega, vsSQL
  If rsMDFeCarrega.RecordCount > 0 Then rsMDFeCarrega.MoveFirst
  Do While Not rsMDFeCarrega.EOF
     mdfeInfMunCarrega = mdfeInfMunCarrega + objMDFeUtil.infMunCarrega(rsMDFeCarrega!CodigoIBGE, rsMDFeCarrega!NomeCidade)
     rsMDFeCarrega.MoveNext
  Loop
  
  'Percuso
  vsSQL = "SELECT [CodigoManifesto], [UF] FROM [TbMDFePercurso] WHERE [CodigoManifesto] = " & CodigoMDFe
  ufPercurso = ""
  RsOpen rsMDFePercurso, vsSQL
  If rsMDFePercurso.RecordCount > 0 Then rsMDFePercurso.MoveFirst
  Do While Not rsMDFePercurso.EOF
     ufPercurso = ufPercurso & rsMDFePercurso!UF & ";"
     rsMDFePercurso.MoveNext
  Loop
  
  Dim xFone As String, MDFecMDF As String, MDFecDV As String, MDFecDF As String
  xFone = Retira(rsEmitente!Fone, "().-", UM_A_UM)
  xFone = Retira(IIf(Left(xFone, 1) = "0", Mid(xFone, 2), xFone), " ", UM_A_UM)                ' número do telefone sem máscara
  If Len(xFone) = 1 Then xFone = ""
  
  mdfeEmit = objMDFeUtil.emit(Retira(rsEmitente!CNPJ, ".-/ ", UM_A_UM), Retira(rsEmitente!InscEst, ".-/ ", UM_A_UM), RemoveAcento(rsEmitente!RazaoSocial), RemoveAcento(rsEmitente!Fantasia), RemoveAcento(rsEmitente!Endereco), rsEmitente!Num, "", RemoveAcento(rsEmitente!bairro), rsEmitente!CodigoIBGE, RemoveAcento(rsEmitente!Cidade), Retira(rsEmitente!CEP, "-. ", UM_A_UM), rsEmitente!UF, xFone, rsEmitente!Email)
  
  MDFecMDF = objMDFeDLL.GetHashCode
  cStat = objMDFeUtil.CriaChaveDFe(Left(rsEmitente!CodigoIBGE, 2), Format(rsMDFe!DataEmissao, "yy"), Format(rsMDFe!DataEmissao, "mm"), Retira(rsEmitente!CNPJ, "./- ", UM_A_UM), "58", rsMDFe!Serie, rsMDFe!nMDF, Left(rsMDFe!TipoEmissao, 1), CStr(MDFecMDF), msgResultado, MDFecDF, MDFecDV, ChaveMDFe)
  
  MDFecMDF = MDFecDF
  
  mdfeIde = objMDFeUtil.ide_v3a(CLng(Left(rsEmitente!CodigoIBGE, 2)), CLng(Left(rsEmitente!MDFeAmbiente, 1)), CLng(Left(rsMDFe!MDFeTipoEmitente, 1)), CLng(Left(rsMDFe!TipoTransportador, 1)), CLng(58), CLng(rsMDFe!Serie), CLng(rsMDFe!nMDF), CLng(MDFecMDF), CLng(MDFecDV), CLng(Left(rsMDFe!ModalidadeMDFe, 1)), Format(rsMDFe!DataEmissao, "yyyy-mm-dd") & "T" & Format(rsMDFe!HoraEmissao, "hh:mm:ss") & UTC, CLng(Left(rsMDFe!TipoEmissao, 1)), "0", "SISC v." & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision), rsMDFe!UFCarregamento, rsMDFe!UFDescarregamento, mdfeInfMunCarrega, "", Format(rsMDFe!DataInicioViagem, "yyyy-mm-dd") & "T" & Format(rsMDFe!HoraInicioViagem, "hh:mm:ss") + UTC, "", "")
  
  'Condutores Rodoviarios
  vsSQL = "SELECT [CodigoManifesto],[CondutorNome],[CondutorCPF] " & _
          "FROM [TbMDFeRodoviarioCondutores] " & _
          "WHERE [CodigoManifesto] = " & CodigoMDFe & " ORDER BY [IdMDFeCondutores]"
  CondutorNome = ""
  RsOpen rsMDFeCondutorRodoviario, vsSQL
  If rsMDFeCondutorRodoviario.RecordCount > 0 Then rsMDFeCondutorRodoviario.MoveFirst
  Do While Not rsMDFeCondutorRodoviario.EOF
     CondutorNome = CondutorNome & objMDFeUtil.condutor(RemoveAcento(rsMDFeCondutorRodoviario!CondutorNome), Retira(rsMDFeCondutorRodoviario!CondutorCPF, ".-/ ", UM_A_UM))
     rsMDFeCondutorRodoviario.MoveNext
  Loop
  
  mdfeVeicTracao = objMDFeUtil.veicPrincipal_v3("", Retira(rsMDFe!VeiculoTracaoPlaca, ".-/ ", UM_A_UM), rsMDFe!VeiculoTracaoRenavam, rsMDFe!VeiculoTracaoTara, rsMDFe!VeiculoTracaoCapKG, rsMDFe!VeiculoTracaoCapM3, "", CondutorNome, Left(rsMDFe!VeiculoTracaoTipoRodado, 2), Left(rsMDFe!VeiculoTracaoTipoCarroceria, 2), rsMDFe!VeiculoTracaoUF)
  
  mdfeInfModal = objMDFeUtil.rodo_v3("3.00", mdfeInfANTT, mdfeVeicTracao, mdfeVeicReboque, "", "")
  
  'Municipio Descarregamento
  vsSQL = "SELECT [CodigoManifesto],[CodigoCidade],[NomeCidade],[UF],[CodigoIBGE] " & _
          "FROM [TbMDFeDescargaDocumentos] INNER JOIN [TbCidades] ON [TbMDFeDescargaDocumentos].[CodigoCidade] = [TbCidades].[IDCidade] " & _
          "WHERE [CodigoManifesto] = " & CodigoMDFe & " ORDER BY [IdMDFeDescargasDocumentos]"
  
  RsOpen rsMDFeDesCarregaDocumentos, vsSQL
  If rsMDFeDesCarregaDocumentos.RecordCount > 0 Then rsMDFeDesCarregaDocumentos.MoveFirst
  Do While Not rsMDFeDesCarregaDocumentos.EOF
     cMunCarrega = rsMDFeDesCarregaDocumentos!CodigoIBGE
     xMunCarrega = RemoveAcento(rsMDFeDesCarregaDocumentos!NomeCidade)
     vsSQL = "SELECT [CodigoManifesto],[ModDocumento],[CodigoCidade],[NomeCidade],[UF],[CodigoIBGE],[ChaveAcesso],[SegCodBarra],[IndicadorReentrega] " & _
             "FROM [TbMDFeDescargaDocumentos] INNER JOIN [TbCidades] ON [TbMDFeDescargaDocumentos].[CodigoCidade] = [TbCidades].[IDCidade] " & _
             "WHERE [CodigoManifesto] = " & CodigoMDFe & " AND CodigoIBGE = " & rsMDFeDesCarregaDocumentos!CodigoIBGE & " ORDER BY [IdMDFeDescargasDocumentos]"
     RsOpen rsMDFeDesCarrega, vsSQL
     If rsMDFeDesCarrega.RecordCount > 0 Then rsMDFeDesCarrega.MoveFirst
     Do While Not rsMDFeDesCarrega.EOF
        If Left(rsMDFeDesCarrega!ModDocumento, 2) = "57" Then
           chCTe = chCTe & objMDFeUtil.infCTe_NT2018002(rsMDFeDesCarrega!ChaveAcesso, "", "", "", "", "")
        ElseIf Left(rsMDFeDesCarrega!ModDocumento, 2) = "55" Then
           chNFe = chNFe & objMDFeUtil.infNFe_v3(rsMDFeDesCarrega!ChaveAcesso, "", "", "", "")
        ElseIf Left(rsMDFeDesCarrega!ModDocumento, 2) = "58" Then
           chMdfe = chMdfe & objMDFeUtil.infMDFe(rsMDFeDesCarrega!ChaveAcesso, "", "", "")
        End If
        rsMDFeDesCarrega.MoveNext
     Loop
     mdfeInfMunDescarga = mdfeInfMunDescarga + objMDFeUtil.infMunDescarga(cMunCarrega, xMunCarrega, chCTe, chNFe, chMdfe)
     rsMDFeDesCarregaDocumentos.MoveNext
  Loop
  mdfeInfDoc = mdfeInfMunDescarga
  
  'Totais
  mdfeTot = objMDFeUtil.tot_v3(rsMDFe!QuantidadeCTe, rsMDFe!QuantidadeNFe, rsMDFe!QuantidadeMDFe, rsMDFe!ValorTotalCarga, Left(rsMDFe!CodigoUnidade, 2), rsMDFe!PesoBrutoCarga)
    
  'Autorizado para Download
  vsSQL = "SELECT [CodigoManifesto], [CNPJCPF] FROM [TbMDFeAutorizados] WHERE [CodigoManifesto] = " & CodigoMDFe & " ORDER BY [IdMDFeAutorizado]"
  autXML = ""
  RsOpen rsMDFeAutoDownload, vsSQL
  If rsMDFeAutoDownload.RecordCount > 0 Then rsMDFeAutoDownload.MoveFirst
  Do While Not rsMDFeAutoDownload.EOF
     mdfeAutXML = mdfeAutXML + objMDFeUtil.autXML(IIf(Len(Retira(rsMDFeAutoDownload!CNPJCPF, "-./ ", UM_A_UM)) = 11, "", rsMDFeAutoDownload!CNPJCPF), IIf(Len(Retira(rsMDFeAutoDownload!CNPJCPF, "-./ ", UM_A_UM)) = 11, rsMDFeAutoDownload!CNPJCPF, ""))
     rsMDFeAutoDownload.MoveNext
  Loop
  
  'Informações
  mdfeInfAdic = objMDFeUtil.infAdic(RemoveAcento(rsMDFe!InformacoesAdicionaisFisco), RemoveAcento(rsMDFe!InformacoesComplementares))
  
  mdfeinfMDFeSupl = objMDFeUtil.infMDFeSupl("https://dfe-portal.svrs.rs.gov.br/mdfe/qrCode", ChaveMDFe, CLng(Left(rsEmitente!MDFeAmbiente, 1)), SQLExecutaRetorno("SELECT CertificadoDigital FROM TbFiliais WHERE IdFilial = " & rsMDFe!IdFilial, "CertificadoDigital", ""), URLQRCode, cResultado, msgResultado)
  
  'Gerar o Arquivo XML
  MDFe = objMDFeUtil.MDFe_v3a("3.00", ChaveMDFe, mdfeIde, mdfeEmit, mdfeInfModal, mdfeInfDoc, "", mdfeTot, "", mdfeAutXML, mdfeInfAdic, "", mdfeinfMDFeSupl)
  
  xCaminhoXML = dirXML & "mdfe\arquivos\gerados\" & ChaveMDFe & "-mdfe.xml"
  
  n = FreeFile(20)
  Open xCaminhoXML For Output Shared As #n
     Print #n, MDFe
  Close #n
    
  MDFeXMLOK = objMDFeDLL.CarregarXML(xCaminhoXML)
  
  vsSQL = "UPDATE [TbMDFe] SET [ChaveMDFe] = '" & ChaveMDFe & "', [MDFeXMLGerado] = 1 WHERE [CodigoManifesto] = " & CodigoMDFe '[MDFeXMLAssinado],[MDFeXMLValidado]
  SQLExecuta vsSQL
  GoTo PodeSair
  
  
AssinarXML:
  ChaveMDFe = SQLExecutaRetorno("SELECT ChaveMDFe FROM TbMDFe WHERE CodigoManifesto = " & CodigoMDFe, "ChaveMDFe", "")
  xCaminhoXML = dirXML & "mdfe\arquivos\gerados\" & ChaveMDFe & "-mdfe.xml"
  iRetorno = objMDFeDLL.CarregarConfiguracoes()
  iRetorno = objMDFeDLL.AssinarXML(xCaminhoXML, xCaminhoXMLAuxiliar)
  msgAlerta = objMDFeDLL.mensagemAlerta
  msgErro = objMDFeDLL.mensagemErro
  If Not iRetorno Then GoTo Caifora
  If Not Vazio(msgAlerta) Then GoTo Caifora
  If Not Vazio(msgErro) Then GoTo Caifora
  MDFeXMLAssinado = True
  vsSQL = "UPDATE [TbMDFe] SET [MDFeXMLAssinado] = 1 WHERE [CodigoManifesto] = " & CodigoMDFe '[MDFeXMLValidado]
  SQLExecuta vsSQL
  
  
ValidarXML:
  ChaveMDFe = SQLExecutaRetorno("SELECT ChaveMDFe FROM TbMDFe WHERE CodigoManifesto = " & CodigoMDFe, "ChaveMDFe", "")
  If Vazio(xCaminhoXMLAuxiliar) Then xCaminhoXMLAuxiliar = dirXML & "mdfe\arquivos\assinado\MDFe" & ChaveMDFe & "-assinado.xml"
  iRetorno = objMDFeDLL.ValidarXML(xCaminhoXMLAuxiliar, False, msgResultado)
  msgAlerta = objMDFeDLL.mensagemAlerta
  msgErro = objMDFeDLL.mensagemErro
  If Not iRetorno Then GoTo Caifora
  If Not Vazio(msgAlerta) Then GoTo Caifora
  If Not Vazio(msgErro) Then GoTo Caifora
  MDFeXMLValidado = True
  vsSQL = "UPDATE [TbMDFe] SET [MDFeXMLValidado] = 1 WHERE [CodigoManifesto] = " & CodigoMDFe
  SQLExecuta vsSQL
  GoTo PodeSair

Transmitir:
  ChaveMDFe = SQLExecutaRetorno("SELECT ChaveMDFe FROM TbMDFe WHERE CodigoManifesto = " & CodigoMDFe, "ChaveMDFe", "")
  xCaminhoXML = dirXML & "mdfe\arquivos\gerados\" & ChaveMDFe & "-mdfe.xml"
  MDFeXMLOK = objMDFeDLL.CarregarXML(xCaminhoXML)
  iRetorno = objMDFeDLL.Enviar(CodigoMDFe, False, ChaveMDFe, MDFeRecibo)
  msgAlerta = objMDFeDLL.mensagemAlerta
  msgErro = objMDFeDLL.mensagemErro
  If Not iRetorno Then GoTo Caifora
  If Not Vazio(msgAlerta) Then GoTo Caifora
  If Not Vazio(msgErro) Then GoTo Caifora
  
  cStat = objMDFeDLL.cStat
  NFeMotivo = objMDFeDLL.xMotivo
  
  vsSQL = "UPDATE [TbMDFe] SET [MDFeRecibo] = " & MDFeRecibo & " WHERE [CodigoManifesto] = " & CodigoMDFe
  SQLExecuta vsSQL
  
  iRetorno = objMDFeDLL.ConsultarRecibo(MDFeRecibo)
  If Not iRetorno Then GoTo Caifora
  If Not Vazio(msgAlerta) Then GoTo Caifora
  If Not Vazio(msgErro) Then GoTo Caifora
  
  cStat = objMDFeDLL.cStat
  NFeMotivo = objMDFeDLL.xMotivo
  
  If cStat = 100 Then
     cStat2 = objMDFeDLL.cStat
     NFeValidate = objMDFeDLL.xMotivo
     NFeNumeroProtocolo = objMDFeDLL.nProtocolo
     
     iRetorno = objMDFeDLL.ConsultarProtocolo(ChaveMDFe)
     If Not iRetorno Then GoTo Caifora
     If Not Vazio(msgAlerta) Then GoTo Caifora
     If Not Vazio(msgErro) Then GoTo Caifora
  End If
  
PodeSair:
  TransmitirMDFe = True
  Set objMDFeUtil = Nothing
  Set objMDFeDLL = Nothing
  Exit Function
  
Caifora:
 If Not Vazio(msgAlerta) Then MsgBox "Alerta: " & msgAlerta, vbExclamation + vbOKOnly, "Alerta: TransmitirMDFe"
 If Not Vazio(msgErro) Then MsgBox "Erro: " & msgErro, vbExclamation + vbOKOnly, "Erro: TransmitirMDFe"
 Set objMDFeUtil = Nothing
 Set objMDFeDLL = Nothing
 TransmitirMDFe = False
 Exit Function
 
 Resume
 
deuErro:
  MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO TransmitirMDFe"
  Set objMDFeUtil = Nothing
  Set objMDFeDLL = Nothing
  TransmitirMDFe = False
End Function

Public Sub ConsultarStatusWSMDFe(Filial As Integer)
Dim objMDFeDLL As MDFeDLL.Util
 On Error GoTo deuErro
    
   Set objMDFeDLL = New MDFeDLL.Util
    
   iRetorno = ConfiguraDLLMDFe(Filial, objMDFeDLL)
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora

   iRetorno = objMDFeDLL.ConsultaStatusWS()
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora

   MsgBox CStr(objMDFeDLL.cStat) & " - " & objMDFeDLL.xMotivo, vbInformation + vbOKOnly, "STATUS WS MDFe"

   Set objMDFeDLL = Nothing

   Exit Sub

Caifora:
 If Not Vazio(msgAlerta) Then MsgBox "Alerta: " & msgAlerta, vbExclamation + vbOKOnly, "Alerta: Status WS MDFe"
 If Not Vazio(msgErro) Then MsgBox "Erro: " & msgErro, vbExclamation + vbOKOnly, "Erro: Status WS MDFe"
 Set objMDFeDLL = Nothing
 Exit Sub
 
deuErro:
  MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO Status WS MDFe"
  Set objMDFeDLL = Nothing

End Sub

Public Sub MDFeReciboConsulta(NumRecibo As Double, Filial As Integer)
Dim objMDFeDLL As MDFeDLL.Util
 On Error GoTo deuErro
    
   Set objMDFeDLL = New MDFeDLL.Util
    
   iRetorno = ConfiguraDLLMDFe(Filial, objMDFeDLL)
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora

   iRetorno = objMDFeDLL.ConsultarRecibo(Trim(Str(NumRecibo)))
   msgAlerta = objMDFeDLL.mensagemAlerta
   msgErro = objMDFeDLL.mensagemErro
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora
   
   cStat = objMDFeDLL.cStat
   NFeMotivo = objMDFeDLL.xMotivo
   NFeNumeroProtocolo = objMDFeDLL.nProtocolo
   NFeDataHora = objMDFeDLL.dhRecbto

   MsgBox CStr(objMDFeDLL.cStat) & " - " & objMDFeDLL.xMotivo, vbInformation + vbOKOnly, "CONSULTA RECIBO MDFe"

   Set objMDFeDLL = Nothing

   Exit Sub

Caifora:
 If Not Vazio(msgAlerta) Then MsgBox "Alerta: " & msgAlerta, vbExclamation + vbOKOnly, "Alerta: CONSULTA RECIBO MDFe"
 If Not Vazio(msgErro) Then MsgBox "Erro: " & msgErro, vbExclamation + vbOKOnly, "Erro: CONSULTA RECIBO MDFe"
 Set objMDFeDLL = Nothing
 Exit Sub
 
deuErro:
  MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO CONSULTA RECIBO MDFe"
  Set objMDFeDLL = Nothing
End Sub

Public Sub MDFeProtocoloConsulta(ChaveMDFe As String, Filial As Integer)
Dim objMDFeDLL As MDFeDLL.Util
 On Error GoTo deuErro
    
   Set objMDFeDLL = New MDFeDLL.Util
    
   iRetorno = ConfiguraDLLMDFe(Filial, objMDFeDLL)
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora

   iRetorno = objMDFeDLL.ConsultarProtocolo(ChaveMDFe)
   msgAlerta = objMDFeDLL.mensagemAlerta
   msgErro = objMDFeDLL.mensagemErro
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora
   
   cStat = objMDFeDLL.cStat
   NFeMotivo = objMDFeDLL.xMotivo
   NFeNumeroProtocolo = objMDFeDLL.nProtocolo
   NFeDataHora = objMDFeDLL.dhRecbto

   'MsgBox CStr(objMDFeDLL.cStat) & " - " & objMDFeDLL.xMotivo, vbInformation + vbOKOnly, "CONSULTA PROTOCOLO MDFe"

   Set objMDFeDLL = Nothing

   Exit Sub

Caifora:
 If Not Vazio(msgAlerta) Then MsgBox "Alerta: " & msgAlerta, vbExclamation + vbOKOnly, "Alerta: CONSULTA PROTOCOLO MDFe"
 If Not Vazio(msgErro) Then MsgBox "Erro: " & msgErro, vbExclamation + vbOKOnly, "Erro: CONSULTA PROTOCOLO MDFe"
 Set objMDFeDLL = Nothing
 Exit Sub
 
deuErro:
  MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO CONSULTA PROTOCOLO MDFe"
  Set objMDFeDLL = Nothing
End Sub

Public Function MDFeCondutorIncluir(ChaveMDFe As String, SeqEvento As Long, CondutorCPF As String, CondutorNome As String, Filial As Integer) As Boolean
Dim objMDFeDLL As MDFeDLL.Util
 On Error GoTo deuErro
    
   Set objMDFeDLL = New MDFeDLL.Util
    
   iRetorno = ConfiguraDLLMDFe(Filial, objMDFeDLL)
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora

   dirXML = SQLExecutaRetorno("SELECT DiretorioXMLMDFe FROM Empresa", "DiretorioXMLMDFe", "")
   dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
      
   xCaminhoXML = dirXML & "mdfe\arquivos\gerados\" & ChaveMDFe & "-mdfe.xml"
    
   iRetorno = objMDFeDLL.IncluirCondutor(xCaminhoXML, SeqEvento, Retira(CondutorCPF, ".-/ ", UM_A_UM), RemoveAcento(CondutorNome))
   msgAlerta = objMDFeDLL.mensagemAlerta
   msgErro = objMDFeDLL.mensagemErro
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora
   
   cStat = objMDFeDLL.cStat
   NFeMotivo = objMDFeDLL.xMotivo
   NFeNumeroProtocolo = objMDFeDLL.nProtocolo
   NFeDataHora = objMDFeDLL.dhRecbto

   'MsgBox CStr(objMDFeDLL.cStat) & " - " & objMDFeDLL.xMotivo, vbInformation + vbOKOnly, "CONDUTOR INCLUIR MDFe"
   MDFeCondutorIncluir = True

   Set objMDFeDLL = Nothing

   Exit Function

Caifora:
 If Not Vazio(msgAlerta) Then MsgBox "Alerta: " & msgAlerta, vbExclamation + vbOKOnly, "Alerta: CONDUTOR INCLUIR MDFe"
 If Not Vazio(msgErro) Then MsgBox "Erro: " & msgErro, vbExclamation + vbOKOnly, "Erro: CONDUTOR INCLUIR MDFe"
 Set objMDFeDLL = Nothing
 Exit Function
 
deuErro:
  MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO CONDUTOR INCLUIR MDFe"
  Set objMDFeDLL = Nothing
End Function

Public Function MDFeEncerrar(ChaveMDFe As String, SeqEvento As Long, Protocolo As String, Filial As Integer) As Boolean
Dim objMDFeDLL As MDFeDLL.Util
 On Error GoTo deuErro
    
   Set objMDFeDLL = New MDFeDLL.Util
    
   iRetorno = ConfiguraDLLMDFe(Filial, objMDFeDLL)
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora

   dirXML = SQLExecutaRetorno("SELECT DiretorioXMLMDFe FROM Empresa", "DiretorioXMLMDFe", "")
   dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
      
   xCaminhoXML = dirXML & "mdfe\arquivos\gerados\" & ChaveMDFe & "-mdfe.xml"
    
   iRetorno = objMDFeDLL.EventoEncerramento(xCaminhoXML, SeqEvento, Protocolo)
   msgAlerta = objMDFeDLL.mensagemAlerta
   msgErro = objMDFeDLL.mensagemErro
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora
   
   cStat = objMDFeDLL.cStat
   NFeMotivo = objMDFeDLL.xMotivo
   NFeNumeroProtocolo = objMDFeDLL.nProtocolo
   NFeDataHora = objMDFeDLL.dhRecbto

   'MsgBox CStr(objMDFeDLL.cStat) & " - " & objMDFeDLL.xMotivo, vbInformation + vbOKOnly, "CONDUTOR INCLUIR MDFe"

   MDFeEncerrar = True

   Set objMDFeDLL = Nothing

   Exit Function

Caifora:
 If Not Vazio(msgAlerta) Then MsgBox "Alerta: " & msgAlerta, vbExclamation + vbOKOnly, "Alerta: ENCERRAR MDFe"
 If Not Vazio(msgErro) Then MsgBox "Erro: " & msgErro, vbExclamation + vbOKOnly, "Erro: ENCERRAR MDFe"
 Set objMDFeDLL = Nothing
 Exit Function
 
deuErro:
  MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO ENCERRAR MDFe"
  Set objMDFeDLL = Nothing
End Function

Public Function MDFeCancelar(ChaveMDFe As String, SeqEvento As Long, Protocolo As String, Justificativa As String, Filial As Integer) As Boolean
Dim objMDFeDLL As MDFeDLL.Util
 On Error GoTo deuErro
    
   Set objMDFeDLL = New MDFeDLL.Util
    
   iRetorno = ConfiguraDLLMDFe(Filial, objMDFeDLL)
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora

   dirXML = SQLExecutaRetorno("SELECT DiretorioXMLMDFe FROM Empresa", "DiretorioXMLMDFe", "")
   dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
      
   xCaminhoXML = dirXML & "mdfe\arquivos\gerados\" & ChaveMDFe & "-mdfe.xml"
    
   iRetorno = objMDFeDLL.CancelarMDFe(xCaminhoXML, 1, Protocolo, Justificativa)
   msgAlerta = objMDFeDLL.mensagemAlerta
   msgErro = objMDFeDLL.mensagemErro
   If Not iRetorno Then GoTo Caifora
   If Not Vazio(msgAlerta) Then GoTo Caifora
   If Not Vazio(msgErro) Then GoTo Caifora
   
   cStat = objMDFeDLL.cStat
   NFeMotivo = objMDFeDLL.xMotivo
   NFeNumeroProtocolo = objMDFeDLL.nProtocolo
   NFeDataHora = objMDFeDLL.dhRecbto

   'MsgBox CStr(objMDFeDLL.cStat) & " - " & objMDFeDLL.xMotivo, vbInformation + vbOKOnly, "CONDUTOR INCLUIR MDFe"

   MDFeCancelar = True

   Set objMDFeDLL = Nothing

   Exit Function

Caifora:
 If Not Vazio(msgAlerta) Then MsgBox "Alerta: " & msgAlerta, vbExclamation + vbOKOnly, "Alerta: CANCELAR MDFe"
 If Not Vazio(msgErro) Then MsgBox "Erro: " & msgErro, vbExclamation + vbOKOnly, "Erro: CANCELAR MDFe"
 Set objMDFeDLL = Nothing
 Exit Function
 
deuErro:
  MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO CANCELAR MDFe"
  Set objMDFeDLL = Nothing
End Function
