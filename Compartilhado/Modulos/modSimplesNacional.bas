Attribute VB_Name = "modSimplesNacional"
Option Explicit

'--- Contexto da nota em emissao (preenchido por SN_IniciarNota, uma vez por NF-e/NFC-e) ---
Private mAtivo As Boolean          'regras do Simples valem nesta nota
Private mCRT As Long
Private mTpRegime As Long
Private mAno As Long
Private mComp As String            'AAAAMM da emissao
Private mCAtivPadrao As String
Private mAnexoPadrao As String
Private mRBT12 As Double
Private mTemRBT12 As Boolean
Private mNItensTrib As Long        'itens que levaram gTribSN (total W59h so existe se > 0)
Private mErro As String
Private mCacheOk(0 To 4) As Boolean
Private mCacheC(0 To 4) As Double
Private mCacheI(0 To 4) As Double
Private mProdCod As Long           'cache do ultimo produto consultado
Private mProdAnexo As String
Private mProdAtiv As String
'==============================================================================
' Regras da NT 2026.005 (Simples Nacional na NF-e/NFC-e, IBS/CBS).
' Modulo SO de decisao: nao acessa banco nem a DLL, so recebe valores e responde.
' Quem monta o XML (modNFe) pergunta aqui o que informar para cada emitente.
'
' Entradas:
'   vCRT       = Empresa.CRT (1=Simples, 2=Simples excesso sublimite, 3=Normal, 4=MEI)
'   tpRegime   = tag tpRegimeIBSCBS (C21a) derivada por SN_TipoRegimeIBSCBS:
'                0 = nao informar | 1 = Simples com regime regular
'                2 = Normal, opcao pendente pelo CRT 1, sem regime regular
'                3 = Normal, opcao pendente pelo CRT 2, sem regime regular
'                4 = Normal, opcao pendente (CRT 1 ou 2), com regime regular
'
' Tabela da NT (itens 2.1 a 2.3):
'   CRT tpRegime  cAtivSN  gIBSCBS            gTribSN
'   1   1         sim      sim                nao
'   2   1         sim      sim                nao
'   1   0         sim      NAO                gIBSSN + gCBSSN
'   2   0         sim      sim (CBS zerada)   so gCBSSN
'   3   2         sim      sim                gIBSSN + gCBSSN
'   3   3         sim      sim                so gCBSSN
'   3   4         sim      sim                nao
'   3   0         nao      sim                nao      (Regime Normal, como sempre foi)
'   4   0         nao      NAO                nao      (MEI nao informa IBSCBS)
' CST 620 (combustivel monofasico): nao informar gTribSN (tratar no chamador).
'==============================================================================

'CRT como numero (a coluna pode vir como texto, nula ou com mais de 1 caractere)
Public Function SN_CRT(ByVal vCRT As Variant) As Long
    On Error Resume Next
    SN_CRT = Val(Left$(Trim$(vCRT & ""), 1))
End Function

'Valor da tag tpRegimeIBSCBS (0 = nao informar a tag)
Public Function SN_TipoRegimeIBSCBS(ByVal vCRT As Variant, ByVal bOptaRegular As Boolean, ByVal nOpcaoPendente As Long) As Long
    Dim nCRT As Long
    nCRT = SN_CRT(vCRT)
    SN_TipoRegimeIBSCBS = 0
    Select Case nCRT
        Case 1, 2
            If bOptaRegular Then SN_TipoRegimeIBSCBS = 1
        Case 3
            If nOpcaoPendente = 1 Or nOpcaoPendente = 2 Then
                If bOptaRegular Then
                    SN_TipoRegimeIBSCBS = 4
                ElseIf nOpcaoPendente = 1 Then
                    SN_TipoRegimeIBSCBS = 2
                Else
                    SN_TipoRegimeIBSCBS = 3
                End If
            End If
    End Select
End Function

'Emitente do Simples (ou com opcao pendente): informa cAtivSN em todos os itens
Public Function SN_InformaCAtivSN(ByVal vCRT As Variant, ByVal nTpRegime As Long) As Boolean
    Dim nCRT As Long
    nCRT = SN_CRT(vCRT)
    SN_InformaCAtivSN = (nCRT = 1 Or nCRT = 2 Or (nCRT = 3 And nTpRegime >= 2))
End Function

'Informa o grupo gIBSCBS (IBS UF, IBS Mun, CBS) no item?
Public Function SN_InformaGIBSCBS(ByVal vCRT As Variant, ByVal nTpRegime As Long) As Boolean
    Select Case SN_CRT(vCRT)
        Case 4
            SN_InformaGIBSCBS = False                 'MEI: nao informa o grupo IBSCBS
        Case 1
            SN_InformaGIBSCBS = (nTpRegime = 1)       'CRT 1 sem regime regular: so gTribSN
        Case Else
            SN_InformaGIBSCBS = True
    End Select
End Function

'CRT 2 sem regime regular: gIBSCBS vai com a aliquota de CBS zerada
Public Function SN_CBSZeradaNoGIBSCBS(ByVal vCRT As Variant, ByVal nTpRegime As Long) As Boolean
    SN_CBSZeradaNoGIBSCBS = (SN_CRT(vCRT) = 2 And nTpRegime = 0)
End Function

'Informa o grupo gTribSN (UB115a) no item e o total gTribSN?
Public Function SN_InformaGTribSN(ByVal vCRT As Variant, ByVal nTpRegime As Long) As Boolean
    Dim nCRT As Long
    nCRT = SN_CRT(vCRT)
    SN_InformaGTribSN = ((nCRT = 1 Or nCRT = 2) And nTpRegime = 0) Or (nCRT = 3 And (nTpRegime = 2 Or nTpRegime = 3))
End Function

'Dentro do gTribSN, informa tambem o gIBSSN (alem do gCBSSN, que e obrigatorio)?
Public Function SN_InformaIBSNoTribSN(ByVal vCRT As Variant, ByVal nTpRegime As Long) As Boolean
    Dim nCRT As Long
    nCRT = SN_CRT(vCRT)
    SN_InformaIBSNoTribSN = (nCRT = 1 And nTpRegime = 0) Or (nCRT = 3 And nTpRegime = 2)
End Function

'True = emitente cujo IBS/CBS segue exatamente o que o sistema sempre fez (Regime Normal puro)
Public Function SN_RegimeNormalPuro(ByVal vCRT As Variant, ByVal nTpRegime As Long) As Boolean
    SN_RegimeNormalPuro = (SN_CRT(vCRT) = 3 And nTpRegime = 0)
End Function

'------------------------------------------------------------------------------
' cAtivSN (I17d)
'------------------------------------------------------------------------------
Public Function SN_CAtivSNValido(ByVal sCodigo As String) As Boolean
    Select Case Trim$(sCodigo)
        Case "01", "02", "03", "04", "90"
            SN_CAtivSNValido = True
    End Select
End Function

Public Function SN_DescricaoCAtivSN(ByVal sCodigo As String) As String
    Select Case Trim$(sCodigo)
        Case "01": SN_DescricaoCAtivSN = "Revenda de mercadorias / venda de mercadorias industrializadas (Anexo I)"
        Case "02": SN_DescricaoCAtivSN = "Demais bens materiais, sem incidencia de ICMS (Anexo I)"
        Case "03": SN_DescricaoCAtivSN = "Industrializadas com incentivo da Zona Franca de Manaus (Anexo II)"
        Case "04": SN_DescricaoCAtivSN = "Demais bens, Zona Franca de Manaus, sem incidencia de ICMS (Anexo II)"
        Case "90": SN_DescricaoCAtivSN = "Operacoes nao tributaveis"
    End Select
End Function

'cAtivSN do item: o do produto; se vazio/invalido, o padrao da empresa; se tambem invalido, "01"
Public Function SN_ResolverCAtivSN(ByVal vProduto As Variant, ByVal vPadraoEmpresa As Variant) As String
    Dim s As String
    s = Trim$(vProduto & "")
    If SN_CAtivSNValido(s) Then
        SN_ResolverCAtivSN = s
        Exit Function
    End If
    s = Trim$(vPadraoEmpresa & "")
    If SN_CAtivSNValido(s) Then
        SN_ResolverCAtivSN = s
    Else
        SN_ResolverCAtivSN = "01"
    End If
End Function

'==============================================================================
' CALCULO DA ALIQUOTA EFETIVA DE IBS/CBS DO SIMPLES (pIBSSN / pCBSSN)
' Tabelas: SimplesNacionalAliquota e SimplesNacionalTetoISS (script 197), vigencia 2027 a 2033+.
' Os percentuais de pIBSSN/pCBSSN sao ALIQUOTAS EFETIVAS DO TRIBUTO:
'   efetiva = (RBT12 x nominal - deduzir) / RBT12      (da faixa do RBT12)
'   pCBSSN  = efetiva x partilha CBS da faixa / 100    (4 casas)
'   pIBSSN  = efetiva x partilha IBS da faixa / 100    (4 casas)
' Base legal: art. 21, III da Res. CGSN 140 (redacao da Res. CGSN 196/2026, art. 1o).
' RBT12 = receita bruta dos 12 meses anteriores ao mes anterior ao da apuracao.
'==============================================================================

'Aliquota efetiva (%) da faixa. RBT12 <= 0 (empresa nova): usa a nominal.
Public Function SN_AliquotaEfetiva(ByVal dRBT12 As Double, ByVal dNominal As Double, ByVal dDeduzir As Double) As Double
    If dRBT12 <= 0 Then
        SN_AliquotaEfetiva = dNominal
    Else
        SN_AliquotaEfetiva = (dRBT12 * dNominal / 100 - dDeduzir) / dRBT12 * 100
    End If
End Function

'Percentual efetivo de um tributo = efetiva x partilha / 100, com 4 casas
Public Function SN_PercTributo(ByVal dEfetiva As Double, ByVal dPartilha As Double) As Double
    SN_PercTributo = Round(dEfetiva * dPartilha / 100, 4)
End Function

'Valor do tributo no item = base x percentual / 100, com 2 casas
Public Function SN_ValorTributo(ByVal dBase As Double, ByVal dPerc As Double) As Double
    SN_ValorTributo = Round(dBase * dPerc / 100, 2)
End Function

'Calcula pCBSSN e pIBSSN a partir de uma linha da tabela de aliquotas
Public Sub SN_CalcularPercentuais(ByVal dRBT12 As Double, ByVal dNominal As Double, ByVal dDeduzir As Double, _
                                  ByVal dPartCBS As Double, ByVal dPartIBS As Double, _
                                  ByRef pCBSSN As Double, ByRef pIBSSN As Double)
    Dim dEfet As Double
    dEfet = SN_AliquotaEfetiva(dRBT12, dNominal, dDeduzir)
    pCBSSN = SN_PercTributo(dEfet, dPartCBS)
    pIBSSN = SN_PercTributo(dEfet, dPartIBS)
End Sub

'Anexo III e IV, 5a faixa: se a efetiva passar do limite, o ISS fica fixo no teto do ano e a
'diferenca (efetiva - teto) vai para CBS e IBS com os percentuais proprios (tabela SimplesNacionalTetoISS).
'Devolve True se aplicou (pCBSSN/pIBSSN recalculados). Sem linha na tabela (ex.: 2033) = nao se aplica.
Public Function SN_AplicaTetoISS(ByVal sAnexo As String, ByVal nAno As Long, ByVal nFaixa As Long, ByVal dEfetiva As Double, _
                                 ByRef pCBSSN As Double, ByRef pIBSSN As Double) As Boolean
    Dim rs As New ADODB.Recordset, sSQL As String
    Dim dTeto As Double, dLimite As Double, dPC As Double, dPI As Double
    SN_AplicaTetoISS = False
    If nFaixa <> 5 Then Exit Function
    sSQL = "SELECT TOP 1 TetoISS, LimiteEfetiva, PartCBS, PartIBS FROM SimplesNacionalTetoISS " & _
           "WHERE Anexo = '" & UCase$(Trim$(sAnexo)) & "' AND " & nAno & " BETWEEN AnoIni AND AnoFim"
    RsOpen rs, sSQL
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    dTeto = rs!TetoISS: dLimite = rs!LimiteEfetiva: dPC = rs!PartCBS: dPI = rs!PartIBS
    rs.Close
    If dEfetiva <= dLimite Then Exit Function
    pCBSSN = SN_PercTributo(dEfetiva - dTeto, dPC)
    pIBSSN = SN_PercTributo(dEfetiva - dTeto, dPI)
    SN_AplicaTetoISS = True
End Function

'Compra governamental 2027-2028: ente = Uniao (tpEnteGov=1) -> pIBSSN=0 e pCBSSN = IBS + CBS.
'Outros entes: cada um com o seu percentual.
Public Sub SN_AjustaCompraGov(ByVal nTpEnteGov As Long, ByRef pCBSSN As Double, ByRef pIBSSN As Double)
    If nTpEnteGov = 1 Then
        pCBSSN = pCBSSN + pIBSSN
        pIBSSN = 0
    End If
End Sub

'Le a linha da tabela (Anexo, ano, RBT12) e devolve os percentuais efetivos do item.
'nCRT decide a 6a faixa (RBT12 acima de 3,6 mi):
'  CRT 1 (sublimite ainda nao ultrapassado): CBS pela 6a faixa; IBS continua no DAS, calculado pela
'         regra da nota de rodape dos anexos: [(RBT12 x nominal 5a - deduzir 5a) / RBT12] x partilha IBS da 5a faixa.
'  CRT 2 (sublimite ultrapassado): IBS sai do DAS (pIBSSN = 0); so a CBS fica no gTribSN.
'Devolve False se nao achar a faixa (nao calcula nada; o chamador decide).
Public Function SN_PercentuaisDoAnexo(ByVal sAnexo As String, ByVal nAno As Long, ByVal nCRT As Long, ByVal dRBT12 As Double, _
                                      ByRef pCBSSN As Double, ByRef pIBSSN As Double) As Boolean
    Dim rs As New ADODB.Recordset, sSQL As String, sRbt As String
    Dim dNom As Double, dDed As Double, dPC As Double, dPI As Double, nFaixa As Long, dEfet As Double
    Dim dNom5 As Double, dDed5 As Double, dPI5 As Double, bTem5 As Boolean
    pCBSSN = 0: pIBSSN = 0
    sRbt = Replace(Format$(IIf(dRBT12 <= 0, 0, dRBT12), "0.00"), ",", ".")
    sSQL = "SELECT TOP 1 Faixa, Nominal, Deduzir, PartCBS, PartIBS FROM SimplesNacionalAliquota " & _
           "WHERE Anexo = '" & UCase$(Trim$(sAnexo)) & "' AND " & nAno & " BETWEEN AnoIni AND AnoFim " & _
           "AND " & sRbt & " BETWEEN RbtDe - 0.01 AND RbtAte ORDER BY Faixa"
    RsOpen rs, sSQL
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    nFaixa = rs!Faixa
    dNom = rs!Nominal: dDed = rs!Deduzir: dPC = rs!PartCBS: dPI = rs!PartIBS
    rs.Close
    dEfet = SN_AliquotaEfetiva(dRBT12, dNom, dDed)
    pCBSSN = SN_PercTributo(dEfet, dPC)
    pIBSSN = SN_PercTributo(dEfet, dPI)
    If nFaixa = 6 And nCRT = 1 Then
        sSQL = "SELECT Nominal, Deduzir, PartIBS FROM SimplesNacionalAliquota " & _
               "WHERE Anexo = '" & UCase$(Trim$(sAnexo)) & "' AND Faixa = 5 AND " & nAno & " BETWEEN AnoIni AND AnoFim"
        RsOpen rs, sSQL
        If Not rs.EOF Then
            dNom5 = rs!Nominal: dDed5 = rs!Deduzir: dPI5 = rs!PartIBS: bTem5 = True
        End If
        rs.Close
        If bTem5 Then pIBSSN = SN_PercTributo(SN_AliquotaEfetiva(dRBT12, dNom5, dDed5), dPI5)
    End If
    Call SN_AplicaTetoISS(sAnexo, nAno, nFaixa, dEfet, pCBSSN, pIBSSN)
    SN_PercentuaisDoAnexo = True
End Function

'==============================================================================
' EMISSAO: contexto da nota e dados por item (usado por modNFe)
' Vigencia: as regras do Simples so entram a partir de 01/01/2027. Para testar em
' homologacao antes disso, crie o arquivo vazio SN_HOMOLOG.txt na pasta do sistema.
'==============================================================================

'Le um campo do recordset sem estourar se a coluna nao existir (cliente sem o script) ou vier Null
Public Function SN_Campo(ByVal rs As ADODB.Recordset, ByVal sCampo As String, ByVal vPadrao As Variant) As Variant
    On Error GoTo semcampo
    If IsNull(rs.Fields(sCampo).Value) Then
        SN_Campo = vPadrao
    Else
        SN_Campo = rs.Fields(sCampo).Value
    End If
    Exit Function
semcampo:
    SN_Campo = vPadrao
End Function

Public Function SN_Vigente(ByVal dtEmissao As Date) As Boolean
    SN_Vigente = (dtEmissao >= DateSerial(2027, 1, 1))
    If Not SN_Vigente Then
        On Error Resume Next
        SN_Vigente = (Dir$(App.Path & "\SN_HOMOLOG.txt") <> "")
    End If
End Function

Private Function SN_IndiceAnexo(ByVal sAnexo As String) As Long
    Select Case UCase$(Trim$(sAnexo))
        Case "I": SN_IndiceAnexo = 0
        Case "II": SN_IndiceAnexo = 1
        Case "III": SN_IndiceAnexo = 2
        Case "IV": SN_IndiceAnexo = 3
        Case "V": SN_IndiceAnexo = 4
        Case Else: SN_IndiceAnexo = -1
    End Select
End Function

'Prepara o contexto da nota. Devolve o tpRegimeIBSCBS a informar (0 = nao informar a tag).
'rsEmpresa = SELECT * FROM Empresa; dtEmissao = data de emissao da nota.
Public Function SN_IniciarNota(ByVal rsEmpresa As ADODB.Recordset, ByVal vEmissao As Variant) As Long
    Dim bOpta As Boolean, nPend As Long, i As Long, rs As New ADODB.Recordset
    Dim dtEmi As Date
    mAtivo = False: mTpRegime = 0: mNItensTrib = 0: mErro = "": mTemRBT12 = False: mRBT12 = 0: mProdCod = -1
    For i = 0 To 4
        mCacheOk(i) = False
    Next i
    SN_IniciarNota = 0
    On Error GoTo erro
    dtEmi = CDate(vEmissao)
    mCRT = SN_CRT(SN_Campo(rsEmpresa, "CRT", ""))
    bOpta = CBool(SN_Campo(rsEmpresa, "OptaRegimeRegularIBSCBS", False))
    nPend = CLng(Val(SN_Campo(rsEmpresa, "SimplesOpcaoPendente", 0) & ""))
    mCAtivPadrao = Trim$(SN_Campo(rsEmpresa, "cAtivSNPadrao", "01") & "")
    mAnexoPadrao = Trim$(SN_Campo(rsEmpresa, "AnexoSNPadrao", "I") & "")
    If SN_IndiceAnexo(mAnexoPadrao) < 0 Then mAnexoPadrao = "I"
    mAno = Year(dtEmi)
    If mAno < 2027 Then mAno = 2027    'so chega aqui antes de 2027 com SN_HOMOLOG.txt: usa a tabela de 2027 (a unica carregada)
    mComp = Format$(dtEmi, "yyyymm")
    mTpRegime = SN_TipoRegimeIBSCBS(mCRT, bOpta, nPend)
    mAtivo = SN_Vigente(dtEmi) And (mCRT = 1 Or mCRT = 2 Or mCRT = 4 Or (mCRT = 3 And mTpRegime >= 2))
    If Not mAtivo Then Exit Function

    If SN_InformaGTribSN(mCRT, mTpRegime) Then
        'RBT12 da competencia; se nao existir, o da competencia anterior mais recente (nao trava a emissao)
        RsOpen rs, "SELECT TOP 1 RBT12 FROM SimplesNacionalApuracao WHERE Competencia <= '" & mComp & "' ORDER BY Competencia DESC"
        If Not rs.EOF Then
            mRBT12 = CDbl(rs!RBT12)
            mTemRBT12 = True
        End If
        rs.Close
    End If
    SN_IniciarNota = mTpRegime
    Exit Function
erro:
    mErro = "Falha ao preparar o Simples Nacional: " & Err.Description
    mAtivo = False
End Function

Public Function SN_NotaAtiva() As Boolean
    SN_NotaAtiva = mAtivo
End Function

Public Function SN_UltimoErro() As String
    SN_UltimoErro = mErro
End Function

Public Function SN_ItemInformaCAtivSN() As Boolean
    SN_ItemInformaCAtivSN = mAtivo And SN_InformaCAtivSN(mCRT, mTpRegime)
End Function

Public Function SN_ItemInformaGIBSCBS() As Boolean
    SN_ItemInformaGIBSCBS = SN_InformaGIBSCBS(mCRT, mTpRegime)
End Function

Public Function SN_ItemCBSZerada() As Boolean
    SN_ItemCBSZerada = SN_CBSZeradaNoGIBSCBS(mCRT, mTpRegime)
End Function

'gTribSN no item: depende do enquadramento e nao vale para combustivel monofasico (CST 620)
Public Function SN_ItemInformaGTribSN(ByVal sCST As String) As Boolean
    SN_ItemInformaGTribSN = mAtivo And SN_InformaGTribSN(mCRT, mTpRegime) And (Trim$(sCST) <> "620")
End Function

Public Function SN_ItemInformaIBSNoTribSN() As Boolean
    SN_ItemInformaIBSNoTribSN = SN_InformaIBSNoTribSN(mCRT, mTpRegime)
End Function

Public Sub SN_RegistraItemTrib()
    mNItensTrib = mNItensTrib + 1
End Sub

'True se algum item levou gTribSN (nesse caso o total W59h/W59i/W59j tem que sair)
Public Function SN_TotalGTribSN() As Boolean
    SN_TotalGTribSN = (mNItensTrib > 0)
End Function

'Anexo e cAtivSN cadastrados no produto (vazio = padrao da empresa). Guarda o ultimo produto.
Public Sub SN_DadosProduto(ByVal vCodProd As Variant, ByRef sAnexoProd As String, ByRef sCAtivProd As String)
    Dim lCod As Long, rs As New ADODB.Recordset
    sAnexoProd = "": sCAtivProd = ""
    lCod = CLng(Val(vCodProd & ""))
    If lCod = mProdCod Then
        sAnexoProd = mProdAnexo: sCAtivProd = mProdAtiv
        Exit Sub
    End If
    mProdCod = lCod: mProdAnexo = "": mProdAtiv = ""
    On Error GoTo semcolunas
    RsOpen rs, "SELECT AnexoSN, cAtivSN FROM produtos WHERE codigo = " & lCod
    If Not rs.EOF Then
        mProdAnexo = Trim$(SN_Campo(rs, "AnexoSN", "") & "")
        mProdAtiv = Trim$(SN_Campo(rs, "cAtivSN", "") & "")
    End If
    rs.Close
semcolunas:
    sAnexoProd = mProdAnexo: sCAtivProd = mProdAtiv
End Sub

'cAtivSN do item: o gravado no item (nota antiga), senao o do produto, senao o padrao da empresa
Public Function SN_CAtivDoItem(ByVal vItem As Variant, ByVal vProduto As Variant) As String
    If SN_CAtivSNValido(Trim$(vItem & "")) Then
        SN_CAtivDoItem = Trim$(vItem & "")
    Else
        SN_CAtivDoItem = SN_ResolverCAtivSN(vProduto, mCAtivPadrao)
    End If
End Function

Public Function SN_AnexoDoItem(ByVal sAnexoProd As String) As String
    If SN_IndiceAnexo(sAnexoProd) >= 0 Then
        SN_AnexoDoItem = UCase$(Trim$(sAnexoProd))
    Else
        SN_AnexoDoItem = mAnexoPadrao
    End If
End Function

'Percentuais efetivos do anexo na competencia: os informados pelo contador (se houver) ou o calculo.
Private Function SN_PercentuaisItem(ByVal sAnexo As String, ByRef pCBSSN As Double, ByRef pIBSSN As Double) As Boolean
    Dim idx As Long, rs As New ADODB.Recordset
    Dim bOvC As Boolean, bOvI As Boolean, dOvC As Double, dOvI As Double, nCRT As Long
    idx = SN_IndiceAnexo(sAnexo)
    If idx < 0 Then
        mErro = "Anexo do Simples Nacional invalido: " & sAnexo
        Exit Function
    End If
    If mCacheOk(idx) Then
        pCBSSN = mCacheC(idx): pIBSSN = mCacheI(idx)
        SN_PercentuaisItem = True
        Exit Function
    End If
    RsOpen rs, "SELECT pIBSSN, pCBSSN FROM SimplesNacionalApuracaoAnexo WHERE Competencia = '" & mComp & "' AND Anexo = '" & UCase$(Trim$(sAnexo)) & "'"
    If Not rs.EOF Then
        If Not IsNull(rs!pCBSSN) Then bOvC = True: dOvC = CDbl(rs!pCBSSN)
        If Not IsNull(rs!pIBSSN) Then bOvI = True: dOvI = CDbl(rs!pIBSSN)
    End If
    rs.Close
    If Not (bOvC And bOvI) Then
        If Not mTemRBT12 Then
            mErro = "RBT12 do Simples Nacional nao cadastrado. Informe em Empresa > Enquadramento IBS/CBS > RBT12 e aliquotas."
            Exit Function
        End If
        nCRT = mCRT
        If mCRT = 3 Then nCRT = IIf(mTpRegime = 2, 1, 2)    'opcao pendente: segue o CRT da opcao
        If Not SN_PercentuaisDoAnexo(sAnexo, mAno, nCRT, mRBT12, pCBSSN, pIBSSN) Then
            mErro = "Nao ha aliquota do Anexo " & sAnexo & " para " & mAno & " (script 197)."
            Exit Function
        End If
    End If
    If bOvC Then pCBSSN = dOvC
    If bOvI Then pIBSSN = dOvI
    mCacheC(idx) = pCBSSN: mCacheI(idx) = pIBSSN: mCacheOk(idx) = True
    SN_PercentuaisItem = True
End Function

'Monta os valores do gTribSN do item. Falso = nao foi possivel (veja SN_UltimoErro).
Public Function SN_MontarItem(ByVal sAnexo As String, ByVal dBase As Double, _
                              ByRef pIBSSN As Double, ByRef vIBSSN As Double, _
                              ByRef pCBSSN As Double, ByRef vCBSSN As Double) As Boolean
    pIBSSN = 0: vIBSSN = 0: pCBSSN = 0: vCBSSN = 0
    If Not SN_PercentuaisItem(sAnexo, pCBSSN, pIBSSN) Then Exit Function
    If Not SN_InformaIBSNoTribSN(mCRT, mTpRegime) Then pIBSSN = 0
    vCBSSN = SN_ValorTributo(dBase, pCBSSN)
    vIBSSN = SN_ValorTributo(dBase, pIBSSN)
    SN_MontarItem = True
End Function
