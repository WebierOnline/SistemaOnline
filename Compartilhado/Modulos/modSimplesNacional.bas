Attribute VB_Name = "modSimplesNacional"
Option Explicit
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
