Attribute VB_Name = "modAvisos"
Option Explicit
' Popup de avisos no login do Tela_Principal (Form frmAvisoPopup, config em
' Meus_Avisos.frm). Cada "aviso" tem uma chave fixa (grava/le em
' Usuario_Avisos.Aviso) e uma funcao geradora de texto abaixo. So aparece
' pro usuario se:
'  1) ele tiver a permissao 'EXIBIR POPUP' (Usuario_permissoes/Usuario_Acessos)
'  2) ele tiver marcado "Me avise" pra aquele aviso especifico (Usuario_Avisos)
'  3) o aviso tiver conteudo relevante no momento (ex: 0 contas vencendo = nao mostra)
' Cada popup e um MDIChild da Tela_Principal (como o Resumo Fiscal): fica aberto ate o usuario
' clicar no X (ou em CLIQUE AQUI), mesmo abrindo outras telas. Fechar so esconde na sessao
' atual - volta a aparecer no proximo login.

Public Const AVISO_CONTAS_PAGAR As String = "CONTAS_PAGAR"
Public Const AVISO_CONTAS_RECEBER As String = "CONTAS_RECEBER"
Public Const AVISO_ESTOQUE_BAIXO As String = "ESTOQUE_BAIXO"
Public Const AVISO_ALERTAS_GERAIS As String = "ALERTAS_GERAIS"

Private Const COR_AVISO_ATENCAO As Long = &HC0FFFF    'amarelo claro
Private Const COR_AVISO_URGENTE As Long = &HC0C0FF    'vermelho claro

Private miPosPopup As Integer

Public Sub MostrarAvisosPopup(ByVal vCodUsuario As Long)
Dim sTexto As String
Dim sDestinos As String

If Not TemPermissaoPopup(vCodUsuario) Then Exit Sub

miPosPopup = 0

If AvisoAtivo(vCodUsuario, AVISO_CONTAS_PAGAR) Then
    sTexto = GerarAvisoContasPagar()
    If sTexto <> "" Then ExibirPopup "CONTAS A PAGAR", sTexto, AVISO_CONTAS_PAGAR, COR_AVISO_ATENCAO
End If

If AvisoAtivo(vCodUsuario, AVISO_CONTAS_RECEBER) Then
    sTexto = GerarAvisoContasReceber()
    If sTexto <> "" Then ExibirPopup "CONTAS A RECEBER", sTexto, AVISO_CONTAS_RECEBER, COR_AVISO_ATENCAO
End If

If AvisoAtivo(vCodUsuario, AVISO_ESTOQUE_BAIXO) Then
    sTexto = GerarAvisoEstoqueBaixo()
    If sTexto <> "" Then ExibirPopup "ESTOQUE BAIXO", sTexto, AVISO_ESTOQUE_BAIXO, COR_AVISO_ATENCAO
End If

If AvisoAtivo(vCodUsuario, AVISO_ALERTAS_GERAIS) Then
    sTexto = GerarAlertasGerais(sDestinos)
    If sTexto <> "" Then ExibirPopupLinhas "ALERTAS GERAIS", sTexto, sDestinos, COR_AVISO_URGENTE
End If
End Sub

Private Sub ExibirPopup(ByVal sTitulo As String, ByVal sTexto As String, ByVal sChave As String, ByVal lCor As Long)
Dim frm As New frmAvisoPopup
Const MARGEM As Long = 150

frm.Carregar sTitulo, sTexto, sChave, lCor
frm.Left = MARGEM + (miPosPopup * (frm.Width + MARGEM))   'posicao relativa a area da Tela_Principal (MDIChild)
frm.Top = MARGEM
frm.Show vbModeless
frm.ZOrder 0
miPosPopup = miPosPopup + 1
End Sub

Private Sub ExibirPopupLinhas(ByVal sTitulo As String, ByVal sTextos As String, ByVal sDestinos As String, ByVal lCor As Long)
Dim frm As New frmAvisoPopup
Const MARGEM As Long = 150

frm.CarregarLinhas sTitulo, sTextos, sDestinos, lCor
frm.Left = MARGEM + (miPosPopup * (frm.Width + MARGEM))
frm.Top = MARGEM
frm.Show vbModeless
frm.ZOrder 0
miPosPopup = miPosPopup + 1
End Sub

Private Function TemPermissaoPopup(ByVal vCodUsuario As Long) As Boolean
TemPermissaoPopup = (SQLExecutaRetorno( _
    "SELECT COUNT(*) r FROM Usuario_permissoes INNER JOIN Usuario_Acessos ON Usuario_permissoes.Codigo = Usuario_Acessos.Cod_Permissao " & _
    "WHERE (Usuario_permissoes.permissao = 'EXIBIR POPUP') AND (Usuario_Acessos.Cod_Usuario = " & vCodUsuario & ")", "r", 0) > 0)
End Function

Private Function AvisoAtivo(ByVal vCodUsuario As Long, ByVal sChave As String) As Boolean
AvisoAtivo = (SQLExecutaRetorno( _
    "SELECT COUNT(*) r FROM Usuario_Avisos WHERE Cod_Usuario = " & vCodUsuario & " AND Aviso = '" & sChave & "' AND Ativo = 1", "r", 0) > 0)
End Function

Public Function GerarAvisoContasPagar() As String
Dim nHoje As Long, nVencidas As Long, nProximos As Long
Dim sTexto As String

nHoje = SQLExecutaRetorno("SELECT COUNT(*) r FROM a_pagar WHERE status LIKE '%PAGAR%' AND CONVERT(date, vencimento) = CONVERT(date, GETDATE())", "r", 0)
nVencidas = SQLExecutaRetorno("SELECT COUNT(*) r FROM a_pagar WHERE status LIKE '%PAGAR%' AND CONVERT(date, vencimento) < CONVERT(date, GETDATE())", "r", 0)
nProximos = SQLExecutaRetorno("SELECT COUNT(*) r FROM a_pagar WHERE status LIKE '%PAGAR%' AND CONVERT(date, vencimento) BETWEEN CONVERT(date, DATEADD(day, 1, GETDATE())) AND CONVERT(date, DATEADD(day, 5, GETDATE()))", "r", 0)

If nHoje = 0 And nVencidas = 0 And nProximos = 0 Then
    GerarAvisoContasPagar = ""
    Exit Function
End If

If nHoje > 0 Then sTexto = sTexto & "Hoje tem " & nHoje & " conta(s) a pagar vencendo." & vbCrLf
If nVencidas > 0 Then sTexto = sTexto & "Existem " & nVencidas & " conta(s) vencida(s)!" & vbCrLf
If nProximos > 0 Then sTexto = sTexto & "Existem " & nProximos & " conta(s) vencendo nos próximos 5 dias." & vbCrLf

GerarAvisoContasPagar = sTexto
End Function

Public Function GerarAvisoContasReceber() As String
Dim nHoje As Long, nVencidas As Long, nProximos As Long
Dim sTexto As String

'Contas a Receber = qualquer parcela ainda aberta (status=0), sem importar
'a origem (venda/OS/aluguel/receber) - usuario confirmou que nao deve
'filtrar por tipo_pedido, so pelo fato de estar em aberto.
nHoje = SQLExecutaRetorno("SELECT COUNT(*) r FROM parcelas WHERE status = 0 AND CONVERT(date, data) = CONVERT(date, GETDATE())", "r", 0)
nVencidas = SQLExecutaRetorno("SELECT COUNT(*) r FROM parcelas WHERE status = 0 AND CONVERT(date, data) < CONVERT(date, GETDATE())", "r", 0)
nProximos = SQLExecutaRetorno("SELECT COUNT(*) r FROM parcelas WHERE status = 0 AND CONVERT(date, data) BETWEEN CONVERT(date, DATEADD(day, 1, GETDATE())) AND CONVERT(date, DATEADD(day, 5, GETDATE()))", "r", 0)

If nHoje = 0 And nVencidas = 0 And nProximos = 0 Then
    GerarAvisoContasReceber = ""
    Exit Function
End If

If nHoje > 0 Then sTexto = sTexto & "Hoje tem " & nHoje & " conta(s) a receber vencendo." & vbCrLf
If nVencidas > 0 Then sTexto = sTexto & "Existem " & nVencidas & " conta(s) vencida(s)!" & vbCrLf
If nProximos > 0 Then sTexto = sTexto & "Existem " & nProximos & " conta(s) vencendo nos próximos 5 dias." & vbCrLf

GerarAvisoContasReceber = sTexto
End Function

Public Function GerarAvisoEstoqueBaixo() As String
Dim nProdutos As Long

nProdutos = SQLExecutaRetorno("SELECT COUNT(*) r FROM produtos WHERE quant_estoque < quant_min AND quant_min > 0", "r", 0)

If nProdutos = 0 Then
    GerarAvisoEstoqueBaixo = ""
    Exit Function
End If

GerarAvisoEstoqueBaixo = "Existem " & nProdutos & " produto(s) com estoque abaixo do mínimo."
End Function

'Alertas gerais: cada pendencia vira uma linha (texto) e um destino (chave do form aberto pelo clique, ver
'lblExibir_Click em frmAvisoPopup). Textos separados por vbLf e destinos por "|". Um "~" dentro do texto
'quebra a linha dentro do mesmo alerta. Novos alertas entram aqui.
Public Function GerarAlertasGerais(ByRef sDestinos As String) As String
Dim sTextos As String
Dim rf As TResumoFiscal
Dim nDias As Long

sDestinos = ""

'1) contabilidade sem e-mail (necessario para o envio automatico do Resumo Fiscal)
If RF_EmailContador() = "" Then AddAlerta sTextos, sDestinos, "Contabilidade sem e-mail cadastrado.", "CONTADOR"

'2) quanto falta emitir em notas de saida para bater a meta fiscal do mes
rf = RF_Calcular(True, Month(Date), Year(Date))
If rf.vRestante1 > 0 Then
    AddAlerta sTextos, sDestinos, "Meta fiscal de " & StrConv(Format(Date, "mmm"), vbProperCase) & "/" & Year(Date) & " faltam:~" & Format(rf.vRestante1, ocMONEY) & " em notas de saída.", "RESUMO_FISCAL"
End If

'3) certificado digital vencendo (avisa a partir de 15 dias antes) ou vencido
'validade lida pela DLL (certificado instalado nesta maquina)
Dim dVencCert As Date
If Not CertificadoDiasRestantes(nDias, dVencCert) Then
    nDias = 99999
    If CertFalhaMotivo = 1 Then AddAlerta sTextos, sDestinos, "Certificado digital não encontrado nesta máquina!", "EMPRESA"
End If
If nDias <= 15 Then
    If nDias < 0 Then
        AddAlerta sTextos, sDestinos, "Certificado digital VENCIDO há " & Abs(nDias) & " dia(s)!", "EMPRESA"
    ElseIf nDias = 0 Then
        AddAlerta sTextos, sDestinos, "Certificado digital vence HOJE!", "EMPRESA"
    Else
        AddAlerta sTextos, sDestinos, "Certificado digital vence em " & nDias & " dia(s) (" & Format(Date + nDias, "dd/mm/yyyy") & ").", "EMPRESA"
    End If
End If

'4) monitor de NF-e (/monitornfe): notas importadas sozinhas aguardando conferência, falhas e notas canceladas
'pelo fornecedor depois de lançadas. Sem as colunas do script 177 as consultas falham e retornam 0.
Dim nNotas As Long
nNotas = Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) r FROM EntradaEstoque e WHERE e.OrigemLancamento = 'MONITOR' " & _
         "AND ISNULL(e.Cancelada, 0) = 0 AND EXISTS (SELECT 1 FROM EntradaEstoqueItens i WHERE i.CodigoNota = e.CodigoNota " & _
         "AND ISNULL(i.Adicionada, 0) = 0)", "r", 0)))
If nNotas > 0 Then AddAlerta sTextos, sDestinos, nNotas & " nota(s) de compra importada(s) automaticamente~aguardando conferência na Entrada de Estoque.", "ENTRADA_ESTOQUE"

nNotas = Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) r FROM TbNotasDestinadasManifestacao d WHERE d.ImportacaoStatus = 2 " & _
         "AND NOT EXISTS (SELECT 1 FROM EntradaEstoque e WHERE e.ChavedeAcesso = d.ChaveAcesso)", "r", 0)))
If nNotas > 0 Then AddAlerta sTextos, sDestinos, nNotas & " nota(s) não puderam ser importadas automaticamente~(veja em Notas Destinadas).", "NOTAS_DESTINADAS"

nNotas = Val(CStr(SQLExecutaRetorno("SELECT COUNT(*) r FROM TbNotasDestinadasManifestacao d WHERE d.TipoEventoCodigo = '110111' " & _
         "AND EXISTS (SELECT 1 FROM EntradaEstoque e WHERE e.ChavedeAcesso = d.ChaveAcesso AND ISNULL(e.Cancelada, 0) = 0)", "r", 0)))
If nNotas > 0 Then AddAlerta sTextos, sDestinos, nNotas & " nota(s) lançada(s) na Entrada de Estoque~foram CANCELADAS pelo fornecedor!", "NOTAS_DESTINADAS"

GerarAlertasGerais = sTextos
End Function

Private Sub AddAlerta(ByRef sTextos As String, ByRef sDestinos As String, ByVal sTexto As String, ByVal sDestino As String)
If sTextos <> "" Then
    sTextos = sTextos & vbLf
    sDestinos = sDestinos & "|"
End If
sTextos = sTextos & sTexto
sDestinos = sDestinos & sDestino
End Sub

'Posiciona uma janela filha da Tela_Principal (ex: Resumo Fiscal) na mesma altura dos popups: ao lado do ultimo
'popup aberto (o mais a direita) ou, se nao houver nenhum, na posicao do primeiro popup.
Public Sub PosicionarAoLadoDosPopups(frmDestino As Form)
Const MARGEM As Long = 150
Dim f As Form
Dim bAchou As Boolean
Dim nMaxLeft As Long
Dim nLeft As Long

nLeft = MARGEM
For Each f In Forms
    If TypeOf f Is frmAvisoPopup Then
        If Not bAchou Then
            bAchou = True
            nMaxLeft = f.Left
            nLeft = f.Left + f.Width + MARGEM
        ElseIf f.Left > nMaxLeft Then
            nMaxLeft = f.Left
            nLeft = f.Left + f.Width + MARGEM
        End If
    End If
Next

'nao deixa sair da area da Tela_Principal
If nLeft + frmDestino.Width > Tela_Principal.ScaleWidth Then nLeft = Tela_Principal.ScaleWidth - frmDestino.Width
If nLeft < 0 Then nLeft = 0

frmDestino.Move nLeft, MARGEM
End Sub
