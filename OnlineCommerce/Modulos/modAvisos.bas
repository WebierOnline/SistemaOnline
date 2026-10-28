Attribute VB_Name = "modAvisos"
Option Explicit
' Popup de avisos no login do Tela_Principal (Form frmAvisoPopup, config em
' Meus_Avisos.frm). Cada "aviso" tem uma chave fixa (grava/le em
' Usuario_Avisos.Aviso) e uma funcao geradora de texto abaixo. So aparece
' pro usuario se:
'  1) ele tiver a permissao 'EXIBIR POPUP' (Usuario_permissoes/Usuario_Acessos)
'  2) ele tiver marcado "Me avise" pra aquele aviso especifico (Usuario_Avisos)
'  3) o aviso tiver conteudo relevante no momento (ex: 0 contas vencendo = nao mostra)
' Fechar (X) so esconde na sessao atual - volta a aparecer no proximo login.

Public Const AVISO_CONTAS_PAGAR As String = "CONTAS_PAGAR"
Public Const AVISO_CONTAS_RECEBER As String = "CONTAS_RECEBER"
Public Const AVISO_ESTOQUE_BAIXO As String = "ESTOQUE_BAIXO"

Private Const COR_AVISO_ATENCAO As Long = &HC0FFFF    'amarelo claro
Private Const COR_AVISO_URGENTE As Long = &HC0C0FF    'vermelho claro

Private miPosPopup As Integer

Public Sub MostrarAvisosPopup(ByVal vCodUsuario As Long)
Dim sTexto As String

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
    If sTexto <> "" Then ExibirPopup "ESTOQUE BAIXO", sTexto, AVISO_ESTOQUE_BAIXO, COR_AVISO_URGENTE
End If
End Sub

Private Sub ExibirPopup(ByVal sTitulo As String, ByVal sTexto As String, ByVal sChave As String, ByVal lCor As Long)
Dim frm As New frmAvisoPopup
Const LARGURA_POPUP As Long = 3300
Const MARGEM As Long = 150
Const TOPO_AREA As Long = 1850

frm.Carregar sTitulo, sTexto, sChave, lCor
frm.Width = LARGURA_POPUP
frm.Left = MARGEM + (miPosPopup * (LARGURA_POPUP + MARGEM))
frm.Top = TOPO_AREA
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
