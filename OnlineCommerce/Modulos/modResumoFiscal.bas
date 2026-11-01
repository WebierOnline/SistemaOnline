Attribute VB_Name = "modResumoFiscal"
Option Explicit
'Resumo Fiscal: calculo, geracao do PDF e envio por e-mail para a contabilidade.
'Usado pela tela Notas_Adesivas (botoes) e pelo envio automatico mensal (tarefa agendada
'"OnlineCommerce_ResumoFiscal*" chama o sistema com o parametro /resumofiscal -> Senha.Form_Load
'-> ExecutarResumoFiscalAuto).

'Acrescimo que os contadores pedem entre o valor de compra (entradas) e o valor emitido em NFCe/NFe (META 1)
Public Const ACRESCIMO_META As Currency = 0.25

'Dia do mes a partir do qual o resumo do MES ANTERIOR pode ser enviado automaticamente
Private Const DIA_ENVIO As Integer = 8

Public Type TResumoFiscal
   sPeriodo As String          '"Outubro de 2026" ou "2026"
   vEntradas As Currency
   vPix As Currency
   vDebito As Currency
   vCredito As Currency
   vTransf As Currency
   vDepos As Currency
   vTotalRec As Currency
   vNFe As Currency
   vNFCe As Currency
   vTotalSaidas As Currency
   vMeta1 As Currency
   vMeta3 As Currency
   vRestante1 As Currency
   vRestante2 As Currency
End Type

Private colTexto As Collection
Private colValor As Collection
Private colNegrito As Collection
Private colTamanho As Collection
Private colCentro As Collection

Public RF_UltimoErro As String

'---------------------------------------------------------------------------
' Calculo
'---------------------------------------------------------------------------
'Condicao SQL do periodo (mes/ano ou so o ano) para o campo de data informado
Private Function CondPeriodo(ByVal sCampo As String, ByVal bMensal As Boolean, ByVal nMes As Integer, ByVal nAno As Integer) As String
If bMensal Then
   CondPeriodo = "(MONTH(" & sCampo & ") = " & nMes & ") AND (YEAR(" & sCampo & ") = " & nAno & ")"
Else
   CondPeriodo = "(YEAR(" & sCampo & ") = " & nAno & ")"
End If
End Function

'Executa um SELECT SUM(...) de uma coluna e devolve o total (NULL/sem linha = 0)
Private Function SomaSQL(ByVal sSQL As String) As Currency
Dim r As ADODB.Recordset

Set r = dbData.OpenRecordset(sSQL)
If Not r.EOF Then
   If Not IsNull(r(0).Value) Then SomaSQL = CCur(r(0).Value)
End If
If r.State <> 0 Then r.Close
Set r = Nothing
End Function

'Total recebido numa forma de pagamento: parcelas pagas + haver (sTipoCartao = "C"/"D" so para CARTAO)
Private Function TotalRecebido(ByVal sForma As String, ByVal sTipoCartao As String, ByVal bMensal As Boolean, ByVal nMes As Integer, ByVal nAno As Integer) As Currency
Dim sExtra As String

If sTipoCartao <> "" Then sExtra = " AND TIPO_CARTAO = '" & sTipoCartao & "'"

TotalRecebido = SomaSQL("SELECT SUM(VALOR_FINAL) FROM parcelas WHERE " & CondPeriodo("PAGAMENTO", bMensal, nMes, nAno) & _
                        " AND STATUS = 1 AND FORMA_PGTO = '" & sForma & "'" & sExtra) + _
                SomaSQL("SELECT SUM(VALOR_HAVER) FROM parcelas_haver WHERE " & CondPeriodo("HAVER", bMensal, nMes, nAno) & _
                        " AND FORMA_PGTO = '" & sForma & "'" & sExtra)
End Function

Private Function Maior(ByVal a As Currency, ByVal b As Currency) As Currency
If a > b Then Maior = a Else Maior = b
End Function

'Calcula todo o resumo do periodo (bMensal = mes nMes/nAno; senao o ano nAno inteiro)
Public Function RF_Calcular(ByVal bMensal As Boolean, ByVal nMes As Integer, ByVal nAno As Integer) As TResumoFiscal
Dim r As TResumoFiscal

If bMensal Then
   r.sPeriodo = StrConv(Format(DateSerial(nAno, nMes, 1), "mmmm"), vbProperCase) & " de " & nAno
Else
   r.sPeriodo = CStr(nAno)
End If

'RECEBIMENTOS (META 2)
r.vPix = TotalRecebido("PIX", "", bMensal, nMes, nAno)
r.vCredito = TotalRecebido("CARTAO", "C", bMensal, nMes, nAno)
r.vDebito = TotalRecebido("CARTAO", "D", bMensal, nMes, nAno)
r.vTransf = TotalRecebido("TRANSFERENCIA", "", bMensal, nMes, nAno)
r.vDepos = TotalRecebido("DEPOSITO", "", bMensal, nMes, nAno)
r.vTotalRec = r.vPix + r.vCredito + r.vDebito + r.vTransf + r.vDepos

'ENTRADAS (valor total das notas de compra importadas por XML)
r.vEntradas = SomaSQL("SELECT SUM(ValorNota) FROM EntradaEstoque WHERE " & CondPeriodo("DataEmissao", bMensal, nMes, nAno))

'SAIDAS FISCAIS: NFCe + NFe (NFe so de saida, sem devolucao/ajuste; entrada propria fica de fora)
r.vNFCe = SomaSQL("SELECT SUM(Valor_NF_Prod - DescontoPromocional) FROM TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) " & _
                  "AND (Inutilizada = 0) AND " & CondPeriodo("DataEmissao", bMensal, nMes, nAno))
r.vNFe = SomaSQL("SELECT SUM(ValorNota) FROM NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) " & _
                 "AND (ISNULL(CONVERT(varchar(1), TipoDocumento), '1') = '1') " & _
                 "AND (ISNULL(LEFT(FinalidadeEmissaoNFe, 1), '') NOT IN ('3', '4')) AND " & CondPeriodo("DataEmissao", bMensal, nMes, nAno))
r.vTotalSaidas = r.vNFCe + r.vNFe

'METAS: 1 = entradas + 25%; 2 = recebimentos; 3 = a maior entre a 1 e a 2
r.vMeta1 = r.vEntradas + (r.vEntradas * ACRESCIMO_META)
r.vMeta3 = Maior(r.vTotalRec, r.vMeta1)

'quanto falta emitir para alcancar a meta 1 e a meta 2 (nunca negativo)
r.vRestante1 = Maior(r.vMeta1 - r.vTotalSaidas, 0)
r.vRestante2 = Maior(r.vTotalRec - r.vTotalSaidas, 0)

RF_Calcular = r
End Function

'Frase de situacao; dPct recebe o % do minimo ja emitido (0 a mais de 1)
Public Function RF_Situacao(ByRef r As TResumoFiscal, ByRef dPct As Double) As String
If r.vMeta1 <= 0 Then
   dPct = 0
   RF_Situacao = "Sem compras no período para calcular."
Else
   dPct = CDbl(r.vTotalSaidas) / CDbl(r.vMeta1)
   If dPct >= 1 Then
      RF_Situacao = "Meta atingida! Mínimo já emitido."
   ElseIf dPct >= 0.6 Then
      RF_Situacao = "Quase lá! Você emitiu " & Format(dPct * 100, "0") & "% do mínimo."
   Else
      RF_Situacao = "Atenção: só " & Format(dPct * 100, "0") & "% do mínimo emitido."
   End If
End If
End Function

'---------------------------------------------------------------------------
' Linhas do relatorio (usadas pelo PDF e pelo relatorio ReportX)
'---------------------------------------------------------------------------
Public Sub RF_LinhasLimpar()
Set colTexto = New Collection
Set colValor = New Collection
Set colNegrito = New Collection
Set colTamanho = New Collection
Set colCentro = New Collection
End Sub

Public Sub RF_LinhaAdd(ByVal sTexto As String, ByVal sValor As String, Optional ByVal bNegrito As Boolean = False, Optional ByVal dTamanho As Double = 9, Optional ByVal bCentro As Boolean = False)
If colTexto Is Nothing Then RF_LinhasLimpar
If sTexto = "" Then sTexto = " "
colTexto.Add sTexto
colValor.Add sValor
colNegrito.Add bNegrito
colTamanho.Add dTamanho
colCentro.Add bCentro
End Sub

Public Function RF_LinhasContar() As Long
If colTexto Is Nothing Then RF_LinhasContar = 0 Else RF_LinhasContar = colTexto.Count
End Function

Public Function RF_LinhaTexto(ByVal i As Long) As String
RF_LinhaTexto = colTexto(i)
End Function

Public Function RF_LinhaValor(ByVal i As Long) As String
RF_LinhaValor = colValor(i)
End Function

Public Function RF_LinhaNegrito(ByVal i As Long) As Boolean
RF_LinhaNegrito = colNegrito(i)
End Function

Public Function RF_LinhaTamanho(ByVal i As Long) As Double
RF_LinhaTamanho = colTamanho(i)
End Function

Public Function RF_LinhaCentro(ByVal i As Long) As Boolean
RF_LinhaCentro = colCentro(i)
End Function

Private Sub LinhaRel(ByVal sTexto As String, ByVal sValor As String, Optional ByVal bNegrito As Boolean = False)
RF_LinhaAdd sTexto, sValor, bNegrito, 9
End Sub

Private Sub TituloRel(ByVal sTexto As String)
RF_LinhaAdd " ", "", False, 6
RF_LinhaAdd sTexto, "", True, 10
End Sub

'Monta as linhas do relatorio com os valores do resumo
Public Sub RF_MontarLinhas(ByRef r As TResumoFiscal)
Dim sRazao As String, sCnpj As String
Dim sPct As String
Dim dPct As Double

sRazao = Trim(CStr(SQLExecutaRetorno("SELECT TOP 1 RAZAO FROM empresa", "RAZAO", "")))
sCnpj = Trim(CStr(SQLExecutaRetorno("SELECT TOP 1 CNPJ FROM empresa", "CNPJ", "")))
sPct = Format(ACRESCIMO_META * 100, "0") & "%"

RF_LinhasLimpar
RF_LinhaAdd " ", "", False, 9
RF_LinhaAdd " ", "", False, 9
RF_LinhaAdd "RESUMO FISCAL - " & UCase(r.sPeriodo), "", True, 12, True
RF_LinhaAdd "Emitido em " & Format(Now, "dd/mm/yyyy hh:nn"), "", False, 8.25, True
RF_LinhaAdd " ", "", False, 6
RF_LinhaAdd "RAZÃO: " & sRazao, "", True, 12
If sCnpj <> "" Then RF_LinhaAdd "CNPJ: " & sCnpj, "", False, 8.25

TituloRel "FALTA EMITIR EM NOTA FISCAL"
LinhaRel RF_Situacao(r, dPct), "R$ " & Format(r.vRestante1, ocMONEY), True

TituloRel "O QUE COMPREI"
LinhaRel "Total das notas de compra (XML)", Format(r.vEntradas, ocMONEY), True

TituloRel "O QUE RECEBI DOS CLIENTES"
LinhaRel "Pix", Format(r.vPix, ocMONEY)
LinhaRel "Cartão de débito", Format(r.vDebito, ocMONEY)
LinhaRel "Cartão de crédito", Format(r.vCredito, ocMONEY)
LinhaRel "Transferência", Format(r.vTransf, ocMONEY)
LinhaRel "Depósito", Format(r.vDepos, ocMONEY)
LinhaRel "Total recebido", Format(r.vTotalRec, ocMONEY), True

TituloRel "O QUE EMITI DE NOTA (vendas)"
LinhaRel "Notas NF-e", Format(r.vNFe, ocMONEY)
LinhaRel "Notas NFC-e", Format(r.vNFCe, ocMONEY)
LinhaRel "Total emitido", Format(r.vTotalSaidas, ocMONEY), True

TituloRel "METAS DO PERÍODO"
LinhaRel "Mínimo recomendado (compras + " & sPct & ")", Format(r.vMeta1, ocMONEY)
LinhaRel "Meta do mês (a maior entre o mínimo e o recebido)", Format(r.vMeta3, ocMONEY)
LinhaRel "Falta emitir para o mínimo", Format(r.vRestante1, ocMONEY), True
LinhaRel "Falta emitir para igualar o que recebi", Format(r.vRestante2, ocMONEY), True

RF_LinhaAdd " ", "", False, 6
RF_LinhaAdd "NF-e: somente notas de saída enviadas (sem canceladas, devoluções, ajustes e entradas próprias).", "", False, 7
RF_LinhaAdd "NFC-e: somente notas enviadas (sem canceladas e inutilizadas).", "", False, 7
End Sub

'---------------------------------------------------------------------------
' PDF
'---------------------------------------------------------------------------
'Pasta RESUMOS dentro da pasta da logomarca da empresa (empresa.CAMINHO); se nao houver, na pasta do sistema
Public Function RF_PastaResumos() As String
Dim sCam As String
Dim sBase As String
Dim sPasta As String

sCam = Trim(CStr(SQLExecutaRetorno("SELECT TOP 1 ISNULL(caminho,'') AS c FROM empresa", "c", "")))
If InStrRev(sCam, "\") > 0 Then sBase = Left(sCam, InStrRev(sCam, "\"))
If sBase = "" Then
   sBase = appPathApp
ElseIf Dir$(sBase, vbDirectory) = "" Then
   sBase = appPathApp
End If

sPasta = sBase & "RESUMOS\"
If Dir$(sPasta, vbDirectory) = "" Then MkDir sPasta
RF_PastaResumos = sPasta
End Function

'Nome do arquivo: ResumoFiscal_Out2026_CNPJxxxxxxxxxxxxxx.pdf (ou Ano2026 quando e o ano todo)
Public Function RF_NomeArquivo(ByVal bMensal As Boolean, ByVal nMes As Integer, ByVal nAno As Integer) As String
Dim sCnpj As String, sSoNum As String, sRef As String
Dim i As Long

sCnpj = Trim(CStr(SQLExecutaRetorno("SELECT TOP 1 ISNULL(CNPJ,'') AS c FROM empresa", "c", "")))
For i = 1 To Len(sCnpj)
   If Mid(sCnpj, i, 1) >= "0" And Mid(sCnpj, i, 1) <= "9" Then sSoNum = sSoNum & Mid(sCnpj, i, 1)
Next

If bMensal Then
   sRef = StrConv(Format(DateSerial(nAno, nMes, 1), "mmm"), vbProperCase) & nAno
Else
   sRef = "Ano" & nAno
End If

RF_NomeArquivo = "ResumoFiscal_" & sRef & "_CNPJ" & sSoNum & ".pdf"
End Function

'Gera o PDF do resumo na pasta RESUMOS e devolve o caminho; "" se falhar (motivo em RF_UltimoErro)
Public Function RF_GerarPDF(ByRef r As TResumoFiscal, ByVal bMensal As Boolean, ByVal nMes As Integer, ByVal nAno As Integer) As String
Dim sArq As String

On Error GoTo deuErro

RF_UltimoErro = ""
sArq = RF_PastaResumos & RF_NomeArquivo(bMensal, nMes, nAno)
RF_MontarLinhas r
If EscreverPDF(sArq) Then RF_GerarPDF = sArq
Exit Function

deuErro:
RF_UltimoErro = "Não foi possível gerar o PDF: " & Err.Description
Err.Clear
End Function

Private Function PdfNum(ByVal d As Double) As String
PdfNum = Replace(Format(d, "0.00"), ",", ".")
End Function

Private Function PdfTexto(ByVal s As String) As String
s = Replace(s, "\", "\\")
s = Replace(s, "(", "\(")
s = Replace(s, ")", "\)")
PdfTexto = s
End Function

'Largura aproximada (pontos) do texto em Helvetica / Helvetica-Bold
Private Function PdfLargura(ByVal s As String, ByVal bNegrito As Boolean, ByVal dTam As Double) As Double
Dim aReg As Variant, aNeg As Variant
Dim i As Long, c As Long, w As Long, nTotal As Long

aReg = Array(278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, _
             556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, _
             1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, _
             667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, _
             333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, _
             556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584)
aNeg = Array(278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, _
             556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, _
             975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, _
             667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, _
             333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, _
             611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584)

For i = 1 To Len(s)
   c = Asc(Mid(s, i, 1))
   If c >= 32 And c <= 126 Then
      If bNegrito Then w = aNeg(c - 32) Else w = aReg(c - 32)
   ElseIf c >= 224 Then
      w = 556                       'letras minusculas acentuadas
   ElseIf c >= 192 Then
      w = 722                       'letras maiusculas acentuadas
   Else
      w = 556
   End If
   nTotal = nTotal + w
Next
PdfLargura = nTotal * dTam / 1000
End Function

'Escreve o PDF (A4, texto simples) com as linhas montadas por RF_MontarLinhas
Private Function EscreverPDF(ByVal sArq As String) As Boolean
Const PG_LARG As Double = 595
Const PG_ALT As Double = 842
Const MARGEM As Double = 42

Dim nLin As Long, i As Long, nPag As Long, k As Long
Dim aPag() As String
Dim y As Double, dTam As Double, dLead As Double, x As Double
Dim sTexto As String, sValor As String
Dim bNeg As Boolean
Dim sFonte As String
Dim sOut As String
Dim aOff() As Long
Dim nObj As Long
Dim nXref As Long
Dim sStream As String
Dim sKids As String
Dim aBytes() As Byte
Dim nArq As Integer

nLin = RF_LinhasContar
nPag = 1
ReDim aPag(1 To 1)
y = PG_ALT - MARGEM

For i = 1 To nLin
   sTexto = RF_LinhaTexto(i)
   sValor = RF_LinhaValor(i)
   bNeg = RF_LinhaNegrito(i)
   dTam = RF_LinhaTamanho(i)
   dLead = dTam * 1.7
   If bNeg Then sFonte = "/F2" Else sFonte = "/F1"

   If y - dLead < MARGEM Then
      nPag = nPag + 1
      ReDim Preserve aPag(1 To nPag)
      y = PG_ALT - MARGEM
   End If
   y = y - dLead

   If Trim(sTexto) <> "" Then
      If RF_LinhaCentro(i) Then
         x = (PG_LARG - PdfLargura(sTexto, bNeg, dTam)) / 2
      Else
         x = MARGEM
      End If
      aPag(nPag) = aPag(nPag) & "BT " & sFonte & " " & PdfNum(dTam) & " Tf " & PdfNum(x) & " " & PdfNum(y) & " Td (" & PdfTexto(sTexto) & ") Tj ET" & vbLf
   End If
   If sValor <> "" Then
      x = PG_LARG - MARGEM - PdfLargura(sValor, True, dTam)
      aPag(nPag) = aPag(nPag) & "BT /F2 " & PdfNum(dTam) & " Tf " & PdfNum(x) & " " & PdfNum(y) & " Td (" & PdfTexto(sValor) & ") Tj ET" & vbLf
   End If
Next

'objetos: 1 catalogo, 2 paginas, 3 e 4 fontes, depois (pagina, conteudo) de cada pagina
nObj = 4 + 2 * nPag
ReDim aOff(1 To nObj)
sOut = "%PDF-1.4" & vbLf

For k = 1 To nPag
   sKids = sKids & (3 + 2 * k) & " 0 R "
Next

aOff(1) = Len(sOut)
sOut = sOut & "1 0 obj" & vbLf & "<< /Type /Catalog /Pages 2 0 R >>" & vbLf & "endobj" & vbLf
aOff(2) = Len(sOut)
sOut = sOut & "2 0 obj" & vbLf & "<< /Type /Pages /Kids [" & sKids & "] /Count " & nPag & " >>" & vbLf & "endobj" & vbLf
aOff(3) = Len(sOut)
sOut = sOut & "3 0 obj" & vbLf & "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>" & vbLf & "endobj" & vbLf
aOff(4) = Len(sOut)
sOut = sOut & "4 0 obj" & vbLf & "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>" & vbLf & "endobj" & vbLf

For k = 1 To nPag
   aOff(3 + 2 * k) = Len(sOut)
   sOut = sOut & (3 + 2 * k) & " 0 obj" & vbLf & "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] " & _
          "/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents " & (4 + 2 * k) & " 0 R >>" & vbLf & "endobj" & vbLf
   sStream = aPag(k)
   aOff(4 + 2 * k) = Len(sOut)
   sOut = sOut & (4 + 2 * k) & " 0 obj" & vbLf & "<< /Length " & Len(sStream) & " >>" & vbLf & "stream" & vbLf & _
          sStream & "endstream" & vbLf & "endobj" & vbLf
Next

nXref = Len(sOut)
sOut = sOut & "xref" & vbLf & "0 " & (nObj + 1) & vbLf & "0000000000 65535 f " & vbLf
For k = 1 To nObj
   sOut = sOut & Format(aOff(k), "0000000000") & " 00000 n " & vbLf
Next
sOut = sOut & "trailer" & vbLf & "<< /Size " & (nObj + 1) & " /Root 1 0 R >>" & vbLf & "startxref" & vbLf & nXref & vbLf & "%%EOF" & vbLf

'grava em ANSI (1 byte por caractere = os offsets acima batem)
aBytes = StrConv(sOut, vbFromUnicode)
If Dir$(sArq) <> "" Then Kill sArq
nArq = FreeFile
Open sArq For Binary Access Write As #nArq
Put #nArq, , aBytes
Close #nArq

EscreverPDF = True
End Function

'---------------------------------------------------------------------------
' E-mail
'---------------------------------------------------------------------------
'E-mail da contabilidade (cadastro do contador); "" se nao houver
Public Function RF_EmailContador() As String
RF_EmailContador = Trim(CStr(SQLExecutaRetorno("SELECT TOP 1 email FROM TbContabilista WHERE (email IS NOT NULL) AND (email <> '')", "email", "")))
End Function

'Envio por e-mail pela DLL da NFe (mesmo caminho do cmdEnviarXML da NFe_Completa / Exportar_XML2).
'bAguardar = True no envio automatico: a entrega e assincrona dentro da DLL e o processo nao pode
'terminar logo apos o EmailEnviar (ver ExportarXMLServidor / EnviarEmailContador.ps1).
Public Function RF_EnviarEmail(ByVal sEmailPara As String, ByVal sAnexo As String, ByVal sPeriodo As String, ByVal bAguardar As Boolean) As Boolean
Dim sistNFe As snfe.Util
Dim pathAnexo() As String
Dim emailCC() As String
Dim corpoEmail As String
Dim vRet As Variant

On Error GoTo deuErro

RF_UltimoErro = ""
Set sistNFe = New snfe.Util
Call ConfiguraDLLNFeNFCe(55, "1", sistNFe)

ReDim emailCC(0)
emailCC(0) = sEmailPara
ReDim pathAnexo(0)
pathAnexo(0) = sAnexo

corpoEmail = "Segue em anexo o Resumo Fiscal referente a " & sPeriodo & "." & _
             "<br><br>" & _
             "Atenciosamente, " & _
             "<br><br>" & _
             Trim(CStr(SQLExecutaRetorno("SELECT TOP 1 RAZAO FROM empresa", "RAZAO", "")))

vRet = sistNFe.EmailEnviar(sEmailPara, "Resumo Fiscal - " & sPeriodo, corpoEmail, pathAnexo, emailCC)
If vRet Then
   If bAguardar Then Sleep 25000
   RF_EnviarEmail = True
Else
   RF_UltimoErro = "A DLL de e-mail retornou falha no envio."
End If
Set sistNFe = Nothing
Exit Function

deuErro:
RF_UltimoErro = Err.Description
Err.Clear
Set sistNFe = Nothing
End Function

'---------------------------------------------------------------------------
' Envio automatico mensal
'---------------------------------------------------------------------------
Private Sub RF_Log(ByVal sMsg As String)
Dim nLog As Integer

On Error Resume Next
nLog = FreeFile
Open appPathApp & "ResumoFiscalAuto_log.txt" For Append As #nLog
Print #nLog, Format(Now, "dd/mm/yyyy hh:nn:ss") & "  " & sMsg
Close #nLog
End Sub

'Chamada pelo Senha.Form_Load quando o sistema e iniciado com /resumofiscal (tarefa agendada, so no servidor).
'Regra: roda todo dia (e ao ligar/logar); a partir do dia DIA_ENVIO envia o resumo do MES ANTERIOR, uma unica vez
'(marcador RESUMOS\enviado_AAAAMM.txt). Nunca envia o mes corrente nem futuro.
Public Sub ExecutarResumoFiscalAuto()
Dim dMes As Date
Dim nMes As Integer, nAno As Integer
Dim sPasta As String, sMarc As String, sLock As String
Dim sEmail As String, sArq As String
Dim r As TResumoFiscal
Dim nArq As Integer
Dim bLock As Boolean

On Error GoTo deuErro

If Not EhServidorLocal() Then Exit Sub
If Day(Date) < DIA_ENVIO Then Exit Sub

dMes = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
nMes = Month(dMes)
nAno = Year(dMes)

sPasta = RF_PastaResumos
sMarc = sPasta & "enviado_" & Format(dMes, "yyyymm") & ".txt"
If Dir$(sMarc) <> "" Then Exit Sub        'ja enviado

'trava contra duas execucoes ao mesmo tempo (tarefa diaria + ao logar)
sLock = sPasta & "enviando.lock"
If Dir$(sLock) <> "" Then
   If DateDiff("n", FileDateTime(sLock), Now) < 30 Then Exit Sub
End If
nArq = FreeFile
Open sLock For Output As #nArq
Print #nArq, Format(Now, "dd/mm/yyyy hh:nn:ss")
Close #nArq
bLock = True

sEmail = RF_EmailContador
If sEmail = "" Then
   RF_Log "Resumo Fiscal " & Format(dMes, "mm/yyyy") & " NAO enviado: e-mail do contador nao cadastrado (TbContabilista)."
   GoTo fim
End If
If InStr(1, sEmail, "@") = 0 Then
   RF_Log "Resumo Fiscal " & Format(dMes, "mm/yyyy") & " NAO enviado: e-mail do contador invalido (" & sEmail & ")."
   GoTo fim
End If

r = RF_Calcular(True, nMes, nAno)
sArq = RF_GerarPDF(r, True, nMes, nAno)
If sArq = "" Then
   RF_Log "Resumo Fiscal " & Format(dMes, "mm/yyyy") & " NAO enviado: " & RF_UltimoErro
   GoTo fim
End If

If RF_EnviarEmail(sEmail, sArq, r.sPeriodo, True) Then
   nArq = FreeFile
   Open sMarc For Output As #nArq
   Print #nArq, "Enviado em " & Format(Now, "dd/mm/yyyy hh:nn:ss") & " para " & sEmail & " (" & sArq & ")"
   Close #nArq
   RF_Log "Resumo Fiscal " & Format(dMes, "mm/yyyy") & " enviado para " & sEmail & "."
Else
   RF_Log "Resumo Fiscal " & Format(dMes, "mm/yyyy") & " NAO enviado: " & RF_UltimoErro
End If

fim:
If bLock Then
   On Error Resume Next
   Kill sLock
End If
Exit Sub

deuErro:
RF_Log "Erro no envio automatico do Resumo Fiscal: " & Err.Description
Err.Clear
Resume fim
End Sub
