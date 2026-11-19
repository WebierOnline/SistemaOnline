Attribute VB_Name = "modZapEnvio"
Option Explicit
'---------------------------------------------------------------------------
' Envio de documentos pelo WhatsApp da empresa (conectado no Online Commerce, Sistema > WhatsApp).
' Envio em si: ZapOnlineEnviar (LicencaOnline.bas). Usado pela OS (modSaidaPDF, frmNFSe) e pela NF-e (NFe_Completa).
'---------------------------------------------------------------------------

'Celular do cliente no cadastro; sem celular, o telefone1 quando tem cara de celular (11 dígitos). "" se não tiver.
Public Function PDFZap_CelularCliente(ByVal lCodCliente As Long) As String
Dim s As String

If lCodCliente <= 0 Then Exit Function
s = Trim$(CStr(SQLExecutaRetorno("SELECT celular FROM cliente WHERE codigo = " & lCodCliente, "celular", "")))
If Len(PDFZap_Digitos(s)) < 10 Then
    s = Trim$(CStr(SQLExecutaRetorno("SELECT telefone1 FROM cliente WHERE codigo = " & lCodCliente, "telefone1", "")))
    If Len(PDFZap_Digitos(s)) <> 11 Then s = ""
End If
PDFZap_CelularCliente = s
End Function

'Pergunta o celular (já preenchido com o do cadastro). "" = cancelou.
Public Function PDFZap_PedirCelular(ByVal sSugestao As String) As String
Dim s As String, d As String

Do
    s = Trim$(InputBox("Informe o celular (com DDD) para enviar o PDF pelo WhatsApp:", "Enviar por WhatsApp", sSugestao))
    If s = "" Then Exit Function
    d = PDFZap_Digitos(s)
    If Len(d) >= 12 And Left$(d, 2) = "55" Then d = Mid$(d, 3)
    If Len(d) = 10 Or Len(d) = 11 Then
        PDFZap_PedirCelular = s
        Exit Function
    End If
    MsgBox "Celular inválido. Digite com o DDD, ex.: (89) 9 9999-9999.", vbExclamation, "Enviar por WhatsApp"
    sSugestao = s
Loop
End Function

'Mesmo número (só os dígitos, sem o 55)?
Public Function PDFZap_MesmoNumero(ByVal a As String, ByVal b As String) As Boolean
a = PDFZap_Digitos(a): b = PDFZap_Digitos(b)
If Len(a) >= 12 And Left$(a, 2) = "55" Then a = Mid$(a, 3)
If Len(b) >= 12 And Left$(b, 2) = "55" Then b = Mid$(b, 3)
PDFZap_MesmoNumero = (a = b)
End Function

'Grava o celular no cadastro do cliente no formato da tela, "(89) 99999-9999". True se atualizou.
Public Function PDFZap_AtualizarCliente(ByVal lCodCliente As Long, ByVal sCelular As String) As Boolean
Dim d As String, s As String

If lCodCliente <= 0 Then Exit Function
d = PDFZap_Digitos(sCelular)
If Len(d) >= 12 And Left$(d, 2) = "55" Then d = Mid$(d, 3)
If Len(d) = 11 Then
    s = "(" & Left$(d, 2) & ") " & Mid$(d, 3, 5) & "-" & Mid$(d, 8)
ElseIf Len(d) = 10 Then
    s = "(" & Left$(d, 2) & ") " & Mid$(d, 3, 4) & "-" & Mid$(d, 7)
Else
    Exit Function
End If
On Error GoTo deuErro
dbData.Execute "UPDATE cliente SET celular = '" & s & "' WHERE codigo = " & lCodCliente
PDFZap_AtualizarCliente = True
Exit Function
deuErro:
Err.Clear
End Function

'Primeiro nome do cliente, só com a inicial maiúscula ("JOAO DA SILVA" -> "Joao"). "" se não tiver.
Public Function PDFZap_PrimeiroNome(ByVal lCodCliente As Long) As String
Dim s As String

If lCodCliente <= 0 Then Exit Function
s = Trim$(CStr(SQLExecutaRetorno("SELECT nome FROM cliente WHERE codigo = " & lCodCliente, "nome", "")))
If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
PDFZap_PrimeiroNome = StrConv(s, vbProperCase)
End Function

'Nome da empresa para assinar a mensagem (fantasia; sem fantasia, a razão social)
Public Function PDFZap_NomeEmpresa() As String
Dim s As String

s = Trim$(CStr(SQLExecutaRetorno("SELECT TOP 1 FANTASIA FROM empresa", "FANTASIA", "")))
If s = "" Then s = Trim$(CStr(SQLExecutaRetorno("SELECT TOP 1 RAZAO FROM empresa", "RAZAO", "")))
PDFZap_NomeEmpresa = s
End Function

Public Function PDFZap_Digitos(ByVal s As String) As String
Dim i As Long, ch As String

For i = 1 To Len(s)
    ch = Mid$(s, i, 1)
    If ch Like "#" Then PDFZap_Digitos = PDFZap_Digitos & ch
Next
End Function

'Emoji (code point acima de FFFF) como par UTF-16 para strings do VB6.
Public Function PDFZap_Emoji(ByVal lCodigo As Long) As String
Dim v As Long

v = lCodigo - &H10000
PDFZap_Emoji = ChrW$(&HD800& + (v \ &H400&)) & ChrW$(&HDC00& + (v And &H3FF&))
End Function


'Celular do fornecedor no cadastro (Celular; sem ele, o Telefone quando tem cara de celular). "" se não tiver.
Public Function PDFZap_CelularFornecedor(ByVal lCodigo As Long) As String
Dim s As String

If lCodigo <= 0 Then Exit Function
s = Trim$(CStr(SQLExecutaRetorno("SELECT Celular FROM fornecedor WHERE codigo = " & lCodigo, "Celular", "")))
If Len(PDFZap_Digitos(s)) < 10 Then
    s = Trim$(CStr(SQLExecutaRetorno("SELECT Telefone FROM fornecedor WHERE codigo = " & lCodigo, "Telefone", "")))
    If Len(PDFZap_Digitos(s)) <> 11 Then s = ""
End If
PDFZap_CelularFornecedor = s
End Function

'Mensagem do WhatsApp da NF-e (tela NFe_Completa), no estilo dos modelos da OS.
'nTipo: 1 = NF-e autorizada (DANFE + XML), 2 = espelho (só PDF, para conferência), 3 = Carta de Correção (PDF + XML).
'sSeqCCe = sequência da carta (só no tipo 3).
Public Function PDFZap_MensagemNFe(ByVal nTipo As Integer, ByVal sNome As String, ByVal sNumero As String, _
                                   ByVal sValor As String, ByVal bCancelada As Boolean, ByVal sEmpresa As String, _
                                   Optional ByVal sSeqCCe As String = "") As String
Dim L As String, sOla As String, sNota As String, sValorTxt As String, sCorpo As String

L = vbLf
sOla = "Olá" & IIf(sNome <> "", ", *" & sNome & "*", "") & "! Tudo bem? " & PDFZap_Emoji(&H1F44B) & L & L
sNota = "*Nota Fiscal eletrônica (NF-e) nº " & sNumero & "*"
sValorTxt = IIf(sValor <> "", ", no valor de *R$ " & sValor & "*", "")

Select Case nTipo
    Case 2  'Espelho
        sCorpo = "Segue em anexo o *espelho* da " & sNota & sValorTxt & ", para sua conferência. " & PDFZap_Emoji(&H1F4C4) & L & L & _
                 "Este documento ainda *não tem valor fiscal*. Confira os dados e, se estiver tudo certo, é só nos avisar por aqui que emitimos a nota! " & _
                 PDFZap_Emoji(&H1F60A)
    Case 3  'Carta de Correção
        sCorpo = "Segue em anexo a *Carta de Correção eletrônica (CC-e)*" & IIf(sSeqCCe <> "", " nº " & sSeqCCe, "") & _
                 " referente à " & sNota & ", em PDF e o arquivo XML do evento. " & PDFZap_Emoji(&H1F4C4) & L & L & _
                 "Guarde junto com a nota para o seu controle."
    Case Else  'NF-e autorizada
        sCorpo = "Segue em anexo a " & sNota & sValorTxt & ", com o DANFE em PDF e o arquivo XML. " & PDFZap_Emoji(&H1F4C4) & L & L & _
                 IIf(bCancelada, ChrW$(&H26A0) & ChrW$(&HFE0F) & " *Atenção: esta NF-e foi CANCELADA.*" & L & L, "") & _
                 "A autenticidade pode ser conferida pela chave de acesso no portal da NF-e (www.nfe.fazenda.gov.br)."
End Select

PDFZap_MensagemNFe = sOla & sCorpo & IIf(sEmpresa <> "", L & L & "Atenciosamente," & L & "*" & sEmpresa & "*", "")
End Function
