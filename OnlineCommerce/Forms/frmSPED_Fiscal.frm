VERSION 5.00
Begin VB.Form frmSPED_Fiscal
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SPED Fiscal (EFD ICMS/IPI)"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTipo
      Caption         =   "Tipo do arquivo"
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8280
      Begin VB.OptionButton optCompleto
         Caption         =   "EFD completa (notas, apuração e inventário)"
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Top             =   300
         Width           =   4035
      End
      Begin VB.OptionButton optInventario
         Caption         =   "Somente o inventário (Bloco H)"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   3735
      End
   End
   Begin VB.Frame fraDados
      Caption         =   "Período, finalidade e inventário"
      Height          =   1455
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   8280
      Begin VB.ComboBox cboInventario
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   6855
      End
      Begin VB.ComboBox cboFinalidade
         Height          =   315
         Left            =   5340
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   300
         Width           =   2775
      End
      Begin VB.TextBox txtDataFim
         Height          =   315
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   7
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox txtDataIni
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   5
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblDicaInventario
         Caption         =   "O inventário de 31/12 vai na EFD de fevereiro do ano seguinte."
         Height          =   255
         Left            =   1260
         TabIndex        =   14
         Top             =   1080
         Width           =   6855
      End
      Begin VB.Label lblInventario
         AutoSize        =   -1  'True
         Caption         =   "Inventário:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   780
         Width           =   750
      End
      Begin VB.Label lblFinalidade
         AutoSize        =   -1  'True
         Caption         =   "Finalidade:"
         Height          =   195
         Left            =   4440
         TabIndex        =   8
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblDataFim
         AutoSize        =   -1  'True
         Caption         =   "até"
         Height          =   195
         Left            =   2550
         TabIndex        =   6
         Top             =   360
         Width           =   225
      End
      Begin VB.Label lblDataIni
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraBlocos
      Caption         =   "Blocos do arquivo (desmarcado = o bloco vai sem dados)"
      Height          =   3075
      Left            =   60
      TabIndex        =   15
      Top             =   2280
      Width           =   8280
      Begin VB.CheckBox chkBloco9
         Caption         =   "Bloco 9 - Controle e Encerramento do Arquivo Digital (obrigatório)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   2730
         Value           =   1  'Checked
         Width           =   7935
      End
      Begin VB.CheckBox chkBloco1
         Caption         =   "Bloco 1 - Outras Informações (obrigatório)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   2460
         Value           =   1  'Checked
         Width           =   7935
      End
      Begin VB.CheckBox chkBlocoK
         Caption         =   "Bloco K - Controle da Produção e do Estoque"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   2190
         Width           =   7935
      End
      Begin VB.CheckBox chkBlocoH
         Caption         =   "Bloco H - Inventário Físico"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   1920
         Width           =   7935
      End
      Begin VB.CheckBox chkBlocoG
         Caption         =   "Bloco G - Controle do Crédito de ICMS do Ativo Permanente - CIAP (o sistema não controla)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   1650
         Width           =   7935
      End
      Begin VB.CheckBox chkBlocoE
         Caption         =   "Bloco E - Apuração do ICMS e do IPI (obrigatório; zerado sem o Bloco C)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   1380
         Value           =   1  'Checked
         Width           =   7935
      End
      Begin VB.CheckBox chkBlocoD
         Caption         =   "Bloco D - Documentos Fiscais II - Serviços (ICMS): CT-e e notas de comunicação"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1110
         Width           =   7935
      End
      Begin VB.CheckBox chkBlocoC
         Caption         =   "Bloco C - Documentos Fiscais I - Mercadorias (ICMS/IPI): NF-e, NFC-e e notas de entrada"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   840
         Width           =   7935
      End
      Begin VB.CheckBox chkBlocoB
         Caption         =   "Bloco B - Escrituração e Apuração do ISS (só Distrito Federal)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   570
         Width           =   7935
      End
      Begin VB.CheckBox chkBloco0
         Caption         =   "Bloco 0 - Abertura, Identificação e Referências (obrigatório)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Value           =   1  'Checked
         Width           =   7935
      End
   End
   Begin VB.Frame fraProdutos
      Caption         =   "Produtos do inventário (conforme o cadastro do produto)"
      Height          =   675
      Left            =   60
      TabIndex        =   26
      Top             =   5400
      Width           =   8280
      Begin VB.CheckBox chkImobilizado
         Caption         =   "Incluir produtos do Ativo Imobilizado"
         Height          =   255
         Left            =   4080
         TabIndex        =   28
         Top             =   300
         Width           =   4035
      End
      Begin VB.CheckBox chkUsoConsumo
         Caption         =   "Incluir produtos de Uso e Consumo"
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   300
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdFechar
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   420
      Left            =   6840
      TabIndex        =   33
      Top             =   7140
      Width           =   1500
   End
   Begin VB.CommandButton cmdPasta
      Caption         =   "Abrir &pasta"
      Height          =   420
      Left            =   3960
      TabIndex        =   32
      Top             =   7140
      Width           =   1500
   End
   Begin VB.CommandButton cmdEnviar
      Caption         =   "&Enviar ao contador"
      Enabled         =   0   'False
      Height          =   420
      Left            =   1980
      TabIndex        =   31
      Top             =   7140
      Width           =   1860
   End
   Begin VB.CommandButton cmdGerar
      Caption         =   "&Gerar arquivo"
      Height          =   420
      Left            =   60
      TabIndex        =   30
      Top             =   7140
      Width           =   1800
   End
   Begin VB.Label lblStatus
      Height          =   915
      Left            =   60
      TabIndex        =   29
      Top             =   6150
      Width           =   8280
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSPED_Fiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Geração manual do SPED Fiscal (EFD ICMS/IPI) pelo SpedGerador.exe (modResumoFiscal.SPED_Gerar) e envio ao contador
' num e-mail só do arquivo (SPED_EnviarEmail). "Somente o inventário" marca só o Bloco H (os obrigatórios 0, E zerado,
' 1 e 9 vão junto, como no arquivo que o SPED.exe gerava para a contabilidade); "EFD completa" marca C e D (e o K se a
' empresa usa). Blocos desmarcados vão sem dados (abertura e encerramento). Uso/consumo e imobilizado são as marcações
' do cadastro do produto.
Dim dInventario() As Date          'data de cada inventário da lista (mesmo índice do cboInventario)
Dim sArquivoGerado As String, sDescGerado As String, sPeriodoGerado As String
Dim bBlocoKEmpresa As Boolean
Dim bCarregando As Boolean

Private Sub Form_Load()
    On Error GoTo ErrL
    bCarregando = True
    cboFinalidade.Clear
    cboFinalidade.AddItem "0 - Remessa do arquivo original"
    cboFinalidade.AddItem "1 - Remessa do arquivo substituto"
    cboFinalidade.ListIndex = 0

    Dim sK As String
    sK = Trim(CStr(SQLExecutaRetorno("SELECT TOP 1 CAST(SPEDBlocoK AS VARCHAR(10)) AS K FROM SPEDEmpresa WHERE IdFilial = 1", "K", "0")))
    bBlocoKEmpresa = (sK <> "" And sK <> "0")

    CarregarInventarios
    optInventario.Value = True
    bCarregando = False
    AplicarTipo
    Exit Sub
ErrL:
    MsgBox "Não foi possível abrir a tela do SPED Fiscal." & vbCrLf & vbCrLf & _
           "Verifique se os scripts do SPED (SPEDEmpresa, InventarioGerado...) já foram aplicados no banco." & vbCrLf & _
           Err.Description, vbExclamation
    bCarregando = False
End Sub

' Inventários com data, do mais novo para o mais antigo; na mesma data o fechado e o anual primeiro (a mesma
' preferência do envio automático).
Private Sub CarregarInventarios()
    Dim rRs As ADODB.Recordset
    Dim sTexto As String
    Dim n As Long

    cboInventario.Clear
    cboInventario.AddItem "(sem inventário)"
    cboInventario.ItemData(cboInventario.NewIndex) = 0
    ReDim dInventario(0)

    RsOpen rRs, "SELECT i.IdInventario, i.PeriodoInventarioTipo, i.PeriodoInventario, i.AnoInventario, i.DataInventario, ISNULL(i.Fechado, 0) AS Fechado, " & _
                "(SELECT COUNT(*) FROM TbInventariosItens ii WHERE ii.IdInventario = i.IdInventario AND ii.MetaCalculado > 0) AS Itens " & _
                "FROM TbInventarios i WHERE i.IdFilial = 1 AND i.DataInventario IS NOT NULL " & _
                "ORDER BY CAST(i.DataInventario AS DATE) DESC, i.Fechado DESC, CASE WHEN i.PeriodoInventarioTipo = 'Anual' THEN 0 ELSE 1 END, i.IdInventario DESC"
    Do While Not rRs.EOF
        sTexto = Format(rRs("DataInventario"), "dd/mm/yyyy") & " - " & rRs("PeriodoInventarioTipo") & ""
        If rRs("PeriodoInventarioTipo") & "" = "Mensal" Then sTexto = sTexto & " " & Format(Val(rRs("PeriodoInventario") & ""), "00") & "/" & rRs("AnoInventario")
        If rRs("PeriodoInventarioTipo") & "" = "Anual" Then sTexto = sTexto & " " & rRs("AnoInventario")
        sTexto = sTexto & IIf(CBool(rRs("Fechado")), " - fechado", " - aberto") & " - " & Format(rRs("Itens"), "#,##0") & " itens com saldo (nº " & rRs("IdInventario") & ")"
        cboInventario.AddItem sTexto
        cboInventario.ItemData(cboInventario.NewIndex) = rRs("IdInventario")
        n = cboInventario.NewIndex
        ReDim Preserve dInventario(n)
        dInventario(n) = CDate(Format(rRs("DataInventario"), "dd/mm/yyyy"))
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
    cboInventario.ListIndex = 0
End Sub

' Primeiro inventário de 31/12 da lista (o mais novo, fechado/anual na frente); sem nenhum, o mais novo.
Private Function InventarioSugerido() As Integer
    Dim i As Integer
    For i = 1 To cboInventario.ListCount - 1
        If Day(dInventario(i)) = 31 And Month(dInventario(i)) = 12 Then
            InventarioSugerido = i
            Exit Function
        End If
    Next
    If cboInventario.ListCount > 1 Then InventarioSugerido = 1
End Function

Private Sub AplicarTipo()
    Dim dIni As Date
    bCarregando = True
    If optInventario.Value Then
        chkBlocoC.Value = vbUnchecked
        chkBlocoD.Value = vbUnchecked
        chkBlocoK.Value = vbUnchecked
        If cboInventario.ListIndex <= 0 Then cboInventario.ListIndex = InventarioSugerido
        chkBlocoH.Value = IIf(cboInventario.ListIndex > 0, vbChecked, vbUnchecked)
        SugerirPeriodoInventario
    Else
        chkBlocoC.Value = vbChecked
        chkBlocoD.Value = vbChecked
        chkBlocoK.Value = IIf(bBlocoKEmpresa, vbChecked, vbUnchecked)
        cboInventario.ListIndex = 0
        chkBlocoH.Value = vbUnchecked
        'mês anterior
        dIni = DateSerial(Year(Date), Month(Date) - 1, 1)
        txtDataIni.Text = Format(dIni, "dd/mm/yyyy")
        txtDataFim.Text = Format(DateSerial(Year(dIni), Month(dIni) + 1, 0), "dd/mm/yyyy")
    End If
    bCarregando = False
    AtualizarProdutos
End Sub

' Período da EFD em que o inventário é escriturado: o de 31/12 vai na de fevereiro do ano seguinte; outra data, na do
' mês seguinte ao inventário.
Private Sub SugerirPeriodoInventario()
    Dim dIni As Date, d As Date
    If cboInventario.ListIndex <= 0 Then Exit Sub
    d = dInventario(cboInventario.ListIndex)
    If Day(d) = 31 And Month(d) = 12 Then
        dIni = DateSerial(Year(d) + 1, 2, 1)
    Else
        dIni = DateSerial(Year(d), Month(d) + 1, 1)
    End If
    txtDataIni.Text = Format(dIni, "dd/mm/yyyy")
    txtDataFim.Text = Format(DateSerial(Year(dIni), Month(dIni) + 1, 0), "dd/mm/yyyy")
End Sub

Private Sub AtualizarProdutos()
    chkUsoConsumo.Enabled = (chkBlocoH.Value = vbChecked)
    chkImobilizado.Enabled = (chkBlocoH.Value = vbChecked)
End Sub

Private Sub optInventario_Click()
    If Not bCarregando Then AplicarTipo
End Sub

Private Sub optCompleto_Click()
    If Not bCarregando Then AplicarTipo
End Sub

Private Sub cboInventario_Click()
    If bCarregando Then Exit Sub
    bCarregando = True
    chkBlocoH.Value = IIf(cboInventario.ListIndex > 0, vbChecked, vbUnchecked)
    If optInventario.Value Then SugerirPeriodoInventario
    bCarregando = False
    AtualizarProdutos
End Sub

Private Sub chkBlocoH_Click()
    If bCarregando Then Exit Sub
    If chkBlocoH.Value = vbChecked And cboInventario.ListIndex <= 0 Then
        bCarregando = True
        cboInventario.ListIndex = InventarioSugerido
        If optInventario.Value Then SugerirPeriodoInventario
        bCarregando = False
    End If
    AtualizarProdutos
End Sub

Private Function DataTexto(ByVal sTexto As String, ByRef dData As Date) As Boolean
    sTexto = Trim(sTexto)
    If Len(sTexto) <> 10 Then Exit Function
    If Not IsDate(sTexto) Then Exit Function
    dData = CDate(sTexto)
    DataTexto = True
End Function

Private Sub cmdGerar_Click()
    Dim dIni As Date, dFim As Date
    Dim sBlocos As String, sParam As String, sRes As String
    Dim sArq As String, sDesc As String, sPer As String, sAvisos As String
    Dim nInv As Long

    If Not DataTexto(txtDataIni.Text, dIni) Then
        MsgBox "Informe a data inicial do período (dd/mm/aaaa).", vbExclamation
        txtDataIni.SetFocus
        Exit Sub
    End If
    If Not DataTexto(txtDataFim.Text, dFim) Then
        MsgBox "Informe a data final do período (dd/mm/aaaa).", vbExclamation
        txtDataFim.SetFocus
        Exit Sub
    End If
    If dFim < dIni Then
        MsgBox "A data final é anterior à inicial.", vbExclamation
        Exit Sub
    End If
    If Month(dIni) <> Month(dFim) Or Year(dIni) <> Year(dFim) Then
        If MsgBox("A EFD ICMS/IPI é mensal e o período informado passa de um mês." & vbCrLf & "Gerar assim mesmo?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    If chkBlocoC.Value = vbChecked Then sBlocos = sBlocos & "C,"
    If chkBlocoD.Value = vbChecked Then sBlocos = sBlocos & "D,"
    If chkBlocoH.Value = vbChecked Then sBlocos = sBlocos & "H,"
    If chkBlocoK.Value = vbChecked Then sBlocos = sBlocos & "K,"
    If sBlocos = "" Then
        MsgBox "Marque ao menos um bloco com dados (C, D, H ou K).", vbExclamation
        Exit Sub
    End If
    sBlocos = Left$(sBlocos, Len(sBlocos) - 1)

    If cboInventario.ListIndex > 0 Then nInv = cboInventario.ItemData(cboInventario.ListIndex)
    If chkBlocoH.Value = vbChecked And nInv = 0 Then
        MsgBox "Escolha o inventário do Bloco H.", vbExclamation
        cboInventario.SetFocus
        Exit Sub
    End If

    sParam = "tipo=manual" & vbCrLf & _
             "pasta=" & SPED_Pasta & vbCrLf & _
             "dataini=" & Format(dIni, "dd") & "/" & Format(dIni, "mm") & "/" & Format(dIni, "yyyy") & vbCrLf & _
             "datafim=" & Format(dFim, "dd") & "/" & Format(dFim, "mm") & "/" & Format(dFim, "yyyy") & vbCrLf & _
             "finalidade=" & cboFinalidade.ListIndex & vbCrLf & _
             "inventario=" & nInv & vbCrLf & _
             "blocos=" & sBlocos & vbCrLf & _
             "usoconsumo=" & IIf(chkUsoConsumo.Value = vbChecked, "1", "0") & vbCrLf & _
             "imobilizado=" & IIf(chkImobilizado.Value = vbChecked, "1", "0")

    'confere (e corrige, quando há fonte confiável) os cadastros que vão no arquivo
    If Not ConferirCadastros(sParam) Then Exit Sub

    sArquivoGerado = ""
    cmdEnviar.Enabled = False
    lblStatus.Caption = "Gerando o arquivo do SPED..."
    Screen.MousePointer = vbHourglass
    cmdGerar.Enabled = False
    DoEvents
    sRes = SPED_Gerar(sParam, sArq, sDesc, sPer, sAvisos)
    cmdGerar.Enabled = True
    Screen.MousePointer = vbDefault

    If sRes = "OK" Then
        sArquivoGerado = sArq
        sDescGerado = sDesc
        sPeriodoGerado = sPer
        cmdEnviar.Enabled = True
        lblStatus.Caption = "Arquivo gerado: " & sArq
        MsgBox "Arquivo do SPED gerado:" & vbCrLf & sArq & IIf(sAvisos <> "", vbCrLf & vbCrLf & "Avisos:" & vbCrLf & sAvisos, ""), vbInformation
    Else
        lblStatus.Caption = "SPED não gerado: " & Mid$(sRes, IIf(Left$(sRes, 6) = "ERRO: ", 7, 1))
        MsgBox "O arquivo do SPED não foi gerado." & vbCrLf & vbCrLf & Mid$(sRes, IIf(Left$(sRes, 6) = "ERRO: ", 7, 1)) & vbCrLf & vbCrLf & _
               "Detalhes em SpedGerador\SpedGerador_log.txt.", vbExclamation
    End If
End Sub

' Antes de gerar: confere os cadastros que vão no arquivo (empresa, contador, fornecedores, produtos), corrige o que tem
' fonte confiável (XML da nota de entrada) depois de confirmar e lista o resto (frmSPED_Conferencia).
Private Function ConferirCadastros(ByVal sParam As String) As Boolean
    Dim sRes As String, nC As Long, nM As Long, nD As Long
    lblStatus.Caption = "Conferindo os cadastros que vão no arquivo..."
    Screen.MousePointer = vbHourglass
    DoEvents
    sRes = SPED_Conferir(sParam, nC, nM, nD)
    Screen.MousePointer = vbDefault
    lblStatus.Caption = ""
    If sRes <> "OK" Then
        ConferirCadastros = (MsgBox("Não foi possível conferir os cadastros:" & vbCrLf & Mid$(sRes, IIf(Left$(sRes, 6) = "ERRO: ", 7, 1)) & vbCrLf & vbCrLf & _
                                    "Gerar o arquivo assim mesmo?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes)
        Exit Function
    End If
    If nC + nM + nD = 0 Then
        ConferirCadastros = True
        Exit Function
    End If
    ConferirCadastros = frmSPED_Conferencia.Mostrar(sParam)
End Function

Private Sub cmdEnviar_Click()
    Dim sEmail As String, sErro As String

    If sArquivoGerado = "" Then Exit Sub
    sEmail = RF_EmailContador
    If InStr(1, sEmail, "@") = 0 Then
        MsgBox "E-mail do contador não cadastrado (menu Contabilidade > Contador).", vbExclamation
        Exit Sub
    End If
    If MsgBox("Enviar o arquivo do SPED para o contador?" & vbCrLf & vbCrLf & sEmail & vbCrLf & sArquivoGerado, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass
    cmdEnviar.Enabled = False
    DoEvents
    If SPED_EnviarEmail(sArquivoGerado, sDescGerado, sPeriodoGerado, False, sErro) Then
        Screen.MousePointer = vbDefault
        lblStatus.Caption = "Enviado para " & sEmail & ": " & sArquivoGerado
        MsgBox "E-mail com o SPED enviado para " & sEmail & "." & vbCrLf & vbCrLf & _
               "A entrega termina em segundo plano: aguarde alguns segundos antes de fechar o sistema.", vbInformation
    Else
        Screen.MousePointer = vbDefault
        MsgBox "O e-mail não foi enviado: " & sErro, vbExclamation
    End If
    cmdEnviar.Enabled = True
End Sub

Private Sub cmdPasta_Click()
    Dim sPasta As String
    On Error Resume Next
    If sArquivoGerado <> "" Then
        Shell "explorer.exe /select," & Chr(34) & sArquivoGerado & Chr(34), vbNormalFocus
        Exit Sub
    End If
    sPasta = SPED_Pasta
    If Dir$(sPasta, vbDirectory) = "" Then
        MsgBox "A pasta do SPED ainda não existe (é criada ao gerar o primeiro arquivo):" & vbCrLf & sPasta, vbInformation
        Exit Sub
    End If
    Shell "explorer.exe " & Chr(34) & sPasta & Chr(34), vbNormalFocus
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub
