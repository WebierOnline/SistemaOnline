VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOleoVencidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRODUTOS PRAZO LIMITE ESTOURADO"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid lstVencidos 
      Height          =   4020
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   7091
      _Version        =   393216
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizar 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   4140
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "ATUALIZAR LISTA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOleoVencidos.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdNotificarEmail 
      Height          =   315
      Left            =   11340
      TabIndex        =   2
      Top             =   4140
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "NOTIFICAR / E-MAIL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOleoVencidos.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdNotificarZap 
      Height          =   315
      Left            =   9000
      TabIndex        =   3
      Top             =   4140
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "NOTIFICAR / WHATSAPP"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOleoVencidos.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmOleoVencidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long

Private Function ToUTF8Bytes(ByVal sTexto As String) As Byte()
    Dim oStream As ADODB.Stream

    Set oStream = New ADODB.Stream
    oStream.Type = adTypeText
    oStream.Charset = "utf-8"
    oStream.Open
    oStream.WriteText sTexto
    oStream.Position = 0
    oStream.Type = adTypeBinary
    oStream.Position = 3 ' pula o BOM (EF BB BF) que o ADODB.Stream grava
    ToUTF8Bytes = oStream.Read
    oStream.Close
End Function

Private Function URLEncodeUTF8(ByVal sTexto As String) As String
    Dim aBytes() As Byte
    Dim i As Long
    Dim b As Byte
    Dim sResult As String

    aBytes = ToUTF8Bytes(sTexto)

    For i = LBound(aBytes) To UBound(aBytes)
        b = aBytes(i)
        Select Case b
            Case 48 To 57, 65 To 90, 97 To 122 ' 0-9 A-Z a-z
                sResult = sResult & Chr(b)
            Case 45, 46, 95, 126 ' - . _ ~
                sResult = sResult & Chr(b)
            Case Else
                sResult = sResult & "%" & Right$("0" & Hex(b), 2)
        End Select
    Next i

    URLEncodeUTF8 = sResult
End Function

Private Function FormatarTelefoneWhatsapp(ByVal sTelefone As String) As String
    Dim sSoNumeros As String
    Dim i As Integer
    Dim c As String

    sSoNumeros = ""
    For i = 1 To Len(sTelefone)
        c = Mid$(sTelefone, i, 1)
        If c >= "0" And c <= "9" Then sSoNumeros = sSoNumeros & c
    Next i

    If Left$(sSoNumeros, 2) <> "55" Then sSoNumeros = "55" & sSoNumeros

    FormatarTelefoneWhatsapp = sSoNumeros
End Function

Private Sub Form_Load()
    ConfigurarGrid
    CarregarGrid
End Sub

Private Sub ConfigurarGrid()
    With lstVencidos
        .Cols = 8
        .Rows = 1
        .FixedCols = 0
        .ColWidth(0) = 700
        .ColWidth(1) = 2000
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 700
        .ColWidth(6) = 3600
        .ColWidth(7) = 0
        .TextMatrix(0, 0) = "COD.OS"
        .TextMatrix(0, 1) = "CLIENTE"
        .TextMatrix(0, 2) = "VEICULO / PLACA / KM"
        .TextMatrix(0, 3) = "TROCA"
        .TextMatrix(0, 4) = "VENCIDO": .CellBackColor = &HE0E0E0
        .TextMatrix(0, 5) = "DIAS"
        .TextMatrix(0, 6) = "ÓLEO"
        .AllowUserResizing = 1
        .SelectionMode = 1
    End With
End Sub

Private Sub CarregarGrid()
    Dim r As ADODB.Recordset
    Dim sql As String
    Dim n As Integer
    Dim sVeiculo As String
    Dim sVeiculoPlacaKm As String
    Dim dVencido As Date

    lstVencidos.Rows = 1

    sql = "SELECT COD_OS, Cliente, FABRICANTE, MODELO, ANO, PLACA, KM, DATA_TERMINO, Oleo, Vencido " & _
          "FROM (" & _
          "    SELECT OS.COD_OS, cliente.nome AS Cliente, " & _
          "        OS_Equipamento_Auto.fabricante AS FABRICANTE, OS_Equipamento_Auto.modelo AS MODELO, OS_Equipamento_Auto.ano AS ANO, " & _
          "        OS_Equipamento_Auto.placa AS PLACA, OS_Equipamento_Auto.km AS KM, OS.DATA_TERMINO, " & _
          "        produtos.descricao AS Oleo, " & _
          "        DATEADD(month, OS_ControleOleo.LIMITE_PRAZO, OS.DATA_TERMINO) AS Vencido, " & _
          "        ROW_NUMBER() OVER (PARTITION BY OS_Equipamento_Auto.placa, produtos.codigo ORDER BY OS.DATA_TERMINO DESC) AS rn " & _
          "    FROM OS " & _
          "    INNER JOIN cliente ON cliente.CODIGO = OS.COD_CLIENTE " & _
          "    INNER JOIN OS_Equipamento_Auto ON OS_Equipamento_Auto.COD_OS = OS.COD_OS " & _
          "    INNER JOIN pedidos_itens ON pedidos_itens.COD_PEDIDO = OS.COD_PEDIDO " & _
          "    INNER JOIN produtos ON produtos.CODIGO = pedidos_itens.COD_PRODUTO " & _
          "    INNER JOIN OS_ControleOleo ON OS_ControleOleo.COD_PRODUTO = produtos.CODIGO " & _
          "    WHERE (OS.DATA_TERMINO IS NOT NULL) AND (OS_ControleOleo.LIMITE_PRAZO > 0) " & _
          ") AS Ult " & _
          "WHERE (rn = 1) AND (Vencido <= GETDATE()) " & _
          "ORDER BY Vencido ASC"

    RsOpen r, sql
    n = 1
    Do While Not r.EOF
        sVeiculo = Trim(ValidateNull(r("FABRICANTE")) & " / " & ValidateNull(r("MODELO")) & " / " & ValidateNull(r("ANO")))
        sVeiculoPlacaKm = sVeiculo & " / " & ValidateNull(r("PLACA")) & " / " & ValidateNull(r("KM"))
        dVencido = r("Vencido")
        With lstVencidos
            .Rows = n + 1
            .Row = n: .Col = 0: .Text = ValidateNull(r("COD_OS"))
            .Row = n: .Col = 1: .Text = ValidateNull(r("Cliente"))
            .Row = n: .Col = 2: .Text = sVeiculoPlacaKm
            .Row = n: .Col = 3: .Text = Format(r("DATA_TERMINO"), "dd/mm/yy")
            .Row = n: .Col = 4: .Text = Format(dVencido, "dd/mm/yy"): .CellBackColor = &HE0E0E0
            .Row = n: .Col = 5: .Text = DateDiff("d", dVencido, Date)
            .Row = n: .Col = 6: .Text = ValidateNull(r("Oleo"))
            .Row = n: .Col = 7: .Text = ValidateNull(r("PLACA"))
        End With
        n = n + 1
        r.MoveNext
    Loop
    If r.State <> 0 Then r.Close

    If lstVencidos.Rows = 1 Then
        MsgBox "Nenhuma troca de óleo vencida encontrada!", vbInformation, "Aviso do Sistema"
    End If
End Sub

Private Sub cmdAtualizar_Click()
    CarregarGrid
End Sub

Private Sub cmdNotificarEmail_Click()
    Dim lCodOS As Long
    Dim lCodCliente As Long
    Dim sPlaca As String
    Dim sNomeCliente As String
    Dim sEmailCliente As String
    Dim sNomeOficina As String
    Dim sCelularOficina As String
    Dim dDataManutencao As Date
    Dim sListaItens As String
    Dim sAssunto As String
    Dim corpoEmail As String
    Dim i As Long
    Dim rInfo As ADODB.Recordset
    Dim sSQLInfo As String
    Dim sistNFe As snfe.Util
    Dim pathAnexo() As String
    Dim emailCC() As String

    If lstVencidos.Row < 1 Then
        MsgBox "Selecione uma linha no grid.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    lCodOS = Val(lstVencidos.TextMatrix(lstVencidos.Row, 0))
    sNomeCliente = lstVencidos.TextMatrix(lstVencidos.Row, 1)

    sSQLInfo = "SELECT OS_Equipamento_Auto.placa, OS.COD_CLIENTE " & _
        "FROM OS INNER JOIN OS_Equipamento_Auto ON OS_Equipamento_Auto.COD_OS = OS.COD_OS " & _
        "WHERE (OS.COD_OS = " & lCodOS & ")"
    RsOpen rInfo, sSQLInfo
    If rInfo.EOF Then
        If rInfo.State <> 0 Then rInfo.Close
        MsgBox "Esta OS não tem dados de veículo cadastrados (OS_Equipamento_Auto).", vbExclamation, "Aviso do Sistema"
        Exit Sub
    End If
    sPlaca = ValidateNull(rInfo("placa"))
    lCodCliente = ValidateNull(rInfo("COD_CLIENTE"))
    If rInfo.State <> 0 Then rInfo.Close

    sListaItens = ""
    dDataManutencao = 0

    If Trim(sPlaca) = "" Then
        ' sem placa cadastrada: nao da pra agrupar por veiculo, notifica so o produto desta OS
        i = lstVencidos.Row
        sListaItens = "<li>&#128680; <b>" & lstVencidos.TextMatrix(i, 6) & "</b>: Venceu em " & lstVencidos.TextMatrix(i, 4) & " (Atrasado há " & lstVencidos.TextMatrix(i, 5) & " dias)</li>"
        dDataManutencao = CDate(lstVencidos.TextMatrix(i, 3))
    Else
        For i = 1 To lstVencidos.Rows - 1
            If lstVencidos.TextMatrix(i, 7) = sPlaca Then
                sListaItens = sListaItens & "<li>&#128680; <b>" & lstVencidos.TextMatrix(i, 6) & "</b>: Venceu em " & lstVencidos.TextMatrix(i, 4) & " (Atrasado há " & lstVencidos.TextMatrix(i, 5) & " dias)</li>"
                If CDate(lstVencidos.TextMatrix(i, 3)) > dDataManutencao Then dDataManutencao = CDate(lstVencidos.TextMatrix(i, 3))
            End If
        Next i
    End If

    If sListaItens = "" Then Exit Sub

    sEmailCliente = SQLExecutaRetorno("SELECT Correio_eletronico FROM cliente WHERE (codigo = " & lCodCliente & ")", "Correio_eletronico", "")
    If Trim(sEmailCliente) = "" Then
        sEmailCliente = InputBox("Cliente sem e-mail cadastrado. Informe o e-mail do destinatário:", "Envio de Email", "")
        If Vazio(sEmailCliente) Then Exit Sub
        If InStr(1, sEmailCliente, "@") = 0 Then
            MsgBox "E-mail inválido! O endereço não contém ""@"".", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If

    sNomeOficina = SQLExecutaRetorno("SELECT Fantasia FROM Empresa", "Fantasia", "")
    sCelularOficina = SQLExecutaRetorno("SELECT CELULAR FROM Empresa", "CELULAR", "")

    corpoEmail = "Olá, " & sNomeCliente & ", tudo bem? &#128295;<br><br>" & _
        "Aqui é da " & sNomeOficina & ".<br><br>" & _
        "Notamos no nosso sistema que você realizou uma manutenção conosco no dia " & Format(dDataManutencao, "dd/mm/yyyy") & ". Algumas das peças substituídas ou revisadas possuem um prazo limite de validade ou quilometragem para garantir a sua segurança.<br><br>" & _
        "Identificamos os seguintes itens pendentes:<br><br>" & _
        "<ul>" & sListaItens & "</ul><br>" & _
        "Rodar com esses componentes fora do prazo do fabricante pode colocar o seu veículo em risco e gerar prejuízos maiores.<br><br>" & _
        "Gostaria de aproveitar para <b>agendar a manutenção</b> e <b>deixar o seu carro 100% protegido</b>?<br><br>" & _
        "Se preferir, basta nos responder por aqui com o dia e horário de sua preferência que nós reservaremos a sua vaga! &#128522;<br><br>" & _
        "Caso deseje agendar pelo whatsapp, segue nosso contato " & sCelularOficina

    sAssunto = "PRODUTOS FORA DO PRAZO LIMITE"

    On Error GoTo TrataErroEmail

    Set sistNFe = New snfe.Util
    iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)

    ReDim pathAnexo(0)
    pathAnexo(0) = ""
    ReDim emailCC(0)
    emailCC(0) = sEmailCliente

    Screen.MousePointer = vbHourglass
    iRetorno = sistNFe.EmailEnviar(sEmailCliente, sAssunto, corpoEmail, pathAnexo, emailCC)
    Screen.MousePointer = vbDefault

    If iRetorno Then
        MsgBox "Email enviado com sucesso!", vbInformation, "Email OK!"
    Else
        MsgBox "Falha ao enviar o email. Verifique a configuração de envio.", vbExclamation, "Erro: Envio de Email"
    End If

    Set sistNFe = Nothing
    Exit Sub

TrataErroEmail:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "Erro: Envio de Email"
    Err.Clear
    Set sistNFe = Nothing
End Sub

Private Sub cmdNotificarZap_Click()
    Dim lCodOS As Long
    Dim lCodCliente As Long
    Dim sPlaca As String
    Dim sNomeCliente As String
    Dim sCelularCliente As String
    Dim sNomeOficina As String
    Dim dDataManutencao As Date
    Dim sListaItens As String
    Dim sMensagem As String
    Dim i As Long
    Dim rInfo As ADODB.Recordset
    Dim sSQLInfo As String
    Dim sTelefone As String
    Dim sURL As String
    Dim hWndZap As Long
    Dim lThreadAtual As Long
    Dim lThreadAlvo As Long
    Dim lProcessoAlvo As Long

    If lstVencidos.Row < 1 Then
        MsgBox "Selecione uma linha no grid.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    If Not WhatsAppEstaAberto() Then
        MsgBox "Atenção: O WhatsApp (App ou aba do navegador) não foi detectado aberto neste computador." & vbCrLf & _
            "Por favor, abra o WhatsApp para que o sistema possa enviar as mensagens.", vbExclamation, "WhatsApp não encontrado"
        Exit Sub
    End If

    lCodOS = Val(lstVencidos.TextMatrix(lstVencidos.Row, 0))
    sNomeCliente = lstVencidos.TextMatrix(lstVencidos.Row, 1)

    sSQLInfo = "SELECT OS_Equipamento_Auto.placa, OS.COD_CLIENTE " & _
        "FROM OS INNER JOIN OS_Equipamento_Auto ON OS_Equipamento_Auto.COD_OS = OS.COD_OS " & _
        "WHERE (OS.COD_OS = " & lCodOS & ")"
    RsOpen rInfo, sSQLInfo
    If rInfo.EOF Then
        If rInfo.State <> 0 Then rInfo.Close
        MsgBox "Esta OS não tem dados de veículo cadastrados (OS_Equipamento_Auto).", vbExclamation, "Aviso do Sistema"
        Exit Sub
    End If
    sPlaca = ValidateNull(rInfo("placa"))
    lCodCliente = ValidateNull(rInfo("COD_CLIENTE"))
    If rInfo.State <> 0 Then rInfo.Close

    sListaItens = ""
    dDataManutencao = 0

    If Trim(sPlaca) = "" Then
        ' sem placa cadastrada: nao da pra agrupar por veiculo, notifica so o produto desta OS
        i = lstVencidos.Row
        sListaItens = "- " & lstVencidos.TextMatrix(i, 6) & ": Venceu em " & lstVencidos.TextMatrix(i, 4) & " (Atrasado há " & lstVencidos.TextMatrix(i, 5) & " dias)" & vbCrLf
        dDataManutencao = CDate(lstVencidos.TextMatrix(i, 3))
    Else
        For i = 1 To lstVencidos.Rows - 1
            If lstVencidos.TextMatrix(i, 7) = sPlaca Then
                sListaItens = sListaItens & "- " & lstVencidos.TextMatrix(i, 6) & ": Venceu em " & lstVencidos.TextMatrix(i, 4) & " (Atrasado há " & lstVencidos.TextMatrix(i, 5) & " dias)" & vbCrLf
                If CDate(lstVencidos.TextMatrix(i, 3)) > dDataManutencao Then dDataManutencao = CDate(lstVencidos.TextMatrix(i, 3))
            End If
        Next i
    End If

    If sListaItens = "" Then Exit Sub

    sCelularCliente = SQLExecutaRetorno("SELECT celular FROM cliente WHERE (codigo = " & lCodCliente & ")", "celular", "")
    If Trim(sCelularCliente) = "" Then
        sCelularCliente = InputBox("Cliente sem celular cadastrado. Informe o número (com DDD):", "Envio de WhatsApp", "")
        If Vazio(sCelularCliente) Then Exit Sub
    End If

    sTelefone = FormatarTelefoneWhatsapp(sCelularCliente)
    If Len(sTelefone) < 12 Then
        MsgBox "Número de celular inválido!", vbExclamation, "Aviso"
        Exit Sub
    End If

    sNomeOficina = SQLExecutaRetorno("SELECT Fantasia FROM Empresa", "Fantasia", "")

    sMensagem = "Olá, " & sNomeCliente & ", tudo bem?" & vbCrLf & vbCrLf & _
        "Aqui é da " & sNomeOficina & "." & vbCrLf & vbCrLf & _
        "Notamos no nosso sistema que você realizou uma manutenção conosco no dia " & Format(dDataManutencao, "dd/mm/yyyy") & ". Algumas das peças substituídas ou revisadas possuem um prazo limite de validade ou quilometragem para garantir a sua segurança." & vbCrLf & vbCrLf & _
        "Identificamos os seguintes itens pendentes:" & vbCrLf & vbCrLf & _
        sListaItens & vbCrLf & _
        "Rodar com esses componentes fora do prazo do fabricante pode colocar o seu veículo em risco e gerar prejuízos maiores." & vbCrLf & vbCrLf & _
        "Gostaria de aproveitar para agendar a manutenção e deixar o seu carro 100% protegido?" & vbCrLf & vbCrLf & _
        "Se preferir, basta nos responder por aqui com o dia e horário de sua preferência que nós reservaremos a sua vaga!"

    sURL = "https://web.whatsapp.com/send?phone=" & sTelefone & "&text=" & URLEncodeUTF8(sMensagem)

    ShellExecute Me.hwnd, "open", sURL, vbNullString, vbNullString, 1

    Sleep 4000

    hWndZap = HwndJanelaWhatsApp()
    If hWndZap <> 0 Then
        lThreadAtual = GetCurrentThreadId()
        lThreadAlvo = GetWindowThreadProcessId(hWndZap, lProcessoAlvo)
        Call AttachThreadInput(lThreadAtual, lThreadAlvo, 1)
        SetForegroundWindow hWndZap
        Call AttachThreadInput(lThreadAtual, lThreadAlvo, 0)
        Sleep 300
    End If

    On Error Resume Next
    SendKeys "~"
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "A mensagem foi carregada no WhatsApp Web, mas não consegui apertar Enter automaticamente." & vbCrLf & "Confira a aba aberta no navegador e clique em enviar manualmente.", vbInformation, "Envio de WhatsApp"
    End If
    On Error GoTo 0
End Sub

