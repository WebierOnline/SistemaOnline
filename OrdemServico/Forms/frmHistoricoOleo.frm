VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmHistoricoOleo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Trocas de Óleo por Placa"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optPlaca 
      Caption         =   "Placa"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   795
   End
   Begin VB.OptionButton optCliente 
      Caption         =   "Cliente"
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   795
   End
   Begin VB.ComboBox cboCliente 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3780
      TabIndex        =   7
      Top             =   60
      Width           =   6555
   End
   Begin MSFlexGridLib.MSFlexGrid lstOleo 
      Height          =   3660
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   6456
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtPlacaF 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   60
      Width           =   1800
   End
   Begin VB.Frame fraCarregando 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   3500
      TabIndex        =   2
      Top             =   1900
      Visible         =   0   'False
      Width           =   2800
      Begin VB.Label lblCarregando 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carregando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   3
         Top             =   90
         Width           =   2680
      End
   End
   Begin VB.Timer tmrCarregando 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   60
      Top             =   4020
   End
   Begin ChamaleonBtn.chameleonButton cmdHistorico 
      Height          =   315
      Left            =   8580
      TabIndex        =   4
      Top             =   4200
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Histórico"
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
      MICON           =   "frmHistoricoOleo.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFiltrar 
      Height          =   315
      Left            =   10380
      TabIndex        =   5
      Top             =   60
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Buscar"
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
      MICON           =   "frmHistoricoOleo.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCodBarra 
      Height          =   315
      Left            =   10260
      TabIndex        =   6
      Top             =   4200
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Copiar Cód. Barra"
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
      MICON           =   "frmHistoricoOleo.frx":0038
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
Attribute VB_Name = "frmHistoricoOleo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sPlacaBusca As String

Private iDots As Integer
Private lCodClienteFiltro As Long
Private bAutoCompletando As Boolean

Private Sub cboCliente_Change()
    Dim sTexto As String
    Dim sItem As String
    Dim i As Integer

    If bAutoCompletando Then Exit Sub

    sTexto = cboCliente.Text
    lCodClienteFiltro = 0

    If Len(sTexto) < 3 Then Exit Sub

    For i = 0 To cboCliente.ListCount - 1
        sItem = cboCliente.List(i)
        If UCase$(Left$(sItem, Len(sTexto))) = UCase$(sTexto) Then
            lCodClienteFiltro = cboCliente.ItemData(i)
            bAutoCompletando = True
            cboCliente.Text = sItem
            cboCliente.SelStart = Len(sTexto)
            cboCliente.SelLength = Len(sItem) - Len(sTexto)
            bAutoCompletando = False
            Exit For
        End If
    Next i
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Form_Load()
    iDots = 0
    fraCarregando.Visible = False
    ConfigurarGrid
    PreencherClientes

    optPlaca.Value = True
    optCliente.Value = False
    txtPlacaF.Enabled = True
    cboCliente.Enabled = False

    If sPlacaBusca <> "" Then
        txtPlacaF.Text = sPlacaBusca
        sPlacaBusca = ""
        CarregarGrid
    End If
End Sub

Private Sub Form_Activate()
    txtPlacaF.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    End If
End Sub

Private Sub ConfigurarGrid()
    With lstOleo
        .Cols = 9
        .Rows = 1
        .ColWidth(0) = 600
        .ColWidth(1) = 1000
        .ColWidth(2) = 2100
        .ColWidth(3) = 2100
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 2200
        .ColWidth(7) = 4100
        .ColWidth(8) = 1400
        .Row = 0: .Col = 0: .Text = "OS"
        .Row = 0: .Col = 1: .Text = "DATA"
        .Row = 0: .Col = 2: .Text = "CLIENTE"
        .Row = 0: .Col = 3: .Text = "VEICULO"
        .Row = 0: .Col = 4: .Text = "PLACA": .CellBackColor = &HE0E0E0
        .Row = 0: .Col = 5: .Text = "KM"
        .Row = 0: .Col = 6: .Text = "PRÓXIMO": .CellBackColor = &HE0E0E0
        .Row = 0: .Col = 7: .Text = "ÓLEO"
        .Row = 0: .Col = 8: .Text = "CÓD. BARRA"
        .AllowUserResizing = 1
        .SelectionMode = 1
    End With
End Sub

Private Sub PreencherClientes()
    Dim rCli As ADODB.Recordset
    Dim sql As String

    cboCliente.Clear
    sql = "SELECT codigo, nome FROM cliente ORDER BY nome"
    RsOpen rCli, sql
    Do While Not rCli.EOF
        cboCliente.AddItem rCli("nome")
        cboCliente.ItemData(cboCliente.NewIndex) = rCli("codigo")
        rCli.MoveNext
    Loop
    If rCli.State <> 0 Then rCli.Close
End Sub

Private Sub optPlaca_Click()
    txtPlacaF.Enabled = True
    cboCliente.Enabled = False
    If Me.Visible Then txtPlacaF.SetFocus
End Sub

Private Sub optCliente_Click()
    txtPlacaF.Enabled = False
    cboCliente.Enabled = True
    If Me.Visible Then cboCliente.SetFocus
End Sub

Private Sub CarregarGrid()
    Dim rOleo As ADODB.Recordset
    Dim sql As String
    Dim sWhere As String
    Dim n As Integer
    Dim sVeiculo As String
    Dim sProximo As String
    Dim lKmAtual As Long
    Dim lLimiteKm As Long
    Dim lLimitePrazo As Long

    If optCliente.Value = True Then
        If lCodClienteFiltro <= 0 Then
            MsgBox "Selecione um cliente para consultar!", vbInformation, "Aviso do Sistema"
            Exit Sub
        End If
        sWhere = "(cliente.codigo = " & lCodClienteFiltro & ") "
    Else
        If Trim(txtPlacaF.Text) = "" Then
            MsgBox "Digite a placa para consultar!", vbInformation, "Aviso do Sistema"
            Exit Sub
        End If
        sWhere = "(OS_Equipamento_Auto.placa LIKE '%" & Replace(Trim(txtPlacaF.Text), "'", "''") & "%') "
    End If

    fraCarregando.Visible = True
    fraCarregando.ZOrder 0
    tmrCarregando.Enabled = True
    Me.Refresh
    DoEvents

    lstOleo.Rows = 1

    sql = "SELECT OS.COD_OS, OS.DATA_ENTRADA, cliente.nome, " & _
          "OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.modelo, OS_Equipamento_Auto.ano, " & _
          "OS_Equipamento_Auto.placa, OS_Equipamento_Auto.km, produtos.descricao AS var_oleo, produtos.cod_barra AS var_codbarra, " & _
          "OS_ControleOleo.LIMITE_KM, OS_ControleOleo.LIMITE_PRAZO " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE " & _
          "INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _
          "INNER JOIN pedidos_itens ON pedidos_itens.COD_PEDIDO = OS.COD_PEDIDO " & _
          "INNER JOIN produtos ON produtos.CODIGO = pedidos_itens.COD_PRODUTO " & _
          "LEFT JOIN OS_ControleOleo ON OS_ControleOleo.COD_PRODUTO = produtos.CODIGO " & _
          "WHERE " & sWhere & _
          "AND (produtos.descricao LIKE '%OLEO%' OR produtos.descricao LIKE '%ÓLEO%') " & _
          "ORDER BY OS.DATA_ENTRADA DESC"

    RsOpen rOleo, sql
    n = 1
    Do While Not rOleo.EOF
        sVeiculo = Trim(ValidateNull(rOleo("fabricante")) & " / " & ValidateNull(rOleo("modelo")) & " / " & ValidateNull(rOleo("ano")))

        lKmAtual = Val(ValidateNull(rOleo("km")))
        lLimiteKm = Val(ValidateNull(rOleo("LIMITE_KM")))
        lLimitePrazo = Val(ValidateNull(rOleo("LIMITE_PRAZO")))
        sProximo = ""
        If lLimiteKm > 0 Then
            sProximo = CStr(lKmAtual + lLimiteKm) & "km"
        End If
        If lLimitePrazo > 0 Then
            If sProximo <> "" Then sProximo = sProximo & " ou "
            sProximo = sProximo & Format(DateAdd("m", lLimitePrazo, rOleo("DATA_ENTRADA")), "dd/mm/yy")
        End If

        With lstOleo
            .Rows = n + 1
            .Row = n: .Col = 0: .Text = ValidateNull(rOleo("COD_OS"))
            .Row = n: .Col = 1: .Text = Format(rOleo("DATA_ENTRADA"), "dd/mm/yy")
            .Row = n: .Col = 2: .Text = ValidateNull(rOleo("nome"))
            .Row = n: .Col = 3: .Text = sVeiculo
            .Row = n: .Col = 4: .Text = ValidateNull(rOleo("placa"))
            .Row = n: .Col = 5: .Text = ValidateNull(rOleo("km"))
            .Row = n: .Col = 6: .Text = sProximo: .CellBackColor = &HE0E0E0
            .Row = n: .Col = 7: .Text = ValidateNull(rOleo("var_oleo"))
            .Row = n: .Col = 8: .Text = ValidateNull(rOleo("var_codbarra"))
        End With
        n = n + 1
        rOleo.MoveNext
    Loop
    If rOleo.State <> 0 Then rOleo.Close

    tmrCarregando.Enabled = False
    fraCarregando.Visible = False

    If lstOleo.Rows = 1 Then
        MsgBox "Nenhuma troca de óleo encontrada!", vbInformation, "Aviso do Sistema"
    End If
End Sub

Private Sub txtPlacaF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        CarregarGrid
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdFiltrar_Click()
    CarregarGrid
End Sub

Private Sub cmdCodBarra_Click()
    If lstOleo.Row < 1 Then
        MsgBox "Selecione um produto no grid.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    Dim sCodBarra As String
    sCodBarra = Trim(lstOleo.TextMatrix(lstOleo.Row, 8))
    If sCodBarra = "" Then
        MsgBox "Produto sem código de barra cadastrado.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    Clipboard.Clear
    Clipboard.SetText sCodBarra
    MsgBox "Código de barra copiado para a área de transferência!", vbInformation, "Aviso do Sistema"
End Sub

Private Sub cmdHistorico_Click()
    If lstOleo.Row < 1 Then
        MsgBox "Selecione um registro.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    Dim sPlaca As String
    sPlaca = Trim(lstOleo.TextMatrix(lstOleo.Row, 4))
    If sPlaca = "" Then
        MsgBox "Registro sem placa.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    OS_Consulta.sPlacaBusca = sPlaca
    OS_Consulta.Show vbModal
    Unload OS_Consulta
End Sub


Private Sub tmrCarregando_Timer()
    iDots = (iDots Mod 3) + 1
    lblCarregando.Caption = "Carregando" & String(iDots, ".")
End Sub
