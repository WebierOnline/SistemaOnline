VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCobrancaRegua 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Régua de cobrança por WhatsApp"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblRegras 
      Caption         =   "Regras (vale a regra da parcela mais atrasada do cliente)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5400
   End
   Begin MSFlexGridLib.MSFlexGrid grdRegras 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdNova 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3780
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Nova"
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExcluir 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3780
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "E&xcluir"
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuração geral"
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   5400
      Begin VB.Label lblHora 
         Caption         =   "Horário de envio:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtHoraIni 
         Height          =   315
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   6
         Top             =   330
         Width           =   735
      End
      Begin VB.Label lblAs 
         Caption         =   "às"
         Height          =   255
         Left            =   2620
         TabIndex        =   7
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox txtHoraFim 
         Height          =   315
         Left            =   2940
         MaxLength       =   5
         TabIndex        =   8
         Top             =   330
         Width           =   735
      End
      Begin VB.Label lblLimite 
         Caption         =   "Limite por dia:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   1500
      End
      Begin VB.TextBox txtLimite 
         Height          =   315
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   10
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblLimite2 
         Caption         =   "mensagens"
         Height          =   255
         Left            =   2620
         TabIndex        =   11
         Top             =   780
         Width           =   1500
      End
      Begin VB.Label lblIntervalo 
         Caption         =   "Intervalo entre mensagens:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtIntMin 
         Height          =   315
         Left            =   2340
         MaxLength       =   4
         TabIndex        =   13
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label lblA 
         Caption         =   "a"
         Height          =   255
         Left            =   3020
         TabIndex        =   14
         Top             =   1200
         Width           =   200
      End
      Begin VB.TextBox txtIntMax 
         Height          =   315
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   15
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label lblSeg 
         Caption         =   "segundos"
         Height          =   255
         Left            =   3920
         TabIndex        =   16
         Top             =   1200
         Width           =   900
      End
      Begin VB.CheckBox chkSabado 
         Caption         =   "Enviar aos sábados"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1620
         Width           =   2200
      End
      Begin VB.CheckBox chkDomingo 
         Caption         =   "Enviar aos domingos"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   1620
         Width           =   2400
      End
      Begin VB.Label lblPix 
         Caption         =   "Chave Pix da empresa (tag {CHAVE_PIX}):"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox txtPix 
         Height          =   315
         Left            =   240
         MaxLength       =   100
         TabIndex        =   20
         Top             =   2280
         Width           =   4935
      End
      Begin VB.Label lblAvisoZap 
         Caption         =   "O limite e o intervalo protegem o número da empresa: muitas mensagens seguidas podem fazer o WhatsApp bloquear o número."
         ForeColor       =   &H00808080&
         Height          =   855
         Left            =   240
         TabIndex        =   21
         Top             =   2700
         Width           =   3615
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarConfig 
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Top             =   3240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Salvar &config."
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
   Begin VB.Frame fraRegra 
      Caption         =   "Regra"
      Height          =   7935
      Left            =   5640
      TabIndex        =   23
      Top             =   120
      Width           =   7440
      Begin VB.Label lblDescricao 
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   25
         Top             =   330
         Width           =   4335
      End
      Begin VB.CheckBox chkAtiva 
         Caption         =   "Ativa"
         Height          =   255
         Left            =   5820
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblQuando 
         Caption         =   "Quando:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   780
         Width           =   1000
      End
      Begin VB.ComboBox cboSituacao 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   750
         Width           =   2535
      End
      Begin VB.Label lblDias 
         Caption         =   "Dias:"
         Height          =   255
         Left            =   4000
         TabIndex        =   29
         Top             =   780
         Width           =   500
      End
      Begin VB.TextBox txtDias 
         Height          =   315
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   30
         Top             =   750
         Width           =   615
      End
      Begin VB.Label lblRepetir 
         Caption         =   "Repetir a cada"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1100
      End
      Begin VB.TextBox txtRepetir 
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   32
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label lblRepetir2 
         Caption         =   "dia(s) (0 = não repete)     Máximo de envios por parcela:"
         Height          =   255
         Left            =   2000
         TabIndex        =   33
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtMax 
         Height          =   315
         Left            =   5820
         MaxLength       =   3
         TabIndex        =   34
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label lblMensagem 
         Caption         =   "Mensagem:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1620
         Width           =   1000
      End
      Begin VB.ComboBox cboTag 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1590
         Width           =   4335
      End
      Begin ChamaleonBtn.chameleonButton cmdInserirTag 
         Height          =   375
         Left            =   5760
         TabIndex        =   37
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Inserir tag"
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
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtMensagem 
         Height          =   2655
         Left            =   240
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   2040
         Width           =   6975
      End
      Begin VB.Label lblPrevia 
         Caption         =   "Prévia com dados de exemplo (os emojis aparecem como [nome]; no celular saem como emoji):"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   4800
         Width           =   6975
      End
      Begin VB.TextBox txtPrevia 
         Height          =   2055
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   5040
         Width           =   6975
      End
      Begin VB.Label lblTeste 
         Caption         =   "Testar no celular:"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   7260
         Width           =   1400
      End
      Begin VB.TextBox txtTeste 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   42
         Top             =   7230
         Width           =   1815
      End
      Begin ChamaleonBtn.chameleonButton cmdTeste 
         Height          =   375
         Left            =   3600
         TabIndex        =   43
         Top             =   7200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Enviar &teste"
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
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvar 
         Height          =   375
         Left            =   5880
         TabIndex        =   44
         Top             =   7200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Salvar regra"
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
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   375
      Left            =   11880
      TabIndex        =   45
      Top             =   8220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Fechar"
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
Attribute VB_Name = "frmCobrancaRegua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Régua de cobrança por WhatsApp (CobrancaRegra) e configuração geral (CobrancaConfig). Lógica em modCobranca.bas.

Private mIdAtual As Long          'regra na edição (0 = nova)
Private mAlterado As Boolean
Private mCarregando As Boolean
Private mCfg As CobConfig
Private mTags() As String

Private Sub Form_Load()
Dim desc() As String, emo() As String, i As Long

mCarregando = True
cboSituacao.AddItem "Antes do vencimento"
cboSituacao.ItemData(cboSituacao.NewIndex) = COB_ANTES
cboSituacao.AddItem "No dia do vencimento"
cboSituacao.ItemData(cboSituacao.NewIndex) = COB_NO_DIA
cboSituacao.AddItem "Após o vencimento"
cboSituacao.ItemData(cboSituacao.NewIndex) = COB_DEPOIS

Cob_ListaTags mTags, desc, emo
For i = 0 To UBound(mTags)
    cboTag.AddItem mTags(i) & "  -  " & desc(i)
Next
cboTag.ListIndex = 0

Cob_GarantirRegrasPadrao
Cob_LerConfig mCfg
txtHoraIni.Text = mCfg.HoraInicio
txtHoraFim.Text = mCfg.HoraFim
txtLimite.Text = mCfg.LimiteDia
txtIntMin.Text = mCfg.IntervaloMin
txtIntMax.Text = mCfg.IntervaloMax
chkSabado.Value = IIf(mCfg.EnviarSabado, vbChecked, vbUnchecked)
chkDomingo.Value = IIf(mCfg.EnviarDomingo, vbChecked, vbUnchecked)
txtPix.Text = mCfg.ChavePix
mCarregando = False

CarregarRegras 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not PodeSairDaRegra() Then Cancel = 1
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

'---------------------------------------------------------------------------
' Lista de regras
'---------------------------------------------------------------------------
Private Sub CarregarRegras(ByVal lSelecionar As Long)
Dim rs As ADODB.Recordset, r As Long, rSel As Long

mCarregando = True
With grdRegras
    .Redraw = False
    .Rows = 1
    .Cols = 3
    .TextMatrix(0, 0) = "Descrição": .ColWidth(0) = 2750
    .TextMatrix(0, 1) = "Quando": .ColWidth(1) = 1950
    .TextMatrix(0, 2) = "Ativa": .ColWidth(2) = 600
    .ColAlignment(2) = flexAlignCenterCenter
    On Error GoTo Fim
    Set rs = dbData.OpenRecordset("SELECT Id, Descricao, Situacao, Dias, Ativa FROM CobrancaRegra ORDER BY Situacao, Dias, Id")
    Do While Not rs.EOF
        .Rows = .Rows + 1
        r = .Rows - 1
        .RowData(r) = rs("Id")
        .TextMatrix(r, 0) = "" & rs("Descricao")
        .TextMatrix(r, 1) = Cob_SituacaoTexto(rs("Situacao"), Val("" & rs("Dias")))
        .TextMatrix(r, 2) = IIf(CBool(rs("Ativa")), "Sim", "Não")
        If rs("Id") = lSelecionar Then rSel = r
        rs.MoveNext
    Loop
    rs.Close
Fim:
    If rSel = 0 And .Rows > 1 Then rSel = 1
    If rSel > 0 Then .Row = rSel: .Col = 0: .ColSel = .Cols - 1
    .Redraw = True
End With
Err.Clear
mCarregando = False
If rSel > 0 Then MostrarRegra grdRegras.RowData(rSel) Else NovaRegra
End Sub

Private Sub grdRegras_Click()
TrocarRegra
End Sub

Private Sub grdRegras_RowColChange()
TrocarRegra
End Sub

Private Sub TrocarRegra()
Dim lId As Long

If mCarregando Then Exit Sub
If grdRegras.Row < 1 Then Exit Sub
lId = grdRegras.RowData(grdRegras.Row)
If lId = mIdAtual Then Exit Sub
If Not PodeSairDaRegra() Then Exit Sub
MostrarRegra lId
End Sub

'Alterou a regra e não salvou? Pergunta. False = ficar na regra atual.
Private Function PodeSairDaRegra() As Boolean
PodeSairDaRegra = True
If Not mAlterado Then Exit Function
Select Case MsgBox("A regra """ & txtDescricao.Text & """ foi alterada. Salvar as alterações?", vbQuestion + vbYesNoCancel, "Régua de cobrança")
    Case vbYes: PodeSairDaRegra = SalvarRegra()
    Case vbNo: mAlterado = False
    Case Else: PodeSairDaRegra = False
End Select
End Function

Private Sub MostrarRegra(ByVal lId As Long)
Dim rs As ADODB.Recordset, i As Long

On Error GoTo Falha
Set rs = dbData.OpenRecordset("SELECT * FROM CobrancaRegra WHERE Id = " & lId)
If rs.EOF Then rs.Close: NovaRegra: Exit Sub
mCarregando = True
mIdAtual = lId
txtDescricao.Text = "" & rs("Descricao")
chkAtiva.Value = IIf(CBool(rs("Ativa")), vbChecked, vbUnchecked)
For i = 0 To cboSituacao.ListCount - 1
    If cboSituacao.ItemData(i) = rs("Situacao") Then cboSituacao.ListIndex = i
Next
txtDias.Text = Val("" & rs("Dias"))
txtRepetir.Text = Val("" & rs("RepetirDias"))
txtMax.Text = Val("" & rs("MaxEnvios"))
txtMensagem.Text = "" & rs("Mensagem")
rs.Close
fraRegra.Caption = "Regra"
mCarregando = False
AjustarDias
mAlterado = False
AtualizarPrevia
Exit Sub
Falha:
mCarregando = False
MsgBox "Não foi possível ler a regra." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Régua de cobrança"
End Sub

Private Sub NovaRegra()
mCarregando = True
mIdAtual = 0
txtDescricao.Text = ""
chkAtiva.Value = vbChecked
cboSituacao.ListIndex = 2
txtDias.Text = "5"
txtRepetir.Text = "0"
txtMax.Text = "1"
txtMensagem.Text = "Olá, *{PRIMEIRO_NOME}*! Tudo bem? {ACENO}" & vbCrLf & vbCrLf & "{LISTA_PARCELAS}" & vbCrLf & vbCrLf & _
                   "Total: *R$ {VALOR_TOTAL}*" & vbCrLf & vbCrLf & "*{EMPRESA}*"
fraRegra.Caption = "Nova regra"
mCarregando = False
AjustarDias
mAlterado = False
AtualizarPrevia
End Sub

Private Sub cmdNova_Click()
If Not PodeSairDaRegra() Then Exit Sub
NovaRegra
txtDescricao.SetFocus
End Sub

Private Sub cmdExcluir_Click()
If mIdAtual = 0 Then NovaRegra: Exit Sub
If MsgBox("Excluir a regra """ & txtDescricao.Text & """?" & vbCrLf & vbCrLf & _
          "O histórico dos envios já feitos continua. Para só pausar a regra, desmarque ""Ativa"".", _
          vbQuestion + vbYesNo + vbDefaultButton2, "Régua de cobrança") <> vbYes Then Exit Sub
On Error GoTo Falha
dbData.Execute "DELETE FROM CobrancaRegra WHERE Id = " & mIdAtual
mAlterado = False
mIdAtual = 0
CarregarRegras 0
Exit Sub
Falha:
MsgBox "Não foi possível excluir a regra." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Régua de cobrança"
End Sub

'---------------------------------------------------------------------------
' Edição
'---------------------------------------------------------------------------
Private Function SituacaoAtual() As Integer
If cboSituacao.ListIndex < 0 Then SituacaoAtual = COB_DEPOIS Else SituacaoAtual = cboSituacao.ItemData(cboSituacao.ListIndex)
End Function

Private Sub AjustarDias()
If SituacaoAtual() = COB_NO_DIA Then
    txtDias.Text = "0"
    txtDias.Enabled = False
Else
    txtDias.Enabled = True
    If Val(txtDias.Text) < 1 Then txtDias.Text = "1"
End If
End Sub

Private Sub cboSituacao_Click()
If mCarregando Then Exit Sub
AjustarDias
mAlterado = True
End Sub

Private Sub txtDescricao_Change()
If Not mCarregando Then mAlterado = True
End Sub

Private Sub chkAtiva_Click()
If Not mCarregando Then mAlterado = True
End Sub

Private Sub txtDias_Change()
If Not mCarregando Then mAlterado = True
End Sub

Private Sub txtRepetir_Change()
If Not mCarregando Then mAlterado = True
End Sub

Private Sub txtMax_Change()
If Not mCarregando Then mAlterado = True
End Sub

Private Sub txtMensagem_Change()
If mCarregando Then Exit Sub
mAlterado = True
AtualizarPrevia
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Sub txtRepetir_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Sub txtLimite_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Sub txtIntMin_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Sub txtIntMax_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Function SoDigito(ByVal KeyAscii As Integer) As Integer
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then SoDigito = KeyAscii Else SoDigito = 0
End Function

Private Sub cmdInserirTag_Click()
If cboTag.ListIndex < 0 Then Exit Sub
txtMensagem.SelText = mTags(cboTag.ListIndex)
txtMensagem.SetFocus
End Sub

Private Function SalvarRegra() As Boolean
Dim nSit As Integer, nDias As Long, nRep As Long, nMax As Long, r As ADODB.Recordset

If Len(Trim$(txtDescricao.Text)) = 0 Then
    MsgBox "Informe a descrição da regra.", vbExclamation, "Régua de cobrança"
    txtDescricao.SetFocus
    Exit Function
End If
If Len(Trim$(txtMensagem.Text)) = 0 Then
    MsgBox "Escreva a mensagem da regra.", vbExclamation, "Régua de cobrança"
    txtMensagem.SetFocus
    Exit Function
End If
nSit = SituacaoAtual()
nDias = Val(txtDias.Text)
If nSit = COB_NO_DIA Then nDias = 0
If nSit <> COB_NO_DIA And (nDias < 1 Or nDias > 365) Then
    MsgBox "Informe os dias (de 1 a 365).", vbExclamation, "Régua de cobrança"
    txtDias.SetFocus
    Exit Function
End If
nRep = Val(txtRepetir.Text)
nMax = Val(txtMax.Text)
If nMax < 1 Then nMax = 1
If nRep = 0 Then nMax = 1

On Error GoTo Falha
If mIdAtual = 0 Then
    Set r = dbData.OpenRecordset("INSERT INTO CobrancaRegra (Descricao, Ativa, Situacao, Dias, RepetirDias, MaxEnvios, Mensagem) OUTPUT INSERTED.Id VALUES (" & _
        Cob_Q(Trim$(txtDescricao.Text)) & ", " & IIf(chkAtiva.Value = vbChecked, 1, 0) & ", " & nSit & ", " & nDias & ", " & nRep & ", " & nMax & ", " & _
        Cob_Q(txtMensagem.Text) & ")")
    If Not r.EOF Then mIdAtual = r(0)
    r.Close
Else
    dbData.Execute "UPDATE CobrancaRegra SET Descricao = " & Cob_Q(Trim$(txtDescricao.Text)) & ", Ativa = " & IIf(chkAtiva.Value = vbChecked, 1, 0) & _
        ", Situacao = " & nSit & ", Dias = " & nDias & ", RepetirDias = " & nRep & ", MaxEnvios = " & nMax & _
        ", Mensagem = " & Cob_Q(txtMensagem.Text) & " WHERE Id = " & mIdAtual
End If
mAlterado = False
SalvarRegra = True
Exit Function
Falha:
MsgBox "Não foi possível salvar a regra." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Régua de cobrança"
End Function

Private Sub cmdSalvar_Click()
If SalvarRegra() Then CarregarRegras mIdAtual
End Sub

'---------------------------------------------------------------------------
' Prévia e teste (dados de exemplo)
'---------------------------------------------------------------------------
Private Function Exemplo() As CobItem
Dim it As CobItem

it.Nome = "MARIA DA SILVA"
it.Mensagem = txtMensagem.Text
it.Lista = "• *Venda nº 1234* (parcela 2/3) – venc. " & Format$(Date - 5, "dd/mm/yyyy") & " – R$ 150,00 – 5 dia(s) em atraso" & vbLf & _
           "• *OS nº 567* – venc. " & Format$(Date, "dd/mm/yyyy") & " – R$ 200,00"
it.QtdParcelas = 2
it.Total = 350
it.DiasAtraso = 5
it.Vencimento = Date - 5
Exemplo = it
End Function

Private Sub AtualizarPrevia()
Dim it As CobItem, cfg As CobConfig

it = Exemplo()
cfg = mCfg
cfg.ChavePix = Trim$(txtPix.Text)
txtPrevia.Text = Cob_Texto(it, cfg, True)
End Sub

Private Sub txtPix_Change()
If Not mCarregando Then AtualizarPrevia
End Sub

Private Sub cmdTeste_Click()
Dim it As CobItem, cfg As CobConfig, sMsg As String

If Len(PDFZap_Digitos(txtTeste.Text)) < 10 Then
    MsgBox "Informe o celular com DDD para o teste.", vbExclamation, "Régua de cobrança"
    txtTeste.SetFocus
    Exit Sub
End If
it = Exemplo()
cfg = mCfg
cfg.ChavePix = Trim$(txtPix.Text)
Me.MousePointer = vbHourglass
If ZapOnlineEnviar(txtTeste.Text, Cob_Texto(it, cfg, False), "", sMsg) Then
    Me.MousePointer = vbDefault
    MsgBox "Mensagem de teste enviada (com dados de exemplo).", vbInformation, "Régua de cobrança"
Else
    Me.MousePointer = vbDefault
    MsgBox sMsg, vbExclamation, ZapOnlineTitulo("Régua de cobrança")
End If
End Sub

'---------------------------------------------------------------------------
' Configuração geral
'---------------------------------------------------------------------------
Private Sub cmdSalvarConfig_Click()
Dim c As CobConfig

c = mCfg
c.HoraInicio = Cob_Hora(txtHoraIni.Text)
c.HoraFim = Cob_Hora(txtHoraFim.Text)
If c.HoraFim <= c.HoraInicio Then
    MsgBox "O horário final deve ser depois do inicial.", vbExclamation, "Régua de cobrança"
    txtHoraFim.SetFocus
    Exit Sub
End If
c.LimiteDia = Val(txtLimite.Text)
If c.LimiteDia < 1 Then
    MsgBox "Informe o limite de mensagens por dia.", vbExclamation, "Régua de cobrança"
    txtLimite.SetFocus
    Exit Sub
End If
c.IntervaloMin = Val(txtIntMin.Text)
c.IntervaloMax = Val(txtIntMax.Text)
If c.IntervaloMin < 5 Then
    MsgBox "O intervalo mínimo entre mensagens é de 5 segundos (recomendado: 20 ou mais).", vbExclamation, "Régua de cobrança"
    txtIntMin.SetFocus
    Exit Sub
End If
If c.IntervaloMax < c.IntervaloMin Then c.IntervaloMax = c.IntervaloMin
c.EnviarSabado = (chkSabado.Value = vbChecked)
c.EnviarDomingo = (chkDomingo.Value = vbChecked)
c.ChavePix = Trim$(txtPix.Text)
If Not Cob_SalvarConfig(c) Then
    MsgBox "Não foi possível salvar a configuração.", vbExclamation, "Régua de cobrança"
    Exit Sub
End If
mCfg = c
txtHoraIni.Text = c.HoraInicio
txtHoraFim.Text = c.HoraFim
txtIntMax.Text = c.IntervaloMax
MsgBox "Configuração salva.", vbInformation, "Régua de cobrança"
End Sub
