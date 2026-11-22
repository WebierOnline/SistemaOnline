VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCampanhaLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Campanhas de promoção por WhatsApp"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   16815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblResumo 
      Caption         =   ""
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16575
   End
   Begin MSFlexGridLib.MSFlexGrid grdCamp 
      Height          =   3600
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   16575
      _ExtentX        =   29237
      _ExtentY        =   6350
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdPausar 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4110
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Pausar"
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
   Begin ChamaleonBtn.chameleonButton cmdRetomar 
      Height          =   375
      Left            =   1380
      TabIndex        =   3
      Top             =   4110
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Retomar"
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
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   4110
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancelar campanha"
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
   Begin ChamaleonBtn.chameleonButton cmdAtualizar 
      Height          =   375
      Left            =   4380
      TabIndex        =   5
      Top             =   4110
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Atualizar"
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
   Begin VB.Label lblDest 
      Caption         =   "Clientes da campanha:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4620
      Width           =   8000
   End
   Begin MSFlexGridLib.MSFlexGrid grdDest 
      Height          =   4335
      Left            =   120
      TabIndex        =   7
      Top             =   4860
      Width           =   16575
      _ExtentX        =   29237
      _ExtentY        =   7647
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   375
      Left            =   15480
      TabIndex        =   8
      Top             =   9300
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
Attribute VB_Name = "frmCampanhaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Acompanhamento das campanhas de promoção (modCampanha): andamento de cada campanha, clientes e erros,
'pausar / retomar / cancelar (o servidor obedece antes da próxima mensagem).

Private Const TITULO As String = "Promoções por WhatsApp"
Private mSemEvento As Boolean
Private mIniciou As Boolean

Private Sub Form_Activate()
If mIniciou Then Exit Sub
mIniciou = True
Carregar 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub cmdAtualizar_Click()
Carregar CampanhaAtual()
End Sub

Private Function CampanhaAtual() As Long
If grdCamp.Row >= 1 And grdCamp.Row < grdCamp.Rows Then CampanhaAtual = grdCamp.RowData(grdCamp.Row)
End Function

Private Sub Carregar(ByVal lSel As Long)
Dim rs As ADODB.Recordset, r As Long, nSel As Long, c As Integer

Me.MousePointer = vbHourglass
lblResumo.Caption = "Promoções enviadas hoje: " & Camp_EnviadasHoje() & " de " & Camp_LimiteDia() & " (limite por dia). " & _
                    "O servidor envia a cada 30 minutos, no horário e nos dias da Régua de cobrança, depois das cobranças."
mSemEvento = True
With grdCamp
    .Redraw = False
    .Rows = 1
    .Cols = 10
    .TextMatrix(0, 0) = "Nº"
    .TextMatrix(0, 1) = "Campanha"
    .TextMatrix(0, 2) = "Conteúdo"
    .TextMatrix(0, 3) = "Criada"
    .TextMatrix(0, 4) = "Início"
    .TextMatrix(0, 5) = "Situação"
    .TextMatrix(0, 6) = "Enviadas"
    .TextMatrix(0, 7) = "Erros"
    .TextMatrix(0, 8) = "Aguardando"
    .TextMatrix(0, 9) = "Puladas"
    For c = 0 To 9
        .ColAlignment(c) = flexAlignLeftCenter
    Next
    For c = 6 To 9
        .ColAlignment(c) = flexAlignCenterCenter
    Next
End With
On Error GoTo Falha
Set rs = dbData.OpenRecordset("SELECT ca.Id, ca.Nome, ca.TipoMidia, ca.CriadaEm, ca.AgendadaPara, ca.Situacao, " & _
    "SUM(CASE WHEN d.Situacao = " & CAMP_DEST_ENVIADA & " THEN 1 ELSE 0 END) AS ENV, " & _
    "SUM(CASE WHEN d.Situacao = " & CAMP_DEST_ERRO & " THEN 1 ELSE 0 END) AS ERR, " & _
    "SUM(CASE WHEN d.Situacao = " & CAMP_DEST_AGUARDANDO & " THEN 1 ELSE 0 END) AS AGU, " & _
    "SUM(CASE WHEN d.Situacao = " & CAMP_DEST_PULADA & " THEN 1 ELSE 0 END) AS PUL " & _
    "FROM Campanha ca LEFT JOIN CampanhaDestinatario d ON d.CodCampanha = ca.Id " & _
    "GROUP BY ca.Id, ca.Nome, ca.TipoMidia, ca.CriadaEm, ca.AgendadaPara, ca.Situacao ORDER BY ca.Id DESC")
Do While Not rs.EOF
    grdCamp.Rows = grdCamp.Rows + 1
    r = grdCamp.Rows - 1
    grdCamp.RowData(r) = rs("Id")
    grdCamp.TextMatrix(r, 0) = rs("Id")
    grdCamp.TextMatrix(r, 1) = "" & rs("Nome")
    grdCamp.TextMatrix(r, 2) = Camp_TipoTexto(Val("" & rs("TipoMidia")))
    grdCamp.TextMatrix(r, 3) = Format$(rs("CriadaEm"), "dd/mm/yyyy hh:nn")
    grdCamp.TextMatrix(r, 4) = IIf(IsNull(rs("AgendadaPara")), "Imediato", Format$(rs("AgendadaPara"), "dd/mm/yyyy hh:nn"))
    grdCamp.TextMatrix(r, 5) = Camp_SituacaoTexto(Val("" & rs("Situacao")))
    grdCamp.TextMatrix(r, 6) = Val("" & rs("ENV"))
    grdCamp.TextMatrix(r, 7) = Val("" & rs("ERR"))
    grdCamp.TextMatrix(r, 8) = Val("" & rs("AGU"))
    grdCamp.TextMatrix(r, 9) = Val("" & rs("PUL"))
    If rs("Id") = lSel Then nSel = r
    rs.MoveNext
Loop
rs.Close
Continua:
On Error Resume Next
AjustarColunas grdCamp, ",1,"
If grdCamp.Rows > 1 Then
    If nSel = 0 Then nSel = 1
    grdCamp.Row = nSel
    grdCamp.Col = 1
End If
grdCamp.Redraw = True
mSemEvento = False
Err.Clear
CarregarDestinatarios
Me.MousePointer = vbDefault
Exit Sub
Falha:
Resume Continua
End Sub

Private Sub CarregarDestinatarios()
Dim rs As ADODB.Recordset, lId As Long, r As Long, c As Integer

lId = CampanhaAtual()
With grdDest
    .Redraw = False
    .Rows = 1
    .Cols = 5
    .TextMatrix(0, 0) = "Cliente"
    .TextMatrix(0, 1) = "Celular"
    .TextMatrix(0, 2) = "Situação"
    .TextMatrix(0, 3) = "Data/hora"
    .TextMatrix(0, 4) = "Detalhe"
    For c = 0 To 4
        .ColAlignment(c) = flexAlignLeftCenter
    Next
End With
AtualizarBotoes
If lId = 0 Then
    lblDest.Caption = "Clientes da campanha:"
    grdDest.Redraw = True
    Exit Sub
End If
lblDest.Caption = "Clientes da campanha nº " & lId & ":"
On Error GoTo Falha
Set rs = dbData.OpenRecordset("SELECT c.Nome, d.Celular, d.Situacao, d.DataHora, d.Erro FROM CampanhaDestinatario d " & _
    "LEFT JOIN cliente c ON c.CODIGO = d.CodCliente WHERE d.CodCampanha = " & lId & " ORDER BY d.Situacao, c.Nome")
Do While Not rs.EOF
    grdDest.Rows = grdDest.Rows + 1
    r = grdDest.Rows - 1
    grdDest.TextMatrix(r, 0) = "" & rs("Nome")
    grdDest.TextMatrix(r, 1) = "" & rs("Celular")
    Select Case Val("" & rs("Situacao"))
        Case CAMP_DEST_ENVIADA: grdDest.TextMatrix(r, 2) = "Enviada"
        Case CAMP_DEST_ERRO: grdDest.TextMatrix(r, 2) = "Erro"
        Case CAMP_DEST_PULADA: grdDest.TextMatrix(r, 2) = "Pulada"
        Case Else: grdDest.TextMatrix(r, 2) = "Aguardando"
    End Select
    grdDest.TextMatrix(r, 3) = IIf(IsNull(rs("DataHora")), "", Format$(rs("DataHora"), "dd/mm/yyyy hh:nn"))
    grdDest.TextMatrix(r, 4) = Replace("" & rs("Erro"), vbCrLf, " ")
    rs.MoveNext
Loop
rs.Close
Continua:
On Error Resume Next
AjustarColunas grdDest, ",0,4,"
grdDest.Redraw = True
Err.Clear
Exit Sub
Falha:
Resume Continua
End Sub

Private Sub AtualizarBotoes()
Dim s As String

If CampanhaAtual() > 0 Then s = grdCamp.TextMatrix(grdCamp.Row, 5)
cmdPausar.Enabled = (s = Camp_SituacaoTexto(CAMP_NA_FILA))
cmdRetomar.Enabled = (s = Camp_SituacaoTexto(CAMP_PAUSADA))
cmdCancelar.Enabled = (s = Camp_SituacaoTexto(CAMP_NA_FILA) Or s = Camp_SituacaoTexto(CAMP_PAUSADA))
End Sub

Private Sub grdCamp_RowColChange()
If Not mSemEvento Then CarregarDestinatarios
End Sub

Private Sub cmdPausar_Click()
Mudar CAMP_PAUSADA, ""
End Sub

Private Sub cmdRetomar_Click()
Mudar CAMP_NA_FILA, ""
End Sub

Private Sub cmdCancelar_Click()
Mudar CAMP_CANCELADA, "Cancelar a campanha nº " & CampanhaAtual() & "?" & vbCrLf & vbCrLf & _
                      "Quem ainda não recebeu não vai receber. Não dá para desfazer."
End Sub

Private Sub Mudar(ByVal nSit As Integer, ByVal sConfirma As String)
Dim lId As Long

lId = CampanhaAtual()
If lId = 0 Then Exit Sub
If Len(sConfirma) > 0 Then
    If MsgBox(sConfirma, vbQuestion + vbYesNo + vbDefaultButton2, TITULO) <> vbYes Then Exit Sub
End If
If Not Camp_MudarSituacao(lId, nSit) Then
    MsgBox "Não foi possível alterar a campanha.", vbExclamation, TITULO
    Exit Sub
End If
Carregar lId
End Sub

Private Sub AjustarColunas(grd As MSFlexGrid, ByVal sElasticas As String)
Dim c As Long, r As Long, w() As Long, wt As Long, nTotal As Long, nDisp As Long, nElast As Long, nDif As Long

On Error Resume Next
Me.FontName = grd.Font.Name
Me.FontSize = grd.Font.Size
Me.FontBold = False
ReDim w(0 To grd.Cols - 1)
For c = 0 To grd.Cols - 1
    For r = 0 To grd.Rows - 1
        wt = Me.TextWidth(grd.TextMatrix(r, c))
        If wt > w(c) Then w(c) = wt
    Next
    w(c) = w(c) + 180
    nTotal = nTotal + w(c)
    If InStr(sElasticas, "," & c & ",") > 0 Then nElast = nElast + w(c)
Next
nDisp = grd.Width - 90
If grd.Rows * grd.RowHeight(0) > grd.Height - 60 Then nDisp = nDisp - 270
If nElast > 0 And nTotal <> nDisp Then
    nDif = nDisp - nTotal
    For c = 0 To grd.Cols - 1
        If InStr(sElasticas, "," & c & ",") > 0 Then
            w(c) = w(c) + CLng(CDbl(nDif) * w(c) / nElast)
            If w(c) < 600 Then w(c) = 600
        End If
    Next
End If
For c = 0 To grd.Cols - 1
    grd.ColWidth(c) = w(c)
Next
Err.Clear
End Sub
