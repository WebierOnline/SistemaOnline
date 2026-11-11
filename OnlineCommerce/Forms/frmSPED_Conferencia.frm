VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSPED_Conferencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SPED Fiscal - Conferência dos cadastros"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdCorrigirAmarelos 
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   7800
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "CORRIGIR OS AMARELOS (automaticamente)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmSPED_Conferencia.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdLista 
      Height          =   5595
      Left            =   60
      TabIndex        =   6
      Top             =   2160
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   9869
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdCorrigir 
      Height          =   615
      Left            =   3660
      TabIndex        =   8
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "CORRIGIR OS VERDES (automaticamente)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSPED_Conferencia.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdMarcar 
      Height          =   615
      Left            =   60
      TabIndex        =   9
      Top             =   7800
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "MARCAR VERDES"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSPED_Conferencia.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdMarcarAmarelos 
      Height          =   615
      Left            =   1860
      TabIndex        =   10
      Top             =   7800
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "MARCAR AMARELOS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSPED_Conferencia.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCorrigirManual 
      Height          =   615
      Left            =   7920
      TabIndex        =   11
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "CORRIGIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSPED_Conferencia.frx":0070
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
      Height          =   615
      Left            =   13860
      TabIndex        =   12
      Top             =   7800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "CANCELAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSPED_Conferencia.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdGerar 
      Height          =   615
      Left            =   9300
      TabIndex        =   13
      Top             =   7800
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "GERAR ARQUIVO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSPED_Conferencia.frx":00A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdManter 
      Height          =   615
      Left            =   10860
      TabIndex        =   14
      Top             =   7800
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "MANTER O ATUAL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSPED_Conferencia.frx":00C4
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
      Height          =   615
      Left            =   12420
      TabIndex        =   15
      Top             =   7800
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "SALVAR LISTA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSPED_Conferencia.frx":00E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblRodape 
      Caption         =   "Somente após a correção de todos o arquivo poderá ser gerado."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   1860
      Width           =   14175
   End
   Begin VB.Label lblVermelhos 
      BackColor       =   &H00DCDCFA&
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1560
      Width           =   14955
   End
   Begin VB.Label lblAmarelos 
      BackColor       =   &H00C8F5FF&
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   14955
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVerdes 
      BackColor       =   &H00DCF5DC&
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   14955
   End
   Begin VB.Label lblTotal 
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   14955
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "RELATÓRIO DE ERROS ENCONTRADOS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   14955
   End
End
Attribute VB_Name = "frmSPED_Conferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Conferência dos cadastros antes de gerar o SPED (chamada pelo frmSPED_Fiscal): lista o que o SpedGerador achou
' (modResumoFiscal.SPED_Conferir).
'   VERDE    = há fonte confiável (XML da nota de entrada ou tabela de cidades): CORRIGIR OS VERDES, vem marcado;
'   AMARELO  = válido, mas diferente da nota de entrada/tabela de conversão: CORRIGIR OS AMARELOS (marcando antes),
'              CORRIGIR (abre o cadastro) ou MANTER O ATUAL (não volta enquanto os valores não mudarem);
'   VERMELHO = sem fonte para corrigir: CORRIGIR abre o cadastro da linha.
' O arquivo só é gerado depois que não sobra nenhum. As correções guardam o valor anterior em SPED_CorrecaoCadastro.
Private Const COL_MARCA As Integer = 0
Private Const COL_SITUACAO As Integer = 1
Private Const COL_TIPO As Integer = 2
Private Const COL_CODIGO As Integer = 3
Private Const COL_NOME As Integer = 4
Private Const COL_CAMPO As Integer = 5
Private Const COL_ATUAL As Integer = 6
Private Const COL_NOVO As Integer = 7
Private Const COL_PROBLEMA As Integer = 8
Private Const COL_FONTE As Integer = 9
Private Const COL_ONDE As Integer = 10
Private Const COL_CHAVE As Integer = 11
Private Const COL_SIT As Integer = 12

Dim sParametros As String
Dim bGerar As Boolean
Dim nVerdes As Long, nAmarelos As Long, nVermelhos As Long

' Mostra a conferência (modal). True = gerar o arquivo; False = cancelado.
Public Function Mostrar(ByVal sParam As String) As Boolean
    sParametros = sParam
    bGerar = False
    MontarGrade
    Carregar
    Me.Show vbModal
    Mostrar = bGerar
    Unload Me
End Function

Private Sub MontarGrade()
    With grdLista
        'Rows/FixedRows ficam no padrão (2 e 1): FixedRows = 1 com Rows = 1 dá erro 30016; o Carregar limpa as linhas
        .Cols = 13
        .FixedCols = 0
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColWidth(COL_MARCA) = 450: .TextMatrix(0, COL_MARCA) = "Corr."
        .ColWidth(COL_SITUACAO) = 1300: .TextMatrix(0, COL_SITUACAO) = "Situação"
        .ColWidth(COL_TIPO) = 1100: .TextMatrix(0, COL_TIPO) = "Tipo"
        .ColWidth(COL_CODIGO) = 750: .TextMatrix(0, COL_CODIGO) = "Código"
        .ColWidth(COL_NOME) = 2600: .TextMatrix(0, COL_NOME) = "Nome"
        .ColWidth(COL_CAMPO) = 850: .TextMatrix(0, COL_CAMPO) = "Campo"
        .ColWidth(COL_ATUAL) = 1000: .TextMatrix(0, COL_ATUAL) = "Atual"
        .ColWidth(COL_NOVO) = 1000: .TextMatrix(0, COL_NOVO) = "Correção"
        .ColWidth(COL_PROBLEMA) = 2600: .TextMatrix(0, COL_PROBLEMA) = "Problema"
        .ColWidth(COL_FONTE) = 1900: .TextMatrix(0, COL_FONTE) = "Fonte da correção"
        .ColWidth(COL_ONDE) = 2200: .TextMatrix(0, COL_ONDE) = "Onde corrigir"
        .ColWidth(COL_CHAVE) = 0
        .ColWidth(COL_SIT) = 0
    End With
End Sub

Private Sub Carregar()
    Dim nArq As Integer
    Dim sLinha As String
    Dim v As Variant
    Dim r As Long, c As Integer
    Dim nCor As Long

    nVerdes = 0: nAmarelos = 0: nVermelhos = 0
    grdLista.rows = 1
    If Dir$(SPED_ArquivoConferencia) <> "" Then
        nArq = FreeFile
        Open SPED_ArquivoConferencia For Input As #nArq
        Do While Not EOF(nArq)
            Line Input #nArq, sLinha
            v = Split(sLinha, vbTab)
            If UBound(v) >= 10 Then
                grdLista.AddItem ""
                r = grdLista.rows - 1
                grdLista.TextMatrix(r, COL_SIT) = v(1)
                Select Case v(1)
                    Case "CORRIGIR"
                        grdLista.TextMatrix(r, COL_SITUACAO) = "Verde"
                        grdLista.TextMatrix(r, COL_MARCA) = "X"
                        nVerdes = nVerdes + 1
                        nCor = RGB(220, 245, 220)
                    Case "DIVERGENTE"
                        grdLista.TextMatrix(r, COL_SITUACAO) = "Amarelo"
                        nAmarelos = nAmarelos + 1
                        nCor = RGB(255, 245, 200)
                    Case Else
                        grdLista.TextMatrix(r, COL_SITUACAO) = "Vermelho"
                        nVermelhos = nVermelhos + 1
                        nCor = RGB(250, 220, 220)
                End Select
                grdLista.TextMatrix(r, COL_CHAVE) = v(0)
                grdLista.TextMatrix(r, COL_TIPO) = v(2)
                grdLista.TextMatrix(r, COL_CODIGO) = v(3)
                grdLista.TextMatrix(r, COL_NOME) = v(4)
                grdLista.TextMatrix(r, COL_CAMPO) = v(5)
                grdLista.TextMatrix(r, COL_ATUAL) = v(6)
                grdLista.TextMatrix(r, COL_NOVO) = v(7)
                grdLista.TextMatrix(r, COL_PROBLEMA) = v(8)
                grdLista.TextMatrix(r, COL_FONTE) = v(9)
                grdLista.TextMatrix(r, COL_ONDE) = v(10)
                grdLista.Row = r
                For c = 0 To COL_ONDE
                    grdLista.Col = c
                    grdLista.CellBackColor = nCor
                Next
            End If
        Loop
        Close #nArq
    End If

    If nVerdes + nAmarelos + nVermelhos = 0 Then
        lblTotal.Caption = "Nenhum erro encontrado nos cadastros que vão no arquivo. O arquivo pode ser gerado."
    Else
        lblTotal.Caption = "Existem " & (nVerdes + nAmarelos + nVermelhos) & " cadastro(s) que precisam ser avaliados/corrigidos antes de gerar o arquivo."
    End If
    lblVerdes.Caption = nVerdes & " ERROS VERDES - Podem ser corrigidos automaticamente clicando no botão 'CORRIGIR OS VERDES (automaticamente)'."
    lblAmarelos.Caption = nAmarelos & " ERROS AMARELOS - Podem ser corrigidos automaticamente clicando 'MARCAR AMARELOS' e depois no botão 'CORRIGIR OS AMARELOS (automaticamente)', " & _
                          "manualmente no botão 'CORRIGIR' ou, se o cadastro estiver certo, mantidos no botão 'MANTER O ATUAL'."
    lblVermelhos.Caption = nVermelhos & " ERROS VERMELHOS - Precisam ser corrigidos manualmente pelo usuário (selecione a linha e clique em 'CORRIGIR')."
    lblRodape.Caption = IIf(nVerdes + nAmarelos + nVermelhos = 0, "", "Somente após a correção de todos o arquivo poderá ser gerado.")
    AtualizarBotoes
End Sub

Private Function Contar(ByVal sSituacao As String, ByVal bMarcados As Boolean) As Long
    Dim r As Long
    For r = 1 To grdLista.rows - 1
        If grdLista.TextMatrix(r, COL_SIT) = sSituacao Then
            If Not bMarcados Or grdLista.TextMatrix(r, COL_MARCA) = "X" Then Contar = Contar + 1
        End If
    Next
End Function

Private Function SituacaoSelecionada() As String
    If grdLista.Row >= 1 And grdLista.rows > 1 Then SituacaoSelecionada = grdLista.TextMatrix(grdLista.Row, COL_SIT)
End Function

Private Sub AtualizarBotoes()
    Dim nVerdesMarc As Long, nAmarMarc As Long, sSel As String
    nVerdesMarc = Contar("CORRIGIR", True)
    nAmarMarc = Contar("DIVERGENTE", True)
    sSel = SituacaoSelecionada

    cmdMarcar.Enabled = (nVerdes > 0)
    cmdMarcar.Caption = IIf(nVerdes > 0 And nVerdesMarc = nVerdes, "DESMARCAR VERDES", "MARCAR VERDES")
    cmdCorrigir.Enabled = (nVerdesMarc > 0)
    cmdMarcarAmarelos.Enabled = (nAmarelos > 0)
    cmdMarcarAmarelos.Caption = IIf(nAmarelos > 0 And nAmarMarc = nAmarelos, "DESMARCAR AMARELOS", "MARCAR AMARELOS")
    cmdCorrigirAmarelos.Enabled = (nAmarMarc > 0)
    cmdCorrigirManual.Enabled = (sSel = "DIVERGENTE" Or sSel = "MANUAL")
    cmdManter.Enabled = (sSel = "DIVERGENTE")
    cmdSalvar.Enabled = (nVerdes + nAmarelos + nVermelhos > 0)
    cmdGerar.Enabled = (nVerdes + nAmarelos + nVermelhos = 0)
End Sub

Private Sub grdLista_Click()
    Dim r As Long
    r = grdLista.MouseRow
    If r >= 1 And grdLista.MouseCol = COL_MARCA Then
        If grdLista.TextMatrix(r, COL_SIT) = "CORRIGIR" Or (grdLista.TextMatrix(r, COL_SIT) = "DIVERGENTE" And grdLista.TextMatrix(r, COL_NOVO) <> "") Then
            grdLista.TextMatrix(r, COL_MARCA) = IIf(grdLista.TextMatrix(r, COL_MARCA) = "X", "", "X")
        End If
    End If
    AtualizarBotoes
End Sub

Private Sub grdLista_RowColChange()
    AtualizarBotoes
End Sub

Private Sub Marcar(ByVal sSituacao As String, ByVal sMarca As String)
    Dim r As Long
    For r = 1 To grdLista.rows - 1
        If grdLista.TextMatrix(r, COL_SIT) = sSituacao And grdLista.TextMatrix(r, COL_NOVO) <> "" Then grdLista.TextMatrix(r, COL_MARCA) = sMarca
    Next
    AtualizarBotoes
End Sub

Private Sub cmdMarcar_Click()
    Marcar "CORRIGIR", IIf(cmdMarcar.Caption = "DESMARCAR VERDES", "", "X")
End Sub

Private Sub cmdMarcarAmarelos_Click()
    Marcar "DIVERGENTE", IIf(cmdMarcarAmarelos.Caption = "DESMARCAR AMARELOS", "", "X")
End Sub

' Chaves das linhas marcadas de uma situação, separadas por vírgula.
Private Function ChavesMarcadas(ByVal sSituacao As String) As String
    Dim r As Long
    For r = 1 To grdLista.rows - 1
        If grdLista.TextMatrix(r, COL_SIT) = sSituacao And grdLista.TextMatrix(r, COL_MARCA) = "X" Then
            ChavesMarcadas = ChavesMarcadas & IIf(ChavesMarcadas <> "", ",", "") & grdLista.TextMatrix(r, COL_CHAVE)
        End If
    Next
End Function

' Corrige os marcados da situação (verdes ou amarelos) com os valores da coluna Correção.
Private Sub CorrigirMarcados(ByVal sSituacao As String, ByVal sCor As String)
    Dim sChaves As String, sRes As String
    Dim n As Long, nFeitos As Long

    sChaves = ChavesMarcadas(sSituacao)
    n = Contar(sSituacao, True)
    If n = 0 Then
        MsgBox "Marque (X) os " & sCor & " a corrigir.", vbInformation
        Exit Sub
    End If
    If MsgBox("Corrigir " & n & " cadastro(s) " & sCor & " com os valores da coluna Correção?" & vbCrLf & vbCrLf & _
              "O valor anterior fica guardado no histórico (tabela SPED_CorrecaoCadastro).", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass
    DoEvents
    sRes = SPED_Corrigir(sParametros, sChaves, nFeitos)
    Screen.MousePointer = vbDefault
    If sRes <> "OK" Then
        MsgBox "Não foi possível corrigir: " & Mid$(sRes, IIf(Left$(sRes, 6) = "ERRO: ", 7, 1)), vbExclamation
        Exit Sub
    End If
    Reconferir
    MsgBox nFeitos & " cadastro(s) corrigido(s).", vbInformation
End Sub

Private Sub cmdCorrigir_Click()
    CorrigirMarcados "CORRIGIR", "verdes"
End Sub

Private Sub cmdCorrigirAmarelos_Click()
    CorrigirMarcados "DIVERGENTE", "amarelos"
End Sub

' Confirma o amarelo selecionado como está (o cadastro está certo): não aparece mais enquanto os valores não mudarem.
Private Sub cmdManter_Click()
    Dim r As Long, sRes As String, n As Long
    r = grdLista.Row
    If r < 1 Or SituacaoSelecionada <> "DIVERGENTE" Then Exit Sub
    If MsgBox("Manter o cadastro como está?" & vbCrLf & vbCrLf & _
              grdLista.TextMatrix(r, COL_TIPO) & " " & grdLista.TextMatrix(r, COL_CODIGO) & " - " & grdLista.TextMatrix(r, COL_NOME) & vbCrLf & _
              grdLista.TextMatrix(r, COL_CAMPO) & ": fica " & grdLista.TextMatrix(r, COL_ATUAL) & " (sugestão era " & grdLista.TextMatrix(r, COL_NOVO) & ")", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    DoEvents
    sRes = SPED_Manter(sParametros, grdLista.TextMatrix(r, COL_CHAVE), n)
    Screen.MousePointer = vbDefault
    If sRes <> "OK" Then
        MsgBox "Não foi possível registrar: " & Mid$(sRes, IIf(Left$(sRes, 6) = "ERRO: ", 7, 1)), vbExclamation
        Exit Sub
    End If
    Reconferir
End Sub

' Abre o cadastro da linha selecionada (amarela ou vermelha); ao fechar, a conferência roda de novo.
Private Sub cmdCorrigirManual_Click()
    Dim r As Long, sGrupo As String, lCodigo As Long
    r = grdLista.Row
    If r < 1 Then Exit Sub
    If SituacaoSelecionada <> "DIVERGENTE" And SituacaoSelecionada <> "MANUAL" Then Exit Sub
    sGrupo = grdLista.TextMatrix(r, COL_TIPO)
    lCodigo = Val(grdLista.TextMatrix(r, COL_CODIGO))

    On Error GoTo ErrL
    Select Case sGrupo
        Case "Fornecedor"
            Load Fornecedor_Cadastro
            Fornecedor_Cadastro.EditarFornecedor lCodigo
            Fornecedor_Cadastro.Show vbModal, Me
            Unload Fornecedor_Cadastro
        Case "Produto"
            Load Produtos_Cadastro
            Produtos_Cadastro.SSTab1.Tab = 0
            Produtos_Cadastro.EditarProduto lCodigo
            Produtos_Cadastro.Show vbModal, Me
            Unload Produtos_Cadastro
        Case "Cliente NF-e"
            Load Clientes_Cadastro
            Clientes_Cadastro.EditarCliente lCodigo
            Clientes_Cadastro.Show vbModal, Me
            Unload Clientes_Cadastro
        Case "Empresa"
            Empresa_Cadastro.Show vbModal, Me
            Unload Empresa_Cadastro
        Case "Contador"
            Contador_Cadastro.Show vbModal, Me
            Unload Contador_Cadastro
        Case Else
            MsgBox "Corrija em: " & grdLista.TextMatrix(r, COL_ONDE), vbInformation
            Exit Sub
    End Select
    Reconferir
    Exit Sub
ErrL:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível abrir o cadastro: " & Err.Description & vbCrLf & "Corrija em: " & grdLista.TextMatrix(r, COL_ONDE), vbExclamation
End Sub

' Roda a conferência de novo e recarrega a lista (depois de corrigir ou manter).
Private Sub Reconferir()
    Dim sRes As String, nC As Long, nM As Long, nD As Long
    Screen.MousePointer = vbHourglass
    DoEvents
    sRes = SPED_Conferir(sParametros, nC, nM, nD)
    Screen.MousePointer = vbDefault
    If sRes <> "OK" Then
        MsgBox "Não foi possível conferir de novo: " & Mid$(sRes, IIf(Left$(sRes, 6) = "ERRO: ", 7, 1)), vbExclamation
        Exit Sub
    End If
    Carregar
End Sub

' Lista em texto (para corrigir os cadastros com calma ou mandar para alguém), aberta no Bloco de Notas.
Private Sub cmdSalvar_Click()
    Dim sPasta As String, sArq As String
    Dim nArq As Integer, r As Long
    Dim oFso As Object

    On Error GoTo ErrL
    sPasta = SPED_Pasta
    Set oFso = CreateObject("Scripting.FileSystemObject")
    If Not oFso.FolderExists(oFso.GetParentFolderName(sPasta)) Then oFso.CreateFolder oFso.GetParentFolderName(sPasta)
    If Not oFso.FolderExists(sPasta) Then oFso.CreateFolder sPasta
    sArq = sPasta & "\Conferencia_cadastros_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"

    nArq = FreeFile
    Open sArq For Output As #nArq
    Print #nArq, "RELATÓRIO DE ERROS ENCONTRADOS - SPED Fiscal - " & Format(Now, "dd/mm/yyyy hh:nn")
    Print #nArq, lblTotal.Caption
    Print #nArq, lblVerdes.Caption
    Print #nArq, lblAmarelos.Caption
    Print #nArq, lblVermelhos.Caption
    Print #nArq, String(100, "-")
    For r = 1 To grdLista.rows - 1
        Print #nArq, UCase(grdLista.TextMatrix(r, COL_SITUACAO)) & " | " & grdLista.TextMatrix(r, COL_TIPO) & " " & grdLista.TextMatrix(r, COL_CODIGO) & _
                     " - " & grdLista.TextMatrix(r, COL_NOME)
        Print #nArq, "   " & grdLista.TextMatrix(r, COL_CAMPO) & ": " & grdLista.TextMatrix(r, COL_PROBLEMA) & _
                     " (atual: " & grdLista.TextMatrix(r, COL_ATUAL) & IIf(grdLista.TextMatrix(r, COL_NOVO) <> "", "; correção: " & grdLista.TextMatrix(r, COL_NOVO), "") & ")"
        If grdLista.TextMatrix(r, COL_FONTE) <> "" Then Print #nArq, "   Fonte: " & grdLista.TextMatrix(r, COL_FONTE)
        Print #nArq, "   Onde corrigir: " & grdLista.TextMatrix(r, COL_ONDE)
    Next
    Close #nArq
    Shell "notepad.exe " & Chr(34) & sArq & Chr(34), vbNormalFocus
    Exit Sub
ErrL:
    MsgBox "Não foi possível salvar a lista: " & Err.Description, vbExclamation
End Sub

Private Sub cmdGerar_Click()
    If nVerdes + nAmarelos + nVermelhos > 0 Then
        MsgBox "Somente após a correção de todos o arquivo poderá ser gerado.", vbExclamation
        Exit Sub
    End If
    bGerar = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    bGerar = False
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'fechar pelo X = cancelar (a função Mostrar descarrega o form)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        bGerar = False
        Me.Hide
    End If
End Sub
