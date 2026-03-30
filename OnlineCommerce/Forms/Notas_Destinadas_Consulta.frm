VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Notas_Destinadas_Consulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOTAS DESTINADAS - CONSULTA"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15180
   Icon            =   "Notas_Destinadas_Consulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   15180
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6120
      Picture         =   "Notas_Destinadas_Consulta.frx":1D82
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   23
      Top             =   4140
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   675
      Left            =   60
      TabIndex        =   8
      Top             =   420
      Width           =   15015
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   540
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendario1 
         Height          =   315
         Left            =   4440
         TabIndex        =   12
         Tag             =   "Calendario"
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Notas_Destinadas_Consulta.frx":2DBA
         PICN            =   "Notas_Destinadas_Consulta.frx":2DD6
         PICH            =   "Notas_Destinadas_Consulta.frx":5129
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendario2 
         Height          =   315
         Left            =   6060
         TabIndex        =   13
         Tag             =   "Calendario"
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Notas_Destinadas_Consulta.frx":747C
         PICN            =   "Notas_Destinadas_Consulta.frx":7498
         PICH            =   "Notas_Destinadas_Consulta.frx":97EB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   3420
         TabIndex        =   14
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFim 
         Height          =   315
         Left            =   5100
         TabIndex        =   15
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdLimparSelecao 
         Height          =   315
         Left            =   11220
         TabIndex        =   17
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Limpa Seleção"
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
         MICON           =   "Notas_Destinadas_Consulta.frx":BB3E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAtualizarGrid 
         Height          =   315
         Left            =   9600
         TabIndex        =   18
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Atualizar Grid"
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
         MICON           =   "Notas_Destinadas_Consulta.frx":BB5A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton chkMarcarTodas 
         Height          =   315
         Left            =   12840
         TabIndex        =   24
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "MARCAR TODAS"
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
         MICON           =   "Notas_Destinadas_Consulta.frx":BB76
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   4860
         TabIndex        =   16
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.TextBox txtUltimoNSU 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1020
      TabIndex        =   5
      Top             =   60
      Width           =   2715
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   9150
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22437
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "07:56"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChamaleonBtn.chameleonButton cmdDANFe 
      Height          =   315
      Left            =   11880
      TabIndex        =   1
      Top             =   8820
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "DANFe"
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
      MICON           =   "Notas_Destinadas_Consulta.frx":BB92
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdImportarNFe 
      Height          =   315
      Left            =   10260
      TabIndex        =   2
      Top             =   8820
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Importar NFe"
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
      MICON           =   "Notas_Destinadas_Consulta.frx":BBAE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   13500
      TabIndex        =   7
      Top             =   60
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Consultar"
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
      MICON           =   "Notas_Destinadas_Consulta.frx":BBCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEnviarManifestacao 
      Height          =   315
      Left            =   8520
      TabIndex        =   19
      Top             =   8820
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Enviar Manifestação"
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
      MICON           =   "Notas_Destinadas_Consulta.frx":BBE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCopiarXML 
      Height          =   315
      Left            =   13500
      TabIndex        =   21
      Top             =   8820
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Copiar Chave"
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
      MICON           =   "Notas_Destinadas_Consulta.frx":BC02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7635
      Left            =   60
      TabIndex        =   22
      Top             =   1140
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13467
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdEnviarManifestacao2 
      Height          =   315
      Left            =   8520
      TabIndex        =   26
      Top             =   8820
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Enviar Manifestações"
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
      MICON           =   "Notas_Destinadas_Consulta.frx":BC1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblRegistroSel 
      AutoSize        =   -1  'True
      Caption         =   "0000 Registro(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2100
      TabIndex        =   25
      Top             =   8820
      Width           =   1410
   End
   Begin VB.Image ImgMarcada 
      Height          =   195
      Left            =   4800
      Picture         =   "Notas_Destinadas_Consulta.frx":BC3A
      Top             =   8880
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDesmarcada 
      Height          =   195
      Left            =   5100
      Picture         =   "Notas_Destinadas_Consulta.frx":E039
      Top             =   8880
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblQuantRegistro 
      AutoSize        =   -1  'True
      Caption         =   "0000 Registro(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   20
      Top             =   8820
      Width           =   1410
   End
   Begin VB.Label lblUltimaConsulta 
      AutoSize        =   -1  'True
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5100
      TabIndex        =   6
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Última Consulta:"
      Height          =   195
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Último NSU:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "Notas_Destinadas_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim r As ADODB.Recordset
Dim printSQL As String
Private moCombo As cComboHelper

Enum TipoOP2                 'usado para o check das parcelas
   MarcarTodos = 1
   DesmarcarTodos = 2
   contar = 3
End Enum

Dim OP As TipoOP2           'usado para o check das parcelas
Dim var_Contador As Integer 'usado para o check das parcelas

'Dim i As Integer
Private Sub MostrarDados()

txtUltimoNSU.Text = SQLExecutaRetorno("SELECT UltimoNSU FROM empresa", "UltimoNSU", "")

'If cboDesc.Text = "" Then
    sSQL = "SELECT SituacaoNF, TipoRetornoWS, TipoEventoID, TipoEvento, TipoNotaWS, Data, CNPJFornecedor, RazaoSocial, ChaveAcesso, DataEmissao, ValorNota, SituacaoNota, ImportadoXML, Enviar, Enviada, CodigoConsulta, IdNFConsulta, IdFilial, TipoEventoCodigo, EventoProtocolo, EventoDataHora " & _
           "FROM NotasDestinadas " & _
           "WHERE 1=0  " & _
           "ORDER BY IdNFConsulta DESC"
'Else
'    sSQL = "SELECT IdInventario, Seq as vSeq, IDProduto as vCodProd, NomeProduto as vDesc, EAN as vEAN, NCM as vNCM, SaldoCalculado, VlrUnitInvent as vVlrUnit, TotalInvent as vTotal, MetaCalculado  as vMeta, SaldoLancado, TotalFisico " & _
           "FROM InventarioGerado " & _
           "WHERE (NomeProduto = '" & cboDesc.Text & "')  " & _
           "ORDER BY Seq"
'End If
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub

Private Sub cboTipo_GotFocus()
cboTipo.Clear

sSQL = "SELECT DISTINCT SituacaoNF FROM NotasDestinadas ORDER BY SituacaoNF;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboTipo.AddItem ValidateNull(r("SituacaoNF"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboTipo
End Sub

Private Sub chkMarcarTodas_Click()
If chkMarcarTodas.Caption = "MARCAR TODAS" Then
   OP = MarcarTodos
   AcaoGrid
   chkMarcarTodas.Caption = "DESMARCAR TODAS"
Else
   OP = DesmarcarTodos
   AcaoGrid
   chkMarcarTodas.Caption = "MARCAR TODAS"
End If

OP = contar
AcaoGrid
End Sub


Private Sub cmdAtualizarGrid_Click()
Dim vSituacao As String
Dim vDataConf As String

If Not Vazio(cboTipo.Text) And cboTipo.Text <> "Todas" Then
    vSituacao = "SituacaoNF = '" & cboTipo.Text & "'"
Else
    vSituacao = "1 = 1"
End If

If IsDate(mskInicio.Text) And IsDate(mskFim.Text) Then
    vDataConf = "Periodo"
ElseIf IsDate(mskInicio.Text) And IsDate(mskFim.Text) = False Then
    vDataConf = "Inicio"
ElseIf IsDate(mskInicio.Text) = False And IsDate(mskFim.Text) = False Then
    vDataConf = "Nada"
End If

If vDataConf = "Periodo" Then
    sSQL = "SELECT SituacaoNF, TipoRetornoWS, TipoEventoID, TipoEvento, TipoNotaWS, Data, CNPJFornecedor, RazaoSocial, ChaveAcesso, DataEmissao, ValorNota, SituacaoNota, ImportadoXML, Enviar, Enviada, CodigoConsulta, IdNFConsulta, IdFilial, TipoEventoCodigo, EventoProtocolo, EventoDataHora " & _
           "FROM NotasDestinadas " & _
           "WHERE " & vSituacao & " AND (DataEmissao >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
           "ORDER BY DataEmissao DESC"
ElseIf vDataConf = "Inicio" Then
    sSQL = "SELECT SituacaoNF, TipoRetornoWS, TipoEventoID, TipoEvento, TipoNotaWS, Data, CNPJFornecedor, RazaoSocial, ChaveAcesso, DataEmissao, ValorNota, SituacaoNota, ImportadoXML, Enviar, Enviada, CodigoConsulta, IdNFConsulta, IdFilial, TipoEventoCodigo, EventoProtocolo, EventoDataHora " & _
           "FROM NotasDestinadas " & _
           "WHERE " & vSituacao & " AND (DataEmissao >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103))  " & _
           "ORDER BY DataEmissao DESC"
ElseIf vDataConf = "Nada" Then
    sSQL = "SELECT SituacaoNF, TipoRetornoWS, TipoEventoID, TipoEvento, TipoNotaWS, Data, CNPJFornecedor, RazaoSocial, ChaveAcesso, DataEmissao, ValorNota, SituacaoNota, ImportadoXML, Enviar, Enviada, CodigoConsulta, IdNFConsulta, IdFilial, TipoEventoCodigo, EventoProtocolo, EventoDataHora " & _
           "FROM NotasDestinadas " & _
           "WHERE " & vSituacao & " " & _
           "ORDER BY DataEmissao DESC"
End If
'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL)

lblQuantRegistro.Caption = r.RecordCount & " Registro(s)"

FormatarGrid r

OP = DesmarcarTodos
AcaoGrid

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub

Private Sub cmdCalendario1_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCalendario2_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskFim = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer

With Grid
   .Clear
   .Cols = 11
   .rows = 2
   
   .ColWidth(0) = 300
   .ColWidth(1) = 0
   .ColWidth(2) = 1600
   .ColWidth(3) = 1800
   .ColWidth(4) = 850
   .ColWidth(5) = 0
   .ColWidth(6) = 4000
   .ColWidth(7) = 4200
   .ColWidth(8) = 1000
   .ColWidth(9) = 870
   .ColWidth(10) = 50
   
   .TextMatrix(0, 1) = "SEL"
   .TextMatrix(0, 2) = "SITUAÇÃO"
   .TextMatrix(0, 3) = "EVENTO"
   .TextMatrix(0, 4) = "DATA"
   .TextMatrix(0, 5) = "CNPJ"
   .TextMatrix(0, 6) = "RAZÃO SOCIAL"
   .TextMatrix(0, 7) = "CHAVE DE ACESSO"
   .TextMatrix(0, 8) = "VALOR"
   .TextMatrix(0, 9) = "EMISSÃO"
   .TextMatrix(0, 10) = "ID"
   
   .Redraw = False
   i = 1
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         '.TextMatrix(.Rows - 1, 1) = rTabela("vSeq")
         .TextMatrix(.rows - 1, 2) = rTabela("SituacaoNF")
         .TextMatrix(.rows - 1, 3) = rTabela("TipoEvento")
         .TextMatrix(.rows - 1, 4) = Format(rTabela("Data"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 5) = rTabela("CNPJFornecedor")
         .TextMatrix(.rows - 1, 6) = rTabela("RazaoSocial")
         .TextMatrix(.rows - 1, 7) = rTabela("ChaveAcesso")
         .TextMatrix(.rows - 1, 8) = Format(rTabela("ValorNota"), ocMONEY)
         .TextMatrix(.rows - 1, 9) = Format(rTabela("DataEmissao"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 10) = rTabela("IdNFConsulta")
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If

   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   
   'MUDAR COR DE FONTE DA COLUNA
   'For i = 1 To .Rows - 1
   '   .Row = i
   '   .Col = 9
   '   .CellForeColor = &HC0&
   '   .CellFontBold = True
   'Next
   
   Grid.Col = 0
   
    For i = 1 To .rows - 1
       Grid.Row = i
       Set Grid.CellPicture = imgDesmarcada
       Grid.CellPictureAlignment = 4
    Next
   
   .rows = .rows - 1
   .Redraw = True
End With

'lblTotalEntrada.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
'lblTotalSaida.Caption = Format(SomaGrid(Grid, 8), ocMONEY)
'lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
End Sub


Private Sub cmdConsultar_Click()
Dim xChaveNF As String, xCNPJ As String, xNomeFornecedor As String, xDataEmissao As String, xValorNF As String, xSituacaoNF As String, e As String
Dim i As Long, xTipoRetorno As String, xSituacaoConfirmacao As String, xTipoNota As String, xTipoEvento As String, xnSeqEvento As String, xCorrecao As String
  cmdConsultar.Enabled = False
  picAguarde.Visible = True
  DoEvents
  vsSQL = "SELECT DATEDIFF(mi, DFeUltimaConsultaData, GETDATE()) As tempo " & _
          "FROM empresa"
  If SQLExecutaRetorno(vsSQL, "tempo", 0) < 30 Then
     MsgBox "A última consulta ocorreu com menos de 30 minutos." & vbNewLine & "Aguarde pelo menos 30 minutos e tente novamente." & vbNewLine, vbInformation + vbOKOnly, "AVISO"
     GoTo Caifora
  End If
  retultNSU = txtUltimoNSU.Text
ConsultaNovamente:
  iRetorno = TransmitirConsultaNFDestinada(retultNSU)
  If iRetorno Then
     If cStat = 138 Then
        vsSQL = "UPDATE empresa SET " & _
                "UltimoNSU = " & retultNSU & ", " & _
                "DFeUltimaConsultaData = CONVERT(DATETIME, '" & Format$(Now, ocDTHR) & "', 103), " & _
                "DFeUltimaConsultaHora = CONVERT(DATETIME, '" & Format$(Now, ocDTHR) & "', 103)"
        SQLExecuta vsSQL
     ElseIf cStat = 137 Or cStat = 656 Then
        vsSQL = "UPDATE empresa SET " & _
                "DFeUltimaConsultaData = CONVERT(DATETIME, '" & Format$(Now, ocDTHR) & "', 103), " & _
                "DFeUltimaConsultaHora = CONVERT(DATETIME, '" & Format$(Now, ocDTHR) & "', 103)"
        SQLExecuta vsSQL
     End If
     If cStat = 138 Then
        Dim sistNFe As snfe.Util
        Set sistNFe = New snfe.Util
        msgResultado = sistNFe.LeArquivoANSI(msgRetWS)
        Parse msgResultado, "#"
        NFeXML = msgResultado
        'Tratamento do XML de Retorno
        For i = 0 To cStat2 - 1
            If retornoTipo(i) = "resNFe" Then
               xTipoRetorno = "resNFe"
               xChaveNF = retornochNFe(i)
               xCNPJ = retornoCNPJ(i)
               xNomeFornecedor = retornoxNome(i)
               xDataEmissao = retornodhEmi(i)
               xValorNF = retornovNF(i)
               xValorNF = Substitui(xValorNF, ".", ",", UM_A_UM)
               xTipoNota = retornotpNF(i)
               xSituacaoNF = retornocSitNFe(i)
               xSituacaoConfirmacao = retornocSitConf(i)
               If Not Vazio(xDataEmissao) Then xDataEmissao = Format(Left(xDataEmissao, 10), "yyyy-mm-dd") ' & " às " & Mid(xDataEmissao, 12, 8)
               NFeValidate = xCNPJ & "|" & xChaveNF & "|" & xNomeFornecedor & "|" & xDataEmissao & "|" & xValorNF & "|" & xSituacaoNF
               vsSQL = "SELECT COUNT(*) r FROM TbNotasDestinadasManifestacao WHERE ChaveAcesso = '" & xChaveNF & "'"
               If SQLExecutaRetorno(vsSQL, "r", 0) = 0 Then
                  vsSQL = "INSERT INTO TbNotasDestinadasManifestacao (Data, TipoRetornoWS, TipoNotaWS, CNPJFornecedor, ChaveAcesso, RazaoSocial, DataEmissao, ValorNota, SituacaoNota, NumeroProtocolo, IdFilial, DataInclusao) VALUES " & _
                          "(CONVERT(DATETIME, '" & Format$(Date, ocDATA) & "', 103), '" & xTipoRetorno & "', " & xTipoNota & ", '" & xCNPJ & "', '" & xChaveNF & "', '" & RemoveAcento(xNomeFornecedor) & "', CONVERT(DATETIME, '" & Format$(xDataEmissao, ocDATA) & "', 103), " & FSQL(xValorNF) & ", " & IIf(Vazio(xSituacaoNF), "0", xSituacaoNF) & ", " & IIf(Vazio(xSituacaoConfirmacao), "0", xSituacaoConfirmacao) & ", 1, CONVERT(DATETIME, '" & Format$(Date, ocDATA) & "', 103))"
                  e$ = SQLExecuta(vsSQL)
                  If Not Vazio(e$) Then
                     MsgBox e$, vbCritical + vbOKOnly, vgAtencao
                     e$ = ""
                  End If
               End If
            ElseIf retornoTipo(i) = "procNFe" Then
               xTipoRetorno = "procNFe"
               xChaveNF = retornochNFe(i)
               xCNPJ = retornoCNPJ(i)
               xNomeFornecedor = retornoxNome(i)
               xDataEmissao = retornodhEmi(i)
               xValorNF = retornovNF(i)
               xValorNF = Substitui(xValorNF, ".", ",", UM_A_UM)
               xTipoNota = retornotpNF(i)
               xSituacaoNF = retornocSitNFe(i)
               xSituacaoConfirmacao = "1"   'retornocSitConf(i)
               vsSQL = "UPDATE TbNotasDestinadasManifestacao SET Enviada = 1, SituacaoConfirmacao = 1, DataHoraProcotolo = '" & Format$(xDataEmissao, ocDATA) & "', 103) & " ' WHERE ChaveAcesso = '" & xChaveNF & "'"
               e$ = SQLExecuta(vsSQL)
               If Not Vazio(e$) Then
                  MsgBox e$, vbCritical + vbOKOnly, vgAtencao
                  e$ = ""
               End If
            ElseIf retornoTipo(i) = "resEvento" Then
               xTipoRetorno = "resEvento"
               xChaveNF = retornochNFe(i)
               xDataEmissao = retornodhEmi(i)
               xTipoNota = "1"
               xValorNF = "0"
               xTipoEvento = retornotpNF(i)
               xnSeqEvento = retornocSitConf(i)
               xCorrecao = retornoxNome(i)
            ElseIf retornoTipo(i) = "procEventoNFe" Then
               xTipoRetorno = "procEventoNFe"
               xChaveNF = retornochNFe(i)
               xDataEmissao = retornodhEmi(i)
               xTipoNota = "1"
               xValorNF = "0"
               xTipoEvento = retornotpNF(i)
               xnSeqEvento = retornocSitConf(i)
               xCorrecao = retornoxNome(i)
               xSituacaoNF = retornocSitNFe(i)
               If Left(xTipoEvento, 2) = 11 Or Left(xTipoEvento, 2) = 21 Then
                  vsSQL = "UPDATE TbNotasDestinadasManifestacao SET TipoEventoCodigo = '" & xTipoEvento & "', TipoEvento = '" & xCorrecao & "', EventoProtocolo = " & xSituacaoNF & ", EventoDataHora = '" & xDataEmissao & "' WHERE ChaveAcesso = '" & xChaveNF & "'"
                  e$ = SQLExecuta(vsSQL)
                  If Not Vazio(e$) Then
                     MsgBox e$, vbCritical + vbOKOnly, vgAtencao
                     e$ = ""
                  End If
               End If
            End If
            xTipoRetorno = ""
            xChaveNF = ""
            xCNPJ = ""
            xNomeFornecedor = ""
            xDataEmissao = ""
            xValorNF = ""
            xTipoNota = ""
            xTipoEvento = ""
            xSituacaoNF = ""
            xSituacaoConfirmacao = ""
            xnSeqEvento = ""
            xCorrecao = ""
            txtUltimoNSU.Text = ""
            txtUltimoNSU.SetFocus
        Next i
        Set sistNFe = Nothing
     End If
     Call cmdAtualizarGrid_Click
     If retindCont = 1 Then GoTo ConsultaNovamente
     picAguarde.Visible = False
     DoEvents
  Else
     cmdConsultar.Enabled = True
  End If
Caifora:
  picAguarde.Visible = False
  DoEvents
  cmdConsultar.Enabled = True
End Sub

Private Sub cmdCopiarXML_Click()
Dim vChave As String, i As Long

For i = 0 To Grid.rows - 1
   Grid.Row = i
   Grid.Col = 0
   
   If Grid.CellPicture = ImgMarcada Then
      vChave = (Grid.TextMatrix(Grid.Row, 7))
   End If
Next

Clipboard.Clear
Clipboard.SetText vChave
End Sub

Private Sub cmdDANFe_Click()
'Dim chaveNFe As String, i As Long
Dim vSituacao As String, vChave As String, i As Long

For i = 0 To Grid.rows - 1
   Grid.Row = i
   Grid.Col = 0
   
   If Grid.CellPicture = ImgMarcada Then
      vSituacao = (Grid.TextMatrix(Grid.Row, 2))
      vChave = (Grid.TextMatrix(Grid.Row, 7))
   End If
Next

If vSituacao = "Sem Manifestar" Then
    MsgBox "A Nota precisa ser manifestada primeiramente!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim sistNFe As snfe.Util
  
  On Error GoTo deuErro
  
  Set sistNFe = New snfe.Util
    
  dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", "")
  dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
  
  'i = Grid.Row
  'chaveNFe = Grid.TextMatrix(i, 7)
  
  xCaminhoXML = dirXML & "nfe\arquivos\ConfRecebto\" & vChave & "-procNFe.xml"

  If Not Existe(xCaminhoXML) Then
     MsgBox "Arquivo XML: " & xCaminhoXML & " não foi localizado!", vbExclamation + vbOKOnly, "ATENÇÃO"
     Exit Sub
  End If

  iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
  Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", False, "", 0, False, False, "", False, False, True, False, True)
 
  Set sistNFe = Nothing
  Exit Sub
 
 Resume
 
deuErro:
 MsgBox Err.Description, vbCritical + vbOKOnly, vgAtencao
 Err.Clear
 Set sistNFe = Nothing
End Sub

Private Sub cmdEnviarManifestacao_Click()
Dim SQL As String, TipoEvento As String, xTipoEvento As String, Justificativa As String
Dim DataHora As Variant, i As Long
Dim vSituacao As String, vEvento As String, vCnpj As String, vChave As String, vID As Integer

For i = 0 To Grid.rows - 1
   Grid.Row = i
   Grid.Col = 0
   
   If Grid.CellPicture = ImgMarcada Then
      vSituacao = (Grid.TextMatrix(Grid.Row, 2))
      vEvento = (Grid.TextMatrix(Grid.Row, 3))
      vCnpj = (Grid.TextMatrix(Grid.Row, 5))
      vChave = (Grid.TextMatrix(Grid.Row, 7))
      vID = (Grid.TextMatrix(Grid.Row, 10))
   End If
Next

If vSituacao = "Sem Manifestar" And vEvento <> "" Then
    MsgBox "A Nota espera um retorno da sefaz para mudar sua situação!", vbInformation, "Aviso do Sistema"
    Exit Sub
ElseIf vSituacao <> "Sem Manifestar" Then
    MsgBox "A Nota já foi manifestada anteriormente!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

On Error GoTo deuErro

     TipoEvento = "1"
tentaNovamente:
     TipoEvento = InputBox("INFORME O TIPO DA MANIFESTAÇÃO" & vbNewLine & vbNewLine & "0 - Confirmação da Operação" & vbNewLine & "1 - Ciência da Operação" & vbNewLine & "2 - Desconhecimento da Operação" & vbNewLine & "3 - Registro da Operação não Realizada", "MANIFESTAÇÃO DO DESTINATÁRIO", TipoEvento)
     If Not IsNumeric(TipoEvento) Then GoTo tentaNovamente
     If Int(TipoEvento) > 4 Then GoTo tentaNovamente
     DataHora = Format(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
     iRetorno = TransmitirManDest(vCnpj, vChave, DataHora, TipoEvento, Justificativa)
     If iRetorno Then
        Select Case TipoEvento
           Case "0"
              xTipoEvento = "0 - Confirmação da Operação"
           Case "1"
              xTipoEvento = "1 - Ciência da Operação"
           Case "2"
              xTipoEvento = "2 - Desconhecimento da Operação"
           Case "3"
              xTipoEvento = "3 - Registro da Operação não Realizada"
        End Select
        vsSQL = "UPDATE TbNotasDestinadasManifestacao SET " & _
                "TipoEvento = '" & xTipoEvento & "', " & _
                "nSeqEvento = 1, " & _
                "Justificativa = '" & Justificativa & "', " & _
                "Data = GETDATE(), " & _
                "Enviada = " & IIf(Int(TipoEvento) <> 1, 1, 0) & ", " & _
                "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
                "DataHoraProcotolo = '" & NFeDataHora & "', " & _
                "Enviar = 0, " & _
                "XMLAutorizado = '" & NFeXML & "' " & _
                "WHERE IDNFConsulta = " & vID
        SQLExecuta vsSQL
     End If
     Call cmdAtualizarGrid_Click
  Exit Sub
  
deuErro:
  Screen.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
  Err.Clear
End Sub

Private Sub cmdEnviarManifestacao2_Click()
Dim SQL As String, TipoEvento As String, xTipoEvento As String, Justificativa As String
Dim DataHora As Variant, i As Long
Dim vSituacao As String, vEvento As String, vCnpj As String, vChave As String, vID As Integer

    TipoEvento = "1"
tentaNovamente:
    TipoEvento = InputBox("INFORME O TIPO DA MANIFESTAÇÃO" & vbNewLine & vbNewLine & "0 - Confirmação da Operação" & vbNewLine & "1 - Ciência da Operação" & vbNewLine & "2 - Desconhecimento da Operação" & vbNewLine & "3 - Registro da Operação não Realizada", "MANIFESTAÇÃO DO DESTINATÁRIO", TipoEvento)
    If Not IsNumeric(TipoEvento) Then GoTo tentaNovamente
    If Int(TipoEvento) > 4 Then GoTo tentaNovamente
    
    For i = 0 To Grid.rows - 1
       Grid.Row = i
       Grid.Col = 0
       
       If Grid.CellPicture = ImgMarcada Then
            vSituacao = (Grid.TextMatrix(Grid.Row, 2))
            vEvento = (Grid.TextMatrix(Grid.Row, 3))
            vCnpj = (Grid.TextMatrix(Grid.Row, 5))
            vChave = (Grid.TextMatrix(Grid.Row, 7))
            vID = (Grid.TextMatrix(Grid.Row, 10))
       
            If vSituacao = "Sem Manifestar" And vEvento <> "" Then
                MsgBox "Você selecionou nota(s) que espera(m) um retorno da sefaz para mudar sua situação!", vbInformation, "Aviso do Sistema"
                Exit Sub
            ElseIf vSituacao <> "Sem Manifestar" Then
                MsgBox "Você selecionou nota(s) que já foi(foram) manifestada(s) anteriormente!", vbInformation, "Aviso do Sistema"
                Exit Sub
            End If
       
            On Error GoTo deuErro
    
            DataHora = Format(Date, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & UTC
            iRetorno = TransmitirManDest(vCnpj, vChave, DataHora, TipoEvento, Justificativa, True)
            If iRetorno Then
                Select Case TipoEvento
                   Case "0"
                      xTipoEvento = "0 - Confirmação da Operação"
                   Case "1"
                      xTipoEvento = "1 - Ciência da Operação"
                   Case "2"
                      xTipoEvento = "2 - Desconhecimento da Operação"
                   Case "3"
                      xTipoEvento = "3 - Registro da Operação não Realizada"
                End Select
                vsSQL = "UPDATE TbNotasDestinadasManifestacao SET " & _
                        "TipoEvento = '" & xTipoEvento & "', " & _
                        "nSeqEvento = 1, " & _
                        "Justificativa = '" & Justificativa & "', " & _
                        "Data = GETDATE(), " & _
                        "Enviada = " & IIf(Int(TipoEvento) <> 1, 1, 0) & ", " & _
                        "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
                        "DataHoraProcotolo = '" & NFeDataHora & "', " & _
                        "Enviar = 0, " & _
                        "XMLAutorizado = '" & NFeXML & "' " & _
                        "WHERE IDNFConsulta = " & vID
                SQLExecuta vsSQL
            End If
       End If
    Next
  Call cmdAtualizarGrid_Click
  Exit Sub
  
deuErro:
  Screen.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
  Err.Clear
End Sub

Private Sub cmdImportarNFe_Click()
Dim i As Long
Dim vSituacao As String, vChave As String

For i = 0 To Grid.rows - 1
   Grid.Row = i
   Grid.Col = 0
   
   If Grid.CellPicture = ImgMarcada Then
      vSituacao = (Grid.TextMatrix(Grid.Row, 2))
      vChave = (Grid.TextMatrix(Grid.Row, 7))
   End If
Next
     
If vSituacao = "Sem Manifestar" Then
    MsgBox "A Nota precisa ser manifestada primeiramente!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

'chaveNFe = Grid.TextMatrix(i, 7)
If Vazio(vChave) Then Exit Sub
dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", "")
dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
xCaminhoXML = dirXML & "nfe\arquivos\ConfRecebto\" & vChave & "-procNFe.xml"
If Existe(xCaminhoXML) = -1 Then
   Load Entrada_Estoque
   Call Entrada_Estoque.ImportarXML(xCaminhoXML)
   Entrada_Estoque.CriarEntradaEstoque Compra
   Entrada_Estoque.Show 1
   Call cmdAtualizarGrid_Click
Else
   MsgBox "Arquivo XML " & xCaminhoXML & " não localizado!" & vbNewLine & "Importação do XML cancelada!", vbExclamation + vbOKOnly, "ATENÇÃO"
End If
End Sub


Private Sub cmdLimparSelecao_Click()
cboTipo.Text = "Todas"
mskInicio.Mask = ""
mskInicio.Text = ""
mskFim.Mask = ""
mskFim.Text = ""

'If cboDesc.Text = "" Then
    sSQL = "SELECT SituacaoNF, TipoRetornoWS, TipoEventoID, TipoEvento, TipoNotaWS, Data, CNPJFornecedor, RazaoSocial, ChaveAcesso, DataEmissao, ValorNota, SituacaoNota, ImportadoXML, Enviar, Enviada, CodigoConsulta, IdNFConsulta, IdFilial, TipoEventoCodigo, EventoProtocolo, EventoDataHora " & _
           "FROM NotasDestinadas " & _
           "WHERE 1=0  " & _
           "ORDER BY IdNFConsulta DESC"
'Else
'    sSQL = "SELECT IdInventario, Seq as vSeq, IDProduto as vCodProd, NomeProduto as vDesc, EAN as vEAN, NCM as vNCM, SaldoCalculado, VlrUnitInvent as vVlrUnit, TotalInvent as vTotal, MetaCalculado  as vMeta, SaldoLancado, TotalFisico " & _
           "FROM InventarioGerado " & _
           "WHERE (NomeProduto = '" & cboDesc.Text & "')  " & _
           "ORDER BY Seq"
'End If
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
var_Contador = 0
MostrarDados
End Sub

Private Sub Grid_Click()
If Grid.Col <> 0 Then Exit Sub

If Grid.CellPicture = imgDesmarcada Then
   Set Grid.CellPicture = ImgMarcada
Else
   Set Grid.CellPicture = imgDesmarcada
End If

OP = contar
AcaoGrid
End Sub

Sub AcaoGrid()
Dim i As Integer
'Dim var_Contador As Integer

Grid.Col = 0

For i = 1 To Grid.rows - 1
   Grid.Row = i
   If OP = MarcarTodos Then Set Grid.CellPicture = ImgMarcada
   If OP = DesmarcarTodos Then Set Grid.CellPicture = imgDesmarcada
   If OP = contar Then
        If Grid.CellPicture = ImgMarcada Then var_Contador = var_Contador + 1
   End If
Next

lblRegistroSel.Caption = var_Contador & " Registro(s) Selecionado(s)"

If var_Contador = 1 Then
    cmdEnviarManifestacao.Visible = True
    cmdEnviarManifestacao.Enabled = True
    cmdEnviarManifestacao2.Visible = False
    cmdImportarNFe.Enabled = True
    cmdDANFe.Enabled = True
    cmdCopiarXML.Enabled = True
ElseIf var_Contador > 1 Then
    cmdEnviarManifestacao2.Visible = True
    cmdEnviarManifestacao2.Enabled = True
    'cmdEnviarManifestacao2.Visible = False
    cmdEnviarManifestacao.Visible = False
    cmdImportarNFe.Enabled = False
    cmdDANFe.Enabled = False
    cmdCopiarXML.Enabled = False
ElseIf var_Contador = 0 Then
    cmdEnviarManifestacao2.Visible = False
    cmdEnviarManifestacao.Visible = True
    cmdEnviarManifestacao.Enabled = False
    cmdImportarNFe.Enabled = False
    cmdDANFe.Enabled = False
    cmdCopiarXML.Enabled = False
End If

var_Contador = 0
End Sub

Private Sub mskFim_GotFocus()
SelectControl mskFim
End Sub
Private Sub mskFim_KeyPress(KeyAscii As Integer)
mskFim.Mask = "##/##/##"
End Sub
Private Sub mskFim_LostFocus()
If mskFim.Text = "" Or mskFim.Text = "__/__/__" Then
   mskFim.Mask = ""
   mskFim.Text = ""
Else
   If Not IsDate(mskFim.Text) Then
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskFim.SetFocus
   End If
End If
End Sub

Private Sub mskInicio_GotFocus()
SelectControl mskInicio
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
mskInicio.Mask = "##/##/##"
End Sub

Private Sub mskInicio_LostFocus()
If mskInicio.Text = "" Or mskInicio.Text = "__/__/__" Then
   mskInicio.Mask = ""
   mskInicio.Text = ""
Else
   If Not IsDate(mskInicio.Text) Then
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskInicio.SetFocus
   End If
End If
End Sub


