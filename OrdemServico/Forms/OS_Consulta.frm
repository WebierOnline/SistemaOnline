VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form OS_Consulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSULTA DE ORDEM DE SERVIÇOS"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmConsultaSimples2 
      Height          =   915
      Left            =   9420
      TabIndex        =   28
      Top             =   480
      Width           =   5055
      Begin VB.OptionButton optDataPrevisao 
         Caption         =   "Previsão"
         Height          =   195
         Left            =   3000
         TabIndex        =   42
         Top             =   240
         Width           =   915
      End
      Begin VB.ComboBox cboAnoConsulta 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.ComboBox cboMesConsulta 
         Height          =   315
         Left            =   60
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.ComboBox cboLocalizar 
         Height          =   315
         Left            =   60
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   4905
      End
      Begin VB.TextBox txtCodClienteLocalizar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4380
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OptionButton optDataEntrada 
         Caption         =   "Entrada"
         Height          =   195
         Left            =   1020
         TabIndex        =   30
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton optDataTermino 
         Caption         =   "Termino"
         Height          =   195
         Left            =   1980
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   915
      End
      Begin ChamaleonBtn.chameleonButton cmdCal1 
         Height          =   315
         Left            =   1320
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "Calendario"
         Top             =   480
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
         MICON           =   "OS_Consulta.frx":0000
         PICN            =   "OS_Consulta.frx":001C
         PICH            =   "OS_Consulta.frx":236F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox mskDataConsulta 
         Height          =   315
         Left            =   60
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPeriodoInicio 
         Height          =   315
         Left            =   60
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPeriodoFim 
         Height          =   315
         Left            =   1980
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdCal2 
         Height          =   315
         Left            =   3240
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "Calendario"
         Top             =   480
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
         MICON           =   "OS_Consulta.frx":46C2
         PICN            =   "OS_Consulta.frx":46DE
         PICH            =   "OS_Consulta.frx":6A31
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblPeriodoAte 
         AutoSize        =   -1  'True
         Caption         =   "até"
         Height          =   195
         Left            =   1680
         TabIndex        =   41
         Top             =   540
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblCriterio 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   60
         TabIndex        =   40
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.OptionButton optFiltroRefinado 
      Caption         =   "Filtro Refinado"
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
      Left            =   1800
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optFiltroSimples 
      Caption         =   "Filtro Simples"
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
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Frame frmConsultaRefina 
      Caption         =   "FILTRO REFINADO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   60
      TabIndex        =   16
      Top             =   1440
      Width           =   14415
      Begin VB.CheckBox chkPlaca 
         Caption         =   "Placa"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
      Begin VB.CheckBox chkChassi 
         Caption         =   "Chassi"
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   300
         Width           =   795
      End
      Begin VB.TextBox txtFiltroRefinado 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   540
         Width           =   3735
      End
   End
   Begin VB.Frame frmConsultaSimples 
      Caption         =   "FILTROS SIMPLES"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   60
      TabIndex        =   9
      Top             =   480
      Width           =   9255
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   4860
         TabIndex        =   26
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cboConsultaCriterios 
         Height          =   315
         Left            =   7620
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboConsultaMostrar 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cboConsultaStatus 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cboTipoServico 
         Height          =   315
         Left            =   3300
         TabIndex        =   2
         Top             =   480
         Width           =   1515
      End
      Begin VB.ComboBox cboIndice 
         Height          =   315
         Left            =   6240
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   4860
         TabIndex        =   27
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Critérios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   7620
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Organização:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   6240
         TabIndex        =   11
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Serviço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3300
         TabIndex        =   12
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Financeiro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Técnico:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Width           =   690
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5655
      Left            =   60
      TabIndex        =   15
      Top             =   2760
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   9975
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibir 
      Height          =   315
      Left            =   12780
      TabIndex        =   5
      Top             =   2400
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Exibir"
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
      MICON           =   "OS_Consulta.frx":8D84
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   8460
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "DETALHAMENTO"
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "OS_Consulta.frx":8DA0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirParcelas 
      Height          =   255
      Left            =   2340
      TabIndex        =   20
      Top             =   8460
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "FINANCEIRO"
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "OS_Consulta.frx":8DBC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   255
      Left            =   4620
      TabIndex        =   21
      Top             =   8460
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "IMPRIMIR"
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
      MICON           =   "OS_Consulta.frx":8DD8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   735
      Left            =   11760
      Top             =   8460
      Width           =   2715
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANT.:"
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
      Left            =   11880
      TabIndex        =   25
      Top             =   8520
      Width           =   780
   End
   Begin VB.Label lblQtda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   12720
      TabIndex        =   24
      Top             =   8520
      Width           =   1635
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   12720
      TabIndex        =   23
      Top             =   8820
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
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
      Left            =   12000
      TabIndex        =   22
      Top             =   8820
      Width           =   675
   End
End
Attribute VB_Name = "OS_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lCodOSSelecionado As Long
Private moCombo As cComboHelper
Dim printSQL As String
Dim sSQL As String
Dim r As ADODB.Recordset

Private Sub Form_Load()
Set moCombo = New cComboHelper
frmConsultaRefina.Enabled = False
frmConsultaSimples2.Enabled = True
lCodOSSelecionado = 0

Preencher_TipoServico
Preencher_Mostrar
Preencher_Status
Preencher_Criterios
Preencher_Indice

cboConsultaMostrar.ListIndex = 0
cboConsultaStatus.ListIndex = 0
cboTipoServico.ListIndex = 0
Preencher_Tipo
cboTipo.Text = "TODOS"
cboIndice.ListIndex = 0

cboMesConsulta_GotFocus
cboMesConsulta.ListIndex = Month(Date) - 1
cboAnoConsulta_GotFocus
cboAnoConsulta.Text = Year(Date)

cboConsultaCriterios.Text = "MENSAL"
AtualizarCamposCriterios

MostrarGrid_OS
End Sub
Private Sub cmdExibirParcelas_Click()
If Grid.Row = 0 Then Exit Sub
If Grid.TextMatrix(Grid.Row, 0) = "" Then Exit Sub
If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub

Dim lPedido As Long
lPedido = CLng(Grid.TextMatrix(Grid.Row, 11))
If lPedido = 0 Then Exit Sub

Vendas_Consulta_Geral_Parcelas.loadInformacoes lPedido
Vendas_Consulta_Geral_Parcelas.Show 1
End Sub
Private Sub cmdExibirPedidos_Click()
If Grid.Row = 0 Then Exit Sub
If Grid.TextMatrix(Grid.Row, 0) = "" Then Exit Sub
If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub
If CLng(Grid.TextMatrix(Grid.Row, 11)) = 0 Then Exit Sub

Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 11)), "OS"
Parcelas_Consulta_Produtos.Show 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
lCodOSSelecionado = 0
End Sub



Private Sub cmdCal1_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty

Set fCal = New Calendario
fCal.Show vbModal

varData = fCal.DateSelected

Unload fCal
Set fCal = Nothing

If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub

If cboConsultaCriterios.Text = "PERÍODO" Then
   mskPeriodoInicio = Format(varData, "dd/mm/yy")
Else
   mskDataConsulta = Format(varData, "dd/mm/yy")
End If
End Sub

Private Sub cmdCal2_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty

Set fCal = New Calendario
fCal.Show vbModal

varData = fCal.DateSelected

Unload fCal
Set fCal = Nothing

If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub

mskPeriodoFim = Format(varData, "dd/mm/yy")
End Sub

Private Sub cboConsultaCriterios_Click()
AtualizarCamposCriterios
If cboConsultaCriterios.Text = "TODOS" Then
   MostrarGrid_OS
ElseIf cboConsultaCriterios.Text = "CLIENTE" Or cboConsultaCriterios.Text = "CÓD. OS" Then
   cboLocalizar.SetFocus
ElseIf cboConsultaCriterios.Text = "DATA" Then
   mskDataConsulta.SetFocus
ElseIf cboConsultaCriterios.Text = "PERÍODO" Then
   mskPeriodoInicio.SetFocus
ElseIf cboConsultaCriterios.Text = "MENSAL" Then
   cboMesConsulta.SetFocus
End If
End Sub

Private Sub AtualizarCamposCriterios()
cboLocalizar.Visible = False
mskDataConsulta.Visible = False
cmdCal1.Visible = False
mskPeriodoInicio.Visible = False
lblPeriodoAte.Visible = False
mskPeriodoFim.Visible = False
cmdCal2.Visible = False
cboMesConsulta.Visible = False
cboAnoConsulta.Visible = False
optDataEntrada.Visible = False
optDataTermino.Visible = False
optDataPrevisao.Visible = False

If cboConsultaCriterios.Text = "TODOS" Then
   cboLocalizar.Text = ""
   lblCriterio(5).Caption = ""
ElseIf cboConsultaCriterios.Text = "CLIENTE" Then
   cboLocalizar.Visible = True
   lblCriterio(5).Caption = "Cliente:"
ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then
   cboLocalizar.Visible = True
   lblCriterio(5).Caption = "Código:"
ElseIf cboConsultaCriterios.Text = "DATA" Then
   mskDataConsulta.Visible = True
   cmdCal1.Visible = True
   optDataEntrada.Visible = True
   optDataTermino.Visible = True
   optDataPrevisao.Visible = True
   lblCriterio(5).Caption = "Data:"
ElseIf cboConsultaCriterios.Text = "PERÍODO" Then
   mskPeriodoInicio.Visible = True
   cmdCal1.Visible = True
   lblPeriodoAte.Visible = True
   mskPeriodoFim.Visible = True
   cmdCal2.Visible = True
   optDataEntrada.Visible = True
   optDataTermino.Visible = True
   optDataPrevisao.Visible = True
   lblCriterio(5).Caption = "Período:"
ElseIf cboConsultaCriterios.Text = "MENSAL" Then
   cboMesConsulta.Visible = True
   cboAnoConsulta.Visible = True
   optDataEntrada.Visible = True
   optDataTermino.Visible = True
   optDataPrevisao.Visible = True
   lblCriterio(5).Caption = "Mês/Ano:"
End If
End Sub

Private Sub cboConsultaCriterios_GotFocus()
Dim itemAtual As String
itemAtual = cboConsultaCriterios.Text
Preencher_Criterios
cboConsultaCriterios.Text = itemAtual
moCombo.AttachTo cboConsultaCriterios
End Sub

Private Sub cboConsultaCriterios_Validate(Cancel As Boolean)
AtualizarCamposCriterios
End Sub

Private Sub cboConsultaMostrar_Change()
''MostrarGrid_OS
End Sub

Private Sub cboConsultaMostrar_Click()
''MostrarGrid_OS
End Sub

Private Sub cboConsultaMostrar_GotFocus()
Dim itemAtual As String
itemAtual = cboConsultaMostrar.Text
Preencher_Mostrar
cboConsultaMostrar.Text = itemAtual
moCombo.AttachTo cboConsultaMostrar
End Sub

Private Sub cboConsultaMostrar_Validate(Cancel As Boolean)
''MostrarGrid_OS
End Sub

Private Sub cboConsultaStatus_Change()
''MostrarGrid_OS
End Sub

Private Sub cboConsultaStatus_Click()
''MostrarGrid_OS
End Sub

Private Sub cboConsultaStatus_GotFocus()
Dim itemAtual As String
itemAtual = cboConsultaStatus.Text
Preencher_Status
cboConsultaStatus.Text = itemAtual
moCombo.AttachTo cboConsultaStatus
End Sub

Private Sub cboConsultaStatus_Validate(Cancel As Boolean)
''MostrarGrid_OS
End Sub

Private Sub cboIndice_Change()
''MostrarGrid_OS
End Sub

Private Sub cboIndice_Click()
''MostrarGrid_OS
End Sub

Private Sub cboIndice_GotFocus()
Dim varNomeAntes As String
varNomeAntes = cboIndice.Text

Preencher_Indice

cboIndice.Text = varNomeAntes
moCombo.AttachTo cboIndice
End Sub

Private Sub cboLocalizar_GotFocus()

If cboConsultaCriterios.Text = "CLIENTE" Then
   cboLocalizar.Clear
   
   sSQL = "SELECT codigo, nome FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboLocalizar.AddItem r("nome")
      cboLocalizar.ItemData(cboLocalizar.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   SelectControl cboLocalizar
   moCombo.AttachTo cboLocalizar
ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then
   cboLocalizar.Clear
ElseIf cboConsultaCriterios.Text = "TODOS" Then
   cboLocalizar.Text = ""
End If
End Sub

Private Sub cboLocalizar_LostFocus()
   On Error GoTo TrataErro

If cboConsultaCriterios.Text = "CLIENTE" Then
   If cboLocalizar.Text = "" Then Exit Sub
   If cboLocalizar.ListIndex = -1 Then txtCodClienteLocalizar.Text = "": Exit Sub
   txtCodClienteLocalizar = cboLocalizar.ItemData(cboLocalizar.ListIndex)
   Exit Sub
End If

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboMesConsulta_GotFocus()
cboMesConsulta.Clear
cboMesConsulta.AddItem "Janeiro"
cboMesConsulta.AddItem "Fevereiro"
cboMesConsulta.AddItem "Março"
cboMesConsulta.AddItem "Abril"
cboMesConsulta.AddItem "Maio"
cboMesConsulta.AddItem "Junho"
cboMesConsulta.AddItem "Julho"
cboMesConsulta.AddItem "Agosto"
cboMesConsulta.AddItem "Setembro"
cboMesConsulta.AddItem "Outubro"
cboMesConsulta.AddItem "Novembro"
cboMesConsulta.AddItem "Dezembro"
moCombo.AttachTo cboMesConsulta
End Sub

Private Sub cboAnoConsulta_GotFocus()
Dim iAno As Integer
Dim i As Integer
cboAnoConsulta.Clear
iAno = Year(Date)
For i = iAno - 5 To iAno + 1
   cboAnoConsulta.AddItem i
Next
moCombo.AttachTo cboAnoConsulta
End Sub

Private Sub cboTipoServico_Change()
''MostrarGrid_OS
End Sub

Private Sub cboTipoServico_Click()
''MostrarGrid_OS
End Sub

Private Sub cboTipoServico_GotFocus()
Dim varNomeAntes As String
varNomeAntes = cboTipoServico.Text

Preencher_TipoServico

cboTipoServico.Text = varNomeAntes
moCombo.AttachTo cboTipoServico
End Sub

Private Sub cmdExibir_Click()
If optFiltroRefinado.Value = True Then
   MostrarGrid_OS_Refinado
Else
   MostrarGrid_OS
End If
End Sub

Private Sub optFiltroSimples_Click()
frmConsultaSimples.Enabled = True
frmConsultaSimples2.Enabled = True
frmConsultaRefina.Enabled = False
End Sub

Private Sub optFiltroRefinado_Click()
frmConsultaSimples.Enabled = False
frmConsultaSimples2.Enabled = False
frmConsultaRefina.Enabled = True
End Sub

Private Sub chkPlaca_Click()
If chkPlaca.Value = 1 Then chkChassi.Value = 0
End Sub

Private Sub chkChassi_Click()
If chkChassi.Value = 1 Then chkPlaca.Value = 0
End Sub

Private Sub MostrarGrid_OS_Refinado()
Dim totalRegistros As Long
Dim campoBusca As String

If vTipoOS <> "Automóveis" And vTipoOS <> "Motocicletas" And vTipoOS <> "Recapadora" Then
   MsgBox "Consulta por Placa/Chassi disponível apenas para veículos!", vbInformation, "Aviso do Sistema"
   Exit Sub
End If

If chkPlaca.Value = 1 Then
   campoBusca = "OS_Equipamento_Auto.PLACA"
ElseIf chkChassi.Value = 1 Then
   campoBusca = "OS_Equipamento_Auto.CHASSI"
Else
   MsgBox "Selecione Placa ou Chassi!", vbInformation, "Aviso do Sistema"
   Exit Sub
End If

If txtFiltroRefinado.Text = "" Then
   MsgBox "Digite a Placa ou o Chassi para consultar!", vbInformation, "Aviso do Sistema"
   Exit Sub
End If

Dim SITUACAO As String
Dim var_STATUS As String

'Status
If cboConsultaStatus.Text = "TODOS" Then
   SITUACAO = ""
ElseIf cboConsultaStatus.Text = "À COMEÇAR" Then
   SITUACAO = "AND (os.status = 'À COMEÇAR') "
ElseIf cboConsultaStatus.Text = "EM EXECUÇÃO" Then
   SITUACAO = "AND (os.status = 'EM EXECUÇÃO') "
ElseIf cboConsultaStatus.Text = "AGUARDANDO" Then
   SITUACAO = "AND (os.status = 'AGUARDANDO') "
ElseIf cboConsultaStatus.Text = "TERMINADO" Then
   SITUACAO = "AND (os.status = 'TERMINADO') "
End If

'Situação
If cboConsultaMostrar.Text = "TODOS" Then
   var_STATUS = ""
ElseIf cboConsultaMostrar.Text = "ABERTOS" Then
   var_STATUS = "AND (status_os = 0) "
ElseIf cboConsultaMostrar.Text = "FECHADOS" Then
   var_STATUS = "AND (status_os = 1) "
End If

'forma de pagamento (TODOS/À VISTA/À PRAZO)
Dim varTipoPagamento As String
If cboTipo.Text = "À VISTA" Then
   varTipoPagamento = "AND (os.tipo_pagamento = 'À Vista') "
ElseIf cboTipo.Text = "À PRAZO" Then
   varTipoPagamento = "AND (os.tipo_pagamento = 'À Prazo') "
Else
   varTipoPagamento = ""
End If

sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento_Auto.fabricante,'') + ' / ' + ISNULL(OS_Equipamento_Auto.modelo,'') + ' / ' + ISNULL(CAST(OS_Equipamento_Auto.ano AS VARCHAR(10)),'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
   "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _
   "WHERE (" & campoBusca & " = '" & txtFiltroRefinado.Text & "') " & _
   SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY os.DATA_TERMINO DESC"

Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid_OS r

printSQL = sSQL

lblQtda.Caption = Format(totalRegistros, "000")

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cmdImprimir_Click()
'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_OS_Consulta.Relatorio.Recordset = r

REL_OS_Consulta.dfQuant.Caption = lblQtda.Caption
REL_OS_Consulta.dfTotal.Caption = lblTotal.Caption
REL_OS_Consulta.lblTitulo.Caption = "RELATÓRIO - CONSULTA DE ORDEM DE SERVIÇOS"
REL_OS_Consulta.rfData.Caption = "Data: " & Format(Now, "dd/mm/yy") & " às " & Format(Now, "hh:nn") & "hs"

'If cboFiltro.Text = "TODOS" Then
'   REL_OS_Consulta.dfTipo.Caption = "Tipo: Todos os registros"
'ElseIf cboFiltro.Text = "PERIODO" Then
'   REL_OS_Consulta.dfTipo.Caption = "Tipo: Intervalo de " & Mask1.Text & " à " & Mask2.Text
'ElseIf cboFiltro.Text = "MÊS" Then
'   REL_OS_Consulta.dfTipo.Caption = "Tipo: Mês = " & cboMes.Text & "/" & cboAno.Text
'ElseIf cboFiltro.Text = "CLIENTE" Then
'   REL_OS_Consulta.dfTipo.Caption = "Cliente = " & cboNome.Text
'Else
'   REL_OS_Consulta.dfTipo.Caption = "Tipo:"
'End If

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
   REL_OS_Consulta.Label4.Caption = "CLIENTE / VEÍCULOS"
Else
   REL_OS_Consulta.Label4.Caption = "CLIENTE / EQUIPAMENTO"
End If

Dim sTipoConsulta As String
Dim sCriterio As String
Dim sFiltrosAdic As String

If optFiltroRefinado.Value = True Then
   sTipoConsulta = "Refinado - " & IIf(chkPlaca.Value = 1, "PLACA", "CHASSI")
   sCriterio = txtFiltroRefinado.Text
Else
   sTipoConsulta = "Simples - " & cboConsultaCriterios.Text
   Select Case cboConsultaCriterios.Text
      Case "CLIENTE"
         sCriterio = txtCodClienteLocalizar.Text
      Case "CÓD. OS"
         sCriterio = cboLocalizar.Text
      Case "DATA"
         sCriterio = mskDataConsulta.Text
      Case "PERÍODO"
         sCriterio = mskPeriodoInicio.Text & " a " & mskPeriodoFim.Text
      Case "MENSAL"
         sCriterio = cboMesConsulta.Text & "/" & cboAnoConsulta.Text
      Case Else
         sCriterio = "TODOS"
   End Select
End If

sFiltrosAdic = "Status: " & cboConsultaStatus.Text & "  |  Situação: " & cboConsultaMostrar.Text & "  |  Pagamento: " & cboTipo.Text & "  |  Tipo Serviço: " & cboTipoServico.Text

REL_OS_Consulta.ReportField9.Visible = True
REL_OS_Consulta.ReportField10.Visible = True
REL_OS_Consulta.rfCons1.Visible = True
REL_OS_Consulta.rfCons1.Caption = sTipoConsulta
REL_OS_Consulta.rfCons2.Caption = sCriterio
REL_OS_Consulta.rfCons3.Caption = sFiltrosAdic

REL_OS_Consulta.Relatorio.NomeImpressora = var_Impressora
REL_OS_Consulta.Relatorio.Ativar
Unload REL_OS_Consulta

Me.Show 1
End Sub

Private Sub Grid_DblClick()
lCodOSSelecionado = Val(Grid.TextMatrix(Grid.Row, 0))
Me.Hide
End Sub

Private Sub MostrarGrid_OS()
Dim totalRegistros As Long

Dim SITUACAO As String
Dim var_STATUS As String
Dim INDICE As String
Dim varTIPO_OS As String

'campo de data usado nos filtros DATA/PERÍODO/MENSAL
Dim campoData As String
If optDataEntrada.Value = True Then
   campoData = "os.DATA_ENTRADA"
Else
   campoData = "os.DATA_TERMINO"
End If

'indice
If cboIndice.Text = "CÓD. OS" Then
   INDICE = "os.COD_OS DESC "
ElseIf cboIndice.Text = "CLIENTE" Then
   INDICE = "cliente.nome DESC "
ElseIf cboIndice.Text = "DATA" Then
   INDICE = campoData & " DESC "
Else
   INDICE = "OS.COD_OS DESC "
End If

'tipo de serviço
If cboTipoServico.Text = "TODOS" Then
   varTIPO_OS = " (os.tipo_os <> 'TODOS') "
ElseIf cboTipoServico.Text = "CONSERTO" Then
   varTIPO_OS = " (os.tipo_os = 'CONSERTO') "
ElseIf cboTipoServico.Text = "MONTAGEM" Then
   varTIPO_OS = " (os.tipo_os = 'MONTAGEM') "
ElseIf cboTipoServico.Text = "ATENDIMENTO" Then
   varTIPO_OS = " (os.tipo_os = 'ATENDIMENTO') "
ElseIf cboTipoServico.Text = "AUTOMAÇÃO" Then
   varTIPO_OS = " (os.tipo_os = 'AUTOMAÇÃO') "
ElseIf cboTipoServico.Text = "CONSULTORIA" Then
   varTIPO_OS = " (os.tipo_os = 'CONSULTORIA') "
ElseIf cboTipoServico.Text = "GARANTIA" Then
   varTIPO_OS = " (os.tipo_os = 'GARANTIA') "
ElseIf cboTipoServico.Text = "ORÇAMENTO" Then
   varTIPO_OS = " (os.tipo_os = 'ORÇAMENTO') "
Else
   varTIPO_OS = " (os.tipo_os <> 'TODOS') "
End If

'Status
If cboConsultaStatus.Text = "TODOS" Then
   SITUACAO = ""
ElseIf cboConsultaStatus.Text = "À COMEÇAR" Then
   SITUACAO = "AND (os.status = 'À COMEÇAR') "
ElseIf cboConsultaStatus.Text = "EM EXECUÇÃO" Then
   SITUACAO = "AND (os.status = 'EM EXECUÇÃO') "
ElseIf cboConsultaStatus.Text = "AGUARDANDO" Then
   SITUACAO = "AND (os.status = 'AGUARDANDO') "
ElseIf cboConsultaStatus.Text = "TERMINADO" Then
   SITUACAO = "AND (os.status = 'TERMINADO') "
End If

'Situação
If cboConsultaMostrar.Text = "TODOS" Then
   var_STATUS = ""
ElseIf cboConsultaMostrar.Text = "ABERTOS" Then
   var_STATUS = "AND (status_os = 0) "
ElseIf cboConsultaMostrar.Text = "FECHADOS" Then
   var_STATUS = "AND (status_os = 1) "
End If


'forma de pagamento (TODOS/À VISTA/À PRAZO)
Dim varTipoPagamento As String
If cboTipo.Text = "À VISTA" Then
   varTipoPagamento = "AND (os.tipo_pagamento = 'À Vista') "
ElseIf cboTipo.Text = "À PRAZO" Then
   varTipoPagamento = "AND (os.tipo_pagamento = 'À Prazo') "
Else
   varTipoPagamento = ""
End If

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
    If cboConsultaCriterios.Text = "CLIENTE" Then
       If txtCodClienteLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento_Auto.fabricante,'') + ' / ' + ISNULL(OS_Equipamento_Auto.modelo,'') + ' / ' + ISNULL(CAST(OS_Equipamento_Auto.ano AS VARCHAR(10)),'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS WHERE " & varTIPO_OS & " and (os.cod_cliente = " & txtCodClienteLocalizar.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then
       If cboLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento_Auto.fabricante,'') + ' / ' + ISNULL(OS_Equipamento_Auto.modelo,'') + ' / ' + ISNULL(CAST(OS_Equipamento_Auto.ano AS VARCHAR(10)),'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS WHERE " & varTIPO_OS & " and (os.cod_os = " & cboLocalizar.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "DATA" Then
       If Not IsDate(mskDataConsulta.Text) Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento_Auto.fabricante,'') + ' / ' + ISNULL(OS_Equipamento_Auto.modelo,'') + ' / ' + ISNULL(CAST(OS_Equipamento_Auto.ano AS VARCHAR(10)),'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS WHERE " & varTIPO_OS & " and (" & campoData & " >= CONVERT(DATETIME, '" & Format(mskDataConsulta.Text, ocDATA) & "', 103)) and (" & campoData & " < DATEADD(day, 1, CONVERT(DATETIME, '" & Format(mskDataConsulta.Text, ocDATA) & "', 103))) " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "PERÍODO" Then
       If Not IsDate(mskPeriodoInicio.Text) Or Not IsDate(mskPeriodoFim.Text) Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento_Auto.fabricante,'') + ' / ' + ISNULL(OS_Equipamento_Auto.modelo,'') + ' / ' + ISNULL(CAST(OS_Equipamento_Auto.ano AS VARCHAR(10)),'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS WHERE " & varTIPO_OS & " and (" & campoData & " >= CONVERT(DATETIME, '" & Format(mskPeriodoInicio.Text, ocDATA) & "', 103)) and (" & campoData & " < DATEADD(day, 1, CONVERT(DATETIME, '" & Format(mskPeriodoFim.Text, ocDATA) & "', 103))) " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "MENSAL" Then
       If cboMesConsulta.Text = "" Or cboAnoConsulta.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento_Auto.fabricante,'') + ' / ' + ISNULL(OS_Equipamento_Auto.modelo,'') + ' / ' + ISNULL(CAST(OS_Equipamento_Auto.ano AS VARCHAR(10)),'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS WHERE " & varTIPO_OS & " and (MONTH(" & campoData & ") = " & (cboMesConsulta.ListIndex + 1) & ") and (YEAR(" & campoData & ") = " & cboAnoConsulta.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    Else
        sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento_Auto.fabricante,'') + ' / ' + ISNULL(OS_Equipamento_Auto.modelo,'') + ' / ' + ISNULL(CAST(OS_Equipamento_Auto.ano AS VARCHAR(10)),'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
            "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _
            "WHERE " & varTIPO_OS & " " & SITUACAO & var_STATUS & _
            varTipoPagamento & "ORDER BY " & INDICE
    End If
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    If cboConsultaCriterios.Text = "CLIENTE" Then
       If txtCodClienteLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (os.cod_cliente = " & txtCodClienteLocalizar.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then
       If cboLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (os.cod_os = " & cboLocalizar.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "DATA" Then
       If Not IsDate(mskDataConsulta.Text) Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (" & campoData & " >= CONVERT(DATETIME, '" & Format(mskDataConsulta.Text, ocDATA) & "', 103)) and (" & campoData & " < DATEADD(day, 1, CONVERT(DATETIME, '" & Format(mskDataConsulta.Text, ocDATA) & "', 103))) " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "PERÍODO" Then
       If Not IsDate(mskPeriodoInicio.Text) Or Not IsDate(mskPeriodoFim.Text) Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (" & campoData & " >= CONVERT(DATETIME, '" & Format(mskPeriodoInicio.Text, ocDATA) & "', 103)) and (" & campoData & " < DATEADD(day, 1, CONVERT(DATETIME, '" & Format(mskPeriodoFim.Text, ocDATA) & "', 103))) " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "MENSAL" Then
       If cboMesConsulta.Text = "" Or cboAnoConsulta.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (MONTH(" & campoData & ") = " & (cboMesConsulta.ListIndex + 1) & ") and (YEAR(" & campoData & ") = " & cboAnoConsulta.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    Else
        sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
            "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS " & _
            "WHERE " & varTIPO_OS & " " & SITUACAO & var_STATUS & _
            varTipoPagamento & "ORDER BY " & INDICE
    End If
ElseIf vTipoOS = "Comunicação Visual" Then
    If cboConsultaCriterios.Text = "CLIENTE" Then
       If txtCodClienteLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (os.cod_cliente = " & txtCodClienteLocalizar.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then
       If cboLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (os.cod_os = " & cboLocalizar.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "DATA" Then
       If Not IsDate(mskDataConsulta.Text) Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (" & campoData & " >= CONVERT(DATETIME, '" & Format(mskDataConsulta.Text, ocDATA) & "', 103)) and (" & campoData & " < DATEADD(day, 1, CONVERT(DATETIME, '" & Format(mskDataConsulta.Text, ocDATA) & "', 103))) " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "PERÍODO" Then
       If Not IsDate(mskPeriodoInicio.Text) Or Not IsDate(mskPeriodoFim.Text) Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (" & campoData & " >= CONVERT(DATETIME, '" & Format(mskPeriodoInicio.Text, ocDATA) & "', 103)) and (" & campoData & " < DATEADD(day, 1, CONVERT(DATETIME, '" & Format(mskPeriodoFim.Text, ocDATA) & "', 103))) " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    ElseIf cboConsultaCriterios.Text = "MENSAL" Then
       If cboMesConsulta.Text = "" Or cboAnoConsulta.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (MONTH(" & campoData & ") = " & (cboMesConsulta.ListIndex + 1) & ") and (YEAR(" & campoData & ") = " & cboAnoConsulta.Text & ") " & _
          SITUACAO & var_STATUS & varTipoPagamento & "ORDER BY " & INDICE
    Else
        sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, (cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL " & _
            "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS " & _
            "WHERE " & varTIPO_OS & " " & SITUACAO & var_STATUS & _
            varTipoPagamento & "ORDER BY " & INDICE
    End If
Else
    FormatarGrid_OS Nothing
    lblQtda.Caption = Format(0, "000")
    Exit Sub
End If
'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid_OS r

printSQL = sSQL

lblQtda.Caption = Format(totalRegistros, "000")

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGrid_OS(rTabela As ADODB.Recordset)
Dim i As Integer
Dim aCor As ColorConstants
Dim totalRegistros As Long

With Grid
   .Rows = 1       'INICIA O GRID COM UMA LINHA
   .FixedCols = 0  'DETERMINA QUE NÃO HAJA COLUNA FIXA
   
   'Abaixo o cabeçalho é criado
   .FormatString = "^CÓD.|^ENTRADA|^TERMINO|^TECNICO|^FINANC.|^CLIENTE|^TIPO|^FORMA|^VALOR|^DESC.|^TOTAL|"
   .ColWidth(0) = 600
   .ColWidth(1) = 1000
   .ColWidth(2) = 1000
   .ColWidth(3) = 1100
   .ColWidth(4) = 900
   .ColWidth(5) = 5350
   .ColWidth(6) = 750
   .ColWidth(7) = 1050
   .ColWidth(8) = 850
   .ColWidth(9) = 650
   .ColWidth(10) = 850
   .ColWidth(11) = 0
    
    'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   .Redraw = False
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         'ALINHAMENTO
         .ColAlignment(5) = 1
         .ColAlignment(8) = 0
         .ColAlignment(7) = 0
         .ColAlignment(8) = 6
         .ColAlignment(9) = 6
         .ColAlignment(10) = 6
         
         'A linha abaixo cria mais linha no grid
         .Rows = .Rows + 1
         
         'Preenche com os dados, e assim sucessivamente
         .TextMatrix(.Rows - 1, 0) = Format(rTabela("cod_os"), "0000")
         If IsNull(rTabela("DATA_ENTRADA")) Then
            .TextMatrix(.Rows - 1, 1) = ""
         Else
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("DATA_ENTRADA"), "dd/mm/yy")
         End If
         If IsNull(rTabela("DATA_TERMINO")) Then
            .TextMatrix(.Rows - 1, 2) = ""
         ElseIf rTabela("var_status") = "TERMINADO" Then
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("DATA_TERMINO"), "dd/mm/yy")
         ElseIf optDataPrevisao.Value = True Then
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("DATA_TERMINO"), "dd/mm/yy") & "(P)"
         Else
            .TextMatrix(.Rows - 1, 2) = ""
         End If
         .TextMatrix(.Rows - 1, 3) = rTabela("var_status")
         .TextMatrix(.Rows - 1, 4) = rTabela("var_status_os") & ""
         
         If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo")) & " / " & ValidateNull(rTabela("ano"))
         ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("equipamento")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo"))
         ElseIf vTipoOS = "Comunicação Visual" Then
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("equipamento")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo"))
         End If
         .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("TIPO_PAGAMENTO"))
         .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("PAGAMENTO"))
         .TextMatrix(.Rows - 1, 8) = Format(rTabela("SUBTOTAL"), ocMONEY)
         .TextMatrix(.Rows - 1, 9) = Format(rTabela("ValorDescReal"), ocMONEY)
         .TextMatrix(.Rows - 1, 10) = Format(rTabela("TOTAL"), ocMONEY)
         .TextMatrix(.Rows - 1, 11) = ValidateNull(rTabela("cod_pedido"))
         rTabela.MoveNext
      Loop
   End If
   
   'agora sim coloco a fução para mudar a cor da coluna e pronto
   'mudar a cor da fonte
   For i = 1 To .Rows - 1
      If UCase(Trim(.TextMatrix(i, 4))) = UCase("ABERTO") Then
         aCor = vbBlue
      Else
         aCor = vbRed
      End If
      
      .Col = 4 'a coluna do aberto ou fechado
      .Row = i
      .CellForeColor = aCor
   Next
   
   'mudar a cor da fonte
   For i = 1 To .Rows - 1
      If UCase(Trim(.TextMatrix(i, 3))) = UCase("À COMEÇAR") Then
         aCor = vbBlack
      ElseIf UCase(Trim(.TextMatrix(i, 3))) = UCase("EM EXECUÇÃO") Then
         aCor = RGB(0, 100, 0)
      ElseIf UCase(Trim(.TextMatrix(i, 3))) = UCase("AGUARDANDO") Then
         aCor = vbBlue
      ElseIf UCase(Trim(.TextMatrix(i, 3))) = UCase("TERMINADO") Then
         aCor = vbRed
      End If
      
      .Col = 3 'a coluna do aberto ou fechado
      .Row = i
      .CellForeColor = aCor
   Next
   
   'colunas TERMINO e FINANC. com fundo cinza claro
   For i = 0 To .Rows - 1
      .Row = i
      .Col = 2
      .CellBackColor = &HE0E0E0
      .Col = 4
      .CellBackColor = &HE0E0E0
   Next

   .Redraw = True
End With

lblTotal.Caption = Format(SomaGrid(Grid, 10), ocMONEY)
End Sub

Private Sub Preencher_Criterios()
cboConsultaCriterios.Clear
cboConsultaCriterios.AddItem "TODOS"
cboConsultaCriterios.AddItem "CÓD. OS"
cboConsultaCriterios.AddItem "CLIENTE"
cboConsultaCriterios.AddItem "DATA"
cboConsultaCriterios.AddItem "PERÍODO"
cboConsultaCriterios.AddItem "MENSAL"
End Sub

Private Sub Preencher_Indice()
   cboIndice.Clear
   cboIndice.AddItem "CÓD. OS"
   cboIndice.AddItem "CLIENTE"
   cboIndice.AddItem "DATA"
End Sub

Private Sub Preencher_Mostrar()
cboConsultaMostrar.Clear
cboConsultaMostrar.AddItem "TODOS"
cboConsultaMostrar.AddItem "ABERTOS"
cboConsultaMostrar.AddItem "FECHADOS"
End Sub

Private Sub Preencher_Status()
cboConsultaStatus.Clear
cboConsultaStatus.AddItem "TODOS"
cboConsultaStatus.AddItem "À COMEÇAR"
cboConsultaStatus.AddItem "EM EXECUÇÃO"
cboConsultaStatus.AddItem "AGUARDANDO"
cboConsultaStatus.AddItem "TERMINADO"
End Sub

Private Sub Preencher_TipoServico()
cboTipoServico.Clear
cboTipoServico.AddItem "TODOS"
cboTipoServico.AddItem "CONSERTO"
cboTipoServico.AddItem "GARANTIA"
cboTipoServico.AddItem "ORÇAMENTO"
End Sub

Private Sub Preencher_Tipo()
cboTipo.Clear
cboTipo.AddItem "TODOS"
cboTipo.AddItem "À VISTA"
cboTipo.AddItem "À PRAZO"
End Sub

Private Sub cboTipo_GotFocus()
Dim itemAtual As String
itemAtual = cboTipo.Text
Preencher_Tipo
cboTipo.Text = itemAtual
moCombo.AttachTo cboTipo
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
Dim i As Integer
Dim Valor As Currency

Valor = 0
For i = 0 To var_Grid.Rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function

