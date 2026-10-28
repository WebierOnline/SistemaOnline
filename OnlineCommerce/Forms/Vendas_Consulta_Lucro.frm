VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Vendas_Consulta_Lucro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE LUCRO"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "Vendas_Consulta_Lucro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   60
      ScaleHeight     =   2265
      ScaleWidth      =   15105
      TabIndex        =   8
      ToolTipText     =   "Imprimir"
      Top             =   1080
      Width           =   15135
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Critérios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2115
         Left            =   120
         TabIndex        =   25
         Top             =   60
         Width           =   2535
         Begin VB.ComboBox cboIndice2 
            Height          =   315
            Left            =   1620
            TabIndex        =   34
            Top             =   1680
            Width           =   855
         End
         Begin VB.ComboBox cboIndice 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox cboCriterioSec 
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   2355
         End
         Begin VB.ComboBox cboCriterioPrinc 
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Criterio"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Consultar por:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Escolha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2115
         Left            =   2700
         TabIndex        =   9
         Top             =   60
         Width           =   12375
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   6480
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   2355
         End
         Begin ChamaleonBtn.chameleonButton cmdCalendario1 
            Height          =   315
            Left            =   1080
            TabIndex        =   23
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
            MICON           =   "Vendas_Consulta_Lucro.frx":23D2
            PICN            =   "Vendas_Consulta_Lucro.frx":23EE
            PICH            =   "Vendas_Consulta_Lucro.frx":4741
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
            Left            =   2820
            TabIndex        =   24
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
            MICON           =   "Vendas_Consulta_Lucro.frx":6A94
            PICN            =   "Vendas_Consulta_Lucro.frx":6AB0
            PICH            =   "Vendas_Consulta_Lucro.frx":8E03
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.ComboBox cboDescricao 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Visible         =   0   'False
            Width           =   5235
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1500
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   1155
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFim 
            Height          =   315
            Left            =   1860
            TabIndex        =   16
            Top             =   480
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizar 
            Height          =   495
            Left            =   120
            TabIndex        =   33
            Top             =   1500
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&Exibir"
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
            MICON           =   "Vendas_Consulta_Lucro.frx":B156
            PICN            =   "Vendas_Consulta_Lucro.frx":B172
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo"
            Height          =   195
            Left            =   6480
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblDescricao 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   795
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano:"
            Height          =   195
            Left            =   1500
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lblInicio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data inicial:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lblFim 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data final:"
            Height          =   195
            Left            =   1860
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblAte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "até"
            Height          =   195
            Left            =   1500
            TabIndex        =   17
            Top             =   540
            Visible         =   0   'False
            Width           =   225
         End
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4380
      Picture         =   "Vendas_Consulta_Lucro.frx":BA4C
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   2
      Top             =   6480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   15105
      TabIndex        =   0
      Top             =   60
      Width           =   15135
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE LUCRO POR PRODUTO AGRUPADOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   7920
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Vendas_Consulta_Lucro.frx":CA84
         Top             =   0
         Width           =   1140
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   10305
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22595
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "07:51"
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5295
      Left            =   60
      TabIndex        =   10
      Top             =   3480
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   9340
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   255
      Left            =   2400
      TabIndex        =   32
      Top             =   8820
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
      MICON           =   "Vendas_Consulta_Lucro.frx":132CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirDetalhes 
      Height          =   255
      Left            =   60
      TabIndex        =   37
      Top             =   8820
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Vendas_Consulta_Lucro.frx":132E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTO:"
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
      Left            =   12600
      TabIndex        =   42
      Top             =   9540
      Width           =   705
   End
   Begin VB.Label lblCusto 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13440
      TabIndex        =   41
      Top             =   9540
      Width           =   1635
   End
   Begin VB.Label lblVenda 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13440
      TabIndex        =   40
      Top             =   9240
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENDAS:"
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
      Left            =   12600
      TabIndex        =   39
      Top             =   9240
      Width           =   825
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   315
      Left            =   5760
      TabIndex        =   38
      Top             =   9540
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LUCROS:"
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
      Left            =   12600
      TabIndex        =   6
      Top             =   9840
      Width           =   825
   End
   Begin VB.Label lblLucro 
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
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   13440
      TabIndex        =   5
      Top             =   9840
      Width           =   1635
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13440
      TabIndex        =   4
      Top             =   8940
      Width           =   1635
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
      Left            =   12600
      TabIndex        =   3
      Top             =   8940
      Width           =   780
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1395
      Left            =   12480
      Top             =   8820
      Width           =   2715
   End
End
Attribute VB_Name = "Vendas_Consulta_Lucro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

Dim posX As Single

Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer

Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Private Sub FormatarGrid_ProdDetalhado(rTabela As ADODB.Recordset)
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 10
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 765
      .ColWidth(2) = 1330
      .ColWidth(3) = 6270
      .ColWidth(4) = 1350
      .ColWidth(5) = 765
      .ColWidth(6) = 1020
      .ColWidth(7) = 969
      .ColWidth(8) = 969
      .ColWidth(9) = 1200
      
      .TextMatrix(0, 1) = "CÓD."
      .TextMatrix(0, 2) = "CÓD. BARRA"
      .TextMatrix(0, 3) = "DESCRIÇÃO"
      .TextMatrix(0, 4) = "FABRICANTE"
      .TextMatrix(0, 5) = "QTDE"
      .TextMatrix(0, 6) = "VENDAS"
      .TextMatrix(0, 7) = "CUSTO"
      .TextMatrix(0, 8) = "LUCRO"
      .TextMatrix(0, 9) = "MARGEM %"
      
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      .ColAlignment(1) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then

         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("cod_produto"), "000000")
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("COD_BARRA"))
            .TextMatrix(.rows - 1, 3) = rTabela("descricao")
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("fabricante"))
            .TextMatrix(.rows - 1, 5) = rTabela("vsomaquant")
            .TextMatrix(.rows - 1, 6) = FormatNumber(rTabela("vSomaTOTAL"), 2)
            .TextMatrix(.rows - 1, 7) = FormatNumber(rTabela("vSomaCUSTO"), 2)
            .TextMatrix(.rows - 1, 8) = FormatNumber(rTabela("vSomaLUCRO"), 2)
            .TextMatrix(.rows - 1, 9) = FormatNumber(rTabela("vSomaMARGEM"), 2)
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 8
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   lblVenda.Caption = FormatNumber(SomaGrid(Grid, 6), 2)
   lblCusto.Caption = FormatNumber(SomaGrid(Grid, 7), 2)
   lblQtda.Caption = SomaGrid(Grid, 5)
   lblLucro.Caption = FormatNumber(SomaGrid(Grid, 8), 2)
picAguarde.Visible = False
End Sub

Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
mskFim.Mask = ""
mskFim.Text = ""
'mskInicio.Mask = ""
'mskInicio.Text = ""
End Sub

Private Sub PreencherCriterio()
'"TODOS" removido - cboCriterioPrinc_GotFocus recarrega a lista so com
'MENSAL/DATA/PERIODO a cada foco (a opcao sumia depois do 1o clique) e
'cmdLocalizar_Click nunca tratou "TODOS" mesmo quando selecionada
cboCriterioPrinc.AddItem "MENSAL"
End Sub

Private Sub PreencherCriterioSec()
cboCriterioSec.AddItem "TODOS"
cboCriterioSec.AddItem "DESCRIÇÃO"
cboCriterioSec.AddItem "CÓD. BARRA"
cboCriterioSec.AddItem "FABRICANTE"
End Sub

Private Sub PreencherIndice()
cboIndice.AddItem "PRODUTO"
cboIndice.AddItem "QUANT."
cboIndice.AddItem "TOTAL"
End Sub

Private Sub PreencherTipoConsulta()
cboTipo.AddItem "POR PRODUTOS"
cboTipo.AddItem "POR SERVIÇOS"
End Sub

Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAno.AddItem i
Next

moCombo.AttachTo cboAno
End Sub

Private Sub cboAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboCriterioPrinc_Click()
cboCriterioPrinc_LostFocus
End Sub

Private Sub cboCriterioPrinc_GotFocus()
cboCriterioPrinc.Clear

'cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "MENSAL"
cboCriterioPrinc.AddItem "DATA"
cboCriterioPrinc.AddItem "PERÍODO"

moCombo.AttachTo cboCriterioPrinc
End Sub

Private Sub cboCriterioPrinc_LostFocus()
If cboCriterioPrinc.Text = "TODOS" Then
    lblInicio.Visible = False
    mskInicio.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
    lblInicio.Visible = False
    mskInicio.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
    lblMes.Visible = True
    cboMes.Visible = True
    lblAno.Visible = True
    cboAno.Visible = True
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
    lblInicio.Visible = True
    lblInicio.Caption = "Data Inicio"
    mskInicio.Visible = True
    mskInicio.Text = Format(Date, "dd/mm/yy")
    lblFim.Visible = True
    mskFim.Visible = True
    lblAte.Visible = True
    cmdCalendario1.Visible = True
    cmdCalendario2.Visible = True
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
ElseIf cboCriterioPrinc.Text = "DATA" Then
    lblInicio.Visible = True
    lblInicio.Caption = "Data"
    mskInicio.Visible = True
    mskInicio.Text = Format(Date, "dd/mm/yy")
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = True
    cmdCalendario2.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
End If

If cboCriterioSec.Text = "DESCRIÇÃO" Or cboCriterioSec.Text = "REFERÊNCIA" Or cboCriterioSec.Text = "FABRICANTE" Then
    If cboCriterioSec.Text = "DESCRIÇÃO" Then
        lblDescricao.Caption = "Descrição"
    ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
        lblDescricao.Caption = "Referência"
    ElseIf cboCriterioSec.Text = "FABRICANTE" Then
        lblDescricao.Caption = "Fabricante"
    End If
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    txtCodBarra.Visible = False
ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
    lblDescricao.Caption = "Cód. Barra"
    lblDescricao.Visible = True
    cboDescricao.Visible = False
    txtCodBarra.Visible = True
Else
End If


LimparObjetos_Consulta
End Sub


Private Sub cboCriterioSec_Click()
cboCriterioSec_LostFocus
End Sub

Private Sub cboCriterioSec_GotFocus()
cboCriterioSec.Clear
cboCriterioSec.AddItem "TODOS"
cboCriterioSec.AddItem "DESCRIÇÃO"
cboCriterioSec.AddItem "CÓD. BARRA"
cboCriterioSec.AddItem "CÓDIGO"
cboCriterioSec.AddItem "FABRICANTE"
moCombo.AttachTo cboCriterioSec
End Sub

Private Sub cboCriterioSec_LostFocus()
If cboCriterioSec.Text = "TODOS" Then
    lblDescricao.Visible = False
    cboDescricao.Visible = False
    txtCodBarra.Visible = False
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Or cboCriterioSec.Text = "FABRICANTE" Then
    If cboCriterioSec.Text = "DESCRIÇÃO" Then
        lblDescricao.Caption = "Descrição"
    ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
        lblDescricao.Caption = "Referência"
    ElseIf cboCriterioSec.Text = "FABRICANTE" Then
        lblDescricao.Caption = "Fabricante"
    End If
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    cboDescricao.Text = ""
    txtCodBarra.Visible = False
ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
    lblDescricao.Caption = "Cód. Barra"
    lblDescricao.Visible = True
    cboDescricao.Visible = False
    txtCodBarra.Visible = True
    txtCodBarra.Text = ""
ElseIf cboCriterioSec.Text = "CÓDIGO" Then
    lblDescricao.Caption = "Código do produto"
    lblDescricao.Visible = True
    cboDescricao.Visible = False
    txtCodBarra.Visible = True
    txtCodBarra.Text = ""
Else
End If
End Sub


Private Sub cboDescricao_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboDescricao.Clear
   
If cboCriterioSec.Text = "DESCRIÇÃO" Then
   sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem r("descricao")
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
   sSQL = "SELECT DISTINCT REF FROM produtos ORDER BY REF;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem r("REF")
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "FABRICANTE" Then
   sSQL = "SELECT DISTINCT FABRICANTE FROM produtos ORDER BY FABRICANTE;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem ValidateNull(r("FABRICANTE"))
      r.MoveNext
   Loop
Else
   Exit Sub
End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboDescricao
End Sub


Private Sub cboIndice_GotFocus()
cboIndice.Clear
cboIndice.AddItem "PRODUTO"
cboIndice.AddItem "QUANT."
cboIndice.AddItem "VENDA"
cboIndice.AddItem "LUCRO"
cboIndice.AddItem "MARGEM"
moCombo.AttachTo cboIndice
End Sub

Private Sub cboIndice2_GotFocus()
cboIndice2.Clear
cboIndice2.AddItem "ASC"
cboIndice2.AddItem "DESC"
moCombo.AttachTo cboIndice2
End Sub


Private Sub cboMes_GotFocus()
cboMes.Clear

cboMes.AddItem "Janeiro"
cboMes.AddItem "Fevereiro"
cboMes.AddItem "Março"
cboMes.AddItem "Abril"
cboMes.AddItem "Maio"
cboMes.AddItem "Junho"
cboMes.AddItem "Julho"
cboMes.AddItem "Agosto"
cboMes.AddItem "Setembro"
cboMes.AddItem "Outubro"
cboMes.AddItem "Novembro"
cboMes.AddItem "Dezembro"

moCombo.AttachTo cboMes
End Sub

Private Sub cboMes_LostFocus()
   cboAno.SetFocus
End Sub

Private Sub cboTipo_Change()
If cboTipo.Text = "POR PRODUTOS" Then
   'cmdExibirPedidos.Visible = True
ElseIf cboTipo.Text = "POR SERVIÇOS" Then
   'cmdExibirPedidos.Visible = True
Else
   Exit Sub
End If
End Sub

Private Sub cboTipo_Click()
cboTipo_Change
End Sub

Private Sub cboTipo_GotFocus()
cboTipo.Clear
cboTipo.AddItem "POR PRODUTOS"
cboTipo.AddItem "POR SERVIÇOS"
moCombo.AttachTo cboTipo
End Sub

Private Sub chameleonButton1_Click()

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


Private Sub cmdExibirDetalhes_Click()
If Grid.Col = 0 Then Exit Sub

If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
   If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
    If cboCriterioPrinc.Text = "MENSAL" Then
        If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
        Vendas_Consulta_PorLucro.loadPedidos Grid.TextMatrix(Grid.Row, 1), cboMes.ListIndex + 1, cboAno
    ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
        If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
        Vendas_Consulta_PorLucro.loadPedidos3 Grid.TextMatrix(Grid.Row, 1), mskInicio.Text, mskFim.Text
    ElseIf cboCriterioPrinc.Text = "DATA" Then
        If Not IsDate(mskInicio.Text) Then Exit Sub
        Vendas_Consulta_PorLucro.loadPedidos2 Grid.TextMatrix(Grid.Row, 1), mskInicio.Text
    End If
   
   Vendas_Consulta_PorLucro.Show 1
End If



End Sub

Public Sub cmdImprimir_Click()
Dim r As ADODB.Recordset

'Dim r As ADODB.Recordset
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = var_ImpNormal Then
      Set Printer = Prt
      Exit For
   End If
Next

Me.Hide    'ver depois como nao exibir

Set r = dbData.OpenRecordset(printSQL)
Set REL_Cons_Venda_PorLucro.Relatorio.Recordset = r

If cboCriterioPrinc.Text = "MENSAL" Then
    REL_Cons_Venda_PorLucro.rfCons1.Caption = "MENSAL"
    REL_Cons_Venda_PorLucro.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
    REL_Cons_Venda_PorLucro.rfCons1.Caption = "PERÍODO"
    REL_Cons_Venda_PorLucro.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " até " & mskFim.Text
ElseIf cboCriterioPrinc.Text = "DATA" Then
    REL_Cons_Venda_PorLucro.rfCons1.Caption = "DATA"
    REL_Cons_Venda_PorLucro.rfCons3.Caption = "DATA = " & mskInicio.Text
End If

If cboCriterioSec.Text = "DESCRIÇÃO" Then
    REL_Cons_Venda_PorLucro.rfCons2.Caption = "DESCRIÇÃO = " & cboDescricao.Text & ""
ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
    REL_Cons_Venda_PorLucro.rfCons2.Caption = "CÓD. BARRA = " & txtCodBarra.Text & ""
ElseIf cboCriterioSec.Text = "FABRICANTE" Then
    REL_Cons_Venda_PorLucro.rfCons2.Caption = "FABRICANTE = " & cboDescricao.Text & ""
End If

REL_Cons_Venda_PorLucro.dfQuant.Caption = lblQtda.Caption
REL_Cons_Venda_PorLucro.dfVenda.Caption = Format(lblVenda.Caption, "##,##0.00")
REL_Cons_Venda_PorLucro.dfCusto.Caption = Format(lblCusto.Caption, "##,##0.00")
REL_Cons_Venda_PorLucro.dfLucro.Caption = Format(lblLucro.Caption, "##,##0.00")

REL_Cons_Venda_PorLucro.Relatorio.NomeImpressora = var_ImpNormal
REL_Cons_Venda_PorLucro.Relatorio.Ativar
Unload REL_Cons_Venda_PorLucro

Me.Show 1   'ver depois como nao exibir
End Sub

Public Sub cmdLocalizar_Click()

totalRegistros = "0"

Dim INDICE As String 'INDICE
If cboIndice.Text = "QUANT." Then
   INDICE = "vSomaQuant "
ElseIf cboIndice.Text = "PRODUTO" Then
   INDICE = "produtos.descricao "
ElseIf cboIndice.Text = "VENDA" Then
   INDICE = "vSomaTOTAL "
ElseIf cboIndice.Text = "LUCRO" Then
   INDICE = "vSomaLUCRO "
ElseIf cboIndice.Text = "MARGEM" Then
   INDICE = "vSomaMARGEM "
Else
   INDICE = "produtos.descricao "
End If

Dim INDICE2 As String 'INDICE
If cboIndice2.Text = "ASC" Then
   INDICE2 = "ASC ;"
ElseIf cboIndice2.Text = "DESC" Then
   INDICE2 = "DESC ;"
Else
   INDICE2 = "ASC ;"
End If

Dim vCriterioSec As String
If cboCriterioSec.Text = "TODOS" Then
    vCriterioSec = ""
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
    vCriterioSec = "AND produtos.DESCRICAO = '" & cboDescricao.Text & "'"
ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
    vCriterioSec = "AND produtos.COD_BARRA = '" & txtCodBarra.Text & "'"
ElseIf cboCriterioSec.Text = "CÓDIGO" Then
    vCriterioSec = "AND produtos.CODIGO = " & txtCodBarra.Text & ""
ElseIf cboCriterioSec.Text = "FABRICANTE" Then
    vCriterioSec = "AND produtos.FABRICANTE = '" & cboDescricao.Text & "'"
End If

If cboCriterioPrinc.Text = "MENSAL" Then
    If cboMes.Text = "" Or cboAno.Text = "" Then
        ShowMsg "Selecione o mês e o ano antes de exibir.", vbInformation
        Exit Sub
    End If
    sSQL = "SELECT pedidos_itens.COD_PRODUTO, SUM(pedidos_itens.QUANTIDADE) AS vSomaQuant, SUM(pedidos_itens.Total) AS vSomaTOTAL, " & _
        "SUM(ISNULL(pedidos_itens.custo, 0) * pedidos_itens.QUANTIDADE) AS vSomaCUSTO, " & _
        "CASE WHEN SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) = 0 THEN 0 ELSE SUM(pedidos_itens.Total) - SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) END AS vSomaLUCRO, " & _
        "produtos.DESCRICAO, produtos.FABRICANTE, produtos.QUANT_ESTOQUE, produtos.COD_BARRA,  " & _
        " CASE WHEN SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) = 0 OR SUM(pedidos_itens.Total) = 0 THEN 0 ELSE (SUM(pedidos_itens.Total) - SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE)) / SUM(pedidos_itens.Total) * 100 END AS vSomaMARGEM " & _
        "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO INNER JOIN pedidos ON pedidos_itens.COD_PEDIDO = pedidos.COD_PEDIDO " & _
        "WHERE (MONTH(pedidos_itens.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens.data) = " & cboAno & ") AND (pedidos_itens.cancelado = 0) AND (pedidos.cancelado = 0) AND (pedidos.TIPO_PEDIDO NOT IN ('ORÇAMENTO', 'CONSIGNADO')) AND (pedidos.STATUS_PEDIDO = 1) " & vCriterioSec & " " & _
        "GROUP BY pedidos_itens.COD_PRODUTO, produtos.DESCRICAO, produtos.FABRICANTE, produtos.QUANT_ESTOQUE, produtos.COD_BARRA " & _
        "ORDER BY " & INDICE & " " & INDICE2 & " "
'Debug.Print sSQL

ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
    If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then
        ShowMsg "Informe as datas inicial e final antes de exibir.", vbInformation
        Exit Sub
    End If
   sSQL = "SELECT pedidos_itens.COD_PRODUTO, SUM(pedidos_itens.QUANTIDADE) AS vSomaQuant, SUM(pedidos_itens.Total) AS vSomaTOTAL, " & _
        "SUM(ISNULL(pedidos_itens.custo, 0) * pedidos_itens.QUANTIDADE) AS vSomaCUSTO, " & _
        "CASE WHEN SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) = 0 THEN 0 ELSE SUM(pedidos_itens.Total) - SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) END AS vSomaLUCRO, " & _
        "produtos.DESCRICAO, produtos.FABRICANTE, produtos.QUANT_ESTOQUE, produtos.COD_BARRA, " & _
        "CASE WHEN SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) = 0 OR SUM(pedidos_itens.Total) = 0 THEN 0 ELSE (SUM(pedidos_itens.Total) - SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE)) / SUM(pedidos_itens.Total) * 100 END AS vSomaMARGEM " & _
        "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO INNER JOIN pedidos ON pedidos_itens.COD_PEDIDO = pedidos.COD_PEDIDO " & _
        "WHERE (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_itens.cancelado = 0) AND (pedidos.cancelado = 0) AND (pedidos.TIPO_PEDIDO NOT IN ('ORÇAMENTO', 'CONSIGNADO')) AND (pedidos.STATUS_PEDIDO = 1) " & vCriterioSec & "  " & _
        "GROUP BY pedidos_itens.COD_PRODUTO, produtos.DESCRICAO, produtos.FABRICANTE, produtos.QUANT_ESTOQUE, produtos.COD_BARRA " & _
        "ORDER BY " & INDICE & " " & INDICE2 & " "
        
ElseIf cboCriterioPrinc.Text = "DATA" Then
    If Not IsDate(mskInicio.Text) Then
        ShowMsg "Informe a data antes de exibir.", vbInformation
        Exit Sub
    End If
   sSQL = "SELECT pedidos_itens.COD_PRODUTO, SUM(pedidos_itens.QUANTIDADE) AS vSomaQuant, SUM(pedidos_itens.Total) AS vSomaTOTAL, " & _
        "SUM(ISNULL(pedidos_itens.custo, 0) * pedidos_itens.QUANTIDADE) AS vSomaCUSTO, " & _
        "CASE WHEN SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) = 0 THEN 0 ELSE SUM(pedidos_itens.Total) - SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) END AS vSomaLUCRO, " & _
        "produtos.DESCRICAO, produtos.FABRICANTE, produtos.QUANT_ESTOQUE, produtos.COD_BARRA, " & _
        "CASE WHEN SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE) = 0 OR SUM(pedidos_itens.Total) = 0 THEN 0 ELSE (SUM(pedidos_itens.Total) - SUM(ISNULL(pedidos_itens.Custo, 0) * pedidos_itens.QUANTIDADE)) / SUM(pedidos_itens.Total) * 100 END AS vSomaMARGEM " & _
        "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO INNER JOIN pedidos ON pedidos_itens.COD_PEDIDO = pedidos.COD_PEDIDO " & _
        "WHERE (pedidos_itens.data = CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.cancelado = 0) AND (pedidos.cancelado = 0) AND (pedidos.TIPO_PEDIDO NOT IN ('ORÇAMENTO', 'CONSIGNADO')) AND (pedidos.STATUS_PEDIDO = 1) " & vCriterioSec & " " & _
        "GROUP BY pedidos_itens.COD_PRODUTO, produtos.DESCRICAO, produtos.FABRICANTE, produtos.QUANT_ESTOQUE, produtos.COD_BARRA " & _
        "ORDER BY " & INDICE & " " & INDICE2 & " "
End If
      
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

'Debug.Print sSQL

FormatarGrid_ProdDetalhado r
printSQL = sSQL

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Double
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Form_Load()

If vOrigemRelatorio = True Then Exit Sub

Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
   
'FORMATAR O GRID
With Grid
   .Clear
   .Cols = 7
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 0
   .ColWidth(3) = 0
   .ColWidth(4) = 0
   .ColWidth(5) = 0
   .ColWidth(6) = 0
End With

StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")

PreencherTipoConsulta
cboTipo.ListIndex = 0

PreencherCriterio
cboCriterioPrinc.ListIndex = 0

PreencherCriterioSec
cboCriterioSec.ListIndex = 0

PreencherIndice
cboIndice.ListIndex = 0

cboIndice2.Text = "ASC"

cboCriterioPrinc_LostFocus

Set moCombo = New cComboHelper

'pre-preenche mes/ano com o atual - criterio padrao e MENSAL, assim o
'usuario so precisa clicar Exibir sem escolher nada na maioria das vezes
cboMes_GotFocus
cboMes.ListIndex = Month(Date) - 1
cboAno_GotFocus
cboAno.ListIndex = 2   'lista e sempre (ano atual -2) a (ano atual +2), 5 itens - o atual fica no meio
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   'tooltip antigo era posicionado pro layout Pedido/Forma de outra tela
   '(copy-paste) e prometia um duplo-clique que nunca existiu aqui (nao ha
   'Grid_DblClick nesse form) - o drill-down real e selecionar a linha e
   'clicar em DETALHAMENTO
   Grid.ToolTipText = "Selecione o produto e clique em DETALHAMENTO para ver os pedidos."
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
   Exit Sub
Else
   If IsDate(mskFim.Text) Then
      cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskFim.SetFocus
      SelectControl mskFim
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
      Exit Sub
   Else
      If IsDate(mskInicio.Text) Then
         If mskFim.Visible = True Then mskFim.SetFocus
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskInicio.SetFocus
         SelectControl mskInicio
      End If
   End If
End Sub

