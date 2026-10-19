VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Ajuste_Consulta_PorProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HISTÓRICO DE AJUSTES"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   Icon            =   "Ajuste_Consulta_PorProdutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   60
      ScaleHeight     =   2205
      ScaleWidth      =   11385
      TabIndex        =   23
      ToolTipText     =   "Imprimir"
      Top             =   1080
      Width           =   11415
      Begin VB.Frame Frame1 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   60
         TabIndex        =   37
         Top             =   60
         Width           =   2295
         Begin VB.ComboBox cboIndice 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   1680
            Width           =   1995
         End
         Begin VB.ComboBox cboCriterioSec 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1995
         End
         Begin VB.ComboBox cboCriterioPrinc 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   1995
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   1480
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Consultar por:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   280
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Filtrar por:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   880
            Width           =   600
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2115
         Left            =   2400
         TabIndex        =   24
         Top             =   60
         Width           =   8955
         Begin ChamaleonBtn.chameleonButton cmdCalendario2 
            Height          =   315
            Left            =   2340
            TabIndex        =   11
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   1080
            Visible         =   0   'False
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
            MICON           =   "Ajuste_Consulta_PorProdutos.frx":23D2
            PICN            =   "Ajuste_Consulta_PorProdutos.frx":23EE
            PICH            =   "Ajuste_Consulta_PorProdutos.frx":4741
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCalendario1 
            Height          =   315
            Left            =   1020
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   1080
            Visible         =   0   'False
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
            MICON           =   "Ajuste_Consulta_PorProdutos.frx":6A94
            PICN            =   "Ajuste_Consulta_PorProdutos.frx":6AB0
            PICH            =   "Ajuste_Consulta_PorProdutos.frx":8E03
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
            Left            =   60
            TabIndex        =   1
            Top             =   480
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.ComboBox cboDescricao 
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   6015
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   60
            TabIndex        =   6
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1080
            Visible         =   0   'False
            Width           =   1155
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   1080
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFim 
            Height          =   315
            Left            =   1380
            TabIndex        =   10
            Top             =   1080
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizar 
            Height          =   495
            Left            =   60
            TabIndex        =   40
            Top             =   1560
            Width           =   1515
            _ExtentX        =   2672
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
            MICON           =   "Ajuste_Consulta_PorProdutos.frx":B156
            PICN            =   "Ajuste_Consulta_PorProdutos.frx":B172
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
            Height          =   495
            Left            =   1620
            TabIndex        =   41
            Top             =   1560
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&Imprimir"
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
            MICON           =   "Ajuste_Consulta_PorProdutos.frx":BA4C
            PICN            =   "Ajuste_Consulta_PorProdutos.frx":BA68
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboTipoAjuste 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   480
            Width           =   1700
         End
         Begin VB.Label lblTipoAjuste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   3000
            TabIndex        =   42
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblFim 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data final:"
            Height          =   195
            Left            =   1380
            TabIndex        =   31
            Top             =   840
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   60
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblInicio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data inicial:"
            Height          =   195
            Left            =   60
            TabIndex        =   29
            Top             =   840
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês:"
            Height          =   195
            Left            =   60
            TabIndex        =   28
            Top             =   840
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano:"
            Height          =   195
            Left            =   1440
            TabIndex        =   27
            Top             =   840
            Visible         =   0   'False
            Width           =   330
         End
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3960
      Picture         =   "Ajuste_Consulta_PorProdutos.frx":BD82
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   15
      Top             =   5160
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
      ScaleWidth      =   11385
      TabIndex        =   12
      Top             =   60
      Width           =   11415
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HISTÓRICO DE AJUSTE POR PRODUTOS"
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
         TabIndex        =   13
         Top             =   240
         Width           =   5895
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Ajuste_Consulta_PorProdutos.frx":CDBA
         Top             =   0
         Width           =   1140
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   22
      Top             =   9510
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16007
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 9427-5280"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 9427-5280"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "07:45"
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
      Height          =   4695
      Left            =   60
      TabIndex        =   26
      Top             =   3360
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8281
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Label lblEstoqueAtual 
      AutoSize        =   -1  'True
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   43
      Top             =   8100
      Width           =   150
   End
   Begin VB.Label lblQtdaAjuste 
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
      Left            =   8880
      TabIndex        =   36
      Top             =   9060
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblQtdaEntrada 
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
      Left            =   8880
      TabIndex        =   35
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label lblQtdaCadastro 
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
      Left            =   8880
      TabIndex        =   34
      Top             =   8460
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SAÍDA:"
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
      Left            =   8040
      TabIndex        =   33
      Top             =   9060
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblQtdaAjusteSoma 
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
      Left            =   9660
      TabIndex        =   32
      Top             =   9060
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REMOÇÃO:"
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
      Left            =   7860
      TabIndex        =   21
      Top             =   8760
      Width           =   960
   End
   Begin VB.Label lblQtdaCadastroSoma 
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
      Left            =   9660
      TabIndex        =   20
      Top             =   8460
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADIÇÃO:"
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
      Left            =   7740
      TabIndex        =   19
      Top             =   8460
      Width           =   1080
   End
   Begin VB.Label lblQtdaEntradaSoma 
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
      Left            =   9660
      TabIndex        =   18
      Top             =   8760
      Width           =   1755
   End
   Begin VB.Label lblQtdaTotal 
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
      Left            =   8880
      TabIndex        =   17
      Top             =   8160
      Width           =   2535
   End
   Begin VB.Label Label8 
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
      Left            =   8160
      TabIndex        =   16
      Top             =   8160
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1335
      Left            =   7620
      Top             =   8100
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   6180
      TabIndex        =   14
      Top             =   8220
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Ajuste_Consulta_PorProdutos"
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

Private Sub Limpar_Grid()
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
End Sub

Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
cboDescricao.Text = ""
mskFim.Mask = ""
mskFim.Text = ""
mskInicio.Mask = ""
mskInicio.Text = ""
End Sub

Private Sub PreencherConsulta()
cboCriterioSec.AddItem "DESCRIÇÃO"
cboCriterioSec.AddItem "CÓD. BARRA"
cboCriterioSec.AddItem "CÓD. PRODUTO"
End Sub

Private Sub PreencherCriterios()
cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "MENSAL"
cboCriterioPrinc.AddItem "PERÍODO"
End Sub

Private Sub PreencherIndice()
cboIndice.AddItem "DATA"
cboIndice.AddItem "QUANT."
cboIndice.AddItem "DESCRIÇÃO"
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
'Não espera perder o foco - exibe os objetos na hora que escolhe. SetFocus
'fica só aqui (não em LostFocus), mesmo motivo do cboCriterioSec: evita
'reentrância (erro 5)
AtualizarCriterioPrinc
DoEvents
If cboAno.Visible Then
   cboAno.SetFocus
ElseIf mskInicio.Visible Then
   mskInicio.SetFocus
End If
End Sub

Private Sub cboCriterioPrinc_GotFocus()
Dim vTexto As String
vTexto = cboCriterioPrinc.Text

cboCriterioPrinc.Clear
PreencherCriterios

cboCriterioPrinc.Text = vTexto
moCombo.AttachTo cboCriterioPrinc
End Sub


Private Sub cboCriterioPrinc_LostFocus()
AtualizarCriterioPrinc
End Sub

Private Sub AtualizarCriterioPrinc()
If cboCriterioPrinc.Text = "TODOS" Then
   lblMes.Visible = False
   lblAno.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
    lblInicio.Visible = False
    mskInicio.Visible = False
    cmdCalendario1.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    cmdCalendario2.Visible = False
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
    lblMes.Visible = True
    lblAno.Visible = True
    cboMes.Visible = True
    cboAno.Visible = True
    lblInicio.Visible = False
    mskInicio.Visible = False
    cmdCalendario1.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    cmdCalendario2.Visible = False
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
    lblMes.Visible = False
    lblAno.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
    lblInicio.Visible = True
    mskInicio.Visible = True
    cmdCalendario1.Visible = True
    lblFim.Visible = True
    mskFim.Visible = True
    cmdCalendario2.Visible = True
Else
End If
End Sub

Private Sub cboCriterioPrinc_Validate(Cancel As Boolean)
If cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
   lblMes.Visible = True
   cboMes.Visible = True
   lblAno.Visible = True
   cboAno.Visible = True
   cboDescricao.Visible = True
   lblDescricao.Visible = True
End If
End Sub

Private Sub cboCriterioSec_Click()
'Não espera perder o foco - exibe os objetos e dá foco na hora que escolhe.
'SetFocus fica só aqui (não em LostFocus) pra evitar reentrância: dar
'SetFocus tira o foco do combo, o que dispararia o LostFocus de verdade de
'novo, numa hora ruim (erro 5)
AtualizarCriterioSec
DoEvents
If txtCodBarra.Visible Then
   txtCodBarra.SetFocus
ElseIf cboDescricao.Visible Then
   cboDescricao.SetFocus
End If
End Sub


Private Sub cboCriterioSec_GotFocus()
Dim vTexto As String
vTexto = cboCriterioSec.Text

cboCriterioSec.Clear
'Só existe o modo "POR PRODUTOS" hoje (cboTipo foi removido - era
'o "POR SERVIÇOS", nunca implementado de verdade)
PreencherConsulta

cboCriterioSec.Text = vTexto

moCombo.AttachTo cboCriterioSec
End Sub

Private Sub cboCriterioSec_LostFocus()
AtualizarCriterioSec
End Sub

Private Sub AtualizarCriterioSec()
If cboCriterioSec.Text = "" Then cboCriterioSec.Text = "DESCRIÇÃO"

If cboCriterioSec.Text = "CÓD. BARRA" Then
   lblDescricao.Caption = "Cód. Barra"
   lblDescricao.Visible = True
   txtCodBarra.Visible = True
   txtCodBarra.Text = ""
   cboDescricao.Visible = False
   cboTipoAjuste.Left = 3060
   lblTipoAjuste.Left = 3060
ElseIf cboCriterioSec.Text = "CÓD. PRODUTO" Then
   lblDescricao.Caption = "Cód. Produto"
   lblDescricao.Visible = True
   txtCodBarra.Visible = True
   txtCodBarra.Text = ""
   cboDescricao.Visible = False
   cboTipoAjuste.Left = 3060
   lblTipoAjuste.Left = 3060
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
   lblDescricao.Caption = "Descrição"
   lblDescricao.Visible = True
   txtCodBarra.Visible = False
   cboDescricao.Visible = True
   cboTipoAjuste.Left = 6120
   lblTipoAjuste.Left = 6120
ElseIf cboCriterioSec.Text = "CATEGORIA" Then
   lblDescricao.Caption = "Categoria"
   lblDescricao.Visible = True
   txtCodBarra.Visible = False
   cboDescricao.Visible = True
   cboTipoAjuste.Left = 6120
   lblTipoAjuste.Left = 6120
End If
End Sub

Private Sub cboDescricao_GotFocus()
cboDescricao.Clear
   
If cboCriterioSec.Text = "DESCRIÇÃO" Then
   sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem r("descricao")
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "CATEGORIA" Then
   sSQL = "SELECT DISTINCT CATEGORIA FROM produtos ORDER BY CATEGORIA;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem ValidateNull(r("CATEGORIA"))
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
Dim vTexto As String
vTexto = cboIndice.Text

cboIndice.Clear
PreencherIndice

cboIndice.Text = vTexto

moCombo.AttachTo cboIndice
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


Private Sub cmdExibirParcelas_Click()
If Grid.Col = 0 Then Exit Sub
   If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
         Vendas_Consulta_Geral_Parcelas.loadInformacoes (Grid.TextMatrix(Grid.Row, 1))
         Vendas_Consulta_Geral_Parcelas.Show 1
   End If
End Sub

Private Sub cmdExibirProdutos_Click()
If Grid.Col = 0 Then Exit Sub
If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
   If Grid.Col = 1 Then
      If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), Grid.TextMatrix(Grid.Row, 7)
      Parcelas_Consulta_Produtos.Show 1
   End If
End If
End Sub

Private Sub cmdImprimir_Click()
'Essa tela ainda não tem um relatório próprio (REL_Entrada_PorProduto não
'bate com os campos de ajuste - ADIÇÃO/REMOÇÃO, sem Cadastro/Entrada). Fica
'pendente até criarmos um relatório dedicado.
ShowMsg "Impressão ainda não implementada pra essa tela.", vbInformation
End Sub

Public Sub cmdLocalizar_Click()
Dim INDICE As String
Dim sFiltro As String

totalRegistros = "0"

If cboCriterioPrinc.Text = "" Then Exit Sub
If cboCriterioSec.Text = "" Then Exit Sub

If cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "TODOS" Then
    If txtCodBarra.Text = "" Then Exit Sub
    sFiltro = "AND (Produtos.COD_BARRA = '" & txtCodBarra.Text & "') "
ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "MENSAL" Then
    If txtCodBarra.Text = "" Then Exit Sub
    If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
    sFiltro = "AND (Produtos.COD_BARRA = '" & txtCodBarra.Text & "') and (MONTH(Produtos_Quant.DATA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant.DATA) = " & cboAno & ") "
ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "PERÍODO" Then
    If txtCodBarra.Text = "" Then Exit Sub
    If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
    sFiltro = "AND (Produtos.COD_BARRA = '" & txtCodBarra.Text & "') and (Produtos_Quant.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
ElseIf cboCriterioSec.Text = "CÓD. PRODUTO" And cboCriterioPrinc.Text = "TODOS" Then
    If txtCodBarra.Text = "" Then Exit Sub
    sFiltro = "AND (Produtos.CODIGO = " & txtCodBarra.Text & ") "
ElseIf cboCriterioSec.Text = "CÓD. PRODUTO" And cboCriterioPrinc.Text = "MENSAL" Then
    If txtCodBarra.Text = "" Then Exit Sub
    If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
    sFiltro = "AND (Produtos.CODIGO = " & txtCodBarra.Text & ") and (MONTH(Produtos_Quant.DATA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant.DATA) = " & cboAno & ") "
ElseIf cboCriterioSec.Text = "CÓD. PRODUTO" And cboCriterioPrinc.Text = "PERÍODO" Then
    If txtCodBarra.Text = "" Then Exit Sub
    If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
    sFiltro = "AND (Produtos.CODIGO = " & txtCodBarra.Text & ") and (Produtos_Quant.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "TODOS" Then
    If cboDescricao.Text = "" Then Exit Sub
    sFiltro = "AND (Produtos.DESCRICAO = '" & cboDescricao.Text & "') "
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "MENSAL" Then
    If cboDescricao.Text = "" Then Exit Sub
    If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
    sFiltro = "AND (Produtos.DESCRICAO = '" & cboDescricao.Text & "') and (MONTH(Produtos_Quant.DATA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant.DATA) = " & cboAno & ") "
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "PERÍODO" Then
    If cboDescricao.Text = "" Then Exit Sub
    If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
    sFiltro = "AND (Produtos.DESCRICAO = '" & cboDescricao.Text & "') and (Produtos_Quant.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
Else
    Exit Sub
End If

'Filtro extra por TIPO (ADIÇÃO/REMOÇÃO/TODOS) - independente do filtro por produto
If cboTipoAjuste.Text = "ADIÇÃO" Then
    sFiltro = sFiltro & "AND (Produtos_Quant.TIPO = 'ADIÇÃO') "
ElseIf cboTipoAjuste.Text = "REMOÇÃO" Then
    sFiltro = sFiltro & "AND (Produtos_Quant.TIPO = 'REMOÇÃO') "
End If

If cboIndice.Text = "QUANT." Then
    INDICE = "ABS(Produtos_Quant.QUANT) DESC"
ElseIf cboIndice.Text = "DESCRIÇÃO" Then
    INDICE = "produtos.descricao"
ElseIf cboIndice.Text = "DATA" Then
    INDICE = "Produtos_Quant.DATA"
Else
    INDICE = "produtos.descricao"
End If

'HISTÓRICO DE AJUSTES = FORMA='AJUSTE' (adição e remoção) + a linha de
'FORMA='CADASTRO' como ponto de partida/referência. Sem entrada/saída/venda.
'QUANT sempre positivo (ABS) + TIPO indica a direção,
'igual à convenção já usada no Grid_Quant do Produtos_Cadastro.frm. Nome do
'usuário via funcionario (funcionario.CODIGO = usuario.Codigo, confirmado).
sSQL = "SELECT Produtos_Quant.codigo, Produtos_Quant.Hora, Produtos_Quant.DATA, " & _
   "ABS(Produtos_Quant.QUANT) AS QUANT, Produtos_Quant.TIPO, Produtos_Quant.FORMA, Produtos_Quant.Estoque, " & _
   "Produtos_Quant.COD_PRODUTO, funcionario.NOME AS varNomeUsuario, produtos.quant_estoque AS varEstoqueAtual " & _
   "FROM Produtos_Quant " & _
   "INNER JOIN produtos ON Produtos_Quant.COD_PRODUTO = produtos.CODIGO " & _
   "LEFT JOIN funcionario ON funcionario.CODIGO = Produtos_Quant.cod_usuario " & _
   "WHERE (Produtos_Quant.FORMA = 'AJUSTE' OR Produtos_Quant.FORMA = 'CADASTRO') " & sFiltro & "ORDER BY " & INDICE

Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If Not r.EOF Then
   lblEstoqueAtual.Caption = "ESTOQUE ATUAL: " & ValidateNull(r("varEstoqueAtual"))
Else
   lblEstoqueAtual.Caption = "ESTOQUE ATUAL: ?"
End If

FormatarGrid_Quant r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL

'Totais do rodapé: ADIÇÃO / REMOÇÃO
Dim soma As Currency
Dim contar As Integer
Dim i As Integer

soma = 0
contar = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 4) = "ADIÇÃO" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
      End If
   Next
End With

lblQtdaCadastroSoma.Caption = soma
lblQtdaCadastro.Caption = Format(contar, "000")

soma = 0
contar = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 4) = "REMOÇÃO" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
      End If
   Next
End With

lblQtdaEntradaSoma.Caption = soma
lblQtdaEntrada.Caption = Format(contar, "000")

End Sub
Private Sub FormatarGrid_Quant(rTabela As ADODB.Recordset)
Dim x As Integer
Dim j As Integer
Dim i As Integer

With Grid
   .Clear
   .Cols = 8
   .rows = 2

   .ColWidth(0) = 0
   .ColWidth(1) = 1000
   .ColWidth(2) = 900
   .ColWidth(3) = 1200
   .ColWidth(4) = 1260
   .ColWidth(5) = 1200
   .ColWidth(6) = 1000
   .ColWidth(7) = 1500

   .ColAlignment(1) = 3
   .ColAlignment(2) = 3
   .ColAlignment(5) = 3
   .ColAlignment(6) = 5
   .ColAlignment(7) = 5

    .TextMatrix(0, 1) = "DATA"
    .TextMatrix(0, 2) = "HORA"
    .TextMatrix(0, 3) = "USUÁRIO"
    .TextMatrix(0, 4) = "TIPO"
    .TextMatrix(0, 5) = "FORMA"
    .TextMatrix(0, 6) = "QUANT"
    .TextMatrix(0, 7) = "SALDO/DIA"

   'colocar os cabeçalho em negrito
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next

   'centralizar o titulo
   For j = 0 To .Cols - 1
      .Row = 0
      .Col = j
      .CellAlignment = flexAlignCenterCenter
   Next

   .Redraw = False

   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
        .TextMatrix(.rows - 1, 1) = Format$(rTabela("DATA"), "dd/mm/yy")
        If Trim(ValidateNull(rTabela("Hora"))) = "" Then
           .TextMatrix(.rows - 1, 2) = ""
        Else
           .TextMatrix(.rows - 1, 2) = Format$(rTabela("Hora"), ocHORA)
        End If
        .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("varNomeUsuario"))
        .TextMatrix(.rows - 1, 4) = rTabela("TIPO")
        .TextMatrix(.rows - 1, 5) = rTabela("FORMA")
        .TextMatrix(.rows - 1, 6) = rTabela("QUANT")
        .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("Estoque"))

         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If

    'Deixar vermelho escuro a linha inteira quando for REMOÇÃO
     For i = 1 To .rows - 1
        If .TextMatrix(i, 4) = "REMOÇÃO" Then
           For j = 0 To .Cols - 1
              .Row = i
              .Col = j
              .CellForeColor = &H80&
           Next
        End If
     Next

    'MUDAR COR DE FONTE DA COLUNA (saldo em estoque)
     For i = 1 To .rows - 1
        .Row = i
        .Col = 7
        .CellBackColor = &HC0FFFF
        .CellFontBold = True
     Next

   .Redraw = True
   .rows = .rows - 1

lblQtdaTotal.Caption = SomaGrid(Grid, 6)
End With
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





'Chamado de fora (Saida_Consulta_PorProdutos.frm, cmdExibirPedidos) pra abrir
'já filtrado no produto, quando a origem do movimento foi FORMA = 'AJUSTE'.
Public Sub FiltrarPorProduto(ByVal CodProduto As Long)
cboCriterioPrinc.ListIndex = 0   'TODOS
cboCriterioSec.ListIndex = 2     'CÓD. PRODUTO
txtCodBarra.Text = CodProduto
cboTipoAjuste.ListIndex = 0      'TODOS
cmdLocalizar_Click
End Sub

'Mesma coisa que FiltrarPorProduto, mas replicando o filtro de periodo
'(TODOS/MENSAL/PERIODO) que estava ativo em quem chamou - usado pelo
'Consulta_HistoricoProduto pra nao perder o filtro ao abrir essa tela
Public Sub FiltrarPorProdutoEPeriodo(ByVal CodProduto As Long, ByVal CriterioPrinc As String, Optional ByVal Mes As Integer = 0, Optional ByVal Ano As String = "", Optional ByVal DataIni As String = "", Optional ByVal DataFim As String = "")
cboCriterioSec.ListIndex = 2     'CÓD. PRODUTO
txtCodBarra.Text = CodProduto
cboTipoAjuste.ListIndex = 0      'TODOS

If CriterioPrinc = "MENSAL" Then
   cboCriterioPrinc.ListIndex = 1   'MENSAL
   cboMes_GotFocus
   If Mes >= 1 And Mes <= 12 Then cboMes.ListIndex = Mes - 1
   cboAno_GotFocus
   cboAno.Text = Ano
ElseIf CriterioPrinc = "PERÍODO" Then
   cboCriterioPrinc.ListIndex = 2   'PERÍODO
   mskInicio.Text = DataIni
   mskFim.Text = DataFim
Else
   cboCriterioPrinc.ListIndex = 0   'TODOS
End If

cmdLocalizar_Click
End Sub

Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
   
Limpar_Grid

lblDescricao.Caption = "Cód. Barra"
lblDescricao.Visible = True
txtCodBarra.Visible = True

PreencherConsulta
cboCriterioSec.ListIndex = 1

PreencherCriterios
cboCriterioPrinc.ListIndex = 0

PreencherIndice
cboIndice.ListIndex = 0

cboTipoAjuste.AddItem "TODOS"
cboTipoAjuste.AddItem "ADIÇÃO"
cboTipoAjuste.AddItem "REMOÇÃO"
cboTipoAjuste.ListIndex = 0

If txtCodBarra.Visible = True Then txtCodBarra.SetFocus

StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
posX = x
Label3 = posX
If Label3.Caption > 0 And Label3.Caption < 149 Then Grid.ToolTipText = ""
If Label3.Caption > 150 And Label3.Caption < 930 Then Grid.ToolTipText = "Dê um duplo-clique para exibir os itens do Pedido."
If Label3.Caption > 931 And Label3.Caption < 7230 Then Grid.ToolTipText = ""
If Label3.Caption > 7231 And Label3.Caption < 8355 Then Grid.ToolTipText = "Dê um duplo-clique para exibir a forma de pgto."
If Label3.Caption > 8356 And Label3.Caption < 9555 Then Grid.ToolTipText = ""
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
