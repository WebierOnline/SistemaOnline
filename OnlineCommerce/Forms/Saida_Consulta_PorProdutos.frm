VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Saida_Consulta_PorProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE VENDAS"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   Icon            =   "Saida_Consulta_PorProdutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
      Height          =   255
      Left            =   60
      TabIndex        =   27
      Top             =   8160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "EXIBIR ENTRADAS"
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
      MICON           =   "Saida_Consulta_PorProdutos.frx":23D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   60
      ScaleHeight     =   2205
      ScaleWidth      =   11385
      TabIndex        =   25
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
         Height          =   1515
         Left            =   60
         TabIndex        =   39
         Top             =   60
         Width           =   6315
         Begin VB.ComboBox cboIndice 
            Height          =   315
            Left            =   4740
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox cboCriterioSec 
            Height          =   315
            Left            =   1740
            TabIndex        =   1
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox cboCriterioPrinc 
            Height          =   315
            Left            =   3240
            TabIndex        =   2
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   60
            TabIndex        =   0
            Top             =   480
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   4740
            TabIndex        =   43
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Consultar por:"
            Height          =   195
            Left            =   1740
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Consulta:"
            Height          =   195
            Left            =   60
            TabIndex        =   41
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Critérios:"
            Height          =   195
            Left            =   3240
            TabIndex        =   40
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Frame8 
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
         ForeColor       =   &H000000C0&
         Height          =   1515
         Left            =   6420
         TabIndex        =   26
         Top             =   60
         Width           =   4875
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
            MICON           =   "Saida_Consulta_PorProdutos.frx":23EE
            PICN            =   "Saida_Consulta_PorProdutos.frx":240A
            PICH            =   "Saida_Consulta_PorProdutos.frx":475D
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
            MICON           =   "Saida_Consulta_PorProdutos.frx":6AB0
            PICN            =   "Saida_Consulta_PorProdutos.frx":6ACC
            PICH            =   "Saida_Consulta_PorProdutos.frx":8E1F
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
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.ComboBox cboDescricao 
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   3855
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
         Begin VB.Label lblFim 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data final:"
            Height          =   195
            Left            =   1380
            TabIndex        =   33
            Top             =   840
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblDescricao 
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   60
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblInicio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data inicial:"
            Height          =   195
            Left            =   60
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   840
            Visible         =   0   'False
            Width           =   330
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   495
         Left            =   8220
         TabIndex        =   12
         Top             =   1620
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
         MICON           =   "Saida_Consulta_PorProdutos.frx":B172
         PICN            =   "Saida_Consulta_PorProdutos.frx":B18E
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
         Left            =   9780
         TabIndex        =   13
         Top             =   1620
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
         MICON           =   "Saida_Consulta_PorProdutos.frx":BA68
         PICN            =   "Saida_Consulta_PorProdutos.frx":BA84
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
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3960
      Picture         =   "Saida_Consulta_PorProdutos.frx":BD9E
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   17
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
      TabIndex        =   14
      Top             =   60
      Width           =   11415
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE SAÍDA POR PRODUTOS"
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
         TabIndex        =   15
         Top             =   240
         Width           =   5895
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Saida_Consulta_PorProdutos.frx":CDD6
         Top             =   0
         Width           =   1140
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   24
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
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "00:48"
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
      TabIndex        =   28
      Top             =   3360
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8281
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
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
      TabIndex        =   38
      Top             =   9060
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
      TabIndex        =   37
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
      TabIndex        =   36
      Top             =   8460
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AJUSTE:"
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
      TabIndex        =   35
      Top             =   9060
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
      TabIndex        =   34
      Top             =   9060
      Width           =   1755
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRADA:"
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   8460
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CADASTRO:"
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   16
      Top             =   8220
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Saida_Consulta_PorProdutos"
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
   .Rows = 2
   
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
cboCriterioSec.AddItem "CATEGORIA"
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

Private Sub PreencherTipo()
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
'cboCriterioPrinc_LostFocus
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
    If cboAno.Visible = True Then cboAno.SetFocus
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
    mskInicio.SetFocus
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

Private Sub cboCriterioSec_GotFocus()
Dim vTexto As String
vTexto = cboCriterioSec.Text

cboCriterioSec.Clear
If cboTipo.Text = "POR PRODUTOS" Then
    PreencherConsulta
ElseIf cboTipo.Text = "POR SERVIÇOS" Then
   'cboCriterioSec.AddItem "TECNICO"
   'cboCriterioSec.AddItem "SERVIÇO"
   'cboCriterioSec.AddItem "DESCRIÇÃO"
   'cboCriterioSec.AddItem "DESCRIÇÃO"
End If

cboCriterioSec.Text = vTexto

moCombo.AttachTo cboCriterioSec
End Sub

Private Sub cboCriterioSec_LostFocus()
If cboCriterioSec.Text = "" Then cboCriterioSec.Text = "DESCRIÇÃO"
   
If cboCriterioSec.Text = "CÓD. BARRA" Then
   lblDescricao.Caption = "Cód. Barra"
   lblDescricao.Visible = True
   txtCodBarra.Visible = True
   cboDescricao.Visible = False
   'If txtCodBarra.Visible = True And cboCriterioPrinc.Text <> "" Then txtCodBarra.SetFocus
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
   lblDescricao.Caption = "Descrição"
   lblDescricao.Visible = True
   txtCodBarra.Visible = False
   cboDescricao.Visible = True
   'If cboDescricao.Visible = True And cboCriterioPrinc.Text <> "" Then cboDescricao.SetFocus
ElseIf cboCriterioSec.Text = "CATEGORIA" Then
   lblDescricao.Caption = "Categoria"
   lblDescricao.Visible = True
   txtCodBarra.Visible = False
   cboDescricao.Visible = True
   'If cboDescricao.Visible = True And cboCriterioPrinc.Text <> "" Then cboDescricao.SetFocus
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

Private Sub cboTipo_Change()
If cboTipo.Text = "POR PRODUTOS" Then
   cmdExibirPedidos.Visible = True
'ElseIf cboTipo.Text = "TODOS" Then
   'cmdExibirProdutos.Visible = False
End If
End Sub

Private Sub cboTipo_GotFocus()
Dim vTexto As String
vTexto = cboTipo.Text

cboTipo.Clear
PreencherTipo

cboTipo.Text = vTexto
moCombo.AttachTo cboTipo
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

Private Sub cmdExibirPedidos_Click()
If Grid.Col = 0 Then Exit Sub

If cboTipo.Text = "POR PRODUTOS" Then
    If IsNumeric(Grid.TextMatrix(Grid.Row, 4)) = True Then
        If Grid.TextMatrix(Grid.Row, 4) <> "0000" Then
          Produtos_Entrada.txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 4))
          'Me.Hide
          Produtos_Entrada.Show
        End If
    End If
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
Dim r As ADODB.Recordset

'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)
   
If cboTipo.Text = "POR PRODUTOS" Then
    Set REL_Entrada_PorProduto.Relatorio.Recordset = r

    REL_Entrada_PorProduto.dfQuantTotal.Caption = lblQtdaTotal.Caption
    REL_Entrada_PorProduto.dfQuantCadastroSoma.Caption = lblQtdaCadastroSoma.Caption
    REL_Entrada_PorProduto.dfQuantEntradaSoma.Caption = lblQtdaEntradaSoma.Caption
    REL_Entrada_PorProduto.dfQuantAjusteSoma.Caption = lblQtdaAjusteSoma.Caption
    
    REL_Entrada_PorProduto.dfQuantCadastro.Caption = lblQtdaCadastro.Caption
    REL_Entrada_PorProduto.dfQuantEntrada.Caption = lblQtdaEntrada.Caption
    REL_Entrada_PorProduto.dfQuantAjuste.Caption = lblQtdaAjuste.Caption
    
    If cboCriterioPrinc.Text = "TODOS" Then
       REL_Entrada_PorProduto.rfCons2.Caption = "TODOS"
    ElseIf cboCriterioPrinc.Text = "MENSAL" Then
       REL_Entrada_PorProduto.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
    ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
       REL_Entrada_PorProduto.rfCons2.Caption = "Intervalo de " & mskInicio.Text & " à " & mskFim.Text
    Else
       REL_Entrada_PorProduto.rfCons2.Caption = "TODOS"
    End If

    If cboCriterioSec.Text = "DESCRIÇÃO" Then
       REL_Entrada_PorProduto.rfCons1.Caption = "PRODUTO = " & cboDescricao.Text & ""
    ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
       REL_Entrada_PorProduto.rfCons1.Caption = "CÓD. BARRA = " & txtCodBarra.Text & ""
    ElseIf cboCriterioSec.Text = "CATEGORIA" Then
       REL_Entrada_PorProduto.rfCons1.Caption = "Categoria = " & cboDescricao.Text & ""
    Else
       REL_Entrada_PorProduto.rfCons1.Caption = "PRODUTO = " & cboDescricao.Text & ""
    End If

    'REL_Entrada_PorProduto.Relatorio.NomeImpressora = var_Impressora
    REL_Entrada_PorProduto.Relatorio.Ativar
    Unload REL_Entrada_PorProduto

End If
 
 Me.Show 1
End Sub

Public Sub cmdLocalizar_Click()
Dim INDICE As String
Dim Tipo As String

totalRegistros = "0"

If cboCriterioPrinc.Text = "" Then Exit Sub
If cboCriterioSec.Text = "" Then Exit Sub

If cboTipo.Text = "POR PRODUTOS" Then

If cboTipo.Text = "POR PRODUTOS" Then
    If cboIndice.Text = "QUANT." Then
        INDICE = "SUM(Produtos_Quant.QUANT) DESC;"
    ElseIf cboIndice.Text = "DESCRIÇÃO" Then
        INDICE = "produtos.descricao;"
    ElseIf cboIndice.Text = "DATA" Then
        INDICE = "Produtos_Quant.Data;"
    Else
        INDICE = "produtos.descricao;"
    End If
End If

sSQL = "SELECT Produtos_Quant.CODIGO, Produtos_Quant.Hora, Produtos_Quant.DATA, Produtos_Quant.QUANT, ISNULL(Produtos_Quant.cod_usuario, '') AS vUsuario, ISNULL(Produtos_Quant.Estoque, ' ') AS vQuantEstoque, Produtos_Quant.FORMA, Produtos_Quant.COD_ENTRADA, Produtos_Quant.COD_PRODUTO, produtos_entrada.NOTAFISCAL, produtos.descricao, isnull(produtos_entrada.CODIGO,0) AS vCodEntrada " & _
       "FROM Produtos_Quant INNER JOIN produtos ON Produtos_Quant.COD_PRODUTO = produtos.CODIGO LEFT OUTER JOIN produtos_entrada ON Produtos_Quant.COD_ENTRADA = produtos_entrada.CODIGO " & _
       "WHERE (Produtos_Quant.TIPO <> 'REMOÇÃO') "
    
    If cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "TODOS" Then
        If txtCodBarra.Text = "" Then Exit Sub
        sSQL = sSQL & "AND (Produtos.COD_BARRA = '" & txtCodBarra.Text & "')"
    ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "MENSAL" Then
        If txtCodBarra.Text = "" Then Exit Sub
        If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
        sSQL = sSQL & "AND (Produtos.COD_BARRA = '" & txtCodBarra.Text & "') and (MONTH(Produtos_Quant.DATA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant.DATA) = " & cboAno & ") "
    ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "PERÍODO" Then
        If txtCodBarra.Text = "" Then Exit Sub
        If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
        sSQL = sSQL & "AND (Produtos.COD_BARRA = '" & txtCodBarra.Text & "') and (Produtos_Quant.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    End If
    
    If cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "TODOS" Then
        If cboDescricao.Text = "" Then Exit Sub
        sSQL = sSQL & "AND (Produtos.DESCRICAO = '" & cboDescricao.Text & "')"
    ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "MENSAL" Then
        If cboDescricao.Text = "" Then Exit Sub
        If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
        sSQL = sSQL & "AND (Produtos.DESCRICAO = '" & cboDescricao.Text & "') and (MONTH(Produtos_Quant.DATA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant.DATA) = " & cboAno & ") "
    ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "PERÍODO" Then
        If cboDescricao.Text = "" Then Exit Sub
        If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
        sSQL = sSQL & "AND (Produtos.DESCRICAO = '" & cboDescricao.Text & "') and (Produtos_Quant.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    End If

    If cboCriterioSec.Text = "CATEGORIA" And cboCriterioPrinc.Text = "TODOS" Then
        If cboDescricao.Text = "" Then Exit Sub
        sSQL = sSQL & "AND (Produtos.CATEGORIA = '" & cboDescricao.Text & "')"
    ElseIf cboCriterioSec.Text = "CATEGORIA" And cboCriterioPrinc.Text = "MENSAL" Then
        If cboDescricao.Text = "" Then Exit Sub
        If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
        sSQL = sSQL & "AND (Produtos.CATEGORIA = '" & cboDescricao.Text & "') and (MONTH(Produtos_Quant.DATA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant.DATA) = " & cboAno & ") "
    ElseIf cboCriterioSec.Text = "CATEGORIA" And cboCriterioPrinc.Text = "PERÍODO" Then
        If cboDescricao.Text = "" Then Exit Sub
        If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
        sSQL = sSQL & "AND (Produtos.CATEGORIA = '" & cboDescricao.Text & "') and (Produtos_Quant.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    End If
   
ElseIf cboTipo.Text = "POR SERVIÇOS" Then


         'TODOS
         If cboCriterioPrinc.Text = "TODOS" And cboCriterioSec.Text = "" Then
            'sSQL = "SELECT os_servicos.cod_produto, os_servicos.descricao as var_desc, SUM(os_servicos.QUANT) AS var_qtde, CUSTO, SUM(CUSTO * QUANT) AS var_total " & _
            '   "FROM produtos INNER JOIN os_servicos ON produtos.codigo = os_servicos.cod_produto " & _
            '   "INNER JOIN produtos_entrada ON os_servicos.cod_pedido = produtos_entrada.cod_pedido " & _
            '   "WHERE (produtos_entrada.tipo_pedido = 'BALCAO' or produtos_entrada.tipo_pedido = 'OFICINA') " & _
               "GROUP BY os_servicos.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, produtos.ref, os_servicos.CUSTO ORDER BY " & INDICE
         End If
    

End If

sSQL = sSQL & " ORDER BY " & INDICE

Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid_Quant r

If r.State <> 0 Then r.Close
Set r = Nothing
   
printSQL = sSQL

'Dividir os totais
Dim soma As Currency
Dim contar As Integer
Dim i As Integer

soma = 0
contar = 0
With Grid
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 6) = "CADASTRO" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 8))
      End If
   Next
End With

lblQtdaCadastroSoma.Caption = soma
lblQtdaCadastro.Caption = Format(contar, "000")

soma = 0
contar = 0
With Grid
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 6) = "ENTRADA" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 8))
      End If
   Next
End With

lblQtdaEntradaSoma.Caption = soma
lblQtdaEntrada.Caption = Format(contar, "000")

soma = 0
contar = 0
With Grid
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 6) = "AJUSTE" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 8))
      End If
   Next
End With

lblQtdaAjusteSoma.Caption = soma
lblQtdaAjuste.Caption = Format(contar, "000")

End Sub
Private Sub FormatarGrid_Quant(rTabela As ADODB.Recordset)
Dim x As Integer
Dim j As Integer
Dim i As Integer

With Grid
   .Clear
   .Cols = 11
   .Rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 850
   .ColWidth(3) = 750
   .ColWidth(4) = 900
   .ColWidth(5) = 3300
   .ColWidth(6) = 1100
   .ColWidth(7) = 1400
   .ColWidth(8) = 1200
   .ColWidth(9) = 1000
   .ColWidth(10) = 1500
   
   .ColAlignment(2) = 5
   .ColAlignment(3) = 5

    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "DATA"
    .TextMatrix(0, 3) = "HORA"
    .TextMatrix(0, 4) = "COD_ENTRADA"
    .TextMatrix(0, 5) = "PRODUTO"
    .TextMatrix(0, 6) = "FORMA"
    .TextMatrix(0, 7) = "NOTA FISCAL"
    .TextMatrix(0, 8) = "QUANT"
    .TextMatrix(0, 9) = "USUÁRIO"
    .TextMatrix(0, 10) = "ESTOQUE/DIA"
   
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
        .TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("Codigo"))
        .TextMatrix(.Rows - 1, 2) = Format$(rTabela("Data"), "dd/mm/yy")
        .TextMatrix(.Rows - 1, 3) = Format$(rTabela("HORA"), ocHORA)
        .TextMatrix(.Rows - 1, 4) = rTabela("vCodEntrada")
        .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("DESCRICAO"))
        .TextMatrix(.Rows - 1, 6) = rTabela("FORMA")
        .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("notafiscal"))
        .TextMatrix(.Rows - 1, 8) = rTabela("QUANT")
        .TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("vUsuario"))
        .TextMatrix(.Rows - 1, 10) = ValidateNull(rTabela("vQuantEstoque"))
         
         rTabela.MoveNext
         .Rows = .Rows + 1
      Loop
   End If
   
    'MUDAR COR DE FONTE DA COLUNA
     For i = 1 To .Rows - 1
        .Row = i
        .Col = 8
        .CellBackColor = &HC0FFFF
        .CellFontBold = True
     Next

   .Redraw = True
   .Rows = .Rows - 1

lblQtdaTotal.Caption = SomaGrid(Grid, 8)
End With
End Sub
Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Double
Dim i As Integer, Valor As Currency

Valor = 0

For i = 0 To var_Grid.Rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function





Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
   
Limpar_Grid

lblDescricao.Caption = "Cód. Barra"
lblDescricao.Visible = True
txtCodBarra.Visible = True

PreencherTipo
cboTipo.ListIndex = 0

PreencherConsulta
cboCriterioSec.ListIndex = 1

PreencherCriterios
cboCriterioPrinc.ListIndex = 0

PreencherIndice
cboIndice.ListIndex = 0

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
