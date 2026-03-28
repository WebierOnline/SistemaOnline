VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form OS_Consulta_Servicos_Pecas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEÇAS & SERVIÇOS"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   Icon            =   "OS_Consulta_Servicos_Pecas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   10365
      TabIndex        =   33
      Top             =   60
      Width           =   10395
      Begin VB.Image Image2 
         Height          =   555
         Left            =   9660
         Picture         =   "OS_Consulta_Servicos_Pecas.frx":23D2
         Top             =   180
         Width           =   600
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PEÇAS E SERVIÇOS"
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
         Left            =   3780
         TabIndex        =   34
         Top             =   240
         Width           =   3060
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "OS_Consulta_Servicos_Pecas.frx":7A90
         Top             =   120
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   5475
      Left            =   60
      ScaleHeight     =   5415
      ScaleWidth      =   10335
      TabIndex        =   24
      Top             =   3480
      Width           =   10395
      Begin VB.TextBox txtTotalOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8940
         TabIndex        =   29
         Top             =   4740
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtQtdaOs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   28
         Top             =   4740
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtTotalVenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8940
         TabIndex        =   27
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtQtdaVenda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   26
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4605
         Left            =   60
         TabIndex        =   25
         Top             =   60
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8123
         _Version        =   393216
         Cols            =   8
         BackColorBkg    =   -2147483633
         Appearance      =   0
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   6960
         TabIndex        =   32
         Top             =   4560
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "OS:"
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
         Left            =   7980
         TabIndex        =   31
         Top             =   4800
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Balcão:"
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
         Left            =   7680
         TabIndex        =   30
         Top             =   5100
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   60
      ScaleHeight     =   2355
      ScaleWidth      =   10335
      TabIndex        =   0
      ToolTipText     =   "Imprimir"
      Top             =   1020
      Width           =   10395
      Begin VB.Frame frmFormaVenda 
         Caption         =   "Vendas"
         Enabled         =   0   'False
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
         Height          =   1575
         Left            =   1620
         TabIndex        =   35
         Top             =   60
         Width           =   1155
         Begin VB.OptionButton OptVendaOficina 
            Caption         =   "Oficina"
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
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton OptVendaTodas 
            Caption         =   "Todas"
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
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton OptVendaBalcao 
            Caption         =   "Balcão"
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
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   540
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sub - Critérios"
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
         Height          =   1575
         Left            =   5640
         TabIndex        =   14
         Top             =   60
         Width           =   4635
         Begin VB.ComboBox cboAno 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3060
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   1515
         End
         Begin VB.ComboBox cboDescricao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            TabIndex        =   16
            Top             =   1080
            Width           =   3315
         End
         Begin VB.ComboBox cboMES 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "OS_Consulta_Servicos_Pecas.frx":DC9C
            Left            =   1260
            List            =   "OS_Consulta_Servicos_Pecas.frx":DC9E
            TabIndex        =   15
            Top             =   240
            Width           =   1755
         End
         Begin MSMask.MaskEdBox Mask2 
            Height          =   315
            Left            =   3300
            TabIndex        =   18
            Top             =   660
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Mask1 
            Height          =   315
            Left            =   1260
            TabIndex        =   19
            Top             =   660
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Label lblCONnome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   420
            TabIndex        =   23
            Top             =   1140
            Width           =   765
         End
         Begin VB.Label lblCONint2 
            AutoSize        =   -1  'True
            Caption         =   "Data &Final:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2460
            TabIndex        =   22
            Top             =   705
            Width           =   765
         End
         Begin VB.Label lblCONint1 
            AutoSize        =   -1  'True
            Caption         =   "Da&ta Inicial:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   720
            Width           =   840
         End
         Begin VB.Label lblCONmes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&scolha o mês:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   1080
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Indice:"
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
         Height          =   1575
         Left            =   2820
         TabIndex        =   9
         Top             =   60
         Width           =   1395
         Begin VB.OptionButton optOrdTipo 
            Caption         =   "Código"
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
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton optOrdDescricao 
            Caption         =   "Descrição"
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
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton optOrdData 
            Caption         =   "Data"
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
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Tipo:"
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
         Height          =   1575
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   1515
         Begin VB.OptionButton optOS 
            Caption         =   "Serviços"
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
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optVendas 
            Caption         =   "Peças"
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
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   540
            Width           =   1035
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Critérios:"
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
         Height          =   1575
         Left            =   4260
         TabIndex        =   1
         Top             =   60
         Width           =   1365
         Begin VB.OptionButton optCritTodos 
            Caption         =   "&Todos"
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
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton optCritDescricao 
            Caption         =   "Descrição"
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
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton optCritIntervalor 
            Caption         =   "Intervalo"
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
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   900
            Width           =   1155
         End
         Begin VB.OptionButton optCritMensal 
            Caption         =   "Mensal"
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
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   600
            Width           =   1035
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimir 
         Height          =   615
         Left            =   7740
         TabIndex        =   5
         Top             =   1680
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1085
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
         MICON           =   "OS_Consulta_Servicos_Pecas.frx":DCA0
         PICN            =   "OS_Consulta_Servicos_Pecas.frx":DCBC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   9000
         TabIndex        =   13
         Top             =   1680
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1085
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
         MICON           =   "OS_Consulta_Servicos_Pecas.frx":DFD6
         PICN            =   "OS_Consulta_Servicos_Pecas.frx":DFF2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   615
         Left            =   6480
         TabIndex        =   40
         Top             =   1680
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1085
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
         MICON           =   "OS_Consulta_Servicos_Pecas.frx":E30C
         PICN            =   "OS_Consulta_Servicos_Pecas.frx":E328
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   41
      Top             =   8985
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14261
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "22:19"
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
End
Attribute VB_Name = "OS_Consulta_Servicos_Pecas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

Dim Tipo As String                  'usado no botao exibir
Dim Texto As String                 'usado pra preencher os combos
Dim i, Posicao As Integer           'usado pra preencher os combos
Dim Posicionar As Boolean           'usado pra preencher os combos

Private Sub Limpar_Objetos_Baixo()
   txtQtdaOs.Text = ""
   txtTotalOS.Text = ""
   txtQtdaVenda.Text = ""
   txtTotalVenda.Text = ""
End Sub

Private Sub Limpar_Objetos()
   'cboVendedor.Text = ""
   'mskInicio.Mask = ""
   'mskInicio.Text = ""
   'mskFim.Mask = ""
   'mskFim.Text = ""
   'mskDia.Mask = ""
   'mskDia.Text = ""
   'cboCliente.Text = ""
   cboMES.Text = ""
   cboAno.Text = ""
End Sub

Private Sub Montar_Grid_OS(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 100
      .ColWidth(1) = 850
      .ColWidth(2) = 850
      .ColWidth(3) = 5750
      .ColWidth(4) = 850
      .ColWidth(5) = 650
      .ColWidth(6) = 850
      
      .TextMatrix(0, 1) = "DATA"
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "DESCRIÇÃO"
      .TextMatrix(0, 4) = "PREÇO"
      .TextMatrix(0, 5) = "QTDA"
      .TextMatrix(0, 6) = "TOTAL"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      .ColAlignment(3) = 1
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            For i = 1 To .Rows - 1
               .Row = i
               .Col = 2
               .CellBackColor = vbYellow
            Next
            
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("cod_os"), "000000")
            .TextMatrix(.Rows - 1, 3) = UCase(rTabela("descricao"))
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.Rows - 1, 5) = rTabela("quantidade")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("total"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub cboAno_GotFocus()
   Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
   
   cboAno.Clear
   
   iAno = Year(Date)
   FirstYear = iAno - 2
   LastYear = iAno + 2
   
   For i = LastYear To FirstYear Step -1
      cboAno.AddItem i
   Next
   
   'For i = iAno To FirstYear Step -1
   '   cboAno.AddItem i
   'Next
   '
   'iAno = iAno + 1
   'For i = iAno To LastYear
   '   cboAno.AddItem
   'Next
End Sub

Private Sub cboAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdExibir_Click
End Sub

Private Sub cboAno_LostFocus()
   If cboAno.Text = "" Then Exit Sub Else cmdExibir.SetFocus
End Sub

Private Sub cboDescricao_LostFocus()
   If cboDescricao.Text = "" Then Exit Sub Else cmdExibir.SetFocus
End Sub

Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMES.Clear
   For vMes = 1 To 12
      cboMES.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMES
End Sub

Private Sub cboMes_LostFocus()
   If cboMES.Text = "" Then Exit Sub Else cboAno.SetFocus
End Sub

Private Sub cboDescricao_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboDescricao.Clear
   
   If optOS.Value = True Then
      sSQL = "SELECT DISTINCT descricao FROM os_servicos ORDER BY descricao;"
      Set r = dbData.OpenRecordset(sSQL)
      
      'ABRIR_BD_com_Data Me.Data2
      'Data2.RecordSource = "SELECT DISTINCT DESCRICAO FROM OS_SERVICOS ORDER BY DESCRICAO"
      'Data2.Refresh
      
      Do While Not r.EOF
         cboDescricao.AddItem r("descricao")
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
   ElseIf optVendas.Value = True Then
      sSQL = "SELECT DISTINCT descricao FROM pedidos_itens ORDER BY descricao;"
      Set r = dbData.OpenRecordset(sSQL)
      
      'ABRIR_BD_com_Data Me.Data2
      'Data2.RecordSource = "SELECT DISTINCT DESCRICAO FROM PEDIDOS_ITENS ORDER BY DESCRICAO"
      'Data2.Refresh
      
      Do While Not r.EOF
         cboDescricao.AddItem r("descricao")
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
   
   moCombo.AttachTo cboDescricao
End Sub

Private Sub cboDescricao_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdExibir_Click
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdExibir_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   Dim vTotal As Currency
   
   Dim INDICE As String       'indice
   Dim FORMA As String        'forma de venda
   
   If optOrdDescricao.Value = True Then
      INDICE = "descricao"
   ElseIf optOrdData.Value = True Then
      INDICE = "data"
   ElseIf optOrdTipo.Value = True Then
      If optOS.Value = True Then
         INDICE = "cod_os"
      ElseIf optVendas.Value = True Then
         INDICE = "tipo_venda"
      End If
   End If
   
   If OptVendaTodas.Value = True Then
      FORMA = "(tipo_venda <> '') "
   ElseIf OptVendaBalcao.Value = True Then
      FORMA = "(tipo_venda = 'BALCÃO') "
   ElseIf OptVendaOficina.Value = True Then
      FORMA = "(tipo_venda = 'OFICINA') "
   End If
   
   Limpar_Objetos_Baixo
   
   'Call Abrir_BancodeDados
   'SQL = "SELECT * FROM CAIXA_DIA ORDER BY DATA"
   'Set RS = BD.OpenRecordset(SQL)
   
   If optCritTodos.Value = True Then
      If optOS.Value = True Then
         'SERVIÇOS
         sSQL = "SELECT *, CASE cod_os WHEN 0 THEN 'OS' ELSE 'OS' END AS var_tipo, CASE cod_os WHEN 0 THEN cod_os ELSE cod_os END AS var_codigo " & _
            "FROM os_servicos ORDER BY " & INDICE
         
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         Montar_Grid_OS r
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         'CALCULAR
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total), 0) AS valor_total FROM os_servicos;"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_total")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaOs.Text = Format(totalRegistros, "00")
         txtTotalOS.Text = Format(vTotal, ocMONEY)
         
    ElseIf optVendas.Value = True Then
         'PEÇAS
         sSQL = "SELECT pedidos_itens.*, pedidos_itens.tipo_venda AS var_tipo, CASE cod_pedido WHEN 0 THEN cod_os ELSE cod_pedido END AS var_codigo " & _
            "FROM pedidos_itens WHERE " & FORMA & " ORDER BY " & INDICE
         
         Set r = dbData.OpenRecordset(sSQL)
         Montar_Grid_Vendas r
         
         'QUANTIDADE DE VENDAS POR BALCAO
         sSQL = "SELECT * FROM pedidos_itens WHERE (cod_pedido <> 0);"
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         If Not r.BOF Then r.Close
         Set r = Nothing
         
         txtQtdaVenda.Text = Format(totalRegistros, "00")
         
         'QUANTIDADE DE VENDAS POR OFICINA
         sSQL = "SELECT * FROM pedidos_itens WHERE (cod_os <> 0)"
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaOs.Text = Format(totalRegistros, "00")
         
         'CALCULAR AS VENDAS POR BALCAO
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total) AS valor_balcao FROM pedidos_itens WHERE (cod_pedido <> 0);"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_balcao")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalVenda.Text = Format(vTotal, ocMONEY)
         
         'CALCULAR AS VENDAS POR OFICINA
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total) AS valor_os FROM pedidos_itens WHERE (cod_os <> 0);"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = ("valor_os")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalOS.Text = Format(vTotal, ocMONEY)
      End If
   
   ElseIf optCritDescricao.Value = True Then
      If cboDescricao.Text = "" Then Exit Sub
      
      If optOS.Value = True Then
         'SERVIÇOS
         sSQL = "SELECT *, CASE cod_os WHEN 0 THEN 'OS' ELSE 'OS' END AS var_tipo, CASE cod_os WHEN 0 THEN cod_os ELSE cod_os END AS var_codigo " & _
            "FROM os_servicos WHERE (descricao = '" & cboDescricao.Text & "') ORDER BY " & INDICE
         
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         Montar_Grid_OS r
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaOs.Text = Format(totalRegistros, "00")
         
         'CALCULAR
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total), 0) AS valor_total FROM os_servicos WHERE (descricao = '" & cboDescricao.Text & "');"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_total")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalOS.Text = Format(vTotal, ocMONEY)
         
      ElseIf optVendas.Value = True Then
         'PEÇAS
         sSQL = "SELECT pedidos_itens.*, pedidos_itens.tipo_venda AS var_tipo, CASE cod_pedido WHEN 0 THEN cod_os ELSE cod_pedido END AS var_codigo " & _
            "FROM pedidos_itens WHERE (descricao = '" & cboDescricao.Text & "') AND " & FORMA & " ORDER BY " & INDICE
         
         Set r = dbData.OpenRecordset(sSQL)
         Montar_Grid_Vendas r
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         'QUANTIDADE DE VENDAS POR BALCÃO
         sSQL = "SELECT * FROM pedidos_itens WHERE (cod_pedido <> 0) AND (descricao = '" & cboDescricao.Text & "');"
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaVenda.Text = Format(totalRegistros, "00")
         
         'QUANTIDADE DE VENDAS POR OFICINA
         sSQL = "SELECT * FROM pedidos_itens WHERE (cod_os <> 0) AND (descricao = '" & cboDescricao.Text & "');"
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         If Not r.BOF Then r.Close
         Set r = Nothing
         
         txtQtdaOs.Text = Format(totalRegistros, "00")
         
         'CALCULAR AS VENDAS POR BALCAO
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total), 0) AS valor_balcao FROM pedidos_itens WHERE (cod_pedido <> 0) AND (descricao = '" & cboDescricao.Text & "');"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_balcao")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalVenda.Text = Format(vTotal, ocMONEY)
         
         'CALCULAR AS VENDAS POR OFICINA
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total )AS valor_os FROM pedidos_itens WHERE (cod_os <> 0) AND (descricao = '" & cboDescricao.Text & "');"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_os")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalOS.Text = Format(vTotal, ocMONEY)
      End If
      
   ElseIf optCritIntervalor.Value = True Then
      If Mask1.Text = "" And Mask2.Text = "" Then Exit Sub
      If Not IsDate(Mask1) = True Or Not IsDate(Mask2) = True Then Exit Sub
      
      If optOS.Value = True Then
         'SERVIÇOS
         sSQL = "SELECT *, CASE cod_os WHEN 0 THEN 'OS' ELSE 'OS' END AS var_tipo, CASE cod_os WHEN 0 THEN cod_os ELSE cod_os AS var_codigo " & _
            "FROM os_servicos WHERE (data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) AND (data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103)) ORDER BY " & INDICE
         
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         Montar_Grid_OS r
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaOs.Text = Format(totalRegistros, "00")
         
         'CALCULAR
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total) AS valor_total FROM os_servicos WHERE (data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) AND (data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103));"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_total")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalOS.Text = Format(vTotal, ocMONEY)
         
      ElseIf optVendas.Value = True Then
         'PEÇAS
         sSQL = "SELECT pedidos_itens.*, pedidos_itens.tipo_venda AS var_tipo, CASE cod_pedido WHEN 0 THEN cod_os ELSE cod_pedido END AS var_codigo " & _
            "FROM pedidos_itens WHERE (data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) AND (data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103)) AND " & FORMA & " ORDER BY " & INDICE
         
         Set r = dbData.OpenRecordset(sSQL)
         Montar_Grid_Vendas r
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         'QUANTIDADE DE VENDAS POR BALCAO
         sSQL = "SELECT * FROM pedidos_itens WHERE (cod_pedido <> 0) AND (data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) AND (data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103));"
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaVenda.Text = Format(totalRegistros, "00")
         
         'QUANTIDADE DE VENDAS POR OFICINA
         sSQL = "SELECT * FROM pedidos_itens WHERE (cod_os <> 0) AND (data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) AND (data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA_EUA) & "', 103));"
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         If Not r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaOs.Text = Format(totalRegistros, "00")
         
         'CALCULAR AS VENDAS POR BALCAO
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total), 0) AS valor_balcao FROM pedidos_itens WHERE (cod_pedido <> 0) AND (data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "') AND (data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103));"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_balcao")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalVenda.Text = Format(vTotal, ocMONEY)
         
         'CALCULAR AS VENDAS POR OFICINA
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total), 0) AS valor_os FROM pedidos_itens WHERE (cod_os <> 0) AND (data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) AND (data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103));"
         Set r = dbData.OpenRecordset(sSQL)
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalOS.Text = Format(vTotal, ocMONEY)
      End If
      
   ElseIf optCritMensal.Value = True Then
      If cboMES.Text = "" And cboAno.Text = "" Then Exit Sub
      
      If optOS.Value = True Then
         sSQL = "SELECT *, CASE cod_os WHEN 0 THEN 'OS' ELSE 'OS' END AS var_tipo, CASE cod_os WHEN 0 THEN cod_os ELSE cod_os AS var_codigo " & _
            "FROM os_servicos WHERE (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ") ORDER BY " & INDICE
         
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         Montar_Grid_OS r
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaOs.Text = Format(totalRegistros, "00")
         
         'CALCULAR
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total), 0) AS valor_total FROM os_servicos WHERE (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ");"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_total")
         If r.State <> 0 Then r.Close
         
         txtTotalOS.Text = Format(vTotal, ocMONEY)
         
      ElseIf optVendas.Value = True Then
         'PEÇAS
         sSQL = "SELECT pedidos_itens.*, pedidos_itens.tipo_venda AS var_tipo, CASE cod_pedido WHEN 0 THEN cod_os ELSE cod_pedido END AS var_codigo " & _
            "FROM pedidos_itens WHERE (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ") AND " & FORMA & " ORDER BY " & INDICE
         
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         Montar_Grid_Vendas r
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         sSQL = "SELECT * FROM pedidos_itens WHERE (cod_pedido <> 0) AND (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ");"
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtQtdaVenda.Text = Format(totalRegistros, "00")
         
         'QUANTIDADE DE VENDAS POR OFICINA
         sSQL = "SELECT * FROM pedidos_itens WHERE (cod_os <> 0) AND (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ");"
         Set r = dbData.OpenRecordset(sSQL, totalRegistros)
         
         txtQtdaOs.Text = Format(totalRegistros, "00")
         
         'CALCULAR AS VENDAS POR BALCAO
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total), 0) AS valor_balcao FROM pedidos_itens WHERE (cod_pedido <> 0) AND (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ");"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_balcao")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalVenda.Text = Format(vTotal, ocMONEY)
         
         'CALCULAR AS VENDAS POR OFICINA
         vTotal = 0
         sSQL = "SELECT ISNULL(SUM(total), 0) AS valor_os FROM pedidos_itens WHERE (cod_os <> 0) AND (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ");"
         Set r = dbData.OpenRecordset(sSQL)
         If Not r.BOF Then vTotal = r("valor_os")
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         txtTotalOS.Text = Format(vTotal, ocMONEY)
      End If
   End If
   
   printSQL = sSQL
End Sub

Private Sub cmdImprimir_Click()
   If optCritTodos.Value = False And optCritDescricao.Value = False And optCritIntervalor.Value = False And optCritMensal.Value = False Then Exit Sub
   
   'colocar o nome da maquina na barra de status
   Dim var_Impressora As String
   Dim oIni As Ini
   Dim r As ADODB.Recordset
   
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   Me.Hide
   
   Set r = dbData.OpenRecordset(printSQL)
   
   Set REL_OS_Servicos_Pecas.Relatorio.Recordset = r
   
   If optOS.Value = True Then
      REL_OS_Servicos_Pecas.lblTitulo.Caption = "RELATÓRIO DE SERVIÇOS"
   Else
      If OptVendaTodas.Value = True Then
         REL_OS_Servicos_Pecas.lblTitulo.Caption = "RELATÓRIO DE PEÇAS"
      ElseIf OptVendaBalcao.Value = True Then
         REL_OS_Servicos_Pecas.lblTitulo.Caption = "RELATÓRIO DE PEÇAS (BALCÃO)"
      ElseIf OptVendaOficina.Value = True Then
         REL_OS_Servicos_Pecas.lblTitulo.Caption = "RELATÓRIO DE PEÇAS (OFICINA)"
      End If
   End If
   
   REL_OS_Servicos_Pecas.dfQuantVenda.Caption = txtQtdaVenda.Text
   REL_OS_Servicos_Pecas.dfTotalVenda.Caption = txtTotalVenda.Text
   REL_OS_Servicos_Pecas.dfQuantOS.Caption = txtQtdaOs.Text
   REL_OS_Servicos_Pecas.dfTotalOS.Caption = txtTotalOS.Text
   
   If optCritTodos.Value = True Then
      REL_OS_Servicos_Pecas.dfTipo.Caption = "Tipo: Todos"
   ElseIf optCritDescricao.Value = True Then
      REL_OS_Servicos_Pecas.dfTipo.Caption = "Tipo: Desc.: = " & cboDescricao.Text & ""
   ElseIf optCritIntervalor.Value = True Then
      REL_OS_Servicos_Pecas.dfTipo.Caption = "Tipo: Intervalo de " & Mask1.Text & " à " & Mask2.Text
   ElseIf optCritMensal.Value = True Then
      REL_OS_Servicos_Pecas.dfTipo.Caption = "Tipo: Mês = " & cboMES.Text & "/" & cboAno.Text
   End If
   
   REL_OS_Servicos_Pecas.Relatorio.NomeImpressora = var_Impressora
   REL_OS_Servicos_Pecas.Relatorio.Ativar
   Unload REL_OS_Servicos_Pecas
   
   Me.Show 1
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   'FORMATAR O GRID
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
   
   Label1.Caption = "Serviços:"
   txtQtdaOs.Visible = True
   txtTotalOS.Visible = True
   Label1.Visible = True
   
   Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Sub Montar_Grid_Vendas(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 100
      .ColWidth(1) = 850
      .ColWidth(2) = 830
      .ColWidth(3) = 710
      .ColWidth(4) = 5100
      .ColWidth(5) = 850
      .ColWidth(6) = 650
      .ColWidth(7) = 850
      
      .TextMatrix(0, 1) = "DATA"
      .TextMatrix(0, 2) = "TIPO"
      .TextMatrix(0, 3) = "CÓD"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "PREÇO"
      .TextMatrix(0, 6) = "QTDA"
      .TextMatrix(0, 7) = "TOTAL"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 3
      .ColAlignment(3) = 3
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            For i = 1 To Grid.Rows - 1
               .Row = i
               .Col = 3
               .CellBackColor = vbYellow
            Next
            
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_tipo")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("var_codigo"), "000000")
            .TextMatrix(.Rows - 1, 4) = UCase(rTabela("descricao"))
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.Rows - 1, 6) = rTabela("quantidade")
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("total"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub Grid_DblClick()
   If Grid.Col = 0 Then Exit Sub
   
   If optOS.Value = True Then
      If Grid.Col = 1 Then
         If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
      Else
         Ordem_Servicos_Consulta_Geral.optCodigo.Value = True
         Ordem_Servicos_Consulta_Geral.txtCodigo.Text = CInt(Grid.TextMatrix(Grid.Row, 2))
         Ordem_Servicos_Consulta_Geral.cmdLocalizar_Click
         Ordem_Servicos_Consulta_Geral.Show 1
      End If
   
   ElseIf optVendas.Value = True Then
      If Grid.TextMatrix(Grid.Row, 2) = "OS" Then
         Ordem_Servicos_Consulta_Geral.optCodigo.Value = True
         Ordem_Servicos_Consulta_Geral.txtCodigo.Text = CInt(Grid.TextMatrix(Grid.Row, 3))
         Ordem_Servicos_Consulta_Geral.cmdLocalizar_Click
         Ordem_Servicos_Consulta_Geral.Show 1
      Else
         If Grid.Col = 1 Then
            If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
         Else
            'Vendas_Consulta_Geral.cboTipo.Text = "CÓDIGO"
            'Vendas_Consulta_Geral.txtCodigo.Text = CInt(Grid.TextMatrix(Grid.Row, 3))
            'Vendas_Consulta_Geral.cmdLocalizar_Click
            'Vendas_Consulta_Geral.Show 1
         End If
      End If
   End If
End Sub

Private Sub MASK1_KeyPress(KeyAscii As Integer)
   Mask1.Mask = "##/##/##"
End Sub

Private Sub Mask1_LostFocus()
   If Mask1.Text = "" Then Exit Sub Else Mask2.SetFocus
End Sub

Private Sub Mask2_KeyPress(KeyAscii As Integer)
   Mask2.Mask = "##/##/##"
End Sub

Private Sub Mask2_LostFocus()
   If Mask2.Text = "" Then Exit Sub Else cmdExibir.SetFocus
End Sub

Private Sub optCritDescricao_Click()
   lblCONmes.Enabled = False
   cboMES.Enabled = False
   cboAno.Enabled = False
   lblCONint1.Enabled = False
   Mask1.Enabled = False
   lblCONint2.Enabled = False
   Mask2.Enabled = False
   lblCONnome.Enabled = True
   cboDescricao.Enabled = True
   cboDescricao.SetFocus
End Sub

Private Sub optCritIntervalor_Click()
   lblCONmes.Enabled = False
   cboMES.Enabled = False
   cboAno.Enabled = False
   lblCONint1.Enabled = True
   Mask1.Enabled = True
   lblCONint2.Enabled = True
   Mask2.Enabled = True
   Mask1.SetFocus
   lblCONnome.Enabled = False
   cboDescricao.Enabled = False
End Sub

Private Sub optCritMensal_Click()
   lblCONmes.Enabled = True
   cboMES.Enabled = True
   cboAno.Enabled = True
   cboMES.SetFocus
   lblCONint1.Enabled = False
   Mask1.Enabled = False
   lblCONint2.Enabled = False
   Mask2.Enabled = False
   lblCONnome.Enabled = False
   cboDescricao.Enabled = False
End Sub

Private Sub optCritTodos_Click()
   lblCONmes.Enabled = False
   cboMES.Enabled = False
   cboAno.Enabled = False
   lblCONint1.Enabled = False
   Mask1.Enabled = False
   lblCONint2.Enabled = False
   Mask2.Enabled = False
   lblCONnome.Enabled = False
   cboDescricao.Enabled = False
   cmdExibir_Click
End Sub

Private Sub optOrdData_Click()
   cmdExibir_Click
End Sub

Private Sub optOrdDescricao_Click()
   cmdExibir_Click
End Sub

Private Sub optOrdTipo_Click()
   cmdExibir_Click
End Sub

Private Sub optOS_Click()
   If optCritTodos.Value = True Then
      cmdExibir_Click
   ElseIf optCritDescricao.Value = True Then
      cmdExibir_Click
   ElseIf optCritIntervalor.Value = True Then
      cmdExibir_Click
   ElseIf optCritMensal.Value = True Then
      cmdExibir_Click
   End If
   
   frmFormaVenda.Enabled = False
   optOrdTipo.Caption = "Código"
   
   Label1.Caption = "Serviços:"
   Label2.Caption = "Serviços"
   txtQtdaOs.Visible = True
   txtTotalOS.Visible = True
   txtQtdaVenda.Visible = False
   txtTotalVenda.Visible = False
   Label1.Visible = True
   Label2.Visible = False
End Sub

Private Sub OptVendaBalcao_Click()
   cmdExibir_Click
End Sub

Private Sub OptVendaOficina_Click()
   cmdExibir_Click
End Sub

Private Sub optVendas_Click()
   If optCritTodos.Value = True Then
      cmdExibir_Click
   ElseIf optCritDescricao.Value = True Then
      cmdExibir_Click
   ElseIf optCritIntervalor.Value = True Then
      cmdExibir_Click
   ElseIf optCritMensal.Value = True Then
      cmdExibir_Click
   End If
   
   frmFormaVenda.Enabled = True
   optOrdTipo.Caption = "Tipo"
   
   Label1.Caption = "Oficina:"
   Label2.Caption = "Balcão:"
   txtQtdaOs.Visible = True
   txtTotalOS.Visible = True
   txtQtdaVenda.Visible = True
   txtTotalVenda.Visible = True
   Label1.Visible = True
   Label2.Visible = True
End Sub

Private Sub OptVendaTodas_Click()
   cmdExibir_Click
End Sub
