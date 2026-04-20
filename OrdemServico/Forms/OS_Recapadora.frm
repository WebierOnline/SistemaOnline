VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form OS_Recapadora 
   Caption         =   "ORDEM DE SERVIÇOS"
   ClientHeight    =   9900
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   12735
   ForeColor       =   &H00008000&
   Icon            =   "OS_Recapadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   12645
      TabIndex        =   105
      Top             =   0
      Width           =   12675
      Begin VB.TextBox txtCodOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   11340
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   10860
         TabIndex        =   108
         Top             =   120
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   3120
         Picture         =   "OS_Recapadora.frx":1EDA
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORDEM DE SERVIÇOS"
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
         Left            =   3900
         TabIndex        =   106
         Top             =   120
         Width           =   3360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   109
      Top             =   9630
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11562
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "20:43"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8865
      Left            =   60
      TabIndex        =   110
      Top             =   720
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   15637
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   452
      TabMaxWidth     =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SITUAÇÃO"
      TabPicture(0)   =   "OS_Recapadora.frx":2441
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPecasServicos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblQuantOS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSomaDesconto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "llblTotalSemDesconto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdExcluir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPedidoPDF"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOrcamentoPDF"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdImpGarantia1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdImpPedido1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdImpOrcamento1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdImpEntrada1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdFinanceiroOS"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdNovoOS"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdEditarOS"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "GridPecasServicos"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Grid_OS"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "optFinanceiroAberto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "optFinanceiroFechado"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame6"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "CADASTRO"
      TabPicture(1)   =   "OS_Recapadora.frx":245D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmParecer"
      Tab(1).Control(1)=   "txtCodPedido"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmSecundario"
      Tab(1).Control(3)=   "frmPrincipal"
      Tab(1).Control(4)=   "cmdCancelarEntrada"
      Tab(1).Control(5)=   "cmdAlterar"
      Tab(1).Control(6)=   "cmdApagar"
      Tab(1).Control(7)=   "cmdGerarEntrada"
      Tab(1).Control(8)=   "cmdNovo"
      Tab(1).Control(9)=   "cmdImpEntrada2"
      Tab(1).Control(10)=   "cmdImpOrcamento2"
      Tab(1).Control(11)=   "cmdImpPedido2"
      Tab(1).Control(12)=   "lblDataAberturaCaixa"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "FINANCEIRO"
      TabPicture(2)   =   "OS_Recapadora.frx":2479
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmVendaFechamento"
      Tab(2).Control(1)=   "cmdFinalizarAV"
      Tab(2).Control(2)=   "cmdFinalizarAP"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "CONSULTA"
      TabPicture(3)   =   "OS_Recapadora.frx":2495
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "Grid"
      Tab(3).Control(2)=   "lblQuant"
      Tab(3).Control(3)=   "lblTotalConsulta"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   " "
      TabPicture(4)   =   "OS_Recapadora.frx":24B1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   " "
      TabPicture(5)   =   "OS_Recapadora.frx":24CD
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblQuantFiltro"
      Tab(5).ControlCount=   1
      Begin VB.Frame frmParecer 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Parecer Técnico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   -71880
         TabIndex        =   245
         Top             =   3780
         Visible         =   0   'False
         Width           =   4935
         Begin VB.TextBox txtParecerTecnico 
            Appearance      =   0  'Flat
            Height          =   2295
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   246
            Top             =   240
            Width           =   4815
         End
         Begin ChamaleonBtn.chameleonButton cmdCancelarParecer 
            Height          =   315
            Left            =   3840
            TabIndex        =   247
            Top             =   2580
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Cancelar"
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
            MICON           =   "OS_Recapadora.frx":24E9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdSalvarParecer 
            Height          =   315
            Left            =   2820
            TabIndex        =   248
            Top             =   2580
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Salvar"
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
            MICON           =   "OS_Recapadora.frx":2505
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
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   420
         Left            =   60
         TabIndex        =   112
         Text            =   "ORDEM DE SERVIÇO"
         Top             =   600
         Width           =   12495
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   234
         Top             =   270
         Width           =   4335
         Begin VB.OptionButton optGarantia 
            Caption         =   "Garantia"
            Height          =   195
            Left            =   3300
            TabIndex        =   238
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optOrcamento 
            Caption         =   "Orçamento"
            Height          =   195
            Left            =   2040
            TabIndex        =   237
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton optServico 
            Caption         =   "Serviço"
            Height          =   195
            Left            =   1020
            TabIndex        =   236
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   235
            Top             =   120
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1275
         Left            =   -74880
         TabIndex        =   213
         Top             =   300
         Width           =   12375
         Begin VB.TextBox txtCodClienteLocalizar 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11700
            TabIndex        =   214
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cboLocalizar 
            Height          =   315
            Left            =   7860
            TabIndex        =   102
            Top             =   480
            Visible         =   0   'False
            Width           =   4425
         End
         Begin VB.ComboBox cboConsultaCriterios 
            Height          =   315
            Left            =   6240
            TabIndex        =   101
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox cboConsultaMostrar 
            Height          =   315
            Left            =   1800
            TabIndex        =   98
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox cboConsultaStatus 
            Height          =   315
            Left            =   60
            TabIndex        =   97
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox cboTipoServico 
            Height          =   315
            Left            =   3300
            TabIndex        =   99
            Top             =   480
            Width           =   1515
         End
         Begin VB.ComboBox cboIndice 
            Height          =   315
            Left            =   4860
            TabIndex        =   100
            Top             =   480
            Width           =   1335
         End
         Begin ChamaleonBtn.chameleonButton cmdExibir 
            Height          =   315
            Left            =   8580
            TabIndex        =   103
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
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
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "OS_Recapadora.frx":2521
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdImprimirConsulta 
            Height          =   315
            Left            =   10440
            TabIndex        =   104
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Imprimir"
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
            MICON           =   "OS_Recapadora.frx":253D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label46 
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
            Left            =   6240
            TabIndex        =   219
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label37 
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
            Left            =   1800
            TabIndex        =   218
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
            Left            =   60
            TabIndex        =   217
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label47 
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
            Left            =   3300
            TabIndex        =   216
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label48 
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
            Left            =   4860
            TabIndex        =   215
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.Frame frmVendaFechamento 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5835
         Left            =   -72840
         TabIndex        =   180
         Top             =   1620
         Visible         =   0   'False
         Width           =   7515
         Begin VB.CommandButton Command1 
            Caption         =   "X"
            Height          =   195
            Left            =   7320
            TabIndex        =   211
            Top             =   60
            Width           =   195
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00C0FFFF&
            Height          =   855
            Left            =   3960
            TabIndex        =   208
            Top             =   2880
            Width           =   3435
            Begin VB.TextBox txtRecebido 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Height          =   360
               Left            =   60
               TabIndex        =   83
               Top             =   420
               Width           =   1875
            End
            Begin VB.TextBox txtTroco 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
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
               Height          =   360
               Left            =   1980
               Locked          =   -1  'True
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   420
               Width           =   1335
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Troco:"
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
               Left            =   1980
               TabIndex        =   210
               Top             =   180
               Width           =   570
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recebido:"
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
               TabIndex        =   209
               Top             =   180
               Width           =   885
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Forma Pagamento"
            Height          =   615
            Left            =   3360
            TabIndex        =   207
            Top             =   240
            Width           =   1695
            Begin VB.ComboBox cboTipoPgto 
               Height          =   315
               Left            =   60
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Quant. de Forma de Pgto"
            Height          =   615
            Left            =   5100
            TabIndex        =   206
            Top             =   240
            Width           =   2295
            Begin VB.ComboBox cboQuantForma 
               Height          =   315
               Left            =   60
               TabIndex        =   78
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Timer tmrDebito 
            Enabled         =   0   'False
            Interval        =   150
            Left            =   180
            Top             =   5400
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Usuário"
            Height          =   615
            Left            =   120
            TabIndex        =   205
            Top             =   240
            Width           =   3195
            Begin VB.TextBox txtFuncAP 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   240
               Width           =   2235
            End
            Begin VB.TextBox txtCodFuncAP 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   675
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0FFFF&
            Height          =   1455
            Left            =   120
            TabIndex        =   195
            Top             =   3900
            Width           =   7275
            Begin VB.ComboBox cboformaPgto 
               Height          =   315
               Left            =   4500
               TabIndex        =   88
               Top             =   420
               Width           =   2175
            End
            Begin VB.ComboBox cboFormaPgtoEntrada 
               Height          =   315
               Left            =   1200
               TabIndex        =   86
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtValorParc 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1620
               Locked          =   -1  'True
               TabIndex        =   91
               Text            =   "0"
               Top             =   1020
               Width           =   1155
            End
            Begin VB.ComboBox cboQuantParc 
               Height          =   315
               Left            =   120
               TabIndex        =   89
               Text            =   "1"
               Top             =   1020
               Width           =   735
            End
            Begin VB.TextBox txtValorRest 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3420
               Locked          =   -1  'True
               TabIndex        =   87
               Text            =   "0"
               Top             =   420
               Width           =   1035
            End
            Begin VB.TextBox txtEntrada 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   85
               Text            =   "0"
               Top             =   420
               Width           =   1035
            End
            Begin VB.ComboBox cboPrazo 
               Height          =   315
               Left            =   900
               TabIndex        =   90
               Text            =   "30"
               Top             =   1020
               Width           =   675
            End
            Begin ChamaleonBtn.chameleonButton cmdCal2 
               Height          =   315
               Left            =   3780
               TabIndex        =   93
               Tag             =   "Calendario"
               Top             =   1020
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
               MICON           =   "OS_Recapadora.frx":2559
               PICN            =   "OS_Recapadora.frx":2575
               PICH            =   "OS_Recapadora.frx":48C8
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
               Left            =   2820
               TabIndex        =   92
               Top             =   1020
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskTermino 
               Height          =   315
               Left            =   4080
               TabIndex        =   94
               Top             =   1020
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label lblFormaParcelas 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de Pagamento"
               Height          =   195
               Left            =   4500
               TabIndex        =   204
               Top             =   180
               Width           =   1515
            End
            Begin VB.Label lblFormaEntrada 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de Pagamento"
               Height          =   195
               Left            =   1200
               TabIndex        =   203
               Top             =   180
               Width           =   1515
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Termino:"
               Height          =   195
               Left            =   4080
               TabIndex        =   202
               Top             =   780
               Width           =   615
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Parc.:"
               Height          =   195
               Left            =   1620
               TabIndex        =   201
               Top             =   780
               Width           =   825
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant:"
               Height          =   195
               Left            =   120
               TabIndex        =   200
               Top             =   780
               Width           =   480
            End
            Begin VB.Label lblValorParc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Rest."
               Height          =   195
               Left            =   3420
               TabIndex        =   199
               Top             =   180
               Width           =   780
            End
            Begin VB.Label lblQuantParc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo:"
               Height          =   195
               Left            =   900
               TabIndex        =   198
               Top             =   780
               Width           =   450
            End
            Begin VB.Label lblInicio 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicio:"
               Height          =   195
               Left            =   2820
               TabIndex        =   197
               Top             =   780
               Width           =   420
            End
            Begin VB.Label lblEntrada 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor:"
               Height          =   195
               Left            =   120
               TabIndex        =   196
               Top             =   180
               Width           =   405
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            Height          =   1815
            Left            =   3960
            TabIndex        =   181
            Top             =   960
            Width           =   3435
            Begin VB.TextBox txtAcrescDinheiro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   81
               Top             =   960
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox txtDescDinheiro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   80
               Top             =   600
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox txtDescItens 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   190
               Top             =   1440
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   1440
               ScaleHeight     =   210
               ScaleWidth      =   1035
               TabIndex        =   187
               TabStop         =   0   'False
               Top             =   1020
               Width           =   1035
               Begin VB.OptionButton optAscrescPorc 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "%"
                  Height          =   210
                  Left            =   600
                  TabIndex        =   189
                  TabStop         =   0   'False
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   435
               End
               Begin VB.OptionButton optAscrescRS 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "R$"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   188
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   555
               End
            End
            Begin VB.TextBox txtAcresc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   186
               ToolTipText     =   "Pressiona a tecla ""ENTER"" para desconto em dinheiro."
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtSubTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   79
               TabStop         =   0   'False
               Text            =   "0,00"
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txtDesc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   185
               ToolTipText     =   "Pressiona a tecla ""ENTER"" para desconto em dinheiro."
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txtTotalDesc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   82
               TabStop         =   0   'False
               Text            =   "0,00"
               Top             =   1320
               Width           =   1455
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   1440
               ScaleHeight     =   210
               ScaleWidth      =   1035
               TabIndex        =   182
               TabStop         =   0   'False
               Top             =   660
               Width           =   1035
               Begin VB.OptionButton optDescPorc 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "%"
                  Height          =   210
                  Left            =   600
                  TabIndex        =   184
                  TabStop         =   0   'False
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   435
               End
               Begin VB.OptionButton optDescRS 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "R$"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   183
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   555
               End
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Acresc.:"
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
               Left            =   660
               TabIndex        =   194
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desc.:"
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
               Left            =   780
               TabIndex        =   193
               Top             =   600
               Width           =   570
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
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
               Left            =   1380
               TabIndex        =   192
               Top             =   1380
               Width           =   510
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal:"
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
               Left            =   960
               TabIndex        =   191
               Top             =   300
               Width           =   840
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdCancelar 
            Height          =   315
            Left            =   6540
            TabIndex        =   96
            Top             =   5400
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Cancelar"
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
            MICON           =   "OS_Recapadora.frx":6C1B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdFinalizar 
            Height          =   315
            Left            =   5640
            TabIndex        =   95
            Top             =   5400
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Finalizar"
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
            MICON           =   "OS_Recapadora.frx":6C37
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblInfoDebito 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "POR FAVOR, ENCAMINHE O CLIENTE PARA A GERÊNCIA"
            Height          =   195
            Left            =   120
            TabIndex        =   212
            Top             =   5400
            Visible         =   0   'False
            Width           =   4365
         End
      End
      Begin VB.OptionButton optFinanceiroFechado 
         Caption         =   "Fechado"
         Height          =   195
         Left            =   11520
         TabIndex        =   153
         Top             =   390
         Width           =   975
      End
      Begin VB.OptionButton optFinanceiroAberto 
         Caption         =   "Aberto"
         Height          =   195
         Left            =   10680
         TabIndex        =   152
         Top             =   390
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   -63840
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   7800
         Width           =   1095
      End
      Begin VB.PictureBox frmSecundario 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   7515
         Left            =   -74880
         ScaleHeight     =   7485
         ScaleWidth      =   10305
         TabIndex        =   119
         Top             =   1260
         Width           =   10335
         Begin VB.PictureBox frmTotaisGeral 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1180
            Left            =   7560
            ScaleHeight     =   1155
            ScaleWidth      =   2685
            TabIndex        =   176
            Top             =   6240
            Visible         =   0   'False
            Width           =   2715
            Begin VB.TextBox txtDescGeral 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   179
               TabStop         =   0   'False
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox txtTotalGeral 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1080
               TabIndex        =   178
               TabStop         =   0   'False
               Top             =   780
               Width           =   1575
            End
            Begin VB.TextBox txtSubtotalGeral 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   177
               TabStop         =   0   'False
               Top             =   60
               Width           =   1575
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
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
               Index           =   2
               Left            =   480
               TabIndex        =   244
               Top             =   780
               Width           =   510
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto:"
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
               Index           =   1
               Left            =   105
               TabIndex        =   243
               Top             =   420
               Width           =   885
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Subtotal:"
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
               Index           =   0
               Left            =   210
               TabIndex        =   242
               Top             =   60
               Width           =   780
            End
         End
         Begin TabDlg.SSTab stProdSer 
            Height          =   2280
            Left            =   60
            TabIndex        =   158
            Top             =   1680
            Visible         =   0   'False
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   4022
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabMaxWidth     =   3528
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Serviços"
            TabPicture(0)   =   "OS_Recapadora.frx":6C53
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "frmServicos"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Produtos"
            TabPicture(1)   =   "OS_Recapadora.frx":6C6F
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "frmProdutos"
            Tab(1).ControlCount=   1
            Begin VB.Frame frmProdutos 
               Caption         =   "Produtos"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   -74880
               TabIndex        =   169
               Top             =   360
               Width           =   9915
               Begin VB.ComboBox cboPecas 
                  BackColor       =   &H00C0FFFF&
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   48
                  Top             =   480
                  Width           =   8235
               End
               Begin VB.TextBox txtCodBarra 
                  Height          =   315
                  Left            =   60
                  MaxLength       =   90
                  TabIndex        =   47
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.TextBox txtCodPeca 
                  Appearance      =   0  'Flat
                  Height          =   255
                  Left            =   3960
                  TabIndex        =   170
                  TabStop         =   0   'False
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.TextBox txtValorPeca 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   60
                  Locked          =   -1  'True
                  TabIndex        =   49
                  Top             =   1080
                  Width           =   1035
               End
               Begin VB.TextBox txtDescPecas 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2940
                  TabIndex        =   52
                  Top             =   1080
                  Width           =   915
               End
               Begin VB.TextBox txtSubtotalPecas 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1860
                  Locked          =   -1  'True
                  TabIndex        =   51
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   1035
               End
               Begin VB.TextBox txtTotalPeca 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3900
                  Locked          =   -1  'True
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   1155
               End
               Begin VB.TextBox txtQuantPeca 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1140
                  TabIndex        =   50
                  Top             =   1080
                  Width           =   675
               End
               Begin ChamaleonBtn.chameleonButton cmdRemoverPecas 
                  Height          =   315
                  Left            =   8580
                  TabIndex        =   55
                  Top             =   1440
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Remover"
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
                  MICON           =   "OS_Recapadora.frx":6C8B
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ChamaleonBtn.chameleonButton cmdAdicionarPecas 
                  Height          =   315
                  Left            =   7320
                  TabIndex        =   54
                  Top             =   1440
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Adicionar"
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
                  MICON           =   "OS_Recapadora.frx":6CA7
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total"
                  Height          =   195
                  Index           =   6
                  Left            =   3900
                  TabIndex        =   232
                  Top             =   840
                  Width           =   360
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Desconto"
                  Height          =   195
                  Index           =   5
                  Left            =   2940
                  TabIndex        =   231
                  Top             =   840
                  Width           =   690
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Subtotal"
                  Height          =   195
                  Index           =   4
                  Left            =   1860
                  TabIndex        =   230
                  Top             =   840
                  Width           =   585
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Quant."
                  Height          =   195
                  Index           =   3
                  Left            =   1140
                  TabIndex        =   229
                  Top             =   840
                  Width           =   480
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Index           =   2
                  Left            =   60
                  TabIndex        =   228
                  Top             =   840
                  Width           =   360
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Descrição"
                  Height          =   195
                  Index           =   1
                  Left            =   1620
                  TabIndex        =   227
                  Top             =   240
                  Width           =   720
               End
               Begin VB.Label lblAvisoF2Pecas 
                  AutoSize        =   -1  'True
                  Caption         =   "Pressione a tecla  [ F2 ]  para obter os produtos."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   5520
                  TabIndex        =   172
                  Top             =   240
                  Width           =   4260
                  WordWrap        =   -1  'True
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cód. Barra"
                  Height          =   195
                  Index           =   0
                  Left            =   60
                  TabIndex        =   171
                  Top             =   240
                  Width           =   750
               End
            End
            Begin VB.Frame frmServicos 
               Caption         =   "Serviço"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   120
               TabIndex        =   159
               Top             =   360
               Width           =   9915
               Begin VB.TextBox txtObsServ 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   37
                  Top             =   1080
                  Width           =   9735
               End
               Begin VB.TextBox txtSerie 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   33
                  Top             =   480
                  Width           =   735
               End
               Begin VB.TextBox txtFogo 
                  Height          =   315
                  Left            =   2580
                  TabIndex        =   34
                  Top             =   480
                  Width           =   735
               End
               Begin VB.ComboBox cboTipo 
                  Height          =   315
                  Left            =   60
                  Sorted          =   -1  'True
                  TabIndex        =   32
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.ComboBox cboMarca 
                  Height          =   315
                  Left            =   4860
                  Sorted          =   -1  'True
                  TabIndex        =   36
                  Top             =   480
                  Width           =   2115
               End
               Begin VB.TextBox txtDote 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   35
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.TextBox mskValorServicoAuto 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4020
                  TabIndex        =   39
                  Top             =   1080
                  Width           =   1155
               End
               Begin VB.TextBox txtSubTotalServicoAuto 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6000
                  Locked          =   -1  'True
                  TabIndex        =   41
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.TextBox txtDescServicoAuto 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7380
                  TabIndex        =   42
                  Top             =   1080
                  Width           =   1035
               End
               Begin VB.TextBox txtTotalServicoAuto 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8460
                  Locked          =   -1  'True
                  TabIndex        =   43
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.TextBox txtQuantServicoAuto 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5220
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.TextBox txtCodServicoAuto 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   3300
                  TabIndex        =   160
                  TabStop         =   0   'False
                  Top             =   900
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.ComboBox cboServicosAuto 
                  Height          =   315
                  Left            =   60
                  Sorted          =   -1  'True
                  TabIndex        =   38
                  Top             =   1080
                  Width           =   3915
               End
               Begin ChamaleonBtn.chameleonButton cmdRemoverServicosAuto 
                  Height          =   315
                  Left            =   8820
                  TabIndex        =   45
                  Top             =   1440
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "R&emover"
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
                  MICON           =   "OS_Recapadora.frx":6CC3
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ChamaleonBtn.chameleonButton cmdAdicionarServicosAuto 
                  Height          =   315
                  Left            =   7800
                  TabIndex        =   44
                  Top             =   1440
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "A&dicionar"
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
                  MICON           =   "OS_Recapadora.frx":6CDF
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label lblObsServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Obs.:"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   226
                  Top             =   840
                  Width           =   375
               End
               Begin VB.Label lblSerieServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Série"
                  Height          =   195
                  Left            =   1800
                  TabIndex        =   225
                  Top             =   240
                  Width           =   360
               End
               Begin VB.Label lblFogoServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fogo"
                  Height          =   195
                  Left            =   2580
                  TabIndex        =   224
                  Top             =   240
                  Width           =   360
               End
               Begin VB.Label lblTipoServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   223
                  Top             =   240
                  Width           =   315
               End
               Begin VB.Label lblMarca 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fabricante"
                  Height          =   195
                  Left            =   4860
                  TabIndex        =   168
                  Top             =   240
                  Width           =   750
               End
               Begin VB.Label lblDote 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Dote No."
                  Height          =   195
                  Left            =   3360
                  TabIndex        =   167
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lblTotalServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total"
                  Height          =   195
                  Left            =   8460
                  TabIndex        =   166
                  Top             =   840
                  Width           =   360
               End
               Begin VB.Label lblDescServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Desconto"
                  Height          =   195
                  Left            =   7380
                  TabIndex        =   165
                  Top             =   840
                  Width           =   690
               End
               Begin VB.Label lblSubTotalServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Subtotal"
                  Height          =   195
                  Left            =   6000
                  TabIndex        =   164
                  Top             =   840
                  Width           =   585
               End
               Begin VB.Label lblDescricaoServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Serviços:"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   163
                  Top             =   840
                  Width           =   660
               End
               Begin VB.Label lblQuantServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Quant:"
                  Height          =   195
                  Left            =   5220
                  TabIndex        =   162
                  Top             =   840
                  Width           =   480
               End
               Begin VB.Label lblValorServ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   4020
                  TabIndex        =   161
                  Top             =   840
                  Width           =   360
               End
            End
         End
         Begin VB.Frame frmParecerCliente 
            Caption         =   "Parecer do Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   60
            TabIndex        =   143
            Top             =   1680
            Width           =   10155
            Begin VB.TextBox txtPareceCliente 
               Appearance      =   0  'Flat
               Height          =   1455
               Left            =   60
               MultiLine       =   -1  'True
               TabIndex        =   56
               Top             =   240
               Width           =   9975
            End
         End
         Begin VB.TextBox txtObs 
            Appearance      =   0  'Flat
            Height          =   915
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   46
            Top             =   1500
            Visible         =   0   'False
            Width           =   10155
         End
         Begin VB.CheckBox chkVeiculo 
            Caption         =   "Mostrar Veículo"
            Height          =   315
            Left            =   9300
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   300
            Width           =   915
         End
         Begin VB.Frame frmSituacao 
            Caption         =   "Situação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   5160
            TabIndex        =   149
            Top             =   3540
            Width           =   5055
            Begin ChamaleonBtn.chameleonButton ccmdIncluirSituacao 
               Height          =   315
               Left            =   2400
               TabIndex        =   157
               Top             =   480
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "+"
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
               MICON           =   "OS_Recapadora.frx":6CFB
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboSituacao 
               Height          =   315
               Left            =   60
               TabIndex        =   61
               Top             =   480
               Width           =   2355
            End
            Begin VB.TextBox txtCodSituacao 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1920
               TabIndex        =   150
               Top             =   180
               Visible         =   0   'False
               Width           =   675
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Situacao 
               Height          =   2295
               Left            =   2700
               TabIndex        =   63
               Top             =   180
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   4048
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverSituacao 
               Height          =   315
               Left            =   1380
               TabIndex        =   64
               Top             =   900
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "R&emover"
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
               MICON           =   "OS_Recapadora.frx":6D17
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarSituacao 
               Height          =   315
               Left            =   60
               TabIndex        =   62
               Top             =   900
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "A&dicionar"
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
               MICON           =   "OS_Recapadora.frx":6D33
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Item"
               Height          =   195
               Left            =   60
               TabIndex        =   151
               Top             =   240
               Width           =   300
            End
         End
         Begin VB.Frame frmAcessorios 
            Caption         =   "Acessórios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   60
            TabIndex        =   144
            Top             =   3540
            Width           =   5055
            Begin ChamaleonBtn.chameleonButton ccmdIncluirAcessório 
               Height          =   315
               Left            =   2400
               TabIndex        =   156
               Top             =   480
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "+"
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
               MICON           =   "OS_Recapadora.frx":6D4F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox txtCodAcessorio 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1920
               TabIndex        =   145
               Top             =   180
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.ComboBox cboAcessorios 
               Height          =   315
               Left            =   60
               TabIndex        =   57
               Top             =   480
               Width           =   2355
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Acessorio 
               Height          =   2295
               Left            =   2700
               TabIndex        =   60
               Top             =   180
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   4048
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverAcessorios 
               Height          =   315
               Left            =   1380
               TabIndex        =   59
               Top             =   900
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "R&emover"
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
               MICON           =   "OS_Recapadora.frx":6D6B
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarAcessorios 
               Height          =   315
               Left            =   60
               TabIndex        =   58
               Top             =   900
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "A&dicionar"
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
               MICON           =   "OS_Recapadora.frx":6D87
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Item"
               Height          =   195
               Left            =   60
               TabIndex        =   146
               Top             =   240
               Width           =   300
            End
         End
         Begin VB.Frame frmEquipamento 
            Caption         =   "Veículo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   60
            TabIndex        =   134
            Top             =   660
            Width           =   10155
            Begin VB.TextBox txtChassi 
               Height          =   315
               Left            =   6150
               TabIndex        =   28
               Top             =   540
               Width           =   1755
            End
            Begin VB.ComboBox cboModelo 
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   1962
               TabIndex        =   24
               Top             =   540
               Width           =   1575
            End
            Begin VB.ComboBox cboFabricante 
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   60
               TabIndex        =   23
               Top             =   540
               Width           =   1875
            End
            Begin VB.ComboBox cboTanque 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   9160
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   540
               Width           =   915
            End
            Begin VB.ComboBox cboCor 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   7932
               TabIndex        =   29
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox txtKm 
               Height          =   315
               Left            =   5328
               TabIndex        =   27
               Top             =   540
               Width           =   795
            End
            Begin VB.TextBox txtPlaca 
               Height          =   315
               Left            =   4386
               TabIndex        =   26
               Top             =   540
               Width           =   915
            End
            Begin VB.TextBox txtAno 
               Height          =   315
               Left            =   3564
               TabIndex        =   25
               Top             =   540
               Width           =   795
            End
            Begin VB.Label lblChassi 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Chassi:"
               Height          =   195
               Left            =   6150
               TabIndex        =   249
               Top             =   300
               Width           =   510
            End
            Begin VB.Label lblFabricante 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fabricante:"
               Height          =   195
               Left            =   60
               TabIndex        =   137
               Top             =   300
               Width           =   795
            End
            Begin VB.Label lblModelo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Modelo:"
               Height          =   195
               Left            =   1962
               TabIndex        =   136
               Top             =   300
               Width           =   570
            End
            Begin VB.Label lblTanque 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanque"
               Height          =   195
               Left            =   9160
               TabIndex        =   142
               Top             =   300
               Width           =   555
            End
            Begin VB.Label lblCor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cor"
               Height          =   195
               Left            =   7932
               TabIndex        =   141
               Top             =   300
               Width           =   240
            End
            Begin VB.Label lblKM 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "KM:"
               Height          =   195
               Left            =   5328
               TabIndex        =   139
               Top             =   300
               Width           =   285
            End
            Begin VB.Label lblPlaca 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Placa:"
               Height          =   195
               Left            =   4386
               TabIndex        =   138
               Top             =   300
               Width           =   450
            End
            Begin VB.Label lblAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   3564
               TabIndex        =   135
               Top             =   300
               Width           =   330
            End
         End
         Begin VB.TextBox txtCodCliente 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   6300
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtCodFuncionario 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   1200
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboCliente 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2400
            TabIndex        =   21
            Top             =   300
            Width           =   6855
         End
         Begin VB.ComboBox cboFuncionario 
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   300
            Width           =   2295
         End
         Begin VB.Frame frmGridServicos 
            Caption         =   "Serviços Adicionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   60
            TabIndex        =   124
            Top             =   3420
            Width           =   10155
            Begin MSFlexGridLib.MSFlexGrid Grid_Servicos 
               Height          =   2115
               Left            =   60
               TabIndex        =   65
               Top             =   600
               Width           =   10035
               _ExtentX        =   17701
               _ExtentY        =   3731
               _Version        =   393216
               WordWrap        =   -1  'True
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.PictureBox frmTotaisProdServ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1180
            Left            =   60
            ScaleHeight     =   1155
            ScaleWidth      =   2625
            TabIndex        =   120
            Top             =   6240
            Visible         =   0   'False
            Width           =   2655
            Begin VB.TextBox txtQuantGeral 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   900
               TabIndex        =   175
               TabStop         =   0   'False
               Top             =   780
               Width           =   435
            End
            Begin VB.TextBox txtQuantPecas 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   900
               Locked          =   -1  'True
               TabIndex        =   174
               TabStop         =   0   'False
               Top             =   420
               Width           =   435
            End
            Begin VB.TextBox txtQuantServicos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   900
               Locked          =   -1  'True
               TabIndex        =   173
               TabStop         =   0   'False
               Top             =   60
               Width           =   435
            End
            Begin VB.TextBox txtTotalServicos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   123
               TabStop         =   0   'False
               Top             =   60
               Width           =   1215
            End
            Begin VB.TextBox txtTotalPecasServicos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1380
               TabIndex        =   122
               TabStop         =   0   'False
               Top             =   780
               Width           =   1215
            End
            Begin VB.TextBox txtTotalPecas 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Subtotal:"
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
               Index           =   5
               Left            =   60
               TabIndex        =   241
               Top             =   780
               Width           =   780
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peças:"
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
               Index           =   4
               Left            =   255
               TabIndex        =   240
               Top             =   420
               Width           =   585
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Serviços:"
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
               Index           =   3
               Left            =   30
               TabIndex        =   239
               Top             =   60
               Width           =   810
            End
         End
         Begin VB.Label lblValidade 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4920
            TabIndex        =   129
            Top             =   6660
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Left            =   2400
            TabIndex        =   128
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recepcionista"
            Height          =   195
            Left            =   60
            TabIndex        =   127
            Top             =   60
            Width           =   1020
         End
      End
      Begin VB.PictureBox frmPrincipal 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   -74880
         ScaleHeight     =   765
         ScaleWidth      =   10305
         TabIndex        =   31
         Top             =   360
         Width           =   10335
         Begin VB.ComboBox cboTipoOS 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4200
            TabIndex        =   13
            Top             =   300
            Width           =   2115
         End
         Begin VB.ComboBox cboStatus 
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   300
            Width           =   1755
         End
         Begin VB.ComboBox cboMecanico 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1860
            TabIndex        =   12
            Top             =   300
            Width           =   2295
         End
         Begin VB.TextBox txtCodMecanico 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton1 
            Height          =   315
            Left            =   9240
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   300
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
            MICON           =   "OS_Recapadora.frx":6DA3
            PICN            =   "OS_Recapadora.frx":6DBF
            PICH            =   "OS_Recapadora.frx":9112
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCal1 
            Height          =   315
            Left            =   7260
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   300
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
            MICON           =   "OS_Recapadora.frx":B465
            PICN            =   "OS_Recapadora.frx":B481
            PICH            =   "OS_Recapadora.frx":D7D4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskDataSaida 
            Height          =   315
            Left            =   8280
            TabIndex        =   17
            Top             =   300
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648384
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskHoraSaida 
            Height          =   315
            Left            =   9600
            TabIndex        =   19
            Top             =   300
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648384
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataEntrada 
            Height          =   315
            Left            =   6360
            TabIndex        =   14
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskHoraEntrada 
            Height          =   315
            Left            =   7620
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   300
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saída (Previsão)"
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
            Left            =   8280
            TabIndex        =   118
            Top             =   60
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6360
            TabIndex        =   117
            Top             =   60
            Width           =   675
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Serviço"
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
            Left            =   4200
            TabIndex        =   116
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   195
            Left            =   60
            TabIndex        =   115
            Top             =   60
            Width           =   450
         End
         Begin VB.Label lblMecanico 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1860
            TabIndex        =   114
            Top             =   60
            Width           =   930
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000008&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   420
         Left            =   120
         TabIndex        =   111
         Text            =   "PEÇAS / SERVIÇOS"
         Top             =   5700
         Width           =   12375
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_OS 
         Height          =   3915
         Left            =   60
         TabIndex        =   0
         Top             =   1020
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6906
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarEntrada 
         Height          =   615
         Left            =   -64485
         TabIndex        =   67
         Top             =   1740
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Cancelar"
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
         MICON           =   "OS_Recapadora.frx":FB27
         PICN            =   "OS_Recapadora.frx":FB43
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   615
         Left            =   -64485
         TabIndex        =   68
         Top             =   2400
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Alterar"
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
         MICON           =   "OS_Recapadora.frx":118D5
         PICN            =   "OS_Recapadora.frx":118F1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdApagar 
         Height          =   615
         Left            =   -64485
         TabIndex        =   69
         Top             =   3060
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Excluir"
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
         MICON           =   "OS_Recapadora.frx":13683
         PICN            =   "OS_Recapadora.frx":1369F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdGerarEntrada 
         Height          =   615
         Left            =   -64500
         TabIndex        =   66
         Top             =   1080
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Salvar"
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
         MICON           =   "OS_Recapadora.frx":15431
         PICN            =   "OS_Recapadora.frx":1544D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   615
         Left            =   -64500
         TabIndex        =   10
         Top             =   360
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Novo"
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
         MICON           =   "OS_Recapadora.frx":171DF
         PICN            =   "OS_Recapadora.frx":171FB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridPecasServicos 
         Height          =   2415
         Left            =   120
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   6120
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   4260
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdEditarOS 
         Height          =   375
         Left            =   1020
         TabIndex        =   2
         Top             =   5100
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Alterar"
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
         MICON           =   "OS_Recapadora.frx":18F8D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovoOS 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   5100
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Novo"
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
         MICON           =   "OS_Recapadora.frx":18FA9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdFinanceiroOS 
         Height          =   375
         Left            =   1860
         TabIndex        =   3
         Top             =   5100
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Financeiro"
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
         MICON           =   "OS_Recapadora.frx":18FC5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImpEntrada2 
         Height          =   615
         Left            =   -64500
         TabIndex        =   70
         Top             =   4260
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Entrada"
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
         MICON           =   "OS_Recapadora.frx":18FE1
         PICN            =   "OS_Recapadora.frx":18FFD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImpOrcamento2 
         Height          =   615
         Left            =   -64500
         TabIndex        =   71
         Top             =   4920
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Orçamento"
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
         MICON           =   "OS_Recapadora.frx":19317
         PICN            =   "OS_Recapadora.frx":19333
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImpEntrada1 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   5100
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Imprimir Entrada"
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
         MICON           =   "OS_Recapadora.frx":1964D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImpOrcamento1 
         Height          =   375
         Left            =   4380
         TabIndex        =   5
         Top             =   5100
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Imprimir Orçamento"
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
         MICON           =   "OS_Recapadora.frx":19669
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImpPedido1 
         Height          =   375
         Left            =   7260
         TabIndex        =   7
         Top             =   5100
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Imprimir Pedido"
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
         MICON           =   "OS_Recapadora.frx":19685
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImpPedido2 
         Height          =   615
         Left            =   -64500
         TabIndex        =   72
         Top             =   5580
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Pedido"
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
         MICON           =   "OS_Recapadora.frx":196A1
         PICN            =   "OS_Recapadora.frx":196BD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImpGarantia1 
         Height          =   375
         Left            =   9840
         TabIndex        =   9
         Top             =   5100
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Garantia"
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
         MICON           =   "OS_Recapadora.frx":199D7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdOrcamentoPDF 
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   5100
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Orçamento PDF"
         ENAB            =   0   'False
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
         MICON           =   "OS_Recapadora.frx":199F3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdPedidoPDF 
         Height          =   375
         Left            =   8580
         TabIndex        =   8
         Top             =   5100
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Pedido PDF"
         ENAB            =   0   'False
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
         MICON           =   "OS_Recapadora.frx":19A0F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdFinalizarAV 
         Height          =   555
         Left            =   -74640
         TabIndex        =   73
         Top             =   540
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Venda à Vista (F10)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "OS_Recapadora.frx":19A2B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdFinalizarAP 
         Height          =   555
         Left            =   -72120
         TabIndex        =   74
         Top             =   540
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Venda à Prazo (F12)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "OS_Recapadora.frx":19A47
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
         Height          =   6855
         Left            =   -74880
         TabIndex        =   220
         Top             =   1680
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   12091
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluir 
         Height          =   375
         Left            =   10740
         TabIndex        =   233
         Top             =   5100
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Excluir"
         ENAB            =   0   'False
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
         MICON           =   "OS_Recapadora.frx":19A63
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblQuant 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "00"
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
         Left            =   -74880
         TabIndex        =   222
         Top             =   8580
         Width           =   225
      End
      Begin VB.Label lblTotalConsulta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "00"
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
         Left            =   -62760
         TabIndex        =   221
         Top             =   8580
         Width           =   225
      End
      Begin VB.Label lblDataAberturaCaixa 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -63960
         TabIndex        =   155
         Top             =   6600
         Width           =   1035
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Financeiro:"
         Height          =   195
         Left            =   9780
         TabIndex        =   154
         Top             =   390
         Width           =   780
      End
      Begin VB.Label llblTotalSemDesconto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "00"
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
         Left            =   8880
         TabIndex        =   148
         Top             =   8540
         Width           =   225
      End
      Begin VB.Label lblSomaDesconto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "00"
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
         Left            =   9840
         TabIndex        =   147
         Top             =   8540
         Width           =   225
      End
      Begin VB.Label lblQuantFiltro 
         AutoSize        =   -1  'True
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
         Left            =   -74880
         TabIndex        =   133
         Top             =   8040
         Width           =   75
      End
      Begin VB.Label lblQuantOS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "00"
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
         Left            =   12300
         TabIndex        =   132
         Top             =   4980
         Width           =   225
      End
      Begin VB.Label lblPecasServicos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "00"
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
         Left            =   10980
         TabIndex        =   131
         Top             =   8540
         Width           =   225
      End
   End
   Begin VB.Menu menu_Cadastrk 
      Caption         =   "&Cadastro"
      Begin VB.Menu menu_Cadastro_Cliente 
         Caption         =   "Cli&ente"
      End
      Begin VB.Menu menu_Cadastro_Pecas 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu menu_Cadastro_Servicos 
         Caption         =   "&Serviços"
      End
      Begin VB.Menu menu_Cadastro_Pneus 
         Caption         =   "Pneus"
      End
      Begin VB.Menu Menu_Cadastro_Acessorios 
         Caption         =   "Acessórios"
      End
      Begin VB.Menu Menu_Cadastro_Situacoes 
         Caption         =   "Situações"
      End
      Begin VB.Menu Menu_Cadastro_Parecer 
         Caption         =   "Parecer Técnico"
      End
   End
   Begin VB.Menu menu_Impressao 
      Caption         =   "&Impressão"
      Begin VB.Menu menu_Impressao_Entrada 
         Caption         =   "&Entrada"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_Impressao_Orcamento 
         Caption         =   "&Orçamento"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_Impressao_Pedido 
         Caption         =   "&Pedido"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_Impressao_Garantia 
         Caption         =   "&Garantia"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "OS_Recapadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim r_Itens As ADODB.Recordset
Dim sSQL_Itens As String
Dim vTabelaServicos As String
Public codPedido As String
Dim rFunc As ADODB.Recordset
Dim rTecnico As ADODB.Recordset
Dim rPedido As ADODB.Recordset
Dim rOS As ADODB.Recordset
Dim rCliente As ADODB.Recordset
Dim rEquip As ADODB.Recordset

Dim xParc As Long, xAcess As Long
Dim xPeca As Long, xServ As Long, xPecaItem As Long

Dim w As Long
Dim vCodPedido As Long
Dim numCol As Integer
Dim numRow As Integer

Dim Texto As String         'usado pra preencher os combos
Dim i, Posicao As Integer   'usado pra preencher os combos
Dim Posicionar As Boolean   'usado pra preencher os combos

Dim OS_FECHADA As Boolean
Dim OS_FINANCEIROABERTO As Boolean
Dim VERIFICAR_QUANTIDADE As Boolean
Dim CAIXA_FECHADO As Boolean

Public oCfg As ConfigItem
Dim bConfFechAP As Boolean
Dim iCopiasAP As Integer
Dim bEntregaAP As Boolean
Dim bImprAP As Integer
Dim bConfImprAP As Boolean
Dim vTabelaServico As String
Dim printSQL As String

Dim NumCopias As Integer
Dim ii As Integer
Dim lNovoCod As Long

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim r As ADODB.Recordset
Dim sSQL As String

Public cCfg As ConfigItem       'arquivo .ini
Public oIni As Ini              'arquivo .ini
Dim var_ImpNormal As String

Dim NFCe_OK As Boolean
Dim PararFechamentoVenda As Boolean

'desconto
Public vTipoDesc As String
Public vLimitarDesc As String
Public vValorDescFixoAV As String
Public vValorDescFixoAP As String
Public vValorDescGradualAV1 As String
Public vValorDescGradualAP1 As String
Public vValorDescGradualAV2 As String
Public vValorDescGradualAP2 As String
Public vValorDescGradualAV3 As String
Public vValorDescGradualAP3 As String
Public vMargemDescGradual1 As String
Public vMargemDescGradual2 As String
Public vMargemDescGradual3 As String
Dim vEtapa As Integer
Public bFechAP As Boolean       'impressão aprazo
Dim Passou_Limite As Boolean
Dim Cliente_Debito As Boolean
Public bFechAV As Boolean

Private Sub AutoNumeracao_Pedido()
sSQL = "SELECT ISNULL(MAX(cod_pedido), 0) AS ultimo_Ped FROM pedidos;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtCodPedido.Text = Format(r("ultimo_Ped") + 1, "000000")
'If r.State <> 0 Then r.Close
'Set r = Nothing
End Sub

Private Sub CalcularTotalServicoAuto()
If txtQuantServicoAuto.Text = "" Then txtQuantServicoAuto.Text = 1
If mskValorServicoAuto.Text = "" Then mskValorServicoAuto.Text = Format(0, ocMONEY)

Dim vQuantServ As Integer
Dim vValorServ As Currency
Dim vSubtotalServ As Currency

vQuantServ = txtQuantServicoAuto.Text
vValorServ = mskValorServicoAuto.Text
vSubtotalServ = vValorServ * vQuantServ

txtSubTotalServicoAuto.Text = Format(vSubtotalServ, ocMONEY)

Dim vDescServico As Currency
Dim vTotalServico As Currency

If txtDescServicoAuto.Text = "" Then txtDescServicoAuto.Text = 0
vDescServico = txtDescServicoAuto.Text
vTotalServico = vSubtotalServ - vDescServico
txtTotalServicoAuto.Text = Format(vTotalServico, ocMONEY)

End Sub

Private Sub Calcular_Prazo()
If cboPrazo.Text = "" Then Exit Sub

If txtEntrada.Text = "0,00" And cboQuantParc.Text = "1" Then
   mskInicio.Text = Format(DateAdd("d", cboPrazo, Date), "dd/mm/yy")
   mskTermino.Text = Format(DateAdd("d", cboPrazo, Date), "dd/mm/yy")

ElseIf txtEntrada.Text = "0,00" And cboQuantParc.Text > "1" Then
   mskInicio.Text = Format(DateAdd("d", cboPrazo, Date), "dd/mm/yy")
   mskTermino.Text = Format(DateAdd("d", cboPrazo * (cboQuantParc.Text), Date), "dd/mm/yy")

ElseIf txtEntrada.Text <> "0,00" And cboQuantParc.Text = "1" Then
   mskInicio.Text = Format(Date, "dd/mm/yy")
   mskTermino.Text = Format(DateAdd("d", cboPrazo * (cboQuantParc.Text), Date), "dd/mm/yy")

ElseIf txtEntrada.Text <> "0,00" And cboQuantParc.Text > "1" Then
   mskInicio.Text = Format(Date, "dd/mm/yy")
   mskTermino.Text = Format(DateAdd("d", cboPrazo * (cboQuantParc.Text), Date), "dd/mm/yy")

End If
End Sub


Private Sub CalcularValorPeca()
Dim vValorPeca As Currency
Dim vQuantPeca As Double
Dim vSubtotalPeca As Currency

If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1: vQuantPeca = 1 Else vQuantPeca = txtQuantPeca.Text
If txtValorPeca.Text = "" Then txtTotalPeca.Text = "0,00": vValorPeca = 0 Else vValorPeca = txtValorPeca.Text

vSubtotalPeca = vValorPeca * vQuantPeca
txtSubtotalPecas.Text = Format(vSubtotalPeca, ocMONEY)

Dim vDescPeca As Currency
Dim vTotalPeca As Currency

If txtDescPecas.Text = "" Then txtDescPecas.Text = Format(0, ocMONEY)

vDescPeca = txtDescPecas.Text
vTotalPeca = vSubtotalPeca - vDescPeca
txtTotalPeca = Format(vTotalPeca, ocMONEY)
End Sub

Private Sub ExibirObjetosServicos()
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    lblMarca.Visible = False
    lblDote.Visible = False
    cboMarca.Visible = False
    txtDote.Visible = False
    cboTipo.Visible = False
    txtSerie.Visible = False
    txtFogo.Visible = False
    lblFogoServ.Visible = False
    lblSerieServ.Visible = False
    lblTipoServ.Visible = False
    lblObsServ.Visible = False
    txtObsServ.Visible = False
    txtObs.Visible = False
    cmdImpEntrada2.Visible = True
    
    lblDescricaoServ.Left = 60
    lblValorServ.Left = 4020
    lblQuantServ.Left = 5220
    lblSubTotalServ.Left = 6000
    lblDescServ.Left = 7380
    lblTotalServ.Left = 8460
    
    lblDescricaoServ.Top = 240
    lblValorServ.Top = 240
    lblQuantServ.Top = 240
    lblSubTotalServ.Top = 240
    lblDescServ.Top = 240
    lblTotalServ.Top = 240
    
    cboServicosAuto.Left = 60
    mskValorServicoAuto.Left = 4020
    txtQuantServicoAuto.Left = 5220
    txtSubTotalServicoAuto.Left = 6000
    txtDescServicoAuto.Left = 7380
    txtTotalServicoAuto.Left = 8460
    
    cboServicosAuto.Top = 480
    mskValorServicoAuto.Top = 480
    txtQuantServicoAuto.Top = 480
    txtSubTotalServicoAuto.Top = 480
    txtDescServicoAuto.Top = 480
    txtTotalServicoAuto.Top = 480
    
    cmdAdicionarServicosAuto.Left = 7800
    cmdRemoverServicosAuto.Left = 8820
    
    cmdAdicionarServicosAuto.Top = 840
    cmdRemoverServicosAuto.Top = 840
    txtTotalServicoAuto.Width = 1335
    frmServicos.Caption = "Serviços"
ElseIf vTipoOS = "Comunicação Visual" Then
    lblMarca.Visible = False
    lblDote.Visible = False
    cboMarca.Visible = False
    txtDote.Visible = False
    cboTipo.Visible = False
    txtSerie.Visible = False
    txtFogo.Visible = False
    lblFogoServ.Visible = False
    lblSerieServ.Visible = False
    lblTipoServ.Visible = False
    lblObsServ.Visible = True
    txtObsServ.Visible = True
    cmdImpEntrada2.Visible = False
        
    lblDescricaoServ.Left = 60
    lblValorServ.Left = 4020
    lblQuantServ.Left = 5220
    lblSubTotalServ.Left = 6000
    lblDescServ.Left = 7380
    lblTotalServ.Left = 8460
    
    lblDescricaoServ.Top = 240
    lblValorServ.Top = 240
    lblQuantServ.Top = 240
    lblSubTotalServ.Top = 240
    lblDescServ.Top = 240
    lblTotalServ.Top = 240
    
    cboServicosAuto.Left = 60
    mskValorServicoAuto.Left = 4020
    txtQuantServicoAuto.Left = 5220
    txtSubTotalServicoAuto.Left = 6000
    txtDescServicoAuto.Left = 7380
    txtTotalServicoAuto.Left = 8460
    
    cboServicosAuto.Top = 480
    mskValorServicoAuto.Top = 480
    txtQuantServicoAuto.Top = 480
    txtSubTotalServicoAuto.Top = 480
    txtDescServicoAuto.Top = 480
    txtTotalServicoAuto.Top = 480
    
    cmdAdicionarServicosAuto.Left = 7800
    cmdRemoverServicosAuto.Left = 8820
    
    cmdAdicionarServicosAuto.Top = 1440
    cmdRemoverServicosAuto.Top = 1440
    txtTotalServicoAuto.Width = 1335
    frmServicos.Caption = "Serviços"
ElseIf vTipoOS = "Recapadora" Then
    lblMarca.Visible = True
    lblDote.Visible = True
    cboMarca.Visible = True
    txtDote.Visible = True
    cboTipo.Visible = True
    txtSerie.Visible = True
    txtFogo.Visible = True
    lblFogoServ.Visible = True
    lblSerieServ.Visible = True
    lblTipoServ.Visible = True
    lblObsServ.Visible = False
    txtObsServ.Visible = False
    cmdImpEntrada2.Visible = True
        
    lblDescricaoServ.Left = 60
    lblValorServ.Left = 4020
    lblQuantServ.Left = 5220
    lblSubTotalServ.Left = 6000
    lblDescServ.Left = 7380
    lblTotalServ.Left = 8460
    
    lblDescricaoServ.Top = 840
    lblValorServ.Top = 840
    lblQuantServ.Top = 840
    lblSubTotalServ.Top = 840
    lblDescServ.Top = 840
    lblTotalServ.Top = 840
    
    cboServicosAuto.Left = 60
    mskValorServicoAuto.Left = 4020
    txtQuantServicoAuto.Left = 5220
    txtSubTotalServicoAuto.Left = 6000
    txtDescServicoAuto.Left = 7380
    txtTotalServicoAuto.Left = 8460
    
    cboServicosAuto.Top = 1080
    mskValorServicoAuto.Top = 1080
    txtQuantServicoAuto.Top = 1080
    txtSubTotalServicoAuto.Top = 1080
    txtDescServicoAuto.Top = 1080
    txtTotalServicoAuto.Top = 1080
    
    cmdAdicionarServicosAuto.Left = 7800
    cmdRemoverServicosAuto.Left = 8820
    
    cmdAdicionarServicosAuto.Top = 1440
    cmdRemoverServicosAuto.Top = 1440
    txtTotalServicoAuto.Width = 1335
    frmServicos.Caption = "Pneus/Serviços"
End If

End Sub

Private Sub LimparObjetos_Equipamento()
cboFabricante.Text = ""
cboModelo.Text = ""
txtAno.Text = ""
txtPlaca.Text = ""
txtKM.Text = ""
txtChassi.Text = ""
cboCor.Text = ""
cboTanque.Text = ""
txtPareceCliente.Text = ""
End Sub

Private Sub LimparTotais()
txtQuantServicos.Text = "0"
txtQuantPecas.Text = "0"
txtQuantGeral.Text = "0"
txtTotalServicos.Text = Format(0, ocMONEY)
txtTotalPecas.Text = Format(0, ocMONEY)
txtTotalPecasServicos.Text = Format(0, ocMONEY)
txtSubtotalGeral.Text = Format(0, ocMONEY)
txtDescGeral.Text = Format(0, ocMONEY)
txtTotalGeral.Text = Format(0, ocMONEY)
End Sub

Private Sub Mostrar_Equipamento_Automoveis()
lblFabricante.Visible = True
lblModelo.Visible = True
lblAno.Visible = True
lblPlaca.Visible = True
lblKM.Visible = True
lblChassi.Visible = True
lblCor.Visible = True
lblTanque.Visible = True
cboFabricante.Visible = True
cboModelo.Visible = True
txtAno.Visible = True
txtPlaca.Visible = True
txtKM.Visible = True
txtChassi.Visible = True
cboCor.Visible = True
cboTanque.Visible = True

cboFabricante.Left = 60
cboModelo.Left = 1962
txtAno.Left = 3564
txtPlaca.Left = 4386
txtKM.Left = 5328
txtChassi.Left = 6150
cboCor.Left = 7932
cboTanque.Left = 9180

'lblFabricante.Top = 300
'lblFabricante.Left = 60
'cboFabricante.Top = 540
'cboFabricante.Left = 60

'lblModelo.Top = 300
'lblModelo.Left = 2640
'cboModelo.Top = 540
'cboModelo.Left = 2640

'lblTanque.Top = 300
'lblTanque.Left = 9240
'cboTanque.Top = 540
'cboTanque.Left = 9240
'cboTanque.Width = 855
lblTanque.Caption = "Tanque"
End Sub

Private Sub Mostrar_Equipamentos_Informatica()
lblFabricante.Visible = True
lblModelo.Visible = True
lblAno.Visible = False
lblPlaca.Visible = False
lblKM.Visible = False
lblCor.Visible = False
lblTanque.Visible = True
cboFabricante.Visible = True
cboModelo.Visible = True
txtAno.Visible = False
txtPlaca.Visible = False
txtKM.Visible = False
cboCor.Visible = False
cboTanque.Visible = True

lblFabricante.Top = 300
lblFabricante.Left = 3840
cboFabricante.Top = 540
cboFabricante.Left = 3840

lblModelo.Top = 300
lblModelo.Left = 6420
cboModelo.Top = 540
cboModelo.Left = 6420

lblTanque.Top = 300
lblTanque.Left = 60
lblTanque.Caption = "Equipamento"
cboTanque.Top = 540
cboTanque.Left = 60
cboTanque.Width = 3735
End Sub

Private Sub Mostrar_ValorRestante()
Dim Valor As Currency
Dim QUANT As Integer
Dim Entrada As Currency
Dim RESULTADO As Currency
Dim VALOR_SENTRADA As Currency

If txtEntrada.Text = "" Then Entrada = 0 Else Entrada = txtEntrada.Text
If txtTotalDesc.Text = "" Then Valor = 0 Else Valor = txtTotalDesc.Text
' QUANT = txtQuantParc.Text

VALOR_SENTRADA = Valor - Entrada
txtValorRest.Text = Format(VALOR_SENTRADA, "##,##0.00")
End Sub

Private Sub LimparObjetos_Avista()
   txtSubtotal.Text = "0,00"
   optDescPorc.Value = True
   'optAVdinheiro.Value = True
   'optDebito.Value = True
   'frmCartao.Visible = False
End Sub

Private Sub LimparObjetos_Prazo()
txtEntrada.Text = Format(0, ocMONEY)
cboPrazo.Text = "30"
txtValorParc.Text = Format(0, ocMONEY)
mskInicio.Mask = ""
mskInicio.Text = ""
optDescRS.Value = True
txtDesc.Text = Format(0, ocMONEY)
cboQuantParc.Text = "1"
End Sub

Private Function Autonumeracao_Parcelas() As Long
Dim sSQL As String
Dim r As ADODB.Recordset
Dim lNovoCod As Long

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultima_parcela FROM parcelas;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lNovoCod = r("ultima_parcela") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

Autonumeracao_Parcelas = lNovoCod
End Function


Private Function Atualizar_Dados_OS() As Boolean
Dim sSQL As String

'Comando de atualização
sSQL = "UPDATE os SET " & _
   "data_entrada = CONVERT(DATETIME, '" & Format$(mskDataEntrada.Text, ocDATA) & "', 103), " & _
   "hora_entrada = '" & Format$(mskHoraEntrada.Text, ocHORA) & "', " & _
   "cod_cliente = " & txtCodCliente.Text & ", " & _
   "cod_funcionario = " & txtCodFuncionario.Text & ", " & _
   "status = '" & cboStatus.Text & "', " & _
   "status_os = 0, " & _
   "SUBTOTAL = " & Replace(CCur(txtTotalPecasServicos.Text), ",", ".") & ", " & _
   "TIPO_DESC = 'P', " & _
   "OBS = '" & txtObs.Text & "', " & _
   "VALOR_DESC = " & Replace(CCur(0), ",", ".") & ", " & _
   "TOTAL = " & Replace(CCur(txtTotalPecasServicos.Text), ",", ".") & ", " & _
   "COD_RESPONSAVEL = " & IIf(txtCodMecanico.Text = "", "Null", txtCodMecanico.Text) & ", " & _
   "DATA_TERMINO = " & IIf(mskDataSaida.Text = "", "Null", "CONVERT(DATETIME, '" & Format$(mskDataSaida.Text, ocDATA) & "', 103)") & ", " & _
   "HORA_TERMINO = " & IIf(mskHoraSaida.Text = "", "Null", "'" & Format$(mskHoraSaida.Text, ocHORA) & "'") & ", " & _
   "tipo_os = '" & cboTipoOS.Text & "' "
'" & Replace(CCur(txtSubTotalServicoAuto.Text), ",", ".") & "
'Condição para atualização
sSQL = sSQL & "WHERE (cod_os = " & txtCodOS.Text & ");"

'Retorna o resultado da atualização
Atualizar_Dados_OS = dbData.Execute(sSQL)
End Function

Private Sub AutoNumeracao_Situacao_Auto()
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo FROM OS_Situacao_Auto;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then xAcess = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub AutoNumeracao_Situacao()
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo FROM OS_Situacao;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then xAcess = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub


Private Sub AutoNumeracao_Acessorio_Auto()
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo FROM OS_acessorios_Auto;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then xAcess = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub AutoNumeracao_Acessorio()
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo FROM OS_acessorios;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then xAcess = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub AutoNumeracao_OS()
sSQL = "SELECT ISNULL(MAX(COD_OS), 0) AS ultima_os FROM os;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtCodOS.Text = Format(r("ultima_os") + 1, "000000")
'If r.State <> 0 Then r.Close
'Set r = Nothing
End Sub

Private Sub AutoNumeracao_PecaItem()
sSQL = "SELECT ISNULL(MAX(item), 0) AS ultima_peca FROM pedidos_itens where cod_pedido = " & txtCodPedido.Text & " ;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then xPecaItem = r("ultima_peca") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub AutoNumeracao_Peca()
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultima_peca FROM pedidos_itens;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then xPeca = Format(r("ultima_peca") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub AutoNumeracao_Servico()
Dim vTabelaServico As String
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    vTabelaServico = "OS_Servicos_Auto"
ElseIf vTipoOS = "Recapadora" Then
    vTabelaServico = "os_servicos_recapadora"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    vTabelaServico = "OS_Servicos_Auto"
ElseIf vTipoOS = "Comunicação Visual" Then
    vTabelaServico = "OS_Servicos_Comunicacao"
End If

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo FROM " & vTabelaServico
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then xServ = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGrid_PecasServicos(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridPecasServicos
      .Clear
      .Cols = 9
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 2000
      .ColWidth(3) = 4200
      .ColWidth(4) = 900
      .ColWidth(5) = 900
      .ColWidth(6) = 1050
      .ColWidth(7) = 900
      .ColWidth(8) = 1050
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "TIPO"
      .TextMatrix(0, 3) = "SERVIÇOS"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "QUANT."
      .TextMatrix(0, 6) = "SUBTOTAL"
      .TextMatrix(0, 7) = "DESC."
      .TextMatrix(0, 8) = "TOTAL"
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("var_COD")
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("var_tipo"))
            If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
                .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("descricao"))
            ElseIf vTipoOS = "Comunicação Visual" Then
                .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("descricao"))
            ElseIf vTipoOS = "Recapadora" Then
                .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("descricao")) & " | " & rTabela("vartipo") & " | " & rTabela("varMEDIDA") & " | " & rTabela("varARO") & " "
            End If
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.Rows - 1, 5) = rTabela("quantidade")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("var_SUBtotal"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("var_DESCONTO"), ocMONEY)
            .TextMatrix(.Rows - 1, 8) = Format(rTabela("var_total"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
      
      lblPecasServicos.Caption = Format(SomaGrid(GridPecasServicos, 8), ocMONEY)
      llblTotalSemDesconto.Caption = Format(SomaGrid(GridPecasServicos, 6), ocMONEY)
      lblSomaDesconto.Caption = Format(SomaGrid(GridPecasServicos, 7), ocMONEY)
   End With
End Sub

Private Sub LimparGrid_Situacao()
Dim i As Integer

With GridPecasServicos
    .Clear
    .Rows = 1       'INICIA O Grid_OS COM UMA LINHA
    .FixedCols = 0  'DETERMINA QUE NÃO HAJA COLUNA FIXA
    
    'Abaixo o cabeçalho é criado
    .FormatString = "^CÓD.|^TECNICO|^FINANCEIRO|^CLIENTE|^VEICULO|^ENTRADA|^PEDIDO"
    .ColWidth(0) = 650
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 4000
    .ColWidth(4) = 3000
    .ColWidth(5) = 1350
    .ColWidth(6) = 1000
     
     'colocar os cabeçalho em negrito
    For i = 0 To .Cols - 1
       .Col = i
       .Row = 0
       .CellFontBold = True
    Next

 .Redraw = False
 
End With
End Sub

Private Sub LimparGrid_Servicos()
   Dim i As Integer
   
   With Grid_Servicos
      .Clear
      .Cols = 6
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 4400
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "SERVIÇOS"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "QUANT."
      .TextMatrix(0, 5) = "TOTAL"
      
      .Redraw = False
      .Rows = .Rows + 1
      i = i + 1
      
      .Rows = .Rows - 1
      .Redraw = True
      
      'lblTotal.Caption = Format(0, ocMONEY)
   End With
End Sub


Private Sub LimparObjetos_Pecas()
txtCodPeca.Text = ""
cboPecas.Text = ""
txtQuantPeca.Text = ""
txtValorPeca.Text = ""
txtTotalPeca.Text = ""
txtCodBarra.Text = ""
txtSubtotalPecas.Text = ""
txtDescPecas.Text = ""
End Sub

Private Sub LimparObjetos_ServicosAuto()
txtCodServicoAuto.Text = ""
cboServicosAuto.Text = ""
txtQuantServicoAuto.Text = ""
mskValorServicoAuto.Text = ""
txtSubTotalServicoAuto.Text = ""
txtDescServicoAuto.Text = ""
txtTotalServicoAuto.Text = ""
txtDote.Text = ""
cboTipo.Text = ""
txtSerie.Text = ""
txtFogo.Text = ""
cboMarca.Text = ""
txtObsServ.Text = ""
End Sub



Private Sub LimparObjetos_Servicos()
txtCodServicoAuto.Text = ""
cboServicosAuto.Text = ""
txtQuantServicoAuto.Text = ""
mskValorServicoAuto.Text = ""
cboMarca.Text = ""
txtDote.Text = ""

End Sub

Private Sub MostrarEquipamento()
If txtCodOS.Text = "" Then Exit Sub

sSQL = "SELECT * FROM OS_Equipamento WHERE (cod_os = " & txtCodOS.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   cboTanque.Text = ValidateNull(r("equipamento"))
   cboFabricante.Text = ValidateNull(r("fabricante"))
   cboModelo.Text = ValidateNull(r("MODELO"))
   txtPareceCliente.Text = ValidateNull(r("PARECER_CLIENTE"))
End If
End Sub

Private Sub MostrarEquipamentoAuto()
If txtCodOS.Text = "" Then Exit Sub

sSQL = "SELECT * FROM OS_Equipamento_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   'cboFabricante.Text = ValidateNull(rTabela("fabricante"))
   cboFabricante.Text = ValidateNull(r("fabricante"))
   cboModelo.Text = ValidateNull(r("MODELO"))
   txtPlaca.Text = ValidateNull(r("PLACA"))
   txtAno.Text = ValidateNull(r("ANO"))
   txtKM.Text = ValidateNull(r("KM"))
   txtChassi.Text = ValidateNull(r("CHASSI"))
   cboCor.Text = ValidateNull(r("COR"))
   cboTanque.Text = ValidateNull(r("TANQUE"))
   txtPareceCliente.Text = ValidateNull(r("PARECER_CLIENTE"))
End If
End Sub


Private Sub MostrarGrid_Servicos()
If txtCodOS.Text = "" Then txtCodOS.Text = 0
If txtCodPedido.Text = "" Then txtCodPedido.Text = 0

'If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
'    sSQL = "SELECT 'SERVIÇO' as var_Tipo, * FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
'ElseIf vTipoOS = "Recapadora" Then
'    sSQL = "SELECT 'SERVIÇO' as var_Tipo, *  FROM os_servicos_recapadora WHERE (cod_os = " & txtCodOS.Text & ");"
'ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
'    sSQL = "SELECT 'SERVIÇO' as var_Tipo, * FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
'End If

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    sSQL = "SELECT 'SERVIÇO' as var_Tipo, COD_OS as var_COD, DESCRICAO, PRECO, QUANTIDADE, SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, TOTAL as var_TOTAL, CODIGO AS var_CODITEM, '' as var_CODPROD FROM OS_Servicos_Auto WHERE (COD_OS = " & txtCodOS.Text & ")" & _
          " UNION ALL "
    sSQL = sSQL & "SELECT 'PRODUTO' AS var_Tipo, pedidos_itens.COD_PEDIDO as var_COD, produtos.descricao as var_desc, preco, quantidade, pedidos_itens.SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, pedidos_itens.TOTAL AS var_TOTAL, pedidos_itens.CODIGO AS var_CODITEM, pedidos_itens.COD_PRODUTO as var_CODPROD " & _
             "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
             "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
             "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") "
ElseIf vTipoOS = "Comunicação Visual" Then
    sSQL = "SELECT 'SERVIÇO' as var_Tipo, COD_OS as var_COD, DESCRICAO, PRECO, QUANTIDADE, SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, TOTAL as var_TOTAL, CODIGO AS var_CODITEM, '' as var_CODPROD, OBS as var_obs FROM OS_Servicos_Comunicacao WHERE (COD_OS = " & txtCodOS.Text & ")" & _
          " UNION ALL "
    sSQL = sSQL & "SELECT 'PRODUTO' AS var_Tipo, pedidos_itens.COD_PEDIDO as var_COD, produtos.descricao as var_desc, preco, quantidade, pedidos_itens.SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, pedidos_itens.TOTAL AS var_TOTAL, pedidos_itens.CODIGO AS var_CODITEM, pedidos_itens.COD_PRODUTO as var_CODPROD, '' as var_obs " & _
             "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
             "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
             "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") "
             Debug.Print sSQL
ElseIf vTipoOS = "Recapadora" Then
    sSQL = "SELECT 'SERVIÇO' as var_Tipo, COD_OS as var_COD, DESCRICAO, PRECO, QUANTIDADE, SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, TOTAL as var_TOTAL, TIPO as var_TipoPneu, SERIE as var_serie, FOGO as var_fogo, ARO as var_aro, BANDA as var_banda, DOTE as var_dote, MEDIDA as var_medida, FABRICANTE as var_fabricante, CODIGO AS var_CODITEM, '' as var_CODPROD FROM os_servicos_recapadora WHERE (COD_OS = " & txtCodOS.Text & ")" & _
          " UNION ALL "
    sSQL = sSQL & "SELECT 'PRODUTO' AS var_Tipo, pedidos_itens.COD_PEDIDO as var_COD, produtos.descricao as var_desc, preco, quantidade, pedidos_itens.SUBTOTAL as var_SUBTOTAL, pedidos_itens.DESCONTO as var_DESCONTO, pedidos_itens.TOTAL AS var_TOTAL, '' as var_TipoPneu, '' as var_serie, '' as var_fogo, '' as var_aro, '' as var_banda, '' as var_dote, '' as var_medida, '' as var_fabricante, pedidos_itens.CODIGO AS var_CODITEM , pedidos_itens.COD_PRODUTO as var_CODPROD " & _
             "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
             "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
             "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") "
End If

Set r = dbData.OpenRecordset(sSQL)
'Debug.Print sSQL

'If r.RecordCount > 0 Then
'    cmdImpOrcamento2.Enabled = True
'Else
'    cmdImpOrcamento2.Enabled = False
'End If

FormatarGrid_Servicos r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGrid_Servicos(rTabela As ADODB.Recordset)
Dim i As Integer
Dim j As Integer
Dim soma As Currency
Dim QUANT As Integer

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    With Grid_Servicos
       .Clear
       .Cols = 11
       .Rows = 2
       
       .ColWidth(0) = 0
       .ColWidth(1) = 0
       .ColWidth(2) = 1000
       .ColWidth(3) = 3000
       .ColWidth(4) = 1000
       .ColWidth(5) = 600
       .ColWidth(6) = 1100
       .ColWidth(7) = 800
       .ColWidth(8) = 1100
       .ColWidth(9) = 0
       .ColWidth(10) = 0
       
       For i = 0 To .Cols - 1
          .Col = i
          .Row = 0
          .CellFontBold = True
       Next
       
       .TextMatrix(0, 1) = "COD"
       .TextMatrix(0, 2) = "TIPO"
       .TextMatrix(0, 3) = "SERVIÇOS"
       .TextMatrix(0, 4) = "VALOR"
       .TextMatrix(0, 5) = "QTDE"
       .TextMatrix(0, 6) = "SUBTOTAL"
       .TextMatrix(0, 7) = "DESC."
       .TextMatrix(0, 8) = "TOTAL"
       .TextMatrix(0, 9) = "ITEM"
       .TextMatrix(0, 10) = "PROD"
    
       .Redraw = False
       
       If Not rTabela Is Nothing Then
          Do While Not rTabela.EOF
             .TextMatrix(.Rows - 1, 1) = rTabela("var_COD")
             .TextMatrix(.Rows - 1, 2) = rTabela("VAR_TIPO")
             .TextMatrix(.Rows - 1, 3) = rTabela("descricao")
             .TextMatrix(.Rows - 1, 4) = Format(rTabela("preco"), ocMONEY)
             .TextMatrix(.Rows - 1, 5) = rTabela("quantidade")
             .TextMatrix(.Rows - 1, 6) = Format(rTabela("var_SUBTOTAL"), ocMONEY)
             .TextMatrix(.Rows - 1, 7) = Format(rTabela("var_desconto"), ocMONEY)
             .TextMatrix(.Rows - 1, 8) = Format(rTabela("var_total"), ocMONEY)
             .TextMatrix(.Rows - 1, 9) = rTabela("var_CODITEM")
             .TextMatrix(.Rows - 1, 10) = rTabela("var_CODPROD")
             
             rTabela.MoveNext
             .Rows = .Rows + 1
             i = i + 1
          Loop
       End If
    
          'MUDAR COR DE FONTE DA COLUNA
          For i = 1 To .Rows - 1
             .Row = i
             .Col = 6
             .CellForeColor = &HC0&
             .CellFontBold = True
          Next

      'Deixar linha de outra cor
      For i = 1 To .Rows - 1
         For j = 0 To .Cols - 1
            .Col = j
            .Row = i
            
            If .TextMatrix(i, 2) = "PRODUTO" Then
               .CellForeColor = vbRed
               '.CellFontBold = True
            End If
         Next
      Next

       .Rows = .Rows - 1
       .Redraw = True
       
   
    'somando os serviços
    soma = 0
    QUANT = 0
    With Grid_Servicos
       For i = 1 To .Rows - 1
          If .TextMatrix(i, 2) = "SERVIÇO" Then
             soma = soma + CCur(.TextMatrix(i, 8))
             QUANT = QUANT + 1
          End If
       Next
    End With

    txtTotalServicos.Text = Format(soma, "#,##0.00")
    txtQuantServicos.Text = Format(QUANT, "000")
    
    'somando as peças
    soma = 0
    QUANT = 0
    With Grid_Servicos
       For i = 1 To .Rows - 1
          If .TextMatrix(i, 2) = "PRODUTO" Then
             soma = soma + CCur(.TextMatrix(i, 8))
             QUANT = QUANT + 1
          End If
       Next
    End With

    txtTotalPecas.Text = Format(soma, "#,##0.00")
    txtQuantPecas.Text = Format(QUANT, "000")

    'somar totais
    txtSubtotalGeral.Text = Format(SomaGrid(Grid_Servicos, 6), ocMONEY)
    txtDescGeral.Text = Format(SomaGrid(Grid_Servicos, 7), ocMONEY)
    txtTotalGeral.Text = Format(SomaGrid(Grid_Servicos, 8), ocMONEY)
    End With
ElseIf vTipoOS = "Comunicação Visual" Then
    With Grid_Servicos
       .Clear
       .Cols = 11
       .Rows = 2
       
       .ColWidth(0) = 0
       .ColWidth(1) = 0
       .ColWidth(2) = 1000
       .ColWidth(3) = 4000
       .ColWidth(4) = 1000
       .ColWidth(5) = 600
       .ColWidth(6) = 1100
       .ColWidth(7) = 800
       .ColWidth(8) = 1100
       .ColWidth(9) = 0
       .ColWidth(10) = 1000
       
       For i = 0 To .Cols - 1
          .Col = i
          .Row = 0
          .CellFontBold = True
       Next
       
       .TextMatrix(0, 1) = "COD"
       .TextMatrix(0, 2) = "TIPO"
       .TextMatrix(0, 3) = "SERVIÇOS"
       .TextMatrix(0, 4) = "VALOR"
       .TextMatrix(0, 5) = "QTDE"
       .TextMatrix(0, 6) = "SUBTOTAL"
       .TextMatrix(0, 7) = "DESC."
       .TextMatrix(0, 8) = "TOTAL"
       .TextMatrix(0, 9) = "ITEM"
       .TextMatrix(0, 10) = "PROD"
    
       .Redraw = False
       
       '.RowHeight(-1) = .RowHeight(1) * 2
       
       If Not rTabela Is Nothing Then
          Do While Not rTabela.EOF
             .TextMatrix(.Rows - 1, 1) = rTabela("var_COD")
             .TextMatrix(.Rows - 1, 2) = rTabela("VAR_TIPO")
             .TextMatrix(.Rows - 1, 3) = rTabela("descricao") & vbCrLf & rTabela("var_obs")
             .TextMatrix(.Rows - 1, 4) = Format(rTabela("preco"), ocMONEY)
             .TextMatrix(.Rows - 1, 5) = rTabela("quantidade")
             .TextMatrix(.Rows - 1, 6) = Format(rTabela("var_SUBTOTAL"), ocMONEY)
             .TextMatrix(.Rows - 1, 7) = Format(rTabela("var_desconto"), ocMONEY)
             .TextMatrix(.Rows - 1, 8) = Format(rTabela("var_total"), ocMONEY)
             .TextMatrix(.Rows - 1, 9) = rTabela("var_CODITEM")
             .TextMatrix(.Rows - 1, 10) = rTabela("var_CODPROD")
             
             rTabela.MoveNext
             .Rows = .Rows + 1
             i = i + 1
          Loop
       End If
    
          'MUDAR COR DE FONTE DA COLUNA
          For i = 1 To .Rows - 1
             .Row = i
             .Col = 6
             .CellForeColor = &HC0&
             .CellFontBold = True
          Next

      'Deixar linha de outra cor
      For i = 1 To .Rows - 1
         For j = 0 To .Cols - 1
            .Col = j
            .Row = i
            
            If .TextMatrix(i, 2) = "PRODUTO" Then
               .CellForeColor = vbRed
               '.CellFontBold = True
            End If
         Next
      Next

       .Rows = .Rows - 1
       .Redraw = True
       
   
    'somando os serviços
    soma = 0
    QUANT = 0
    With Grid_Servicos
       For i = 1 To .Rows - 1
          If .TextMatrix(i, 2) = "SERVIÇO" Then
             soma = soma + CCur(.TextMatrix(i, 8))
             QUANT = QUANT + 1
          End If
       Next
    End With

    txtTotalServicos.Text = Format(soma, "#,##0.00")
    txtQuantServicos.Text = Format(QUANT, "000")
    
    'somando as peças
    soma = 0
    QUANT = 0
    With Grid_Servicos
       For i = 1 To .Rows - 1
          If .TextMatrix(i, 2) = "PRODUTO" Then
             soma = soma + CCur(.TextMatrix(i, 8))
             QUANT = QUANT + 1
          End If
       Next
    End With

    txtTotalPecas.Text = Format(soma, "#,##0.00")
    txtQuantPecas.Text = Format(QUANT, "000")

    'somar totais
    txtSubtotalGeral.Text = Format(SomaGrid(Grid_Servicos, 6), ocMONEY)
    txtDescGeral.Text = Format(SomaGrid(Grid_Servicos, 7), ocMONEY)
    txtTotalGeral.Text = Format(SomaGrid(Grid_Servicos, 8), ocMONEY)
    End With
ElseIf vTipoOS = "Recapadora" Then
    With Grid_Servicos
       .Clear
       .Cols = 11
       .Rows = 2
       
       .ColWidth(0) = 0
       .ColWidth(1) = 0
       .ColWidth(2) = 800
       .ColWidth(3) = 5500
       .ColWidth(4) = 750
       .ColWidth(5) = 500
       .ColWidth(6) = 850
       .ColWidth(7) = 550
       .ColWidth(8) = 800
       .ColWidth(9) = 0
       .ColWidth(10) = 0
       
       For i = 0 To .Cols - 1
          .Col = i
          .Row = 0
          .CellFontBold = True
       Next
       
       .TextMatrix(0, 1) = "COD"
       .TextMatrix(0, 2) = "TIPO"
       .TextMatrix(0, 3) = "DETALHES"
       .TextMatrix(0, 4) = "VALOR"
       .TextMatrix(0, 5) = "QTDE"
       .TextMatrix(0, 6) = "SUBTOTAL"
       .TextMatrix(0, 7) = "DESC."
       .TextMatrix(0, 8) = "TOTAL"
       .TextMatrix(0, 9) = "ITEM"
       .TextMatrix(0, 10) = "PROD"
    
       .Redraw = False
       
       If Not rTabela Is Nothing Then
          Do While Not rTabela.EOF
             .TextMatrix(.Rows - 1, 1) = rTabela("var_COD")
             .TextMatrix(.Rows - 1, 2) = rTabela("VAR_TIPO")
            If .TextMatrix(.Rows - 1, 2) = "SERVIÇO" Then
             .TextMatrix(.Rows - 1, 3) = rTabela("descricao") & " | " & ValidateNull(rTabela("var_TipoPneu")) & " | " & ValidateNull(rTabela("var_serie")) & " | " & ValidateNull(rTabela("var_fogo")) & " | " & ValidateNull(rTabela("var_aro")) & " | " & ValidateNull(rTabela("var_banda")) & " | " & ValidateNull(rTabela("var_dote")) & " | " & ValidateNull(rTabela("var_medida")) & " | " & ValidateNull(rTabela("var_fabricante"))
            Else
                .TextMatrix(.Rows - 1, 3) = rTabela("descricao")
            End If
            
             .TextMatrix(.Rows - 1, 4) = Format(rTabela("preco"), ocMONEY)
             .TextMatrix(.Rows - 1, 5) = rTabela("quantidade")
             .TextMatrix(.Rows - 1, 6) = Format(rTabela("var_subtotal"), ocMONEY)
             .TextMatrix(.Rows - 1, 7) = Format(rTabela("var_desconto"), ocMONEY)
             .TextMatrix(.Rows - 1, 8) = Format(rTabela("var_total"), ocMONEY)
             .TextMatrix(.Rows - 1, 9) = rTabela("var_CODITEM")
             .TextMatrix(.Rows - 1, 10) = rTabela("var_CODPROD")
             rTabela.MoveNext
             .Rows = .Rows + 1
             i = i + 1
          Loop
       End If
    
          'MUDAR COR DE FONTE DA COLUNA
          For i = 1 To .Rows - 1
             .Row = i
             .Col = 8
             .CellForeColor = &HC0&
             .CellFontBold = True
          Next
       
       .Rows = .Rows - 1
       .Redraw = True

    'somando os serviços

    soma = 0
    QUANT = 0
    With Grid_Servicos
       For i = 1 To .Rows - 1
          If .TextMatrix(i, 2) = "SERVIÇO" Then
             soma = soma + CCur(.TextMatrix(i, 8))
             QUANT = QUANT + 1
          End If
       Next
    End With

    txtTotalServicos.Text = Format(soma, "#,##0.00")
    txtQuantServicos.Text = Format(QUANT, "000")
    
    'somando as peças
    soma = 0
    QUANT = 0
    With Grid_Servicos
       For i = 1 To .Rows - 1
          If .TextMatrix(i, 2) = "PRODUTO" Then
             soma = soma + CCur(.TextMatrix(i, 8))
             QUANT = QUANT + 1
          End If
       Next
    End With

    txtTotalPecas.Text = Format(soma, "#,##0.00")
    txtQuantPecas.Text = Format(QUANT, "000")

    'somar totais
    txtSubtotalGeral.Text = Format(SomaGrid(Grid_Servicos, 6), ocMONEY)
    txtDescGeral.Text = Format(SomaGrid(Grid_Servicos, 7), ocMONEY)
    txtTotalGeral.Text = Format(SomaGrid(Grid_Servicos, 8), ocMONEY)

    End With
End If
End Sub


Private Sub LimparObjetos_Entrada()
txtCodCliente.Text = ""
txtCodFuncionario.Text = ""
mskDataEntrada.Mask = ""
mskDataEntrada.Text = ""
mskHoraEntrada.Mask = ""
mskHoraEntrada.Text = ""
cboCliente.Text = ""
cboFuncionario.Text = ""
cboStatus.Text = ""
txtCodMecanico.Text = ""
cboMecanico.Text = ""
cboTipoOS.Text = ""
txtObs.Text = ""
mskDataSaida.Mask = ""
mskDataSaida.Text = ""
mskHoraSaida.Mask = ""
mskHoraSaida.Text = ""
txtTotalPecasServicos.Text = Format(0, "##,##0.00")
LimparObjetos_Equipamento
LimparGrid_Servicos
End Sub
Private Sub Mostrar_Entrada(rTabela As ADODB.Recordset)
'Se o parametro passado é Nothing, sai da rotina
If rTabela Is Nothing Then Exit Sub

If Not rTabela.BOF Then
   mskDataEntrada.Text = Format(rTabela("data_entrada"), "dd/mm/yy")
   mskHoraEntrada.Text = Format(rTabela("hora_entrada"), ocHRMN)
   txtCodCliente.Text = ValidateNull(rTabela("cod_cliente"))
   txtCodFuncionario.Text = ValidateNull(rTabela("cod_funcionario"))
   cboStatus.Text = ValidateNull(rTabela("status"))
   txtCodMecanico.Text = ValidateNull(rTabela("COD_RESPONSAVEL"))
   mskDataSaida.Text = Format(rTabela("DATA_TERMINO"), "dd/mm/yy")
   mskHoraSaida.Text = Format(rTabela("DATA_TERMINO"), ocHRMN)
   cboTipoOS.Text = ValidateNull(rTabela("tipo_os"))
   txtObs.Text = ValidateNull(rTabela("OBS"))
   txtCodPedido.Text = ValidateNull(rTabela("cod_Pedido"))
End If

'sSQL = "SELECT cod_cliente, cod_pedido FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
'Set r = dbData.OpenRecordset(sSQL)
'If Not r.BOF Then txtCodCliente.Text = rTabela("cod_cliente")
'If r.State <> 0 Then r.Close
'Set r = Nothing

If txtCodCliente.Text = "" Then Exit Sub

sSQL = "SELECT codigo, nome, celular FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then cboCliente.Text = r("nome") & IIf(Trim(ValidateNull(r("celular"))) = "", "", "     (" & Right$(ValidateNull(r("celular")), 9) & ")")
If r.State <> 0 Then r.Close
Set r = Nothing

End Sub

Private Sub MostrarGrid_PecasServicos()
Dim totalRegistros As Long

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    sSQL = "SELECT 'SERVIÇO' as var_Tipo, COD_OS as var_COD, DESCRICAO, PRECO, QUANTIDADE, SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, TOTAL as var_TOTAL FROM OS_Servicos_Auto WHERE (COD_OS = " & Grid_OS.TextMatrix(Grid_OS.Row, 0) & ")" & _
          " UNION ALL "
    sSQL = sSQL & "SELECT 'PRODUTO' AS var_Tipo, pedidos_itens.COD_PEDIDO as var_COD, produtos.descricao as var_desc, preco, quantidade, pedidos_itens.SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, pedidos_itens.TOTAL AS var_TOTAL " & _
             "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
             "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
             "WHERE (pedidos_itens.cod_pedido = " & Grid_OS.TextMatrix(Grid_OS.Row, 7) & ") "
             'Debug.Print sSQL
ElseIf vTipoOS = "Comunicação Visual" Then
    sSQL = "SELECT 'SERVIÇO' as var_Tipo, COD_OS as var_COD, DESCRICAO, PRECO, QUANTIDADE, SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, TOTAL as var_TOTAL FROM OS_Servicos_Comunicacao WHERE (COD_OS = " & Grid_OS.TextMatrix(Grid_OS.Row, 0) & ")" & _
          " UNION ALL "
    sSQL = sSQL & "SELECT 'PRODUTO' AS var_Tipo, pedidos_itens.COD_PEDIDO as var_COD, produtos.descricao as var_desc, preco, quantidade, pedidos_itens.SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, pedidos_itens.TOTAL AS var_TOTAL " & _
             "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
             "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
             "WHERE (pedidos_itens.cod_pedido = " & Grid_OS.TextMatrix(Grid_OS.Row, 5) & ") "
ElseIf vTipoOS = "Recapadora" Then
    sSQL = "SELECT 'SERVIÇO' as var_Tipo, COD_OS as var_COD, DESCRICAO, PRECO, QUANTIDADE, SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, TOTAL as var_TOTAL, DOTE as varTipo, MEDIDA as varMedida, FABRICANTE as varAro, BANDA as varBanda FROM os_servicos_recapadora WHERE (COD_OS = " & Grid_OS.TextMatrix(Grid_OS.Row, 0) & ")" & _
          " UNION ALL "
    sSQL = sSQL & "SELECT 'PRODUTO' AS var_Tipo, pedidos_itens.COD_PEDIDO as var_COD, produtos.descricao as var_desc, preco, quantidade, pedidos_itens.SUBTOTAL as var_SUBTOTAL, pedidos_itens.DESCONTO as var_DESCONTO, pedidos_itens.TOTAL AS var_TOTAL, '' as varTipo, '' as varMedida, '' as varAro, '' as varBanda " & _
             "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
             "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
             "WHERE (pedidos_itens.cod_pedido = " & Grid_OS.TextMatrix(Grid_OS.Row, 6) & ") "
End If
      
Set r = dbData.OpenRecordset(sSQL, totalRegistros)
FormatarGrid_PecasServicos r

'cmdEditarOS.Enabled = True

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub MostrarGrid_OS_Situacao()
Dim totalRegistros As Long
Dim SITUACAO As String
Dim var_STATUS As String
Dim INDICE As String
Dim varTIPO_OS As String
Dim vStatusFinanceiro As String

INDICE = "os.COD_OS DESC "

If optFinanceiroFechado.Value = True Then
    vStatusFinanceiro = "and os.status_os = 1"
ElseIf optFinanceiroAberto.Value = True Then
    vStatusFinanceiro = "and os.status_os = 0"
End If

SITUACAO = ""
var_STATUS = ""

If optTodos.Value = True Then
    varTIPO_OS = " (os.tipo_os <> 'TODOS') "
ElseIf optServico.Value = True Then
    varTIPO_OS = " (os.tipo_os = 'CONSERTO') "
ElseIf optOrcamento.Value = True Then
    varTIPO_OS = " (os.tipo_os = 'ORÇAMENTO') "
ElseIf optGarantia.Value = True Then
    varTIPO_OS = " (os.tipo_os = 'GARANTIA') "
End If

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
   sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS.COD_PEDIDO, OS.TIPO_OS, OS_Equipamento_Auto.FABRICANTE, OS_Equipamento_Auto.ANO, OS_Equipamento_Auto.MODELO, OS_Equipamento_Auto.PLACA, os.DATA_ENTRADA, os.HORA_ENTRADA, os.STATUS AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_Financeiro,  os.STATUS_OS, os.STATUS, os.SUBTOTAL, os.TOTAL, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.ValorDescReal " & _
      "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _
      "WHERE " & varTIPO_OS & " " & SITUACAO & " " & var_STATUS & vStatusFinanceiro & _
      "ORDER BY " & INDICE
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
   sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS.COD_PEDIDO, OS.TIPO_OS, OS_Equipamento.FABRICANTE, OS_Equipamento.EQUIPAMENTO, OS_Equipamento.MODELO, os.DATA_ENTRADA, os.HORA_ENTRADA, os.STATUS AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_Financeiro, os.STATUS_OS, os.STATUS, os.SUBTOTAL, os.TOTAL, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.ValorDescReal " & _
      "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS " & _
      "WHERE " & varTIPO_OS & " " & SITUACAO & " " & var_STATUS & vStatusFinanceiro & _
      "ORDER BY " & INDICE
ElseIf vTipoOS = "Comunicação Visual" Then
   'sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS.COD_PEDIDO, OS_Equipamento.FABRICANTE, OS_Equipamento.EQUIPAMENTO, OS_Equipamento.MODELO, os.DATA_ENTRADA, os.HORA_ENTRADA, os.STATUS AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_Financeiro, os.STATUS_OS, os.STATUS, os.SUBTOTAL, os.TOTAL, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.ValorDescReal " & _
      "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " " & SITUACAO & " " & var_STATUS & vStatusFinanceiro & _
      "ORDER BY " & INDICE
    'sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS.COD_PEDIDO, OS.DATA_ENTRADA, OS.HORA_ENTRADA, OS.STATUS AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_Financeiro, OS.STATUS_OS, OS.STATUS, OS.SUBTOTAL, OS.TOTAL, OS.TIPO_PAGAMENTO, OS.PAGAMENTO, OS.ValorDescReal " & _
        "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE WHERE (OS.TIPO_OS <> 'TODOS') AND (OS.STATUS_OS = 0) " & _
        "ORDER BY " & INDICE
        
   sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS.COD_PEDIDO, OS.TIPO_OS, os.DATA_ENTRADA, os.HORA_ENTRADA, os.STATUS AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_Financeiro, os.STATUS_OS, os.STATUS, os.SUBTOTAL, os.TOTAL, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.ValorDescReal " & _
      "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE " & _
      "WHERE " & varTIPO_OS & " " & SITUACAO & " " & var_STATUS & vStatusFinanceiro & _
      "ORDER BY " & INDICE

'SELECT DISTINCT OS.COD_OS, cliente.Nome, OS.COD_PEDIDO, OS.DATA_ENTRADA, OS.HORA_ENTRADA, OS.STATUS AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_Financeiro, OS.STATUS_OS, OS.STATUS, OS.SUBTOTAL, OS.TOTAL, OS.TIPO_PAGAMENTO, OS.PAGAMENTO, OS.ValorDescReal
'FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE
      'Debug.Print sSQL
End If

Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid_OS_Situacao r

lblQuantOS.Caption = totalRegistros

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub MostrarGrid_OS()
Dim totalRegistros As Long

Dim SITUACAO As String
Dim var_STATUS As String
Dim INDICE As String
Dim varTIPO_OS As String

'indice
If cboIndice.Text = "CÓD. OS" Then
   INDICE = "os.COD_OS DESC "
ElseIf cboIndice.Text = "TIPO DE SERVIÇO" Then
   INDICE = "os.TIPO_OS DESC "
ElseIf cboIndice.Text = "CLIENTE" Then
   INDICE = "cliente.nome DESC "
ElseIf cboIndice.Text = "DATA" Then
   INDICE = "os.DATA_ENTRADA DESC "
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

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
    If cboConsultaCriterios.Text = "CLIENTE" Then
       If txtCodClienteLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, os.status AS var_status, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS WHERE " & varTIPO_OS & " and (cod_cliente = " & txtCodClienteLocalizar.Text & ") " & _
          "ORDER BY " & INDICE
          Debug.Print sSQL
       
    ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then
       If cboLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, os.status AS var_status, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS WHERE " & varTIPO_OS & " and (os.cod_os = " & cboLocalizar.Text & ") " & _
          "ORDER BY " & INDICE
    Else
        sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento_Auto.FABRICANTE, OS_Equipamento_Auto.ANO, OS_Equipamento_Auto.MODELO, os.STATUS AS var_status, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.STATUS_OS, os.STATUS, os.SUBTOTAL, os.TOTAL, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.ValorDescReal " & _
                "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _
                "WHERE " & varTIPO_OS & " " & SITUACAO & var_STATUS & _
                "ORDER BY " & INDICE
    End If
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    If cboConsultaCriterios.Text = "CLIENTE" Then
       If txtCodClienteLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento.fabricante, OS_Equipamento.modelo, OS_Equipamento.equipamento, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.* " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (cod_cliente = " & txtCodClienteLocalizar.Text & ") " & _
          "ORDER BY " & INDICE
       
    ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then
       If cboLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.* " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (os.cod_os = " & cboLocalizar.Text & ") " & _
          "ORDER BY " & INDICE
    Else
        sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento.FABRICANTE, OS_Equipamento.equipamento, OS_Equipamento.MODELO, os.STATUS AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.STATUS_OS, os.STATUS, os.SUBTOTAL, os.TOTAL, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.ValorDescReal " & _
                "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS " & _
                "WHERE " & varTIPO_OS & " " & SITUACAO & var_STATUS & _
                "ORDER BY " & INDICE
    End If
ElseIf vTipoOS = "Comunicação Visual" Then
    If cboConsultaCriterios.Text = "CLIENTE" Then
       If txtCodClienteLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento.fabricante, OS_Equipamento.modelo, OS_Equipamento.equipamento, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.* " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (cod_cliente = " & txtCodClienteLocalizar.Text & ") " & _
          "ORDER BY " & INDICE
       
    ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then
       If cboLocalizar.Text = "" Then Exit Sub
       sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.* " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS WHERE " & varTIPO_OS & " and (os.cod_os = " & cboLocalizar.Text & ") " & _
          "ORDER BY " & INDICE
    Else
        sSQL = "SELECT DISTINCT OS.COD_OS, cliente.Nome, OS_Equipamento.FABRICANTE, OS_Equipamento.equipamento, OS_Equipamento.MODELO, os.STATUS AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.STATUS_OS, os.STATUS, os.SUBTOTAL, os.TOTAL, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.ValorDescReal " & _
                "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento ON OS.COD_OS = OS_Equipamento.COD_OS " & _
                "WHERE " & varTIPO_OS & " " & SITUACAO & var_STATUS & _
                "ORDER BY " & INDICE
    End If
End If
'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid_OS r

printSQL = sSQL

lblQuant.Caption = "QUANTIDADE: " & Format(totalRegistros, "000")

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub MostrarTipoOS()
cboTipoOS.Clear
cboTipoOS.AddItem "CONSERTO"
cboTipoOS.AddItem "GARANTIA"
cboTipoOS.AddItem "ORÇAMENTO"
cboTipoOS.AddItem "CONFECÇÃO"
End Sub

Private Sub Preencher_Criterios()
cboConsultaCriterios.Clear
cboConsultaCriterios.AddItem "TODOS"
cboConsultaCriterios.AddItem "CÓD. OS"
cboConsultaCriterios.AddItem "CLIENTE"
End Sub

Private Sub Preencher_Indice()
   cboIndice.Clear
   cboIndice.AddItem "CÓD. OS"
   cboIndice.AddItem "TIPO DE SERVIÇO"
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

Private Sub Somar_Totais()
Dim Servicos As Currency
Dim Pecas As Currency
Dim Total As Currency

If txtTotalServicos.Text <> "" Then Servicos = txtTotalServicos.Text Else Servicos = 0
If txtTotalPecas.Text <> "" Then Pecas = txtTotalPecas.Text Else Pecas = 0
Total = Servicos + Pecas

'txtSubTotal.Text = Format(Total, ocMONEY)
txtTotalPecasServicos.Text = Format(Total, ocMONEY)

If txtQuantServicos.Text <> "" Then Servicos = txtQuantServicos.Text Else Servicos = 0
If txtQuantPecas.Text <> "" Then Pecas = txtQuantPecas.Text Else Pecas = 0
Total = Servicos + Pecas

txtQuantGeral.Text = Format(Total, "000")
End Sub




Private Sub cboAcessorios_GotFocus()
cboAcessorios.Clear

sSQL = "SELECT DISTINCT acessorio, codigo FROM OS_acessorios ORDER BY acessorio;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboAcessorios.AddItem r("acessorio")
   cboAcessorios.ItemData(cboAcessorios.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboAcessorios
End Sub


Private Sub cboAcessorios_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboAcessorios_LostFocus()
On Error GoTo TrataErro

If cboAcessorios.Text = "" Then txtCodAcessorio.Text = "": Exit Sub
If cboAcessorios.ListIndex = -1 Then txtCodAcessorio.Text = "": Exit Sub
txtCodAcessorio = cboAcessorios.ItemData(cboAcessorios.ListIndex)
Exit Sub

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboCliente_GotFocus()
Dim varNomeAntes As String
Dim varCodAntes As String

varNomeAntes = cboCliente.Text
varCodAntes = txtCodCliente.Text

cboCliente.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboCliente.AddItem r("nome")
   cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboCliente.Text = varNomeAntes
txtCodCliente.Text = varCodAntes

cboCliente.SelStart = 0
cboCliente.SelLength = Len(cboCliente)

moCombo.AttachTo cboCliente
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CboCliente_LostFocus()
On Error GoTo TrataErro

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    cboFabricante.SetFocus
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    cboTanque.SetFocus
End If

If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub

If cmdAlterar.Enabled = False Then
   If cboCliente.ListIndex = -1 Then
      'txtCodCliente.Text = ""
      'Exit Sub
   End If
End If

txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)



Exit Sub

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboConsultaCriterios_Click()
If cboConsultaCriterios.Text = "TODOS" Then
   cboLocalizar.Text = ""
   cboLocalizar.Visible = False
   MostrarGrid_OS
Else
   cboLocalizar.Visible = True
   cboLocalizar.SetFocus
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
If cboConsultaCriterios.Text = "TODOS" Then
   cboLocalizar.Text = ""
   cboLocalizar.Visible = False
Else
   cboLocalizar.Visible = True
End If
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


Private Sub cboCor_GotFocus()
'Limpa a lista atual
cboCor.Clear

sSQL = "SELECT DISTINCT cor FROM OS_Equipamento_Auto ORDER BY cor;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboCor.AddItem ValidateNull(r("cor"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboCor
End Sub


Private Sub cboCor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboCor_LostFocus()
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    cboTanque.SetFocus
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    'cboTanque.SetFocus
End If
End Sub

Private Sub cboFabricante_GotFocus()
Dim varNomeAntes As String
varNomeAntes = cboFabricante.Text
   
cboFabricante.Clear

If vTipoOS = "Automóveis" Then
    sSQL = "SELECT DISTINCT fabricante FROM OS_Fabricantes_Carro ORDER BY fabricante;"
ElseIf vTipoOS = "Motocicletas" Then
    sSQL = "SELECT DISTINCT fabricante FROM OS_Fabricante_Moto ORDER BY fabricante;"
ElseIf vTipoOS = "Recapadora" Then
    sSQL = "SELECT DISTINCT fabricante FROM OS_Fabricante_Caminhao ORDER BY fabricante;"
ElseIf vTipoOS = "Informática" Then
    sSQL = "SELECT DISTINCT fabricante FROM OS_Equipamento ORDER BY fabricante;"
ElseIf vTipoOS = "Comunicação Visual" Then
    sSQL = "SELECT DISTINCT fabricante FROM OS_Equipamento ORDER BY fabricante;"
ElseIf vTipoOS = "Celular" Then
    sSQL = "SELECT DISTINCT fabricante FROM OS_Equipamento ORDER BY fabricante;"
End If
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFabricante.AddItem ValidateNull(r("fabricante"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

   cboFabricante.Text = varNomeAntes
   
SelectControl cboFabricante
moCombo.AttachTo cboFabricante
End Sub




Private Sub cboFabricante_LostFocus()
cboFabricante.Text = TirarEspaco(cboFabricante.Text)
End Sub


Private Sub cboFormaPgtoEntrada_Change()
Calcular_Desconto
End Sub

Private Sub cboFormaPgtoEntrada_Click()
Calcular_Desconto
End Sub


Private Sub cboFormaPgtoEntrada_GotFocus()
cboFormaPgtoEntrada.AddItem "1 - DINHEIRO"
cboFormaPgtoEntrada.AddItem "3 - CARTÃO - DÉBITO"
cboFormaPgtoEntrada.AddItem "4 - CARTÃO - CRÉDITO"
cboFormaPgtoEntrada.AddItem "5 - CHEQUE"
cboFormaPgtoEntrada.AddItem "6 - BOLETO"
cboFormaPgtoEntrada.AddItem "7 - TRANSFERÊNCIA"
cboFormaPgtoEntrada.AddItem "8 - DEPOSITO"
cboFormaPgtoEntrada.AddItem "10 - PIX"
End Sub

Private Sub cboModelo_GotFocus()
Dim varNomeAntes As String
varNomeAntes = cboModelo.Text
   
cboModelo.Clear

If vTipoOS = "Automóveis" Then
    sSQL = "SELECT DISTINCT MODELO FROM OS_Modelo_Carro ORDER BY MODELO;"
ElseIf vTipoOS = "Motocicletas" Then
    sSQL = "SELECT DISTINCT MODELO FROM OS_Modelo_Moto ORDER BY MODELO;"
ElseIf vTipoOS = "Recapadora" Then
    sSQL = "SELECT DISTINCT MODELO FROM OS_Modelo_Caminhao ORDER BY MODELO;"
ElseIf vTipoOS = "Informática" Then
    sSQL = "SELECT DISTINCT MODELO FROM OS_Equipamento ORDER BY MODELO;"
ElseIf vTipoOS = "Comunicação Visual" Then
    sSQL = "SELECT DISTINCT MODELO FROM OS_Equipamento ORDER BY MODELO;"
ElseIf vTipoOS = "Celular" Then
    sSQL = "SELECT DISTINCT MODELO FROM OS_Equipamento ORDER BY MODELO;"

End If
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboModelo.AddItem ValidateNull(r("MODELO"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

   cboModelo.Text = varNomeAntes

SelectControl cboModelo
moCombo.AttachTo cboModelo
End Sub

Private Sub cboModelo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub






Private Sub cboFuncionario_GotFocus()
Dim varNomeAntes As String
Dim varCodAntes As String

varNomeAntes = cboFuncionario.Text
varCodAntes = txtCodFuncionario.Text

cboFuncionario.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM funcionario WHERE (cargo <> 'mecânico') ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFuncionario.AddItem r("nome")
   cboFuncionario.ItemData(cboFuncionario.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

txtCodFuncionario.Text = varCodAntes
cboFuncionario.Text = varNomeAntes

cboFuncionario.SelStart = 0
cboFuncionario.SelLength = Len(cboFuncionario)
   
   moCombo.AttachTo cboFuncionario
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFuncionario_LostFocus()
   On Error GoTo TrataErro
   
   If cboFuncionario.Text = "" Then txtCodFuncionario.Text = "": Exit Sub
   
   If cmdAlterar.Enabled = False Then
      If cboFuncionario.ListIndex = -1 Then
         'txtCodFuncionario.Text = ""
         'Exit Sub
      End If
   End If
   
   txtCodFuncionario = cboFuncionario.ItemData(cboFuncionario.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
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

Private Sub cboMecanico_GotFocus()
Dim varNomeAntes As String
Dim varCodAntes As String

varNomeAntes = cboMecanico.Text
varCodAntes = txtCodMecanico.Text

cboMecanico.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM funcionario order by nome;"
'WHERE (cargo IN ('tecnico', 'aux. tecnico')) ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboMecanico.AddItem r("nome")
   cboMecanico.ItemData(cboMecanico.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboMecanico.Text = varNomeAntes
txtCodMecanico.Text = varCodAntes

moCombo.AttachTo cboMecanico
End Sub

Private Sub cboMecanico_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboMecanico_LostFocus()
   On Error GoTo TrataErro
   
   If cboMecanico.Text = "" Then txtCodMecanico.Text = "": Exit Sub
   
   If cmdAlterar.Enabled = False Then
      If cboMecanico.ListIndex = -1 Then
         'txtCodMecanico.Text = ""
         'Exit Sub
      End If
   End If
   
   txtCodMecanico = cboMecanico.ItemData(cboMecanico.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboFabricante_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboModelo_LostFocus()
cboModelo.Text = TirarEspaco(cboModelo.Text)
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    'txtPareceCliente.SetFocus
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    txtPareceCliente.SetFocus
End If
End Sub

Private Sub cboPecas_Change()
If vTipoConsPecas <> 1 Then
    If Len(cboPecas.Text) > 3 Then
        sSQL = "SELECT DISTINCT descricao, codigo FROM produtos WHERE (descricao LIKE '" & cboPecas & "%')  ORDER BY descricao;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboPecas.AddItem ValidateNull(r("descricao"))
            cboPecas.ItemData(cboPecas.NewIndex) = r("codigo")
           r.MoveNext
        Loop
    End If
End If
End Sub

Private Sub cboPrazo_Change()
Calcular_Prazo
End Sub

Private Sub cboPrazo_Click()
   Calcular_Prazo
End Sub
Private Sub cboQuantForma_Change()
cboQuantForma_LostFocus
Mostrar_ValorRestante
End Sub

Private Sub cboQuantForma_Click()
cboQuantForma_Change
End Sub


Private Sub cboQuantForma_GotFocus()
Dim varTexto As String
varTexto = cboQuantForma.Text
    cboQuantForma.Clear
If cboTipoPgto.Text = "À VISTA" Then
    cboQuantForma.AddItem "1 - FORMA"
    cboQuantForma.AddItem "2 - FORMAS"
ElseIf cboTipoPgto.Text = "À PRAZO" Then
    cboQuantForma.AddItem "1 - SEM ENTRADA"
    cboQuantForma.AddItem "2 - COM ENTRADA"
End If

cboQuantForma.Text = varTexto
SelectControl cboQuantForma
moCombo.AttachTo cboQuantForma
End Sub


Private Sub cboQuantForma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      cmdFinalizar_Click
End If
End Sub

Private Sub cboQuantForma_LostFocus()
If cboQuantForma.Text = "1 - FORMA" Then
    lblEntrada.Enabled = False
    txtEntrada.Enabled = False
    lblFormaEntrada.Enabled = False
    cboFormaPgtoEntrada.Enabled = False
    cboformaPgto.Enabled = True
    lblFormaParcelas.Enabled = True
    lblValorParc.Enabled = True
    txtValorRest.Enabled = True
    'txtValorRest.Locked = True
ElseIf cboQuantForma.Text = "2 - FORMAS" Then
    lblEntrada.Enabled = True
    txtEntrada.Enabled = True
    lblFormaEntrada.Enabled = True
    cboFormaPgtoEntrada.Enabled = True
    cboformaPgto.Enabled = True
    lblFormaParcelas.Enabled = True
    lblValorParc.Enabled = True
    txtValorRest.Enabled = True
    'txtValorRest.Locked = True
ElseIf cboQuantForma.Text = "1 - SEM ENTRADA" Then
    lblEntrada.Enabled = False
    txtEntrada.Enabled = False
    lblFormaEntrada.Enabled = False
    cboFormaPgtoEntrada.Enabled = False
    cboformaPgto.Enabled = True
    lblFormaParcelas.Enabled = True
    lblValorParc.Enabled = True
    txtValorRest.Enabled = True
    txtEntrada.Text = Format(0, ocMONEY)
    'txtValorRest.Locked = True
ElseIf cboQuantForma.Text = "2 - COM ENTRADA" Then
    lblEntrada.Enabled = True
    txtEntrada.Enabled = True
    lblFormaEntrada.Enabled = True
    cboFormaPgtoEntrada.Enabled = True
    cboformaPgto.Enabled = True
    lblFormaParcelas.Enabled = True
    lblValorParc.Enabled = True
    txtValorRest.Enabled = True
    'txtValorRest.Locked = True
End If
End Sub


Private Sub cboQuantParc_LostFocus()
   Calcular_Parcelas
   Calcular_Prazo
End Sub

Private Sub cboQuantParc_Validate(Cancel As Boolean)
If cboQuantParc.Text = "" Then cboQuantParc = "1"
End Sub


Private Sub cboServicosAuto_GotFocus()
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    cboServicosAuto.Clear
    sSQL = "SELECT * FROM os_Servicos ORDER BY servico;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
        
       cboServicosAuto.AddItem r("servico")
       cboServicosAuto.ItemData(cboServicosAuto.NewIndex) = r("codigo")
       r.MoveNext
    Loop
    If r.State <> 0 Then r.Close
    Set r = Nothing
ElseIf vTipoOS = "Comunicação Visual" Then
    cboServicosAuto.Clear
    sSQL = "SELECT * FROM os_Servicos ORDER BY servico;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
        
       cboServicosAuto.AddItem r("servico")
       cboServicosAuto.ItemData(cboServicosAuto.NewIndex) = r("codigo")
       r.MoveNext
    Loop
    If r.State <> 0 Then r.Close
    Set r = Nothing
ElseIf vTipoOS = "Recapadora" Then
    sSQL = "SELECT * FROM os_Servicos WHERE (TIPO = '" & cboTipo.Text & "') ORDER BY servico;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboServicosAuto.AddItem r("servico") & " | " & r("TIPO") & " | " & r("MEDIDA") & " | " & r("ARO") & " | " & r("BANDA")
       cboServicosAuto.ItemData(cboServicosAuto.NewIndex) = r("codigo")
       r.MoveNext
    Loop
    'desativei daqui pra frente=============
    'Dim itemAtual As String
    'itemAtual = cboServicosAuto.Text
    'cboServicosAuto.Clear
    'cboServicosAuto.AddItem "RECAPAGEM"
    'cboServicosAuto.AddItem "CONSERTO"
    'cboServicosAuto.AddItem "ENCHIMENTO"
    'cboServicosAuto.AddItem "TELA DE AÇO"
    'cboServicosAuto.AddItem "ENGALOCHAMENTO"
    'cboServicosAuto.AddItem "DUPLAGEM"
    'cboServicosAuto.Text = itemAtual
    moCombo.AttachTo cboServicosAuto
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    cboServicosAuto.Clear
    sSQL = "SELECT * FROM os_Servicos ORDER BY servico;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
        
       cboServicosAuto.AddItem r("servico")
       cboServicosAuto.ItemData(cboServicosAuto.NewIndex) = r("codigo")
       r.MoveNext
    Loop
    If r.State <> 0 Then r.Close
    Set r = Nothing
ElseIf vTipoOS = "Comunicação Visual" Then
    cboServicosAuto.Clear
    sSQL = "SELECT * FROM os_Servicos ORDER BY servico;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
        
       cboServicosAuto.AddItem r("servico")
       cboServicosAuto.ItemData(cboServicosAuto.NewIndex) = r("codigo")
       r.MoveNext
    Loop
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If

moCombo.AttachTo cboServicosAuto
End Sub


Private Sub cboServicosAuto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboServicosAuto_LostFocus()
On Error GoTo TrataErro

If cboServicosAuto.Text = "" Then txtCodServicoAuto.Text = "": Exit Sub
txtCodServicoAuto = cboServicosAuto.ItemData(cboServicosAuto.ListIndex)
Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboSituacao_GotFocus()
cboSituacao.Clear

sSQL = "SELECT DISTINCT situacao, codigo FROM OS_situacao ORDER BY situacao;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboSituacao.AddItem r("situacao")
   cboSituacao.ItemData(cboSituacao.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboSituacao
End Sub


Private Sub cboSituacao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboSituacao_LostFocus()
On Error GoTo TrataErro

If cboSituacao.Text = "" Then txtCodSituacao.Text = "": Exit Sub
If cboSituacao.ListIndex = -1 Then txtCodSituacao.Text = "": Exit Sub
txtCodSituacao = cboSituacao.ItemData(cboSituacao.ListIndex)
Exit Sub

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboStatus_Change()
If cboStatus.Text = "À COMEÇAR" Then
   'cmdImprimirEntrada.Enabled = False
   lblMecanico.Enabled = False
   cboMecanico.Enabled = False
   cmdFinalizarAP.Visible = False
   cmdFinalizarAV.Visible = False
   frmServicos.Enabled = True
ElseIf cboStatus.Text = "EM EXECUÇÃO" Or cboStatus.Text = "AGUARDANDO" Then
   'cmdImprimirEntrada.Enabled = True
   lblMecanico.Enabled = True
   cboMecanico.Enabled = True
   cmdFinalizarAP.Visible = False
   cmdFinalizarAV.Visible = False
   frmServicos.Enabled = True
ElseIf cboStatus.Text = "TERMINADO" Then
   'dbData.Execute "UPDATE OS SET DATA_TERMINO = '" & Format(Date, ocDATA) & "', DATA_TERMINO = '" & Format(Now, ocHORA) & "' WHERE (cod_os = " & txtCodOS.Text & ");"
   lblMecanico.Enabled = True
   cboMecanico.Enabled = True
   frmServicos.Enabled = False
End If
End Sub

Private Sub cboStatus_Click()
   cboStatus_Change
End Sub

Private Sub cboStatus_GotFocus()
   Dim itemAtual As String
   itemAtual = cboStatus.Text
   cboStatus.Clear
   cboStatus.AddItem "À COMEÇAR"
   cboStatus.AddItem "EM EXECUÇÃO"
   cboStatus.AddItem "AGUARDANDO"
   cboStatus.AddItem "TERMINADO"
   cboStatus.Text = itemAtual
   moCombo.AttachTo cboStatus
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboStatus_LostFocus()
cboStatus_Change
If cboStatus.Text = "TERMINADO" Then
    If frmPrincipal.Enabled = True And cboMecanico.Enabled = True Then
        'cboMecanico.SetFocus
    End If
End If
End Sub

Private Sub cboTanque_GotFocus()
Dim varNomeAntes As String

varNomeAntes = cboTanque.Text

cboTanque.Clear

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
    cboTanque.AddItem "VAZIO"
    cboTanque.AddItem "CHEIO"
    cboTanque.AddItem "1/4"
    cboTanque.AddItem "3/4"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    sSQL = "SELECT DISTINCT EQUIPAMENTO FROM OS_Equipamento ORDER BY EQUIPAMENTO;"
    Set r = dbData.OpenRecordset(sSQL)

    Do While Not r.EOF
       cboTanque.AddItem ValidateNull(r("EQUIPAMENTO"))
       r.MoveNext
    Loop
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
ElseIf vTipoOS = "Comunicação Visual" Then
    sSQL = "SELECT DISTINCT EQUIPAMENTO FROM OS_Equipamento ORDER BY EQUIPAMENTO;"
    Set r = dbData.OpenRecordset(sSQL)

    Do While Not r.EOF
       cboTanque.AddItem ValidateNull(r("EQUIPAMENTO"))
       r.MoveNext
    Loop
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If

cboTanque.Text = varNomeAntes
SelectControl cboTanque
moCombo.AttachTo cboTanque
End Sub


Private Sub cboTanque_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboTanque_LostFocus()
cboTanque.Text = TirarEspaco(cboTanque.Text)
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    txtPareceCliente.SetFocus
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    cboFabricante.SetFocus
End If
End Sub

Public Function TirarEspaco(ByVal Value As String) As String
Dim bRepete As Boolean
Value = Replace$(Value, "'", vbNullString)
Do
  Value = Replace$(Value, "  ", " ")
  bRepete = InStr(1, Value, "  ", vbTextCompare)
  Value = Trim(Value)
Loop Until Not bRepete

TirarEspaco = Value
End Function

Private Sub cboMarca_GotFocus()
cboMarca.Clear

sSQL = "SELECT DISTINCT FABRICANTE FROM OS_Servicos_Recapadora ORDER BY FABRICANTE;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboMarca.AddItem ValidateNull(r("FABRICANTE"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboMarca
End Sub


Private Sub cboMarca_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboTipo_GotFocus()
Dim itemAtual As String
itemAtual = cboTipo.Text
cboTipo.Clear
cboTipo.AddItem "AGRICOLA"
cboTipo.AddItem "CARGA"
cboTipo.AddItem "CAMINHONETE"
cboTipo.Text = itemAtual
moCombo.AttachTo cboTipo
End Sub


Private Sub cboTipo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboTipoOS_GotFocus()
Dim varNomeAntes As String

varNomeAntes = cboTipoOS.Text

MostrarTipoOS

cboTipoOS.Text = varNomeAntes

moCombo.AttachTo cboTipoOS
End Sub

Private Sub cboTipoOS_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



Private Sub cboTipoPgto_Change()
If cboTipoPgto.Text = "À VISTA" Then
    txtEntrada.Enabled = False
    cboPrazo.Enabled = False
    txtValorRest.Enabled = False
    cboQuantParc.Enabled = False
    txtValorParc.Enabled = False
    mskInicio.Enabled = False
    mskTermino.Enabled = False
    lblEntrada.Enabled = False
    lblQuantParc.Enabled = False
    lblValorParc.Enabled = False
    'Label6.Enabled = False 'saber
    'Label7.Enabled = False
    lblInicio.Enabled = False
    Label17.Enabled = False
    cboFormaPgtoEntrada.Enabled = False
    lblFormaEntrada.Enabled = False
    'BuscarClienteConsumidor
    cboformaPgto.Text = "1 - DINHEIRO"
    cboFormaPgtoEntrada.Text = "1 - DINHEIRO"
ElseIf cboTipoPgto.Text = "À PRAZO" Then
    txtEntrada.Enabled = True
    cboPrazo.Enabled = True
    txtValorRest.Enabled = True
    cboQuantParc.Enabled = True
    txtValorParc.Enabled = True
    mskInicio.Enabled = True
    mskTermino.Enabled = True
    lblEntrada.Enabled = True
    lblQuantParc.Enabled = True
    lblValorParc.Enabled = True
    'Label6.Enabled = True  'saber
    'Label7.Enabled = True  'saber
    lblInicio.Enabled = True
    Label17.Enabled = True
    cboFormaPgtoEntrada.Enabled = True
    lblFormaEntrada.Enabled = True
    cboformaPgto.Text = "1 - PROMISSÓRIA"
    cboFormaPgtoEntrada.Text = "1 - DINHEIRO"
    
ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
    txtEntrada.Enabled = False
    cboPrazo.Enabled = False
    txtValorRest.Enabled = False
    cboQuantParc.Enabled = False
    txtValorParc.Enabled = False
    mskInicio.Enabled = False
    mskTermino.Enabled = False
    lblEntrada.Enabled = False
    lblQuantParc.Enabled = False
    lblValorParc.Enabled = False
    'Label6.Enabled = False 'saber
    'Label7.Enabled = False 'saber
    lblInicio.Enabled = False
    Label17.Enabled = False
    'BuscarClienteConsumidor
    cboFormaPgtoEntrada.Enabled = True
    lblFormaEntrada.Enabled = True
    cboformaPgto.Text = "1 - DINHEIRO"
    cboFormaPgtoEntrada.Text = "1 - DINHEIRO"
End If

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


Private Sub ccmdIncluirAcessório_Click()
If cboAcessorios.Text = "" Or txtCodOS.Text = "" Then Exit Sub

'CHECAR SE A OS ESTÁ FECHADA
'Verificar_OS_Fechada
'If OS_FECHADA = True Then Exit Sub

'ADICIONAR NA TABELA OS SERVIÇOS
AutoNumeracao_Acessorio
dbData.Execute "INSERT INTO OS_acessorios (codigo, acessorio) VALUES(" & xAcess & ", '" & cboAcessorios.Text & "')"

cboAcessorios.Text = ""
txtCodAcessorio.Text = ""
cboAcessorios.SetFocus
End Sub

Private Sub ccmdIncluirSituacao_Click()
If cboSituacao.Text = "" Or txtCodOS.Text = "" Then Exit Sub

'CHECAR SE A OS ESTÁ FECHADA
'Verificar_OS_Fechada
'If OS_FECHADA = True Then Exit Sub

'ADICIONAR NA TABELA OS SERVIÇOS
AutoNumeracao_Situacao
dbData.Execute "INSERT INTO OS_Situacao (codigo, SITUACAO) VALUES(" & xAcess & ", '" & cboSituacao.Text & "')"

cboSituacao.Text = ""
txtCodSituacao.Text = ""
cboSituacao.SetFocus
End Sub


Private Sub chameleonButton1_Click()
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

mskDataSaida = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub Verificar_OS_FechadaePaga()
sSQL = "SELECT cod_os, status_os FROM os WHERE (cod_os = " & txtCodOS.Text & ") AND (status_os = 0);"
Set r = dbData.OpenRecordset(sSQL)

OS_FINANCEIROABERTO = (r.RecordCount <> 0)

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Verificar_OS_Fechada()
sSQL = "SELECT cod_os, status_os FROM os WHERE (cod_os = " & txtCodOS.Text & ") AND (status_os = 1);"
Set r = dbData.OpenRecordset(sSQL)

OS_FECHADA = False

If r.RecordCount <> 0 Then
   ShowMsg "ESTA O.S. JÁ ESTÁ FECHADA!", vbExclamation
   OS_FECHADA = True
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Verificar_Caixa()
sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
    'MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
    varCodCaixa = CInt(0)
    StatusBar1.Panels(3).Text = Format(varCodCaixa, "0000")
    CAIXA_FECHADO = True
    
Else
    If CDate(r("DATA_ABERTURA")) <> Date Then
        MsgBox "A data do caixa aberto é diferente da data atual!", vbInformation, "Aviso do Sistema"
        'lblAlerta.Visible = True
        'lblRotuloAberto.Visible = True
        'lblDataAberturaCaixa.Visible = True
    Else
        'lblAlerta.Visible = False
        'lblRotuloAberto.Visible = False
        'lblDataAberturaCaixa.Visible = False
    End If
    varCodCaixa = CInt(r("codcaixa"))
    StatusBar1.Panels(3).Text = Format(varCodCaixa, "0000")
    lblDataAberturaCaixa.Caption = CDate(r("DATA_ABERTURA"))
    CAIXA_FECHADO = False
End If
End Sub


Private Sub Verificar_Limite()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim Limite As Currency
   Dim LimiteAtual As Currency
   Dim Parcelas_Abertas As Currency
   
   Passou_Limite = False
   
   'ver o limite do cliente
   If txtCodCliente.Text = "" Then Exit Sub
   
   sSQL = "SELECT codigo, limite_credito FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then Limite = r("limite_credito")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   If Limite = 0 Then
      Passou_Limite = False
      Exit Sub
   End If
   
   'ver as parcelas em aberto do cliente
   Parcelas_Abertas = 0
   
   sSQL = "SELECT pedidos.cod_cliente, ISNULL(SUM(parcelas.valor), 0) AS somas_parcelas " & _
      "FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
      "WHERE (pedidos.cod_cliente = " & txtCodCliente.Text & ") AND (parcelas.status = 0) " & _
      "GROUP BY pedidos.cod_cliente;"
   
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then Parcelas_Abertas = r("somas_parcelas")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   If Limite <= Parcelas_Abertas Then
      ShowMsg "O cliente passou do limite de crédito dele!", vbInformation
      Passou_Limite = True
   Else
      Passou_Limite = False
   End If
   
End Sub


Private Sub Mostrar_Desconto()
If vTipoDesc = "1" Then
    txtDesc.Text = Format(0, ocPESO)
ElseIf vTipoDesc = "2" Then
    If cboTipoPgto.Text = "À VISTA" Then
        txtDesc.Text = Format(vValorDescFixoAV, ocPESO)
    ElseIf cboTipoPgto.Text = "À PRAZO" Then
        txtDesc.Text = Format(vValorDescFixoAP, ocPESO)
    ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
        txtDesc.Text = Format(vValorDescFixoAV, ocPESO)
    End If
ElseIf vTipoDesc = "3" Then
    Dim vSubtotal As Currency
    vSubtotal = txtSubtotal.Text

If vSubtotal <= vMargemDescGradual1 Then
    vEtapa = 1
ElseIf vSubtotal > vMargemDescGradual1 And vSubtotal <= vMargemDescGradual2 Then
    vEtapa = 2
ElseIf vSubtotal > vMargemDescGradual2 Then
    vEtapa = 3
Else
    vEtapa = 1
End If

    If cboTipoPgto.Text = "À VISTA" Then
        If vEtapa = 1 Then
            txtDesc.Text = Format(vValorDescGradualAV1, ocPESO)
        ElseIf vEtapa = 2 Then
            txtDesc.Text = Format(vValorDescGradualAV2, ocPESO)
        ElseIf vEtapa = 3 Then
            txtDesc.Text = Format(vValorDescGradualAV3, ocPESO)
        End If
    ElseIf cboTipoPgto.Text = "À PRAZO" Then
        If vEtapa = 1 Then
            txtDesc.Text = Format(vValorDescGradualAP1, ocPESO)
        ElseIf vEtapa = 2 Then
            txtDesc.Text = Format(vValorDescGradualAP2, ocPESO)
        ElseIf vEtapa = 3 Then
            txtDesc.Text = Format(vValorDescGradualAP3, ocPESO)
        End If
    ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
        If vEtapa = 1 Then
            txtDesc.Text = Format(vValorDescGradualAV1, ocPESO)
        ElseIf vEtapa = 2 Then
            txtDesc.Text = Format(vValorDescGradualAV2, ocPESO)
        ElseIf vEtapa = 3 Then
            txtDesc.Text = Format(vValorDescGradualAV3, ocPESO)
        End If
    End If
End If
End Sub


Private Sub chkVeiculo_Click()
If chkVeiculo.Value = Checked Then
    frmEquipamento.Visible = True
    txtObs.Visible = False
Else
    frmEquipamento.Visible = False
    txtObs.Visible = True
End If
End Sub

Private Sub cmdAdicionarAcessorios_Click()
If txtCodAcessorio.Text = "" Or txtCodOS.Text = "" Then Exit Sub

'CHECAR SE A OS ESTÁ FECHADA
Verificar_OS_Fechada
If OS_FECHADA = True Then Exit Sub

'ADICIONAR NA TABELA OS SERVIÇOS
AutoNumeracao_Acessorio_Auto
dbData.Execute "INSERT INTO OS_acessorios_Auto (codigo, cod_os, cod_acessorio, acessorio) VALUES(" & xAcess & ", " & txtCodOS.Text & ", " & txtCodAcessorio.Text & ", '" & cboAcessorios.Text & "')"

MostrarGrid_Acessorios

txtCodAcessorio.Text = ""
cboAcessorios.Text = ""
cboAcessorios.SetFocus
End Sub

Private Sub cmdAdicionarPecas_Click()
   Dim QUANT As Integer
   Dim Valor As Currency
   Dim Total As Currency
   Dim sSQL As String
   
   If txtCodPeca.Text = "" Or txtCodOS.Text = "" Or txtCodPedido.Text = "" Then Exit Sub
   If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1
   If txtSubtotalPecas.Text = "" Then Exit Sub
   If txtValorPeca.Text = "" Or txtValorPeca.Text = "0,00" Then Exit Sub
   
   'CHECAR SE A OS ESTÁ FECHADA
   Verificar_OS_Fechada
   If OS_FECHADA = True Then Exit Sub
   
   'VERIFICAR O STATUS
   'If cboStatus.Text = "À COMEÇAR" Then
   '   ShowMsg "Não é possivel adicionar peças em uma OS com Status = À COMEÇAR!", vbExclamation
   '   Exit Sub
   'End If
   
   'Verifica_Quantidade do Estoque
   Verifica_QuantEstoque
   If VERIFICAR_QUANTIDADE = True Then
      LimparObjetos_Pecas
      Exit Sub
   End If
   
   'calcular o total das peças no grid
   If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1
   
   'If txtQuantPeca.Text <> "" Or txtValorPeca.Text <> "" Then
   '   QUANT = txtQuantPeca.Text
   '   Valor = txtValorPeca.Text
   '   Total = Valor * QUANT
   'End If
   
   'adicionar na tabela PEDIDOS_ITENS
   AutoNumeracao_Peca
   AutoNumeracao_PecaItem

   sSQL = "INSERT INTO pedidos_itens (" & _
      "codigo, " & _
      "item, " & _
      "cod_pedido, " & _
      "cod_produto, " & _
      "preco, " & _
      "quantidade, " & _
      "SUBTOTAL, " & _
      "DESCONTO, " & _
      "TOTAL, " & _
      "data, " & _
      "tipo_venda) " & _
      "VALUES (" & _
      xPeca & ", " & _
      xPecaItem & ", " & _
      "" & txtCodPedido.Text & ", " & _
      "" & txtCodPeca.Text & ", " & _
      "" & Replace(CCur(txtValorPeca.Text), ",", ".") & ", " & _
      "" & Replace(CDbl(txtQuantPeca.Text), ",", ".") & ", " & _
      "" & Replace(CCur(txtSubtotalPecas.Text), ",", ".") & ", " & _
      "" & Replace(CCur(txtDescPecas.Text), ",", ".") & ", " & _
      "" & Replace(CCur(txtTotalPeca.Text), ",", ".") & ", " & _
      "CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & _
      "'OFICINA')"

   dbData.Execute sSQL
   
MostrarGrid_Servicos

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Recapadora" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Comunicação Visual" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
End If

LimparObjetos_Pecas
cboPecas.SetFocus
Somar_Totais
End Sub

Private Sub Verifica_QuantEstoque()
Dim oCfg As ConfigItem
Dim bEstNeg As Boolean

If txtCodPeca.Text = "" Then Exit Sub
If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1

Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
bEstNeg = CBool(oCfg.Value)
Set oCfg = Nothing

If bEstNeg = False Then
   sSQL = "SELECT codigo, quant_estoque FROM produtos WHERE (codigo = " & txtCodPeca.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   VERIFICAR_QUANTIDADE = False
   
   If Not r.BOF Then
      If r("quant_estoque") < CDbl(txtQuantPeca.Text) And r("quant_estoque") <> 0 Then
         ShowMsg "ESSA QUANTIDADE É INVÁLIDA!" & vbCrLf & "SEU ESTOQUE ATUAL É DE " & r("quant_estoque") & " PRODUTO(S)", vbExclamation
         LimparObjetos_Pecas
         VERIFICAR_QUANTIDADE = True
         
      ElseIf r("quant_estoque") = 0 Then
         ShowMsg "ESSA QUANTIDADE É INVÁLIDA!" & vbCrLf & "SEU ESTOQUE ATUAL É DE 0 PRODUTO(S)", vbExclamation
         LimparObjetos_Pecas
         VERIFICAR_QUANTIDADE = True
         
      End If
   End If
Else
   Exit Sub
End If
End Sub

Private Sub cmdAdicionarServicosAuto_Click()

If txtCodServicoAuto.Text = "" Or txtCodOS.Text = "" Then Exit Sub
If txtQuantServicoAuto.Text = "" Then txtQuantServicoAuto.Text = 1
If mskValorServicoAuto.Text = "" Or mskValorServicoAuto.Text = "0,00" Then Exit Sub

'CHECAR SE A OS ESTÁ FECHADA
Verificar_OS_Fechada
If OS_FECHADA = True Then Exit Sub

'ADICIONAR NA TABELA OS SERVIÇOS
AutoNumeracao_Servico

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    dbData.Execute "INSERT INTO OS_Servicos_Auto (codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data) VALUES (" & _
       xServ & ", " & txtCodOS.Text & ", '" & vServico & "', " & Replace(CCur(mskValorServicoAuto.Text), ",", ".") & ", " & _
       txtQuantServicoAuto.Text & ", " & Replace(CCur(txtSubTotalServicoAuto.Text), ",", ".") & ", " & Replace(CCur(txtDescServicoAuto.Text), ",", ".") & ", " & Replace(CCur(txtTotalServicoAuto.Text), ",", ".") & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103))"
ElseIf vTipoOS = "Recapadora" Then
    dbData.Execute "INSERT INTO os_servicos_recapadora (codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data, DOTE, TIPO, SERIE, FOGO, FABRICANTE, MEDIDA, ARO, BANDA) VALUES (" & _
       xServ & ", " & txtCodOS.Text & ", '" & vServico & "', " & Replace(CCur(mskValorServicoAuto.Text), ",", ".") & ", " & _
       txtQuantServicoAuto.Text & ", " & Replace(CCur(txtSubTotalServicoAuto.Text), ",", ".") & ", " & Replace(CCur(txtDescServicoAuto.Text), ",", ".") & ", " & Replace(CCur(txtTotalServicoAuto.Text), ",", ".") & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), '" & txtDote.Text & "', '" & cboTipo.Text & "', '" & txtSerie.Text & "', '" & txtFogo.Text & "', '" & cboMarca.Text & "', '" & vMedida & "', '" & vAro & "', '" & vBanda & "')"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "INSERT INTO OS_Servicos_Auto (codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data) VALUES (" & _
       xServ & ", " & txtCodOS.Text & ", '" & vServico & "', " & Replace(CCur(mskValorServicoAuto.Text), ",", ".") & ", " & _
       txtQuantServicoAuto.Text & ", " & Replace(CCur(txtSubTotalServicoAuto.Text), ",", ".") & ", " & Replace(CCur(txtDescServicoAuto.Text), ",", ".") & ", " & Replace(CCur(txtTotalServicoAuto.Text), ",", ".") & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103))"
ElseIf vTipoOS = "Comunicação Visual" Then
    dbData.Execute "INSERT INTO OS_Servicos_Comunicacao (codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data, obs) VALUES (" & _
       xServ & ", " & txtCodOS.Text & ", '" & vServico & "', " & Replace(CCur(mskValorServicoAuto.Text), ",", ".") & ", " & _
       txtQuantServicoAuto.Text & ", " & Replace(CCur(txtSubTotalServicoAuto.Text), ",", ".") & ", " & Replace(CCur(txtDescServicoAuto.Text), ",", ".") & ", " & Replace(CCur(txtTotalServicoAuto.Text), ",", ".") & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), '" & txtObsServ.Text & "')"
End If

MostrarGrid_Servicos

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Recapadora" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Comunicação Visual" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"

End If

LimparObjetos_ServicosAuto
If cboTipo.Visible = True Then cboTipo.SetFocus Else cboServicosAuto.SetFocus
Somar_Totais
End Sub
Private Sub cmdAdicionarSituacao_Click()
If txtCodSituacao.Text = "" Or txtCodOS.Text = "" Then Exit Sub

'CHECAR SE A OS ESTÁ FECHADA
Verificar_OS_Fechada
If OS_FECHADA = True Then Exit Sub

'ADICIONAR NA TABELA OS SERVIÇOS
AutoNumeracao_Situacao_Auto
dbData.Execute "INSERT INTO OS_Situacao_Auto (codigo, cod_os, COD_SITUACAO, SITUACAO) VALUES(" & xAcess & ", " & txtCodOS.Text & ", " & txtCodSituacao.Text & ", '" & cboSituacao.Text & "')"

MostrarGrid_Situacao

txtCodSituacao.Text = ""
cboSituacao.Text = ""
cboSituacao.SetFocus
End Sub

Private Sub cmdAlterar_Click()

If cboStatus.Text = "TERMINADO" Then
    'ver a quantidade de peças e serviços da ordem de serviços
    If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
        vTabelaServicos = "OS_Servicos_Auto"
    ElseIf vTipoOS = "Recapadora" Then
        vTabelaServicos = "OS_Servicos_recapadora"
    ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
        vTabelaServicos = "OS_Servicos_Auto"
    ElseIf vTipoOS = "Comunicação Visual" Then
        vTabelaServicos = "OS_Servicos_Comunicacao"
    End If

   sSQL = "SELECT cod_os FROM " & vTabelaServicos & " WHERE (cod_os = " & txtCodOS & ")"
   Set r_Itens = dbData.OpenRecordset(sSQL)
   
   If r_Itens.EOF Then
        MsgBox "Não é permitido finalizar uma ordem de serviços somente com produtos!", vbExclamation, "Aviso do Sistema"
        Exit Sub
   End If
End If

If txtCodOS.Text = "" Then
   ShowMsg "OS VAZIA! Selecione uma OS na guia FILTRO!", vbInformation
   Exit Sub
End If

If txtCodCliente.Text = "" Then
   ShowMsg "Este cliente não encontra-se cadastrado!", vbInformation
   cboCliente.SetFocus
   Exit Sub
End If

If txtCodFuncionario.Text = "" Then
   ShowMsg "Este funcionário não encontra-se cadastrado!", vbInformation
   cboFuncionario.SetFocus
   Exit Sub
End If

If cboStatus.Text = "EM EXECUÇÃO" Or cboStatus.Text = "AGUARDANDO" Or cboStatus.Text = "TERMINADO" Then
   If cboMecanico.Text = "" Then
      ShowMsg "Indique o nome do do responsavel técnico pelo equipamento!", vbInformation
      cboMecanico.SetFocus
      Exit Sub
   End If
    cmdImpEntrada2.Enabled = True
    cmdImpOrcamento2.Enabled = True
    cmdImpPedido2.Enabled = True
End If
 
'Faz a atualização de forma direta e verifica se houve algum erro
If Not Atualizar_Dados_OS Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

'editar tabela pedidos
dbData.Execute "UPDATE pedidos SET cod_cliente = " & txtCodCliente.Text & " WHERE (cod_pedido = " & txtCodPedido.Text & ");"

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    dbData.Execute "UPDATE OS_Equipamento_Auto SET fabricante = '" & cboFabricante.Text & "', modelo = '" & cboModelo.Text & "', placa = '" & txtPlaca.Text & "', ano = '" & txtAno.Text & "', km = '" & txtKM.Text & "', CHASSI = '" & txtChassi.Text & "', COR = '" & cboCor.Text & "', TANQUE = '" & cboTanque.Text & "', PARECER_CLIENTE = '" & txtPareceCliente.Text & "' WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Recapadora" Then
    dbData.Execute "UPDATE OS_Equipamento_Auto SET fabricante = '" & cboFabricante.Text & "', modelo = '" & cboModelo.Text & "', placa = '" & txtPlaca.Text & "', ano = '" & txtAno.Text & "', km = '" & txtKM.Text & "', CHASSI = '" & txtChassi.Text & "', COR = '" & cboCor.Text & "', TANQUE = '" & cboTanque.Text & "', PARECER_CLIENTE = '" & txtPareceCliente.Text & "' WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "UPDATE OS_Equipamento SET fabricante = '" & cboFabricante.Text & "', modelo = '" & cboModelo.Text & "', equipamento = '" & cboTanque.Text & "', PARECER_CLIENTE = '" & txtPareceCliente.Text & "' WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Comunicação Visual" Then
    'dbData.Execute "UPDATE OS_Equipamento SET fabricante = '" & cboFabricante.Text & "', modelo = '" & cboModelo.Text & "', equipamento = '" & cboTanque.Text & "', PARECER_CLIENTE = '" & txtPareceCliente.Text & "' WHERE (cod_os = " & txtCodOS.Text & ");"
Else

End If

If cboTipoOS.Text = "CONSERTO" Or cboTipoOS.Text = "MONTAGEM" Or cboTipoOS.Text = "ASSISTENCIA" Or cboTipoOS.Text = "AUTOMAÇÃO" Or cboTipoOS.Text = "CONSULTORIA" Or cboTipoOS.Text = "CONFECÇÃO" Then
   'CHECAR SE A OS ESTÁ FECHADA & PAGA
   Verificar_OS_FechadaePaga
   
   If OS_FINANCEIROABERTO = True Then
      If cboStatus.Text = "TERMINADO" Then
         SSTab1.Tab = 2
         cmdFinalizarAV.Visible = True
         cmdFinalizarAP.Visible = True
         cmdFinalizarAV.Enabled = True
         cmdFinalizarAP.Enabled = True
      End If
   Else
      cmdFinalizarAV.Visible = False
      cmdFinalizarAP.Visible = False
   End If

ElseIf cboTipoOS.Text = "GARANTIA" And cboStatus.Text = "TERMINADO" Then
   'ATUALIZAR A TABELA OS
   dbData.Execute "UPDATE os SET status_os = 1 WHERE (cod_os = " & txtCodOS.Text & ");"

   'ATUALIZANDO A TABELA PEDIDOS
   dbData.Execute "UPDATE pedidos SET tipo_desc = null, valor_desc = null, tipo_acrescimo = null, valor_acrescimo = null, subtotal = null, total = null, tipo_pagamento = null, pagamento = null, entrada = null, tipo_pedido = 'OFICINA', maquina = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', status_pedido = 1, validade = null WHERE (cod_pedido = " & txtCodPedido.Text & ");"

ElseIf cboTipoOS.Text = "ORÇAMENTO" And cboStatus.Text = "TERMINADO" Then
   'ATUALIZAR A TABELA OS
   dbData.Execute "UPDATE os SET status_os = 0 WHERE (cod_os = " & txtCodOS.Text & ");"

   'ATUALIZANDO A TABELA PEDIDOS
   dbData.Execute "UPDATE pedidos SET tipo_desc = 'P', valor_desc = 0, tipo_acrescimo = 'P', valor_acrescimo = 0, subtotal = " & Replace(CCur(txtTotalPecasServicos.Text), ",", ".") & ", total = " & Replace(CCur(txtTotalPecasServicos.Text), ",", ".") & ", tipo_pagamento = 'À Vista', pagamento = 'AVULSO', entrada = 0, tipo_pedido = 'OFICINA', maquina = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', status_pedido = 1 WHERE (cod_pedido = " & txtCodPedido.Text & ");"
    ', validade = CONVERT(DATETIME, '" & Format(lblValidade.Caption, ocDATA) & "', 103)  'desativei essa parte por causa do campo nao encontrado
   'menu_Impressao_Orcamento_Click
End If

''MostrarGrid_OS
ShowMsg "ALTERAÇÃO DOS DADOS" & vbCr & "Confirmada com sucesso!!", vbExclamation

If cboStatus.Text <> "TERMINADO" Then
    MostrarGrid_OS
    MostrarGrid_OS_Situacao
    If cboStatus.Text <> "EM EXECUÇÃO" Then
        LimparObjetos_Entrada
        LimparObjetos_Servicos
        LimparObjetos_Pecas
        txtCodOS.Text = ""
        txtCodPedido.Text = ""
        Form_Load
    Else
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
            frmParecerCliente.Visible = False
            frmAcessorios.Visible = False
            frmSituacao.Visible = False
            'frmServicos.Visible = True
            frmGridServicos.Visible = True
            frmTotaisGeral.Visible = True
            frmTotaisProdServ.Visible = True
            stProdSer.Visible = True
        ElseIf vTipoOS = "Recapadora" Then
            frmParecerCliente.Visible = False
            frmAcessorios.Visible = False
            frmSituacao.Visible = False
            'frmServicos.Visible = True
            frmGridServicos.Visible = True
            frmTotaisGeral.Visible = True
            frmTotaisProdServ.Visible = True
            stProdSer.Visible = True
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            frmParecerCliente.Visible = False
            frmAcessorios.Visible = False
            frmSituacao.Visible = False
            'frmServicos.Visible = True
            frmGridServicos.Visible = True
            frmTotaisGeral.Visible = True
            frmTotaisProdServ.Visible = True
            stProdSer.Visible = True
        ElseIf vTipoOS = "Comunicação Visual" Then
            frmParecerCliente.Visible = False
            frmAcessorios.Visible = False
            frmSituacao.Visible = False
            'frmServicos.Visible = True
            frmGridServicos.Visible = True
            frmTotaisGeral.Visible = True
            frmTotaisProdServ.Visible = True
            stProdSer.Visible = True
        End If
    
    End If
End If
End Sub

Private Sub cmdApagar_Click()
If txtCodOS.Text = "" Or txtCodCliente.Text = "" Or txtCodFuncionario.Text = "" Then Exit Sub

If ShowMsg("Excluir essa O.S. ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

Retorna_Produtos_Estoque

'EXCLUIR NA TABELA OS
dbData.Execute "DELETE FROM os WHERE (cod_os = " & txtCodOS.Text & ");"

'EXCLUIR NA TABELA PEDIDOS_ITENS
dbData.Execute "DELETE FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ");"

'EXCLUIR NA TABELA PEDIDOS
dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"

'EXCLUIR NA TABELA PARCELAS
dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & txtCodPedido.Text & ");"

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    dbData.Execute "DELETE FROM OS_Equipamento_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM os_servicos_recapadora WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Recapadora" Then
    dbData.Execute "DELETE FROM OS_Equipamento_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM os_servicos_recapadora WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "DELETE FROM OS_Equipamento WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM os_servicos_recapadora WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Comunicação Visual" Then
    'dbData.Execute "DELETE FROM OS_Equipamento WHERE (cod_pedido = " & txtCodPedido.Text & ");"
    'dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
    dbData.Execute "DELETE FROM OS_Servicos_Comunicacao WHERE (cod_os = " & txtCodOS.Text & ");"
    'dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"
End If

LimparObjetos_Entrada
LimparObjetos_Servicos
LimparObjetos_Pecas
txtCodOS.Text = ""
txtCodPedido.Text = ""
MostrarGrid_OS
MostrarGrid_OS_Situacao
Form_Load
End Sub

Private Sub Retorna_Produtos_Estoque()
'Dim i As Integer

'For i = 1 To Grid_Pecas.Rows - 1
'   dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(CDbl(Grid_Pecas.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Pecas.TextMatrix(i, 2) & ");"
'Next
End Sub

Private Sub cmdCal1_Click()
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

mskDataEntrada = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdCal2_Click()
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

Private Sub cmdCancelar_Click()
LimparObjetos_Prazo
frmVendaFechamento.Visible = False
cmdFinalizarAV.Enabled = True
cmdFinalizarAP.Enabled = True
txtTotalPecasServicos.Text = Format(txtSubtotal.Text, ocMONEY)
End Sub

Private Sub cmdCancelarParecer_Click()
frmParecer.Visible = False
End Sub

Private Sub cmdEditarOS_Click()
Dim posit As Long
posit = Grid_OS.Row

If txtCodOS.Text <> "" And cmdGerarEntrada.Enabled = True Then
    MsgBox "A Ordem de Serviço iniciada ainda não foi salvo", vbInformation, "Aviso do Sistema"
    SSTab1.Tab = 1
    Exit Sub
End If

SSTab1.Tab = 1
frmSecundario.Enabled = True
cboStatus.Enabled = True
cmdGerarEntrada.Enabled = False
cmdCancelarEntrada.Enabled = False
cmdAlterar.Enabled = True
cmdApagar.Enabled = True
cmdNovo.Enabled = True

txtCodOS.Text = ""
txtCodOS.Text = (Grid_OS.TextMatrix(Grid_OS.Row, 0))

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    frmEquipamento.Visible = True
    txtObs.Visible = False
    chkVeiculo.Visible = True
    cboCliente.Width = 6795
ElseIf vTipoOS = "Recapadora" Then
    frmEquipamento.Visible = False
    txtObs.Visible = True
    chkVeiculo.Visible = True
    cboCliente.Width = 6795
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    frmEquipamento.Visible = True
    txtObs.Visible = False
    chkVeiculo.Visible = True
    cboCliente.Width = 6795
ElseIf vTipoOS = "Comunicação Visual" Then
    frmEquipamento.Visible = False
    txtObs.Visible = True
    chkVeiculo.Visible = False
    cboCliente.Width = 7755
End If

If (Trim(Grid_OS.TextMatrix(posit, 1))) = ("À COMEÇAR") Then
    If vTipoOS = "Comunicação Visual" Then
        stProdSer.Visible = True
    Else
        stProdSer.Visible = False
    End If
Else
    stProdSer.Visible = True
    stProdSer.Tab = 0
End If
End Sub

Private Sub cmdExcluir_Click()
If ShowMsg("Tem certeza que deseja excluir essa Ordem de Serviço ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

'buscando os dados do formulário
i = Grid_OS.Row

vCodOS = Grid_OS.TextMatrix(i, 0)
codPedido = Grid_OS.TextMatrix(i, 7)

Retorna_Produtos_Estoque

'EXCLUIR NA TABELA OS
dbData.Execute "DELETE FROM os WHERE (cod_os = " & vCodOS & ");"

'EXCLUIR NA TABELA PEDIDOS_ITENS
dbData.Execute "DELETE FROM pedidos_itens WHERE (cod_pedido = " & codPedido & ");"

'EXCLUIR NA TABELA PEDIDOS
dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & codPedido & ");"

'EXCLUIR NA TABELA PARCELAS
dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & codPedido & ");"

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    dbData.Execute "DELETE FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM os_servicos_recapadora WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & vCodOS & ");"
ElseIf vTipoOS = "Recapadora" Then
    dbData.Execute "DELETE FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM os_servicos_recapadora WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & vCodOS & ");"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "DELETE FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM os_servicos_recapadora WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & vCodOS & ");"
ElseIf vTipoOS = "Comunicação Visual" Then
    'dbData.Execute "DELETE FROM OS_Equipamento WHERE (cod_pedido = " & codPedido & ");"
    'dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & vCodOS & ");"
    dbData.Execute "DELETE FROM OS_Servicos_Comunicacao WHERE (cod_os = " & vCodOS & ");"
    'dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & vCodOS & ");"
End If

'LimparObjetos_Entrada
'LimparObjetos_Servicos
'LimparObjetos_Pecas
'txtCodOS.Text = ""
'txtCodPedido.Text = ""
MostrarGrid_OS
MostrarGrid_OS_Situacao
'Form_Load
End Sub

Private Sub cmdExibir_Click()
MostrarGrid_OS
End Sub

Private Sub Imprimir_Pedido()
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

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

'buscando os dados do formulário
'i = Grid_OS.Row

vCodOS = txtCodOS.Text

If txtCodPedido.Text = "" Then
    vCodPedido = codPedido
Else
    codPedido = txtCodPedido.Text
    vCodPedido = txtCodPedido.Text
End If

If vCodOS = "00000" Then MsgBox "Pedido gerado anterior as alterações não permite reimpressão de pedidos. Somente orçamento!", vbInformation, "Aviso do Sistema": Exit Sub

'ver a quantidade de peças e serviços da ordem de serviços
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Recapadora" Then
    vTabelaServicos = "OS_Servicos_recapadora"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Comunicação Visual" Then
    vTabelaServicos = "OS_Servicos_Comunicacao"
End If

'somando os produtos
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM pedidos_itens WHERE (cod_pedido = " & vCodPedido & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
Dim vQuantProduto As Double
Dim vTotalProduto As Currency
vQuantProduto = ValidateNull(r_Itens("VarQuant"))
vTotalProduto = ValidateNull(r_Itens("VarSoma"))
'somando os serviços
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM " & vTabelaServicos & " WHERE (cod_os = " & vCodOS & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
'Debug.Print sSQL_Itens
Dim vQuantServico As Double
Dim vTotalServico As Currency
vQuantServico = ValidateNull(r_Itens("VarQuant"))
vTotalServico = ValidateNull(r_Itens("VarSoma"))
Dim vSomaTotais As Currency
Dim vSomaQuant As Double
vSomaTotais = vTotalProduto + vTotalServico
vSomaQuant = vQuantProduto + vQuantServico

'TOTAIS RESUMIDO
REL_OS_Completo.txtQuantServicos.Caption = " " & Format(vQuantServico, "000")
REL_OS_Completo.txtQuantPecas.Caption = " " & Format(vQuantProduto, "000")
REL_OS_Completo.txtQuantGeral.Caption = " " & Format(vSomaQuant, "000")

REL_OS_Completo.txtTotalServicos.Caption = " " & FormatNumber(vTotalServico, 2)
REL_OS_Completo.txtTotalPecas.Caption = " " & FormatNumber(vTotalProduto, 2)
REL_OS_Completo.txtTotalPecasServicos.Caption = " " & FormatNumber(vSomaTotais, 2)
    
sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & vCodPedido & ")"
sSQL_Itens = sSQL_Itens & " UNION ALL "
sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & vCodOS & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

'Debug.Print sSQL_Itens

'buscar a ordem de serviços para impressão
sSQL = "SELECT * FROM os WHERE (cod_os = " & vCodOS & ");"
Set r = dbData.OpenRecordset(sSQL)

varImpPDF = False
Me.Hide
If r("TIPO_PAGAMENTO") = "À Prazo" Then
    'If r_Itens.RecordCount > 16 Then
        ''REL_OS_PedidoPrazo_Grande.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
        ''Unload REL_OS_PedidoPrazo_Grande

        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") "
        sSQL = sSQL & " UNION ALL "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & vCodOS & ") order by vTipo"
        
        'sSQL = sSQL & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & codPedido & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL, TIPO_DESC FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")

        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "RELATÓRIO DA ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas txtCodPedido.Text
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        'REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        
        If rOS("TIPO_DESC") = "R" Then
            REL_OS_Completo.rfDesc.Caption = FormatNumber(0, 2)
        Else
            REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        End If
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rPedido("DATA_COMPRA"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "VENDA À PRAZO"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If

        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo

    'Else
    '    REL_OS_PedidoPrazo.loadPedidos txtCodPedido.Text, "OFICINA"
    '    Unload REL_OS_PedidoPrazo
    'End If
Else
    'If r_Itens.RecordCount > 16 Then
        ''REL_OS_PedidoVista_Grande.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
        ''Unload REL_OS_PedidoVista_Grande

        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & vCodPedido & ") "
        sSQL = sSQL & " UNION ALL "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & vCodOS & ") order by vTipo"
        'Debug.Print sSQL
        'sSQL = sSQL & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & txtCodPedido.Text & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL, TIPO_DESC FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")
        
        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "RELATÓRIO DA ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas txtCodPedido.Text
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        'REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        If rOS("TIPO_DESC") = "R" Then
            REL_OS_Completo.rfDesc.Caption = FormatNumber(0, 2)
        Else
            REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        End If
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rPedido("DATA_COMPRA"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "VENDA À VISTA"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If

        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo

    'Else
    '    REL_OS_PedidoVista.loadPedidos txtCodPedido.Text, "OFICINA"
    '    Unload REL_OS_PedidoVista
    'End If
End If
Me.Show

'ANTERIOR
''ver a quantidade de peças e serviços da ordem de serviços
'If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
'    vTabelaServicos = "OS_Servicos_Auto"
'ElseIf vTipoOS = "Recapadora" Then
'    vTabelaServicos = "OS_Servicos_recapadora"
'ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
'    vTabelaServicos = "OS_Servicos_Auto"
'ElseIf vTipoOS = "Comunicação Visual" Then
'    vTabelaServicos = "OS_Servicos_Comunicacao"
'End If
    
'sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ")"
'sSQL_Itens = sSQL_Itens & " UNION "
'sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & txtCodOS.Text & ")"
'Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

''buscar a ordem de serviços para impressão
''sSQL = "SELECT * FROM os WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ");"
''Set r = dbData.OpenRecordset(sSQL)

'Me.Hide
'If cboTipoPgto.Text = "À PRAZO" And cboTipoOS.Text <> "ORÇAMENTO" Then
'    If r_Itens.RecordCount > 16 Then
'        REL_OS_PedidoPrazo_Grande.loadPedidos txtCodPedido.Text, "OFICINA"
'        Unload REL_OS_PedidoPrazo
'    Else
'        REL_OS_PedidoPrazo.loadPedidos txtCodPedido.Text, "OFICINA"
'        Unload REL_OS_PedidoPrazo
'    End If

'ElseIf cboTipoPgto.Text = "À VISTA" And cboTipoOS.Text <> "ORÇAMENTO" Then
'    If r_Itens.RecordCount > 16 Then
'        REL_OS_PedidoVista_Grande.loadPedidos txtCodPedido.Text, "OFICINA"
'        Unload REL_OS_PedidoVista_Grande
'    Else
'        REL_OS_PedidoVista.loadPedidos txtCodPedido.Text, "OFICINA"
'        Unload REL_OS_PedidoVista
'    End If
'ElseIf cboTipoOS.Text = "ORÇAMENTO" Then
'    If r_Itens.RecordCount > 16 Then
'        REL_Pedido_Orcamento.loadPedidos txtCodPedido.Text, "OFICINA"
'        Unload REL_Pedido_Orcamento
'    Else
'        REL_Pedido_Orcamento.loadPedidos txtCodPedido.Text, "OFICINA"
'        Unload REL_Pedido_Orcamento
'    End If
'End If
'Me.Show

''If cboTipoPgto.Text = "À PRAZO" And cboTipoOS.Text <> "ORÇAMENTO" Then
''   REL_OS_PedidoPrazo.loadPedidos txtCodPedido.Text, "OFICINA"
''   Unload REL_OS_PedidoPrazo
''ElseIf cboTipoPgto.Text = "À VISTA" And cboTipoOS.Text <> "ORÇAMENTO" Then
''   REL_OS_PedidoVista.loadPedidos txtCodPedido.Text, "OFICINA"
''   Unload REL_OS_PedidoVista
''ElseIf cboTipoOS.Text = "ORÇAMENTO" Then
''   REL_Pedido_Orcamento.loadPedidos txtCodPedido.Text, "OFICINA"
''   Unload REL_Pedido_Orcamento
''End If
End Sub

Private Sub cmdFinalizar_Click()
If txtTotalPecasServicos.Text = "" Then Exit Sub
If txtCodPedido.Text = "" Then Exit Sub
If cboTipoPgto.Text = "" Then Exit Sub
If txtCodCliente.Text = "" Then cboCliente.Text = "": cboCliente.SetFocus: Exit Sub
If txtFuncAP.Text = "" Then ShowMsg "Digite o código do funcionário!", vbInformation: txtCodFuncAP.SetFocus: Exit Sub
If txtCodPedido.Text = "" Then MsgBox "Cód. Pedido em Branco": Exit Sub
If txtCodOS.Text = "" Then Exit Sub Else vCodOS = txtCodOS.Text

If cboQuantForma.Text = "2 - FORMAS" Then
    If txtEntrada.Text = "" Or txtEntrada.Text = "0,00" Then
        ShowMsg "Você esqueceu de colocar um dos valores!", vbInformation: txtEntrada.SetFocus: Exit Sub
    End If
End If

cmdFinalizar.Enabled = False

Dim varValorRealDesc As Currency
Dim varValorRealAcresc As Currency
Dim NumCopias As Integer
Dim ii As Integer
Dim lNovoCod As Long
'Dim varHora As String       'Saber a hora do pagamento da parcela

'Usando na NFCe
Dim vCPF As String
'Dim sistNFe As snfe.Util

'verificar se o caixa ainda está aberto
sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (codcaixa = " & varCodCaixa & ") AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    If r("status") = 0 Then
        'MsgBox "aberto"
    Else
        Verificar_Caixa
        If CAIXA_FECHADO = True Then
            MsgBox "Não existe nenhum caixa aberto para essa venda!", vbInformation, "Aviso do Sistema"
            Exit Sub
        End If
    End If
Else
    Verificar_Caixa
    If CAIXA_FECHADO = True Then
        MsgBox "Não existe nenhum caixa aberto para essa venda!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
End If


'VERIFICAR SE EMITE NFCE =============================================
'NFCe_OK = False
'PararFechamentoVenda = False
    
'    If vConfImprimeNFCeLocal = "SIM" Then    'se essa maquina irá imprimir nfce localmente
'        'definir a impressora da nfce
'        Dim oIni As Ini
'        Set oIni = New Ini
'        oIni.Arquivo = appPathApp & "config.ini"
'        var_ImpNFCe = oIni.LerTexto("IMPRESSORA_NFCE", "impressora")
'        Set oIni = Nothing
        
'        Dim Prt As Printer
'        Dim oldPrinter As String
        
'        oldPrinter = Printer.DeviceName
        
'        For Each Prt In Printers
'           If Prt.DeviceName = var_ImpNFCe Then
'              Set Printer = Prt
'              Exit For
'           End If
'        Next
        
'        If cboTipoPgto.Text = "À VISTA" Then
'            If vNFCeConfImp = "SIM" Then
'                If MsgBox("Impressora Pronta?", vbQuestion + vbYesNo, "NFCe") = vbYes Then
'                    If PararFechamentoVenda = True Then
'                        NFCe_OK = False
'                        Exit Sub
'                    Else
'                        NFCe_OK = True
'                    End If
'                Else
'                    NFCe_OK = False
'                End If
'            Else
'                If PararFechamentoVenda = True Then
'                    NFCe_OK = False
'                    Exit Sub
'                Else
'                    NFCe_OK = True
'                End If
'            End If
'        Else
'            If vNFCeConfPrazo = "SIM" Then
'                If vNFCeConfImp = "SIM" Then
'                    If MsgBox("Impressora Pronta?", vbQuestion + vbYesNo, "NFCe") = vbYes Then
'                        If PararFechamentoVenda = True Then
'                            NFCe_OK = False
'                            Exit Sub
'                        Else
'                            NFCe_OK = True
'                        End If
'                    Else
'                        NFCe_OK = False
'                    End If
'                Else
'                    If PararFechamentoVenda = True Then
'                        NFCe_OK = False
'                        Exit Sub
'                    Else
'                        NFCe_OK = True
'                    End If
'                End If
'            Else
'                NFCe_OK = False
'            End If
'        End If
'    Else
'        NFCe_OK = False
'    End If

'TIPO DE CARTAO PARCELAS===========================================
Dim varTipoCartao As String
varTipoCartao = "NULL"
If cboformaPgto.Text = "3 - CARTÃO - DÉBITO" Then
   varTipoCartao = "'D'"
ElseIf cboformaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
   varTipoCartao = "'C'"
Else
    varTipoCartao = "NULL"
End If

'FORMA DE PAGAMENTO RESTANTE============================================
Dim var_PAGAMENTO As String
If cboformaPgto.Text = "1 - DINHEIRO" Then
   var_PAGAMENTO = "DINHEIRO"
ElseIf cboformaPgto.Text = "2 - PROMISSÓRIA" Then
   var_PAGAMENTO = "PROMISSORIA"
ElseIf cboformaPgto.Text = "3 - CARTÃO - DÉBITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboformaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboformaPgto.Text = "5 - CHEQUE" Then
   var_PAGAMENTO = "CHEQUE"
ElseIf cboformaPgto.Text = "6 - BOLETO" Then
   var_PAGAMENTO = "BOLETO"
ElseIf cboformaPgto.Text = "7 - TRANSFERÊNCIA" Then
   var_PAGAMENTO = "TRANSFERENCIA"
ElseIf cboformaPgto.Text = "8 - DEPOSITO" Then
   var_PAGAMENTO = "DEPOSITO"
ElseIf cboformaPgto.Text = "9 - FINANCEIRA" Then
   var_PAGAMENTO = "FINANCEIRA"
ElseIf cboformaPgto.Text = "10 - PIX" Then
   var_PAGAMENTO = "PIX"
End If

'se houver 2 opçoes de pagamento
If cboQuantForma.Text = "2 - FORMAS" Or cboQuantForma.Text = "2 - COM ENTRADA" Then
    'TIPO DE CARTAO ENTRADA============================================
    Dim varTipoCartaoEntrada As String
    varTipoCartaoEntrada = "NULL"
    
    If cboFormaPgtoEntrada.Text = "3 - CARTÃO - DÉBITO" Then
       varTipoCartaoEntrada = "'D'"
    ElseIf cboFormaPgtoEntrada.Text = "4 - CARTÃO - CRÉDITO" Then
       varTipoCartaoEntrada = "'C'"
    Else
        varTipoCartaoEntrada = "NULL"
    End If
    
    'FORMA DE PAGAMENTO ENTRADA ===============================================
    Dim var_PGTO_Entrada As String
    If cboFormaPgtoEntrada.Text = "1 - DINHEIRO" Then
       var_PGTO_Entrada = "DINHEIRO"
    ElseIf cboFormaPgtoEntrada.Text = "2 - PROMISSÓRIA" Then
       var_PGTO_Entrada = "PROMISSORIA"
    ElseIf cboFormaPgtoEntrada.Text = "3 - CARTÃO - DÉBITO" Then
       var_PGTO_Entrada = "CARTAO"
    ElseIf cboFormaPgtoEntrada.Text = "4 - CARTÃO - CRÉDITO" Then
       var_PGTO_Entrada = "CARTAO"
    ElseIf cboFormaPgtoEntrada.Text = "5 - CHEQUE" Then
       var_PGTO_Entrada = "CHEQUE"
    ElseIf cboFormaPgtoEntrada.Text = "6 - BOLETO" Then
       var_PGTO_Entrada = "BOLETO"
    ElseIf cboFormaPgtoEntrada.Text = "7 - TRANSFERÊNCIA" Then
       var_PGTO_Entrada = "TRANSFERENCIA"
    ElseIf cboFormaPgtoEntrada.Text = "8 - DEPOSITO" Then
       var_PGTO_Entrada = "DEPOSITO"
    ElseIf cboFormaPgtoEntrada.Text = "9 - FINANCEIRA" Then
       var_PGTO_Entrada = "FINANCEIRA"
    ElseIf cboFormaPgtoEntrada.Text = "10 - PIX" Then
       var_PGTO_Entrada = "PIX"
    Else
        var_PGTO_Entrada = "DINHEIRO"
    End If
End If

'quantidade de formas de pagamento
Dim varDivisaoPgto As String
varDivisaoPgto = cboQuantForma.Text

'calcular desconto em dinheiro ===================================
If optDescRS.Value = True Then
    If txtDesc.Text = "0,00" Then
        varValorRealDesc = Format(0, ocPESO)
    Else
        varValorRealDesc = Format(txtDesc.Text, ocPESO)
    End If
ElseIf optDescPorc.Value = True Then
    If txtDesc.Text = "0,00" Then
        varValorRealDesc = Format(0, ocPESO)
    Else
        varValorRealDesc = Format(((txtSubtotal.Text * txtDesc.Text) / 100), ocPESO)
    End If
End If

'calcular acrescimo em dinheiro==========================================================
If optAscrescRS.Value = True Then
    If txtAcresc.Text = "0,00" Then
        varValorRealAcresc = Format(0, ocMONEY)
    Else
        varValorRealAcresc = Format(CCur(txtAcresc.Text), ocMONEY)
    End If
ElseIf optAscrescPorc.Value = True Then
    If txtAcresc.Text = "0,00" Then
        varValorRealAcresc = Format(0, ocMONEY)
    Else
        varValorRealAcresc = Format(((CCur(txtSubtotal.Text) * CCur(txtAcresc.Text)) / 100), ocMONEY)
    End If
End If

'declarar quem receber os produtos
'If vDeclararRecebedor = "SIM" Then
'    Dim vRevebedor As String
'    If ShowMsg("Deseja declarar o recebedor?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'        vRevebedor = InputBox("Informe o nome do recebedor:", "ENTREGA DAS MERCADORIAS", "")
'    Else
'        vRevebedor = cboCliente.Text
'    End If
    
'    dbData.Execute "INSERT INTO pedidos_recebedor (cod_pedido, recebedor) VALUES (" & txtCodPedido.Text & ", '" & vRevebedor & "');"
'    vRevebedor = ""
'End If


If cboTipoPgto.Text = "À PRAZO" Then
    Dim var_Vencimento As Date
    Dim Var_NumParc As Integer
    Dim arrayParc() As Currency
    If txtCodCliente = "1" Then MsgBox "IDENTIFIQUE O CLIENTE DA COMPRA!", vbExclamation, "Aviso do sistema": Exit Sub
           
    'função para checar o limite do cliente..
    If txtCodCliente.Text <> "1" Then
         Verificar_Limite
    End If
        'If Passou_Limite = True Then Exit Sub
           
    'tabela configurações
    If bFechAP Then
       If ShowMsg("Deseja finalizar essa compra?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    End If
           
    'Solicita autorização da gerência
    'If txtCodCliente.Text <> "1" Then
    '     If Passou_Limite Or Cliente_Debito Then
    '        Dim fLib As LiberarVenda
    '        Dim bCancel As Boolean
    '        Dim lGerente As Long
    '
    '        Set fLib = New LiberarVenda
    '        Load fLib
    '
    '        fLib.Show vbModal
    '        bCancel = fLib.Cancelled
    '        lGerente = fLib.Gerente
                  
    '        Unload fLib
    '        Set fLib = Nothing
    '
    '        If bCancel Then
    '           cboCliente.Text = ""
    '           txtCodCliente.Text = ""
    '           Exit Sub
     '       End If
    '     End If
    ' End If
           
        ''If varLiberarVendaDevedor = True Then
        
    'colocar a data da Ultima compra de cada produro
    'For i = 1 To Grid.Rows - 1
    '   dbData.Execute "UPDATE produtos SET ult_compra = '" & Format$(Date, "yyyy-dd-MM") & "' WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
    'Next

        'ATUALIZAR A TABELA OS
        dbData.Execute "UPDATE os SET status_os = 1, tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", tipo_pagamento = 'À Prazo', pagamento = '" & var_PAGAMENTO & "', ValorDescReal = " & Replace(CCur(varValorRealDesc), ",", ".") & ", entrada = " & Replace(CCur(txtEntrada.Text), ",", ".") & " WHERE (cod_os = " & vCodOS & ");"

           
           'ATUALIZANDO A TABELA PEDIDOS
            sSQL = "UPDATE pedidos SET " & _
                 "cod_pedido = " & txtCodPedido.Text & ", " & _
                 "cod_cliente = " & txtCodCliente.Text & ", " & _
                 "data_compra = '" & Format$(Date, "yyyy-dd-MM") & "', " & _
                 "tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', " & _
                 "valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", " & _
                 "ValorDescReal = " & Replace(CCur(varValorRealDesc), ",", ".") & ", " & _
                 "ValorAcrescReal = " & Replace(CCur(varValorRealAcresc), ",", ".") & ", " & _
                 "TIPO_ACRESCIMO = '" & IIf(optAscrescRS.Value = True, "R", "P") & "', " & _
                 "VALOR_ACRESCIMO = " & Replace(CCur(txtAcresc.Text), ",", ".") & ", " & _
                 "entrada = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
                 "subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", " & _
                 "total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
                 "tipo_pagamento = 'À Prazo', pagamento = '" & var_PAGAMENTO & "', tipo_cartao = " & varTipoCartao & ", " & _
                 "cod_funcionario = " & txtCodFuncAP.Text & ", " & _
                 "tipo_pedido = 'OFICINA', " & _
                 "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', " & _
                 "MAQUINA = '" & IIf(StatusBar1.Panels(4).Text = "", "PDV01", StatusBar1.Panels(4).Text) & "', " & _
                 "codcaixa = " & varCodCaixa & ", " & _
                 "status_pedido = 1 " & _
                 "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
              dbData.Execute sSQL
           
           'COM ENTRADA =========================================================================
           If txtEntrada.Text <> "0,00" And txtValorParc.Text <> "0,00" Then
             
              'criar a entrada
              lNovoCod = Autonumeracao_Parcelas
              
              dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, cod_os,  numero, data, valor, status, TIPO, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, VALOR_FINAL) VALUES (" & _
                 lNovoCod & ", " & txtCodPedido.Text & ", " & vCodOS & ", 1, '" & Format$(Date, "yyyy-dd-MM") & "', " & _
                 Replace(CCur(txtEntrada.Text), ",", ".") & ", 0, 'OS', 0, 0, 0, 0, " & Replace(CCur(txtEntrada.Text), ",", ".") & ");"
              
              'criar da segunda parcela em diante
              var_Vencimento = Format(DateAdd("m", Val(1), mskInicio.Text), "dd/mm/yy")
              Var_NumParc = 2
              
              CalcularParcelas (CCur(txtTotalDesc) - CCur(txtEntrada)), CInt(cboQuantParc), arrayParc
              
              For i = 1 To CInt(cboQuantParc)
                 lNovoCod = Autonumeracao_Parcelas
                 
                 dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor, status, VALOR_FINAL, TIPO, DIAS_ATRAZO, JUROS, MULTA, DESCONTO) VALUES (" & _
                    lNovoCod & ", " & txtCodPedido.Text & ", " & vCodOS & ", " & Var_NumParc & ", '" & Format$(var_Vencimento, "yyyy-dd-MM") & "', " & _
                    Replace(arrayParc(i), ",", ".") & ", 0, " & Replace(arrayParc(i), ",", ".") & ", 'OS', 0, 0, 0, 0);"
                 
                 var_Vencimento = Format(DateAdd("m", Val(1), var_Vencimento), "dd/mm/yy")
                 Var_NumParc = Var_NumParc + 1
              Next
              
           'SEM ENTRADA =========================================================================
           ElseIf txtEntrada.Text = "0,00" And txtValorParc.Text <> "0,00" Then
              
              'parcelas
              var_Vencimento = CDate(mskInicio.Text)
              Var_NumParc = 1
              
              CalcularParcelas CCur(txtTotalDesc), CInt(cboQuantParc), arrayParc
              
              'criar as parcelas
              For i = 1 To CInt(cboQuantParc)
                 lNovoCod = Autonumeracao_Parcelas
                 
                 dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor, status, VALOR_FINAL, TIPO, DIAS_ATRAZO, JUROS, MULTA, DESCONTO) VALUES (" & _
                    lNovoCod & ", " & txtCodPedido.Text & ", " & vCodOS & ", " & Var_NumParc & ", '" & Format$(var_Vencimento, "yyyy-dd-MM") & "', " & _
                    Replace(arrayParc(i), ",", ".") & ", 0, " & Replace(arrayParc(i), ",", ".") & ", 'OS', 0, 0, 0, 0);"
                 
                 var_Vencimento = Format(DateAdd("m", Val(1), var_Vencimento), "dd/mm/yy")
                 Var_NumParc = Var_NumParc + 1
              Next
              
           End If
           
           'dar baixa na parcela de entrada ou compra à vista
           'If lblEstornar.Caption = "ESTORNO" Then
           '     If txtEntrada.Text <> "0,00" Then
           '        dbData.Execute "UPDATE parcelas SET " & _
           '           "status = 1, valor_final = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
           '           "pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
           '           "hora = '" & Format(txtHoraCompra, ocHORA) & "', " & _
           '           "forma_pgto = '" & var_PGTO_Entrada & "', " & _
           '           "tipo = 'PARCELA', tipo_cartao = " & varTipoCartaoEntrada & ", " & _
           '           "CODCAIXA = " & varCodCaixa & ", " & _
           '           "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "' " & _
           '           "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
           '     End If
           '
           '     'dar baixa nas parcelas de de cartão
           '     If cboformaPgto.Text = "3 - CARTÃO - DÉBITO" Or cboformaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
           '        dbData.Execute "update parcelas set pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', Status = 1, valor_final = VALOR, hora = '" & Format(txtHoraCompra, ocHORA) & "', forma_pgto = 'CARTAO', caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', CODCAIXA = " & varCodCaixa & "  WHERE (cod_pedido = " & txtCodPedido.Text & ")"
           '     End If
           'Else
                If txtEntrada.Text <> "0,00" Then
                   dbData.Execute "UPDATE parcelas SET " & _
                      "status = 1, valor_final = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
                      "pagamento = '" & Format$(Date, "yyyy-dd-MM") & "', " & _
                      "hora = '" & Format(Now, ocHORA) & "', " & _
                      "forma_pgto = '" & var_PGTO_Entrada & "', " & _
                      "tipo = 'PARCELA', tipo_cartao = " & varTipoCartaoEntrada & ", " & _
                      "CODCAIXA = " & varCodCaixa & ", " & _
                      "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "' " & _
                      "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
                End If
                
                'dar baixa nas parcelas de de cartão
                If cboformaPgto.Text = "3 - CARTÃO - DÉBITO" Or cboformaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
                   dbData.Execute "update parcelas set pagamento = '" & Format$(Date, "yyyy-dd-MM") & "', Status = 1, valor_final = VALOR, hora = '" & Format(Now, ocHORA) & "', forma_pgto = 'CARTAO', caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', CODCAIXA = " & varCodCaixa & " WHERE (cod_pedido = " & txtCodPedido.Text & ")"
                End If
           'End If
           
           'txtHoraCompra.Text = ""
           
           'Colocando a data da ultima compra
           'execSQL "UPDATE CLIENTE SET Ultima_Compra = #" & Format(Date, "MM/dd/yyyy") & "# WHERE CODIGO = " & txtCodCliente.Text

'DESATIVEI PARA ANALISAR DEPOIS COMO DAR DESCONTO, POIS O DESCONTO É PARA PEÇA E SERVIÇOS
        'calcular subtotal de cada item
        'sSQL = "UPDATE pedidos_itens SET subtotal = preco * quantidade where (cod_pedido = " & txtCodPedido.Text & ")"
        'dbData.Execute sSQL
    
        'calcular desconto de cada item
        'If optDescRS.Value = True Then
        '    sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(txtDescItens.Text), ",", ".") & " / 100) where (cod_pedido = " & txtCodPedido.Text & ")"
        'Else
        '    sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(txtDesc.Text), ",", ".") & " / 100) where (cod_pedido = " & txtCodPedido.Text & ")"
        'End If
        'dbData.Execute sSQL
        
        'sSQL = "UPDATE pedidos_itens SET total = subtotal - desconto, data = '" & Format$(Date, "yyyy-dd-MM") & "' where (cod_pedido = " & txtCodPedido.Text & ")"
        'dbData.Execute sSQL
'DESATIVEI ATÉ AQUI
        
        'Retirar da tabela PRODUTOS as QUANTIDADES mencionadas no grid
        For i = 1 To Grid_Servicos.Rows - 1  'analizar essa linha
            If Grid_Servicos.TextMatrix(i, 2) = "PRODUTO" Then
                dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid_Servicos.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Servicos.TextMatrix(i, 10) & ");"
            End If
        Next
        
        If iCopiasAP <> 0 Then  'saber a quantidade de copias
           If bEntregaAP = True Then
              If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                 NumCopias = iCopiasAP + 1
              Else
                 NumCopias = iCopiasAP
              End If
           Else
              NumCopias = iCopiasAP
           End If
        Else
           NumCopias = "1"
        End If
        
        If bImprAP = True Then       'Confirma se vai ter impressão
           If bConfImprAP = True Then
              If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                 For ii = 1 To NumCopias
                    Imprimir_Pedido
                 Next
              End If
           Else
              For ii = 1 To NumCopias
                 Imprimir_Pedido
              Next
           End If
        End If
        
ElseIf cboTipoPgto.Text = "À VISTA" Then
           If bFechAV Then
              If ShowMsg("Deseja finalizar essa compra?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
           End If
           
           'colocar a data da Ultima compra de cada produro
           'For i = 1 To Grid.Rows - 1
           '   dbData.Execute "UPDATE produtos SET ult_compra = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "' WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
           'Next

            'ATUALIZAR A TABELA OS
            dbData.Execute "UPDATE os SET status_os = 1, tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", tipo_pagamento = 'À Vista', pagamento = '" & varDivisaoPgto & "', ValorDescReal = " & Replace(CCur(varValorRealDesc), ",", ".") & ", entrada = 0 WHERE (cod_os = " & vCodOS & ");"

          
           'ATUALIZANDO A TABELA PEDIDOS
           sSQL = "UPDATE pedidos SET " & _
              "cod_pedido = " & txtCodPedido.Text & ", " & _
              "cod_cliente = " & txtCodCliente.Text & ", " & _
              "data_compra = '" & Format$(Date, "yyyy-dd-MM") & "', " & _
              "tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', " & _
              "valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", " & _
              "ValorDescReal = " & Replace(CCur(varValorRealDesc), ",", ".") & ", " & _
              "ValorAcrescReal = " & Replace(CCur(varValorRealAcresc), ",", ".") & ", " & _
              "TIPO_ACRESCIMO = '" & IIf(optAscrescRS.Value = True, "R", "P") & "', " & _
              "VALOR_ACRESCIMO = " & Replace(CCur(txtAcresc.Text), ",", ".") & ", " & _
              "entrada = 0, " & _
              "subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", " & _
              "total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
              "tipo_pagamento = 'À Vista', pagamento = '" & varDivisaoPgto & "', tipo_cartao = " & varTipoCartao & ", " & _
              "cod_funcionario = " & txtCodFuncAP.Text & ",  " & _
              "tipo_pedido = 'OFICINA', caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', maquina = '" & IIf(StatusBar1.Panels(4).Text = "", "PDV01", StatusBar1.Panels(4).Text) & "', codcaixa = " & varCodCaixa & ", " & _
              "status_pedido = 1" & _
              "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
           dbData.Execute sSQL
           
           '===========================================CRIAR E DAR BAIXA EM PARCELAS ==========================================
            If cboTipoPgto.Text = "À VISTA" And cboQuantForma.Text = "1 - FORMA" Then
                'autonumeração das parcelas
                lNovoCod = Autonumeracao_Parcelas
                
                'Criando as Parcelas
                sSQL = "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor, VALOR_FINAL, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, TIPO) VALUES (" & _
                   lNovoCod & ", " & txtCodPedido.Text & ", " & vCodOS & ", 1, '" & Format$(Date, "yyyy-dd-MM") & "', " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", 0, 0, 0, 0, 'OS');"
                dbData.Execute sSQL
                
                'DAR BAIXA NA PARCELA =====
                sSQL = "UPDATE parcelas SET " & _
                "status = 1, " & _
                "valor_final = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
                "pagamento = '" & Format$(Date, "yyyy-dd-MM") & "', " & _
                "hora = '" & Format(Now, ocHORA) & "', " & _
                "forma_pgto = '" & var_PAGAMENTO & "', " & _
                "tipo = 'OS', " & _
                "tipo_cartao = " & varTipoCartao & ", " & _
                "CODCAIXA = " & varCodCaixa & ", " & _
                "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "' " & _
                "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
                   
                dbData.Execute sSQL
                'txtHoraCompra.Text = ""
            
            ElseIf cboTipoPgto.Text = "À VISTA" And cboQuantForma.Text = "2 - FORMAS" Then
                'autonumeração das parcelas
                lNovoCod = Autonumeracao_Parcelas
                
                'Criando as Parcelas
                sSQL = "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor, VALOR_FINAL, DIAS_ATRAZO, JUROS, MULTA, DESCONTO) VALUES (" & _
                   lNovoCod & ", " & txtCodPedido.Text & ", " & vCodOS & ", 1, '" & Format$(Date, "yyyy-dd-MM") & "', " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & Replace(CCur(txtEntrada.Text), ",", ".") & ", 0, 0, 0, 0);"
                dbData.Execute sSQL
            
                'autonumeração das parcelas
                lNovoCod = Autonumeracao_Parcelas
                
                'Criando as Parcelas
                sSQL = "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor, VALOR_FINAL, DIAS_ATRAZO, JUROS, MULTA, DESCONTO) VALUES (" & _
                   lNovoCod & ", " & txtCodPedido.Text & ", " & vCodOS & ", 2, '" & Format$(Date, "yyyy-dd-MM") & "', " & Replace(CCur(txtValorRest.Text), ",", ".") & ", " & Replace(CCur(txtValorRest.Text), ",", ".") & ", 0, 0, 0, 0);"
                dbData.Execute sSQL
            
                'DAR BAIXA NA PARCELA =====
                'compra com estorno pega a data e hora do estorno
                'If lblEstornar.Caption = "ESTORNO" Then
                '    varHora = Format(txtHoraCompra, ocHORA)
                'Else
                '    varHora = Format(Now, ocHORA)
                'End If
                
                'parcela 1
                sSQL = "UPDATE parcelas SET " & _
                "status = 1, " & _
                "valor_final = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
                "pagamento = '" & Format$(Date, "yyyy-dd-MM") & "', " & _
                "hora = '" & Format(Now, ocHORA) & "', " & _
                "forma_pgto = '" & var_PGTO_Entrada & "', " & _
                "tipo = 'OS', " & _
                "tipo_cartao = " & varTipoCartaoEntrada & ", " & _
                "CODCAIXA = " & varCodCaixa & ", " & _
                "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "' " & _
                "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
                dbData.Execute sSQL
            
                'parcela 2
                sSQL = "UPDATE parcelas SET " & _
                "status = 1, " & _
                "valor_final = " & Replace(CCur(txtValorRest.Text), ",", ".") & ", " & _
                "pagamento = '" & Format$(Date, "yyyy-dd-MM") & "', " & _
                "hora = '" & Format(Now, ocHORA) & "', " & _
                "forma_pgto = '" & var_PAGAMENTO & "', " & _
                "tipo = 'OS', " & _
                "tipo_cartao = " & varTipoCartao & ", " & _
                "CODCAIXA = " & varCodCaixa & ", " & _
                "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "' " & _
                "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 2);"
                dbData.Execute sSQL
                
                'txtHoraCompra.Text = ""
            End If
           
           'autonumeração das parcelas
           'lNovoCod = Autonumeracao_Parcelas
        
           'Criando as Parcelas
           'sSQL = "INSERT INTO parcelas (codigo, cod_pedido, cod_os, numero, data, valor) VALUES (" & _
           '   lNovoCod & ", " & txtCodPedido.Text & ", 1, '" & Format$(date, "yyyy-dd-MM") & "', " & Replace(CCur(txtTotalPecasServicos.Text), ",", ".") & ");"
           'dbData.Execute sSQL
           
           'Colocando a data da ultima compra
           ''execSQL "UPDATE CLIENTE SET Ultima_Compra = #" & Format(Date, "mm/dd/yyyy") & "# WHERE CODIGO = " & txtCodCliente.Text
           
        
           'dar baixa na parcela de entrada ou compra à vista
           'If lblEstornar.Caption = "ESTORNO" Then
           '   sSQL = "UPDATE parcelas SET " & _
            '  "status = 1, " & _
           '   "valor_final = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
           '   "pagamento = '" & Format$(date, "yyyy-dd-MM") & "', " & _
           '   "hora = '" & Format(txtHoraCompra, ocHORA) & "', " & _
           '   "forma_pgto = '" & var_PAGAMENTO & "', " & _
           '   "tipo = 'OS', " & _
           '   "tipo_cartao = " & varTipoCartao & ", " & _
           '   "CODCAIXA = " & varCodCaixa & ", " & _
           '   "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "' " & _
           '   "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
           'Else
           '   sSQL = "UPDATE parcelas SET " & _
           '   "status = 1, " & _
           '   "valor_final = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
           '   "pagamento = '" & Format$(date, "yyyy-dd-MM") & "', " & _
           '   "hora = '" & Format(Now, ocHORA) & "', " & _
           '   "forma_pgto = '" & var_PAGAMENTO & "', " & _
           '   "tipo = 'OS', tipo_cartao = " & varTipoCartao & ", " & _
           '   "CODCAIXA = " & varCodCaixa & ", " & _
           '   "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "' " & _
           '   "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
           'End If
           'dbData.Execute sSQL
           'txtHoraCompra.Text = ""

        'calcular itens do pedido
        'sSQL = "UPDATE pedidos_itens SET subtotal = preco * quantidade where (cod_pedido = " & txtCodPedido.Text & ")"
        'dbData.Execute sSQL
        
'DESATIVEI PARA ANALISAR DEPOIS COMO DAR DESCONTO, POIS O DESCONTO É PARA PEÇA E SERVIÇOS
        'If optDescRS.Value = True Then
        '    sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(txtDescItens.Text), ",", ".") & " / 100) where (cod_pedido = " & txtCodPedido.Text & ")"
        'Else
        '    sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(txtDesc.Text), ",", ".") & " / 100) where (cod_pedido = " & txtCodPedido.Text & ")"
        'End If
        'dbData.Execute sSQL
        
        'sSQL = "UPDATE pedidos_itens SET total = subtotal - desconto, data = '" & Format$(Date, "yyyy-dd-MM") & "' where (cod_pedido = " & txtCodPedido.Text & ")"
        'dbData.Execute sSQL
'DESATIVEI ATÉ AQUI
        
        'Retirar da tabela PRODUTOS as QUANTIDADES mencionadas no grid
        For i = 1 To Grid_Servicos.Rows - 1  'analizar essa linha
            If Grid_Servicos.TextMatrix(i, 2) = "PRODUTO" Then
                dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid_Servicos.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Servicos.TextMatrix(i, 10) & ");"
            End If
        Next
        
        If iCopiasAP <> 0 Then  'saber a quantidade de copias
           If bEntregaAP = True Then
              If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                 NumCopias = iCopiasAP + 1
              Else
                 NumCopias = iCopiasAP
              End If
           Else
              NumCopias = iCopiasAP
           End If
        Else
           NumCopias = "1"
        End If
        
        If bImprAP = True Then       'Confirma se vai ter impressão
           If bConfImprAP = True Then
              If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                 For ii = 1 To NumCopias
                    Imprimir_Pedido
                 Next
              End If
           Else
              For ii = 1 To NumCopias
                 Imprimir_Pedido
              Next
           End If
        End If
End If

LimparObjetos_Entrada
LimparObjetos_Prazo
LimparGrid_Situacao
LimparTotais
txtCodOS.Text = ""
frmVendaFechamento.Visible = False
MostrarGrid_OS
MostrarGrid_OS_Situacao
SSTab1.Tab = 0

cmdFinalizar.Enabled = True

'If CAIXA_FECHADO = True Then
'Else
'    lblTipoPedido.Caption = ""
'    cmdFinalizarAvista.Enabled = True
'    cmdFinalizarPrazo.Enabled = True
'    cmdOrçamento.Enabled = True
'    cmdCancelarPedido.Enabled = True
'    cmdRemover.Enabled = True
'    cmdAvancado.Enabled = True
'    cmdInfProduto.Enabled = True
'    Grid.Enabled = True
'    txtCodBarra.Enabled = True
'    txtValor.Enabled = True
'    txtQuant.Enabled = True
'    txtTotal.Enabled = True
'End If

End Sub

Private Sub Calcular_Parcelas2()
'If txtTotalDesc.Text = "0,00" Or txtValorRest.Text = "0,00" Or cboQuantParc.Text = "" Then Exit Sub

Dim var_ValorRest As Currency
Dim QUANT As Integer
Dim RESULTADO As Currency

var_ValorRest = txtValorRest.Text
If cboQuantParc.Text = "0" Then cboQuantParc.Text = "1"
QUANT = cboQuantParc.Text

RESULTADO = CCur(var_ValorRest / QUANT)
txtValorParc = Format(RESULTADO, ocMONEY)
End Sub
Private Sub cmdFinalizarAP_Click()
If cboTipoOS.Text <> "CONSERTO" Then MsgBox "Somente é possivel geral financeiro para uma OS de conserto!", vbInformation, "Aviso do Sistema": SSTab1.Tab = 1: cboTipoOS.SetFocus: Exit Sub
If txtTotalPecasServicos.Text = "" Or txtTotalPecasServicos.Text = "0,00" Then Exit Sub
Dim varTipoPgto As String
Dim varTipoCartao As String

If IsDate(lblDataAberturaCaixa.Caption) = False Then
    MsgBox "Caixa Fechado"
    Exit Sub
End If

If CDate(lblDataAberturaCaixa.Caption) <> Date Then
    If MsgBox("A data do caixa aberto é diferente da data atual. Continuar mesmo assim?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
        
        cboTipoPgto.Text = "À PRAZO"
        frmVendaFechamento.Visible = True
        LimparObjetos_Prazo
        txtSubtotal.Text = txtSubtotalGeral.Text
        txtAcresc.Text = FormatNumber(0, 2)
        
        'If lblEstornar.Caption = "ESTORNO" Then
        '    sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
        '    Set r = dbData.OpenRecordset(sSQL)
        '
        '    'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
        '    'txtCodCliente.Text = ""
        '
        '    If Not r.EOF Then
        '
        '        If r("TIPO_PEDIDO") = "ORÇAMENTO" Then
        '            mskInicio.Text = Format(Date, "dd/mm/yy")
        '            mskTermino.Text = Format(Date, "dd/mm/yy")
        '            txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
        '        Else
        '            txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
        '            'txtHoraCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
        '            mskInicio.Text = Format(r("data_compra"), "dd/mm/yy")
        '            mskTermino.Text = Format(r("data_compra"), "dd/mm/yy")
        '        End If
                
        '        txtCodFuncAP.Text = ValidateNull(r("cod_funcionario"))
        '        txtCodCliente.Text = ValidateNull(r("COD_CLIENTE"))
        '        'lblTipoPedido.Caption = ValidateNull(r("tipo_pedido"))
        '        varTipoPgto = ValidateNull(r("pagamento"))
        '        varTipoCartao = ValidateNull(r("TIPO_CARTAO"))
                
        '        If r("TIPO_DESC") = "P" Then
        '            optDescPorc.Value = True
        '        Else
        '            optDescRS.Value = True
        '        End If
                
        '        txtDesc.Text = FormatNumber(r("VALOR_DESC"), 3)
        
        '         If varTipoPgto = "DINHEIRO" Then
        '             cboformaPgto.Text = "1 - DINHEIRO"
        '         ElseIf varTipoPgto = "PROMISSORIA" Then
        '             cboformaPgto.Text = "2 - PROMISSÓRIA"
        '         ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "D" Then
        '             cboformaPgto.Text = "3 - CARTÃO - DÉBITO"
        '         ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "C" Then
        '             cboformaPgto.Text = "4 - CARTÃO - CRÉDITO"
        '         ElseIf varTipoPgto = "CHEQUE" Then
        '             cboformaPgto.Text = "5 - CHEQUE"
        '         ElseIf varTipoPgto = "BOLETO" Then
        '             cboformaPgto.Text = "6 - BOLETO"
        '         ElseIf varTipoPgto = "FINANCEIRA" Then
        '             cboformaPgto.Text = "9 - FINANCEIRA"
        '         End If
                 
        '        cboformaPgto.Text = "1 - DINHEIRO"
        '        cboQuantForma.Text = "1 - FORMA"
                
        '        txtRecebido.SetFocus
        '    End If
        '        Calcular_Desconto
        '        'Calcular_Prazo
        'Else
            
            cboformaPgto.Text = "2 - PROMISSÓRIA"
            cboQuantForma.Text = "1 - SEM ENTRADA"
            optDescPorc.Value = False
            optDescPorc.Value = True
            cboQuantForma_LostFocus
            
            'limpar campo funcionario
            'If varLoginFunc <> "" Then
               'If varLoginFunc = "2" Then
                  'If lblEstornar.Caption <> "ESTORNO" Then
                     'txtCodFuncAP.Text = ""
                     'txtFuncAP.Text = ""
                     'txtCodFuncAP.SetFocus
                  'Else
                     'cboCliente.SetFocus
                 ' End If
               'Else
                  'cboCliente.SetFocus
               'End If
            'End If
            
            'If lblEstornar.Caption = "ESTORNO" Then
            
                mskInicio.Text = Format(Date, "dd/mm/yy")
                mskTermino.Text = Format(Date, "dd/mm/yy")
                'optDescPorc.Value = True
                'cboCliente.Text = ""
                'BuscarClienteConsumidor
                Mostrar_Desconto
                Calcular_Desconto
                'Calcular_Prazo
                'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
                'If varLoginFunc = "2" Then txtFuncAP.Text = ""
                'If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else txtRecebido.SetFocus
            
            'HabilitaObjetosVenda True
        'End If
        
        cmdFinalizarAV.Enabled = False
        cmdFinalizarAP.Enabled = False
        'cmdOrçamento.Enabled = False
        'cmdCancelarPedido.Enabled = False
        'cmdRemover.Enabled = False
        'cmdAvancado.Enabled = False
        'cmdInfProduto.Enabled = False
        'Grid.Enabled = False
        'txtCodBarra.Enabled = False
        'txtValor.Enabled = False
        'txtQuant.Enabled = False
        'txtTotal.Enabled = False
        If txtDescGeral > 0 Then
            txtDesc.Text = txtDescGeral.Text
            txtDesc.Enabled = False
            optDescRS.Value = True
            optDescRS.Enabled = False
            optDescPorc.Enabled = False
            txtCodFuncAP.SetFocus
        Else
            txtDesc.Text = "0"
            txtDesc.Enabled = True
            optDescRS.Value = True
            optDescRS.Enabled = True
            optDescPorc.Enabled = True
            txtCodFuncAP.SetFocus
        End If
    End If
Else
        cboTipoPgto.Text = "À PRAZO"
        frmVendaFechamento.Visible = True
        LimparObjetos_Prazo
        txtSubtotal.Text = txtSubtotalGeral.Text
        optDescRS.Value = True
        txtDesc.Text = txtDescGeral.Text
        txtAcresc.Text = FormatNumber(0, 2)
        
        'If lblEstornar.Caption = "ESTORNO" Then
        '    sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
         '   Set r = dbData.OpenRecordset(sSQL)
            
        '    'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
        '    'txtCodCliente.Text = ""
            
        '    If Not r.EOF Then
        '        If r("TIPO_PEDIDO") = "ORÇAMENTO" Then
        '            mskInicio.Text = Format(Date, "dd/mm/yy")
        '            mskTermino.Text = Format(Date, "dd/mm/yy")
        '            txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
        '        Else
        '            txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
        '            'txtHoraCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
        '            mskInicio.Text = Format(r("data_compra"), "dd/mm/yy")
        '            mskTermino.Text = Format(r("data_compra"), "dd/mm/yy")
        '        End If
                
        '        txtCodFuncAP.Text = ValidateNull(r("cod_funcionario"))
        '        txtCodCliente.Text = ValidateNull(r("COD_CLIENTE"))
        '        'lblTipoPedido.Caption = ValidateNull(r("tipo_pedido"))
        '        varTipoPgto = ValidateNull(r("pagamento"))
        '        varTipoCartao = ValidateNull(r("TIPO_CARTAO"))
                
        '        If r("TIPO_DESC") = "P" Then
        '            optDescPorc.Value = True
        '        Else
        '            optDescRS.Value = True
        '        End If
                
        '        txtDesc.Text = FormatNumber(r("VALOR_DESC"), 3)
        
        '         If varTipoPgto = "DINHEIRO" Then
        '             cboformaPgto.Text = "1 - DINHEIRO"
        '         ElseIf varTipoPgto = "PROMISSORIA" Then
        '             cboformaPgto.Text = "2 - PROMISSÓRIA"
        '         ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "D" Then
        '             cboformaPgto.Text = "3 - CARTÃO - DÉBITO"
        '         ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "C" Then
        '             cboformaPgto.Text = "4 - CARTÃO - CRÉDITO"
        '         ElseIf varTipoPgto = "CHEQUE" Then
        '             cboformaPgto.Text = "5 - CHEQUE"
        '         ElseIf varTipoPgto = "BOLETO" Then
        '             cboformaPgto.Text = "6 - BOLETO"
        '         ElseIf varTipoPgto = "FINANCEIRA" Then
        '             cboformaPgto.Text = "9 - FINANCEIRA"
        '         End If
                 
        '        cboformaPgto.Text = "1 - DINHEIRO"
        '        cboQuantForma.Text = "1 - FORMA"
        '
        '        txtRecebido.SetFocus
        '    End If
        '        Calcular_Desconto
        '        'Calcular_Prazo
        'Else
            
            cboformaPgto.Text = "2 - PROMISSÓRIA"
            cboQuantForma.Text = "1 - SEM ENTRADA"
            optDescPorc.Value = False
            optDescPorc.Value = True
            cboQuantForma_LostFocus
            
            'limpar campo funcionario
            'If varLoginFunc <> "" Then
               'If varLoginFunc = "2" Then
                  'If lblEstornar.Caption <> "ESTORNO" Then
                  '   txtCodFuncAP.Text = ""
                  '   txtFuncAP.Text = ""
                  '   txtCodFuncAP.SetFocus
                  'Else
                  '   'cboCliente.SetFocus
                  'End If
               'Else
                  'cboCliente.SetFocus
               'End If
            'End If
            
            'If lblEstornar.Caption = "ESTORNO" Then
            
                mskInicio.Text = Format(Date, "dd/mm/yy")
                mskTermino.Text = Format(Date, "dd/mm/yy")
                'optDescPorc.Value = True
                'cboCliente.Text = ""
                'BuscarClienteConsumidor
                Mostrar_Desconto
                Calcular_Desconto
                'Calcular_Prazo
                'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
                'If varLoginFunc = "2" Then txtFuncAP.Text = ""
                'If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else txtRecebido.SetFocus
            
            'HabilitaObjetosVenda True
        'End If
        
        cmdFinalizarAV.Enabled = False
        cmdFinalizarAP.Enabled = False
        'cmdOrçamento.Enabled = False
        'cmdCancelarPedido.Enabled = False
        'cmdRemover.Enabled = False
        'cmdAvancado.Enabled = False
        'cmdInfProduto.Enabled = False
        'Grid.Enabled = False
        'txtCodBarra.Enabled = False
        'txtValor.Enabled = False
        'txtQuant.Enabled = False
        'txtTotal.Enabled = False
        If txtDescGeral > 0 Then
            txtDesc.Text = txtDescGeral.Text
            txtDesc.Enabled = False
            optDescRS.Value = True
            optDescRS.Enabled = False
            optDescPorc.Enabled = False
            txtCodFuncAP.SetFocus
        Else
            txtDesc.Text = "0"
            txtDesc.Enabled = True
            optDescRS.Value = True
            optDescRS.Enabled = True
            optDescPorc.Enabled = True
            txtCodFuncAP.SetFocus
        End If
End If

'Calcular_Desconto
Calcular_Parcelas
Calcular_Prazo
cmdFinalizar.Enabled = True
End Sub

Private Sub cmdFinalizarAV_Click()
If cboTipoOS.Text <> "CONSERTO" Then MsgBox "Somente é possivel geral financeiro para uma OS de conserto!", vbInformation, "Aviso do Sistema": SSTab1.Tab = 1: cboTipoOS.SetFocus: Exit Sub

Dim varTipoPgto As String
Dim varTipoCartao As String

If IsDate(lblDataAberturaCaixa.Caption) = False Then
    MsgBox "Caixa Fechado"
    Exit Sub
End If

If CDate(lblDataAberturaCaixa.Caption) <> Date Then
    If MsgBox("A data do caixa aberto é diferente da data atual. Continuar mesmo assim?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
        If txtTotalPecasServicos.Text = "" Or txtTotalPecasServicos.Text = "0,00" Then Exit Sub
        cboTipoPgto.Text = "À VISTA"
        frmVendaFechamento.Visible = True
        LimparObjetos_Prazo
        txtSubtotal.Text = txtSubtotalGeral.Text
        txtAcresc.Text = FormatNumber(0, 2)
        
           
            cboformaPgto.Text = "1 - DINHEIRO"
            cboQuantForma.Text = "1 - FORMA"
            optDescPorc.Value = False
            optDescPorc.Value = True
            cboQuantForma_LostFocus
            
            'limpar campo funcionario
            'If varLoginFunc <> "" Then
               'If varLoginFunc = "2" Then
                  'If lblEstornar.Caption <> "ESTORNO" Then
                     'txtCodFuncAP.Text = ""
                     'txtFuncAP.Text = ""
                     'txtCodFuncAP.SetFocus
                  'Else
                     'cboCliente.SetFocus
                 ' End If
               'Else
                  'cboCliente.SetFocus
               'End If
            'End If
            
            'If lblEstornar.Caption = "ESTORNO" Then
            
                mskInicio.Text = Format(Date, "dd/mm/yy")
                mskTermino.Text = Format(Date, "dd/mm/yy")
                'optDescPorc.Value = True
                'cboCliente.Text = ""
                'BuscarClienteConsumidor
                Mostrar_Desconto
                Calcular_Desconto
                'Calcular_Prazo
                'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
                'If varLoginFunc = "2" Then txtFuncAP.Text = ""
                'If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else txtRecebido.SetFocus
            
            'HabilitaObjetosVenda True
        'End If
        
        cmdFinalizarAV.Enabled = False
        cmdFinalizarAP.Enabled = False
        'cmdOrçamento.Enabled = False
        'cmdCancelarPedido.Enabled = False
        'cmdRemover.Enabled = False
        'cmdAvancado.Enabled = False
        'cmdInfProduto.Enabled = False
        'Grid.Enabled = False
        'txtCodBarra.Enabled = False
        'txtValor.Enabled = False
        'txtQuant.Enabled = False
        'txtTotal.Enabled = False
        If txtDescGeral > 0 Then
            txtDesc.Text = txtDescGeral.Text
            txtDesc.Enabled = False
            optDescRS.Value = True
            optDescRS.Enabled = False
            optDescPorc.Enabled = False
            txtRecebido.SetFocus
        Else
            txtDesc.Text = "0"
            txtDesc.Enabled = True
            optDescRS.Value = True
            optDescRS.Enabled = True
            optDescPorc.Enabled = True
            txtRecebido.SetFocus
        End If
    End If
Else
        If txtTotalPecasServicos.Text = "" Or txtTotalPecasServicos.Text = "0,00" Then Exit Sub
        cboTipoPgto.Text = "À VISTA"
        frmVendaFechamento.Visible = True
        LimparObjetos_Prazo
        txtSubtotal.Text = txtSubtotalGeral.Text
        txtAcresc.Text = FormatNumber(0, 2)
        
        'If lblEstornar.Caption = "ESTORNO" Then
        '    sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
         '   Set r = dbData.OpenRecordset(sSQL)
            
        '    'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
        '    'txtCodCliente.Text = ""
            
        '    If Not r.EOF Then
        '        If r("TIPO_PEDIDO") = "ORÇAMENTO" Then
        '            mskInicio.Text = Format(Date, "dd/mm/yy")
        '            mskTermino.Text = Format(Date, "dd/mm/yy")
        '            txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
        '        Else
        '            txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
        '            'txtHoraCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
        '            mskInicio.Text = Format(r("data_compra"), "dd/mm/yy")
        '            mskTermino.Text = Format(r("data_compra"), "dd/mm/yy")
        '        End If
                
        '        txtCodFuncAP.Text = ValidateNull(r("cod_funcionario"))
        '        txtCodCliente.Text = ValidateNull(r("COD_CLIENTE"))
        '        'lblTipoPedido.Caption = ValidateNull(r("tipo_pedido"))
        '        varTipoPgto = ValidateNull(r("pagamento"))
        '        varTipoCartao = ValidateNull(r("TIPO_CARTAO"))
                
        '        If r("TIPO_DESC") = "P" Then
        '            optDescPorc.Value = True
        '        Else
        '            optDescRS.Value = True
        '        End If
                
        '        txtDesc.Text = FormatNumber(r("VALOR_DESC"), 3)
        
        '         If varTipoPgto = "DINHEIRO" Then
        '             cboformaPgto.Text = "1 - DINHEIRO"
        '         ElseIf varTipoPgto = "PROMISSORIA" Then
        '             cboformaPgto.Text = "2 - PROMISSÓRIA"
        '         ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "D" Then
        '             cboformaPgto.Text = "3 - CARTÃO - DÉBITO"
        '         ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "C" Then
        '             cboformaPgto.Text = "4 - CARTÃO - CRÉDITO"
        '         ElseIf varTipoPgto = "CHEQUE" Then
        '             cboformaPgto.Text = "5 - CHEQUE"
        '         ElseIf varTipoPgto = "BOLETO" Then
        '             cboformaPgto.Text = "6 - BOLETO"
        '         ElseIf varTipoPgto = "FINANCEIRA" Then
        '             cboformaPgto.Text = "9 - FINANCEIRA"
        '         End If
                 
        '        cboformaPgto.Text = "1 - DINHEIRO"
        '        cboQuantForma.Text = "1 - FORMA"
        '
        '        txtRecebido.SetFocus
        '    End If
        '        Calcular_Desconto
        '        'Calcular_Prazo
        'Else
            
            cboformaPgto.Text = "1 - DINHEIRO"
            cboQuantForma.Text = "1 - FORMA"
            optDescPorc.Value = False
            optDescPorc.Value = True
            cboQuantForma_LostFocus
            
            'limpar campo funcionario
            'If varLoginFunc <> "" Then
               'If varLoginFunc = "2" Then
                  'If lblEstornar.Caption <> "ESTORNO" Then
                  '   txtCodFuncAP.Text = ""
                  '   txtFuncAP.Text = ""
                  '   txtCodFuncAP.SetFocus
                  'Else
                  '   'cboCliente.SetFocus
                  'End If
               'Else
                  'cboCliente.SetFocus
               'End If
            'End If
            
            'If lblEstornar.Caption = "ESTORNO" Then
            
                mskInicio.Text = Format(Date, "dd/mm/yy")
                mskTermino.Text = Format(Date, "dd/mm/yy")
                'optDescPorc.Value = True
                'cboCliente.Text = ""
                'BuscarClienteConsumidor
                Mostrar_Desconto
                Calcular_Desconto
                'Calcular_Prazo
                'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
                'If varLoginFunc = "2" Then txtFuncAP.Text = ""
                'If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else txtRecebido.SetFocus
            
            'HabilitaObjetosVenda True
        'End If
        
        cmdFinalizarAV.Enabled = False
        cmdFinalizarAP.Enabled = False
        'cmdOrçamento.Enabled = False
        'cmdCancelarPedido.Enabled = False
        'cmdRemover.Enabled = False
        'cmdAvancado.Enabled = False
        'cmdInfProduto.Enabled = False
        'Grid.Enabled = False
        'txtCodBarra.Enabled = False
        'txtValor.Enabled = False
        'txtQuant.Enabled = False
        'txtTotal.Enabled = False
        If txtDescGeral > 0 Then
            txtDesc.Text = txtDescGeral.Text
            txtDesc.Enabled = False
            optDescRS.Value = True
            optDescRS.Enabled = False
            optDescPorc.Enabled = False
        Else
            txtDesc.Text = "0"
            txtDesc.Enabled = True
            optDescRS.Value = True
            optDescRS.Enabled = True
            optDescPorc.Enabled = True
        End If
End If
Calcular_Desconto
Calcular_Parcelas
Calcular_Prazo
cmdFinalizar.Enabled = True
End Sub

Private Sub cmdFinanceiroOS_Click()
Dim posit As Long
posit = Grid_OS.Row
txtCodOS.Text = ""
txtCodOS.Text = (Grid_OS.TextMatrix(Grid_OS.Row, 0))
SSTab1.Tab = 2
If (Trim(Grid_OS.TextMatrix(posit, 2))) = ("TERMINADO") Then
    If (Trim(Grid_OS.TextMatrix(posit, 3))) = ("ABERTO") Then
        cmdFinalizarAP.Enabled = True
        cmdFinalizarAV.Enabled = True
    End If
End If
End Sub

Private Sub cmdGerarEntrada_Click()
'On Error GoTo TrataErro

If txtCodFuncionario.Text = "" Then
   ShowMsg "Faltou escolher o recepcionista!", vbInformation
   cboFuncionario.SetFocus
   Exit Sub
End If

'If Not IsDate(mskHoraSaida.Text) = True Then
'   ShowMsg "Falta a hora de previsão de saída!", vbInformation
'   mskHoraSaida.SetFocus
'   Exit Sub
'End If
If Not IsDate(mskHoraSaida.Text) = True Then mskHoraSaida.Text = "00:00"
'If txtKM.Text = "" Then ShowMsg "Quilometragem não especificada", vbInformation, "Aviso do Sistema": txtKM.SetFocus: Exit Sub

If txtCodCliente.Text = "" Then
   ShowMsg "Cliente não encontra-se cadastrado no sistema", vbInformation
   cboCliente.SetFocus
   Exit Sub
End If

If txtCodFuncionario.Text = "" Then
   ShowMsg "Funcionário não encontra-se cadastrado no sistema", vbInformation
   cboFuncionario.SetFocus
   Exit Sub
End If

'OS
If Not Atualizar_Dados_OS Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

'alterar dados do pedido
dbData.Execute "UPDATE pedidos SET cod_cliente = " & txtCodCliente.Text & ", cod_funcionario = " & txtCodFuncionario.Text & ", data_entrega = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(5).Text, ocDATA) & "', 103), data_compra = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(5).Text, ocDATA) & "', 103) WHERE (cod_pedido = " & txtCodPedido.Text & ");"

'Equipamento
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    dbData.Execute "UPDATE OS_Equipamento_Auto SET fabricante = '" & cboFabricante.Text & "', modelo = '" & cboModelo.Text & "', placa = '" & txtPlaca.Text & "', ano = '" & txtAno.Text & "', km = '" & txtKM.Text & "',  CHASSI = '" & txtChassi.Text & "', COR = '" & cboCor.Text & "', TANQUE = '" & cboTanque.Text & "', PARECER_CLIENTE = '" & txtPareceCliente.Text & "' WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Recapadora" Then
    dbData.Execute "UPDATE OS_Equipamento_Auto SET fabricante = '" & cboFabricante.Text & "', modelo = '" & cboModelo.Text & "', placa = '" & txtPlaca.Text & "', ano = '" & txtAno.Text & "', km = '" & txtKM.Text & "',  CHASSI = '" & txtChassi.Text & "', COR = '" & cboCor.Text & "', TANQUE = '" & cboTanque.Text & "', PARECER_CLIENTE = '" & txtPareceCliente.Text & "' WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "UPDATE OS_Equipamento SET fabricante = '" & cboFabricante.Text & "', modelo = '" & cboModelo.Text & "', equipamento = '" & cboTanque.Text & "', PARECER_CLIENTE = '" & txtPareceCliente.Text & "' WHERE (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Comunicação Visual" Then
    'dbData.Execute "UPDATE OS_Equipamento SET fabricante = '" & cboFabricante.Text & "', modelo = '" & cboModelo.Text & "', equipamento = '" & cboTanque.Text & "', PARECER_CLIENTE = '" & txtPareceCliente.Text & "' WHERE (cod_os = " & txtCodOS.Text & ");"
End If

cmdGerarEntrada.Enabled = False
cmdCancelarEntrada.Enabled = False
cmdNovo.Enabled = True
''MostrarGrid_OS

LimparObjetos_Entrada
LimparObjetos_Servicos
LimparObjetos_Pecas
txtCodOS.Text = ""
txtCodPedido.Text = ""
MostrarGrid_OS
MostrarGrid_OS_Situacao
Form_Load
   
'TrataErro:
   'If Err.Number = 3022 Then
   '   MsgBox "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation, "Aviso do Sistema"
   '   Exit Sub
   'End If
End Sub

Private Sub cmdImpEntrada1_Click()
frmSecundario.Enabled = True
cmdGerarEntrada.Enabled = False
cmdCancelarEntrada.Enabled = False
cmdAlterar.Enabled = False
cmdApagar.Enabled = False
cmdNovo.Enabled = True

txtCodOS.Text = ""
txtCodOS.Text = (Grid_OS.TextMatrix(Grid_OS.Row, 0))

menu_Impressao_Entrada_Click
End Sub

Private Sub cmdImpEntrada2_Click()
menu_Impressao_Entrada_Click
End Sub

Private Sub cmdImpOrcamento1_Click()
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

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

'buscando os dados do formulário
i = Grid_OS.Row
vCodOS = Grid_OS.TextMatrix(i, 0)
codPedido = Grid_OS.TextMatrix(i, 7)

'ver a quantidade de peças e serviços da ordem de serviços
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Recapadora" Then
    vTabelaServicos = "OS_Servicos_recapadora"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Comunicação Visual" Then
    vTabelaServicos = "OS_Servicos_Comunicacao"
End If

'somando os produtos
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM pedidos_itens WHERE (cod_pedido = " & Grid_OS.TextMatrix(i, 7) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
Dim vQuantProduto As Double
Dim vTotalProduto As Currency
vQuantProduto = ValidateNull(r_Itens("VarQuant"))
vTotalProduto = ValidateNull(r_Itens("VarSoma"))
'somando os serviços
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
'Debug.Print sSQL_Itens
Dim vQuantServico As Double
Dim vTotalServico As Currency
vQuantServico = ValidateNull(r_Itens("VarQuant"))
vTotalServico = ValidateNull(r_Itens("VarSoma"))
Dim vSomaTotais As Currency
Dim vSomaQuant As Double
vSomaTotais = vTotalProduto + vTotalServico
vSomaQuant = vQuantProduto + vQuantServico

'TOTAIS RESUMIDO
REL_OS_Completo.txtQuantServicos.Caption = " " & Format(vQuantServico, "000")
REL_OS_Completo.txtQuantPecas.Caption = " " & Format(vQuantProduto, "000")
REL_OS_Completo.txtQuantGeral.Caption = " " & Format(vSomaQuant, "000")

REL_OS_Completo.txtTotalServicos.Caption = " " & FormatNumber(vTotalServico, 2)
REL_OS_Completo.txtTotalPecas.Caption = " " & FormatNumber(vTotalProduto, 2)
REL_OS_Completo.txtTotalPecasServicos.Caption = " " & FormatNumber(vSomaTotais, 2)

    
sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Grid_OS.TextMatrix(i, 7) & ")"
sSQL_Itens = sSQL_Itens & " UNION "
sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

varImpPDF = False
Me.Hide

'If r_Itens.RecordCount > 16 Then
        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") "
        sSQL = sSQL & " UNION "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ") order by vTipo"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & codPedido & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL, TIPO_DESC, DATA_ENTRADA, DATA_TERMINO, OBS FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")
        
        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "ORÇAMENTO DA ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas Grid_OS.TextMatrix(i, 7)
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)

        If rOS("TIPO_DESC") = "R" Then
            REL_OS_Completo.rfDesc.Caption = FormatNumber(0, 2)
        Else
            REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        End If
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rOS("DATA_ENTRADA"), "dd/mm/yy")
        REL_OS_Completo.rfDataSaida.Caption = Format(rOS("DATA_TERMINO"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "ORÇAMENTO"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
        REL_OS_Completo.txtParecerTitulo.Visible = True
        REL_OS_Completo.txtParecer.Visible = True
        REL_OS_Completo.txtParecer.Caption = rOS("OBS")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR, CHASSI FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If
        
        'DADOS DO VEICULO/EQUIPAMENTO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
            REL_OS_Completo.txtChassi.Caption = IIf(IsNull(rEquip!CHASSI) = True, "", rEquip!CHASSI)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo
'Else
'    With REL_Pedido_Orcamento
'        REL_Pedido_Orcamento.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
'    End With
'    Unload REL_Pedido_Orcamento
'End If

Me.Show
End Sub

Private Sub cmdImpOrcamento2_Click()
menu_Impressao_Orcamento_Click
End Sub


Private Sub cmdImpPedido1_Click()
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

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

'buscando os dados do formulário
i = Grid_OS.Row

vCodOS = Grid_OS.TextMatrix(i, 0)
codPedido = Grid_OS.TextMatrix(i, 7)

If Grid_OS.TextMatrix(i, 6) = "00000" Then MsgBox "Pedido gerado anterior as alterações não permite reimpressão de pedidos. Somente orçamento!", vbInformation, "Aviso do Sistema": Exit Sub

'ver a quantidade de peças e serviços da ordem de serviços
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Recapadora" Then
    vTabelaServicos = "OS_Servicos_recapadora"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Comunicação Visual" Then
    vTabelaServicos = "OS_Servicos_Comunicacao"
End If

'somando os produtos
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM pedidos_itens WHERE (cod_pedido = " & Grid_OS.TextMatrix(i, 7) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
Dim vQuantProduto As Double
Dim vTotalProduto As Currency
vQuantProduto = ValidateNull(r_Itens("VarQuant"))
vTotalProduto = ValidateNull(r_Itens("VarSoma"))
'somando os serviços
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
'Debug.Print sSQL_Itens
Dim vQuantServico As Double
Dim vTotalServico As Currency
vQuantServico = ValidateNull(r_Itens("VarQuant"))
vTotalServico = ValidateNull(r_Itens("VarSoma"))
Dim vSomaTotais As Currency
Dim vSomaQuant As Double
vSomaTotais = vTotalProduto + vTotalServico
vSomaQuant = vQuantProduto + vQuantServico

'TOTAIS RESUMIDO
REL_OS_Completo.txtQuantServicos.Caption = " " & Format(vQuantServico, "000")
REL_OS_Completo.txtQuantPecas.Caption = " " & Format(vQuantProduto, "000")
REL_OS_Completo.txtQuantGeral.Caption = " " & Format(vSomaQuant, "000")

REL_OS_Completo.txtTotalServicos.Caption = " " & FormatNumber(vTotalServico, 2)
REL_OS_Completo.txtTotalPecas.Caption = " " & FormatNumber(vTotalProduto, 2)
REL_OS_Completo.txtTotalPecasServicos.Caption = " " & FormatNumber(vSomaTotais, 2)
    
sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Grid_OS.TextMatrix(i, 7) & ")"
sSQL_Itens = sSQL_Itens & " UNION ALL "
sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

'Debug.Print sSQL_Itens

'buscar a ordem de serviços para impressão
sSQL = "SELECT * FROM os WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ");"
Set r = dbData.OpenRecordset(sSQL)

varImpPDF = False
Me.Hide
If r("TIPO_PAGAMENTO") = "À Prazo" Then
    'If r_Itens.RecordCount > 16 Then
        ''REL_OS_PedidoPrazo_Grande.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
        ''Unload REL_OS_PedidoPrazo_Grande

        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") "
        sSQL = sSQL & " UNION ALL "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ") order by vTipo"
        
        ''sSQL = sSQL & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & codPedido & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL, TIPO_DESC FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")
        
        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas Grid_OS.TextMatrix(i, 6)
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        
        If rOS("TIPO_DESC") = "R" Then
            REL_OS_Completo.rfDesc.Caption = FormatNumber(0, 2)
        Else
            REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        End If
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rPedido("DATA_COMPRA"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "VENDA À PRAZO"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If

        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo

    'Else
    '    REL_OS_PedidoPrazo.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
    '    Unload REL_OS_PedidoPrazo
    'End If
Else
    'If r_Itens.RecordCount > 16 Then
        ''REL_OS_PedidoVista_Grande.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
        ''Unload REL_OS_PedidoVista_Grande

        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") "
        sSQL = sSQL & " UNION ALL "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ") order by vTipo"
        
        ''sSQL = sSQL & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & codPedido & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL, TIPO_DESC FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")
        
        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "RELATÓRIO DA ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas Grid_OS.TextMatrix(i, 7)
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        'REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        
        If rOS("TIPO_DESC") = "R" Then
            REL_OS_Completo.rfDesc.Caption = FormatNumber(0, 2)
        Else
            REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        End If
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rPedido("DATA_COMPRA"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "VENDA À VISTA"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If

        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo

    'Else
    '    REL_OS_PedidoVista.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
    '    Unload REL_OS_PedidoVista
    'End If
End If
Me.Show
End Sub

Private Sub cmdImpPedido2_Click()
'menu_Impressao_Pedido_Click
If txtCodOS.Text = "" Then Exit Sub
If cboStatus.Text <> "TERMINADO" Then MsgBox "Somente é permitido a impressão de PEDIDOS de Ordem de Serviço terminado!", vbInformation, "Aviso dos Sistema": Exit Sub
'vCodOS = txtCodOS.Text
If txtCodPedido = "0" Then MsgBox "Pedido gerado anterior as alterações não permite reimpressão de pedidos. Somente orçamento!", vbInformation, "Aviso do Sistema": Exit Sub

Imprimir_Pedido

'sSQL = "SELECT * FROM os WHERE (cod_os = " & txtCodOS.Text & ");"
'Set r = dbData.OpenRecordset(sSQL)

'If r("TIPO_PAGAMENTO") = "À Prazo" Then
'    With REL_OS_PedidoPrazo
'        .txtQuantServicos.Caption = " " & Format(txtQuantServicos.Text, "000")
'        .txtQuantPecas.Caption = " " & Format(txtQuantPecas.Text, "000")
'        .txtQuantGeral.Caption = " " & Format(txtQuantGeral.Text, "000")
'
'        .txtTotalServicos.Caption = " " & FormatNumber(txtTotalServicos.Text, 2)
'        .txtTotalPecas.Caption = " " & FormatNumber(txtTotalPecas.Text, 2)
'        .txtTotalPecasServicos.Caption = " " & FormatNumber(txtTotalPecasServicos.Text, 2)
'    End With
'
'    REL_OS_PedidoPrazo.loadPedidos vCodOS, "OFICINA"
'    Unload REL_OS_PedidoPrazo
'Else
'    With REL_OS_PedidoVista
'        .txtQuantServicos.Caption = " " & Format(txtQuantServicos.Text, "000")
'        .txtQuantPecas.Caption = " " & Format(txtQuantPecas.Text, "000")
'        .txtQuantGeral.Caption = " " & Format(txtQuantGeral.Text, "000")
'
'        .txtTotalServicos.Caption = " " & FormatNumber(txtTotalServicos.Text, 2)
'        .txtTotalPecas.Caption = " " & FormatNumber(txtTotalPecas.Text, 2)
'        .txtTotalPecasServicos.Caption = " " & FormatNumber(txtTotalPecasServicos.Text, 2)
'    End With
'
'    REL_OS_PedidoVista.loadPedidos vCodOS, "OFICINA"
'    Unload REL_OS_PedidoVista
'End If
End Sub

Private Sub cmdImprimirConsulta_Click()
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

REL_OS_Consulta.dfQuant.Caption = lblQuant.Caption
REL_OS_Consulta.dfTotal.Caption = "TOTAL: " & lblTotalConsulta.Caption
REL_OS_Consulta.lblTitulo.Caption = "RELATÓRIO - CONSULTA DE ORDEM DE SERVIÇOS"

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

REL_OS_Consulta.Relatorio.NomeImpressora = var_Impressora
REL_OS_Consulta.Relatorio.Ativar
Unload REL_OS_Consulta

Me.Show 1
End Sub

Private Sub cmdNovo_Click()
LimparObjetos_Entrada
LimparTotais
LimparObjetos_Servicos
LimparObjetos_Pecas
AutoNumeracao_Pedido
AutoNumeracao_OS
mskDataEntrada.Text = Format(Date, "dd/mm/yy")
mskDataSaida.Text = Format(Date, "dd/mm/yy")
mskHoraEntrada.Text = Format(Time, "hh:mm")
cboStatus_GotFocus
cboStatus.ListIndex = 0

dbData.Execute "INSERT INTO pedidos (cod_pedido, status_pedido, tipo_pedido, orcamento, reaberto, cancelado, TIPO_DESC, ENTRADA) VALUES (" & txtCodPedido.Text & ", 0, 'OFICINA', 0, 0, 0, 'R', 0)"
dbData.Execute "INSERT INTO os (cod_os, cod_pedido, TIPO_DESC, ENTRADA) VALUES (" & txtCodOS.Text & ", " & txtCodPedido.Text & ", 'R', 0)"

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    dbData.Execute "INSERT INTO OS_Equipamento_Auto (cod_os) VALUES (" & txtCodOS.Text & ")"
ElseIf vTipoOS = "Recapadora" Then
    dbData.Execute "INSERT INTO OS_Equipamento_Auto (cod_os) VALUES (" & txtCodOS.Text & ")"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "INSERT INTO OS_Equipamento (cod_os) VALUES (" & txtCodOS.Text & ")"
ElseIf vTipoOS = "Comunicação Visual" Then
    'dbData.Execute "INSERT INTO OS_Equipamento (cod_os) VALUES (" & txtCodOS.Text & ")"
End If

cmdGerarEntrada.Enabled = True
cmdCancelarEntrada.Enabled = True
cboStatus.Enabled = False
cmdNovo.Enabled = False
cmdAlterar.Enabled = False
cmdApagar.Enabled = False

If cboStatus.Text = "À COMEÇAR" Then
    frmAcessorios.Visible = True
    frmSituacao.Visible = True
    frmParecerCliente.Visible = True
    
    If vTipoOS = "Comunicação Visual" Then
        frmTotaisGeral.Visible = True
        frmTotaisProdServ.Visible = True
        frmGridServicos.Visible = True
    Else
        frmTotaisGeral.Visible = False
        frmTotaisProdServ.Visible = False
        frmGridServicos.Visible = False
    End If
    frmSecundario.Enabled = True
    stProdSer.Visible = False
    frmEquipamento.Visible = True
    cboTipoOS.Text = "CONSERTO"
Else
    frmAcessorios.Visible = False
    frmSituacao.Visible = False
    frmParecerCliente.Visible = False
    frmGridServicos.Visible = True
    stProdSer.Visible = True
    frmTotaisGeral.Visible = True
    frmTotaisProdServ.Visible = True
    frmSecundario.Enabled = True
    frmEquipamento.Visible = True
    cboTipoOS.Text = "CONSERTO"
End If

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    frmAcessorios.Visible = True
    frmSituacao.Visible = True
ElseIf vTipoOS = "Recapadora" Then
    frmAcessorios.Visible = False
    frmSituacao.Visible = False
    frmEquipamento.Visible = False
    txtObs.Visible = True
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    frmAcessorios.Visible = True
    txtObs.Visible = False
    frmSituacao.Visible = True
ElseIf vTipoOS = "Comunicação Visual" Then
    frmAcessorios.Visible = False
    frmSituacao.Visible = False
    frmParecerCliente.Visible = False
    frmGridServicos.Visible = True
    stProdSer.Visible = True
    frmTotaisGeral.Visible = True
    frmTotaisProdServ.Visible = True
    'frmSecundario.Enabled = False
    frmEquipamento.Visible = False
    cboTipoOS.Text = "CONFECÇÃO"
    txtObs.Visible = True
End If

cmdImpEntrada2.Enabled = False
cmdImpOrcamento2.Enabled = False
cmdImpPedido2.Enabled = False

cboFuncionario.SetFocus
End Sub

Private Sub cmdNovoOS_Click()
If txtCodOS.Text <> "" And cmdGerarEntrada.Enabled = True Then
    MsgBox "A Ordem de Serviço iniciada ainda não foi salvo", vbInformation, "Aviso do Sistema"
    SSTab1.Tab = 1
    Exit Sub
End If

SSTab1.Tab = 1
cmdNovo_Click
End Sub



Private Sub cmdOrcamentoPDF_Click()
varImpPDF = True
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

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

'buscando os dados do formulário
i = Grid_OS.Row
vCodOS = Grid_OS.TextMatrix(i, 0)
codPedido = Grid_OS.TextMatrix(i, 7)

'ver a quantidade de peças e serviços da ordem de serviços
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Recapadora" Then
    vTabelaServicos = "OS_Servicos_recapadora"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Comunicação Visual" Then
    vTabelaServicos = "OS_Servicos_Comunicacao"
End If

'somando os produtos
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM pedidos_itens WHERE (cod_pedido = " & Grid_OS.TextMatrix(i, 7) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
Dim vQuantProduto As Double
Dim vTotalProduto As Currency
vQuantProduto = ValidateNull(r_Itens("VarQuant"))
vTotalProduto = ValidateNull(r_Itens("VarSoma"))
'somando os serviços
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
'Debug.Print sSQL_Itens
Dim vQuantServico As Double
Dim vTotalServico As Currency
vQuantServico = ValidateNull(r_Itens("VarQuant"))
vTotalServico = ValidateNull(r_Itens("VarSoma"))
Dim vSomaTotais As Currency
Dim vSomaQuant As Double
vSomaTotais = vTotalProduto + vTotalServico
vSomaQuant = vQuantProduto + vQuantServico

'TOTAIS RESUMIDO
REL_OS_Completo.txtQuantServicos.Caption = " " & Format(vQuantServico, "000")
REL_OS_Completo.txtQuantPecas.Caption = " " & Format(vQuantProduto, "000")
REL_OS_Completo.txtQuantGeral.Caption = " " & Format(vSomaQuant, "000")

REL_OS_Completo.txtTotalServicos.Caption = " " & FormatNumber(vTotalServico, 2)
REL_OS_Completo.txtTotalPecas.Caption = " " & FormatNumber(vTotalProduto, 2)
REL_OS_Completo.txtTotalPecasServicos.Caption = " " & FormatNumber(vSomaTotais, 2)

    
sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Grid_OS.TextMatrix(i, 7) & ")"
sSQL_Itens = sSQL_Itens & " UNION "
sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

'varImpPDF = False
Me.Hide

'If r_Itens.RecordCount > 16 Then
        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") "
        sSQL = sSQL & " UNION "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ") order by vTipo"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & codPedido & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL, DATA_ENTRADA, DATA_TERMINO, OBS FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")

        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "ORÇAMENTO DA ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas Grid_OS.TextMatrix(i, 7)
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rOS("DATA_ENTRADA"), "dd/mm/yy")
        REL_OS_Completo.rfDataSaida.Caption = Format(rOS("DATA_TERMINO"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "ORÇAMENTO"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
            REL_OS_Completo.txtParecerTitulo.Visible = True
        REL_OS_Completo.txtParecer.Visible = True
        REL_OS_Completo.txtParecer.Caption = rOS("OBS")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR, CHASSI FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If
        
        'DADOS DO VEICULO/EQUIPAMENTO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
            REL_OS_Completo.txtChassi.Caption = IIf(IsNull(rEquip!CHASSI) = True, "", rEquip!CHASSI)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Visualizar = False
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo
'Else
'    With REL_Pedido_Orcamento
'        REL_Pedido_Orcamento.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
'    End With
'    Unload REL_Pedido_Orcamento
'End If

Me.Show

End Sub

Private Sub cmdPedidoPDF_Click()
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
varImpPDF = True

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

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

'buscando os dados do formulário
i = Grid_OS.Row

vCodOS = Grid_OS.TextMatrix(i, 0)
codPedido = Grid_OS.TextMatrix(i, 7)

If Grid_OS.TextMatrix(i, 6) = "00000" Then MsgBox "Pedido gerado anterior as alterações não permite reimpressão de pedidos. Somente orçamento!", vbInformation, "Aviso do Sistema": Exit Sub

'ver a quantidade de peças e serviços da ordem de serviços
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Recapadora" Then
    vTabelaServicos = "OS_Servicos_recapadora"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Comunicação Visual" Then
    vTabelaServicos = "OS_Servicos_Comunicacao"
End If

'somando os produtos
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM pedidos_itens WHERE (cod_pedido = " & Grid_OS.TextMatrix(i, 7) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
Dim vQuantProduto As Double
Dim vTotalProduto As Currency
vQuantProduto = r_Itens("VarQuant")
vTotalProduto = r_Itens("VarSoma")
'somando os serviços
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
'Debug.Print sSQL_Itens
Dim vQuantServico As Double
Dim vTotalServico As Currency
vQuantServico = r_Itens("VarQuant")
vTotalServico = r_Itens("VarSoma")
Dim vSomaTotais As Currency
Dim vSomaQuant As Double
vSomaTotais = vTotalProduto + vTotalServico
vSomaQuant = vQuantProduto + vQuantServico

'TOTAIS RESUMIDO
REL_OS_Completo.txtQuantServicos.Caption = " " & Format(vQuantServico, "000")
REL_OS_Completo.txtQuantPecas.Caption = " " & Format(vQuantProduto, "000")
REL_OS_Completo.txtQuantGeral.Caption = " " & Format(vSomaQuant, "000")

REL_OS_Completo.txtTotalServicos.Caption = " " & FormatNumber(vTotalServico, 2)
REL_OS_Completo.txtTotalPecas.Caption = " " & FormatNumber(vTotalProduto, 2)
REL_OS_Completo.txtTotalPecasServicos.Caption = " " & FormatNumber(vSomaTotais, 2)
    
sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Grid_OS.TextMatrix(i, 7) & ")"
sSQL_Itens = sSQL_Itens & " UNION ALL "
sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

'Debug.Print sSQL_Itens

'buscar a ordem de serviços para impressão
sSQL = "SELECT * FROM os WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ");"
Set r = dbData.OpenRecordset(sSQL)

varImpPDF = False
Me.Hide
If r("TIPO_PAGAMENTO") = "À Prazo" Then
    'If r_Itens.RecordCount > 16 Then
        ''REL_OS_PedidoPrazo_Grande.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
        ''Unload REL_OS_PedidoPrazo_Grande

        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") "
        sSQL = sSQL & " UNION ALL "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ") order by vTipo"
        
        ''sSQL = sSQL & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & codPedido & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")

        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "RELATÓRIO DA ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas Grid_OS.TextMatrix(i, 7)
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rPedido("DATA_COMPRA"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "VENDA À PRAZO"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If

        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Visualizar = False
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo

    'Else
    '    REL_OS_PedidoPrazo.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
    '    Unload REL_OS_PedidoPrazo
    'End If
Else
    'If r_Itens.RecordCount > 16 Then
        ''REL_OS_PedidoVista_Grande.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
        ''Unload REL_OS_PedidoVista_Grande

        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") "
        sSQL = sSQL & " UNION ALL "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ") order by vTipo"
        
        ''sSQL = sSQL & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & codPedido & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")

        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "RELATÓRIO DA ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas Grid_OS.TextMatrix(i, 7)
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rPedido("DATA_COMPRA"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "VENDA À VISTA"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If

        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Visualizar = False
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo

    'Else
        'REL_OS_PedidoVista.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
        'Unload REL_OS_PedidoVista
    'End If
End If
Me.Show

'i = Grid_OS.Row

'vCodOS = Grid_OS.TextMatrix(i, 0)
'
'If Grid_OS.TextMatrix(i, 6) = "00000" Then MsgBox "Pedido gerado anterior as alterações não permite reimpressão de pedidos. Somente orçamento!", vbInformation, "Aviso do Sistema": Exit Sub

'ver a quantidade de peças e serviços da ordem de serviços
'If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
'    vTabelaServicos = "OS_Servicos_Auto"
'ElseIf vTipoOS = "Recapadora" Then
'    vTabelaServicos = "OS_Servicos_recapadora"
'ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
'    vTabelaServicos = "OS_Servicos_Auto"
'ElseIf vTipoOS = "Comunicação Visual" Then
'    vTabelaServicos = "OS_Servicos_Comunicacao"
'End If

'somando os produtos
'sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM pedidos_itens WHERE (cod_pedido = " & Grid_OS.TextMatrix(i, 6) & ")"
'Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
'Dim vQuantProduto As Double
'Dim vTotalProduto As Currency
'vQuantProduto = r_Itens("VarQuant")
'vTotalProduto = r_Itens("VarSoma")
''somando os serviços
'sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
'Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
''Debug.Print sSQL_Itens
'Dim vQuantServico As Double
'Dim vTotalServico As Currency
'vQuantServico = r_Itens("VarQuant")
'vTotalServico = r_Itens("VarSoma")
'Dim vSomaTotais As Currency
'Dim vSomaQuant As Double
'vSomaTotais = vTotalProduto + vTotalServico
'vSomaQuant = vQuantProduto + vQuantServico

'TOTAIS RESUMIDO
'REL_OS_Completo.txtQuantServicos.Caption = " " & Format(vQuantServico, "000")
'REL_OS_Completo.txtQuantPecas.Caption = " " & Format(vQuantProduto, "000")
'REL_OS_Completo.txtQuantGeral.Caption = " " & Format(vSomaQuant, "000")

'REL_OS_Completo.txtTotalServicos.Caption = " " & FormatNumber(vTotalServico, 2)
'REL_OS_Completo.txtTotalPecas.Caption = " " & FormatNumber(vTotalProduto, 2)
'REL_OS_Completo.txtTotalPecasServicos.Caption = " " & FormatNumber(vSomaTotais, 2)
    
'sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Grid_OS.TextMatrix(i, 6) & ")"
'sSQL_Itens = sSQL_Itens & " UNION "
'sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ")"
'Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

'buscar a ordem de serviços para impressão
'sSQL = "SELECT * FROM os WHERE (cod_os = " & Grid_OS.TextMatrix(i, 0) & ");"
'Set r = dbData.OpenRecordset(sSQL)

'varImpPDF = True
'Me.Hide
'If r("TIPO_PAGAMENTO") = "À Prazo" Then
'    If r_Itens.RecordCount > 16 Then
'        REL_OS_PedidoPrazo_Grande.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
'        Unload REL_OS_PedidoPrazo_Grande
'    Else
'        REL_OS_PedidoPrazo.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
'        Unload REL_OS_PedidoPrazo
'    End If
'Else
'    If r_Itens.RecordCount > 16 Then
'        REL_OS_PedidoVista_Grande.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
'        Unload REL_OS_PedidoVista_Grande
'    Else
'        REL_OS_PedidoVista.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
'        Unload REL_OS_PedidoVista
'    End If
'End If
'Me.Show
End Sub


Private Sub cmdRemoverAcessorios_Click()
   On Error GoTo erro
   
   If Grid_Acessorio.TextMatrix(Grid_Acessorio.Row, 1) = "" Then GoSub erro
   If ShowMsg("Deseja excluir o acessório: " & Grid_Acessorio.TextMatrix(Grid_Acessorio.Row, 4) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (codigo = " & Grid_Acessorio.TextMatrix(Grid_Acessorio.Row, 1) & ") AND (cod_os = " & txtCodOS.Text & ");"
   
   MostrarGrid_Acessorios
   Exit Sub
   
erro:
   ShowMsg "Não existe nenhum acessório para ser excluido!", vbExclamation
   Exit Sub
End Sub

Private Sub LimparGrid_Acessorios()
   Dim i As Integer
   
   With Grid_Acessorio
      .Visible = False
      
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 0
      .ColWidth(4) = 2900
      
      .RowHeight(0) = 0
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "COD_ACESSORIO"
      .TextMatrix(0, 4) = "ACESSÓRIO"
      
      .Redraw = False
      
    
      .Rows = .Rows - 1
      .Redraw = True
      .Visible = True
   End With
End Sub
Private Sub FormatarGrid_Situacao(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Situacao
      .Visible = False
      
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 0
      .ColWidth(4) = 2900
      
      .RowHeight(0) = 0
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "COD_ACESSORIO"
      .TextMatrix(0, 4) = "ACESSÓRIO"
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = rTabela("cod_os")
            .TextMatrix(.Rows - 1, 3) = rTabela("cod_situacao")
            .TextMatrix(.Rows - 1, 4) = rTabela("situacao")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
      .Visible = True
   End With
End Sub
Private Sub FormatarGrid_Acessorios(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Acessorio
      .Visible = False
      
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 0
      .ColWidth(4) = 2900
      
      .RowHeight(0) = 0
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "COD_ACESSORIO"
      .TextMatrix(0, 4) = "ACESSÓRIO"
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = rTabela("cod_os")
            .TextMatrix(.Rows - 1, 3) = rTabela("cod_acessorio")
            .TextMatrix(.Rows - 1, 4) = rTabela("acessorio")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
      .Visible = True
   End With
End Sub

Private Sub MostrarGrid_Situacao()
If txtCodOS.Text = "" Then txtCodOS.Text = 0

sSQL = "SELECT * FROM OS_situacao_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Situacao r

If r.State <> 0 Then r.Close
End Sub
Private Sub MostrarGrid_Acessorios()
If txtCodOS.Text = "" Then txtCodOS.Text = 0

sSQL = "SELECT * FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Acessorios r

If r.State <> 0 Then r.Close
End Sub

Private Sub cmdRemoverPecas_Click()
On Error GoTo erro
i = Grid_Servicos.Row

'CHECAR SE A OS ESTÁ FECHADA
Verificar_OS_Fechada
If OS_FECHADA = True Then Exit Sub

'REMOVER O ITEM DA LISTA
If Grid_Servicos.TextMatrix(Grid_Servicos.Row, 1) = "" Then GoSub erro
If Grid_Servicos.TextMatrix(i, 2) = "SERVIÇO" Then MsgBox "Esse botão só permite remover produtos!", vbInformation, "Aviso do Sistema": Exit Sub
If ShowMsg("Deseja remover da lista a peça: " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 3) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    dbData.Execute "DELETE FROM pedidos_itens WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9) & ") AND (cod_pedido = " & txtCodPedido.Text & ");"
ElseIf vTipoOS = "Recapadora" Then
    dbData.Execute "DELETE FROM pedidos_itens WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 12) & ") AND (cod_pedido = " & txtCodPedido.Text & ");"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "DELETE FROM pedidos_itens WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9) & ") AND (cod_pedido = " & txtCodPedido.Text & ");"
ElseIf vTipoOS = "Comunicação Visual" Then
    dbData.Execute "DELETE FROM pedidos_itens WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9) & ") AND (cod_pedido = " & txtCodPedido.Text & ");"
End If

MostrarGrid_Servicos

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Recapadora" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Comunicação Visual" Then
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
End If
Somar_Totais
Exit Sub

erro:
   ShowMsg "Não existe nenhuma peça para ser removido!", vbExclamation
   Exit Sub
End Sub

Private Sub cmdRemoverServicosAuto_Click()
On Error GoTo erro

i = Grid_Servicos.Row
'CHECAR SE A OS ESTÁ FECHADA
Verificar_OS_Fechada
If OS_FECHADA = True Then Exit Sub

'REMOVER O ITEM DA LISTA
If Grid_Servicos.TextMatrix(Grid_Servicos.Row, 1) = "" Then GoSub erro
If Grid_Servicos.TextMatrix(i, 2) = "PRODUTO" Then MsgBox "Esse botão só permite remover serviços!", vbInformation, "Aviso do Sistema": Exit Sub
If ShowMsg("Deseja remover da lista o serviço: " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 3) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    dbData.Execute "DELETE FROM OS_Servicos_Auto WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9) & ") AND (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Recapadora" Then
    dbData.Execute "DELETE FROM os_servicos_recapadora WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 12) & ") AND (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    dbData.Execute "DELETE FROM OS_Servicos_Auto WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9) & ") AND (cod_os = " & txtCodOS.Text & ");"
ElseIf vTipoOS = "Comunicação Visual" Then
    dbData.Execute "DELETE FROM OS_Servicos_Comunicacao WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9) & ") AND (cod_os = " & txtCodOS.Text & ");"
End If

MostrarGrid_Servicos

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Recapadora" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Recapadora AS OS_Servicos_Recapadora_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
ElseIf vTipoOS = "Comunicação Visual" Then
    'atualizar tabela OS
    dbData.Execute "UPDATE OS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_OS = " & txtCodOS.Text & ")"
    
    'atualizar tabela PEDIDOS
    dbData.Execute "UPDATE PEDIDOS SET " & _
    "SUBTOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Subtotal), 0) FROM pedidos_itens WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "ValorDescReal = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "VALOR_DESC = (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(desconto), 0) FROM pedidos_itens AS pedidos_itens_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")), " & _
    "TOTAL = (SELECT ISNULL(SUM(preco * quantidade), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_2 WHERE (cod_os = " & txtCodOS.Text & ")) - (SELECT ISNULL(SUM(desconto), 0) FROM OS_Servicos_Auto AS OS_Servicos_Auto_1 WHERE (cod_os = " & txtCodOS.Text & ")) + (SELECT ISNULL(SUM(Total), 0) FROM pedidos_itens AS pedidos_itens_2 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")) " & _
    "Where (COD_PEDIDO = " & txtCodPedido.Text & ")"
End If

Somar_Totais
Exit Sub
   
erro:
   ShowMsg "Não existe nenhum serviço para ser removido!", vbExclamation
   Exit Sub
End Sub


Private Sub cmdRemoverSituacao_Click()
   On Error GoTo erro
   
   If Grid_Situacao.TextMatrix(Grid_Situacao.Row, 1) = "" Then GoSub erro
   If ShowMsg("Deseja excluir o acessório: " & Grid_Situacao.TextMatrix(Grid_Situacao.Row, 4) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   dbData.Execute "DELETE FROM OS_situacao_Auto WHERE (codigo = " & Grid_Situacao.TextMatrix(Grid_Situacao.Row, 1) & ") AND (cod_os = " & txtCodOS.Text & ");"
   
   MostrarGrid_Situacao
   Exit Sub
   
erro:
   ShowMsg "Não existe nenhum acessório para ser excluido!", vbExclamation
   Exit Sub
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub cmdSalvarParecer_Click()
If txtCodOS.Text = "" Then Exit Sub
If txtParecerTecnico.Text = "" Then Exit Sub
dbData.Execute "UPDATE OS SET OBS = '" & txtParecerTecnico.Text & "' WHERE (COD_OS = " & txtCodOS.Text & ");"
txtParecerTecnico.Text = ""
frmParecer.Visible = False
End Sub

Private Sub Form_Load()
Set oCfg = sysConfig("TIPO_OS")
vTipoOS = oCfg.Value
Set oCfg = Nothing

If vTipoOS = "Automóveis" Then
    menu_Cadastro_Pneus.Visible = False
    frmEquipamento.Caption = "Automóvel"
    Mostrar_Equipamento_Automoveis
    chkVeiculo.Enabled = True
    chkVeiculo.Caption = "Mostrar Veículo"
ElseIf vTipoOS = "Motocicletas" Then
    menu_Cadastro_Pneus.Visible = False
    frmEquipamento.Caption = "Motocicleta"
    Mostrar_Equipamento_Automoveis
    chkVeiculo.Enabled = True
    chkVeiculo.Caption = "Mostrar Veículo"
ElseIf vTipoOS = "Motores" Then
    menu_Cadastro_Pneus.Visible = False
    frmEquipamento.Caption = "Equipamento"
ElseIf vTipoOS = "Gráfica Rápida" Then
    menu_Cadastro_Pneus.Visible = False
    frmEquipamento.Caption = "Produto"
ElseIf vTipoOS = "Informática" Then
    menu_Cadastro_Pneus.Visible = False
    frmEquipamento.Caption = "Equipamento"
    Mostrar_Equipamentos_Informatica
    chkVeiculo.Enabled = False
    chkVeiculo.Caption = "Mostrar Equip."
ElseIf vTipoOS = "Celular" Then
    menu_Cadastro_Pneus.Visible = False
    frmEquipamento.Caption = "Equipamento"
    Mostrar_Equipamentos_Informatica
    chkVeiculo.Enabled = False
    chkVeiculo.Caption = "Mostrar Equip."
ElseIf vTipoOS = "Recapadora" Then
    menu_Cadastro_Pneus.Visible = True
    Mostrar_Equipamento_Automoveis
    frmEquipamento.Caption = "Veículo"
    chkVeiculo.Enabled = False
    chkVeiculo.Caption = "Mostrar Veículo"
ElseIf vTipoOS = "Comunicação Visual" Then
    frmEquipamento.Caption = "Equipamento"
    Mostrar_Equipamentos_Informatica
    chkVeiculo.Enabled = False
    chkVeiculo.Caption = "Mostrar Equip."
    Menu_Cadastro_Acessorios.Visible = False
    Menu_Cadastro_Situacoes.Visible = False
    menu_Cadastro_Pneus.Visible = False
ElseIf vTipoOS = "Agrícola" Then
    menu_Cadastro_Pneus.Visible = False
    frmEquipamento.Caption = "Maquina"
End If

vTipoConsPecas = 0

SSTab1.Tab = 0
cmdNovo.Enabled = True
cmdGerarEntrada.Enabled = False
cmdCancelarEntrada.Enabled = False
cmdAlterar.Enabled = False
cmdApagar.Enabled = False
menu_Impressao_Pedido.Enabled = False
'cmdEditarOS.Visible = False
'cboStatus.Enabled = False
frmSecundario.Enabled = False
txtDesc.Text = 0
LimparGrid_Servicos
LimparGrid_Acessorios
'lblTotal.Caption = Format(0, ocMONEY)
'lblTotalPeca.Caption = Format(0, ocMONEY)
Preencher_TipoServico
Preencher_Mostrar
Preencher_Status
Preencher_Criterios
Preencher_Indice
cboConsultaMostrar.ListIndex = 0
cboConsultaStatus.ListIndex = 0
cboConsultaCriterios.ListIndex = 0
cboTipoServico.ListIndex = 0
cboIndice.ListIndex = 0
MostrarGrid_OS
MostrarGrid_OS_Situacao
lblValidade.Caption = Format(DateAdd("m", 1, Date), "dd/mm/yy")
SSTab1.TabVisible(5) = False
cmdEditarOS.Enabled = False
cmdFinanceiroOS.Enabled = False
cmdImpEntrada1.Enabled = False
cmdImpOrcamento1.Enabled = False
cmdImpEntrada2.Enabled = False
cmdImpOrcamento2.Enabled = False
cmdImpPedido1.Enabled = False
cmdImpGarantia1.Enabled = False
cmdImpPedido2.Enabled = False
'frmServicos.Visible = False
frmGridServicos.Visible = False
frmTotaisGeral.Visible = False
frmTotaisProdServ.Visible = False
ExibirObjetosServicos

'colocar o nome da maquina na barra de status
Dim var_Maquina As String
Dim var_Caixa As String

Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
Set oIni = Nothing

StatusBar1.Panels(2).Text = var_Caixa
StatusBar1.Panels(4).Text = var_Maquina

'tipos de descontos e valores
Set oCfg = sysConfig("LIMITEDESCONTO")
vLimitarDesc = oCfg.Value
Set oCfg = Nothing

Set oCfg = sysConfig("TIPODESCONTO")
vTipoDesc = oCfg.Value
Set oCfg = Nothing
    
If vTipoDesc = "1" Then
    Set oCfg = sysConfig("DESC_AV")
    vValorDescFixoAV = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AP")
    vValorDescFixoAP = oCfg.Value
    Set oCfg = Nothing
ElseIf vTipoDesc = "2" Then
    Set oCfg = sysConfig("DESC_AV")
    vValorDescFixoAV = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AP")
    vValorDescFixoAP = oCfg.Value
    Set oCfg = Nothing
ElseIf vTipoDesc = "3" Then
    Set oCfg = sysConfig("DESC_MARGEM_AV1")
    vMargemDescGradual1 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_MARGEM_AV2")
    vMargemDescGradual2 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_MARGEM_AV3")
    vMargemDescGradual3 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AV1")
    vValorDescGradualAV1 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AV2")
    vValorDescGradualAV2 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AV3")
    vValorDescGradualAV3 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AP1")
    vValorDescGradualAP1 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AP2")
    vValorDescGradualAP2 = oCfg.Value
    Set oCfg = Nothing

    Set oCfg = sysConfig("DESC_AP3")
    vValorDescGradualAP3 = oCfg.Value
    Set oCfg = Nothing
End If

   Set oCfg = sysConfig("COPIAS_AP")
   iCopiasAP = CInt(oCfg.Value)
   
   Set oCfg = sysConfig("ENTREGA_AP")
   bEntregaAP = CBool(oCfg.Value)
   
   Set oCfg = sysConfig("IMP_AP")
   bImprAP = CBool(oCfg.Value)
   
   Set oCfg = sysConfig("CONF_IMPRESSAO_AP")
   bConfImprAP = CBool(oCfg.Value)
    
    Set oCfg = sysConfig("CONF_FECHAMENTO_AP")
    bFechAP = CBool(oCfg.Value)
   ' Set oCfg = Nothing
   
    Set oCfg = sysConfig("CONF_FECHAMENTO_AV")
    bFechAV = CBool(oCfg.Value)
    'Set oCfg = Nothing


Verificar_Caixa

StatusBar1.Panels(5).Text = Format(Date, "dd/mm/yy")

Set moCombo = New cComboHelper
End Sub

Private Sub ConsultarCaixaAtual()
sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    StatusBar1.Panels(3).Text = ValidateNull(r("codcaixa"))
Else
    StatusBar1.Panels(3).Text = 0
End If
End Sub
Private Sub FormatarGrid_OS_Situacao(rTabela As ADODB.Recordset)
Dim i As Integer
Dim aCor As ColorConstants
Dim totalRegistros As Long

With Grid_OS
   .Rows = 1       'INICIA O Grid_OS COM UMA LINHA
   .FixedCols = 0  'DETERMINA QUE NÃO HAJA COLUNA FIXA
   
   'Abaixo o cabeçalho é criado
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
   .FormatString = "^CÓD.|^TIPO|^SITUAÇÃO|^PGTO|^CLIENTE|^VEICULO|^ENTRADA|^PEDIDO"
    .ColWidth(0) = 630
    .ColWidth(1) = 1000
    .ColWidth(2) = 1200
    .ColWidth(3) = 900
    .ColWidth(4) = 4000
    .ColWidth(5) = 3200
    .ColWidth(6) = 1350
    .ColWidth(7) = 0
    
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
         .ColAlignment(3) = 2
         .ColAlignment(4) = 1
         .ColAlignment(5) = 1
         
         'A linha abaixo cria mais linha no Grid_OS
         .Rows = .Rows + 1
         
         'Preenche com os dados, e assim sucessivamente
         .TextMatrix(.Rows - 1, 0) = Format(rTabela("cod_os"), "00000")
         .TextMatrix(.Rows - 1, 1) = rTabela("TIPO_OS")
         .TextMatrix(.Rows - 1, 2) = rTabela("var_status")
         .TextMatrix(.Rows - 1, 3) = rTabela("var_status_Financeiro")
         .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("nome"))
         .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("fabricante")) & " | " & ValidateNull(rTabela("modelo")) & " | " & ValidateNull(rTabela("ANO")) & " | " & ValidateNull(rTabela("PLACA"))
         .TextMatrix(.Rows - 1, 6) = Format(rTabela("DATA_ENTRADA"), "dd/mm/yy") & " - " & Format(rTabela("HORA_ENTRADA"), ocHRMN)
         .TextMatrix(.Rows - 1, 7) = Format(rTabela("cod_PEDIDO"), "00000")
         rTabela.MoveNext
      Loop
   End If
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
   .FormatString = "^CÓD.|^TECNICO|^FINANCEIRO|^CLIENTE|^EQUIPAMENTO|^ENTRADA|^PEDIDO"
   .ColWidth(0) = 650
   .ColWidth(1) = 1500
   .ColWidth(2) = 1200
   .ColWidth(3) = 4000
   .ColWidth(4) = 3000
   .ColWidth(5) = 1350
   .ColWidth(6) = 0
    
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
         .ColAlignment(2) = 2
         .ColAlignment(3) = 1
         .ColAlignment(4) = 1
         
         'A linha abaixo cria mais linha no Grid_OS
         .Rows = .Rows + 1
         
         'Preenche com os dados, e assim sucessivamente
         .TextMatrix(.Rows - 1, 0) = Format(rTabela("cod_os"), "00000")
         .TextMatrix(.Rows - 1, 1) = rTabela("var_status")
         .TextMatrix(.Rows - 1, 2) = rTabela("var_status_Financeiro")
         .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("nome"))
         .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("Equipamento")) & " | " & ValidateNull(rTabela("Fabricante")) & " | " & ValidateNull(rTabela("Modelo"))
         .TextMatrix(.Rows - 1, 5) = Format(rTabela("DATA_ENTRADA"), "dd/mm/yy") & " - " & Format(rTabela("HORA_ENTRADA"), ocHRMN)
         .TextMatrix(.Rows - 1, 6) = Format(rTabela("cod_PEDIDO"), "00000")
         rTabela.MoveNext
      Loop
   End If
ElseIf vTipoOS = "Comunicação Visual" Then
   .FormatString = "^CÓD.|^TECNICO|^FINANCEIRO|^CLIENTE|^ENTRADA|^PEDIDO|^PEDIDO"
   .ColWidth(0) = 650
   .ColWidth(1) = 1500
   .ColWidth(2) = 1200
   .ColWidth(3) = 7000
   .ColWidth(4) = 1700
   .ColWidth(5) = 0
   .ColWidth(6) = 0
    
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
         .ColAlignment(2) = 2
         .ColAlignment(3) = 1
         .ColAlignment(4) = 1
         
         'A linha abaixo cria mais linha no Grid_OS
         .Rows = .Rows + 1
         
         'Preenche com os dados, e assim sucessivamente
         .TextMatrix(.Rows - 1, 0) = Format(rTabela("cod_os"), "00000")
         .TextMatrix(.Rows - 1, 1) = rTabela("var_status")
         .TextMatrix(.Rows - 1, 2) = rTabela("var_status_Financeiro")
         .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("nome"))
         .TextMatrix(.Rows - 1, 4) = Format(rTabela("DATA_ENTRADA"), "dd/mm/yy") & " - " & Format(rTabela("HORA_ENTRADA"), ocHRMN)
         .TextMatrix(.Rows - 1, 5) = Format(rTabela("cod_PEDIDO"), "00000")
         .TextMatrix(.Rows - 1, 6) = Format(rTabela("cod_PEDIDO"), "00000")
         rTabela.MoveNext
      Loop
   End If
End If

 'Colocar a coluna em negrito
   For i = 1 To .Rows - 1
      .Row = i
      .Col = 1
      .CellFontBold = True
   Next
 
   'mudar a cor da fonte
   For i = 1 To .Rows - 1
      If UCase(Trim(.TextMatrix(i, 2))) = UCase("À COMEÇAR") Then
         aCor = vbBlack
      ElseIf UCase(Trim(.TextMatrix(i, 2))) = UCase("EM EXECUÇÃO") Then
         aCor = &H8000&
      ElseIf UCase(Trim(.TextMatrix(i, 2))) = UCase("AGUARDANDO") Then
         aCor = vbBlue
      ElseIf UCase(Trim(.TextMatrix(i, 2))) = UCase("TERMINADO") Then
         aCor = vbRed
      End If
      
      .Col = 2 'a coluna do aberto ou fechado
      .Row = i
      .CellForeColor = aCor
   Next
   
   .Redraw = True
End With
End Sub
Private Sub FormatarGrid_OS(rTabela As ADODB.Recordset)
Dim i As Integer
Dim aCor As ColorConstants
Dim totalRegistros As Long

With Grid
   .Rows = 1       'INICIA O GRID COM UMA LINHA
   .FixedCols = 0  'DETERMINA QUE NÃO HAJA COLUNA FIXA
   
   'Abaixo o cabeçalho é criado
   .FormatString = "^CÓD.|^TECNICO|^FINANC.|^CLIENTE|^TIPO|^FORMA|^VALOR|^DESC.|^TOTAL"
   .ColWidth(0) = 650
   .ColWidth(1) = 1250
   .ColWidth(2) = 1000
   .ColWidth(3) = 5350
   .ColWidth(4) = 750
   .ColWidth(5) = 750
   .ColWidth(6) = 850
   .ColWidth(7) = 650
   .ColWidth(8) = 850
    
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
         .ColAlignment(3) = 1
         .ColAlignment(6) = 0
         .ColAlignment(5) = 0
         .ColAlignment(6) = 6
         .ColAlignment(7) = 6
         .ColAlignment(8) = 6
         
         'A linha abaixo cria mais linha no grid
         .Rows = .Rows + 1
         
         'Preenche com os dados, e assim sucessivamente
         .TextMatrix(.Rows - 1, 0) = Format(rTabela("cod_os"), "0000")
         .TextMatrix(.Rows - 1, 1) = rTabela("var_status")
         .TextMatrix(.Rows - 1, 2) = rTabela("var_status_os") & ""
         
         If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo")) & " / " & ValidateNull(rTabela("ano"))
         ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("equipamento")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo"))
         ElseIf vTipoOS = "Comunicação Visual" Then
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("equipamento")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo"))
         End If
         .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("TIPO_PAGAMENTO"))
         .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("PAGAMENTO"))
         .TextMatrix(.Rows - 1, 6) = Format(rTabela("SUBTOTAL"), ocMONEY)
         .TextMatrix(.Rows - 1, 7) = Format(rTabela("ValorDescReal"), ocMONEY)
         .TextMatrix(.Rows - 1, 8) = Format(rTabela("TOTAL"), ocMONEY)
         rTabela.MoveNext
      Loop
   End If
   
   'agora sim coloco a fução para mudar a cor da coluna e pronto
   'mudar a cor da fonte
   For i = 1 To .Rows - 1
      If UCase(Trim(.TextMatrix(i, 2))) = UCase("ABERTO") Then
         aCor = vbBlue
      Else
         aCor = vbRed
      End If
      
      .Col = 2 'a coluna do aberto ou fechado
      .Row = i
      .CellForeColor = aCor
   Next
   
   'mudar a cor da fonte
   For i = 1 To .Rows - 1
      If UCase(Trim(.TextMatrix(i, 1))) = UCase("À COMEÇAR") Then
         aCor = vbBlack
      ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("EM EXECUÇÃO") Then
         aCor = vbGreen
      ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("AGUARDANDO") Then
         aCor = vbBlue
      ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("TERMINADO") Then
         aCor = vbRed
      End If
      
      .Col = 1 'a coluna do aberto ou fechado
      .Row = i
      .CellForeColor = aCor
   Next
   
   .Redraw = True
End With

lblTotalConsulta.Caption = Format(SomaGrid(Grid, 8), ocMONEY)
End Sub
Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
'Dim i As Integer
Dim Valor As Currency

Valor = 0
For i = 0 To var_Grid.Rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If txtCodOS.Text <> "" And cmdGerarEntrada.Enabled = True Then
    MsgBox "A Ordem de Serviço iniciada ainda não foi salvo", vbInformation, "Aviso do Sistema"
    Cancel = True
    SSTab1.Tab = 1
    Exit Sub
Else
    KillProcess "Ordem de Serviço"
End If
End Sub

Private Sub Grid_DblClick()
If txtCodOS.Text <> "" And cmdGerarEntrada.Enabled = True Then
    MsgBox "A Ordem de Serviço iniciada ainda não foi salvo", vbInformation, "Aviso do Sistema"
    SSTab1.Tab = 1
    Exit Sub
End If

SSTab1.Tab = 1
frmSecundario.Enabled = True
cboStatus.Enabled = True
cmdGerarEntrada.Enabled = False
cmdCancelarEntrada.Enabled = False
cmdAlterar.Enabled = True
cmdApagar.Enabled = True
cmdNovo.Enabled = True
txtCodOS.Text = ""
txtCodOS.Text = (Grid.TextMatrix(Grid.Row, 0))
End Sub

Private Sub cmdCancelarEntrada_Click()
If txtCodOS.Text = "" Then Exit Sub

If ShowMsg("Cancelando a OS todos os produtos adicionado até agora serão perdidos!" & vbCrLf & "Deseja cancelar essa OS ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

   'EXCLUIR NA TABELA OS
   dbData.Execute "DELETE FROM os WHERE (cod_os = " & txtCodOS.Text & ");"
   
   'EXCLUIR NA TABELA PEDIDOS_ITENS
   dbData.Execute "DELETE FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ");"
   
   'EXCLUIR NA TABELA PEDIDOS
   dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
   
   'EXCLUIR NA TABELA PARCELAS
   dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & txtCodPedido.Text & ");"

    If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
        dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM OS_Equipamento_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"
    ElseIf vTipoOS = "Recapadora" Then
        dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM OS_Equipamento_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"
    ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
        dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM OS_Equipamento WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"
    ElseIf vTipoOS = "Comunicação Visual" Then
        'dbData.Execute "DELETE FROM OS_acessorios_Auto WHERE (cod_os = " & txtCodOS.Text & ");"
        'dbData.Execute "DELETE FROM OS_Equipamento WHERE (cod_os = " & txtCodOS.Text & ");"
        dbData.Execute "DELETE FROM OS_Servicos_Comunicacao WHERE (cod_os = " & txtCodOS.Text & ");"
        'dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"
    End If

   'EXCLUIR NA TABELA SITUAÇÃO
   'dbData.Execute "DELETE FROM os_situacao_auto WHERE (cod_os = " & txtCodOS.Text & ");"

LimparObjetos_Entrada
LimparObjetos_Servicos
LimparObjetos_Pecas
txtCodOS.Text = ""
txtCodPedido.Text = ""
Form_Load
End Sub

Private Sub Grid_OS_Click()
If Grid_OS.Col <> 0 Then Exit Sub

MostrarGrid_PecasServicos

'Dim i As Integer
'For i = 1 To Grid_OS.Rows - 1
'   If UCase(Trim(Grid_OS.TextMatrix(i, 1))) = UCase("À COMEÇAR") Then
'      cmdEditarOS.Enabled = True
'   ElseIf UCase(Trim(Grid_OS.TextMatrix(i, 1))) = UCase("EM EXECUÇÃO") Then
'      cmdEditarOS.Enabled = True
'   ElseIf UCase(Trim(Grid_OS.TextMatrix(i, 1))) = UCase("AGUARDANDO") Then
'      cmdEditarOS.Enabled = True
'   ElseIf UCase(Trim(Grid_OS.TextMatrix(i, 1))) = UCase("TERMINADO") Then
'      cmdEditarOS.Enabled = False
'   End If
'Next

Dim posit As Long
posit = Grid_OS.Row

If (Trim(Grid_OS.TextMatrix(posit, 2))) = ("À COMEÇAR") Then
    'MsgBox Trim(Grid_OS.TextMatrix(posit, 1))
      cmdEditarOS.Enabled = True
      cmdFinanceiroOS.Enabled = False
      cmdImpEntrada1.Enabled = True
      cmdImpOrcamento1.Enabled = False
      cmdImpPedido1.Enabled = False
      cmdOrcamentoPDF.Enabled = False
      cmdPedidoPDF.Enabled = False
   ElseIf (Trim(Grid_OS.TextMatrix(posit, 2))) = ("EM EXECUÇÃO") Then
   'MsgBox Trim(Grid_OS.TextMatrix(posit, 1))
      cmdEditarOS.Enabled = True
      cmdFinanceiroOS.Enabled = False
      cmdImpEntrada1.Enabled = True
      cmdImpOrcamento1.Enabled = True
      cmdImpPedido1.Enabled = False
      cmdOrcamentoPDF.Enabled = True
      cmdPedidoPDF.Enabled = False
   ElseIf (Trim(Grid_OS.TextMatrix(posit, 2))) = ("AGUARDANDO") Then
   'MsgBox Trim(Grid_OS.TextMatrix(posit, 1))
      cmdEditarOS.Enabled = True
      cmdFinanceiroOS.Enabled = False
      cmdImpEntrada1.Enabled = True
      cmdImpOrcamento1.Enabled = True
      cmdImpPedido1.Enabled = False
      cmdOrcamentoPDF.Enabled = True
      cmdPedidoPDF.Enabled = False
   ElseIf (Trim(Grid_OS.TextMatrix(posit, 2))) = ("TERMINADO") Then
    If (Trim(Grid_OS.TextMatrix(posit, 3))) = ("ABERTO") Then
        cmdFinanceiroOS.Enabled = True
        cmdImpPedido1.Enabled = False
        cmdImpGarantia1.Enabled = False
        cmdPedidoPDF.Enabled = False
    Else
        cmdFinanceiroOS.Enabled = False
        cmdImpPedido1.Enabled = True
        cmdImpGarantia1.Enabled = True
        cmdPedidoPDF.Enabled = True
    End If
      cmdEditarOS.Enabled = False
      cmdImpEntrada1.Enabled = True
      cmdImpOrcamento1.Enabled = True
      cmdOrcamentoPDF.Enabled = True
   End If
If optFinanceiroAberto.Value = True Then cmdExcluir.Enabled = True
End Sub




Private Sub Menu_Cadastro_Acessorios_Click()
OS_Automoveis_Acessorios.Show 1
End Sub

Private Sub menu_Cadastro_Cliente_Click()
Clientes_Cadastro.Show 1
End Sub


Private Sub Menu_Cadastro_Parecer_Click()
If txtCodOS.Text = "" Then Exit Sub

sSQL = "SELECT OBS FROM os WHERE (cod_os = " & txtCodOS.Text & ");"
Set rs = dbData.OpenRecordset(sSQL)

If Not rs.BOF Then
   txtParecerTecnico.Text = ValidateNull(rs("OBS"))
End If


SSTab1.Tab = 1
frmParecer.Visible = True
End Sub

Private Sub menu_Cadastro_Pecas_Click()
Produtos_Cadastro.Show 1
End Sub

Private Sub menu_Cadastro_Pneus_Click()
OS_Recapadora_Pneus.Show 1
End Sub

Private Sub menu_Cadastro_Servicos_Click()
If vTipoOS = "Automóveis" Then
    OS_CAD_Servicos_Geral.Show 1
ElseIf vTipoOS = "Motocicletas" Then
    OS_CAD_Servicos_Geral.Show 1
ElseIf vTipoOS = "Recapadora" Then
    OS_CAD_Servicos_Recapadora.Show 1
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    OS_CAD_Servicos_Geral.Show 1
ElseIf vTipoOS = "Comunicação Visual" Then
    OS_CAD_Servicos_Geral.Show 1
End If
End Sub



Private Sub Menu_Cadastro_Situacoes_Click()
OS_Situacao.Show 1
End Sub

Private Sub menu_Impressao_Entrada_Click()
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

If txtCodOS.Text = "" Or txtCodCliente.Text = "" Then
'If txtCodOS.Text = "" Or txtCodCliente.Text = "" Or cboFabricante.Text = "" Then
   ShowMsg "Não é possível imprimir uma Ordem de Serviço em branco!", vbInformation
   Exit Sub
End If

Me.Hide

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    With REL_OS_Entrada_Automoveis
        .txtOS.Caption = " " & Format(txtCodOS.Text, "000000")
        .txtCliente.Caption = " " & UCase(cboCliente.Text)
        .txtSaida.Caption = " " & Format(mskDataSaida.Text, "dd/mm/yy") & " - " & Format(mskHoraSaida.Text, "hh:mm")
        .txtDataEntrada.Caption = " " & Format(mskDataEntrada.Text, "dd/mm/yy") & " - " & Format(mskHoraEntrada.Text, "hh:mm")
        .txtFuncionario.Caption = " " & UCase(cboFuncionario)
        .txtFabricante.Caption = " " & UCase(cboFabricante.Text)
        .txtModelo.Caption = " " & UCase(cboModelo.Text)
        .txtAno.Caption = " " & UCase(txtAno.Text)
        .txtCor.Caption = " " & UCase(cboCor.Text)
        .txtPlaca1.Caption = " " & UCase(txtPlaca.Text)
        .txtChassi.Caption = " " & UCase(txtChassi.Text)
        .txtKM.Caption = " " & UCase(txtKM.Text)
        .txtTanque.Caption = " " & UCase(cboTanque.Text)
        .txtDescricao.Caption = " " & UCase(txtPareceCliente.Text)
        '.txtEquipamento.Caption = " " & UCase(cboFabricante.Text) & " - " & UCase(cboModelo.Text) & " - " & UCase(txtAno.Text) & " - " & UCase(txtPlaca.Text)
        .Preencher_Acessorios txtCodOS.Text
        .Preencher_Situacao txtCodOS.Text
        .Relatorio.NumeroRegistros = 1
        .Relatorio.NomeImpressora = var_ImpNormal
        .Relatorio.Ativar
    End With
    Unload REL_OS_Entrada_Automoveis
ElseIf vTipoOS = "Recapadora" Then
    With REL_OS_Entrada_Automoveis
        .txtOS.Caption = " " & Format(txtCodOS.Text, "000000")
        .txtCliente.Caption = " " & UCase(cboCliente.Text)
        .txtSaida.Caption = " " & Format(mskDataSaida.Text, "dd/mm/yy") & " - " & Format(mskHoraSaida.Text, "hh:mm")
        .txtDataEntrada.Caption = " " & Format(mskDataEntrada.Text, "dd/mm/yy") & " - " & Format(mskHoraEntrada.Text, "hh:mm")
        .txtFuncionario.Caption = " " & UCase(cboFuncionario)
        .txtFabricante.Caption = " " & UCase(cboFabricante.Text)
        .txtModelo.Caption = " " & UCase(cboModelo.Text)
        .txtAno.Caption = " " & UCase(txtAno.Text)
        .txtCor.Caption = " " & UCase(cboCor.Text)
        .txtPlaca1.Caption = " " & UCase(txtPlaca.Text)
        .txtKM.Caption = " " & UCase(txtKM.Text)
        .txtTanque.Caption = " " & UCase(cboTanque.Text)
        .txtDescricao.Caption = " " & UCase(txtPareceCliente.Text)
        '.txtEquipamento.Caption = " " & UCase(cboFabricante.Text) & " - " & UCase(cboModelo.Text) & " - " & UCase(txtAno.Text) & " - " & UCase(txtPlaca.Text)
        .Preencher_Acessorios txtCodOS.Text
        .Preencher_Situacao txtCodOS.Text
        .Relatorio.NumeroRegistros = 1
        .Relatorio.NomeImpressora = var_ImpNormal
        .Relatorio.Ativar
    End With
    Unload REL_OS_Entrada_Automoveis
ElseIf vTipoOS = "Informática" Then
    With REL_OS_Entrada_Informatica
        .txtOS.Caption = " " & Format(txtCodOS.Text, "000000")
        .txtCliente.Caption = " " & UCase(cboCliente.Text)
        .txtSaida.Caption = " " & Format(mskDataSaida.Text, "dd/mm/yy") & " - " & Format(mskHoraSaida.Text, "hh:mm")
        .txtDataEntrada.Caption = " " & Format(mskDataEntrada.Text, "dd/mm/yy") & " - " & Format(mskHoraEntrada.Text, "hh:mm")
        .txtFuncionario.Caption = " " & UCase(cboFuncionario)
        
        '.txtFabricante.Caption = " " & UCase(cboFabricante.Text)
        '.txtModelo.Caption = " " & UCase(cboModelo.Text)
        '.txtEquipamento.Caption = " " & UCase(cboTanque.Text)
        
        .txtEquipamento.Caption = " " & UCase(cboTanque.Text)
        .txtMarca.Caption = " " & UCase(cboFabricante.Text)
        .txtModelo.Caption = " " & UCase(cboModelo.Text)
        
        .txtDescricao.Caption = " " & UCase(txtPareceCliente.Text)
        .Preencher_Acessorios txtCodOS.Text
        .Preencher_Situacao txtCodOS.Text
        .Relatorio.NumeroRegistros = 1
        .Relatorio.NomeImpressora = var_ImpNormal
        .Relatorio.Ativar
    End With
    Unload REL_OS_Entrada_Informatica
ElseIf vTipoOS = "Celular" Then
    With REL_OS_Entrada_Celular
        .txtOS.Caption = " " & Format(txtCodOS.Text, "000000")
        .txtCliente.Caption = " " & UCase(cboCliente.Text)
        .txtSaida.Caption = " " & Format(mskDataSaida.Text, "dd/mm/yy") & " - " & Format(mskHoraSaida.Text, "hh:mm")
        .txtDataEntrada.Caption = " " & Format(mskDataEntrada.Text, "dd/mm/yy") & " - " & Format(mskHoraEntrada.Text, "hh:mm")
        .txtFuncionario.Caption = " " & UCase(cboFuncionario)
        
        '.txtFabricante.Caption = " " & UCase(cboFabricante.Text)
        '.txtModelo.Caption = " " & UCase(cboModelo.Text)
        '.txtEquipamento.Caption = " " & UCase(cboTanque.Text)
        
        .txtEquipamento.Caption = " " & UCase(cboTanque.Text)
        .txtMarca.Caption = " " & UCase(cboFabricante.Text)
        .txtModelo.Caption = " " & UCase(cboModelo.Text)
        
        .txtDescricao.Caption = " " & UCase(txtPareceCliente.Text)
        .Preencher_Acessorios txtCodOS.Text
        .Preencher_Situacao txtCodOS.Text
        .Relatorio.NumeroRegistros = 1
        .Relatorio.NomeImpressora = var_ImpNormal
        .Relatorio.Ativar
    End With
    Unload REL_OS_Entrada_Celular
ElseIf vTipoOS = "Comunicação Visual" Then
    With REL_OS_Entrada_Informatica
        .txtOS.Caption = " " & Format(txtCodOS.Text, "000000")
        .txtCliente.Caption = " " & UCase(cboCliente.Text)
        .txtSaida.Caption = " " & Format(mskDataSaida.Text, "dd/mm/yy") & " - " & Format(mskHoraSaida.Text, "hh:mm")
        .txtDataEntrada.Caption = " " & Format(mskDataEntrada.Text, "dd/mm/yy") & " - " & Format(mskHoraEntrada.Text, "hh:mm")
        .txtFuncionario.Caption = " " & UCase(cboFuncionario)
        
        '.txtFabricante.Caption = " " & UCase(cboFabricante.Text)
        '.txtModelo.Caption = " " & UCase(cboModelo.Text)
        '.txtEquipamento.Caption = " " & UCase(cboTanque.Text)
        
        .txtEquipamento.Caption = " " & UCase(cboTanque.Text)
        .txtMarca.Caption = " " & UCase(cboFabricante.Text)
        .txtModelo.Caption = " " & UCase(cboModelo.Text)
        
        .txtDescricao.Caption = " " & UCase(txtPareceCliente.Text)
        .Preencher_Acessorios txtCodOS.Text
        .Preencher_Situacao txtCodOS.Text
        .Relatorio.NumeroRegistros = 1
        .Relatorio.NomeImpressora = var_ImpNormal
        .Relatorio.Ativar
    End With
    Unload REL_OS_Entrada_Informatica
End If
Me.Show 1
End Sub

Private Sub menu_Impressao_Garantia_Click()
   'colocar o nome da maquina na barra de status
   Dim var_Impressora As String
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   If txtCodOS.Text = "" Or txtCodCliente.Text = "" Or cboFabricante.Text = "" Then
      ShowMsg "Não é possível imprimir uma Ordem de Serviço em branco!", vbInformation
      Exit Sub
   End If
   
   Me.Hide
   
   With REL_Garantia
      .txtNumero.Caption = " " & Format(txtCodOS.Text, "000000")
      .rfCodCliente.Caption = " " & txtCodCliente.Text
      .rfModelo.Caption = " " & UCase(cboModelo.Text) & "-" & cboFabricante.Text
      '.frCor.Caption = " " & UCase(cboCor.Text)
      '.frPlaca.Caption = " " & UCase(txtPlaca1.Text) & "-" & txtPlaca2.Text
      '.rfQuilometragem.Caption = " " & txtKM.Text
      
      '.rfQuiloPrimeira.Caption = " " & CInt(txtKM.Text) + CInt(500)
      .rfQuiloSegunda.Caption = " " & .rfQuiloPrimeira.Caption + 1000
      .rfQuiloTerceira.Caption = " " & .rfQuiloSegunda.Caption + 1000
      .rfQuiloQuarta.Caption = " " & .rfQuiloTerceira.Caption + 1000
      
      .Relatorio.NumeroRegistros = 1
      .Relatorio.NomeImpressora = var_Impressora
      .Relatorio.Ativar
   End With
   
   Unload REL_Garantia
   
   Me.Show 1
End Sub

Private Sub menu_Impressao_Orcamento_Click()
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

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

'buscando os dados do formulário
'i = Grid_OS.Row
vCodOS = txtCodOS.Text
codPedido = txtCodPedido.Text

'ver a quantidade de peças e serviços da ordem de serviços
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Recapadora" Then
    vTabelaServicos = "OS_Servicos_recapadora"
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    vTabelaServicos = "OS_Servicos_Auto"
ElseIf vTipoOS = "Comunicação Visual" Then
    vTabelaServicos = "OS_Servicos_Comunicacao"
End If

'somando os produtos
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM pedidos_itens WHERE (cod_pedido = " & codPedido & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
Dim vQuantProduto As Double
Dim vTotalProduto As Currency
vQuantProduto = ValidateNull(r_Itens("VarQuant"))
vTotalProduto = ValidateNull(r_Itens("VarSoma"))
'somando os serviços
sSQL_Itens = "SELECT COUNT(*) as VarQuant, SUM(total) as VarSoma FROM " & vTabelaServicos & " WHERE (cod_os = " & vCodOS & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)
'Debug.Print sSQL_Itens
Dim vQuantServico As Double
Dim vTotalServico As Currency
vQuantServico = ValidateNull(r_Itens("VarQuant"))
vTotalServico = ValidateNull(r_Itens("VarSoma"))
Dim vSomaTotais As Currency
Dim vSomaQuant As Double
vSomaTotais = vTotalProduto + vTotalServico
vSomaQuant = vQuantProduto + vQuantServico

'TOTAIS RESUMIDO
REL_OS_Completo.txtQuantServicos.Caption = " " & Format(vQuantServico, "000")
REL_OS_Completo.txtQuantPecas.Caption = " " & Format(vQuantProduto, "000")
REL_OS_Completo.txtQuantGeral.Caption = " " & Format(vSomaQuant, "000")

REL_OS_Completo.txtTotalServicos.Caption = " " & FormatNumber(vTotalServico, 2)
REL_OS_Completo.txtTotalPecas.Caption = " " & FormatNumber(vTotalProduto, 2)
REL_OS_Completo.txtTotalPecasServicos.Caption = " " & FormatNumber(vSomaTotais, 2)

    
sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & codPedido & ")"
sSQL_Itens = sSQL_Itens & " UNION ALL "
sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & vCodOS & ")"
Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

varImpPDF = False
Me.Hide

'If r_Itens.RecordCount > 16 Then
        sSQL = "SELECT produtos.descricao as var_desc, 'PRODUTO' as vTipo, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") "
        sSQL = sSQL & " UNION ALL "
        sSQL = sSQL & "SELECT descricao as var_desc, 'SERVIÇO' as vTipo, quantidade, preco, subtotal, desconto, total, '0' as vCodProd " & _
                "FROM " & vTabelaServicos & " " & _
                "WHERE (cod_os = " & vCodOS & ") order by vTipo"
        Set r = dbData.OpenRecordset(sSQL)

        Set rPedido = dbData.OpenRecordset("SELECT COD_CLIENTE, DATA_COMPRA FROM pedidos WHERE (COD_PEDIDO  = " & codPedido & ");")
        Set rOS = dbData.OpenRecordset("SELECT COD_FUNCIONARIO, SUBTOTAL, VALOR_DESC, TOTAL, ValorDescReal, COD_RESPONSAVEL, TIPO_DESC FROM OS WHERE (COD_OS = " & vCodOS & ");")
        Set rCliente = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (codigo = " & rPedido("COD_CLIENTE") & " );")
        Set rFunc = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & " );")
        Set rTecnico = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rOS("COD_RESPONSAVEL") & " );")

        Me.Hide
        
        Set REL_OS_Completo.ReportMain1.Recordset = r
        
        REL_OS_Completo.txtDHead.Caption = "ORÇAMENTO DA ORDEM DE SERVIÇO Nº " & vCodOS
        REL_OS_Completo.Mostrar_Parcelas txtCodPedido.Text
        REL_OS_Completo.rfSubTotal.Caption = FormatNumber(rOS("SUBTOTAL"), 2)
        REL_OS_Completo.txtDescontoRS.Caption = FormatNumber(rOS("ValorDescReal"), 2)
        REL_OS_Completo.rfTotal.Caption = FormatNumber(rOS("TOTAL"), 2)
        'REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        
        If rOS("TIPO_DESC") = "R" Then
            REL_OS_Completo.rfDesc.Caption = FormatNumber(0, 2)
        Else
            REL_OS_Completo.rfDesc.Caption = FormatNumber(rOS("VALOR_DESC"), 2)
        End If
        
        'DADOS DO CLIENTE
        REL_OS_Completo.rfCliente.Caption = rCliente("nome")
        REL_OS_Completo.rfData.Caption = Format(rPedido("DATA_COMPRA"), "dd/mm/yy")
        REL_OS_Completo.rfForma.Caption = "ORÇAMENTO"
        REL_OS_Completo.rfFunc.Caption = rFunc("nome")
        REL_OS_Completo.rfTecnico.Caption = rTecnico("nome")
        
        'DADOS DO VEICULO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        ElseIf vTipoOS = "Comunicação Visual" Then
             Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
        End If
        
        'DADOS DO VEICULO/EQUIPAMENTO
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
            REL_OS_Completo.frTitParc.Caption = "VEÍCULO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!ANO) = True, "", rEquip!ANO)
            REL_OS_Completo.txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
            REL_OS_Completo.txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            REL_OS_Completo.frTitParc.Caption = "EQUIPAMENTO"
            REL_OS_Completo.txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
            REL_OS_Completo.txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
            REL_OS_Completo.txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = "Equipamento:"
            REL_OS_Completo.ReportField10.Caption = "Fabricante:"
            REL_OS_Completo.ReportField12.Caption = "Modelo:"
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        ElseIf vTipoOS = "Comunicação Visual" Then
            REL_OS_Completo.frTitParc.Caption = ""
            REL_OS_Completo.txtFabricante.Visible = False
            REL_OS_Completo.txtModelo.Visible = False
            REL_OS_Completo.txtAno.Visible = False
            REL_OS_Completo.txtPlaca.Visible = False
            REL_OS_Completo.txtCor.Visible = False
            REL_OS_Completo.txtKM.Visible = False
            REL_OS_Completo.ReportField15.Visible = False
            REL_OS_Completo.ReportField14.Visible = False
            REL_OS_Completo.ReportField18.Visible = False
            REL_OS_Completo.ReportField2.Caption = ""
            REL_OS_Completo.ReportField10.Caption = ""
            REL_OS_Completo.ReportField12.Caption = ""
            REL_OS_Completo.ReportField15.Caption = ""
            REL_OS_Completo.ReportField14.Caption = ""
            REL_OS_Completo.ReportField18.Caption = ""
        End If
        REL_OS_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_OS_Completo.ReportMain1.Ativar
        Unload REL_OS_Completo
'Else
'    With REL_Pedido_Orcamento
'        REL_Pedido_Orcamento.loadPedidos Grid_OS.TextMatrix(i, 6), "OFICINA"
'    End With
'    Unload REL_Pedido_Orcamento
'End If

Me.Show

'parte anterior
'If txtCodOS.Text = "" Then Exit Sub
'vCodOS = txtCodOS.Text

'ver a quantidade de peças e serviços da ordem de serviços
'If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
'    vTabelaServicos = "OS_Servicos_Auto"
'ElseIf vTipoOS = "Recapadora" Then
'    vTabelaServicos = "OS_Servicos_recapadora"
'ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
'    vTabelaServicos = "OS_Servicos_Auto"
'ElseIf vTipoOS = "Comunicação Visual" Then
'    vTabelaServicos = "OS_Servicos_Comunicacao"
'End If

'sSQL_Itens = "SELECT pedidos_itens.codigo FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ")"
'sSQL_Itens = sSQL_Itens & " UNION "
'sSQL_Itens = sSQL_Itens & "SELECT codigo FROM " & vTabelaServicos & " WHERE (cod_os = " & txtCodOS.Text & ")"
'Set r_Itens = dbData.OpenRecordset(sSQL_Itens)

'Me.Hide
'If r_Itens.RecordCount > 16 Then
'   With REL_Pedido_Orcamento_Grande
'        .txtQuantServicos.Caption = " " & Format(txtQuantServicos.Text, "000")
'        .txtQuantPecas.Caption = " " & Format(txtQuantPecas.Text, "000")
'        .txtQuantGeral.Caption = " " & Format(txtQuantGeral.Text, "000")
'
'        .txtTotalServicos.Caption = " " & FormatNumber(txtTotalServicos.Text, 2)
'        .txtTotalPecas.Caption = " " & FormatNumber(txtTotalPecas.Text, 2)
'        .txtTotalPecasServicos.Caption = " " & FormatNumber(txtTotalPecasServicos.Text, 2)
'    End With
'
'    REL_Pedido_Orcamento_Grande.loadPedidos txtCodPedido.Text, "OFICINA"
'    Unload REL_Pedido_Orcamento_Grande
'Else
'    With REL_Pedido_Orcamento
'        .txtQuantServicos.Caption = " " & Format(txtQuantServicos.Text, "000")
'        .txtQuantPecas.Caption = " " & Format(txtQuantPecas.Text, "000")
'        .txtQuantGeral.Caption = " " & Format(txtQuantGeral.Text, "000")
'
'        .txtTotalServicos.Caption = " " & FormatNumber(txtTotalServicos.Text, 2)
'        .txtTotalPecas.Caption = " " & FormatNumber(txtTotalPecas.Text, 2)
'        .txtTotalPecasServicos.Caption = " " & FormatNumber(txtTotalPecasServicos.Text, 2)
'    End With

'    REL_Pedido_Orcamento.loadPedidos txtCodPedido.Text, "OFICINA"
'    'Unload REL_Pedido_Orcamento
'    Unload REL_Pedido_Orcamento
'End If
'Me.Show
End Sub

Private Sub menu_Impressao_Pedido_Click()
cmdImpPedido1_Click
End Sub

Private Sub mskDataSaida_GotFocus()
   SelectControl mskDataSaida
End Sub

Private Sub mskDataSaida_KeyPress(KeyAscii As Integer)
   mskDataSaida.Mask = "##/##/##"
End Sub

Private Sub mskDataSaida_LostFocus()
   If mskDataSaida.Text = "" Or mskDataSaida.Text = "__/__/__" Then
      mskDataSaida.Mask = ""
      mskDataSaida.Text = ""
   Else
      If Not IsDate(mskDataSaida.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskDataSaida.SetFocus
      End If
   End If
End Sub

Private Sub mskHoraSaida_GotFocus()
   SelectControl mskHoraSaida
End Sub

Private Sub mskHoraSaida_KeyPress(KeyAscii As Integer)
   mskHoraSaida.Mask = "##:##"
End Sub

Private Sub mskHoraSaida_LostFocus()
   If mskHoraSaida.Text = "" Or mskHoraSaida.Text = "__:__" Then
      mskHoraSaida.Mask = ""
      mskHoraSaida.Text = ""
   Else
      If Not IsDate(mskHoraSaida.Text) Then
         ShowMsg "HORA INVÁLIDA!" & vbCrLf & "A hora digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskHoraSaida.SetFocus
      End If
   End If
End Sub

Private Sub mskValorServicoAuto_Change()
CalcularTotalServicoAuto
End Sub

Private Sub mskInicio_GotFocus()
Calcular_Prazo
SelectControl mskInicio
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
   If Not IsDate(mskInicio.Text) Then Exit Sub
   mskInicio.Mask = "##/##/##"
End Sub

Private Sub mskTermino_Change()
If Not IsDate(mskTermino.Text) Then Exit Sub
mskTermino.Mask = "##/##/##"
End Sub

Private Sub mskTermino_GotFocus()
SelectControl mskTermino
End Sub
Private Sub mskValorServicoAuto_GotFocus()
SelectControl mskValorServicoAuto
End Sub


Private Sub mskValorServicoAuto_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
   KeyAscii = 0
   cmdAdicionarServicosAuto_Click
End If
End Sub


Private Sub mskValorServicoAuto_LostFocus()
If mskValorServicoAuto.Text = "" Then
   mskValorServicoAuto.Text = Format(0, ocMONEY)
Else
   mskValorServicoAuto.Text = Format(mskValorServicoAuto, ocMONEY)
End If
End Sub



Private Sub optFinanceiroAberto_Click()
MostrarGrid_OS_Situacao
LimparGrid_Situacao
If optFinanceiroAberto.Value = True Then cmdExcluir.Enabled = True
End Sub
Private Sub optFinanceiroFechado_Click()
MostrarGrid_OS_Situacao
LimparGrid_Situacao
If optFinanceiroFechado.Value = True Then cmdExcluir.Enabled = False
End Sub

Private Sub optGarantia_Click()
MostrarGrid_OS_Situacao
LimparGrid_Situacao
End Sub

Private Sub optOrcamento_Click()
MostrarGrid_OS_Situacao
LimparGrid_Situacao
End Sub

Private Sub optServico_Click()
MostrarGrid_OS_Situacao
LimparGrid_Situacao
End Sub

Private Sub optTodos_Click()
MostrarGrid_OS_Situacao
LimparGrid_Situacao
End Sub


Private Sub stProdSer_Click(PreviousTab As Integer)
If stProdSer.Tab = 0 Then
    If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
        If cboServicosAuto.Enabled = True Then cboServicosAuto.SetFocus
    ElseIf vTipoOS = "Recapadora" Then
        If cboTipo.Enabled = True Then cboTipo.SetFocus
    End If
ElseIf stProdSer.Tab = 1 Then
    If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
End If
End Sub

Private Sub txtAcresc_Change()
   'On Error GoTo Erro
   
   If txtAcresc.Text = "" Or txtSubtotal.Text = "" Then
      txtAcresc.Text = FormatNumber(0, 2)
      SelectControl txtAcresc
      Exit Sub
   End If
   
   Calcular_Desconto
   Exit Sub
   
'Erro:
'   ShowMsg "O valor digitado é inválido!", vbExclamation
'   txtAcresc.Text = 0
End Sub





Private Sub txtAcrescDinheiro_GotFocus()
SelectControl txtAcrescDinheiro
End Sub


Private Sub txtAcrescDinheiro_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
    txtAcrescDinheiro_LostFocus
End If
End Sub


Private Sub txtAcrescDinheiro_LostFocus()
Dim ValueTotal As Double
Dim ValueDiscount As Double
Dim Percent As Double

ValueTotal = txtSubtotal.Text

If txtAcrescDinheiro.Text = "" Then
    ValueDiscount = 0
Else
    ValueDiscount = txtAcrescDinheiro.Text
End If

Percent = (ValueDiscount / ValueTotal) * 100
txtAcresc.Text = FormatNumber(Percent, 2)
txtAcrescDinheiro.Visible = False
txtAcresc.Visible = True
txtAcresc_LostFocus
txtRecebido.SetFocus
End Sub
Private Sub txtAcresc_GotFocus()
   SelectControl txtAcresc
End Sub

Private Sub txtAcresc_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
   
If KeyAscii = 13 Then
      txtAcrescDinheiro.Visible = True
      txtAcresc.Visible = False
      txtAcrescDinheiro.Text = ""
      txtAcrescDinheiro.SetFocus
End If
End Sub

Private Sub txtAcresc_LostFocus()
On Error GoTo erro
   
If txtAcresc.Text = "" Or txtSubtotal.Text = "" Then
   txtAcresc.Text = FormatNumber(0, 2)
   SelectControl txtAcresc
   Exit Sub
End If

Calcular_Desconto
txtAcresc.Text = FormatNumber(txtAcresc.Text, 2)
Exit Sub
   
erro:
   MsgBox "O valor digitado é inválido!", vbExclamation, "Aviso do Sistema"
   txtAcresc.Text = 0
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtObsServ_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtRecebido_Change()
Calcular_Troco
End Sub

Private Sub txtRecebido_GotFocus()
   SelectControl txtRecebido
End Sub

Private Sub txtRecebido_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
      cmdFinalizar_Click
End If
End Sub

Private Sub txtRecebido_LostFocus()
If txtRecebido.Text = "" Then
    txtRecebido.Text = Format(0, ocMONEY)
    txtTroco.Text = Format(0, ocMONEY)
Else
    txtRecebido.Text = Format(txtRecebido.Text, ocMONEY)
End If
Calcular_Troco
End Sub
Private Sub txtAno_GotFocus()
txtAno.SelStart = 0
txtAno.SelLength = Len(txtAno)
End Sub


Private Sub txtCodBarra_GotFocus()
SelectControl txtCodBarra
End Sub


Private Sub txtCodBarra_LostFocus()
If vTipoConsPecas = 0 Or vTipoConsPecas = 1 Then
    If txtCodBarra.Text = "" Then
        vTipoConsPecas = 0
        txtCodPeca.Text = ""
        cboPecas.Locked = False
        txtValorPeca.Text = Format(0, ocMONEY)
        txtQuantPeca.Text = "1"
        txtTotalPeca.Text = Format(0, ocMONEY)
        Exit Sub
    End If
        
        If txtCodBarra.Text <> "" Then txtCodBarra.Text = Format(txtCodBarra.Text, "00000")
        sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, tamanho, REF, fabricante, quant_estoque, unid_medida, CFOP, NCM, ICMSCST, ICMSAliq, EAN  FROM produtos WHERE (COD_BARRA = '" & txtCodBarra.Text & "') AND (ativo = 1);"
        Set r = dbData.OpenRecordset(sSQL)
        
        If Not r.BOF Then
           txtCodPeca.Text = r("var_codprod")
           
           'If tipoEmpresa = 4 Then
           '    cboPecas.Text = ValidateNull(r("var_desc")) & " /  " & ValidateNull(r("tamanho")) & " / " & ValidateNull(r("fabricante")) & " /  " & r("REF")
           '    'cboPecas2.Text = ValidateNull(r("var_desc"))
           'Else
              cboPecas.Text = ValidateNull(r("var_desc"))
'          '    txtICMS.Text = Format(ValidateNull(r("ICMSAliq")), "##,##0.00")
           'End If
           
            vTipoConsPecas = 1
            cboPecas.Locked = True
        Else
           ShowMsg "Produto Inexistente!", vbCritical
           vTipoConsPecas = 0
           txtCodBarra.Text = ""
           txtCodBarra.SetFocus
           Exit Sub
        End If
        
        MostrarValorVenda
        txtQuantPeca.SetFocus
    'End If
End If

On Local Error Resume Next

End Sub
Private Sub MostrarValorVenda()
Dim vrVenda As Currency
If txtCodPeca.Text = "" Then Exit Sub

'mostrar o ultimo preço de compra
sSQL = "SELECT TOP 1 VALOR_VV FROM Produtos_Precos WHERE (COD_PRODUTO = " & txtCodPeca & ") ORDER BY codigo DESC;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then vrVenda = r("VALOR_VV")
If r.State <> 0 Then r.Close
Set r = Nothing

txtValorPeca.Text = Format(vrVenda, ocMONEY)
txtQuantPeca.Text = "1"
End Sub



Private Sub txtCodFuncAP_Change()
If txtCodFuncAP.Text = "" Then Exit Sub
txtFuncAP.Text = ""

sSQL = "SELECT codigo, nome, sobrenome FROM funcionario WHERE (codigo = " & txtCodFuncAP.Text & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtFuncAP.Text = r("nome") & " " & r("sobrenome")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub txtCodFuncAP_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtCodServicoAuto_Change()
If txtCodServicoAuto.Text = "" Then mskValorServicoAuto.Text = Format(0, ocMONEY): Exit Sub

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    sSQL = "SELECT * FROM os_Servicos WHERE (codigo = " & txtCodServicoAuto.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then mskValorServicoAuto.Text = Format(r("valor"), ocMONEY)
    If Not r.BOF Then vServico = r("servico")
ElseIf vTipoOS = "Recapadora" Then
    sSQL = "SELECT * FROM os_Servicos WHERE (codigo = " & txtCodServicoAuto.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then mskValorServicoAuto.Text = Format(r("valor"), ocMONEY)
    
    If Not r.BOF Then vMedida = ValidateNull(r("medida"))
    If Not r.BOF Then vAro = ValidateNull(r("aro"))
    If Not r.BOF Then vBanda = ValidateNull(r("banda"))
    If Not r.BOF Then vServico = r("servico")
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    sSQL = "SELECT * FROM os_Servicos WHERE (codigo = " & txtCodServicoAuto.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then mskValorServicoAuto.Text = Format(r("valor"), ocMONEY)
    If Not r.BOF Then vServico = r("servico")
ElseIf vTipoOS = "Comunicação Visual" Then
    sSQL = "SELECT * FROM os_Servicos WHERE (codigo = " & txtCodServicoAuto.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then mskValorServicoAuto.Text = Format(r("valor"), ocMONEY)
    If Not r.BOF Then vServico = r("servico")
End If
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub txtDesc_Change()
On Error GoTo erro
   
If txtDesc.Text = "" Or txtSubtotal.Text = "" Then
   txtDesc.Text = FormatNumber(0, 2)
   SelectControl txtDesc
   Exit Sub
End If

Calcular_Desconto
Exit Sub
   
erro:
   ShowMsg "O valor digitado é inválido!", vbExclamation
   txtDesc.Text = 0
End Sub

Private Sub txtDesc_GotFocus()
   SelectControl txtDesc
End Sub




Private Sub txtDescDinheiro_GotFocus()
SelectControl txtDescDinheiro
End Sub

Private Sub txtDescDinheiro_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
    txtDescDinheiro_LostFocus
End If
End Sub

Private Sub txtDescDinheiro_LostFocus()
Dim ValueTotal As Double
Dim ValueDiscount As Double
Dim Percent As Double

ValueTotal = txtSubtotal.Text

If txtDescDinheiro.Text = "" Then
    ValueDiscount = 0
Else
    ValueDiscount = txtDescDinheiro.Text
End If

Percent = (ValueDiscount / ValueTotal) * 100
txtDesc.Text = FormatNumber(Percent, 2)
txtDescDinheiro.Visible = False
txtDesc.Visible = True
txtDesc_LostFocus
txtRecebido.SetFocus
End Sub


Private Sub txtDesc_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)

If KeyAscii = 13 Then
      txtDescDinheiro.Visible = True
      txtDesc.Visible = False
      txtDescDinheiro.Text = ""
      txtDescDinheiro.SetFocus
End If
End Sub

Private Sub txtDesc_LostFocus()
'On Error GoTo erro
Dim vDesc As Double
   
If txtDesc.Text = "" Or txtSubtotal.Text = "" Then
    txtDesc.Text = FormatNumber(0, 2)
    vDesc = 0
Else
    vDesc = txtDesc.Text
End If


If vTipoDesc = "1" Then     'desconto manual
    If vLimitarDesc = 1 Then
        If cboTipoPgto.Text = "À VISTA" Then
            If vDesc > vValorDescFixoAV Then
            MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
            txtDesc.Text = FormatNumber(0, 2)
            End If
        Else
            If vDesc > vValorDescFixoAP Then
            MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
            txtDesc.Text = FormatNumber(0, 2)
            End If
        End If
    Else
    End If
ElseIf vTipoDesc = "2" Then 'desconto fixo
    If vLimitarDesc = 1 Then
        If cboTipoPgto.Text = "À VISTA" Then
            If vDesc > vValorDescFixoAV Then
            MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
            txtDesc.Text = FormatNumber(vValorDescFixoAV, 2)
            End If
        Else
            If vDesc > vValorDescFixoAP Then
            MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
            txtDesc.Text = FormatNumber(vValorDescFixoAP, 2)
            End If
        End If
    Else
    End If
ElseIf vTipoDesc = "3" Then     'desconto gradativo
    Dim vValorDescGradual As Currency
    If cboTipoPgto.Text = "À VISTA" Then
        If vEtapa = 1 Then
            vValorDescGradual = vValorDescGradualAV1
        ElseIf vEtapa = 2 Then
            vValorDescGradual = vValorDescGradualAV2
        ElseIf vEtapa = 3 Then
            vValorDescGradual = vValorDescGradualAV3
        End If
    Else
        If vEtapa = 1 Then
            vValorDescGradual = vValorDescGradualAP1
        ElseIf vEtapa = 2 Then
            vValorDescGradual = vValorDescGradualAP2
        ElseIf vEtapa = 3 Then
            vValorDescGradual = vValorDescGradualAP3
        End If
    End If

    If vLimitarDesc = 1 Then
        If vDesc > vValorDescGradual Then
        MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
        txtDesc.Text = FormatNumber(vValorDescGradual, 2)
        End If
    End If
End If
   
Calcular_Desconto

txtDesc.Text = FormatNumber(txtDesc.Text, 2)
Exit Sub
   
'erro:
'   ShowMsg "O valor digitado é inválido!", vbExclamation
'   txtDesc.Text = 0
End Sub

Private Sub txtDescPecas_GotFocus()
SelectControl txtDescPecas
End Sub

Private Sub txtDescPecas_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtDescPecas_LostFocus()
CalcularValorPeca
End Sub


Private Sub txtDescServicoAuto_GotFocus()
SelectControl txtDescServicoAuto
End Sub

Private Sub txtDescServicoAuto_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtDescServicoAuto_LostFocus()
CalcularTotalServicoAuto

If txtDescServicoAuto.Text = "" Then
   txtDescServicoAuto.Text = Format(0, ocMONEY)
Else
   txtDescServicoAuto.Text = Format(txtDescServicoAuto, ocMONEY)
End If
End Sub
Private Sub txtEntrada_Change()
   txtEntrada_Click
End Sub
Private Sub txtEntrada_Click()
   If txtTotalPecasServicos.Text = "" Then
      Exit Sub
   Else
      Mostrar_ValorRestante
      Calcular_Parcelas
      Calcular_Prazo
   End If
End Sub

Private Sub txtEntrada_GotFocus()
SelectControl txtEntrada
End Sub

Private Sub txtEntrada_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtEntrada_LostFocus()
   txtEntrada_Click
   If txtEntrada = "" Then txtEntrada = Format(0, ocMONEY) Else txtEntrada = Format(txtEntrada, ocMONEY)
End Sub
Private Sub txtEntrada_Validate(Cancel As Boolean)
If txtEntrada.Text = "" Then txtEntrada.Text = "0,00"
End Sub

Private Sub cboPecas_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboPecas_LostFocus()
If vTipoConsPecas <> 1 Then
    If vTipoConsPecas = 0 Or vTipoConsPecas = 2 Then
       'txtCodBarra.Text = ""
       
       'If cboPecas.Text = "" Then txtCodPeca.Text = "": Exit Sub
       'If cboPecas.ListIndex = -1 Then txtCodPeca.Text = "": Exit Sub
       
        If cboPecas.Text = "" Then
            'txtEAN.Text = ""
            txtCodPeca.Text = ""
            vTipoConsPecas = 0
            txtCodBarra.Locked = False
            txtCodBarra.Text = ""
            'txtUnid.Text = ""
            'txtCFOP.Text = ""
            'txtCST.Text = ""
            'txtNCM.Text = ""
            'txtICMS.Text = ""
            txtValorPeca.Text = "0"
            txtQuantPeca.Text = "0"
            txtTotalPeca.Text = "0"
            Exit Sub
        End If
        
        If cboPecas.ListIndex = -1 Then txtCodPeca.Text = "": vTipoConsPecas = 0: txtCodBarra.Locked = False: cboPecas.Text = "": txtCodBarra.Text = "": Exit Sub
    
       txtCodPeca = cboPecas.ItemData(cboPecas.ListIndex)
       
        If txtCodPeca.Text = "" Then
            vTipoConsPecas = 0
            txtCodBarra.Locked = False
            cboPecas.Text = ""
            txtCodBarra.Text = ""
            Exit Sub
        End If
       
       sSQL = "SELECT codigo, descricao, EAN, COD_BARRA, unid_medida  FROM produtos WHERE (codigo = " & txtCodPeca.Text & ");"
       Set r = dbData.OpenRecordset(sSQL)
       
        If Not r.BOF Then
            txtCodBarra.Text = r("COD_BARRA")
            vTipoConsPecas = 2
            txtCodBarra.Locked = True
            MostrarValorVenda
            CalcularValorPeca
            If txtQuantPeca.Enabled = True Then txtQuantPeca.SetFocus
        ElseIf r.BOF Then
            ShowMsg "Produto não cadastrado.", vbExclamation
            vTipoConsPecas = 0
            cboPecas.Text = ""
            txtCodBarra.Text = ""
            txtValorPeca.Text = "0"
            txtQuantPeca.Text = "0"
            txtTotalPeca.Text = "0"
            txtCodBarra.Locked = False
            If r.State <> 0 Then r.Close
        End If
    End If
End If
End Sub


Private Sub cboPecas_Validate(Cancel As Boolean)
''lstBusca.Visible = False
   
'Dim ItemLst As ListItem
'Dim fGrid As Object
'Dim bCancel As Boolean
'Dim vProd() As String
'Dim rPos As RECT
'Dim lLft As Long, lTop As Long

'Dim cCfg As ConfigItem
'Dim tipoEmpresa As Integer

'Set cCfg = sysConfig("TIPO_EMPRESA")
'tipoEmpresa = cCfg.Value
'Set cCfg = Nothing

'If cboPecas.Text = "" Then Exit Sub
'If cboPecas.Text <> "" And txtCodPeca.Text <> "" Then Exit Sub

'If cboPecas.Text <> "" And txtCodPeca.Text = "" Then
'   DoEvents
'   'lblInfoBusca.Visible = True
'   'lblInfoBusca.Refresh
'   Screen.MousePointer = vbHourglass
   
'   'Otimizando a conslta
'   sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
      "produtos.quant_estoque AS var_quant, (SELECT  TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
      "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
      "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
      "produtos_entrada.data_entrada DESC, produtos_entrada.hora_entrada) AS venda " & _
      "FROM produtos WHERE (descricao LIKE '%" & cboPecas.Text & "%') AND (produtos.ativo = 1) " & _
      "ORDER BY descricao;"
'      Debug.Print sSQL
   
'   Set r = dbData.OpenRecordset(sSQL)
'End If
   
'   GetWindowRect cboPecas.hwnd, rPos
'   lLft = rPos.Left * Screen.TwipsPerPixelX - 160
'   lTop = rPos.Top * Screen.TwipsPerPixelY + cboPecas.Height
   
'   If tipoEmpresa = 5 Then
'      Set fGrid = New BuscaGrid_Automotivo
'   Else
'      'Set fGrid = New BuscaGrid_Comum
'   End If
   
'   Load fGrid
'   LockWindowUpdate fGrid.lstBusca.hwnd
   
'If cboPecas.Text <> "" Then
'   If Not r Is Nothing Then
'      Do While Not r.EOF
'         'primeira coluna
'         Set ItemLst = fGrid.lstBusca.ListItems.Add(, , r("var_cod"))
'         'segunda e terceira coluna, que são sub itens da coluna 1
'         ItemLst.SubItems(1) = r("var_codbarra")
'         ItemLst.SubItems(2) = ValidateNull(r("var_desc")) & " /  " & ValidateNull(r("var_fab"))
      
'      If tipoEmpresa = 5 Then
'         If Not IsNull(r("var_quant")) Then ItemLst.SubItems(4) = r("var_quant")
'         If Not IsNull(r("venda")) Then ItemLst.SubItems(5) = Format(r("venda"), ocMONEY)
         
'            'Compartibilidade
'            Dim sSQL_Comp As String
'            Dim var_Comp As String
'            Dim rS2 As ADODB.Recordset
            
'            sSQL_Comp = "Select MODELO, ANO From PRODUTOS_COMP Where COD_PRODUTO = " & r("var_cod")
'            Set rS2 = dbData.OpenRecordset(sSQL_Comp)
            
'            Do While Not rS2.EOF
'            var_Comp = var_Comp & rS2!Modelo & "(" & rS2!Ano & "),  "
'            rS2.MoveNext
'            Loop
            
'            If Not IsNull(var_Comp) Then ItemLst.SubItems(3) = var_Comp
'            var_Comp = ""
'      Else
'         If Not IsNull(r("var_quant")) Then ItemLst.SubItems(3) = r("var_quant")
'         If Not IsNull(r("venda")) Then ItemLst.SubItems(4) = Format(r("venda"), ocMONEY)
'      End If
         
'         r.MoveNext
'      Loop
      
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
'   End If
'End If

'   'lblInfoBusca.Visible = False
'   Screen.MousePointer = vbDefault
   
'   LockWindowUpdate 0
'   fGrid.Move lLft, lTop
'   fGrid.Show vbModal
   
'   bCancel = fGrid.Cancelled
'   vProd = fGrid.InfoProduct
   
'   Unload fGrid
'   Set fGrid = Nothing
   
'   If Not bCancel Then
'     If tipoEmpresa = 5 Then
'         txtCodPeca.Text = vProd(1)      'lstBusca.SelectedItem
'         cboPecas.Text = vProd(3)        'lstBusca.SelectedItem.ListSubItems.Item(1).Text
'         txtValorPeca.Text = vProd(5)    'lstBusca.SelectedItem.ListSubItems.Item(2).Text
'      Else
'         txtCodPeca.Text = vProd(1)      'lstBusca.SelectedItem
'         cboPecas.Text = vProd(3)        'lstBusca.SelectedItem.ListSubItems.Item(1).Text
'         txtValorPeca.Text = vProd(4)    'lstBusca.SelectedItem.ListSubItems.Item(2).Text
'      End If
'      Cancel = True
'      'GoTo ValidarBusca
'   End If
End Sub

Private Sub txtPareceCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtPlaca_GotFocus()
txtPlaca.SelStart = 0
txtPlaca.SelLength = Len(txtPlaca)
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtQuantPeca_Change()
CalcularValorPeca
End Sub

Private Sub txtQuantPeca_GotFocus()
SelectControl txtQuantPeca
CalcularValorPeca
End Sub

Private Sub txtQuantServicoAuto_GotFocus()
SelectControl txtQuantServicoAuto
End Sub

Private Sub txtQuantServicoAuto_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
   If txtQuantServicoAuto.Text = "" Then txtQuantServicoAuto.Text = 1
   cmdAdicionarServicosAuto_Click
End If
End Sub
Private Sub txtQuantServicoAuto_LostFocus()
CalcularTotalServicoAuto
End Sub


Private Sub Calcular_DescontoAP()
   If txtSubtotal.Text = "" Or txtSubtotal.Text = "0,00" Then Exit Sub
   If txtDesc.Text = "" Then txtDesc.Text = FormatNumber(0, 2)
   If txtAcresc.Text = "" Then txtAcresc.Text = FormatNumber(0, 2)
   
   If txtDesc.Text <> "0,00" And txtAcresc.Text = "0,00" Then
      If optDescRS.Value = True Then
         txtTotalDesc.Text = FormatNumber(CCur(txtSubtotal.Text) - CCur(txtDesc.Text), 2)
      ElseIf optDescPorc.Value = True Then
         txtTotalDesc.Text = FormatNumber(CCur(txtSubtotal.Text) - ((CCur(txtSubtotal.Text) * CCur(txtDesc.Text)) / 100), 2)
      End If
      
   ElseIf txtAcresc.Text <> "0,00" And txtDesc.Text = "0,00" Then
      If optAscrescRS.Value = True Then
         txtTotalDesc.Text = FormatNumber(CCur(txtSubtotal.Text) + CCur(txtAcresc.Text), 2)
      ElseIf optAscrescPorc.Value = True Then
         txtTotalDesc.Text = FormatNumber(CCur(txtSubtotal.Text) + ((CCur(txtSubtotal.Text) * CCur(txtAcresc.Text)) / 100), 2)
      End If
      
   Else
      txtTotalDesc.Text = FormatNumber(txtSubtotal.Text, 2)
      'optDescRS.Value = True
      'optAscrescRS.Value = True
   End If
   
   Mostrar_ValorRestante
End Sub


Private Sub Calcular_Parcelas()
Dim var_ValorRest As Currency
Dim QUANT As Integer
Dim RESULTADO As Currency

If txtTotalDesc.Text = "0,00" Or txtValorRest.Text = "0,00" Or cboQuantParc.Text = "" Then Exit Sub

var_ValorRest = txtValorRest.Text
QUANT = cboQuantParc.Text
RESULTADO = CCur(var_ValorRest / QUANT)
txtValorParc = Format(RESULTADO, ocMONEY)
End Sub

Private Sub Calcular_Troco()
Dim VAR_GERAL As Currency, VAR_RECEBIDO As Currency, var_Troco As Currency

If txtTotalDesc.Text = "" Or txtRecebido.Text = "" Then Exit Sub

If txtRecebido.Text = "0,00" Or txtRecebido.Text = "" Then
   txtTroco.Text = Format(0, ocMONEY)
Else
   VAR_GERAL = txtTotalDesc.Text
   VAR_RECEBIDO = txtRecebido.Text
    If VAR_RECEBIDO > VAR_GERAL Then
        var_Troco = VAR_RECEBIDO - VAR_GERAL
    Else
        var_Troco = 0
    End If
   txtTroco.Text = Format(var_Troco, ocMONEY)
End If
End Sub
Public Function aNumeros(ByVal KeyAscii As Integer, Optional Virgula As Boolean = False, Optional Ponto As Boolean = False) As Integer
   'FUNÇÃO PARA PERMITIR NUMEROS, VIRGULAS E PONTO
   Select Case KeyAscii
      Case IIf(Virgula = True, 44, 0), IIf(Ponto = True, 46, 0), 8, 13, 48 To 57
         aNumeros = KeyAscii
      Case Else
         aNumeros = 0
   End Select
End Function

Private Sub txtSubTotal_Validate(Cancel As Boolean)
Calcular_Desconto
End Sub

Private Sub Calcular_Desconto()
'CALCULAR O VALOR DAS PARCELAS
If txtSubtotal.Text = "" Or txtSubtotal.Text = "0,00" Then Exit Sub
If txtDesc.Text = "" Then txtDesc.Text = FormatNumber(0, 2)
If txtAcresc.Text = "" Then txtAcresc.Text = FormatNumber(0, 2)

Dim varValorSubTotalDebito As Currency
Dim varValorSubTotalCredito As Currency
Dim varSubTotalBruto As Currency
Dim varSubTotalLiquido As Currency

varSubTotalBruto = txtSubtotal.Text

    
   If txtDesc.Text <> "0,00" And txtAcresc.Text = "0,00" Then     'com desconto sem acrescimo
      
      If optDescRS.Value = True Then
         txtTotalDesc.Text = Format(CCur(txtSubtotal.Text) - CCur(txtDesc.Text), ocMONEY)
      ElseIf optDescPorc.Value = True Then
         'txtTotalDesc.Text = Format(CCur(txtSubTotal.Text) - ((CCur(txtSubTotal.Text) * CDbl(txtDesc.Text)) / 100), ocMONEY)
         txtTotalDesc.Text = Format(CCur(txtSubtotal.Text) - Round(((CCur(txtSubtotal.Text) * CCur(txtDesc.Text)) / 100), 2), ocMONEY)
      End If
     
   
   ElseIf txtAcresc.Text <> "0,00" And txtDesc.Text = "0,00" Then    'sem desconto com acrescim0
      If optAscrescRS.Value = True Then
         txtTotalDesc.Text = Format(CCur(txtSubtotal.Text) + CCur(txtAcresc.Text), ocMONEY)
      ElseIf optAscrescPorc.Value = True Then
         txtTotalDesc.Text = Format(CCur(txtSubtotal.Text) + ((CCur(txtSubtotal.Text) * CDbl(txtAcresc.Text)) / 100), ocMONEY)
      End If
      
   Else
      txtTotalDesc.Text = Format(txtSubtotal.Text, ocMONEY)
   End If
'End If
   
   Mostrar_ValorRestante

If optDescRS.Value = True Then
   MostrarDescItem
Else
    txtDescItens.Text = "0"
End If

Calcular_Troco
End Sub
Private Sub MostrarDescItem()
Dim varValorDescProc As Double

If txtTotalDesc.Text = "" Then Exit Sub
If txtSubtotal.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency

B = txtTotalDesc.Text
a = txtSubtotal.Text

varValorDescProc = ((B - a) / a) * 100
txtDescItens.Text = Abs(FormatNumber(varValorDescProc, 6))
End Sub


Private Sub txtSubTotal_GotFocus()
   SelectControl txtSubtotal
End Sub

Private Sub txtSubTotal_LostFocus()
txtSubtotal.Text = Format(txtSubtotal, ocMONEY)
End Sub

Function ChecarLimite() As Boolean
Dim Limite As Currency
Dim Total As Currency
Dim LimiteAtual As Currency

sSQL = "SELECT * FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then Limite = r("limite_credito")
If r.State <> 0 Then r.Close
Set r = Nothing

If Limite = 0 Then
   ChecarLimite = True
   Exit Function
End If

Total = 0
sSQL = "SELECT os.cod_cliente, SUM(os.total) AS total FROM parcelas INNER JOIN os ON parcelas.codigo = os.codigo WHERE (os.cod_cliente = " & txtCodCliente.Text & ") AND (parcelas.status = 0) GROUP BY os.cod_cliente;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then Total = r("total")
If r.State <> 0 Then r.Close
Set r = Nothing

LimiteAtual = Limite - Total

If Left(LimiteAtual, 1) = "-" Then
   LimiteAtual = Mid(LimiteAtual, 2, Len(LimiteAtual))
End If

If LimiteAtual < (CCur(txtTotalPecasServicos.Text) - CCur(txtEntrada.Text)) Then
   ShowMsg "O CLIENTE POSSUE UM TOTAL DE R$ " & FormatNumber(Total, 2) & " EM COMPRAS NÃO PAGAS E O VALOR DA COMPRA É DE R$ " & FormatNumber(txtTotalPecasServicos.Text, 2) & " E O SALDO DELE É R$ " & FormatNumber(Limite - Total), vbExclamation
   ChecarLimite = False
Else
   ChecarLimite = True
End If
End Function

Private Sub mskDataEntrada_GotFocus()
   SelectControl mskDataEntrada
End Sub

Private Sub mskDataEntrada_KeyPress(KeyAscii As Integer)
   mskDataEntrada.Mask = "##/##/##"
End Sub

Private Sub mskDataEntrada_LostFocus()
   If mskDataEntrada.Text = "" Or mskDataEntrada.Text = "__/__/__" Then
      mskDataEntrada.Mask = ""
      mskDataEntrada.Text = ""
   Else
      If Not IsDate(mskDataEntrada.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskDataEntrada.SetFocus
      End If
   End If
End Sub

Private Sub mskHoraEntrada_GotFocus()
   SelectControl mskHoraEntrada
End Sub

Private Sub mskHoraEntrada_KeyPress(KeyAscii As Integer)
   mskHoraEntrada.Mask = "##:##"
End Sub

Private Sub mskHoraEntrada_LostFocus()
   If mskHoraEntrada.Text = "" Or mskHoraEntrada.Text = "__:__" Then
      mskHoraEntrada.Mask = ""
      mskHoraEntrada.Text = ""
Else
    If Not IsDate(mskHoraEntrada.Text) Then
        MsgBox "HORA INVÁLIDA!" & vbCrLf & "A hora digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation, "Aviso do Sistema"
        mskHoraEntrada.SetFocus
    End If
End If
End Sub

Private Sub txtSubtotalPecas_GotFocus()
SelectControl txtSubtotalPecas
End Sub


Private Sub txtSubTotalServicoAuto_GotFocus()
SelectControl txtQuantServicoAuto
End Sub


Private Sub txtSubTotalServicoAuto_LostFocus()
If txtSubTotalServicoAuto.Text = "" Then
   txtSubTotalServicoAuto.Text = Format(0, ocMONEY)
Else
   txtSubTotalServicoAuto.Text = Format(txtSubTotalServicoAuto, ocMONEY)
End If
End Sub


Private Sub txtTotalPeca_GotFocus()
SelectControl txtTotalPeca
End Sub


Private Sub txtTotalPecas_Change()
Somar_Totais
End Sub

Private Sub txtTotalServicoAuto_GotFocus()
SelectControl txtQuantServicoAuto
End Sub


Private Sub txtTotalServicos_Change()
Somar_Totais
End Sub

Private Sub txtValorPeca_GotFocus()
SelectControl txtValorPeca
End Sub

Private Sub txtValorParc_GotFocus()
If txtTotalPecasServicos.Text = "" Then
   Exit Sub
Else
   Mostrar_ValorRestante
End If

SelectControl txtValorParc
End Sub

Private Sub txtValorParc_LostFocus()
   If txtValorParc = "" Then txtValorParc = Format(0, ocMONEY) Else txtValorParc = Format(txtValorParc, ocMONEY)
End Sub
Private Sub txtValorPeca_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdAdicionarPecas_Click
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
   If cmdGerarEntrada.Enabled = True Then mskDataEntrada.SetFocus
ElseIf SSTab1.Tab = 1 Then
   ''If frmServico.Enabled = True Then cboServicos.SetFocus
ElseIf SSTab1.Tab = 2 Then
ElseIf SSTab1.Tab = 3 Then
'      cboStatus.SetFocus
ElseIf SSTab1.Tab = 4 Then
'      optAV.SetFocus
ElseIf SSTab1.Tab = 5 Then
'      optStatusTodos.SetFocus
End If
End Sub

Private Sub TxtCodCliente_Change()
If txtCodCliente.Text = "" Then Exit Sub

If cmdAlterar.Enabled = True Then
   sSQL = "SELECT codigo, nome, celular FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cboCliente.Text = r("nome") & IIf(Trim(ValidateNull(r("celular"))) = "", "", "     (" & Right$(ValidateNull(r("celular")), 9) & ")")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub

Private Sub txtCodFuncionario_Change()
If txtCodFuncionario.Text = "" Then Exit Sub
If txtCodFuncionario.Text = 0 Then Exit Sub

'txtCodFunc.Text = txtCodFuncionario.Text
'txtCodFuncAP.Text = txtCodFuncionario.Text

If cmdAlterar.Enabled = True Then
   sSQL = "SELECT * FROM funcionario WHERE (codigo = " & txtCodFuncionario.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cboFuncionario.Text = r("nome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub

Private Sub txtCodMecanico_Change()
If txtCodMecanico.Text = "" Then Exit Sub

If cmdAlterar.Enabled = True Then
   sSQL = "SELECT * FROM funcionario WHERE (codigo = " & txtCodMecanico.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cboMecanico.Text = r("nome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub

Private Sub txtCodOS_Change()
If txtCodOS.Text = "" Then
   'imgCancelar.Enabled = False
   cmdGerarEntrada.Enabled = False
   cmdFinalizarAV.Visible = False
   cmdFinalizarAP.Visible = False
   Exit Sub
Else
   'imgCancelar.Enabled = True
   cmdGerarEntrada.Enabled = True
End If

LimparObjetos_Entrada

sSQL = "SELECT * FROM os WHERE (cod_os = " & txtCodOS.Text & ");"
Set rs = dbData.OpenRecordset(sSQL)

Mostrar_Entrada rs

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
    MostrarEquipamentoAuto
    MostrarGrid_Acessorios
    MostrarGrid_Situacao
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    MostrarEquipamento
    MostrarGrid_Acessorios
    MostrarGrid_Situacao
ElseIf vTipoOS = "Recapadora" Then
    'MostrarEquipamento
End If

'CHECAR SE A OS ESTÁ FECHADA & PAGA
Verificar_OS_FechadaePaga

If OS_FINANCEIROABERTO = True Then
    If cboStatus.Text = "TERMINADO" Then
        frmSecundario.Enabled = True
        cmdApagar.Enabled = True
        cmdAlterar.Enabled = True
        cmdFinalizarAV.Visible = True
        cmdFinalizarAP.Visible = True
        cmdImpPedido2.Enabled = False
        menu_Impressao_Pedido.Enabled = False
    Else
        cmdImpPedido2.Enabled = False
        menu_Impressao_Pedido.Enabled = False
   End If
Else
    frmSecundario.Enabled = False
    cmdApagar.Enabled = False
    cmdAlterar.Enabled = False
    cmdFinalizarAV.Visible = False
    cmdFinalizarAP.Visible = False
    frmVendaFechamento.Visible = False
    cmdImpPedido2.Enabled = True
    menu_Impressao_Pedido.Enabled = True
   
    'pegar a forma de pagamento da os
    sSQL = "SELECT TIPO_PAGAMENTO FROM os WHERE (cod_os = " & txtCodOS.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)

    If Not r.BOF Then
        If r("TIPO_PAGAMENTO") = "À Prazo" Then
            frmVendaFechamento.Visible = True
            cmdFinalizar.Enabled = False
            cmdCancelar.Enabled = False
        Else
            frmVendaFechamento.Visible = True
            cmdFinalizar.Enabled = False
            cmdFinalizar.Enabled = False
        End If
    End If
    
    If r.State <> 0 Then r.Close
    Set r = Nothing

End If
  
If txtCodOS.Text <> "" Then
   'lblCarro1a.Caption = cboFabricante.Text & " /  " & cboModelo.Text & " /  " & cboModelo.Text
   'lblCarro2a.Caption = cboCliente.Text
End If

If cboStatus.Text = "À COMEÇAR" Then
    If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
        frmAcessorios.Visible = True
        frmSituacao.Visible = True
        frmParecerCliente.Visible = True
        'frmServicos.Visible = False
        frmAcessorios.Visible = True
        frmSituacao.Visible = True
        frmGridServicos.Visible = False
    ElseIf vTipoOS = "Recapadora" Then
        frmAcessorios.Visible = True
        frmSituacao.Visible = True
        frmParecerCliente.Visible = True
        'frmServicos.Visible = False
        frmAcessorios.Visible = False
        frmSituacao.Visible = False
        frmGridServicos.Visible = False
    ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
        frmAcessorios.Visible = True
        frmSituacao.Visible = True
        frmParecerCliente.Visible = True
        'frmServicos.Visible = False
        frmAcessorios.Visible = True
        frmSituacao.Visible = True
        frmGridServicos.Visible = False
    ElseIf vTipoOS = "Comunicação Visual" Then
        'frmServicos.Visible = False
        frmAcessorios.Visible = False
        frmSituacao.Visible = False
        frmParecerCliente.Visible = False
        stProdSer.Visible = True
        frmGridServicos.Visible = True
        frmTotaisGeral.Visible = True
        frmTotaisProdServ.Visible = True
    End If
    
    cmdImpEntrada2.Enabled = False
    cmdImpOrcamento2.Enabled = False
    cmdImpPedido2.Enabled = False
Else
    frmAcessorios.Visible = False
    frmSituacao.Visible = False
    frmParecerCliente.Visible = False
    frmGridServicos.Visible = True
    frmTotaisGeral.Visible = True
    frmTotaisProdServ.Visible = True
    stProdSer.Visible = True

    If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
        'frmServicos.Visible = True
    ElseIf vTipoOS = "Recapadora" Then
        'frmServicos.Visible = True
    ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
        'frmServicos.Visible = True
    ElseIf vTipoOS = "Comunicação Visual" Then
        'frmServicos.Visible = True
    End If
    
    cmdImpEntrada2.Enabled = True
    cmdImpOrcamento2.Enabled = True
    cmdImpPedido2.Enabled = True
End If

MostrarGrid_Servicos
Somar_Totais

cmdGerarEntrada.Enabled = False
cmdNovo.Enabled = True
If cboStatus.Enabled = True Then cboStatus.SetFocus
End Sub

Private Sub cboPecas_GotFocus()
If vTipoConsPecas <> 1 Then
    moCombo.AttachTo cboPecas
    
    If vTipoConsPecas = 0 Or vTipoConsPecas = 2 Then
    
        If cboPecas.ListIndex = -1 Then
            cboPecas.Clear
            If Len(cboPecas.Text) > 3 Then
             sSQL = "SELECT DISTINCT descricao, codigo FROM produtos where quant_estoque > 0 ORDER BY descricao;"
             'sSQL = "SELECT DISTINCT descricao, codigo FROM produtos WHERE (descricao LIKE '%" & cboPecas.Text & "%') ORDER BY descricao;"
             Set r = dbData.OpenRecordset(sSQL)
            
             Do While Not r.EOF
                cboPecas.AddItem ValidateNull(r("descricao"))
                 cboPecas.ItemData(cboPecas.NewIndex) = r("codigo")
                r.MoveNext
             Loop
            End If
        End If
    End If
    moCombo.AttachTo cboPecas
End If
End Sub

Private Sub cboPecas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then OS_Consulta_Pecas.Show 1
End Sub

Private Sub txtQuantPeca_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
   If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1
   cmdAdicionarPecas_Click
End If
End Sub

Private Sub txtQuantPeca_LostFocus()
CalcularValorPeca
End Sub





Private Sub txtTotalDesc_Change()
   txtTotalPecasServicos.Text = Format(txtTotalDesc.Text, "##,##0.00")
End Sub

Private Sub txtValorRest_Change()
   Calcular_Parcelas
End Sub
Private Sub txtValorRest_GotFocus()
SelectControl txtValorRest
End Sub
Private Sub cboformaPgto_Change()
Calcular_Desconto
End Sub

Private Sub cboformaPgto_Click()
Calcular_Desconto
End Sub


Private Sub cboFormaPgto_GotFocus()
Dim varTexto As String
varTexto = cboformaPgto.Text
    cboformaPgto.Clear
    Preencher_FormaPgto
cboformaPgto.Text = varTexto
SelectControl cboformaPgto
moCombo.AttachTo cboformaPgto
End Sub
Private Sub cboQuantParc_Change()
Calcular_Parcelas
Calcular_Prazo
End Sub
Private Sub Preencher_FormaPgto()
If cboTipoPgto.Text = "À VISTA" Then
    cboformaPgto.AddItem "1 - DINHEIRO"
    cboformaPgto.AddItem "3 - CARTÃO - DÉBITO"
    cboformaPgto.AddItem "4 - CARTÃO - CRÉDITO"
    cboformaPgto.AddItem "5 - CHEQUE"
    cboformaPgto.AddItem "7 - TRANSFERÊNCIA"
    cboformaPgto.AddItem "8 - DEPOSITO"
    cboformaPgto.AddItem "9 - FINANCEIRA"
    cboformaPgto.AddItem "10 - PIX"
Else
    cboformaPgto.AddItem "2 - PROMISSÓRIA"
    cboformaPgto.AddItem "5 - CHEQUE"
    cboformaPgto.AddItem "6 - BOLETO"
End If
End Sub

Private Sub cboQuantParc_Click()
   Calcular_Parcelas
   Calcular_Prazo
End Sub

Private Sub cboQuantParc_GotFocus()
Dim i As Integer
Dim varTexto As String
varTexto = cboQuantParc.Text

   cboQuantParc.Clear
   For i = 1 To 12
      cboQuantParc.AddItem i
   Next
cboQuantParc.Text = varTexto
SelectControl cboQuantParc
moCombo.AttachTo cboQuantParc
End Sub
Private Sub cboPrazo_GotFocus()
Dim varTexto As String
varTexto = cboPrazo.Text
   cboPrazo.Clear
   cboPrazo.AddItem "5"
   cboPrazo.AddItem "10"
   cboPrazo.AddItem "15"
   cboPrazo.AddItem "20"
   cboPrazo.AddItem "30"
cboPrazo.Text = varTexto
SelectControl cboPrazo
   moCombo.AttachTo cboPrazo
End Sub
Private Sub optDescPorc_Click()
Calcular_Desconto
If frmVendaFechamento.Visible = True Then
    If txtDesc.Enabled = True Then
        txtDesc.SetFocus
    Else
        txtRecebido.SetFocus
    End If
End If
End Sub

Private Sub optDescRS_Click()
Calcular_Desconto
If txtDesc.Enabled = True Then txtDesc.SetFocus
End Sub
