VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Empresa_Cadastro 
   Caption         =   "LICENÇA"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   Icon            =   "Empresa_Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmBancario 
      Caption         =   "Dados Bancário"
      Height          =   1395
      Left            =   4620
      TabIndex        =   96
      Top             =   4080
      Width           =   6915
      Begin VB.TextBox txtPix 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         TabIndex        =   39
         Top             =   960
         Width           =   5775
      End
      Begin VB.ComboBox cboTipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   540
         TabIndex        =   37
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3180
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtFavorecido 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   38
         Top             =   600
         Width           =   3915
      End
      Begin VB.TextBox txtConta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         TabIndex        =   36
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox txtBanco 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   660
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chave Pix"
         Height          =   195
         Left            =   120
         TabIndex        =   102
         Top             =   1005
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   101
         Top             =   645
         Width           =   315
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agência:"
         Height          =   195
         Left            =   2460
         TabIndex        =   100
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Favorecido"
         Height          =   195
         Left            =   1980
         TabIndex        =   99
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta: "
         Height          =   195
         Left            =   4260
         TabIndex        =   98
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   285
         Width           =   465
      End
   End
   Begin VB.Frame frmPagamento 
      Caption         =   "Pagamentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4620
      TabIndex        =   76
      Top             =   5580
      Width           =   6915
      Begin VB.TextBox txtCodDesbloqueio 
         Height          =   315
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   82
         Top             =   1740
         Width           =   1155
      End
      Begin VB.CheckBox chkProximo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Próximo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2100
         TabIndex        =   81
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtDia 
         Height          =   315
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   77
         Top             =   240
         Width           =   915
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   1095
         Left            =   120
         TabIndex        =   79
         Top             =   600
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   1931
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdGerarPagamentos 
         Height          =   315
         Left            =   3060
         TabIndex        =   80
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Gerar Parcela"
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
         MICON           =   "Empresa_Cadastro.frx":23D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdDesbroquear 
         Height          =   315
         Left            =   3900
         TabIndex        =   84
         Top             =   1740
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Desbloqueio"
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
         MICON           =   "Empresa_Cadastro.frx":23EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdDesbTemp 
         Height          =   315
         Left            =   5340
         TabIndex        =   87
         Top             =   1740
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Temporário"
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
         MICON           =   "Empresa_Cadastro.frx":240A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Desbroqueio:"
         Height          =   195
         Left            =   120
         TabIndex        =   83
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dia de Pgto:"
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   285
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   11385
      TabIndex        =   73
      Top             =   60
      Width           =   11415
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   8940
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   480
         Picture         =   "Empresa_Cadastro.frx":2426
         Top             =   60
         Width           =   615
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LICENÇA"
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
         Left            =   1320
         TabIndex        =   75
         Top             =   240
         Width           =   1380
      End
   End
   Begin VB.Frame FraConfiguraçãoFiscal 
      Caption         =   "Configuração Fiscal"
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
      Left            =   4620
      TabIndex        =   58
      Top             =   1080
      Width           =   6915
      Begin VB.ComboBox cboDIFAL 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   30
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox chkOffline 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Offline NFCe"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2280
         TabIndex        =   31
         Top             =   2640
         Width           =   1275
      End
      Begin VB.TextBox txtRegime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4920
         MaxLength       =   1
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox chkContigenciaNFCe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contigência NFCe"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   33
         Top             =   2640
         Width           =   1635
      End
      Begin VB.CheckBox chkContigenciaNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contigência NFe"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   32
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cboPerfil 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6060
         TabIndex        =   28
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtAliqUF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         TabIndex        =   85
         Top             =   2580
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdGerarLicenca 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Left            =   6480
         TabIndex        =   25
         Top             =   1440
         Width           =   315
      End
      Begin VB.TextBox txtAmbienteNF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtpICMSSN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3660
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtNFCeCSC 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   27
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtNFCeIDToken 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   26
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtLicencaDLL 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Top             =   1440
         Width           =   5355
      End
      Begin VB.CommandButton cmdCertificadoDigital 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Left            =   6480
         TabIndex        =   23
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtCertificadoDigital 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   1080
         Width           =   5355
      End
      Begin VB.CommandButton cmdDiretorioXML 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Left            =   6480
         TabIndex        =   21
         Top             =   720
         Width           =   315
      End
      Begin VB.TextBox txtDiretorioXML 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   5355
      End
      Begin VB.TextBox txtCRT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3840
         MaxLength       =   1
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtCodigoIBGE 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   16
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblREGIME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGIME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6300
         TabIndex        =   104
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblDifal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IPI Compõe Difal"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   103
         Top             =   2205
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regime"
         Height          =   195
         Left            =   4320
         TabIndex        =   95
         Top             =   405
         Width           =   540
      End
      Begin VB.Label lblCRTNome 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6540
         TabIndex        =   94
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil"
         Height          =   195
         Left            =   5640
         TabIndex        =   90
         Top             =   1845
         Width           =   345
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Aliq UF"
         Height          =   195
         Left            =   60
         TabIndex        =   86
         Top             =   2640
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblAmbienteStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ambiente"
         Height          =   195
         Left            =   2220
         TabIndex        =   68
         Top             =   530
         Width           =   660
      End
      Begin VB.Label lblAmbiente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ambiente:"
         Height          =   195
         Left            =   2220
         TabIndex        =   67
         Top             =   350
         Width           =   705
      End
      Begin VB.Label lblAliqSN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Cred. ICMS SN"
         Height          =   195
         Left            =   2340
         TabIndex        =   66
         Top             =   2220
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblCSC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CSC"
         Height          =   195
         Left            =   2040
         TabIndex        =   65
         Top             =   1845
         Width           =   315
      End
      Begin VB.Label lblIDToken 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Token"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   1845
         Width           =   675
      End
      Begin VB.Label lblLicençaDLL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licença DLL"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   1485
         Width           =   915
      End
      Begin VB.Label lblCertDigital 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cert. Digital"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   1125
         Width           =   810
      End
      Begin VB.Label lblDirXML 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dir. XML"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   765
         Width           =   615
      End
      Begin VB.Label lblCRT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRT"
         Height          =   195
         Left            =   3480
         TabIndex        =   60
         Top             =   400
         Width           =   330
      End
      Begin VB.Label lblCódigoIBGE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código IBGE"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   400
         Width           =   915
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ChamaleonBtn.chameleonButton cmdSalvar 
      Height          =   375
      Left            =   60
      TabIndex        =   40
      Top             =   7380
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Salvar"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Empresa_Cadastro.frx":3130
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
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   60
      ScaleHeight     =   6225
      ScaleWidth      =   4425
      TabIndex        =   44
      Top             =   1080
      Width           =   4455
      Begin VB.TextBox txtIEMunicipal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1440
         Width           =   1635
      End
      Begin VB.TextBox txtFantasia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton optJuridico 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jurídica"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2520
         TabIndex        =   93
         Top             =   30
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optFisica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Física"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   92
         Top             =   30
         Width           =   795
      End
      Begin ChamaleonBtn.chameleonButton cmdCopiar 
         Height          =   285
         Left            =   3180
         TabIndex        =   88
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "C"
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
         MICON           =   "Empresa_Cadastro.frx":314C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsultarIE 
         Height          =   285
         Left            =   2940
         TabIndex        =   89
         Top             =   1140
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "..."
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
         MICON           =   "Empresa_Cadastro.frx":3168
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtCodUF 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2100
         TabIndex        =   72
         Top             =   2640
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtCodCid 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3840
         TabIndex        =   71
         Top             =   2640
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox cboCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   2970
         Width           =   2955
      End
      Begin VB.ComboBox cboEstado 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   2640
         Width           =   795
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   6
         Top             =   2040
         Width           =   675
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2340
         Width           =   2895
      End
      Begin VB.TextBox txtCaminho 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   200
         TabIndex        =   14
         Top             =   4500
         Width           =   2895
      End
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   1320
         ScaleHeight     =   1065
         ScaleWidth      =   2865
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4800
         Width           =   2895
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   13
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox txtIE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1140
         Width           =   1635
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1740
         Width           =   2895
      End
      Begin VB.TextBox txtRazao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   1
         Top             =   540
         Width           =   2895
      End
      Begin MSMask.MaskEdBox mskCep 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   3900
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCNPJ 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTelefone 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   3300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCelular 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdSped 
         Height          =   285
         Left            =   3420
         TabIndex        =   91
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "S"
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
         MICON           =   "Empresa_Cadastro.frx":3184
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insc. Municipal:"
         Height          =   195
         Left            =   180
         TabIndex        =   105
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num.:"
         Height          =   195
         Left            =   840
         TabIndex        =   70
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   780
         TabIndex        =   69
         Top             =   2340
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Celular:"
         Height          =   195
         Left            =   720
         TabIndex        =   57
         Top             =   3600
         Width           =   525
      End
      Begin VB.Label lblProcurar 
         AutoSize        =   -1  'True
         Caption         =   "[ &PROCURAR ]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   3120
         TabIndex        =   41
         Top             =   5940
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Logomarca:"
         Height          =   285
         Left            =   420
         TabIndex        =   55
         Top             =   4500
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         Height          =   285
         Left            =   900
         TabIndex        =   54
         Top             =   3900
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   285
         Left            =   780
         TabIndex        =   53
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         Height          =   285
         Left            =   720
         TabIndex        =   52
         Top             =   2640
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fixo:"
         Height          =   195
         Left            =   900
         TabIndex        =   51
         Top             =   3300
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   285
         Left            =   540
         TabIndex        =   50
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   285
         Left            =   705
         TabIndex        =   49
         Top             =   2970
         Width           =   540
      End
      Begin VB.Label lblCNPJ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ:"
         Height          =   285
         Left            =   825
         TabIndex        =   48
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insc. Estadual:"
         Height          =   195
         Left            =   225
         TabIndex        =   47
         Top             =   1140
         Width           =   1050
      End
      Begin VB.Label lblRazao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razão:"
         Height          =   285
         Left            =   765
         TabIndex        =   46
         Top             =   540
         Width           =   510
      End
      Begin VB.Label lblFantazia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fantasia:"
         Height          =   285
         Left            =   630
         TabIndex        =   45
         Top             =   240
         Width           =   645
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdAlterar 
      Height          =   375
      Left            =   1320
      TabIndex        =   42
      Top             =   7380
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Alterar"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Empresa_Cadastro.frx":31A0
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
      Left            =   2580
      TabIndex        =   43
      Top             =   7380
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Excluir"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Empresa_Cadastro.frx":31BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   56
      Top             =   7830
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "18:29"
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
Attribute VB_Name = "Empresa_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim lNovoCod As Long
Dim sSQL As String
Dim r As ADODB.Recordset

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hwndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

Private moCombo As cComboHelper
Private Caminho As String

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Dim certificadoDataVencimento As String

Private Sub GerarCodDesbloqueio()


End Sub

Private Function Inserir_Pagamentos() As Boolean
'Dim sSQL As String

'If txtCodigoIBGE.Text = "" Then txtCodigoIBGE.Text = 0
'If txtCRT.Text = "" Then txtCRT.Text = 0
'If txtAmbienteNF.Text = "" Then txtAmbienteNF.Text = 2
'If txtpICMSSN.Text = "" Then txtpICMSSN.Text = 0

'Comando de inclusão
'sSQL = "INSERT INTO empresa (" & _
   "fantasia, razao, cnpj, ie, endereco, cidade, estado, telefone, celular, cep, email, caminho, CodigoIBGE, CRT, DiretorioXML, CertificadoDigital, NFCeIDToken, NFCeCSC, LicencaDLL, BAIRRO, AmbienteNF, pCreditoICMSSimplesNacional, numero, Perfil) VALUES ('" & _
   txtFantasia.Text & "', '" & txtRazao.Text & "', '" & mskCNPJ.Text & "', '" & txtIE.Text & "', '" & _
   txtEndereco.Text & "', '" & cboCidade.Text & "', '" & cboEstado.Text & "', '" & mskTelefone.Text & "', '" & mskCelular.Text & "','" & _
   mskCep.Text & "', '" & txtEmail.Text & "', '" & txtCaminho.Text & "', " & txtCodigoIBGE.Text & ", " & txtCRT.Text & ", '" & txtDiretorioXML.Text & "', '" & txtCertificadoDigital.Text & "', '" & LPad(txtNFCeIDToken.Text, 6, "0") & "', '" & txtNFCeCSC.Text & "', '" & txtLicencaDLL.Text & "', '" & txtBairro.Text & "', " & txtAmbienteNF.Text & ", " & FSQL(txtpICMSSN.Text, 2) & ", " & txtNum.Text & ", '" & cboperfil.Text & "')"

'Retorna o resultado da atualização
'Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Inserir_Dados() As Boolean
Dim sSQL As String, bOpt As Boolean

If cboDIFAL.Text = "SIM" Then
    bOpt = True
ElseIf cboDIFAL.Text = "NÃO" Then
    bOpt = False
End If

If txtCodigoIBGE.Text = "" Then txtCodigoIBGE.Text = 0
If txtCRT.Text = "" Then txtCRT.Text = 0
If txtAmbienteNF.Text = "" Then txtAmbienteNF.Text = 2
If txtpICMSSN.Text = "" Then txtpICMSSN.Text = 0

'Comando de inclusão
sSQL = "INSERT INTO empresa (" & _
   "fantasia, razao, cnpj, ie, endereco, cidade, estado, telefone, celular, cep, email, caminho, CodigoIBGE, CRT, DiretorioXML, CertificadoDigital, NFCeIDToken, NFCeCSC, LicencaDLL, BAIRRO, AmbienteNF, pCreditoICMSSimplesNacional, numero, Perfil, pAliqUF, UltimoNSU, ContigenciaNFe, ContigenciaNFCe, Banco, Agencia, Conta, Tipo, Favorecido, Pix, WhatsAppApiKey, IPICompoeDIFAL, VencimentoCert, RegimeTributario, NFCeOffline, IEMunicipal ) VALUES ('" & _
   txtFantasia.Text & "', '" & txtRazao.Text & "', '" & mskCNPJ.Text & "', '" & txtIE.Text & "', '" & _
   txtEndereco.Text & "', '" & cboCidade.Text & "', '" & cboEstado.Text & "', '" & mskTelefone.Text & "', '" & mskCelular.Text & "','" & _
   mskCep.Text & "', '" & txtEmail.Text & "', '" & txtCaminho.Text & "', " & txtCodigoIBGE.Text & ", " & txtCRT.Text & ", '" & txtDiretorioXML.Text & "', '" & txtCertificadoDigital.Text & "', '" & LPad(txtNFCeIDToken.Text, 6, "0") & "', '" & txtNFCeCSC.Text & "', '" & txtLicencaDLL.Text & "', '" & txtBairro.Text & "', " & txtAmbienteNF.Text & ", " & FSQL(txtpICMSSN.Text, 2) & ", " & txtNum.Text & ", '" & cboPerfil.Text & "', " & FSQL(txtAliqUF.Text, 2) & ", 0, " & Abs(chkContigenciaNFe.Value) & ", " & Abs(chkContigenciaNFCe.Value) & ", '" & txtBanco.Text & "', '" & txtAgencia.Text & "', '" & txtConta.Text & "', '" & cboTipo.Text & "', '" & txtFavorecido.Text & "', '" & txtPix.Text & "', 0, '" & Abs(bOpt) & "', null, " & txtRegime.Text & ", " & Abs(chkOffline.Value) & ", '" & txtIEMunicipal.Text & "' )"

'Retorna o resultado da atualização
Inserir_Dados = dbData.Execute(sSQL)

End Function

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String, bOpt As Boolean

    If cboDIFAL.Text = "SIM" Then
        bOpt = True
    ElseIf cboDIFAL.Text = "NÃO" Then
        bOpt = False
    End If
   
   If txtCodigoIBGE.Text = "" Then txtCodigoIBGE.Text = 0
   If txtCRT.Text = "" Then txtCRT.Text = 0
   If txtAmbienteNF.Text = "" Then txtAmbienteNF.Text = 2
   If txtpICMSSN.Text = "" Then txtpICMSSN.Text = 0
   If txtCRT.Text = "1" Or txtCRT.Text = "2" Then txtRegime.Text = "1" Else txtRegime.Text = 3
   
   'Comando de atualização
   sSQL = "UPDATE empresa SET " & _
      "fantasia = '" & txtFantasia.Text & "', " & _
      "razao = '" & txtRazao.Text & "', " & _
      "cnpj = '" & mskCNPJ.Text & "', " & _
      "ie = '" & txtIE.Text & "', " & _
      "endereco = '" & txtEndereco.Text & "', " & _
      "bairro = '" & txtBairro.Text & "', " & _
      "cidade = '" & cboCidade.Text & "', " & _
      "estado = '" & cboEstado.Text & "', " & _
      "telefone = '" & mskTelefone.Text & "', " & _
      "celular = '" & mskCelular.Text & "', " & _
      "cep = '" & mskCep.Text & "', " & _
      "email = '" & txtEmail.Text & "', " & _
      "caminho = '" & txtCaminho.Text & "', " & _
      "CodigoIBGE = " & txtCodigoIBGE.Text & ", " & _
      "CRT = " & txtCRT.Text & ", " & _
      "DiretorioXML = '" & txtDiretorioXML.Text & "', " & _
      "CertificadoDigital = '" & txtCertificadoDigital.Text & "', " & _
      "NFCeIDToken = '" & LPad(txtNFCeIDToken.Text, 6, "0") & "', " & _
      "NFCeCSC = '" & txtNFCeCSC.Text & "', " & _
      "LicencaDLL = '" & txtLicencaDLL.Text & "', " & _
      "AmbienteNF = " & txtAmbienteNF.Text & ", " & _
      "pCreditoICMSSimplesNacional = " & FSQL(txtpICMSSN.Text, 2) & ", pAliqUF = " & FSQL(txtAliqUF.Text, 2) & ", " & _
      "Numero = " & txtNum.Text & ", ContigenciaNFe = " & Abs(chkContigenciaNFe.Value) & ", ContigenciaNFCe = " & Abs(chkContigenciaNFCe.Value) & ", Perfil = '" & cboPerfil.Text & "', Banco = '" & txtBanco.Text & "', Agencia = '" & txtAgencia.Text & "', Conta = '" & txtConta.Text & "', Tipo = '" & cboTipo.Text & "', Favorecido = '" & txtFavorecido.Text & "' , " & _
      "Pix = '" & txtPix.Text & "',  IPICompoeDIFAL = '" & Abs(bOpt) & "', RegimeTributario = " & txtRegime.Text & ", NFCeOffline = " & Abs(chkOffline.Value) & ", IEMunicipal = '" & txtIEMunicipal.Text & "'"
      
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Function REGIME(sREGIME As Integer) As String
Select Case sREGIME
    Case 1
        REGIME$ = "Simples Nacional"
    Case 2
        REGIME$ = "Simples Nacional - Excesso de Sublimite"
    Case 3
        REGIME$ = "Lucro Presumido"
    Case 4
        REGIME$ = "Lucro Real"
    Case 5
        REGIME$ = "Lucro Arbitrado"
End Select
End Function
Private Function CRT(sCRT As Integer) As String
Select Case sCRT
    Case 1
        CRT$ = "Simples Nacional"
    Case 2
        CRT$ = "Simples Nacional"
    Case 3
        CRT$ = "Regime Normal"
    Case 4
        CRT$ = "Microempreendedor Individual (MEI)"
End Select
End Function
Private Sub Limpar_Objetos()
txtFantasia.Text = ""
txtRazao.Text = ""
mskCNPJ.Mask = ""
mskCNPJ.Text = ""
txtIE.Text = ""
txtEndereco.Text = ""
txtBairro.Text = ""
cboCidade.Text = ""
cboEstado.Text = ""
mskTelefone.Mask = ""
mskTelefone.Text = ""
mskCelular.Mask = ""
mskCelular.Text = ""
mskCep.Mask = ""
mskCep.Text = ""
txtEmail.Text = ""
txtCaminho.Text = ""
txtCodigoIBGE.Text = ""
txtCRT.Text = ""
txtCertificadoDigital.Text = ""
txtDiretorioXML.Text = ""
txtNFCeIDToken.Text = ""
txtNFCeCSC.Text = ""
txtLicencaDLL.Text = ""
txtAmbienteNF.Text = ""
txtpICMSSN.Text = "1,11"
cboPerfil.Text = ""
txtRegime.Text = ""
cboDIFAL.Text = "NÃO"
chkOffline.Value = False
chkContigenciaNFe.Value = False
chkContigenciaNFCe.Value = False
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
   Dim nroReg As Long
   
   If Not rTabela Is Nothing Then
      txtFantasia.Text = ValidateNull(rTabela("fantasia"))
      txtRazao.Text = rTabela("razao")
      mskCNPJ.Text = rTabela("cnpj")
      txtIE.Text = ValidateNull(rTabela("ie"))
      txtIEMunicipal.Text = ValidateNull(rTabela("IEMunicipal"))
      txtEndereco.Text = rTabela("endereco")
      txtBairro.Text = rTabela("bairro")
      cboCidade.Text = rTabela("cidade")
      cboEstado.Text = rTabela("estado")
      mskTelefone.Text = ValidateNull(rTabela("telefone"))
      mskCelular.Text = ValidateNull(rTabela("celular"))
      mskCep.Text = rTabela("cep")
      txtEmail.Text = ValidateNull(rTabela("email"))
      txtCaminho.Text = ValidateNull(rTabela("caminho"))
      txtCodigoIBGE.Text = rTabela("CodigoIBGE")
      txtCRT.Text = rTabela("CRT")
      txtDiretorioXML.Text = rTabela("DiretorioXML")
      txtCertificadoDigital.Text = rTabela("CertificadoDigital")
      txtNFCeIDToken.Text = rTabela("NFCeIDToken")
      txtNFCeCSC.Text = rTabela("NFCeCSC")
      txtLicencaDLL.Text = rTabela("LicencaDLL")
      txtAmbienteNF.Text = rTabela("AmbienteNF")
      txtpICMSSN.Text = Format(rTabela("pCreditoICMSSimplesNacional"), "#0.00")
      txtAliqUF.Text = Format(rTabela("pAliqUF"), "#0.00")
      txtNum.Text = rTabela("Numero")
      cboPerfil.Text = rTabela("Perfil")
      nroReg = rTabela.RecordCount
      chkOffline = Abs(rTabela("NFCeOffline"))
      chkContigenciaNFe = Abs(rTabela("ContigenciaNFe"))
      chkContigenciaNFCe = Abs(rTabela("ContigenciaNFCe"))
      txtBanco.Text = ValidateNull(rTabela("Banco"))
      txtAgencia.Text = ValidateNull(rTabela("Agencia"))
      txtConta.Text = ValidateNull(rTabela("Conta"))
      cboTipo.Text = ValidateNull(rTabela("Tipo"))
      txtFavorecido.Text = ValidateNull(rTabela("Favorecido"))
      txtPix.Text = ValidateNull(rTabela("Pix"))
    txtRegime = ValidateNull(rTabela("RegimeTributario"))
    cboDIFAL.Text = IIf(ValidateNull(rTabela("IPICompoeDIFAL")) = 1, "SIM", "NÃO")
    If txtCRT.Text = 1 Or txtCRT.Text = 2 Then
        lblAliqSN.Visible = True
        txtpICMSSN.Visible = True
        lblDifal.Enabled = False
        cboDIFAL.Enabled = False
    Else
        lblAliqSN.Visible = False
        txtpICMSSN.Visible = False
        cboDIFAL.Enabled = True
        lblDifal.Enabled = True
    End If

   Else
      nroReg = 0
   End If
   
   On Local Error Resume Next
   If Not IsNull(rTabela("caminho")) Then Set picLogo.Picture = LoadPicture(rTabela("caminho"))
   lblCRTNome.Caption = CRT(txtCRT.Text)
   If Not Vazio(txtAmbienteNF.Text) Then
      If Val(txtAmbienteNF.Text) = 1 Then lblAmbienteStatus.Caption = "Produção"
      If Val(txtAmbienteNF.Text) = 2 Then lblAmbienteStatus.Caption = "Homologação"
   End If
   
   cmdSalvar.Enabled = Not (nroReg >= 1)
   cmdAlterar.Enabled = (nroReg >= 1)
   cmdExcluir.Enabled = (nroReg >= 1)
End Sub

Private Sub Mostrar_Pagamentos()
sSQL = "SELECT *, (CASE WHEN bloqueio = 1 THEN 'SIM' ELSE 'NÃO' END) AS vTextBloq, (CASE WHEN pago = 1 THEN 'SIM' ELSE 'NÃO' END) AS vTextPago FROM  licenca_pagamentos ORDER BY data_vencimento desc;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer, x As Integer

With Grid
   .Clear
   .Cols = 8
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 1000
   .ColWidth(3) = 1000
   .ColWidth(4) = 1000
   .ColWidth(5) = 1000
   .ColWidth(6) = 900
   .ColWidth(7) = 900

   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "REF."
   .TextMatrix(0, 3) = "VENC."
   .TextMatrix(0, 4) = "BLOQ."
   .TextMatrix(0, 5) = "LIBER."
   .TextMatrix(0, 6) = "BLOQ."
   .TextMatrix(0, 7) = "PAGO"
   
   .Redraw = False
   
   i = 1
   'codigo, dia_vencimento, mes_ref, data_vencimento, data_bloqueio, data_liberacao, bloqueio, pago
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.rows - 1, 2) = rTabela("mes_ref")
         .TextMatrix(.rows - 1, 3) = rTabela("data_vencimento")
         .TextMatrix(.rows - 1, 4) = rTabela("data_bloqueio")
         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("data_liberacao"))
         .TextMatrix(.rows - 1, 6) = rTabela("vTextbloq")
         .TextMatrix(.rows - 1, 7) = rTabela("vTextPago")
         rTabela.MoveNext
         
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
   'MUDAR COR DE FONTE DA COLUNA
   'For i = 1 To .Rows - 1
   '   .Row = i
   '   .Col = 3
   '   .CellForeColor = &HC0&
   '   .CellFontBold = True
   'Next
   
   .rows = .rows - 1
   .Redraw = True
End With
End Sub


Private Sub MostrarTipoEmpresa()
If optJuridico.Value = True Then
    lblCNPJ.Caption = "CNPJ:"
    lblRazao.Caption = "Razão:"
Else
    lblCNPJ.Caption = "CPF:"
    lblRazao.Caption = "Nome:"
End If
End Sub

Private Sub Verificar_Registro()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM empresa"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.RecordCount = 1 Then
      Mostrar_Dados r
      MsgBox "Não é possivel cadastrar mais de uma empresa", vbInformation, "Aviso do Sistema"
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cboCidade_Click()
cboCidade_LostFocus
End Sub

Private Sub cboCidade_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodUF.Text <> "" Then
    'Limpa a lista
    cboCidade.Clear
    
    If txtCodUF.Text = "" Then Exit Sub
    
    sSQL = "SELECT DISTINCT NOME, IdEstado, ID FROM CIDADE WHERE (IdEstado = " & txtCodUF.Text & ") ORDER BY NOME"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboCidade.AddItem ValidateNull(r("NOME"))
       cboCidade.ItemData(cboCidade.NewIndex) = r("ID")
       r.MoveNext
    Loop
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If

moCombo.AttachTo cboCidade
End Sub

Private Sub cboCidade_LostFocus()
On Error GoTo TrataErro

If txtCodUF.Text <> "" Then
    If cboCidade.Text = "" Then txtCodCid.Text = "": txtCodigoIBGE.Text = "": Exit Sub
    If cboCidade.ListIndex = -1 Then txtCodCid.Text = "": Exit Sub
    
    txtCodCid = cboCidade.ItemData(cboCidade.ListIndex)
End If

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboDIFAL_GotFocus()
cboDIFAL.Clear
cboDIFAL.AddItem "NÃO"
cboDIFAL.AddItem "SIM"
moCombo.AttachTo cboDIFAL
End Sub


Private Sub cboEstado_Click()
cboEstado_LostFocus
End Sub

Private Sub cboEstado_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim varTexto As String

'varTexto = cboEstado.Text

'Limpa a lista
cboEstado.Clear

sSQL = "SELECT DISTINCT UF, IdEstado FROM CIDADE UF ORDER BY UF"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboEstado.AddItem ValidateNull(r("UF"))
   cboEstado.ItemData(cboEstado.NewIndex) = r("IdEstado")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

'cboEstado.Text = varTexto

moCombo.AttachTo cboEstado
End Sub

Private Sub cboEstado_LostFocus()
On Error GoTo TrataErro

If cboEstado.Text = "" Then txtCodUF.Text = "": Exit Sub
If cboEstado.ListIndex = -1 Then txtCodUF.Text = "": Exit Sub

txtCodUF = cboEstado.ItemData(cboEstado.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboPerfil_GotFocus()
cboPerfil.Clear
cboPerfil.AddItem "A"
cboPerfil.AddItem "B"
cboPerfil.AddItem "C"
cboPerfil.AddItem "D"
cboPerfil.AddItem "E"
cboPerfil.AddItem "F"
moCombo.AttachTo cboPerfil
End Sub


Private Sub cboTipo_GotFocus()
cboTipo.Clear
cboTipo.AddItem "FÍSICA"
cboTipo.AddItem "JURÍDICA"
cboTipo.AddItem "POUPANÇA"
moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cmdAlterar_Click()
If txtFantasia.Text = "" Or txtRazao.Text = "" Or cboCidade.Text = "" Then Exit Sub

If Not ValidarRegimeCRT() Then
    txtRegime.SetFocus
    Exit Sub
End If


'Faz a atualização de forma direta e verifica se houve algum erro
If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub txtRegime_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso do Sistema": Exit Sub
End Sub

Private Sub txtRegime_Validate(Cancel As Boolean)
    Dim nCRT As Integer
    Dim nRegime As Integer
    
    ' Garante que estamos lidando com números para evitar erro de tipo
    nCRT = Val(txtCRT.Text)
    nRegime = Val(txtRegime.Text)
    
    ' Se o campo estiver vazio, não valida agora (deixa para o botão salvar)
    If txtRegime.Text = "" Then Exit Sub

    Select Case nCRT
        Case 1 ' Simples Nacional
            If nRegime <> 1 Then
                MsgBox "Erro: Para CRT 1, o Regime deve ser 1 (Simples Nacional).", vbCritical, "Validação Fiscal"
                Cancel = True ' Prende o foco no campo txtRegime
            End If
            
        Case 2 ' Simples Nacional - Excesso
            If nRegime <> 2 Then
                MsgBox "Erro: Para CRT 2, o Regime deve ser 2 (Excesso de Sublimite).", vbCritical, "Validação Fiscal"
                Cancel = True
            End If
            
        Case 3 ' Regime Normal
            ' Verifica se o regime não pertence ao grupo do Regime Normal (3, 4 ou 5)
            If nRegime < 3 Or nRegime > 5 Then
                MsgBox "Erro: Para CRT 3, o Regime deve ser 3 (Presumido), 4 (Real) ou 5 (Arbitrado).", vbCritical, "Validação Fiscal"
                Cancel = True
            End If
            
        Case Else
            MsgBox "Por favor, preencha primeiro um CRT válido (1, 2 ou 3).", vbExclamation
            txtCRT.SetFocus
            Cancel = True
    End Select
End Sub
Private Sub cmdCertificadoDigital_Click()
Dim cStat As Integer, NFeMotivo As String, nomeCertificado As String, NroSerie As String, CNPJ As String, InicioValidade As String, FimValidade As String
Dim sistCertificado As snfe.CertDigital
Set sistCertificado = New snfe.CertDigital

 On Error GoTo deuErro
    
    iRetorno = sistCertificado.Seleciona
    
    If iRetorno Then
       nomeCertificado = sistCertificado.retCertDigital.nomeCertificado
       CNPJ = sistCertificado.retCertDigital.CNPJ
       NroSerie = sistCertificado.retCertDigital.NumeroSerie
       InicioValidade = sistCertificado.retCertDigital.DataInicio
       FimValidade = sistCertificado.retCertDigital.DataExpira
        
       'MsgBox "Subject Name: " & nomeCertificado & vbNewLine & "CNPJ: " & CNPJ & vbNewLine & "Número de Série: " & NroSerie & vbNewLine & "Validade: " & FimValidade, vbInformation
    Else
       nomeCertificado = "Operação cancelada!"
    End If
    
    If Not Vazio(NroSerie) Then
       txtCertificadoDigital.Text = NroSerie
       SQLExecuta "UPDATE empresa SET VencimentoCert = " & FdtSQL(FimValidade)
    End If
    
    Set sistCertificado = Nothing
    Exit Sub
    
deuErro:
    MsgBox Err.Description, vbExclamation + vbOKOnly, "ERRO"
    Set sistCertificado = Nothing
End Sub

Private Sub cmdConsultarIE_Click()
ShellExecute hwnd, "open", "https://dfe-portal.svrs.rs.gov.br/Nfe/Ccc", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdCopiar_Click()
txtFantasia.Text = UCase(txtFantasia.Text)
txtRazao.Text = UCase(txtRazao.Text)
txtFantasia.Text = TirarEspaco(txtFantasia.Text)
txtRazao.Text = TirarEspaco(txtRazao.Text)
Clipboard.Clear
Clipboard.SetText txtRazao.Text & mskCNPJ.Text
End Sub

Private Sub cmdDesbroquear_Click()
Dim i As Integer
i = Grid.Row

sSQL = "SELECT codigo, bloqueio, mes_ref, COD_DESBLOQUEIO, COD_TEMP, Debloqueio_Temp FROM licenca_pagamentos WHERE (codigo = " & Grid.TextMatrix(i, 1) & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    If txtCodDesbloqueio.Text <> "" Then
        If txtCodDesbloqueio.Text = r("COD_DESBLOQUEIO") Then
            dbData.Execute "UPDATE licenca_pagamentos SET bloqueio = 0, pago = 1, data_liberacao = '" & Format$(Date, "yyyy-dd-MM") & "' WHERE (codigo = " & Grid.TextMatrix(i, 1) & ");"
            MsgBox "MÊS REFERENTE FOI DESBLOQUEADO" & vbCrLf & "Tente novamente fazer o login no sistema", vbInformation
            Mostrar_Pagamentos
        Else
             MsgBox "Código de desbloqueio errado" & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
             Exit Sub
        End If
    Else
        MsgBox "Digite somente um código de desbloqueio!" & vbCrLf & "Os dois códigos não podem ser preenchidos ao mesmo tempo.", vbInformation
        Exit Sub
    End If
End If
End Sub

Private Sub cmdDesbTemp_Click()
Dim i As Integer
i = Grid.Row

sSQL = "SELECT codigo, bloqueio, mes_ref, COD_DESBLOQUEIO, COD_TEMP, Debloqueio_Temp, data_bloqueio FROM licenca_pagamentos WHERE (codigo = " & Grid.TextMatrix(i, 1) & ");"
Set r = dbData.OpenRecordset(sSQL)

Dim vDataBloq As Date
vDataBloq = r("data_bloqueio")

If Not r.BOF Then
    If txtCodDesbloqueio.Text <> "" Then
        If txtCodDesbloqueio.Text = r("COD_TEMP") Then
            If r("Debloqueio_Temp") = 0 Then
                dbData.Execute "UPDATE licenca_pagamentos SET bloqueio = 0, Debloqueio_Temp = 1, data_bloqueio = '" & Format$(vDataBloq + 3, "yyyy-dd-MM") & "' WHERE (codigo = " & Grid.TextMatrix(i, 1) & ");"
                MsgBox "VOCÊ USOU UM CÓD. TEMPORÁRIO" & vbCrLf & "Você ganhou mais 3 dias de desbloqueio!", vbInformation
                Mostrar_Pagamentos
            Else
                MsgBox "Você já usou esse cód. de desbloqueio." & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
                Exit Sub
            End If
        Else
             MsgBox "Código de desbloqueio temporário errado" & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
             Exit Sub
        End If
    Else
        MsgBox "Digite somente um código de desbloqueio!" & vbCrLf & "Os dois códigos não podem ser preenchidos ao mesmo tempo.", vbInformation
        Exit Sub
    End If
End If
End Sub


Private Sub cmdDiretorioXML_Click()
  'Opens a Treeview control that displays the directories in a computer
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = vbCr & vbCr & "Selecione a Pasta Desejada:"
  
  With tBrowseInfo
    .hwndOwner = Me.hwnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtDiretorioXML.Text = sBuffer
  End If
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   If txtFantasia.Text = "" Or txtRazao.Text = "" Or cboCidade.Text = "" Then Exit Sub
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Deseja excluir essa empresa?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Não é necessário consulta o registro antes de exclui-lo
   'sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM empresa;"
   bRet = dbData.Execute(sSQL)
    
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   Limpar_Objetos
   Form_Load
End Sub

Private Sub cmdGerarLicenca_Click()
Dim objNFeNFCe As snfe.Util
 If Vazio(mskCNPJ.Text) Then GoTo faltaDados
 
 On Error GoTo deuErro
 Set objNFeNFCe = New snfe.Util
 txtLicencaDLL.Text = objNFeNFCe.GerarLicenca("02.382.419/0001-80", mskCNPJ.Text)
 Set objNFeNFCe = Nothing
 Exit Sub

faltaDados:
 MsgBox "Falta informar o CNPJ/CPF do Emitente!", vbExclamation + vbOKOnly, "ATENÇÃO"
 Set objNFeNFCe = Nothing
 Exit Sub
 
deuErro:
 MsgBox Err.Description, vbExclamation + vbOKOnly, "ERRO AO GERAR LICENÇA"
 Err.Clear
 Set objNFeNFCe = Nothing
End Sub

Private Sub cmdGerarPagamentos_Click()
If txtDia.Text = "" Then Exit Sub

'Call GerarCodDesbloqueio
sSQL = "SELECT cnpj, razao FROM empresa"
Set r = dbData.OpenRecordset(sSQL)

Dim vCnpj As Integer
Dim vQuantRazao As Integer
If Not r.BOF Then
    vCnpj = SomarDigitos(r("cnpj"))
    vQuantRazao = Len(r("razao"))
End If

'começa a criação
Dim vDataInicio As Date
Dim vDia As Integer
Dim vMes As Integer
Dim vMesInt As String
Dim vAno As Integer
Dim vMesRef As String

vAno = Year(Date)

Dim vDataBloqueio As String

Autonumeracao_Pagamentos
Dim vmestest As String
'descobrir o mês de criação do bloqueio
If chkProximo.Value = 1 Then
    vmestest = Format(DateAdd("m", Val(1), Date), "dd/mm/yy")
    vMes = Format(DatePart("m", vmestest), "m")
    'vMes = DateAdd("m", 1, Date) + 1
    vMes = Format(Date, "m") + 1
Else
    vMes = Format(Date, "m")
End If

'saber o numero do ultimo dia daquele mês
Dim vUltimoDiaMes As Integer
vUltimoDiaMes = Day(DateSerial(vAno, vMes + 1, 0))
vDia = vUltimoDiaMes 'sabe o ultimo dia daquele mês

'If vMes = 2 Then
'    If AnoBisexto(Year(Date)) Then
'        vDia = 29
'    Else
'        vDia = 28
'    End If
'Else
'    vDia = txtDia.Text
'End If

'vMes = Format(Date, "m")

'data por extenso para preenchimento de campo
If chkProximo.Value = 1 Then
    vDataInicio = vDia & " / " & vMes & " / " & vAno
    vMesInt = Format(vDataInicio, "mmmm")
    vAno = Year(vDataInicio)
    vMesRef = vMesInt & "/" & vAno
Else
    vDataInicio = vDia & " / " & vMes & " / " & vAno
    vMesInt = Format(vDataInicio, "mmmm")
    vAno = Year(vDataInicio)
    vMesRef = vMesInt & "/" & vAno
End If

vDataBloqueio = Format(DateAdd("d", Val(5), vDataInicio), "dd/mm/yy")

'codigo de desbloqueio
    Dim vNumeroMes As Integer
    If vMesInt = "janeiro" Then
        vNumeroMes = 1
    ElseIf vMesInt = "fevereiro" Then
        vNumeroMes = 2
    ElseIf vMesInt = "março" Then
        vNumeroMes = 3
    ElseIf vMesInt = "abril" Then
        vNumeroMes = 4
    ElseIf vMesInt = "maio" Then
        vNumeroMes = 5
    ElseIf vMesInt = "junho" Then
        vNumeroMes = 6
    ElseIf vMesInt = "julho" Then
        vNumeroMes = 7
    ElseIf vMesInt = "agosto" Then
        vNumeroMes = 8
    ElseIf vMesInt = "setembro" Then
        vNumeroMes = 9
    ElseIf vMesInt = "outubro" Then
        vNumeroMes = 10
    ElseIf vMesInt = "novembro" Then
        vNumeroMes = 11
    ElseIf vMesInt = "dezembro" Then
        vNumeroMes = 12
    End If
    
    Dim vCodDesbloqueio As String
    Dim vCodDesbTemp As String
    
    'Desbloqueio
    If vNumeroMes Mod 2 = 0 Then
        'MsgBox "Par!"
        vCodDesbloqueio = Left(vCnpj, 1) & "" & Left(vQuantRazao, 1) & "" & Len(vMesInt) & "" & vNumeroMes & "" & UCase(Mid(vMesInt, 3, 1))
    Else
        'MsgBox "Ímpar!"
        vCodDesbloqueio = Mid(vCnpj, 2, 1) & "" & Mid(vQuantRazao, 2, 1) & "" & Len(vMesInt) - 1 & "" & vNumeroMes & "" & UCase(Mid(vMesInt, 2, 1))
    End If

    'Desbloqueio temporario
    If vNumeroMes Mod 2 = 0 Then
        'MsgBox "Par!"
        vCodDesbTemp = Left(vCodDesbloqueio, 1) & "" & Left(vCodDesbloqueio, 1) & "" & vNumeroMes + 1 & "" & UCase(Mid(vMesInt, 4, 1))
    Else
        'MsgBox "Ímpar!"
        vCodDesbTemp = Mid(vCodDesbloqueio, 2, 1) & "" & Mid(vCodDesbloqueio, 2, 1) & "" & Len(vMesInt) - 1 & "" & vNumeroMes + 1 & "" & UCase(Mid(vMesInt, 4, 1))
    End If
    
sSQL = "SELECT codigo FROM licenca_pagamentos;"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
    If r.RecordCount = 0 Then
        dbData.Execute "INSERT INTO  licenca_pagamentos (codigo, dia_vencimento, mes_ref, data_vencimento, data_bloqueio, bloqueio, pago, COD_DESBLOQUEIO, COD_TEMP, Debloqueio_Temp) VALUES (" & _
            lNovoCod & ", " & vDia & ", '" & vMesRef & "', '" & Format$(vDataInicio, "yyyy-dd-MM") & "', '" & Format$(vDataBloqueio, "yyyy-dd-MM") & "', 0, 0, '" & vCodDesbloqueio & "', '" & vCodDesbTemp & "', 0);"
    End If
End If

Mostrar_Pagamentos
End Sub
Private Function AnoBisexto(ValAno As Single) As Boolean
If (ValAno Mod 4 = 0) And ((ValAno Mod 100 <> 0) Or (ValAno Mod 400 = 0)) Then
        AnoBisexto = True
    Else
        AnoBisexto = False
End If
End Function
Public Function SomarDigitos(CNPJ As String) As Integer
    Dim s As Integer
    Dim i As Integer
    For i = 1 To Len(CNPJ)
      If IsNumeric(Mid(CNPJ, i, 1)) Then
        s = s + Mid(CNPJ, i, 1)
      End If
    Next
    SomarDigitos = s
End Function
Private Function Autonumeracao_Pagamentos() As Long
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS Ultimo_Pgto FROM licenca_pagamentos;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lNovoCod = r("Ultimo_Pgto") + 1
Else
    lNovoCod = 1
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Function


Private Sub cmdSalvar_Click()
'Se os dados não foram informados, sai da rotina
If txtFantasia.Text = "" Or txtRazao.Text = "" Or cboCidade.Text = "" Then Exit Sub
If txtAliqUF.Text = "" Then txtAliqUF.Text = "12"

If Not ValidarRegimeCRT() Then
    txtRegime.SetFocus
    Exit Sub
End If

'Faz a inserção de forma direta e verifica se houve algum erro
If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Function ValidarRegimeCRT() As Boolean
    Dim nCRT As Integer
    Dim nRegime As Integer
    
    ' Converte os valores dos campos para Inteiro
    nCRT = Val(txtCRT.Text)
    nRegime = Val(txtRegime.Text)
    
    ValidarRegimeCRT = True ' Começa como verdadeiro
    
    Select Case nCRT
        Case 1 ' Simples Nacional
            If nRegime <> 1 Then
                MsgBox "Inconsistência: Para CRT 1, o Regime deve ser 1 (Simples Nacional).", vbCritical, "Erro Fiscal"
                ValidarRegimeCRT = False
            End If
            
        Case 2 ' Simples Nacional - Excesso de Sublimite
            If nRegime <> 2 Then
                MsgBox "Inconsistência: Para CRT 2, o Regime deve ser 2 (Excesso de Sublimite).", vbCritical, "Erro Fiscal"
                ValidarRegimeCRT = False
            End If
            
        Case 3 ' Regime Normal
            ' Verifica se o regime não é um dos permitidos para o CRT 3
            If nRegime < 3 Or nRegime > 5 Then
                MsgBox "Inconsistência: Para CRT 3, o Regime deve ser 3 (Presumido), 4 (Real) ou 5 (Arbitrado).", vbCritical, "Erro Fiscal"
                ValidarRegimeCRT = False
            End If
            
        Case Else
            MsgBox "CRT Inválido! Use 1, 2 ou 3.", vbExclamation, "Aviso"
            ValidarRegimeCRT = False
    End Select
End Function
Private Sub cmdSped_Click()
txtFantasia.Text = UCase(txtFantasia.Text)
txtRazao.Text = UCase(txtRazao.Text)
txtFantasia.Text = TirarEspaco(txtFantasia.Text)
txtRazao.Text = TirarEspaco(txtRazao.Text)
Dim vCPF As String
vCPF = "000,000,000-00"

Clipboard.Clear
'Clipboard.SetText txtRazao.Text & mskCNPJ.Text

Clipboard.SetText "<ConfiguracaoEmpresa><CNPJ>" & mskCNPJ.Text & "</CNPJ><RazaoSocial>" & txtRazao.Text & "</RazaoSocial><CPF>" & vCPF & "</CPF><Endereco>" & txtEndereco & ", " & txtNum & ", " & txtBairro & "</Endereco><UF>PI</UF><CodigoSiafi>0</CodigoSiafi><CodigoIBGE>" & txtCodigoIBGE & "</CodigoIBGE></ConfiguracaoEmpresa>"
'"Cliente = " & cboNome.Text & ""
End Sub

Private Sub Form_Load()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long

Set moCombo = New cComboHelper

Caminho = appPathApp

sSQL = "SELECT * FROM empresa"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If totalRegistros >= 1 Then
   Mostrar_Dados r
Else
   cmdSalvar.Enabled = True
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   Limpar_Objetos
   cboPerfil.Text = "A"
End If

txtDia.Text = "30"
Mostrar_Pagamentos
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")

If Tela_Principal.StatusBar1.Panels(2).Text <> "PROGRAMADOR" Then
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   'cmdGerarPagamentos.Enabled = False
   Exit Sub
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
Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub lblProcurar_Click()
   Dim FSys As FileSystemObject 'referencia que nao deixa copiar arquivos duplicados (PROJECT / REFERENCES e selecionar MICROSOFT SCRIPTING RUNTIME)
   Set FSys = New FileSystemObject
   
   CommonDialog1.Filter = "Imagens JPG(*.jpg)|*.jpg"
   CommonDialog1.ShowOpen
   txtCaminho.Text = CommonDialog1.FileName
   
   If (CommonDialog1.FileName = "") Then Exit Sub
   
   If Not FSys.FileExists(Caminho & CommonDialog1.FileTitle) Then  'se o arquivo nao existir na pasta ele copia
      FileCopy txtCaminho.Text, Caminho & CommonDialog1.FileTitle
   End If
      txtCaminho.Text = Caminho & CommonDialog1.FileTitle
      picLogo.Picture = LoadPicture(txtCaminho.Text) 'mostrar a imagem
   Set FSys = Nothing
End Sub

Private Sub mskCelular_KeyPress(KeyAscii As Integer)
mskCelular.Mask = "(##) #####-####"
End Sub

Private Sub mskCelular_LostFocus()
If mskCelular.Text = "(__) _____-____" Then
   mskCelular.Mask = ""
   mskCelular.Text = ""
End If
End Sub

Private Sub mskCep_KeyPress(KeyAscii As Integer)
   mskCep.Mask = "##.###-###"
End Sub

Private Sub mskCep_LostFocus()
   If mskCep.Text = "__.___-___" Then
      mskCep.Mask = ""
      mskCep.Text = ""
   End If
End Sub

Private Sub mskCNPJ_KeyPress(KeyAscii As Integer)
If optJuridico.Value = True Then
   mskCNPJ.Mask = "##.###.###/####-##"
Else
    mskCNPJ.Mask = "###.###.###-##"
End If
End Sub

Private Sub mskCNPJ_LostFocus()
If mskCNPJ.Text = "__.___.___/____-__" Then
   mskCNPJ.Mask = ""
   mskCNPJ.Text = ""
End If
End Sub

Private Sub mskTelefone_KeyPress(KeyAscii As Integer)
mskTelefone.Mask = "(##) ####-####"
End Sub

Private Sub mskTelefone_LostFocus()
If mskTelefone.Text = "(__) ____-____" Then
   mskTelefone.Mask = ""
   mskTelefone.Text = ""
End If
End Sub

Private Sub optFisica_Click()
MostrarTipoEmpresa
End Sub

Private Sub optJuridico_Click()
MostrarTipoEmpresa
End Sub

Private Sub txtAliqUF_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii = Asc(".") Then
        KeyAscii = Asc(",")
    ElseIf KeyAscii = Asc(",") Then
        KeyAscii = Asc(",")
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online.Info": Exit Sub
End Sub


Private Sub txtAliqUF_LostFocus()
If txtAliqUF.Text = "" Then txtAliqUF.Text = "0"
Moeda txtAliqUF
End Sub


Private Sub txtAmbienteNF_Change()
   If Not Vazio(txtAmbienteNF.Text) Then
      If Val(txtAmbienteNF.Text) = 1 Then lblAmbienteStatus.Caption = "Produção"
      If Val(txtAmbienteNF.Text) = 2 Then lblAmbienteStatus.Caption = "Homologação"
   End If
End Sub

Private Sub txtAmbienteNF_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("2") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "EkklesiaSoft": Exit Sub
End Sub

Private Sub txtBairro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboCidade_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodCid_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodCid.Text = "" Then txtCodigoIBGE.Text = "": Exit Sub

sSQL = "SELECT CodigoMunicipio, ID FROM CIDADE WHERE (id = " & txtCodCid.Text & ")"
Set r = dbData.OpenRecordset(sSQL)

   If Not r.BOF Then
    txtCodigoIBGE.Text = ValidateNull(r("CodigoMunicipio"))
   End If
   
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub txtCodDesbloqueio_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCodigoIBGE_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "EkklesiaSoft": Exit Sub
End Sub

Private Sub txtCRT_Change()
If Not Vazio(txtCRT.Text) Then
   lblCRTNome.Caption = CRT(txtCRT.Text)
End If
End Sub

Private Sub txtCRT_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso do Sistema": Exit Sub
End Sub

Private Sub txtCRT_LostFocus()
lblCRTNome.Caption = CRT(txtCRT.Text)
If txtCRT.Text = 1 Or txtCRT.Text = 2 Then
    lblAliqSN.Visible = True
    txtpICMSSN.Visible = True
    cboDIFAL.Text = "NÃO"
    lblDifal.Enabled = False
    cboDIFAL.Enabled = False
    txtRegime.Text = "1"
Else
    lblAliqSN.Visible = False
    txtpICMSSN.Visible = False
    cboDIFAL.Text = "NÃO"
    cboDIFAL.Enabled = True
    lblDifal.Enabled = True
    txtRegime.Text = "3"
End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFantasia_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFantasia_LostFocus()
txtFantasia.Text = TirarEspaco(txtFantasia.Text)
End Sub


Private Sub txtNFCeIDToken_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "EkklesiaSoft": Exit Sub
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtpICMSSN_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii = Asc(".") Then
        KeyAscii = Asc(",")
    ElseIf KeyAscii = Asc(",") Then
        KeyAscii = Asc(",")
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online.Info": Exit Sub
End Sub

Private Sub txtpICMSSN_LostFocus()
    If txtpICMSSN.Text = "" Then txtpICMSSN.Text = "0"
    Moeda txtpICMSSN
End Sub

Private Sub txtRazao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRazao_LostFocus()
txtFantasia.Text = TirarEspaco(txtFantasia.Text)
End Sub


Private Sub txtRegime_Change()
txtRegime_LostFocus
End Sub

Private Sub txtRegime_LostFocus()
If Not Vazio(txtRegime.Text) Then
   lblREGIME.Caption = REGIME(txtRegime.Text)
End If

End Sub


