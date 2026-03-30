VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Configuracao_Geral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONFIGURAÇÕES"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "Configuracao_Geral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   8865
      TabIndex        =   9
      Top             =   60
      Width           =   8895
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIGURAÇÕES GERAIS"
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
         Left            =   1140
         TabIndex        =   10
         Top             =   180
         Width           =   4005
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   300
         Picture         =   "Configuracao_Geral.frx":23D2
         Stretch         =   -1  'True
         Top             =   60
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   9930
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "18:30"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8955
      Left            =   60
      TabIndex        =   11
      Top             =   900
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   15796
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabMaxWidth     =   2999
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "GERAL"
      TabPicture(0)   =   "Configuracao_Geral.frx":82DC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(1)=   "FraConfiguracao"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(3)=   "FrameBackup"
      Tab(0).Control(4)=   "cmdSalvarGeral"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "PDV"
      TabPicture(1)   =   "Configuracao_Geral.frx":82F8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "ADICIONAIS"
      TabPicture(2)   =   "Configuracao_Geral.frx":8314
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame13"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(3)=   "cmdSalvarBalanca"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "IMPRESSÃO"
      TabPicture(3)   =   "Configuracao_Geral.frx":8330
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboDeclararRecebedor"
      Tab(3).Control(1)=   "txtFormaParcelasImpressao"
      Tab(3).Control(2)=   "cboFormaParcelasImpressao"
      Tab(3).Control(3)=   "cboINIImpNormal"
      Tab(3).Control(4)=   "Frame7"
      Tab(3).Control(5)=   "Frame8"
      Tab(3).Control(6)=   "Frame17"
      Tab(3).Control(7)=   "Frame2"
      Tab(3).Control(8)=   "cmdSalvarImpressao"
      Tab(3).Control(9)=   "Label61"
      Tab(3).Control(10)=   "Label49"
      Tab(3).Control(11)=   "Label14"
      Tab(3).Control(12)=   "Label48"
      Tab(3).ControlCount=   13
      TabCaption(4)   =   "PROGRAMADOR"
      TabPicture(4)   =   "Configuracao_Geral.frx":834C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chameleonButton1"
      Tab(4).Control(1)=   "txtCodDesbloqueioTemp"
      Tab(4).Control(2)=   "txtCodDesbloqueio"
      Tab(4).Control(3)=   "txtFantasia"
      Tab(4).Control(4)=   "cboAno"
      Tab(4).Control(5)=   "cboMes"
      Tab(4).Control(6)=   "txtRazao"
      Tab(4).Control(7)=   "Grid"
      Tab(4).Control(8)=   "mskCPF"
      Tab(4).Control(9)=   "cmdAdicionar"
      Tab(4).Control(10)=   "cmdMostrarSenha"
      Tab(4).Control(11)=   "cmdPrepara"
      Tab(4).Control(12)=   "cmdPrepara2"
      Tab(4).Control(13)=   "cmdNovo"
      Tab(4).Control(14)=   "cmdLocalizar"
      Tab(4).Control(15)=   "cmdMarcar"
      Tab(4).Control(16)=   "cmdDesmarcar"
      Tab(4).Control(17)=   "cmdDesmarcarTodos"
      Tab(4).Control(18)=   "Label60"
      Tab(4).Control(19)=   "Label59"
      Tab(4).Control(20)=   "Label58"
      Tab(4).Control(21)=   "Label57"
      Tab(4).Control(22)=   "Label56"
      Tab(4).Control(23)=   "Label54"
      Tab(4).Control(24)=   "Label53"
      Tab(4).ControlCount=   25
      Begin VB.Frame Frame13 
         Caption         =   "Cashback"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   219
         Top             =   3180
         Width           =   8655
         Begin VB.TextBox txtCashbackValidade 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2340
            TabIndex        =   229
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtCashbackAP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            TabIndex        =   227
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cboCashbackPrazo 
            Height          =   315
            Left            =   2325
            TabIndex        =   224
            Top             =   720
            Width           =   1395
         End
         Begin VB.TextBox txtCashbackAV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            TabIndex        =   223
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cboCashbackVista 
            Height          =   315
            Left            =   2325
            TabIndex        =   220
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Validade (dias):"
            Height          =   195
            Left            =   1230
            TabIndex        =   228
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Percentagem:"
            Height          =   195
            Left            =   3840
            TabIndex        =   226
            Top             =   720
            Width           =   990
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Venda à prazo:"
            Height          =   195
            Left            =   1200
            TabIndex        =   225
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Percentagem:"
            Height          =   195
            Left            =   3840
            TabIndex        =   222
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Venda à vista:"
            Height          =   195
            Left            =   1260
            TabIndex        =   221
            Top             =   360
            Width           =   1020
         End
      End
      Begin VB.ComboBox cboDeclararRecebedor 
         Height          =   315
         Left            =   -69660
         TabIndex        =   208
         Top             =   2160
         Width           =   3435
      End
      Begin VB.TextBox txtFormaParcelasImpressao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -66780
         TabIndex        =   207
         Top             =   1140
         Width           =   495
      End
      Begin VB.ComboBox cboFormaParcelasImpressao 
         Height          =   315
         Left            =   -69660
         TabIndex        =   205
         Top             =   1440
         Width           =   3405
      End
      Begin VB.ComboBox cboINIImpNormal 
         Height          =   315
         Left            =   -69660
         TabIndex        =   202
         Top             =   780
         Width           =   3405
      End
      Begin VB.Frame Frame9 
         Caption         =   "Juros"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   182
         Top             =   4080
         Width           =   8655
         Begin VB.ComboBox cboTipoJuros 
            Height          =   315
            Left            =   4020
            TabIndex        =   188
            Top             =   960
            Width           =   1275
         End
         Begin VB.TextBox txtTipoJuros 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5310
            TabIndex        =   187
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtJurosMes 
            Height          =   315
            Left            =   4020
            TabIndex        =   184
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtJuroDia 
            Height          =   315
            Left            =   4020
            Locked          =   -1  'True
            TabIndex        =   183
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cobrar juros sobre:"
            Height          =   195
            Left            =   2580
            TabIndex        =   189
            Top             =   960
            Width           =   1320
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Juros ao mês:"
            Height          =   195
            Left            =   2940
            TabIndex        =   186
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Juros ao Dia:"
            Height          =   195
            Left            =   3000
            TabIndex        =   185
            Top             =   600
            Width           =   930
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "NFCe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74820
         TabIndex        =   171
         Top             =   6660
         Width           =   4995
         Begin VB.ComboBox cboCombinarImpNFCe 
            Height          =   315
            Left            =   1995
            TabIndex        =   180
            Top             =   1680
            Width           =   2415
         End
         Begin VB.ComboBox cboConfirmaPrazoNFCe 
            Height          =   315
            Left            =   1995
            TabIndex        =   178
            Top             =   1320
            Width           =   2415
         End
         Begin VB.ComboBox cboConfirmaCPFNFCe 
            Height          =   315
            Left            =   1995
            TabIndex        =   176
            Top             =   960
            Width           =   2415
         End
         Begin VB.ComboBox cboImprimirNFCe 
            Height          =   315
            Left            =   1995
            TabIndex        =   173
            Top             =   240
            Width           =   2415
         End
         Begin VB.ComboBox cboConfirmaImpressaoNFCe 
            Height          =   315
            Left            =   1995
            TabIndex        =   172
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Combinar Impressão?:"
            Height          =   195
            Left            =   360
            TabIndex        =   181
            Top             =   1680
            Width           =   1560
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vendas À Prazo?:"
            Height          =   195
            Left            =   660
            TabIndex        =   179
            Top             =   1320
            Width           =   1275
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Confirmar CPF?:"
            Height          =   195
            Left            =   780
            TabIndex        =   177
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Imprimir?:"
            Height          =   195
            Left            =   1275
            TabIndex        =   175
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Confirmar Impressão?:"
            Height          =   195
            Left            =   360
            TabIndex        =   174
            Top             =   600
            Width           =   1560
         End
      End
      Begin ChamaleonBtn.chameleonButton chameleonButton1 
         Height          =   315
         Left            =   -66600
         TabIndex        =   166
         Top             =   5460
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Configuracao_Geral.frx":8368
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtCodDesbloqueioTemp 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   162
         Top             =   6420
         Width           =   975
      End
      Begin VB.TextBox txtCodDesbloqueio 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   -70980
         Locked          =   -1  'True
         TabIndex        =   156
         Top             =   6420
         Width           =   975
      End
      Begin VB.TextBox txtFantasia 
         Height          =   315
         Left            =   -74880
         TabIndex        =   153
         Top             =   5460
         Width           =   2835
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   -73320
         TabIndex        =   151
         Top             =   6420
         Width           =   1155
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   -74880
         TabIndex        =   149
         Top             =   6420
         Width           =   1515
      End
      Begin VB.TextBox txtRazao 
         Height          =   315
         Left            =   -72000
         TabIndex        =   155
         Top             =   5460
         Width           =   3735
      End
      Begin VB.Frame Frame6 
         Caption         =   "Aluguel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   139
         Top             =   1800
         Width           =   8655
         Begin VB.ComboBox cboHabilitarAluguel 
            Height          =   315
            Left            =   2205
            TabIndex        =   141
            Top             =   360
            Width           =   1395
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   2205
            TabIndex        =   140
            Top             =   720
            Width           =   3675
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Habilitar Aluguel:"
            Height          =   195
            Left            =   975
            TabIndex        =   143
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Aluguel:"
            Height          =   195
            Left            =   945
            TabIndex        =   142
            Top             =   780
            Width           =   1230
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Desconto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   180
         TabIndex        =   116
         Top             =   6840
         Width           =   8535
         Begin VB.Frame Frame12 
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
            Height          =   1755
            Left            =   60
            TabIndex        =   210
            Top             =   240
            Width           =   2295
            Begin VB.ComboBox cboTipoDesconto 
               Height          =   315
               Left            =   120
               TabIndex        =   217
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtTipoDesc 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1140
               TabIndex        =   216
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CheckBox chkCartaoCredito 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Excluir Cartão Crédito"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   180
               TabIndex        =   214
               Top             =   1500
               Width           =   1995
            End
            Begin VB.CheckBox chkCartaoDebito 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Excluir Cartão Débito"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   180
               TabIndex        =   213
               Top             =   1320
               Width           =   1995
            End
            Begin VB.CheckBox chkLimiteDesc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Limitar Desconto"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   180
               TabIndex        =   212
               Top             =   960
               Width           =   1515
            End
            Begin VB.CheckBox chkLimiteGerente 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Permissão do Gerente"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   180
               TabIndex        =   211
               Top             =   1140
               Width           =   1935
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Desc.:"
               Height          =   195
               Left            =   120
               TabIndex        =   218
               Top             =   240
               Width           =   1050
            End
         End
         Begin VB.Frame frmValorDesc 
            Caption         =   "Descontos"
            Height          =   855
            Left            =   2400
            TabIndex        =   130
            Top             =   240
            Width           =   2115
            Begin VB.TextBox txtValorDescAV 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   132
               Top             =   480
               Width           =   915
            End
            Begin VB.TextBox txtValorDescAP 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1080
               TabIndex        =   131
               Top             =   480
               Width           =   915
            End
            Begin VB.Label lblDescAV 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "À Vista(%):"
               Height          =   195
               Left            =   120
               TabIndex        =   134
               Top             =   240
               Width           =   750
            End
            Begin VB.Label lblDescAP 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "À Prazo(%):"
               Height          =   195
               Left            =   1080
               TabIndex        =   133
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame frmLimites 
            Caption         =   "Limites"
            Height          =   1575
            Left            =   4560
            TabIndex        =   117
            Top             =   240
            Width           =   2715
            Begin VB.TextBox txtDescAV3 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   960
               TabIndex        =   126
               Top             =   1200
               Width           =   795
            End
            Begin VB.TextBox txtDescAP1 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1800
               TabIndex        =   125
               Top             =   480
               Width           =   795
            End
            Begin VB.TextBox txtDescAP2 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1800
               TabIndex        =   124
               Top             =   840
               Width           =   795
            End
            Begin VB.TextBox txtDescAP3 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1800
               TabIndex        =   123
               Top             =   1200
               Width           =   795
            End
            Begin VB.TextBox txtDescMargemAV3 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   122
               Top             =   1200
               Width           =   795
            End
            Begin VB.TextBox txtDescMargemAV1 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   121
               Top             =   480
               Width           =   795
            End
            Begin VB.TextBox txtDescMargemAV2 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   120
               Top             =   840
               Width           =   795
            End
            Begin VB.TextBox txtDescAV2 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   960
               TabIndex        =   119
               Top             =   840
               Width           =   795
            End
            Begin VB.TextBox txtDescAV1 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   960
               TabIndex        =   118
               Top             =   480
               Width           =   795
            End
            Begin VB.Label lblMargem 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Margem(R$):"
               Height          =   195
               Left            =   120
               TabIndex        =   129
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lblLimiteAP 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "À Prazo:"
               Height          =   195
               Left            =   1800
               TabIndex        =   128
               Top             =   240
               Width           =   600
            End
            Begin VB.Label lblLimiteAV 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "À Vista:"
               Height          =   195
               Left            =   1080
               TabIndex        =   127
               Top             =   240
               Width           =   540
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem de Serviços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   115
         Top             =   420
         Width           =   8655
         Begin VB.ComboBox cboTipoOS 
            Height          =   315
            Left            =   2205
            TabIndex        =   136
            Top             =   720
            Width           =   3675
         End
         Begin VB.ComboBox cboHabilitarOS 
            Height          =   315
            Left            =   2205
            TabIndex        =   135
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Ordem de Serviços:"
            Height          =   195
            Left            =   120
            TabIndex        =   138
            Top             =   780
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Habilitar Ordem de Serviços:"
            Height          =   195
            Left            =   150
            TabIndex        =   137
            Top             =   360
            Width           =   2010
         End
      End
      Begin VB.Frame FraConfiguracao 
         Caption         =   "Balança e Peso"
         Height          =   1515
         Left            =   -74880
         TabIndex        =   110
         Top             =   6660
         Width           =   8655
         Begin VB.Frame Frame11 
            Caption         =   "Balança/Pesagem"
            Height          =   675
            Left            =   3540
            TabIndex        =   197
            Top             =   780
            Width           =   3375
            Begin VB.TextBox txtQtdeDigitosBalanca 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2760
               TabIndex        =   199
               Top             =   300
               Width           =   495
            End
            Begin VB.TextBox txtIniciaisBalanca 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1080
               TabIndex        =   198
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Digito Inicial:"
               Height          =   195
               Left            =   120
               TabIndex        =   201
               Top             =   300
               Width           =   900
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde Digitos:"
               Height          =   195
               Left            =   1800
               TabIndex        =   200
               Top             =   300
               Width           =   915
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Etiqueta/Pesagem"
            Height          =   675
            Left            =   120
            TabIndex        =   192
            Top             =   780
            Width           =   3375
            Begin VB.TextBox txtIniciaisEtiqueta 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1080
               TabIndex        =   196
               Top             =   300
               Width           =   495
            End
            Begin VB.TextBox txtQtdeDigitosEtiqueta 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2760
               TabIndex        =   195
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde Digitos:"
               Height          =   195
               Left            =   1800
               TabIndex        =   194
               Top             =   300
               Width           =   915
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Digito Inicial:"
               Height          =   195
               Left            =   120
               TabIndex        =   193
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.ComboBox cboBalancaPorta 
            Height          =   315
            ItemData        =   "Configuracao_Geral.frx":8384
            Left            =   3840
            List            =   "Configuracao_Geral.frx":83AC
            Style           =   2  'Dropdown List
            TabIndex        =   112
            Top             =   360
            Width           =   1395
         End
         Begin VB.ComboBox ComBalancaModelo 
            Height          =   315
            ItemData        =   "Configuracao_Geral.frx":83FA
            Left            =   840
            List            =   "Configuracao_Geral.frx":841C
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblPorta 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Porta:"
            Height          =   195
            Left            =   3300
            TabIndex        =   114
            Top             =   360
            Width           =   420
         End
         Begin VB.Label lblModelo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo:"
            Height          =   195
            Left            =   120
            TabIndex        =   113
            Top             =   360
            Width           =   570
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Configuração"
         Height          =   3615
         Left            =   -74880
         TabIndex        =   80
         Top             =   480
         Width           =   8655
         Begin VB.ComboBox cboLimitarCompra 
            Height          =   315
            Left            =   4020
            TabIndex        =   190
            Top             =   3120
            Width           =   1155
         End
         Begin VB.ComboBox cboMultiplasRef 
            Height          =   315
            Left            =   4020
            TabIndex        =   144
            Top             =   2760
            Width           =   1155
         End
         Begin VB.TextBox txtTipoCaixa 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5280
            TabIndex        =   109
            Top             =   2400
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cboTipoCaixa 
            Height          =   315
            Left            =   4020
            TabIndex        =   8
            Top             =   2400
            Width           =   1275
         End
         Begin VB.TextBox txtQuantDiasDesativar 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   5220
            TabIndex        =   7
            Top             =   2040
            Width           =   675
         End
         Begin VB.ComboBox cboDesativarClientes 
            Height          =   315
            Left            =   4020
            TabIndex        =   6
            Top             =   2040
            Width           =   1155
         End
         Begin VB.ComboBox cboTipoReciboHaver 
            Height          =   315
            Left            =   4020
            TabIndex        =   5
            Top             =   1680
            Width           =   2295
         End
         Begin VB.ComboBox cboTipoReciboPgto 
            Height          =   315
            Left            =   4020
            TabIndex        =   4
            Top             =   1320
            Width           =   2295
         End
         Begin VB.ComboBox cboValorVenda 
            Height          =   315
            Left            =   4020
            TabIndex        =   3
            Top             =   960
            Width           =   3675
         End
         Begin VB.TextBox txtValorVenda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7725
            TabIndex        =   82
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtTipoCadastroProduto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7785
            TabIndex        =   81
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cboTipoEmpresa 
            Height          =   315
            Left            =   4020
            TabIndex        =   1
            Top             =   240
            Width           =   3675
         End
         Begin VB.ComboBox cboIncluirPrecos 
            Height          =   315
            Left            =   4020
            TabIndex        =   2
            Top             =   600
            Width           =   3675
         End
         Begin ChamaleonBtn.chameleonButton cmdDesativarClientes 
            Height          =   315
            Left            =   5940
            TabIndex        =   101
            Top             =   2040
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Desativar Todos"
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
            MICON           =   "Configuracao_Geral.frx":847C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Limitar Valor de Compra:"
            Height          =   195
            Left            =   2250
            TabIndex        =   191
            Top             =   3120
            Width           =   1710
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Multiplas Referências:"
            Height          =   195
            Left            =   2400
            TabIndex        =   145
            Top             =   2760
            Width           =   1560
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de caixa:"
            Height          =   195
            Left            =   2910
            TabIndex        =   108
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Desativar Clientes em débito:"
            Height          =   195
            Left            =   1920
            TabIndex        =   102
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Recibo de Haver:"
            Height          =   195
            Left            =   2100
            TabIndex        =   87
            Top             =   1680
            Width           =   1845
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Recibo Pagamento:"
            Height          =   195
            Left            =   1950
            TabIndex        =   86
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Valor de Venda:"
            Height          =   195
            Left            =   2235
            TabIndex        =   85
            Top             =   960
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de empresa:"
            Height          =   195
            Left            =   2745
            TabIndex        =   84
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Aceitar entrada de preço e quantidade no cadastro:"
            Height          =   195
            Left            =   300
            TabIndex        =   83
            Top             =   600
            Width           =   3660
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Principal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6435
         Left            =   180
         TabIndex        =   55
         Top             =   360
         Width           =   8535
         Begin VB.TextBox txtTempoLogoff 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   4080
            TabIndex        =   106
            Top             =   5940
            Width           =   1155
         End
         Begin VB.ComboBox cboLogoffAutomatico 
            Height          =   315
            Left            =   2850
            TabIndex        =   105
            Top             =   5940
            Width           =   1215
         End
         Begin VB.ComboBox cboTipoLogin 
            Height          =   315
            Left            =   2850
            TabIndex        =   103
            Top             =   5580
            Width           =   2415
         End
         Begin VB.ComboBox cboAcrescCredito 
            Height          =   315
            Left            =   2850
            TabIndex        =   98
            Top             =   4860
            Width           =   2415
         End
         Begin VB.TextBox txtValorAcrescCredito 
            Height          =   315
            Left            =   2850
            TabIndex        =   97
            Top             =   5220
            Width           =   2415
         End
         Begin VB.TextBox cboAcrescCreditoConf 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5310
            TabIndex        =   96
            Top             =   4860
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cboAcrescDebito 
            Height          =   315
            Left            =   2850
            TabIndex        =   93
            Top             =   4140
            Width           =   2415
         End
         Begin VB.TextBox txtValorAcrescDebito 
            Height          =   315
            Left            =   2850
            TabIndex        =   92
            Top             =   4500
            Width           =   2415
         End
         Begin VB.TextBox txtAcrescDebitoConf 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5310
            TabIndex        =   91
            Top             =   4140
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cboSegAvancado 
            Height          =   315
            Left            =   2850
            TabIndex        =   89
            Top             =   3780
            Width           =   2415
         End
         Begin VB.TextBox txtCaminho 
            Height          =   315
            Left            =   2880
            TabIndex        =   66
            Top             =   180
            Width           =   4335
         End
         Begin VB.TextBox txtCaminhoCupom 
            Height          =   315
            Left            =   2880
            TabIndex        =   65
            Top             =   540
            Width           =   4335
         End
         Begin VB.TextBox txtTipoIndPDV 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5325
            TabIndex        =   64
            Top             =   2340
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cboIdentVendedor 
            Height          =   315
            Left            =   2880
            TabIndex        =   63
            Top             =   2340
            Width           =   2415
         End
         Begin VB.ComboBox cboEstoqueNegativo 
            Height          =   315
            Left            =   2880
            TabIndex        =   62
            Top             =   900
            Width           =   2415
         End
         Begin VB.ComboBox cboClienteDebito 
            Height          =   315
            Left            =   2880
            TabIndex        =   61
            Top             =   1260
            Width           =   2415
         End
         Begin VB.ComboBox cboIdentificarMaquina 
            Height          =   315
            Left            =   2880
            TabIndex        =   60
            Top             =   1980
            Width           =   2415
         End
         Begin VB.TextBox txtQuantDiasBloqueiar 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2880
            TabIndex        =   59
            Top             =   1620
            Width           =   2415
         End
         Begin VB.ComboBox cboConfFerchamentoAV 
            Height          =   315
            Left            =   2865
            TabIndex        =   58
            Top             =   2700
            Width           =   2415
         End
         Begin VB.ComboBox cboConfFerchamentoAP 
            Height          =   315
            Left            =   2865
            TabIndex        =   57
            Top             =   3060
            Width           =   2415
         End
         Begin VB.ComboBox cboConfFerchamentoORC 
            Height          =   315
            Left            =   2865
            TabIndex        =   56
            Top             =   3420
            Width           =   2415
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   7260
            Top             =   3840
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ChamaleonBtn.chameleonButton cmdBuscarPlano 
            Height          =   315
            Left            =   7275
            TabIndex        =   67
            Top             =   180
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Busca"
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
            MICON           =   "Configuracao_Geral.frx":8498
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdBuscarLogo 
            Height          =   315
            Left            =   7275
            TabIndex        =   68
            Top             =   540
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Busca"
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
            MICON           =   "Configuracao_Geral.frx":84B4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdSalvarPDV 
            Height          =   675
            Left            =   6480
            TabIndex        =   215
            Top             =   5520
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1191
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
            MICON           =   "Configuracao_Geral.frx":84D0
            PICN            =   "Configuracao_Geral.frx":84EC
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
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logoff Automático:"
            Height          =   195
            Left            =   1380
            TabIndex        =   107
            Top             =   5940
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Login:"
            Height          =   195
            Left            =   1725
            TabIndex        =   104
            Top             =   5580
            Width           =   1020
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Acrescimo Venda - Cartão Crédito:"
            Height          =   195
            Left            =   315
            TabIndex        =   100
            Top             =   4860
            Width           =   2430
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do Acréscimo:"
            Height          =   195
            Left            =   1305
            TabIndex        =   99
            Top             =   5220
            Width           =   1410
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Acrescimo Venda - Cartão Débito:"
            Height          =   195
            Left            =   345
            TabIndex        =   95
            Top             =   4140
            Width           =   2400
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do Acréscimo:"
            Height          =   195
            Left            =   1305
            TabIndex        =   94
            Top             =   4500
            Width           =   1410
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Segurança nas Opções Avançadas:"
            Height          =   195
            Left            =   180
            TabIndex        =   90
            Top             =   3780
            Width           =   2580
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Plano de Fundo:"
            Height          =   195
            Left            =   1635
            TabIndex        =   78
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logomarca (Cupom):"
            Height          =   195
            Left            =   1335
            TabIndex        =   77
            Top             =   540
            Width           =   1470
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Identificação do vendedor:"
            Height          =   195
            Left            =   915
            TabIndex        =   76
            Top             =   2340
            Width           =   1905
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Venda (estoque negativo):"
            Height          =   195
            Left            =   945
            TabIndex        =   75
            Top             =   900
            Width           =   1875
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bloquear Clientes em Débito:"
            Height          =   195
            Left            =   750
            TabIndex        =   74
            Top             =   1260
            Width           =   2040
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Identificação da Maquina:"
            Height          =   195
            Left            =   945
            TabIndex        =   73
            Top             =   1980
            Width           =   1845
         End
         Begin VB.Label lblDiasBloqueio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Quant. dias (Bloqueio):"
            Height          =   195
            Left            =   1215
            TabIndex        =   72
            Top             =   1620
            Width           =   1605
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Confirmar fechamendo à Vista :"
            Height          =   195
            Left            =   570
            TabIndex        =   71
            Top             =   2700
            Width           =   2205
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Confirmar fechamendo à Prazo:"
            Height          =   195
            Left            =   555
            TabIndex        =   70
            Top             =   3060
            Width           =   2220
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Confirmar fechamendo Orçamento:"
            Height          =   195
            Left            =   315
            TabIndex        =   69
            Top             =   3420
            Width           =   2460
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "PDV - Orçamento"
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
         Left            =   -74820
         TabIndex        =   43
         Top             =   4860
         Width           =   4995
         Begin VB.ComboBox cboConfirmaImpressaoORC 
            Height          =   315
            Left            =   1995
            TabIndex        =   53
            Top             =   660
            Width           =   2415
         End
         Begin VB.ComboBox cboImprimirORC 
            Height          =   315
            Left            =   1995
            TabIndex        =   47
            Top             =   300
            Width           =   2415
         End
         Begin VB.ComboBox cboTipoImpressaoORC 
            Height          =   315
            Left            =   1995
            TabIndex        =   46
            Top             =   1020
            Width           =   2415
         End
         Begin VB.TextBox txtNumCopiaORC 
            Height          =   315
            Left            =   1995
            TabIndex        =   45
            Top             =   1380
            Width           =   2415
         End
         Begin VB.TextBox txtTipoImpressaoORC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   44
            Top             =   1020
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Imprimir?:"
            Height          =   195
            Left            =   1275
            TabIndex        =   51
            Top             =   300
            Width           =   660
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Confirmar Impressão?:"
            Height          =   195
            Left            =   375
            TabIndex        =   50
            Top             =   660
            Width           =   1560
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Impressão:"
            Height          =   195
            Left            =   510
            TabIndex        =   49
            Top             =   1020
            Width           =   1425
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Número de Copias:"
            Height          =   195
            Left            =   555
            TabIndex        =   48
            Top             =   1380
            Width           =   1350
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "PDV - Vendas à Prazo"
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
         Left            =   -74820
         TabIndex        =   31
         Top             =   2700
         Width           =   4995
         Begin VB.ComboBox cboImprimirAP 
            Height          =   315
            Left            =   1995
            TabIndex        =   37
            Top             =   300
            Width           =   2415
         End
         Begin VB.ComboBox cboConfirmaImpressaoAP 
            Height          =   315
            Left            =   1995
            TabIndex        =   36
            Top             =   660
            Width           =   2415
         End
         Begin VB.ComboBox cboNotaEntregaAP 
            Height          =   315
            Left            =   1995
            TabIndex        =   35
            Top             =   1020
            Width           =   2415
         End
         Begin VB.ComboBox cboTipoImpressaoAP 
            Height          =   315
            Left            =   1995
            TabIndex        =   34
            Top             =   1380
            Width           =   2415
         End
         Begin VB.TextBox txtNumCopiaAP 
            Height          =   315
            Left            =   1995
            TabIndex        =   33
            Top             =   1740
            Width           =   2415
         End
         Begin VB.TextBox txtTipoImpressaoAP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   32
            Top             =   1380
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Imprimir?:"
            Height          =   195
            Left            =   1275
            TabIndex        =   42
            Top             =   300
            Width           =   660
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Confirmar Impressão?:"
            Height          =   195
            Left            =   375
            TabIndex        =   41
            Top             =   660
            Width           =   1560
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nota de Entrega:"
            Height          =   195
            Left            =   720
            TabIndex        =   40
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Impressão:"
            Height          =   195
            Left            =   510
            TabIndex        =   39
            Top             =   1380
            Width           =   1425
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Número de Copias:"
            Height          =   195
            Left            =   555
            TabIndex        =   38
            Top             =   1740
            Width           =   1350
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "PDV - Vendas à Vista"
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
         Left            =   -74820
         TabIndex        =   19
         Top             =   480
         Width           =   4995
         Begin VB.TextBox txtTipoImpressaoAV 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   25
            Top             =   1380
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtNumCopia 
            Height          =   315
            Left            =   1995
            TabIndex        =   24
            Top             =   1740
            Width           =   2415
         End
         Begin VB.ComboBox cboTipoImpressaoAV 
            Height          =   315
            Left            =   1995
            TabIndex        =   23
            Top             =   1380
            Width           =   2415
         End
         Begin VB.ComboBox cboNotaEntregaAV 
            Height          =   315
            Left            =   1995
            TabIndex        =   22
            Top             =   1020
            Width           =   2415
         End
         Begin VB.ComboBox cboConfirmaImpressaoAV 
            Height          =   315
            Left            =   1995
            TabIndex        =   21
            Top             =   660
            Width           =   2415
         End
         Begin VB.ComboBox cboImprimirAV 
            Height          =   315
            Left            =   1995
            TabIndex        =   20
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Número de Copias:"
            Height          =   195
            Left            =   555
            TabIndex        =   30
            Top             =   1740
            Width           =   1350
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Impressão:"
            Height          =   195
            Left            =   510
            TabIndex        =   29
            Top             =   1380
            Width           =   1425
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nota de Entrega:"
            Height          =   195
            Left            =   720
            TabIndex        =   28
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Confirmar Impressão?:"
            Height          =   195
            Left            =   375
            TabIndex        =   27
            Top             =   660
            Width           =   1560
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Imprimir?:"
            Height          =   195
            Left            =   1275
            TabIndex        =   26
            Top             =   300
            Width           =   660
         End
      End
      Begin VB.Frame FrameBackup 
         Caption         =   "Configuração de Backup"
         Height          =   1155
         Left            =   -74880
         TabIndex        =   12
         Top             =   5460
         Width           =   8655
         Begin VB.TextBox txtPastaBackup 
            Height          =   315
            Left            =   1200
            TabIndex        =   15
            Top             =   360
            Width           =   7155
         End
         Begin VB.CheckBox chkBackupAutomatico 
            Caption         =   "Backup Automático"
            Height          =   195
            Left            =   2760
            TabIndex        =   13
            Top             =   825
            Width           =   1935
         End
         Begin MSMask.MaskEdBox txtBackupHorario 
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            Top             =   780
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "99:99"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPastaBackup 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pasta Backup:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblPROCURARPastaBackup 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   7260
            TabIndex        =   17
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblHorario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horário"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   780
            Width           =   510
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarImpressao 
         Height          =   615
         Left            =   -68460
         TabIndex        =   52
         Top             =   7980
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Configuracao_Geral.frx":A27E
         PICN            =   "Configuracao_Geral.frx":A29A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarGeral 
         Height          =   615
         Left            =   -68400
         TabIndex        =   79
         Top             =   8220
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Configuracao_Geral.frx":C02C
         PICN            =   "Configuracao_Geral.frx":C048
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarBalanca 
         Height          =   615
         Left            =   -68400
         TabIndex        =   88
         Top             =   7980
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Configuracao_Geral.frx":DDDA
         PICN            =   "Configuracao_Geral.frx":DDF6
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
         Height          =   4635
         Left            =   -74880
         TabIndex        =   146
         Top             =   420
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8176
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox mskCPF 
         Height          =   315
         Left            =   -68220
         TabIndex        =   157
         Top             =   5460
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdAdicionar 
         Height          =   315
         Left            =   -74880
         TabIndex        =   159
         Top             =   5820
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
         MICON           =   "Configuracao_Geral.frx":FB88
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMostrarSenha 
         Height          =   315
         Left            =   -72120
         TabIndex        =   158
         Top             =   6420
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Mostrar"
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
         MICON           =   "Configuracao_Geral.frx":FBA4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdPrepara 
         Height          =   315
         Left            =   -69960
         TabIndex        =   160
         Top             =   6420
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
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
         MICON           =   "Configuracao_Geral.frx":FBC0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdPrepara2 
         Height          =   315
         Left            =   -68580
         TabIndex        =   164
         Top             =   6420
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
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
         MICON           =   "Configuracao_Geral.frx":FBDC
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
         Height          =   315
         Left            =   -73860
         TabIndex        =   165
         Top             =   5820
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Configuracao_Geral.frx":FBF8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   315
         Left            =   -72780
         TabIndex        =   167
         Top             =   5820
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Localizar"
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
         MICON           =   "Configuracao_Geral.frx":FC14
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
         Height          =   315
         Left            =   -69780
         TabIndex        =   168
         Top             =   5820
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Marcar"
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
         MICON           =   "Configuracao_Geral.frx":FC30
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdDesmarcar 
         Height          =   315
         Left            =   -68760
         TabIndex        =   169
         Top             =   5820
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Desmarcar"
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
         MICON           =   "Configuracao_Geral.frx":FC4C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdDesmarcarTodos 
         Height          =   315
         Left            =   -67680
         TabIndex        =   170
         Top             =   5820
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Desmarcar Todos"
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
         MICON           =   "Configuracao_Geral.frx":FC68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Declarar Recebedor:"
         Height          =   195
         Left            =   -69660
         TabIndex        =   209
         Top             =   1920
         Width           =   1485
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Parcelas/Impressão/Vendas:"
         Height          =   195
         Left            =   -69660
         TabIndex        =   206
         Top             =   1200
         Width           =   2070
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Impressoras Instaladas:"
         Height          =   195
         Left            =   -69660
         TabIndex        =   204
         Top             =   540
         Width           =   1650
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Impressora Normal:"
         Height          =   195
         Left            =   -72480
         TabIndex        =   203
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Temporário"
         Height          =   195
         Left            =   -69540
         TabIndex        =   163
         Top             =   6180
         Width           =   795
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Certo"
         Height          =   195
         Left            =   -70980
         TabIndex        =   161
         Top             =   6180
         Width           =   375
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fantasia"
         Height          =   195
         Left            =   -74880
         TabIndex        =   154
         Top             =   5220
         Width           =   600
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mês"
         Height          =   195
         Left            =   -74880
         TabIndex        =   152
         Top             =   6180
         Width           =   300
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ano"
         Height          =   195
         Left            =   -73320
         TabIndex        =   150
         Top             =   6180
         Width           =   285
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   -68220
         TabIndex        =   148
         Top             =   5220
         Width           =   405
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razão"
         Height          =   195
         Left            =   -71940
         TabIndex        =   147
         Top             =   5220
         Width           =   465
      End
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Confirmar fechamendo da venda:"
      Height          =   195
      Left            =   540
      TabIndex        =   54
      Top             =   5040
      Width           =   2355
   End
End
Attribute VB_Name = "Configuracao_Geral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private moCombo As cComboHelper
Private Caminho As String
Dim oCfg As ConfigItem
Dim sSQL As String
Dim r As ADODB.Recordset
Dim i As Integer



Private Sub AtualizarBackup()
Dim sSQL As String
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtPastaBackup.Text & "' WHERE (config_nome = 'BACKUP_PASTA');"
   dbData.Execute sSQL
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtBackupHorario.Text & "' WHERE (config_nome = 'BACKUP_HORARIO');"
   dbData.Execute sSQL
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(chkBackupAutomatico.Value) & "' WHERE (config_nome = 'BACKUP_AUTOMATICO');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("BACKUP_PASTA").Value = txtPastaBackup.Text
   sysConfig("BACKUP_HORARIO").Value = txtBackupHorario.Text
   sysConfig("BACKUP_AUTOMATICO").Value = Abs(chkBackupAutomatico.Value)
End Sub

Private Sub AtualizarBalanca()
Dim sSQL As String
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & ComBalancaModelo.Text & "' WHERE (config_nome = 'PDV_BALANCA_MODELO');"
   dbData.Execute sSQL
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & cboBalancaPorta.Text & "' WHERE (config_nome = 'PDV_BALANCA_PORTA');"
   dbData.Execute sSQL
      
   'Atualiza a configuração carregada na memória
   sysConfig("PDV_BALANCA_MODELO").Value = ComBalancaModelo.Text
   sysConfig("PDV_BALANCA_PORTA").Value = cboBalancaPorta.Text
End Sub

Private Sub AtualizarHabilitarAluguel()
Dim sSQL As String, bOpt As Boolean

If cboHabilitarAluguel.Text = "SIM" Then
    bOpt = True
ElseIf cboHabilitarAluguel.Text = "NÃO" Then
    bOpt = False
End If

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'ALUGUEL');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("ALUGUEL").Value = Abs(bOpt)
End Sub

Private Sub AtualizarHabilitarOS()
Dim sSQL As String, bOpt As Boolean

If cboHabilitarOS.Text = "SIM" Then
    bOpt = True
ElseIf cboHabilitarOS.Text = "NÃO" Then
    bOpt = False
End If

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'OS');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("OS").Value = Abs(bOpt)
End Sub

Private Sub AP_Mostrar_Copia()
   Set oCfg = sysConfig("COPIAS_AP")
   txtNumCopiaAP.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub AtualizaFundoPDV()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtCaminho.Text & "' WHERE (config_nome = 'FUNDO_PDV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("FUNDO_PDV").Value = txtCaminho.Text

End Sub

Private Sub AtualizarAPConfImpressao()
   Dim sSQL As String, bOpt As Boolean
   
   If cboConfirmaImpressaoAP.Text = "SIM" Then
       bOpt = True
   ElseIf cboConfirmaImpressaoAP.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_IMPRESSAO_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_IMPRESSAO_AP").Value = Abs(bOpt)
End Sub

Private Sub AtualizarAPEntregar()
Dim sSQL As String, bOpt As Boolean

If cboNotaEntregaAP.Text = "SIM" Then
    bOpt = True
ElseIf cboNotaEntregaAP.Text = "NÃO" Then
    bOpt = False
End If

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'ENTREGA_AP');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("ENTREGA_AP").Value = Abs(bOpt)
End Sub

Private Sub AtualizarAPImprimir()
   Dim sSQL As String, bOpt As Boolean
   
   If cboImprimirAP.Text = "SIM" Then
       bOpt = True
   ElseIf cboImprimirAP.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'IMP_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMP_AP").Value = Abs(bOpt)
End Sub

Private Sub Atualizar_ValorCartaoCredito()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtValorAcrescCredito.Text & "' WHERE (config_nome = 'ACRESC_CREDITO_VALOR');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("ACRESC_CREDITO_VALOR").Value = txtValorAcrescCredito.Text
End Sub

Private Sub Atualizar_ValorCartaoDebito()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtValorAcrescDebito.Text & "' WHERE (config_nome = 'ACRESC_DEBITO_VALOR');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("ACRESC_DEBITO_VALOR").Value = txtValorAcrescDebito.Text
End Sub

Private Sub AtualizarAPDesc()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtValorDescAP.Text & "' WHERE (config_nome = 'DESC_AP');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("DESC_AP").Value = txtValorDescAP.Text
End Sub

Private Sub AtualizarAPNumCopias()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtNumCopiaAP.Text & "' WHERE (config_nome = 'COPIAS_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("COPIAS_AP").Value = txtNumCopiaAP.Text
End Sub

Private Sub AtualizarAPTipoImpressao()
   Dim sSQL As String
   
   If txtTipoImpressaoAP.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoImpressaoAP.Text & "' WHERE (config_nome = 'IMPRIMIR_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMPRIMIR_AP").Value = txtTipoImpressaoAP.Text
End Sub

Private Sub AtualizaNFCeConfImpressao()
   Dim sSQL As String, bOpt As Boolean
   
   If cboConfirmaImpressaoAV.Text = "SIM" Then
       bOpt = True
   ElseIf cboConfirmaImpressaoAV.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_IMPRESSAO_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_IMPRESSAO_AV").Value = Abs(bOpt)
End Sub

Private Sub AtualizarNFCeCombinarImp()
Dim sSQL As String
   
If cboCombinarImpNFCe.Text = "" Then Exit Sub
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboCombinarImpNFCe & "' WHERE (config_nome = 'COMBINARIMPNFCE');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("COMBINARIMPNFCE").Value = cboCombinarImpNFCe
End Sub
Private Sub AtualizarNFCeConfPrazo()
Dim sSQL As String
   
If cboConfirmaPrazoNFCe.Text = "" Then Exit Sub
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboConfirmaPrazoNFCe & "' WHERE (config_nome = 'CONFPRAZONFCE');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("CONFCPFNFCE").Value = cboConfirmaPrazoNFCe
End Sub

Private Sub AtualizarNFCeConfCPF()
Dim sSQL As String
   
If cboConfirmaCPFNFCe.Text = "" Then Exit Sub
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboConfirmaCPFNFCe & "' WHERE (config_nome = 'CONFCPFNFCE');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("CONFCPFNFCE").Value = cboConfirmaCPFNFCe
End Sub

Private Sub AtualizarNFCeConfImpressao()
Dim sSQL As String
   
If cboConfirmaImpressaoNFCe.Text = "" Then Exit Sub
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboConfirmaImpressaoNFCe & "' WHERE (config_nome = 'CONFIMPNFCE');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("CONF_IMPRESSAO_AV").Value = cboConfirmaImpressaoNFCe
End Sub
Private Sub AtualizarAVConfImpressao()
   Dim sSQL As String, bOpt As Boolean
   
   If cboConfirmaImpressaoAV.Text = "SIM" Then
       bOpt = True
   ElseIf cboConfirmaImpressaoAV.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_IMPRESSAO_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_IMPRESSAO_AV").Value = Abs(bOpt)
End Sub

Private Sub AtualizarAVEntregar()
   Dim sSQL As String, bOpt As Boolean
   
   If cboNotaEntregaAV.Text = "SIM" Then
       bOpt = True
   ElseIf cboNotaEntregaAV.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'ENTREGA_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("ENTREGA_AV").Value = Abs(bOpt)
End Sub

Private Sub AtualizarNFCeImprimir()
Dim sSQL As String

If cboImprimirNFCe.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboImprimirNFCe.Text & "' WHERE (config_nome = 'IMPRIMINFCE');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("IMP_AV").Value = cboImprimirNFCe.Text
End Sub

Private Sub AtualizarAVImprimir()
   Dim sSQL As String, bOpt As Boolean
   
   If cboImprimirAV.Text = "SIM" Then
       bOpt = True
   ElseIf cboImprimirAV.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'IMP_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMP_AV").Value = Abs(bOpt)
End Sub
Private Sub AtualizarDescGradual()
Dim sSQL As String
 
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtDescMargemAV1.Text & "' WHERE (config_nome = 'DESC_MARGEM_AV1');"
dbData.Execute sSQL

sSQL = "UPDATE configuracao SET config_valor = '" & txtDescAV1.Text & "' WHERE (config_nome = 'DESC_AV1');"
dbData.Execute sSQL

sSQL = "UPDATE configuracao SET config_valor = '" & txtDescAP1.Text & "' WHERE (config_nome = 'DESC_AP1');"
dbData.Execute sSQL

sSQL = "UPDATE configuracao SET config_valor = '" & txtDescMargemAV2.Text & "' WHERE (config_nome = 'DESC_MARGEM_AV2');"
dbData.Execute sSQL

sSQL = "UPDATE configuracao SET config_valor = '" & txtDescAV2.Text & "' WHERE (config_nome = 'DESC_AV2');"
dbData.Execute sSQL

sSQL = "UPDATE configuracao SET config_valor = '" & txtDescAP2.Text & "' WHERE (config_nome = 'DESC_AP2');"
dbData.Execute sSQL

sSQL = "UPDATE configuracao SET config_valor = '" & txtDescMargemAV3.Text & "' WHERE (config_nome = 'DESC_MARGEM_AV3');"
dbData.Execute sSQL

sSQL = "UPDATE configuracao SET config_valor = '" & txtDescAV3.Text & "' WHERE (config_nome = 'DESC_AV3');"
dbData.Execute sSQL

sSQL = "UPDATE configuracao SET config_valor = '" & txtDescAP3.Text & "' WHERE (config_nome = 'DESC_AP3');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("DESC_MARGEM_AV1").Value = txtDescMargemAV1.Text
sysConfig("DESC_AV1").Value = txtDescAV1.Text
sysConfig("DESC_AP1").Value = txtDescAP1.Text
sysConfig("DESC_MARGEM_AV2").Value = txtDescMargemAV2.Text
sysConfig("DESC_AV2").Value = txtDescAV2.Text
sysConfig("DESC_AP2").Value = txtDescAP2.Text
sysConfig("DESC_MARGEM_AV3").Value = txtDescMargemAV3.Text
sysConfig("DESC_AV3").Value = txtDescAV3.Text
sysConfig("DESC_AP3").Value = txtDescAP3.Text
End Sub

Private Sub AtualizarAVDesc()
Dim sSQL As String
 
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtValorDescAV.Text & "' WHERE (config_nome = 'DESC_AV');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("DESC_AV").Value = txtValorDescAV.Text
End Sub

Private Sub AtualizarTipoJuros()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoJuros.Text & "' WHERE (config_nome = 'TIPO_JUROS');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TIPO_JUROS").Value = txtTipoJuros.Text
End Sub

Private Sub AtualizarTipoImpressaoParcelas()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtFormaParcelasImpressao.Text & "' WHERE (config_nome = 'TIPOIMPRESSAOPARCELAS');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TIPOIMPRESSAOPARCELAS").Value = txtFormaParcelasImpressao.Text
End Sub
Private Sub AtualizarTipoCaixa()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoCaixa.Text & "' WHERE (config_nome = 'TIPOCAIXA');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TIPOCAIXA").Value = txtTipoCaixa.Text
End Sub
Private Sub AtualizarValorVenda()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtValorVenda.Text & "' WHERE (config_nome = 'TIPOVALORVENDA');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TIPOVALORVENDA").Value = txtValorVenda.Text
End Sub
Private Sub AtualizarAVNumCopias()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtNumCopia.Text & "' WHERE (config_nome = 'COPIAS_AV');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("COPIAS_AV").Value = txtNumCopia.Text
End Sub
Private Sub AtualizarAVTipoImpressao()
   Dim sSQL As String
   
   If txtTipoCadastroProduto.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoImpressaoAV.Text & "' WHERE (config_nome = 'IMPRIMIR_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMPRIMIR_AV").Value = txtTipoImpressaoAV.Text
End Sub

Private Sub AtualizarLogoffAutomatico()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboLogoffAutomatico.Text & "' WHERE (config_nome = 'LOGOFFAUTOMATICO');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("LOGOFFAUTOMATICO").Value = cboLogoffAutomatico.Text

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtTempoLogoff.Text & "' WHERE (config_nome = 'TEMPOLOGOFF');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TEMPOLOGOFF").Value = txtTempoLogoff.Text

txtTempoLogoff.Enabled = False
End Sub

Private Sub AtualizarBloquearCliente()
Dim sSQL As String, bOpt As Boolean

If cboClienteDebito.Text = "SIM" Then
    bOpt = True
ElseIf cboClienteDebito.Text = "NÃO" Then
    bOpt = False
End If

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'BLOQUEIAR_CLIENTE');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("BLOQUEIAR_CLIENTE").Value = Abs(bOpt)

If txtQuantDiasBloqueiar.Text = "" Then txtQuantDiasBloqueiar.Text = "0"

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtQuantDiasBloqueiar.Text & "' WHERE (config_nome = 'DIAS_BLOQUEIO');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("DIAS_BLOQUEIO").Value = txtQuantDiasBloqueiar.Text

'txtQuantDiasBloqueiar.Enabled = False
End Sub

Private Sub AtualizarConfFechamentoAP()
   Dim sSQL As String, bOpt As Boolean
   
   If cboConfFerchamentoAP.Text = "SIM" Then
       bOpt = True
   ElseIf cboConfFerchamentoAP.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_FECHAMENTO_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_FECHAMENTO_AP").Value = Abs(bOpt)
End Sub
Private Sub AtualizarConfFechamentoORC()
   Dim sSQL As String, bOpt As Boolean
   
   If cboConfFerchamentoORC.Text = "SIM" Then
       bOpt = True
   ElseIf cboConfFerchamentoORC.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_FECHAMENTO_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_FECHAMENTO_ORC").Value = Abs(bOpt)
End Sub

Private Sub AtualizarConfFechamentoAV()
   Dim sSQL As String, bOpt As Boolean
   
   If cboConfFerchamentoAV.Text = "SIM" Then
       bOpt = True
   ElseIf cboConfFerchamentoAV.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_FECHAMENTO_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_FECHAMENTO_AV").Value = Abs(bOpt)
End Sub

Private Sub AtualizarQuantDigitosBalanca()
Dim sSQL As String, bOpt As Boolean

If txtQtdeDigitosBalanca.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtQtdeDigitosBalanca.Text & "' WHERE (config_nome = 'QTDEDIGITOSBALANCA');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("QTDEDIGITOSBALANCA").Value = txtQtdeDigitosBalanca.Text
End Sub

Private Sub AtualizarQuantDigitosEtiquetas()
Dim sSQL As String, bOpt As Boolean

If txtQtdeDigitosEtiqueta.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtQtdeDigitosEtiqueta.Text & "' WHERE (config_nome = 'QTDEDIGITOSETIQUETAS');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("QTDEDIGITOSETIQUETAS").Value = txtQtdeDigitosEtiqueta.Text
End Sub
Private Sub AtualizarIniciaisBalanca()
Dim sSQL As String, bOpt As Boolean

If txtIniciaisBalanca.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtIniciaisBalanca.Text & "' WHERE (config_nome = 'INICIAISBALANCA');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("INICIAISBALANCA").Value = txtIniciaisBalanca.Text
End Sub
Private Sub AtualizarIniciaisEtiquetas()
Dim sSQL As String, bOpt As Boolean

If txtIniciaisEtiqueta.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtIniciaisEtiqueta.Text & "' WHERE (config_nome = 'INICIAISETIQUETAS');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("INICIAISETIQUETAS").Value = txtIniciaisEtiqueta.Text
End Sub
Private Sub AtualizarIndentPDV()
   Dim sSQL As String, bOpt As Boolean
   
   If txtTipoIndPDV.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoIndPDV.Text & "' WHERE (config_nome = 'IDENT_PDV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IDENT_PDV").Value = txtTipoIndPDV.Text
End Sub


Private Sub AtualizarConfCartaoCredito()
   Dim sSQL As String, bOpt As Boolean
   
   If cboAcrescCreditoConf.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & cboAcrescCreditoConf.Text & "' WHERE (config_nome = 'ACRESC_CREDITO');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("ACRESC_CREDITO").Value = cboAcrescCreditoConf.Text
End Sub

Private Sub AtualizarConfCartaoDebito()
   Dim sSQL As String, bOpt As Boolean
   
   If txtAcrescDebitoConf.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtAcrescDebitoConf.Text & "' WHERE (config_nome = 'ACRESC_DEBITO');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("ACRESC_DEBITO").Value = txtAcrescDebitoConf.Text
End Sub
Private Sub AtualizarTipoLogin()
Dim sSQL As String

If cboTipoLogin.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboTipoLogin.Text & "' WHERE (config_nome = 'TIPOLOGIN');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TIPOLOGIN").Value = cboTipoLogin.Text
End Sub
Private Sub AtualizarIdentMaquina()
   Dim sSQL As String, bOpt As Boolean
   
   If cboIdentificarMaquina.Text = "SIM" Then
       bOpt = True
   ElseIf cboIdentificarMaquina.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'IDENT_MAQ');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IDENT_MAQ").Value = Abs(bOpt)
End Sub

Private Sub AtualizarImgCupom()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtCaminhoCupom.Text & "' WHERE (config_nome = 'LOGO_CUPOM');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("LOGO_CUPOM").Value = txtCaminhoCupom.Text
End Sub

Private Sub AtualizarIncluirPreco()
Dim sSQL As String, bOpt As Boolean

If cboIncluirPrecos.Text = "SIM" Then
    bOpt = True
ElseIf cboIncluirPrecos.Text = "NÃO" Then
    bOpt = False
End If

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'INCLUIR_PRECO');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("INCLUIR_PRECO").Value = Abs(bOpt)
End Sub

Private Sub AtualizarJuros()
Dim sSQL As String
    
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtJurosMes.Text & "' WHERE (config_nome = 'JUROS_MES');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("JUROS_MES").Value = txtJurosMes.Text
    
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtJuroDia.Text & "' WHERE (config_nome = 'JUROS_DIA');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("JUROS_DIA").Value = txtJuroDia.Text
End Sub

Private Sub AtualizarORCConfImpressao()
   Dim sSQL As String, bOpt As Boolean
   
   If cboConfirmaImpressaoORC.Text = "SIM" Then
       bOpt = True
   ElseIf cboConfirmaImpressaoORC.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_IMPRESSAO_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_IMPRESSAO_ORC").Value = Abs(bOpt)
End Sub

Private Sub AtualizarORCImprimir()
   Dim sSQL As String, bOpt As Boolean
   
   If cboImprimirORC.Text = "SIM" Then
       bOpt = True
   ElseIf cboImprimirORC.Text = "NÃO" Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'IMP_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMP_ORC").Value = Abs(bOpt)
End Sub

Private Sub AtualizarTipoDesconto()
Dim sSQL As String
Dim vTipoDesc As Integer
Dim vLimiteGerente As String
Dim vDescCredito As String
Dim vDescDebito As String

If cboTipoDesconto.Text = "MANUAL" Then
    vTipoDesc = "1"
ElseIf cboTipoDesconto.Text = "FIXO" Then
    vTipoDesc = "2"
ElseIf cboTipoDesconto.Text = "GRADUAL" Then
    vTipoDesc = "3"
End If

'Tipo de Desconto
sSQL = "UPDATE configuracao SET config_valor = '" & vTipoDesc & "' WHERE (config_nome = 'TIPODESCONTO');"
dbData.Execute sSQL

sysConfig("TIPODESCONTO").Value = vTipoDesc


'Se vai limitar o desconto
sSQL = "UPDATE configuracao SET config_valor = '" & Abs(chkLimiteDesc.Value) & "' WHERE (config_nome = 'LIMITEDESCONTO');"
dbData.Execute sSQL

sysConfig("LIMITEDESCONTO").Value = Abs(chkLimiteDesc.Value)


'Se vai pedir a senha do gerente para liberar desconto
If chkLimiteGerente.Value = Checked Then
    vLimiteGerente = "SIM"
Else
    vLimiteGerente = "NÃO"
End If

sSQL = "UPDATE configuracao SET config_valor = '" & vLimiteGerente & "' WHERE (config_nome = 'LIMITEGERENTE');"
dbData.Execute sSQL

sysConfig("LIMITEGERENTE").Value = vLimiteGerente


'Se vai ter desconto nas vendas de cartão de crédito
If chkCartaoCredito.Value = Checked Then
    vDescCredito = "SIM"
Else
    vDescCredito = "NÃO"
End If

sSQL = "UPDATE configuracao SET config_valor = '" & vDescCredito & "' WHERE (config_nome = 'DESCCARTAOCREDITO');"
dbData.Execute sSQL

sysConfig("DESCCARTAOCREDITO").Value = vDescCredito

'Se vai ter desconto nas vendas de cartão de crédito
If chkCartaoDebito.Value = Checked Then
    vDescDebito = "SIM"
Else
    vDescDebito = "NÃO"
End If

'Se vai ter desconto nas vendas de cartão de dédito
sSQL = "UPDATE configuracao SET config_valor = '" & vDescDebito & "' WHERE (config_nome = 'DESCCARTAODEBITO');"
dbData.Execute sSQL

sysConfig("DESCCARTAODEDITO").Value = vDescDebito

End Sub



Private Sub AtualizarORCNumCopias()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtNumCopiaORC.Text & "' WHERE (config_nome = 'COPIAS_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("COPIAS_ORC").Value = txtNumCopiaORC.Text
End Sub

Private Sub AtualizarORCTipoImpressao()
   Dim sSQL As String
   
   If txtTipoImpressaoORC.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoImpressaoORC.Text & "' WHERE (config_nome = 'IMPRIMIR_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMPRIMIR_ORC").Value = txtTipoImpressaoORC.Text
End Sub

Private Sub AtualizarSegurancaAvancada()
Dim sSQL As String

If cboSegAvancado.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboSegAvancado.Text & "' WHERE (config_nome = 'SEGURANCAAVANCADA');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("SEGURANCAAVANCADA").Value = cboSegAvancado.Text
End Sub

Private Sub AtualizarTipoRecHaver()
Dim sSQL As String

If cboTipoReciboHaver.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboTipoReciboHaver.Text & "' WHERE (config_nome = 'TIPORECHAVER');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TIPORECHAVER").Value = cboTipoReciboHaver.Text
End Sub
Private Sub AtualizarTipoRecPgto()
Dim sSQL As String

If cboTipoReciboPgto.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboTipoReciboPgto.Text & "' WHERE (config_nome = 'TIPORECPGTO');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TIPORECPGTO").Value = cboTipoReciboPgto.Text
End Sub
Private Sub AtualizarTipoEmpresa()
Dim sSQL As String

If txtTipoCadastroProduto.Text = "" Then Exit Sub

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoCadastroProduto.Text & "' WHERE (config_nome = 'TIPO_EMPRESA');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("TIPO_EMPRESA").Value = txtTipoCadastroProduto.Text
End Sub
Private Sub AtualizarLimitarCompra()
'Dim sSQL As String
If cboLimitarCompra.Text = "" Then Exit Sub
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboLimitarCompra.Text & "' WHERE (config_nome = 'LIMITARCOMPRA');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("LIMITARCOMPRA").Value = cboLimitarCompra.Text
End Sub

Private Sub AtualizarCashback()
':::::::::::::::::::CASHBACK A VISTA
If cboCashbackVista.Text = "" Then Exit Sub
   
'SIM ou NÂO
sSQL = "UPDATE configuracao SET config_valor = '" & cboCashbackVista.Text & "' WHERE (config_nome = 'CASHBACKAV');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("CASHBACKAV").Value = cboCashbackVista.Text


   'PORCENTAGEM AV
   sSQL = "UPDATE configuracao SET config_valor = '" & txtCashbackAV.Text & "' WHERE (config_nome = 'CASHBACKVALORAV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CASHBACKVALORAV").Value = txtCashbackAV.Text



'::::::::::::::::::CASHBACK A PRAZO
If cboCashbackPrazo.Text = "" Then Exit Sub
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboCashbackPrazo.Text & "' WHERE (config_nome = 'CASHBACKAP');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("CASHBACKAP").Value = cboCashbackPrazo.Text


   'PORCENTAGEM AP
   sSQL = "UPDATE configuracao SET config_valor = '" & txtCashbackAP.Text & "' WHERE (config_nome = 'CASHBACKVALORAP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CASHBACKVALORAP").Value = txtCashbackAP.Text


'::::::::::::::::::VALIDADE
sSQL = "UPDATE configuracao SET config_valor = '" & txtCashbackValidade.Text & "' WHERE (config_nome = 'CASHBACKVALIDADE');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("CASHBACKVALIDADE").Value = txtCashbackValidade.Text


End Sub


Private Sub AtualizarMultiplasRef()
'Dim sSQL As String
If cboMultiplasRef.Text = "" Then Exit Sub
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboMultiplasRef.Text & "' WHERE (config_nome = 'MULTIPLASREF');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("MULTIPLASREF").Value = cboMultiplasRef.Text
End Sub
Private Sub AtualizarDesativarClientes()
Dim sSQL As String
If cboDesativarClientes.Text = "" Then Exit Sub
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboDesativarClientes.Text & "' WHERE (config_nome = 'DESATIVARCLIENTES');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("DESATIVARCLIENTES").Value = cboDesativarClientes.Text


If txtQuantDiasDesativar.Text = "" Then txtQuantDiasDesativar.Text = "0"
   
'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & txtQuantDiasDesativar.Text & "' WHERE (config_nome = 'QUANTDIASDESATIVAR');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("QUANTDIASDESATIVAR").Value = txtQuantDiasDesativar.Text

End Sub
Private Sub AtualizarTipoOS()
Dim sSQL As String
If cboTipoOS.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & cboTipoOS.Text & "' WHERE (config_nome = 'TIPO_OS');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("TIPO_OS").Value = cboTipoOS.Text
End Sub
Private Sub AtualizarDeclararRecebedor()
Dim sSQL As String

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & cboDeclararRecebedor.Text & "' WHERE (config_nome = 'DECLARARRECEBEDOR');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("DECLARARRECEBEDOR").Value = Abs(bOpt)
End Sub

Private Sub AtualizarVenderNegativo()
Dim sSQL As String, bOpt As Boolean

If cboEstoqueNegativo.Text = "SIM" Then
    bOpt = True
ElseIf cboEstoqueNegativo.Text = "NÃO" Then
    bOpt = False
End If

'Atualiza a base de dados
sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'ESTOQUE_NEGATIVO');"
dbData.Execute sSQL

'Atualiza a configuração carregada na memória
sysConfig("ESTOQUE_NEGATIVO").Value = Abs(bOpt)
End Sub
Private Sub LimparEmpresa()
txtFantasia.Text = ""
txtRazao.Text = ""
mskCPF.Mask = ""
mskCPF.Text = ""
End Sub

Private Sub MostrarEmpresa()
'sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado FROM  empresas_desbloueio ORDER BY FANTASIA;"
'Set r = dbData.OpenRecordset(sSQL)

'FormatarGrid r
End Sub

Private Sub ORC_Mostrar_Copia()
   Set oCfg = sysConfig("COPIAS_ORC")
   txtNumCopiaORC.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_TipoDesconto()
Set oCfg = sysConfig("TIPODESCONTO")
Dim vTipoDesc As Integer
vTipoDesc = oCfg.Value
Set oCfg = Nothing

If vTipoDesc = 1 Then
   cboTipoDesconto.Text = "MANUAL"
ElseIf vTipoDesc = 2 Then
   cboTipoDesconto.Text = "FIXO"
ElseIf vTipoDesc = 3 Then
   cboTipoDesconto.Text = "GRADUAL"
End If

Set oCfg = sysConfig("LIMITEDESCONTO")
chkLimiteDesc.Value = Abs(oCfg.Value)

LimiteGerente_Mostrar
End Sub

Private Sub Mostrar_TipoJuros()
Set oCfg = sysConfig("TIPO_JUROS")
txtTipoJuros.Text = oCfg.Value
Set oCfg = Nothing

If txtTipoJuros.Text = 1 Then
   cboTipoJuros.Text = "Parcela"
ElseIf txtTipoJuros.Text = 0 Then
   cboTipoJuros.Text = "Restante"
End If
End Sub
Private Sub Mostrar_TipoImpressaoParcelas()
Set oCfg = sysConfig("TIPOIMPRESSAOPARCELAS")
txtFormaParcelasImpressao.Text = oCfg.Value
Set oCfg = Nothing

If txtFormaParcelasImpressao.Text = 1 Then
   cboFormaParcelasImpressao.Text = "1 - Resumido"
ElseIf txtFormaParcelasImpressao.Text = 2 Then
   cboFormaParcelasImpressao.Text = "2 - Detalhado"
End If
End Sub
Private Sub Mostrar_TipoCaixa()
Set oCfg = sysConfig("TIPOCAIXA")
txtTipoCaixa.Text = oCfg.Value
Set oCfg = Nothing

If txtTipoCaixa.Text = 1 Then
   cboTipoCaixa.Text = "1 - Único"
ElseIf txtTipoCaixa.Text = 2 Then
   cboTipoCaixa.Text = "2 - Multiplus"
End If
End Sub
Private Sub Mostrar_TipoValorVenda()
Set oCfg = sysConfig("TIPOVALORVENDA")
txtValorVenda.Text = oCfg.Value
Set oCfg = Nothing

If txtValorVenda.Text = 1 Then
   cboValorVenda.Text = "1 - Único"
ElseIf txtValorVenda.Text = 2 Then
   cboValorVenda.Text = "2 - Multiplus"
End If
End Sub


Private Sub AV_Mostrar_Copia()
   Set oCfg = sysConfig("COPIAS_AV")
   txtNumCopia.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_LogoCupom()
   Set oCfg = sysConfig("LOGO_CUPOM")
   txtCaminhoCupom.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_Fundo()
   Set oCfg = sysConfig("FUNDO_PDV")
   txtCaminho.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_Aluguel()
Set oCfg = sysConfig("Aluguel")

If CBool(oCfg.Value) = True Then
     cboHabilitarAluguel.Text = "SIM"
 Else
     cboHabilitarAluguel.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub Mostrar_OS()
Set oCfg = sysConfig("OS")

If CBool(oCfg.Value) = True Then
     cboHabilitarOS.Text = "SIM"
 Else
     cboHabilitarOS.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub
Private Sub Mostrar_Conf_CartaoCredito()
Set oCfg = sysConfig("ACRESC_CREDITO")

cboAcrescCreditoConf.Text = oCfg.Value
If cboAcrescCreditoConf.Text = 1 Then
   cboAcrescCredito.Text = "SIM"
ElseIf cboAcrescCreditoConf.Text = 0 Then
   cboAcrescCredito.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub Mostrar_Conf_CartaoDebito()
Set oCfg = sysConfig("ACRESC_DEBITO")

txtAcrescDebitoConf.Text = oCfg.Value
If txtAcrescDebitoConf.Text = 1 Then
   cboAcrescDebito.Text = "SIM"
ElseIf txtAcrescDebitoConf.Text = 0 Then
   cboAcrescDebito.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub
Private Sub Mostrar_Tipo_Identificacao()
   Set oCfg = sysConfig("IDENT_PDV")
   
   txtTipoIndPDV.Text = oCfg.Value
   If txtTipoIndPDV.Text = 1 Then
      cboIdentVendedor.Text = "LOGIN"
   ElseIf txtTipoIndPDV.Text = 2 Then
      cboIdentVendedor.Text = "CÓD. FUNC."
   End If
   
   Set oCfg = Nothing
End Sub
Private Sub MostrarTipoLogin()
Set oCfg = sysConfig("TIPOLOGIN")

If Not IsNull(oCfg.Value) Then cboTipoLogin.Text = oCfg.Value

Set oCfg = Nothing
End Sub
Private Sub Mostrar_SegurancaAvancada()
Set oCfg = sysConfig("SEGURANCAAVANCADA")

If Not IsNull(oCfg.Value) Then cboSegAvancado.Text = oCfg.Value

Set oCfg = Nothing
End Sub

Private Sub Mostrar_TipoRecHaver()
Set oCfg = sysConfig("TIPORECHAVER")

If Not IsNull(oCfg.Value) Then cboTipoReciboHaver.Text = oCfg.Value

Set oCfg = Nothing
End Sub

Private Sub Mostrar_TipoRecPgto()
Set oCfg = sysConfig("TIPORECPGTO")

If Not IsNull(oCfg.Value) Then cboTipoReciboPgto.Text = oCfg.Value

Set oCfg = Nothing
End Sub


Private Sub Mostrar_Tipo_OS()
   Set oCfg = sysConfig("TIPO_OS")
   
   cboTipoOS.Text = oCfg.Value
   
   Set oCfg = Nothing
End Sub
Private Sub Mostrar_Tipo_Empresa()
   Set oCfg = sysConfig("TIPO_EMPRESA")
   
   txtTipoCadastroProduto.Text = oCfg.Value
   If txtTipoCadastroProduto.Text = 1 Then
      cboTipoEmpresa.Text = "Varejo"
   ElseIf txtTipoCadastroProduto.Text = 2 Then
      cboTipoEmpresa.Text = "Farmacia"
   ElseIf txtTipoCadastroProduto.Text = 3 Then
      cboTipoEmpresa.Text = "Restaurante/Lannchonete"
   ElseIf txtTipoCadastroProduto.Text = 4 Then
      cboTipoEmpresa.Text = "Sapataria/Vestuário"
   ElseIf txtTipoCadastroProduto.Text = 5 Then
      cboTipoEmpresa.Text = "Autopeça/Motopeça"
   ElseIf txtTipoCadastroProduto.Text = 6 Then
      cboTipoEmpresa.Text = "Academia"
   ElseIf txtTipoCadastroProduto.Text = 7 Then
      cboTipoEmpresa.Text = "Escola"
   ElseIf txtTipoCadastroProduto.Text = 8 Then
      cboTipoEmpresa.Text = "Consultorio"
   Else
      Exit Sub
   End If
   
   Set oCfg = Nothing
End Sub
Private Sub MostrarIdentMaquina()
   Set oCfg = sysConfig("IDENT_MAQ")
   
   If CBool(oCfg.Value) = True Then
      cboIdentificarMaquina.Text = "SIM"
   Else
      cboIdentificarMaquina.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarIncluirPreco()
Set oCfg = sysConfig("INCLUIR_PRECO")

If CBool(oCfg.Value) = True Then
     cboIncluirPrecos.Text = "SIM"
Else
     cboIncluirPrecos.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub MostrarEstoqueNegativo()
   Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
   
   If CBool(oCfg.Value) = True Then
      cboEstoqueNegativo.Text = "SIM"
   Else
      cboEstoqueNegativo.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub
Private Sub AP_MostrarEntrega()
   Set oCfg = sysConfig("ENTREGA_AP")
    
   If CBool(oCfg.Value) = True Then
      cboNotaEntregaAP.Text = "SIM"
   Else
      cboNotaEntregaAP.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AV_MostrarEntrega()
Set oCfg = sysConfig("ENTREGA_AV")

If CBool(oCfg.Value) = True Then
   cboNotaEntregaAV.Text = "SIM"
Else
   cboNotaEntregaAV.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub DescDebito_Mostrar()
Set oCfg = sysConfig("DESCCARTAODEDITO")

'chkLimiteGerente.Value = Abs(oCfg.Value)
If oCfg.Value = "SIM" Then
   chkCartaoDebito.Value = Checked
Else
   chkCartaoDebito.Value = Unchecked
End If

Set oCfg = Nothing
End Sub

Private Sub DescCredito_Mostrar()
Set oCfg = sysConfig("DESCCARTAOCREDITO")

'chkLimiteGerente.Value = Abs(oCfg.Value)
If oCfg.Value = "SIM" Then
   chkCartaoCredito.Value = Checked
Else
   chkCartaoCredito.Value = Unchecked
End If

Set oCfg = Nothing
End Sub

Private Sub LimiteGerente_Mostrar()
Set oCfg = sysConfig("LIMITEGERENTE")

'chkLimiteGerente.Value = Abs(oCfg.Value)
If oCfg.Value = "SIM" Then
   chkLimiteGerente.Value = Checked
Else
   chkLimiteGerente.Value = Unchecked
End If

Set oCfg = Nothing
End Sub

Private Sub NFCe_MostrarImp()
Set oCfg = sysConfig("IMPRIMINFCE")

If oCfg.Value = "SIM" Then
   cboImprimirNFCe.Text = "SIM"
Else
   cboImprimirNFCe.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub AP_MostrarImp()
   Set oCfg = sysConfig("IMP_AP")
   
   If CBool(oCfg.Value) = True Then
      cboImprimirAP.Text = "SIM"
   Else
      cboImprimirAP.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub


Private Sub ORC_MostrarImp()
   Set oCfg = sysConfig("IMP_ORC")
      
   If CBool(oCfg.Value) = True Then
      cboImprimirORC.Text = "SIM"
   Else
      cboImprimirORC.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AV_MostrarImp()
   Set oCfg = sysConfig("IMP_AV")
   
   If CBool(oCfg.Value) = True Then
      cboImprimirAV.Text = "SIM"
   Else
      cboImprimirAV.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AP_MostrarConfImpressao()
   Set oCfg = sysConfig("CONF_IMPRESSAO_AP")
   
   If CBool(oCfg.Value) = True Then
      cboConfirmaImpressaoAP.Text = "SIM"
   Else
      cboConfirmaImpressaoAP.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarLogoffAutomatico()
Set oCfg = sysConfig("LOGOFFAUTOMATICO")

If oCfg.Value = "SIM" Then
   Set oCfg = sysConfig("TEMPOLOGOFF")
   txtTempoLogoff.Text = oCfg.Value
   cboLogoffAutomatico.Text = "SIM"
   txtTempoLogoff.Enabled = True
Else
   cboLogoffAutomatico.Text = "NÃO"
   txtTempoLogoff.Enabled = False
End If

Set oCfg = Nothing
End Sub

Private Sub MostrarConfDeclararRecebedor()
Set oCfg = sysConfig("DECLARARRECEBEDOR")

If oCfg.Value = "SIM" Then
   cboDeclararRecebedor.Text = "SIM"
Else
   cboDeclararRecebedor.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub
Private Sub MostrarConfBloqueioCliente()
   Set oCfg = sysConfig("BLOQUEIAR_CLIENTE")
   
   If CBool(oCfg.Value) = True Then
      Set oCfg = sysConfig("DIAS_BLOQUEIO")
      txtQuantDiasBloqueiar.Text = oCfg.Value
      cboClienteDebito.Text = "SIM"
   Else
      cboClienteDebito.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub
Private Sub ORC_MostrarConfImpressao()
   Set oCfg = sysConfig("CONF_IMPRESSAO_ORC")
   
   If CBool(oCfg.Value) = True Then
      cboConfirmaImpressaoORC.Text = "SIM"
   Else
      cboConfirmaImpressaoORC.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub NFCe_MostrarCombinarImpNFCe()
Set oCfg = sysConfig("COMBINARIMPNFCE")

If oCfg.Value = "SIM" Then
   cboCombinarImpNFCe.Text = "SIM"
Else
   cboCombinarImpNFCe.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub NFCe_MostrarConfPrazo()
Set oCfg = sysConfig("CONFPRAZONFCE")

If oCfg.Value = "SIM" Then
   cboConfirmaPrazoNFCe.Text = "SIM"
Else
   cboConfirmaPrazoNFCe.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub NFCe_MostrarConfCPF()
Set oCfg = sysConfig("CONFCPFNFCE")

If oCfg.Value = "SIM" Then
   cboConfirmaCPFNFCe.Text = "SIM"
Else
   cboConfirmaCPFNFCe.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub NFCe_MostrarConfImpressao()
Set oCfg = sysConfig("CONFIMPNFCE")

If oCfg.Value = "SIM" Then
   cboConfirmaImpressaoNFCe.Text = "SIM"
Else
   cboConfirmaImpressaoNFCe.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub AV_MostrarConfImpressao()
   Set oCfg = sysConfig("CONF_IMPRESSAO_AV")
   
   If CBool(oCfg.Value) = True Then
      cboConfirmaImpressaoAV.Text = "SIM"
   Else
      cboConfirmaImpressaoAV.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarIniciaisBalanca()
Set oCfg = sysConfig("INICIAISBALANCA")

txtIniciaisBalanca.Text = oCfg.Value

Set oCfg = Nothing
End Sub

Private Sub MostrarQtdeDigitosBalanca()
Set oCfg = sysConfig("QTDEDIGITOSBALANCA")

txtQtdeDigitosBalanca.Text = oCfg.Value

Set oCfg = Nothing
End Sub
Private Sub MostrarQtdeDigitosEtiquetas()
Set oCfg = sysConfig("QTDEDIGITOSETIQUETAS")

txtQtdeDigitosEtiqueta.Text = oCfg.Value

Set oCfg = Nothing
End Sub
Private Sub MostrarIniciaisEtiquetas()
Set oCfg = sysConfig("INICIAISETIQUETAS")

txtIniciaisEtiqueta.Text = oCfg.Value

Set oCfg = Nothing
End Sub

Private Sub MostrarLimitarCompra()
Set oCfg = sysConfig("LIMITARCOMPRA")

If oCfg.Value = "SIM" Then
    cboLimitarCompra.Text = "SIM"
Else
   cboLimitarCompra.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub
Private Sub MostrarCashback()
Set oCfg = sysConfig("CASHBACKAV")

If oCfg.Value = "SIM" Then
    cboCashbackVista.Text = "SIM"
Else
    cboCashbackVista.Text = "NÃO"
End If

'Set oCfg = Nothing

Set oCfg = sysConfig("CASHBACKAP")

If oCfg.Value = "SIM" Then
    cboCashbackPrazo.Text = "SIM"
Else
    cboCashbackPrazo.Text = "NÃO"
End If

'Set oCfg = Nothing

Set oCfg = sysConfig("CASHBACKVALORAV")
txtCashbackAV.Text = FormatNumber(oCfg.Value, 2)

Set oCfg = sysConfig("CASHBACKVALORAP")
txtCashbackAP.Text = FormatNumber(oCfg.Value, 2)

Set oCfg = sysConfig("CASHBACKVALIDADE")
txtCashbackValidade.Text = oCfg.Value

Set oCfg = Nothing

End Sub

Private Sub MostrarMultiplasRef()
Set oCfg = sysConfig("MULTIPLASREF")

If oCfg.Value = "SIM" Then
    cboMultiplasRef.Text = "SIM"
Else
   cboMultiplasRef.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub
Private Sub MostrarClienteDesativar()
Set oCfg = sysConfig("DESATIVARCLIENTES")

'If IsNull(oCfg.Value) Then Exit Sub

If oCfg.Value = "SIM" Then
    Set oCfg = sysConfig("QUANTDIASDESATIVAR")
    txtQuantDiasDesativar.Text = oCfg.Value
    cboDesativarClientes.Text = "SIM"
Else
   cboDesativarClientes.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub MostrarFechamentoAP()
   Set oCfg = sysConfig("CONF_FECHAMENTO_AP")
   
   If CBool(oCfg.Value) = True Then
      cboConfFerchamentoAP.Text = "SIM"
   Else
      cboConfFerchamentoAP.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarFechamentoORC()
   Set oCfg = sysConfig("CONF_FECHAMENTO_ORC")
   
   If CBool(oCfg.Value) = True Then
      cboConfFerchamentoORC.Text = "SIM"
   Else
      cboConfFerchamentoORC.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarFechamentoAV()
   Set oCfg = sysConfig("CONF_FECHAMENTO_AV")
   
   If CBool(oCfg.Value) = True Then
      cboConfFerchamentoAV.Text = "SIM"
   Else
      cboConfFerchamentoAV.Text = "NÃO"
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub ORC_MostrarImp2()
   'Set oCfg = sysConfig("IMP2_MAQ_ORC")
   'cboAVMaqCupGuiORC.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP2_COMPART_ORC")
   'cboAVImpCupGuiORC.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub AV_MostrarImp2()
   'Set oCfg = sysConfig("IMP2_MAQ_AV")
   'cboAVMaqCupGui.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP2_COMPART_AV")
   'cboAVImpCupGui.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub AP_MostrarImp2()
   'Set oCfg = sysConfig("IMP2_MAQ_AP")
   'cboAVMaqCupGuiAP.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP2_COMPART_AP")
   'cboAVImpCupGuiAP.Text = oCfg.Value
End Sub

Private Sub AP_MostrarImp3()
   'Set oCfg = sysConfig("IMP3_MAQ_AP")
   'cboAVMaqCupSerAP.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP3_COMPART_AP")
   'cboAVImpCupSerAP.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub ORC_MostrarImp3()
   'Set oCfg = sysConfig("IMP3_MAQ_ORC")
   'cboAVMaqCupSerORC.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP3_COMPART_ORC")
   'cboAVImpCupSerORC.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub AV_MostrarImp3()
   'Set oCfg = sysConfig("IMP3_MAQ_AV")
   'cboAVMaqCupSer.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP3_COMPART_AV")
   'cboAVImpCupSer.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub AP_MostrarTipoImpressao()
   Set oCfg = sysConfig("IMPRIMIR_AP")
   
   txtTipoImpressaoAP.Text = oCfg.Value
   If txtTipoImpressaoAP.Text = 1 Then
      cboTipoImpressaoAP.Text = "FOLHA"
   ElseIf txtTipoImpressaoAP.Text = 2 Then
      cboTipoImpressaoAP.Text = "FOLHA"
   ElseIf txtTipoImpressaoAP.Text = 3 Then
      cboTipoImpressaoAP.Text = "CUPOM"
   Else
      Exit Sub
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub ORC_MostrarTipoImpressao()
   Set oCfg = sysConfig("IMPRIMIR_ORC")
   
   txtTipoImpressaoORC.Text = oCfg.Value
   If txtTipoImpressaoORC.Text = 1 Then
      cboTipoImpressaoORC.Text = "FOLHA"
   ElseIf txtTipoImpressaoORC.Text = 2 Then
      cboTipoImpressaoORC.Text = "FOLHA"
   ElseIf txtTipoImpressaoORC.Text = 3 Then
      cboTipoImpressaoORC.Text = "CUPOM"
   Else
      Exit Sub
   End If
End Sub

Private Sub AV_MostrarTipoImpressao()
   Set oCfg = sysConfig("IMPRIMIR_AV")
   
   txtTipoImpressaoAV.Text = oCfg.Value
    If txtTipoImpressaoAV.Text = 1 Then
        cboTipoImpressaoAV.Text = "FOLHA"
    ElseIf txtTipoImpressaoAV.Text = 2 Then
        cboTipoImpressaoAV.Text = "FOLHA"
    ElseIf txtTipoImpressaoAV.Text = 3 Then
        cboTipoImpressaoAV.Text = "CUPOM"
    Else
        Exit Sub
    End If
End Sub

Private Sub cboAcrescCredito_GotFocus()
Dim var_Texto As String
var_Texto = cboAcrescCredito.Text
   cboAcrescCredito.Clear
   cboAcrescCredito.AddItem "SIM"
   cboAcrescCredito.AddItem "NÃO"
cboAcrescCredito.Text = var_Texto
End Sub

Private Sub cboAcrescCredito_LostFocus()
If cboAcrescCredito.Text = "SIM" Then
   cboAcrescCreditoConf.Text = "1"
ElseIf cboAcrescCredito.Text = "NÃO" Then
   cboAcrescCreditoConf.Text = "0"
End If
End Sub


Private Sub cboAcrescDebito_GotFocus()
Dim var_Texto As String
var_Texto = cboAcrescDebito.Text
   cboAcrescDebito.Clear
   cboAcrescDebito.AddItem "SIM"
   cboAcrescDebito.AddItem "NÃO"
cboAcrescDebito.Text = var_Texto
End Sub

Private Sub cboAcrescDebito_LostFocus()
If cboAcrescDebito.Text = "SIM" Then
   txtAcrescDebitoConf.Text = "1"
ElseIf cboAcrescDebito.Text = "NÃO" Then
   txtAcrescDebitoConf.Text = "0"
End If
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


Private Sub cboBalancaPorta_GotFocus()
Dim var_Texto As String
var_Texto = cboBalancaPorta.Text
   cboBalancaPorta.Clear
   cboBalancaPorta.AddItem "COM1"
   cboBalancaPorta.AddItem "COM2"
   cboBalancaPorta.AddItem "COM3"
   cboBalancaPorta.AddItem "COM4"
   cboBalancaPorta.AddItem "COM5"
   cboBalancaPorta.AddItem "COM6"
   cboBalancaPorta.AddItem "COM7"
   cboBalancaPorta.AddItem "COM8"
   cboBalancaPorta.AddItem "COM9"
cboBalancaPorta.Text = var_Texto
End Sub


Private Sub cboCashbackPrazo_GotFocus()
Dim var_Texto As String
var_Texto = cboCashbackPrazo.Text
   cboCashbackPrazo.Clear
   cboCashbackPrazo.AddItem "SIM"
   cboCashbackPrazo.AddItem "NÃO"
cboCashbackPrazo.Text = var_Texto
End Sub


Private Sub cboCashbackVista_GotFocus()
Dim var_Texto As String
var_Texto = cboCashbackVista.Text
   cboCashbackVista.Clear
   cboCashbackVista.AddItem "SIM"
   cboCashbackVista.AddItem "NÃO"
cboCashbackVista.Text = var_Texto
End Sub


Private Sub cboClienteDebito_Change()
'If cboClienteDebito.Text = "SIM" Then
'    lblDiasBloqueio.Enabled = True
'    txtQuantDiasBloqueiar.Enabled = True
'Else
'    lblDiasBloqueio.Enabled = False
'    txtQuantDiasBloqueiar.Enabled = False
'End If
End Sub

Private Sub cboClienteDebito_GotFocus()
Dim var_Texto As String
var_Texto = cboClienteDebito.Text
   cboClienteDebito.Clear
   cboClienteDebito.AddItem "SIM"
   cboClienteDebito.AddItem "NÃO"
cboClienteDebito.Text = var_Texto
End Sub


Private Sub cboCombinarImpNFCe_GotFocus()
Dim var_Texto As String
var_Texto = cboCombinarImpNFCe.Text
   cboCombinarImpNFCe.Clear
   cboCombinarImpNFCe.AddItem "SIM"
   cboCombinarImpNFCe.AddItem "NÃO"
cboCombinarImpNFCe.Text = var_Texto
End Sub


Private Sub cboConfFerchamentoAP_GotFocus()
Dim var_Texto As String
var_Texto = cboConfFerchamentoAP.Text
   cboConfFerchamentoAP.Clear
   cboConfFerchamentoAP.AddItem "SIM"
   cboConfFerchamentoAP.AddItem "NÃO"
cboConfFerchamentoAP.Text = var_Texto
End Sub


Private Sub cboConfFerchamentoAV_GotFocus()
Dim var_Texto As String
var_Texto = cboConfFerchamentoAV.Text
   cboConfFerchamentoAV.Clear
   cboConfFerchamentoAV.AddItem "SIM"
   cboConfFerchamentoAV.AddItem "NÃO"
cboConfFerchamentoAV.Text = var_Texto
End Sub


Private Sub cboConfFerchamentoORC_GotFocus()
Dim var_Texto As String
var_Texto = cboConfFerchamentoORC.Text
   cboConfFerchamentoORC.Clear
   cboConfFerchamentoORC.AddItem "SIM"
   cboConfFerchamentoORC.AddItem "NÃO"
cboConfFerchamentoORC.Text = var_Texto
End Sub


Private Sub cboConfirmaCPFNFCe_GotFocus()
Dim var_Texto As String
var_Texto = cboConfirmaCPFNFCe.Text
   cboConfirmaCPFNFCe.Clear
   cboConfirmaCPFNFCe.AddItem "SIM"
   cboConfirmaCPFNFCe.AddItem "NÃO"
cboConfirmaCPFNFCe.Text = var_Texto
End Sub


Private Sub cboConfirmaImpressaoAP_GotFocus()
Dim var_Texto As String
var_Texto = cboConfirmaImpressaoAP.Text
   cboConfirmaImpressaoAP.Clear
   cboConfirmaImpressaoAP.AddItem "SIM"
   cboConfirmaImpressaoAP.AddItem "NÃO"
cboConfirmaImpressaoAP.Text = var_Texto
End Sub


Private Sub cboConfirmaImpressaoAV_GotFocus()
Dim var_Texto As String
var_Texto = cboConfirmaImpressaoAV.Text
   cboConfirmaImpressaoAV.Clear
   cboConfirmaImpressaoAV.AddItem "SIM"
   cboConfirmaImpressaoAV.AddItem "NÃO"
cboConfirmaImpressaoAV.Text = var_Texto
End Sub


Private Sub cboConfirmaImpressaoNFCe_GotFocus()
Dim var_Texto As String
var_Texto = cboConfirmaImpressaoNFCe.Text
   cboConfirmaImpressaoNFCe.Clear
   cboConfirmaImpressaoNFCe.AddItem "SIM"
   cboConfirmaImpressaoNFCe.AddItem "NÃO"
cboConfirmaImpressaoNFCe.Text = var_Texto
End Sub


Private Sub cboConfirmaImpressaoORC_GotFocus()
Dim var_Texto As String
var_Texto = cboConfirmaImpressaoORC.Text
   cboConfirmaImpressaoORC.Clear
   cboConfirmaImpressaoORC.AddItem "SIM"
   cboConfirmaImpressaoORC.AddItem "NÃO"
cboConfirmaImpressaoORC.Text = var_Texto
End Sub


Private Sub cboConfirmaPrazoNFCe_GotFocus()
Dim var_Texto As String
var_Texto = cboConfirmaPrazoNFCe.Text
   cboConfirmaPrazoNFCe.Clear
   cboConfirmaPrazoNFCe.AddItem "SIM"
   cboConfirmaPrazoNFCe.AddItem "NÃO"
cboConfirmaPrazoNFCe.Text = var_Texto
End Sub


Private Sub cboDeclararRecebedor_GotFocus()
Dim var_Texto As String
var_Texto = cboDeclararRecebedor.Text
   cboDeclararRecebedor.Clear
   cboDeclararRecebedor.AddItem "SIM"
   cboDeclararRecebedor.AddItem "NÃO"
cboDeclararRecebedor.Text = var_Texto
End Sub


Private Sub cboDesativarClientes_Click()
cboDesativarClientes_LostFocus
End Sub

Private Sub cboDesativarClientes_GotFocus()
Dim var_Texto As String
var_Texto = cboDesativarClientes.Text
   cboDesativarClientes.Clear
   cboDesativarClientes.AddItem "SIM"
   cboDesativarClientes.AddItem "NÃO"
cboDesativarClientes.Text = var_Texto
End Sub


Private Sub cboDesativarClientes_LostFocus()
If cboDesativarClientes.Text = "SIM" Then
    txtQuantDiasDesativar.Enabled = True
Else
    txtQuantDiasDesativar.Enabled = False
End If
End Sub


Private Sub cboEstoqueNegativo_GotFocus()
Dim var_Texto As String
var_Texto = cboEstoqueNegativo.Text
   cboEstoqueNegativo.Clear
   cboEstoqueNegativo.AddItem "SIM"
   cboEstoqueNegativo.AddItem "NÃO"
cboEstoqueNegativo.Text = var_Texto
End Sub


Private Sub cboFormaParcelasImpressao_GotFocus()
Dim var_Texto As String
var_Texto = cboFormaParcelasImpressao.Text
   cboFormaParcelasImpressao.Clear
   cboFormaParcelasImpressao.AddItem "1 - Resumido"
   cboFormaParcelasImpressao.AddItem "2 - Detalhado"
cboFormaParcelasImpressao.Text = var_Texto
End Sub


Private Sub cboFormaParcelasImpressao_LostFocus()
If cboFormaParcelasImpressao.Text = "1 - Resumido" Then
    txtFormaParcelasImpressao.Text = 1
ElseIf cboFormaParcelasImpressao.Text = "2 - Detalhado" Then
    txtFormaParcelasImpressao.Text = 2
Else
    txtFormaParcelasImpressao.Text = 0
End If
End Sub


Private Sub cboHabilitarAluguel_GotFocus()
Dim var_Texto As String
var_Texto = cboHabilitarAluguel.Text
   cboHabilitarAluguel.Clear
   cboHabilitarAluguel.AddItem "SIM"
   cboHabilitarAluguel.AddItem "NÃO"
cboHabilitarAluguel.Text = var_Texto
End Sub



Private Sub cboHabilitarOS_GotFocus()
Dim var_Texto As String
var_Texto = cboHabilitarOS.Text
   cboHabilitarOS.Clear
   cboHabilitarOS.AddItem "SIM"
   cboHabilitarOS.AddItem "NÃO"
cboHabilitarOS.Text = var_Texto
End Sub


Private Sub cboIdentificarMaquina_GotFocus()
Dim var_Texto As String
var_Texto = cboIdentificarMaquina.Text
   cboIdentificarMaquina.Clear
   cboIdentificarMaquina.AddItem "SIM"
   cboIdentificarMaquina.AddItem "NÃO"
cboIdentificarMaquina.Text = var_Texto
End Sub


Private Sub cboIdentVendedor_GotFocus()
Dim var_Texto As String
var_Texto = cboIdentVendedor.Text
   cboIdentVendedor.Clear
   cboIdentVendedor.AddItem "LOGIN"
   cboIdentVendedor.AddItem "CÓD. FUNC."
cboIdentVendedor.Text = var_Texto
End Sub


Private Sub cboIdentVendedor_LostFocus()
If cboIdentVendedor.Text = "LOGIN" Then
   txtTipoIndPDV.Text = "1"
   cboSegAvancado.Text = "NÃO"
ElseIf cboIdentVendedor.Text = "CÓD. FUNC." Then
   txtTipoIndPDV.Text = "2"
   cboSegAvancado.Text = "SIM"
End If
End Sub



Private Sub cboImprimirAP_GotFocus()
Dim var_Texto As String
var_Texto = cboImprimirAP.Text
   cboImprimirAP.Clear
   cboImprimirAP.AddItem "SIM"
   cboImprimirAP.AddItem "NÃO"
cboImprimirAP.Text = var_Texto
End Sub


Private Sub cboImprimirAV_GotFocus()
Dim var_Texto As String
var_Texto = cboImprimirAV.Text
   cboImprimirAV.Clear
   cboImprimirAV.AddItem "SIM"
   cboImprimirAV.AddItem "NÃO"
cboImprimirAV.Text = var_Texto
End Sub


Private Sub cboImprimirNFCe_GotFocus()
Dim var_Texto As String
var_Texto = cboImprimirNFCe.Text
   cboImprimirNFCe.Clear
   cboImprimirNFCe.AddItem "SIM"
   cboImprimirNFCe.AddItem "NÃO"
cboImprimirNFCe.Text = var_Texto
End Sub


Private Sub cboImprimirORC_GotFocus()
Dim var_Texto As String
var_Texto = cboImprimirORC.Text
   cboImprimirORC.Clear
   cboImprimirORC.AddItem "SIM"
   cboImprimirORC.AddItem "NÃO"
cboImprimirORC.Text = var_Texto
End Sub


Private Sub cboIncluirPrecos_GotFocus()
Dim var_Texto As String
var_Texto = cboIncluirPrecos.Text
   cboIncluirPrecos.Clear
   cboIncluirPrecos.AddItem "SIM"
   cboIncluirPrecos.AddItem "NÃO"
cboIncluirPrecos.Text = var_Texto
End Sub



Private Sub cboLimitarCompra_GotFocus()
Dim var_Texto As String
var_Texto = cboLimitarCompra.Text
   cboLimitarCompra.Clear
   cboLimitarCompra.AddItem "SIM"
   cboLimitarCompra.AddItem "NÃO"
cboLimitarCompra.Text = var_Texto
End Sub


Private Sub cboLogoffAutomatico_Click()
If cboLogoffAutomatico.Text = "SIM" Then
    txtTempoLogoff.Enabled = True
Else
    txtTempoLogoff.Enabled = False
End If
End Sub

Private Sub cboLogoffAutomatico_GotFocus()
Dim var_Texto As String
var_Texto = cboLogoffAutomatico.Text
   cboLogoffAutomatico.Clear
   cboLogoffAutomatico.AddItem "SIM"
   cboLogoffAutomatico.AddItem "NÃO"
cboLogoffAutomatico.Text = var_Texto
End Sub


Private Sub cboLogoffAutomatico_Validate(Cancel As Boolean)
If cboLogoffAutomatico.Text = "SIM" Then
    txtTempoLogoff.Enabled = True
Else
    txtTempoLogoff.Enabled = False
End If
End Sub

Private Sub cboMes_GotFocus()
Dim vMes As Integer

cboMes.Clear

For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMes
End Sub


Private Sub cboMultiplasRef_GotFocus()
Dim var_Texto As String
var_Texto = cboMultiplasRef.Text
   cboMultiplasRef.Clear
   cboMultiplasRef.AddItem "SIM"
   cboMultiplasRef.AddItem "NÃO"
cboMultiplasRef.Text = var_Texto
End Sub


Private Sub cboNotaEntregaAP_GotFocus()
Dim var_Texto As String
var_Texto = cboNotaEntregaAP.Text
   cboNotaEntregaAP.Clear
   cboNotaEntregaAP.AddItem "SIM"
   cboNotaEntregaAP.AddItem "NÃO"
cboNotaEntregaAP.Text = var_Texto
End Sub


Private Sub cboNotaEntregaAV_GotFocus()
Dim var_Texto As String
var_Texto = cboNotaEntregaAV.Text
   cboNotaEntregaAV.Clear
   cboNotaEntregaAV.AddItem "SIM"
   cboNotaEntregaAV.AddItem "NÃO"
cboNotaEntregaAV.Text = var_Texto
End Sub


Private Sub cboSegAvancado_GotFocus()
Dim var_Texto As String
var_Texto = cboSegAvancado.Text
   cboSegAvancado.Clear
   cboSegAvancado.AddItem "SIM"
   cboSegAvancado.AddItem "NÃO"
cboSegAvancado.Text = var_Texto
End Sub

Private Sub cboTipoCaixa_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoCaixa.Text
   cboTipoCaixa.Clear
   cboTipoCaixa.AddItem "1 - Único"
   cboTipoCaixa.AddItem "2 - Multiplus"
cboTipoCaixa.Text = var_Texto
End Sub


Private Sub cboTipoCaixa_LostFocus()
If cboTipoCaixa.Text = "1 - Único" Then
    txtTipoCaixa.Text = 1
ElseIf cboTipoCaixa.Text = "2 - Multiplus" Then
    txtTipoCaixa.Text = 2
Else
    txtTipoCaixa.Text = 0
End If
End Sub


Private Sub cboTipoDesconto_Change()
If cboTipoDesconto.Text = "MANUAL" Then
    frmValorDesc.Visible = False
    lblDescAV.Visible = False
    lblDescAP.Visible = False
    txtValorDescAV.Visible = False
    txtValorDescAP.Visible = False
    frmLimites.Visible = False
    lblMargem.Visible = False
    lblLimiteAV.Visible = False
    lblLimiteAP.Visible = False
    txtDescMargemAV1.Visible = False
    txtDescMargemAV2.Visible = False
    txtDescMargemAV3.Visible = False
    txtDescAV1.Visible = False
    txtDescAV2.Visible = False
    txtDescAV3.Visible = False
    txtDescAP1.Visible = False
    txtDescAP2.Visible = False
    txtDescAP3.Visible = False
ElseIf cboTipoDesconto.Text = "FIXO" Then
    frmValorDesc.Visible = True
    frmValorDesc.Caption = "Descontos"
    lblDescAV.Visible = True
    lblDescAP.Visible = True
    txtValorDescAV.Visible = True
    txtValorDescAP.Visible = True
    frmLimites.Visible = False
    lblMargem.Visible = False
    lblLimiteAV.Visible = False
    lblLimiteAP.Visible = False
    txtDescMargemAV1.Visible = False
    txtDescMargemAV2.Visible = False
    txtDescMargemAV3.Visible = False
    txtDescAV1.Visible = False
    txtDescAV2.Visible = False
    txtDescAV3.Visible = False
    txtDescAP1.Visible = False
    txtDescAP2.Visible = False
    txtDescAP3.Visible = False
ElseIf cboTipoDesconto.Text = "GRADUAL" Then
    frmValorDesc.Visible = False
    lblDescAV.Visible = False
    lblDescAP.Visible = False
    txtValorDescAV.Visible = False
    txtValorDescAP.Visible = False
    frmLimites.Visible = True
    lblMargem.Visible = True
    lblLimiteAV.Visible = True
    lblLimiteAP.Visible = True
    txtDescMargemAV1.Visible = True
    txtDescMargemAV2.Visible = True
    txtDescMargemAV3.Visible = True
    txtDescAV1.Visible = True
    txtDescAV2.Visible = True
    txtDescAV3.Visible = True
    txtDescAP1.Visible = True
    txtDescAP2.Visible = True
    txtDescAP3.Visible = True
End If

End Sub

Private Sub cboTipoDesconto_Click()
cboTipoDesconto_Change
End Sub


Private Sub cboTipoDesconto_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoDesconto.Text
   cboTipoDesconto.Clear
   cboTipoDesconto.AddItem "MANUAL"
   cboTipoDesconto.AddItem "FIXO"
   cboTipoDesconto.AddItem "GRADUAL"
cboTipoDesconto.Text = var_Texto
End Sub


Private Sub cboTipoDesconto_LostFocus()
If cboTipoDesconto.Text = "MANUAL" Then
   txtTipoDesc.Text = "1"
ElseIf cboTipoDesconto.Text = "FIXO" Then
   txtTipoDesc.Text = "2"
ElseIf cboTipoDesconto.Text = "GRADUAL" Then
    txtTipoDesc.Text = "3"
End If
End Sub


Private Sub cboTipoEmpresa_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoEmpresa.Text
   cboTipoEmpresa.Clear
   cboTipoEmpresa.AddItem "Varejo"
   cboTipoEmpresa.AddItem "Farmacia"
   cboTipoEmpresa.AddItem "Restaurante/Lanchonete"
   cboTipoEmpresa.AddItem "Sapataria/Vestuário"
   cboTipoEmpresa.AddItem "Autopeça/Motopeça"
   cboTipoEmpresa.AddItem "Academia"
   cboTipoEmpresa.AddItem "Escola"
   cboTipoEmpresa.AddItem "Consultorio"
cboTipoEmpresa.Text = var_Texto
End Sub


Private Sub cboTipoEmpresa_Validate(Cancel As Boolean)
If cboTipoEmpresa.Text = "Varejo" Then
   txtTipoCadastroProduto.Text = "1"
ElseIf cboTipoEmpresa.Text = "Farmacia" Then
   txtTipoCadastroProduto.Text = "2"
ElseIf cboTipoEmpresa.Text = "Restaurante/Lanchonete" Then
   txtTipoCadastroProduto.Text = "3"
ElseIf cboTipoEmpresa.Text = "Sapataria/Vestuário" Then
   txtTipoCadastroProduto.Text = "4"
ElseIf cboTipoEmpresa.Text = "Autopeça/Motopeça" Then
   txtTipoCadastroProduto.Text = "5"
ElseIf cboTipoEmpresa.Text = "Academia" Then
   txtTipoCadastroProduto.Text = "6"
ElseIf cboTipoEmpresa.Text = "Escola" Then
   txtTipoCadastroProduto.Text = "7"
ElseIf cboTipoEmpresa.Text = "Consultorio" Then
   txtTipoCadastroProduto.Text = "8"
Else
   txtTipoCadastroProduto.Text = "1"
End If
End Sub


Private Sub cboTipoImpressaoAP_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoImpressaoAP.Text
   cboTipoImpressaoAP.Clear
   cboTipoImpressaoAP.AddItem "FOLHA"
   cboTipoImpressaoAP.AddItem "CUPOM"
cboTipoImpressaoAP.Text = var_Texto
End Sub


Private Sub cboTipoImpressaoAP_LostFocus()
If cboTipoImpressaoAP.Text = "FOLHA" Then
   txtTipoImpressaoAP.Text = "1"
ElseIf cboTipoImpressaoAP.Text = "CUPOM" Then
   txtTipoImpressaoAP.Text = "3"
End If
End Sub


Private Sub cboTipoImpressaoAV_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoImpressaoAV.Text
   cboTipoImpressaoAV.Clear
   cboTipoImpressaoAV.AddItem "FOLHA"
   cboTipoImpressaoAV.AddItem "CUPOM"
cboTipoImpressaoAV.Text = var_Texto
End Sub


Private Sub cboTipoImpressaoAV_LostFocus()
If cboTipoImpressaoAV.Text = "FOLHA" Then
   txtTipoImpressaoAV.Text = "1"
ElseIf cboTipoImpressaoAV.Text = "CUPOM" Then
   txtTipoImpressaoAV.Text = "3"
End If
End Sub


Private Sub cboTipoImpressaoORC_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoImpressaoORC.Text
   cboTipoImpressaoORC.Clear
   cboTipoImpressaoORC.AddItem "FOLHA"
   cboTipoImpressaoORC.AddItem "CUPOM"
cboTipoImpressaoORC.Text = var_Texto
End Sub


Private Sub cboTipoImpressaoORC_LostFocus()
If cboTipoImpressaoORC.Text = "FOLHA" Then
   txtTipoImpressaoORC.Text = "1"
ElseIf cboTipoImpressaoORC.Text = "CUPOM" Then
   txtTipoImpressaoORC.Text = "3"
End If
End Sub


Private Sub cboTipoJuros_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoJuros.Text
   cboTipoJuros.Clear
   cboTipoJuros.AddItem "Parcela"
   cboTipoJuros.AddItem "Restante"
cboTipoJuros.Text = var_Texto
End Sub


Private Sub cboTipoJuros_LostFocus()
If cboTipoJuros.Text = "Parcela" Then
    txtTipoJuros.Text = 1
ElseIf cboTipoJuros.Text = "Restante" Then
    txtTipoJuros.Text = 0
Else
    txtTipoCaixa.Text = 0
End If
End Sub


Private Sub cboTipoLogin_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoLogin.Text
   cboTipoLogin.Clear
   cboTipoLogin.AddItem "NOME"
   cboTipoLogin.AddItem "CPF"
cboTipoLogin.Text = var_Texto
End Sub

Private Sub cboTipoOS_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoOS.Text
   cboTipoOS.Clear
   cboTipoOS.AddItem "Automóveis"
   cboTipoOS.AddItem "Motocicletas"
   cboTipoOS.AddItem "Motores"
   cboTipoOS.AddItem "Gráfica Rápida"
   cboTipoOS.AddItem "Comunicação Visual"
   cboTipoOS.AddItem "Informática"
   cboTipoOS.AddItem "Celular"
   cboTipoOS.AddItem "Recapadora"
   cboTipoOS.AddItem "Agrícola"
cboTipoOS.Text = var_Texto
End Sub


Private Sub Mostrar_Backup()
On Error Resume Next

Set oCfg = sysConfig("BACKUP_PASTA")
txtPastaBackup.Text = oCfg.Value

Set oCfg = sysConfig("BACKUP_HORARIO")
txtBackupHorario.Text = oCfg.Value

Set oCfg = sysConfig("BACKUP_AUTOMATICO")
chkBackupAutomatico.Value = Abs(oCfg.Value)

Set oCfg = Nothing
End Sub

Private Sub cboTipoReciboHaver_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoReciboHaver.Text
   cboTipoReciboHaver.Clear
   cboTipoReciboHaver.AddItem "CUPOM"
   cboTipoReciboHaver.AddItem "FOLHA"
cboTipoReciboHaver.Text = var_Texto
End Sub


Private Sub cboTipoReciboPgto_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoReciboPgto.Text
   cboTipoReciboPgto.Clear
   cboTipoReciboPgto.AddItem "CUPOM"
   cboTipoReciboPgto.AddItem "FOLHA"
cboTipoReciboPgto.Text = var_Texto
End Sub


Private Sub cboValorVenda_GotFocus()
Dim var_Texto As String
var_Texto = cboValorVenda.Text
   cboValorVenda.Clear
   cboValorVenda.AddItem "1 - Único"
   cboValorVenda.AddItem "2 - Multiplus"
cboValorVenda.Text = var_Texto
End Sub


Private Sub cboValorVenda_LostFocus()
If cboValorVenda.Text = "1 - Único" Then
    txtValorVenda.Text = 1
ElseIf cboValorVenda.Text = "2 - Multiplus" Then
    txtValorVenda.Text = 2
Else
    txtValorVenda.Text = 0
End If
End Sub


Private Sub Balanca_Mostrar()
   On Error Resume Next
   
   Set oCfg = sysConfig("PDV_BALANCA_MODELO")
   ComBalancaModelo.Text = oCfg.Value
   
   Set oCfg = sysConfig("PDV_BALANCA_PORTA")
   cboBalancaPorta.Text = oCfg.Value
   
   Set oCfg = Nothing
End Sub

Private Sub chameleonButton1_Click()
Clipboard.Clear
Clipboard.SetText mskCPF.Text
End Sub


Private Sub chkLimiteDesc_Click()
If chkLimiteDesc.Value = Checked Then
    If cboTipoDesconto.Text = "MANUAL" Then
        frmValorDesc.Visible = True
        frmValorDesc.Caption = "Limites"
        lblDescAV.Visible = True
        lblDescAP.Visible = True
        txtValorDescAV.Visible = True
        txtValorDescAP.Visible = True
    End If
Else
    If cboTipoDesconto.Text <> "FIXO" Then
        frmValorDesc.Visible = False
        lblDescAV.Visible = False
        lblDescAP.Visible = False
        txtValorDescAV.Visible = False
        txtValorDescAP.Visible = False
    End If
End If
End Sub

Private Sub cmdAdicionar_Click()
If txtFantasia.Text = "" Or txtRazao.Text = "" Or mskCPF.Text = "" Then Exit Sub

sSQL = "SELECT CNPJ FROM  empresas_desbloueio WHERE CNPJ = '" & mskCPF.Text & "';"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    MsgBox "Empresa já cadastrada!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

MostrarEmpresa
LimparEmpresa

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim x As Integer

With Grid
   .Clear
   .Cols = 6
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 2500
   .ColWidth(2) = 4000
   .ColWidth(3) = 1800
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "FANTASIA"
   .TextMatrix(0, 2) = "RAZÃO."
   .TextMatrix(0, 3) = "CNPJ"
   .TextMatrix(0, 4) = "CODIGO"
   
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("FANTASIA")
         .TextMatrix(.rows - 1, 2) = rTabela("RAZAO")
         .TextMatrix(.rows - 1, 3) = rTabela("CNPJ")
         .TextMatrix(.rows - 1, 4) = rTabela("CODIGO")
         .TextMatrix(.rows - 1, 5) = rTabela("vMarcado")
         rTabela.MoveNext
         
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
   
   For i = 1 To .rows - 1
       For j = 0 To .Cols - 1
          .Col = j
          .Row = i
    
          If .TextMatrix(i, 5) = "NÃO" Then
             .CellForeColor = vbBlack
          ElseIf .TextMatrix(i, 5) = "SIM" Then
             .CellForeColor = vbRed
          Else
             .CellForeColor = vbBlack
          End If
          
       Next
    Next
   
   .rows = .rows - 1
   .Redraw = True
End With
End Sub



Private Function Inserir_Dados() As Boolean
Dim vNovoCodigo As Integer

'autonumeração
sSQL = "SELECT MAX(CODIGO) r FROM empresas_desbloueio "
vNovoCodigo = SQLExecutaRetorno(sSQL, "r", 0) + 1

'Comando de inclusão
sSQL = "INSERT INTO empresas_desbloueio (" & _
   "fantasia, razao, cnpj, codigo, marcado) VALUES ('" & _
   txtFantasia.Text & "', '" & txtRazao.Text & "', '" & mskCPF.Text & "', " & vNovoCodigo & ", 0)"

'Retorna o resultado da atualização
Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdBuscarLogo_Click()
CommonDialog1.Filter = "Imagens JPG(*.jpg)|*.jpg"
CommonDialog1.ShowOpen
txtCaminhoCupom.Text = CommonDialog1.FileName
End Sub

Private Sub cmdBuscarPlano_Click()
CommonDialog1.Filter = "Imagens JPG(*.jpg)|*.jpg"
CommonDialog1.ShowOpen
txtCaminho.Text = CommonDialog1.FileName
End Sub

Private Sub cmdDesativarClientes_Click()
Dim sSQL As String
Dim varQuantDias As Integer

Set oCfg = sysConfig("QUANTDIASDESATIVAR")
varQuantDias = oCfg.Value

Set oCfg = sysConfig("DESATIVARCLIENTES")

If oCfg.Value = "SIM" Then
   
     sSQL = "UPDATE cliente SET status = 0 " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE PARCELAS.STATUS = 0 AND datediff(day ,parcelas.data,getdate() ) > " & varQuantDias & " ;"
    dbData.Execute sSQL
Else
   cboDesativarClientes.Text = "NÃO"
End If

Set oCfg = Nothing
End Sub

Private Sub SelecionaImpressora(lstCtl As Control, ByVal Buffer As String)
Dim intI    As Integer
Dim strS    As String

Do
    intI = InStr(Buffer, Chr(0))
    If intI > 0 Then
        strS = Left(Buffer, intI - 1)
        If Len(Trim(strS)) Then lstCtl.AddItem strS
        Buffer = Mid(Buffer, intI + 1)
    Else
        If Len(Trim(Buffer)) Then lstCtl.AddItem Buffer
        Buffer = ""
    End If
Loop While intI > 0
End Sub

Private Sub cmdDesmarcar_Click()
i = Grid.Row
dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 0 WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
MostrarEmpresa
End Sub

Private Sub cmdDesmarcarTodos_Click()
dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 0;"
MostrarEmpresa
End Sub

Private Sub cmdLocalizar_Click()
sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado FROM  empresas_desbloueio where (FANTASIA LIKE '%" & txtFantasia.Text & "%')"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cmdMarcar_Click()
i = Grid.Row
dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 1 WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
MostrarEmpresa
End Sub

Private Sub cmdMostrarSenha_Click()
If cboMes.Text = "" Then Exit Sub
If cboAno.Text = "" Then Exit Sub

Dim vCnpj As Integer
Dim vQuantRazao As Integer

vCnpj = SomarDigitos(mskCPF.Text)
vQuantRazao = Len(txtRazao.Text)

Dim vNumeroMes As Integer
If cboMes.Text = "Janeiro" Then
    vNumeroMes = 1
ElseIf cboMes.Text = "Fevereiro" Then
    vNumeroMes = 2
ElseIf cboMes.Text = "Março" Then
    vNumeroMes = 3
ElseIf cboMes.Text = "Abril" Then
    vNumeroMes = 4
ElseIf cboMes.Text = "Maio" Then
    vNumeroMes = 5
ElseIf cboMes.Text = "Junho" Then
    vNumeroMes = 6
ElseIf cboMes.Text = "Julho" Then
    vNumeroMes = 7
ElseIf cboMes.Text = "Agosto" Then
    vNumeroMes = 8
ElseIf cboMes.Text = "Setembro" Then
    vNumeroMes = 9
ElseIf cboMes.Text = "Outubro" Then
    vNumeroMes = 10
ElseIf cboMes.Text = "Novembro" Then
    vNumeroMes = 11
ElseIf cboMes.Text = "Dezembro" Then
    vNumeroMes = 12
End If

'começa a criação
Dim vDataInicio As Date
Dim vDia As Integer
Dim vMes As Integer
Dim vMesInt As String
Dim vAno As Integer
Dim vMesRef As String

vDia = 30
vMes = vNumeroMes
vAno = cboAno

Dim vDataBloqueio As String

'Autonumeracao_Pagamentos

'If chkProximo.Value = 1 Then
'    vDataInicio = vDia & " / " & vMes & " / " & vAno
'    vDataInicio = Format(DateAdd("m", Val(1), vDataInicio), "dd/mm/yy")
'    vMesInt = Format(vDataInicio, "mmmm")
'    vAno = Year(vDataInicio)
'    vMesRef = vMesInt & "/" & vAno
'Else
    vDataInicio = vDia & " / " & vMes & " / " & vAno
    vMesInt = Format(vDataInicio, "mmmm")
'    vAno = Year(vDataInicio)
'    vMesRef = vMesInt & "/" & vAno
'End If

'vDataBloqueio = Format(DateAdd("d", Val(5), vDataInicio), "dd/mm/yy")

'codigo de desbloqueio
    
'    If vMesInt = "janeiro" Then
'        vNumeroMes = 1
'    ElseIf vMesInt = "fevereiro" Then
'        vNumeroMes = 2
'    ElseIf vMesInt = "março" Then
'        vNumeroMes = 3
'    ElseIf vMesInt = "abril" Then
'        vNumeroMes = 4
'    ElseIf vMesInt = "maio" Then
'        vNumeroMes = 5
'    ElseIf vMesInt = "junho" Then
'        vNumeroMes = 6
'    ElseIf vMesInt = "julho" Then
'        vNumeroMes = 7
'    ElseIf vMesInt = "agosto" Then
'        vNumeroMes = 8
'    ElseIf vMesInt = "setembro" Then
'        vNumeroMes = 9
'    ElseIf vMesInt = "outubro" Then
'        vNumeroMes = 10
'    ElseIf vMesInt = "novembro" Then
 '       vNumeroMes = 11
 '   ElseIf vMesInt = "dezembro" Then
'        vNumeroMes = 12
'    End If
    
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
    
txtCodDesbloqueio.Text = vCodDesbloqueio
txtCodDesbloqueioTemp.Text = vCodDesbTemp
End Sub

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

Private Sub cmdNovo_Click()
LimparEmpresa
MostrarEmpresa
End Sub

Private Sub cmdPrepara_Click()
Clipboard.Clear
Clipboard.SetText "[MENSAGEM AUTOMÁTICA]: Seu código de desbloqueio é: " & txtCodDesbloqueio.Text & "   - Obs: O último caractere é uma letra. "
End Sub

Private Sub cmdPrepara2_Click()
Clipboard.Clear
Clipboard.SetText "Seu código de desbloqueio temporário é: " & txtCodDesbloqueioTemp.Text
End Sub


Private Sub cmdSalvarBalanca_Click()
AtualizarHabilitarOS
AtualizarTipoOS
AtualizarHabilitarAluguel
AtualizarCashback
MsgBox "Informações Salvas", vbInformation, "Aviso do Sistema"
End Sub

Private Sub cmdSalvarGeral_Click()
AtualizarJuros
AtualizarTipoEmpresa
AtualizarIncluirPreco
AtualizarTipoRecPgto
AtualizarTipoRecHaver
AtualizarValorVenda
AtualizarBackup
AtualizarDesativarClientes
AtualizarTipoCaixa
AtualizarBalanca
AtualizarIniciaisEtiquetas
AtualizarQuantDigitosEtiquetas
AtualizarIniciaisBalanca
AtualizarQuantDigitosBalanca
AtualizarMultiplasRef
AtualizarLimitarCompra
AtualizarTipoJuros
MsgBox "Informações Salvas", vbInformation, "Aviso do Sistema"
End Sub

Private Sub cmdSalvarImpressao_Click()
AtualizarAPImprimir
AtualizarAPConfImpressao
AtualizarAPEntregar
AtualizarAPNumCopias
AtualizarAPTipoImpressao

AtualizarAVImprimir
AtualizarAVConfImpressao
AtualizarAVEntregar
AtualizarAVNumCopias
AtualizarAVTipoImpressao

AtualizarORCImprimir
AtualizarORCConfImpressao
AtualizarORCTipoImpressao
AtualizarORCNumCopias

AtualizarNFCeImprimir
AtualizarNFCeConfImpressao
AtualizarNFCeConfCPF
AtualizarNFCeConfPrazo
AtualizarNFCeCombinarImp

AtualizarTipoImpressaoParcelas
AtualizarDeclararRecebedor

MsgBox "Informações Salvas", vbInformation, "Aviso do Sistema"
End Sub


Private Sub cmdSalvarPDV_Click()
AtualizaFundoPDV
AtualizarImgCupom
AtualizarIndentPDV
AtualizarIdentMaquina
AtualizarBloquearCliente
AtualizarVenderNegativo
AtualizarConfFechamentoAV
AtualizarConfFechamentoAP
AtualizarConfFechamentoORC
AtualizarSegurancaAvancada
AtualizarConfCartaoDebito
AtualizarConfCartaoCredito
Atualizar_ValorCartaoDebito
Atualizar_ValorCartaoCredito
AtualizarTipoLogin
AtualizarLogoffAutomatico

AtualizarTipoDesconto
AtualizarAVDesc
AtualizarAPDesc
AtualizarDescGradual

MsgBox "Informações Salvas", vbInformation, "Aviso do Sistema"
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper

'GUIA: GERAL
Mostrar_Tipo_Empresa
Mostrar_Dados_Juros
Mostrar_Backup
MostrarIncluirPreco
Mostrar_TipoValorVenda
Mostrar_TipoRecPgto
Mostrar_TipoRecHaver
MostrarClienteDesativar
Mostrar_TipoCaixa
MostrarMultiplasRef
MostrarLimitarCompra
Mostrar_TipoJuros
 
'GUIA: BALANÇA
Balanca_Mostrar
MostrarIniciaisEtiquetas
MostrarQtdeDigitosEtiquetas
MostrarIniciaisBalanca
MostrarQtdeDigitosBalanca

'GUIA: PDV
Mostrar_Fundo
Mostrar_LogoCupom
MostrarConfBloqueioCliente
MostrarEstoqueNegativo
Mostrar_Tipo_Identificacao
MostrarIdentMaquina
MostrarFechamentoAV
MostrarFechamentoAP
MostrarFechamentoORC
Mostrar_SegurancaAvancada
Mostrar_Conf_CartaoCredito
Mostrar_Conf_CartaoDebito
Mostrar_ValorCartaoCredito
Mostrar_ValorCartaoDebito
MostrarTipoLogin
MostrarLogoffAutomatico
Mostrar_TipoDesconto
Mostrar_DescGradual
MostrarConfDeclararRecebedor
DescDebito_Mostrar
DescCredito_Mostrar

'GUIA: ADICIONAIS
Mostrar_OS
Mostrar_Tipo_OS
Mostrar_Aluguel
MostrarCashback
   
'GUIA: IMPRESSÃO
AV_MostrarImp
AV_MostrarConfImpressao
AV_MostrarEntrega
AV_MostrarTipoImpressao
AV_Mostrar_Desc
AV_Mostrar_Copia

AP_MostrarImp
AP_MostrarConfImpressao
AP_MostrarEntrega
AP_MostrarTipoImpressao
AP_Mostrar_Desc
AP_Mostrar_Copia

ORC_MostrarImp
ORC_MostrarConfImpressao
ORC_MostrarTipoImpressao
ORC_Mostrar_Copia

NFCe_MostrarImp
NFCe_MostrarConfImpressao
NFCe_MostrarConfCPF
NFCe_MostrarConfPrazo
NFCe_MostrarCombinarImpNFCe

Mostrar_TipoImpressaoParcelas


MostrarEmpresa

'**************************************** CARREGA LISTA DE IMPRESSORAS **********************************************
Dim lngImpr     As Long
Dim Buffer      As String

' API GetProfileString DECLARADA NO MÓDULO modGeral
Buffer = Space(8192)
lngImpr = GetProfileString("PrinterPorts", vbNullString, "", Buffer, Len(Buffer))

SelecionaImpressora cboINIImpNormal, Buffer
cboINIImpNormal.AddItem "LPT1"
'********************************************************************************************************************
   
'MOSTRAR FORM
   SSTab1.Tab = 0
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   Caminho = appPathApp
End Sub

Private Sub Mostrar_ValorCartaoCredito()
   Set oCfg = sysConfig("ACRESC_CREDITO_VALOR")
   txtValorAcrescCredito.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = Nothing
End Sub
Private Sub Mostrar_ValorCartaoDebito()
   Set oCfg = sysConfig("ACRESC_DEBITO_VALOR")
   txtValorAcrescDebito.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = Nothing
End Sub
Private Sub Mostrar_DescGradual()
Set oCfg = sysConfig("DESC_MARGEM_AV1")
txtDescMargemAV1.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = sysConfig("DESC_AV1")
txtDescAV1.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = sysConfig("DESC_AP1")
txtDescAP1.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = sysConfig("DESC_MARGEM_AV2")
txtDescMargemAV2.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = sysConfig("DESC_AV2")
txtDescAV2.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = sysConfig("DESC_AP2")
txtDescAP2.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = sysConfig("DESC_MARGEM_AV3")
txtDescMargemAV3.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = sysConfig("DESC_AV3")
txtDescAV3.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = sysConfig("DESC_AP3")
txtDescAP3.Text = Format(oCfg.Value, "##,##0.00")

Set oCfg = Nothing
End Sub
Private Sub AV_Mostrar_Desc()
   Set oCfg = sysConfig("DESC_AV")
   txtValorDescAV.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = Nothing
End Sub

Private Sub AP_Mostrar_Desc()
   Set oCfg = sysConfig("DESC_AP")
   txtValorDescAP.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_Dados_Juros()
Set oCfg = sysConfig("JUROS_MES")
txtJurosMes.Text = Format(oCfg.Value, "##,##0.00")
Set oCfg = sysConfig("JUROS_DIA")
txtJuroDia.Text = Format(oCfg.Value, "##,##0.00")
Set oCfg = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub


Private Sub Grid_Click()
LimparEmpresa
txtFantasia.Text = (Grid.TextMatrix(Grid.Row, 1))
txtRazao.Text = (Grid.TextMatrix(Grid.Row, 2))
mskCPF.Text = (Grid.TextMatrix(Grid.Row, 3))
End Sub

Private Sub lblPROCURARPastaBackup_Click()
Dim bi As BrowseInfo 'declara as variaveis
Dim rtn&
Dim pidl&
Dim path As String
Dim pos As Integer
Dim t As Long, SpecIn As String, saida As String

    bi.hOwner = Me.hwnd 'centraliza o dialogo na tela
    bi.lpszTitle = "Procura destino do Backup..." 'define o titulo do texto
    bi.ulFlags = BIF_RETURNONLYFSDIRS 'o tipo de pasta para retornar
    pidl& = SHBrowseForFolder(bi) 'exibe o dialogo
    
    path = Space(512) 'define o tamanho maximo
    t = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'obtem o caminho selecionado
    
    pos% = InStr(path$, Chr$(0)) 'extrai o caminho da string
    SpecIn = Left(path$, pos - 1) 'define o caminho extraido
    
    If Right$(SpecIn, 1) = "\" Then 'esteja certo de que a barra "\" esta no fim do caminho
       saida = SpecIn 'se nao estiver , nao faça nada
    Else 'senao
       saida = SpecIn + "\" 'inclui a barra "\" no fim do caminho
    End If
    
    txtPastaBackup.Text = saida
End Sub


Private Sub mskCPF_KeyPress(KeyAscii As Integer)
mskCPF.Mask = "##.###.###/####-##"
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
'If SSTab1.Tab = 4 Then
'    If Tela_Principal.StatusBar1.Panels(2).Text <> "PROGRAMADOR" Then MsgBox "Somento o programador por executar essa tarefa!", vbInformation, "Aviso do Sistema": SSTab1.Tab = 0: Exit Sub
'End If
End Sub

Private Sub txtCashbackAP_LostFocus()
txtCashbackAP.Text = FormatNumber(txtCashbackAP.Text, 2)
End Sub


Private Sub txtCashbackAV_LostFocus()
txtCashbackAV.Text = FormatNumber(txtCashbackAV.Text, 2)
End Sub


Private Sub txtDescAP1_GotFocus()
SelectControl txtDescAP1
End Sub


Private Sub txtDescAP1_LostFocus()
If txtDescAP1.Text = "" Then Exit Sub
txtDescAP1.Text = Format(txtDescAP1, "##,##0.00")
End Sub


Private Sub txtDescAP2_GotFocus()
SelectControl txtDescAP2
End Sub


Private Sub txtDescAP2_LostFocus()
If txtDescAP2.Text = "" Then Exit Sub
txtDescAP2.Text = Format(txtDescAP2, "##,##0.00")
End Sub


Private Sub txtDescAP3_GotFocus()
SelectControl txtDescAP3
End Sub


Private Sub txtDescAP3_LostFocus()
If txtDescAP3.Text = "" Then Exit Sub
txtDescAP3.Text = Format(txtDescAP3, "##,##0.00")
End Sub


Private Sub txtDescAV1_GotFocus()
SelectControl txtDescAV1
End Sub


Private Sub txtDescAV1_LostFocus()
If txtDescAV1.Text = "" Then Exit Sub
txtDescAV1.Text = Format(txtDescAV1, "##,##0.00")
End Sub


Private Sub txtDescAV2_GotFocus()
SelectControl txtDescAV2
End Sub


Private Sub txtDescAV2_LostFocus()
If txtDescAV2.Text = "" Then Exit Sub
txtDescAV2.Text = Format(txtDescAV2, "##,##0.00")
End Sub


Private Sub txtDescAV3_GotFocus()
SelectControl txtDescAV3
End Sub


Private Sub txtDescAV3_LostFocus()
If txtDescAV3.Text = "" Then Exit Sub
txtDescAV3.Text = Format(txtDescAV3, "##,##0.00")
End Sub


Private Sub txtDescMargemAV1_GotFocus()
SelectControl txtDescMargemAV1
End Sub


Private Sub txtDescMargemAV1_LostFocus()
If txtDescMargemAV1.Text = "" Then Exit Sub
txtDescMargemAV1.Text = Format(txtDescMargemAV1, "##,##0.00")
End Sub


Private Sub txtDescMargemAV2_GotFocus()
SelectControl txtDescMargemAV2
End Sub


Private Sub txtDescMargemAV2_LostFocus()
If txtDescMargemAV2.Text = "" Then Exit Sub
txtDescMargemAV2.Text = Format(txtDescMargemAV2, "##,##0.00")
End Sub


Private Sub txtDescMargemAV3_GotFocus()
SelectControl txtDescMargemAV3
End Sub


Private Sub txtDescMargemAV3_LostFocus()
If txtDescMargemAV3.Text = "" Then Exit Sub
txtDescMargemAV3.Text = Format(txtDescMargemAV3, "##,##0.00")
End Sub


Private Sub txtJuroDia_GotFocus()
   SelectControl txtJuroDia
End Sub

Private Sub txtJurosMes_GotFocus()
   SelectControl txtJurosMes
End Sub

Private Sub txtJurosMes_LostFocus()
   If txtJurosMes.Text = "" Then Exit Sub
   txtJurosMes.Text = Format(txtJurosMes, "##,##0.00")
   txtJuroDia.Text = Format((txtJurosMes / 30), "##,##0.00")
End Sub

Private Sub txtNumCopia_GotFocus()
   SelectControl txtNumCopia
End Sub

Private Sub txtNumCopiaAP_GotFocus()
   SelectControl txtNumCopiaAP
End Sub

Private Sub txtNumCopiaORC_GotFocus()
   SelectControl txtNumCopiaORC
End Sub

Private Sub txtValorAcrescCredito_LostFocus()
txtValorAcrescCredito.Text = Format(txtValorAcrescCredito, "##,##0.00")
End Sub


Private Sub txtValorAcrescDebito_LostFocus()
txtValorAcrescDebito.Text = Format(txtValorAcrescDebito, "##,##0.00")
End Sub


Private Sub txtValorDescAP_GotFocus()
   SelectControl txtValorDescAP
End Sub

Private Sub txtValorDescAP_LostFocus()
   If txtValorDescAP.Text = "" Then Exit Sub
   txtValorDescAP.Text = Format(txtValorDescAP, "##,##0.00")
End Sub

Private Sub txtValorDescAV_GotFocus()
SelectControl txtValorDescAV
End Sub

Private Sub txtValorDescAV_LostFocus()
If txtValorDescAV.Text = "" Then Exit Sub
txtValorDescAV.Text = Format(txtValorDescAV, "##,##0.00")
End Sub

