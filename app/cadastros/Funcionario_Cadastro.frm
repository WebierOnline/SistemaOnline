VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Funcionario_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FUNCIONÁRIOS"
   ClientHeight    =   8490
   ClientLeft      =   -870
   ClientTop       =   435
   ClientWidth     =   11685
   Icon            =   "Funcionario_Cadastro.frx":0000
   LinkTopic       =   "Form49"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   11505
      TabIndex        =   107
      Top             =   60
      Width           =   11535
      Begin VB.TextBox txtCadastro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   60
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   420
         Picture         =   "Funcionario_Cadastro.frx":23D2
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Label47 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FUNCIONÁRIOS"
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
         TabIndex        =   108
         Top             =   240
         Width           =   2340
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   56
      Top             =   1020
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12515
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DADOS PESSOAIS"
      TabPicture(0)   =   "Funcionario_Cadastro.frx":2EAE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmDados"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "EXTRA"
      TabPicture(1)   =   "Funcionario_Cadastro.frx":2ECA
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frmEmpresa"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frmDocumento"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frmBancario"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "SISTEMA"
      TabPicture(2)   =   "Funcionario_Cadastro.frx":2EE6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "optAdmissao"
      Tab(2).Control(1)=   "optNome"
      Tab(2).Control(2)=   "optAtivo"
      Tab(2).Control(3)=   "Grid"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame1 
         Caption         =   "Comissões"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2475
         Left            =   120
         TabIndex        =   123
         Top             =   1320
         Width           =   9015
         Begin VB.Frame Frame8 
            Caption         =   "Serviços"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   7140
            TabIndex        =   163
            Top             =   240
            Width           =   2295
            Begin VB.TextBox Text12 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   169
               TabStop         =   0   'False
               Top             =   1020
               Width           =   975
            End
            Begin VB.TextBox Text11 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   168
               TabStop         =   0   'False
               Top             =   1020
               Width           =   435
            End
            Begin VB.TextBox Text10 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   167
               TabStop         =   0   'False
               Top             =   660
               Width           =   975
            End
            Begin VB.TextBox Text9 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   166
               TabStop         =   0   'False
               Top             =   660
               Width           =   435
            End
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   165
               TabStop         =   0   'False
               Top             =   300
               Width           =   975
            End
            Begin VB.TextBox txtComServicos1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   164
               TabStop         =   0   'False
               Top             =   300
               Width           =   435
            End
            Begin VB.Label Label72 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   175
               Top             =   1080
               Width           =   120
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               Caption         =   "Depois:"
               Height          =   195
               Left            =   60
               TabIndex        =   174
               Top             =   1080
               Width           =   540
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   173
               Top             =   720
               Width           =   120
            End
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
               Height          =   195
               Left            =   60
               TabIndex        =   172
               Top             =   720
               Width           =   285
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   171
               Top             =   360
               Width           =   120
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
               Height          =   195
               Left            =   60
               TabIndex        =   170
               Top             =   360
               Width           =   285
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Vendas à prazo"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   4800
            TabIndex        =   150
            Top             =   240
            Width           =   2295
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   156
               TabStop         =   0   'False
               Top             =   1020
               Width           =   975
            End
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   155
               TabStop         =   0   'False
               Top             =   1020
               Width           =   435
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   154
               TabStop         =   0   'False
               Top             =   660
               Width           =   975
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   153
               TabStop         =   0   'False
               Top             =   660
               Width           =   435
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   152
               TabStop         =   0   'False
               Top             =   300
               Width           =   975
            End
            Begin VB.TextBox txtComPrazo1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   151
               TabStop         =   0   'False
               Top             =   300
               Width           =   435
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   162
               Top             =   1080
               Width           =   120
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               Caption         =   "Depois:"
               Height          =   195
               Left            =   60
               TabIndex        =   161
               Top             =   1080
               Width           =   540
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   160
               Top             =   720
               Width           =   120
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
               Height          =   195
               Left            =   60
               TabIndex        =   159
               Top             =   720
               Width           =   285
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   158
               Top             =   360
               Width           =   120
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
               Height          =   195
               Left            =   60
               TabIndex        =   157
               Top             =   360
               Width           =   285
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Recebido"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   2460
            TabIndex        =   137
            Top             =   240
            Width           =   2295
            Begin VB.TextBox txtComRecebidos1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   300
               Width           =   435
            End
            Begin VB.TextBox txtComRECAlvo1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   142
               TabStop         =   0   'False
               Top             =   300
               Width           =   975
            End
            Begin VB.TextBox txtComRecebidos2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   141
               TabStop         =   0   'False
               Top             =   660
               Width           =   435
            End
            Begin VB.TextBox txtComRECAlvo2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   140
               TabStop         =   0   'False
               Top             =   660
               Width           =   975
            End
            Begin VB.TextBox txtComRecebidos3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   139
               TabStop         =   0   'False
               Top             =   1020
               Width           =   435
            End
            Begin VB.TextBox txtComRECAlvo3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   138
               TabStop         =   0   'False
               Top             =   1020
               Width           =   975
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
               Height          =   195
               Left            =   60
               TabIndex        =   149
               Top             =   360
               Width           =   285
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   148
               Top             =   360
               Width           =   120
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
               Height          =   195
               Left            =   60
               TabIndex        =   147
               Top             =   720
               Width           =   285
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   146
               Top             =   720
               Width           =   120
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "Depois:"
               Height          =   195
               Left            =   60
               TabIndex        =   145
               Top             =   1080
               Width           =   540
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   144
               Top             =   1080
               Width           =   120
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Vendas á vista"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   2295
            Begin VB.TextBox txtComVistaAlvo3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   134
               TabStop         =   0   'False
               Top             =   1020
               Width           =   975
            End
            Begin VB.TextBox txtComVista3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   133
               TabStop         =   0   'False
               Top             =   1020
               Width           =   435
            End
            Begin VB.TextBox txtComVistaAlvo2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   130
               TabStop         =   0   'False
               Top             =   660
               Width           =   975
            End
            Begin VB.TextBox txtComVista2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   129
               TabStop         =   0   'False
               Top             =   660
               Width           =   435
            End
            Begin VB.TextBox txtComVistaAlvo1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   660
               TabIndex        =   126
               TabStop         =   0   'False
               Top             =   300
               Width           =   975
            End
            Begin VB.TextBox txtComVista1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   125
               TabStop         =   0   'False
               Top             =   300
               Width           =   435
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   136
               Top             =   1080
               Width           =   120
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "Depois:"
               Height          =   195
               Left            =   60
               TabIndex        =   135
               Top             =   1080
               Width           =   540
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   132
               Top             =   720
               Width           =   120
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
               Height          =   195
               Left            =   60
               TabIndex        =   131
               Top             =   720
               Width           =   285
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2100
               TabIndex        =   128
               Top             =   360
               Width           =   120
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
               Height          =   195
               Left            =   60
               TabIndex        =   127
               Top             =   360
               Width           =   285
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Salário"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   915
         Left            =   5220
         TabIndex        =   118
         Top             =   420
         Width           =   3915
         Begin VB.TextBox mskSalario 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cboFormaPgto 
            Height          =   315
            Left            =   1440
            TabIndex        =   119
            Top             =   480
            Width           =   1875
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Salário:"
            Height          =   195
            Left            =   60
            TabIndex        =   122
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pagamento:"
            Height          =   195
            Left            =   1440
            TabIndex        =   121
            Top             =   240
            Width           =   1560
         End
      End
      Begin VB.OptionButton optAdmissao 
         Caption         =   "Admissão"
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
         Left            =   -74820
         TabIndex        =   116
         Top             =   540
         Width           =   1275
      End
      Begin VB.OptionButton optNome 
         Caption         =   "Nome"
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
         Left            =   -73560
         TabIndex        =   115
         Top             =   540
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optAtivo 
         Caption         =   "Ativo"
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
         Left            =   -72600
         TabIndex        =   114
         Top             =   540
         Width           =   1275
      End
      Begin VB.Frame frmBancario 
         Caption         =   "Bancário"
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
         Left            =   120
         TabIndex        =   101
         Top             =   6000
         Width           =   9015
         Begin VB.ComboBox cboTipoConta 
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   480
            Width           =   2235
         End
         Begin VB.TextBox txtConta 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   6060
            TabIndex        =   51
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtAgencia 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4380
            TabIndex        =   50
            Top             =   480
            Width           =   1635
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   2340
            TabIndex        =   49
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   60
            TabIndex        =   105
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
            Height          =   195
            Left            =   6060
            TabIndex        =   104
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
            Height          =   195
            Left            =   4380
            TabIndex        =   103
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   2340
            TabIndex        =   102
            Top             =   240
            Width           =   510
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Horários"
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
         Left            =   120
         TabIndex        =   96
         Top             =   3840
         Width           =   3555
         Begin VB.ComboBox cboDescanco 
            Height          =   315
            Left            =   1680
            TabIndex        =   36
            Top             =   300
            Width           =   1755
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   1680
            TabIndex        =   37
            Top             =   660
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskIntervalo 
            Height          =   315
            Left            =   1680
            TabIndex        =   38
            Top             =   1020
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskTermino 
            Height          =   315
            Left            =   1680
            TabIndex        =   39
            Top             =   1380
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Descanço Semanal:"
            Height          =   195
            Left            =   120
            TabIndex        =   100
            Top             =   300
            Width           =   1440
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Termino:"
            Height          =   195
            Left            =   120
            TabIndex        =   99
            Top             =   1380
            Width           =   615
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo:"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   1020
            Width           =   660
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   660
            Width           =   420
         End
      End
      Begin VB.Frame frmDocumento 
         Caption         =   "Documentação"
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
         Left            =   3720
         TabIndex        =   87
         Top             =   3840
         Width           =   5415
         Begin VB.TextBox txtPIS 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1260
            TabIndex        =   47
            Top             =   1740
            Width           =   1515
         End
         Begin VB.TextBox txtTitulozona 
            Height          =   315
            Left            =   3480
            TabIndex        =   46
            Top             =   1020
            Width           =   1815
         End
         Begin VB.TextBox txtTitulo 
            Height          =   315
            Left            =   1260
            TabIndex        =   45
            Top             =   1380
            Width           =   1515
         End
         Begin VB.TextBox txtRGorgao 
            Height          =   315
            Left            =   3480
            TabIndex        =   42
            Top             =   300
            Width           =   1815
         End
         Begin VB.TextBox txtCTserie 
            Height          =   315
            Left            =   3480
            TabIndex        =   44
            Top             =   660
            Width           =   1815
         End
         Begin VB.TextBox txtCT 
            Height          =   315
            Left            =   1260
            TabIndex        =   43
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox txtRG 
            Height          =   315
            Left            =   1260
            TabIndex        =   41
            Top             =   660
            Width           =   1515
         End
         Begin MSMask.MaskEdBox mskCPF 
            Height          =   315
            Left            =   1260
            TabIndex        =   40
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "PIS:"
            Height          =   315
            Left            =   180
            TabIndex        =   95
            Top             =   1740
            Width           =   300
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zona:"
            Height          =   315
            Left            =   2940
            TabIndex        =   94
            Top             =   1020
            Width           =   420
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Titulo Eleitoral:"
            Height          =   315
            Left            =   180
            TabIndex        =   93
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Orgão:"
            Height          =   315
            Left            =   2880
            TabIndex        =   92
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Série:"
            Height          =   315
            Left            =   2955
            TabIndex        =   91
            Top             =   660
            Width           =   405
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CLT:"
            Height          =   195
            Left            =   180
            TabIndex        =   90
            Top             =   1020
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "CPF:"
            Height          =   315
            Left            =   180
            TabIndex        =   89
            Top             =   300
            Width           =   345
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "RG:"
            Height          =   315
            Left            =   180
            TabIndex        =   88
            Top             =   660
            Width           =   285
         End
      End
      Begin VB.Frame frmEmpresa 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   120
         TabIndex        =   74
         Top             =   420
         Width           =   5055
         Begin VB.ComboBox cboSetor 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   480
            Width           =   1755
         End
         Begin VB.TextBox txtCargo 
            BackColor       =   &H00C0FFFF&
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   3000
            TabIndex        =   35
            Top             =   480
            Width           =   1935
         End
         Begin MSMask.MaskEdBox mskAdmissao 
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Setor:"
            Height          =   195
            Left            =   1200
            TabIndex        =   106
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Admissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   3000
            TabIndex        =   75
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame frmDados 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         TabIndex        =   57
         Top             =   360
         Width           =   9015
         Begin VB.CheckBox chkInativo 
            Caption         =   "Ativo"
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
            Height          =   195
            Left            =   120
            TabIndex        =   117
            ToolTipText     =   "Deixa o funcionario inativo na empresa."
            Top             =   6360
            Width           =   1515
         End
         Begin VB.Frame Frame3 
            Caption         =   "Filhos (menores de 14 anos)"
            Height          =   2055
            Left            =   0
            TabIndex        =   83
            Top             =   4260
            Width           =   8895
            Begin VB.TextBox txtNome4 
               DataSource      =   "Data1"
               Height          =   315
               Left            =   1980
               MaxLength       =   80
               TabIndex        =   32
               Top             =   1620
               Width           =   6795
            End
            Begin VB.TextBox txtIdade4 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   1620
               Width           =   615
            End
            Begin VB.TextBox txtNome3 
               DataSource      =   "Data1"
               Height          =   315
               Left            =   1980
               MaxLength       =   80
               TabIndex        =   29
               Top             =   1260
               Width           =   6795
            End
            Begin VB.TextBox txtIdade3 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   1260
               Width           =   615
            End
            Begin VB.TextBox txtNome2 
               DataSource      =   "Data1"
               Height          =   315
               Left            =   1980
               MaxLength       =   80
               TabIndex        =   26
               Top             =   900
               Width           =   6795
            End
            Begin VB.TextBox txtIdade2 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   900
               Width           =   615
            End
            Begin VB.TextBox txtNome1 
               DataSource      =   "Data1"
               Height          =   315
               Left            =   1980
               MaxLength       =   80
               TabIndex        =   23
               Top             =   540
               Width           =   6795
            End
            Begin VB.TextBox txtIdade1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   540
               Width           =   615
            End
            Begin MSMask.MaskEdBox mskNasc1 
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   540
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNasc2 
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   900
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNasc3 
               Height          =   315
               Left            =   120
               TabIndex        =   27
               Top             =   1260
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNasc4 
               Height          =   315
               Left            =   120
               TabIndex        =   30
               Top             =   1620
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Nome:"
               Height          =   195
               Left            =   2040
               TabIndex        =   86
               Top             =   300
               Width           =   465
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "Nascimento:"
               Height          =   195
               Left            =   120
               TabIndex        =   85
               Top             =   300
               Width           =   885
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Idade:"
               Height          =   195
               Left            =   1380
               TabIndex        =   84
               Top             =   300
               Width           =   450
            End
         End
         Begin VB.TextBox txtNatural 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   6840
            TabIndex        =   18
            Top             =   3240
            Width           =   2055
         End
         Begin VB.TextBox txtMae 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4560
            MaxLength       =   80
            TabIndex        =   20
            Top             =   3900
            Width           =   4335
         End
         Begin VB.TextBox txtPai 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            MaxLength       =   80
            TabIndex        =   19
            Top             =   3900
            Width           =   4395
         End
         Begin VB.ComboBox cboEscolaridade 
            Height          =   315
            Left            =   5760
            TabIndex        =   15
            Top             =   2580
            Width           =   3135
         End
         Begin VB.TextBox txtApelido 
            BackColor       =   &H00C0FFFF&
            DataSource      =   "Data1"
            Height          =   315
            Left            =   6960
            MaxLength       =   25
            TabIndex        =   2
            Top             =   600
            Width           =   1965
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            ItemData        =   "Funcionario_Cadastro.frx":2F02
            Left            =   1800
            List            =   "Funcionario_Cadastro.frx":2F04
            TabIndex        =   11
            Top             =   2580
            Width           =   675
         End
         Begin VB.ComboBox cboCidade 
            Height          =   315
            ItemData        =   "Funcionario_Cadastro.frx":2F06
            Left            =   120
            List            =   "Funcionario_Cadastro.frx":2F08
            TabIndex        =   10
            Top             =   2580
            Width           =   1635
         End
         Begin VB.TextBox txtNome 
            BackColor       =   &H00C0FFFF&
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   600
            Width           =   1965
         End
         Begin VB.TextBox txtEndereco 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   1260
            Width           =   3375
         End
         Begin VB.TextBox txtReferencia 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   5280
            TabIndex        =   5
            Top             =   1260
            Width           =   2055
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   3120
            TabIndex        =   9
            Top             =   1920
            Width           =   5775
         End
         Begin VB.TextBox txtIdade 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   5100
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2580
            Width           =   615
         End
         Begin VB.ComboBox cboSexo 
            Height          =   315
            ItemData        =   "Funcionario_Cadastro.frx":2F0A
            Left            =   2520
            List            =   "Funcionario_Cadastro.frx":2F0C
            TabIndex        =   12
            Top             =   2580
            Width           =   1395
         End
         Begin VB.TextBox txtBairro 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   3540
            TabIndex        =   4
            Top             =   1260
            Width           =   1695
         End
         Begin VB.ComboBox cboEstadoCivil 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   3240
            Width           =   1755
         End
         Begin VB.TextBox txtConjugue 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1920
            MaxLength       =   80
            TabIndex        =   17
            Top             =   3240
            Width           =   4875
         End
         Begin VB.TextBox txtSobreNome 
            BackColor       =   &H00C0FFFF&
            DataSource      =   "Data1"
            Height          =   315
            Left            =   2160
            TabIndex        =   1
            Top             =   600
            Width           =   4725
         End
         Begin MSMask.MaskEdBox mskCEP 
            Height          =   315
            Left            =   7380
            TabIndex        =   6
            Top             =   1260
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskNascimento 
            Height          =   315
            Left            =   3960
            TabIndex        =   13
            Top             =   2580
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskTelefone 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCelular 
            Height          =   315
            Left            =   1620
            TabIndex        =   8
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Naturalidade:"
            Height          =   195
            Left            =   6840
            TabIndex        =   82
            Top             =   3000
            Width           =   945
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Mãe:"
            Height          =   195
            Left            =   4620
            TabIndex        =   81
            Top             =   3660
            Width           =   360
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Pai:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   3660
            Width           =   270
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Escolaridade:"
            Height          =   195
            Left            =   5760
            TabIndex        =   79
            Top             =   2340
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login (até 25):"
            Height          =   195
            Left            =   6960
            TabIndex        =   78
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Idade:"
            Height          =   195
            Left            =   5100
            TabIndex        =   73
            Top             =   2340
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço*:"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   1020
            Width           =   795
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Ponto de Referência:"
            Height          =   195
            Left            =   5280
            TabIndex        =   70
            Top             =   1020
            Width           =   1515
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Correio Eletrônico:"
            Height          =   195
            Left            =   3120
            TabIndex        =   69
            Top             =   1680
            Width           =   1290
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Cidade*:"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   2340
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "UF*:"
            Height          =   195
            Left            =   1800
            TabIndex        =   66
            Top             =   2340
            Width           =   315
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Sexo:"
            Height          =   195
            Left            =   2520
            TabIndex        =   65
            Top             =   2340
            Width           =   405
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Data de Nasc."
            Height          =   195
            Left            =   3960
            TabIndex        =   64
            Top             =   2340
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   3540
            TabIndex        =   63
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Left            =   1620
            TabIndex        =   62
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Left            =   7380
            TabIndex        =   61
            Top             =   1020
            Width           =   330
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   3000
            Width           =   870
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Conjugue:"
            Height          =   195
            Left            =   1920
            TabIndex        =   59
            Top             =   3000
            Width           =   1410
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobrenome:"
            Height          =   195
            Left            =   2160
            TabIndex        =   58
            Top             =   360
            Width           =   855
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   6195
         Left            =   -74880
         TabIndex        =   113
         Top             =   780
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   10927
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   110
      Top             =   8220
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16272
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "22:41"
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
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   615
      Left            =   9420
      TabIndex        =   53
      Top             =   2700
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Funcionario_Cadastro.frx":2F0E
      PICN            =   "Funcionario_Cadastro.frx":2F2A
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
      Left            =   9420
      TabIndex        =   54
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Funcionario_Cadastro.frx":4CBC
      PICN            =   "Funcionario_Cadastro.frx":4CD8
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
      Height          =   615
      Left            =   9420
      TabIndex        =   55
      Top             =   4020
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Funcionario_Cadastro.frx":6A6A
      PICN            =   "Funcionario_Cadastro.frx":6A86
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
      Left            =   9420
      TabIndex        =   52
      Top             =   2040
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
      MICON           =   "Funcionario_Cadastro.frx":8818
      PICN            =   "Funcionario_Cadastro.frx":8834
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
      Left            =   9420
      TabIndex        =   111
      Top             =   1380
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Funcionario_Cadastro.frx":A5C6
      PICN            =   "Funcionario_Cadastro.frx":A5E2
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
      Left            =   9420
      TabIndex        =   112
      Top             =   7500
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Funcionario_Cadastro.frx":C374
      PICN            =   "Funcionario_Cadastro.frx":C390
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
Attribute VB_Name = "Funcionario_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
'variaveis da funcao de calcular a idade
Dim iAnos As Integer
Dim iMesAniv As Integer
Dim iMesAtual As Integer

Private Sub AutoNumeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_funcionario FROM funcionario;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigo = r("cod_funcionario") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Calcular_Idade()
   '------função para calcular a idade-----------------
   On Error GoTo TrataErro2
   
   If mskNascimento.Text = "" Then Exit Sub
   
   Dim iAnos As Integer
   Dim iMesAniv As Integer
   Dim iMesAtual As Integer
   
   iAnos = DateDiff("yyyy", CDate(mskNascimento.Text), Date)
   iMesAniv = Month(CDate(mskNascimento.Text))
   iMesAtual = Month(Date)
   
   If iMesAtual >= iMesAniv Then
      txtIdade.Text = iAnos
   Else
      txtIdade.Text = (iAnos - 1)
   End If
   
   Exit Sub
   
TrataErro2:
   If Err.Number = 13 Then
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskNascimento.SetFocus
   End If
End Sub

Private Function Inserir_Dados() As Boolean
   'A insersão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   Dim vNasc(0 To 4) As String
   Dim vIdade(0 To 4) As String
   Dim vData(1 To 2) As String
   
   vNasc(0) = IIf(Trim(mskNascimento.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNascimento.Text, ocDATA) & "', 103)")
   vNasc(1) = IIf(Trim(mskNasc1.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNasc1.Text, ocDATA) & "', 103)")
   vNasc(2) = IIf(Trim(mskNasc2.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNasc2.Text, ocDATA) & "', 103)")
   vNasc(3) = IIf(Trim(mskNasc3.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNasc3.Text, ocDATA) & "', 103)")
   vNasc(4) = IIf(Trim(mskNasc4.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNasc4.Text, ocDATA) & "', 103)")
   
   vIdade(0) = Val(Trim(txtIdade.Text))
   vIdade(1) = Val(Trim(txtIdade1.Text))
   vIdade(2) = Val(Trim(txtIdade2.Text))
   vIdade(3) = Val(Trim(txtIdade3.Text))
   vIdade(4) = Val(Trim(txtIdade4.Text))
   
   vData(1) = IIf(Trim(mskInicio.Text) = "", "Null", "CONVERT(DATETIME, '" & mskInicio.Text & "', 103)")
   vData(2) = IIf(Trim(mskTermino.Text) = "", "Null", "CONVERT(DATETIME, '" & mskTermino.Text & "', 103)")
   
   'Validação dos dados
   If Trim(mskSalario) = "" Then mskSalario = 0
   
   'Comando de atualização
   sSQL = "INSERT INTO funcionario (" & _
      "codigo, data_cadastro, nome, sobrenome, apelido, endereco, bairro, referencia, cep, telefone, celular, " & _
      "email, cidade, estado, sexo, nascimento, idade, escolaridade, estado_civil, conjugue, naturalidade, pai, " & _
      "mae, filho_nasc1, filho_nasc2, filho_nasc3, filho_nasc4, filho_idade1, filho_idade2, filho_idade3, filho_idade4, " & _
      "filho_nome1, filho_nome2, filho_nome3, filho_nome4, data_admissao, setor, cargo, salario, forma_pgto, descanco, " & _
      "inicio, intervalo, termino, cpf, rg, rg_orgao, ct, ct_serie, titulo, titulo_zona, pis, banco, agencia, conta, " & _
      "tipo_conta, ativo, Comissao_Avista1, Comissao_Avista2, Comissao_Avista3, Valor_Comissao1, Valor_Comissao2, Valor_Comissao3, Comissao_Recebido1, Comissao_Recebido2, Comissao_Recebido3, Valor_ComissaoRec1, Valor_ComissaoRec2, Valor_ComissaoRec3, Comissao_Prazo1, Comissao_Servico1) VALUES ("
   
   sSQL = sSQL & _
      txtCodigo.Text & ", " & IIf((txtCadastro.Text = ""), "Null", "CONVERT(DATETIME, '" & Format$(txtCadastro.Text, ocDATA) & "', 103)") & ", '" & txtNome.Text & "', '" & txtSobreNome.Text & "', '" & _
      txtApelido.Text & "', '" & txtEndereco.Text & "', '" & txtBairro.Text & "', '" & txtReferencia.Text & "', '" & mskCEP.Text & "', '" & _
      mskTelefone.Text & "', '" & mskCelular.Text & "', '" & txtEmail.Text & "', '" & cboCidade.Text & "', '" & cboEstado.Text & "', '" & _
      cboSexo.Text & "', " & vNasc(0) & ", " & vIdade(0) & ", '" & cboEscolaridade.Text & "', '" & _
      cboEstadoCivil.Text & "', '" & txtConjugue.Text & "', '" & txtNatural.Text & "', '" & txtPai.Text & "', '" & txtMae.Text & "', " & _
      vNasc(1) & ", " & vNasc(2) & ", " & vNasc(3) & ", " & vNasc(4) & ", " & vIdade(1) & ", " & vIdade(2) & ", " & vIdade(3) & ", " & _
      vIdade(4) & ", '" & txtNome1.Text & "', '" & txtNome2.Text & "', '" & txtNome3.Text & "', '" & txtNome4.Text & "', '" & _
      Format$(mskAdmissao.Text, ocDATA_EUA) & "', '" & cboSetor.Text & "', '" & txtCargo.Text & "', " & Replace(CCur(mskSalario.Text), ",", ".") & ", '" & _
      cboFormaPgto.Text & "', '" & cboDescanco.Text & "', " & vData(1) & ", '" & mskIntervalo.Text & "', " & vData(2) & ", '" & mskCPF.Text & "', '" & _
      txtRG.Text & "', '" & txtRGorgao.Text & "', '" & txtCT.Text & "', '" & txtCTserie.Text & "', '" & txtTitulo.Text & "', '" & txtTitulozona.Text & "', '" & _
      txtPIS.Text & "', '" & cboBanco.Text & "', '" & txtAgencia.Text & "', '" & txtConta.Text & "', '" & cboTipoConta.Text & "', " & Abs(chkInativo.Value) & ", " & Replace(CDbl(txtComVista1.Text), ",", ".") & ",  " & Replace(CDbl(txtComVista2.Text), ",", ".") & ",  " & Replace(CDbl(txtComVista3.Text), ",", ".") & ", " & Replace(CDbl(txtComVistaAlvo1.Text), ",", ".") & ",  " & Replace(CDbl(txtComVistaAlvo2.Text), ",", ".") & ",  " & Replace(CDbl(txtComVistaAlvo3.Text), ",", ".") & ", " & Replace(CDbl(txtComRecebidos1.Text), ",", ".") & ", " & Replace(CDbl(txtComRecebidos2.Text), ",", ".") & ", " & Replace(CDbl(txtComRecebidos3.Text), ",", ".") & ", " & Replace(CDbl(txtComRECAlvo1.Text), ",", ".") & ", " & Replace(CDbl(txtComRECAlvo2.Text), ",", ".") & ", " & Replace(CDbl(txtComRECAlvo3.Text), ",", ".") & ", " & Replace(CDbl(txtComPrazo1.Text), ",", ".") & ", " & Replace(CDbl(txtComServicos1.Text), ",", ".") & ");"

   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados_Acesso() As Boolean
Dim sSQL As String

'Validaçao dos dados
If txtApelido.Text = "" Then txtApelido.SetFocus:

'Comando de atualização
sSQL = "UPDATE usuario SET login = '" & txtApelido.Text & "' WHERE (codigo = " & txtCodigo.Text & ");"

'Retorna o resultado da atualização
Atualizar_Dados_Acesso = dbData.Execute(sSQL)
End Function
Private Function Atualizar_Dados() As Boolean
   Dim sSQL As String
   
   'Validaçao dos dados
   If Trim(mskSalario.Text) = "" Then mskSalario = 0
   
   'Comando de atualização
   sSQL = "UPDATE funcionario SET " & _
      "data_cadastro = CONVERT(DATETIME, '" & Format$(txtCadastro.Text, ocDATA) & "', 103), " & _
      "nome = '" & txtNome.Text & "', " & _
      "sobrenome = '" & txtSobreNome.Text & "', " & _
      "apelido = '" & txtApelido.Text & "', " & _
      "endereco = '" & txtEndereco.Text & "', " & _
      "bairro = '" & txtBairro.Text & "', " & _
      "referencia = '" & txtReferencia.Text & "', " & _
      "cep = '" & mskCEP.Text & "', " & _
      "telefone = '" & mskTelefone.Text & "', " & _
      "celular = '" & mskCelular.Text & "', " & _
      "email = '" & txtEmail.Text & "', " & _
      "cidade = '" & cboCidade.Text & "', " & _
      "estado = '" & cboEstado.Text & "', " & _
      "sexo = '" & cboSexo.Text & "', " & _
      "nascimento = CONVERT(DATETIME, '" & Format$(mskNascimento.Text, ocDATA) & "', 103), " & _
      "idade = " & txtIdade.Text & ", " & _
      "escolaridade = '" & cboEscolaridade.Text & "', " & _
      "estado_civil = '" & cboEstadoCivil.Text & "', " & _
      "conjugue = '" & txtConjugue.Text & "', " & _
      "naturalidade = '" & txtNatural.Text & "', "
   
   sSQL = sSQL & _
      "pai = '" & txtPai.Text & "', " & _
      "mae = '" & txtMae.Text & "', " & _
      "filho_nasc1 = " & IIf(Trim(mskNasc1.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNasc1.Text, ocDATA) & "', 103)") & ", " & _
      "filho_nasc2 = " & IIf(Trim(mskNasc2.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNasc2.Text, ocDATA) & "', 103)") & ", " & _
      "filho_nasc3 = " & IIf(Trim(mskNasc3.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNasc3.Text, ocDATA) & "', 103)") & ", " & _
      "filho_nasc4 = " & IIf(Trim(mskNasc4.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskNasc4.Text, ocDATA) & "', 103)") & ", " & _
      "filho_idade1 = " & txtIdade1.Text & ", filho_idade2 = " & txtIdade2.Text & ", filho_idade3 = " & txtIdade3.Text & ", filho_idade4 = " & txtIdade4.Text & ", filho_nome1 = '" & txtNome1.Text & "', filho_nome2 = '" & txtNome2.Text & "', filho_nome3 = '" & txtNome3.Text & "', filho_nome4 = '" & txtNome4.Text & "', " & _
      "data_admissao = " & IIf(Trim(mskAdmissao.Text) = "", "Null", "CONVERT(DATETIME, '" & Format$(mskAdmissao.Text, ocDATA) & "', 103)") & ", " & _
      "setor = '" & cboSetor.Text & "', " & _
      "cargo = '" & txtCargo.Text & "', " & _
      "salario = " & Replace(CCur(mskSalario.Text), ",", ".") & ", " & _
      "Comissao_Avista1 = " & Replace(CDbl(txtComVista1.Text), ",", ".") & ", Comissao_Avista2 = " & Replace(CDbl(txtComVista2.Text), ",", ".") & ", Comissao_Avista3 = " & Replace(CDbl(txtComVista3.Text), ",", ".") & ", Comissao_Prazo1 = " & Replace(CDbl(txtComPrazo1.Text), ",", ".") & ", Comissao_Servico1 = " & Replace(CDbl(txtComServicos1.Text), ",", ".") & ", " & _
      "Valor_Comissao1 = " & Replace(CDbl(txtComVistaAlvo1.Text), ",", ".") & ", Valor_Comissao2 = " & Replace(CDbl(txtComVistaAlvo2.Text), ",", ".") & ", Valor_Comissao3 = " & Replace(CDbl(txtComVistaAlvo3.Text), ",", ".") & ", " & _
      "Comissao_Recebido1 = " & Replace(CDbl(txtComRecebidos1.Text), ",", ".") & ", Comissao_Recebido2 = " & Replace(CDbl(txtComRecebidos2.Text), ",", ".") & ", Comissao_Recebido3 = " & Replace(CDbl(txtComRecebidos3.Text), ",", ".") & ", " & _
      "Valor_ComissaoRec1 = " & Replace(CDbl(txtComRECAlvo1.Text), ",", ".") & ", Valor_ComissaoRec2 = " & Replace(CDbl(txtComRECAlvo2.Text), ",", ".") & ", Valor_ComissaoRec3 = " & Replace(CDbl(txtComRECAlvo3.Text), ",", ".") & ", " & _
      "descanco = '" & cboDescanco.Text & "', forma_pgto = '" & cboFormaPgto.Text & "', " & _
      "inicio = " & IIf(Trim(mskInicio.Text) = "", "Null", "CONVERT(DATETIME, '" & mskInicio.Text & "', 103)") & ", " & _
      "intervalo = '" & mskIntervalo.Text & "', " & _
      "termino = " & IIf(Trim(mskTermino.Text) = "", "Null", "CONVERT(DATETIME, '" & mskTermino.Text & "', 103)") & ", "
   
   sSQL = sSQL & _
      "cpf = '" & mskCPF.Text & "', " & _
      "rg = '" & txtRG.Text & "', " & _
      "rg_orgao = '" & txtRGorgao.Text & "', " & _
      "ct = '" & txtCT.Text & "', " & _
      "ct_serie = '" & txtCTserie.Text & "', " & _
      "titulo = '" & txtTitulo.Text & "', " & _
      "titulo_zona = '" & txtTitulozona.Text & "', " & _
      "pis = '" & txtPIS.Text & "', " & _
      "banco = '" & cboBanco.Text & "', " & _
      "agencia = '" & txtAgencia.Text & "', " & _
      "conta = '" & txtConta.Text & "', " & _
      "tipo_conta = '" & cboTipoConta.Text & "', " & _
      "ativo = " & Abs(chkInativo.Value)
      
   'Condição para atualização
   sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)

   
   'Atualiza o CPF dos usuarios
   sSQL = "UPDATE usuario SET CPF = '" & mskCPF.Text & "' WHERE (codigo = " & txtCodigo.Text & ");"
   dbData.Execute sSQL
End Function


Private Sub Campos_Brancos()
SSTab1.Tab = 0
txtCadastro.Text = ""
If cmdAlterar.Enabled = False Then txtCodigo.Text = ""
txtNome.Text = ""
txtSobreNome.Text = ""
txtApelido.Text = ""
txtEndereco.Text = ""
txtBairro.Text = ""
txtReferencia.Text = ""
mskCEP.Mask = ""
mskCEP.Text = ""
mskTelefone.Mask = ""
mskTelefone.Text = ""
mskCelular.Mask = ""
mskCelular.Text = ""
txtEmail.Text = ""
cboCidade.Text = ""
cboEstado.Text = ""
cboSexo.Text = ""
mskNascimento.Mask = ""
mskNascimento.Text = ""
txtIdade.Text = ""
cboEscolaridade.Text = ""
cboEstadoCivil.Text = ""
txtConjugue.Text = ""
txtNatural.Text = ""
txtPai.Text = ""
txtMae.Text = ""
mskNasc1.Mask = ""
mskNasc1.Text = ""
mskNasc2.Mask = ""
mskNasc2.Text = ""
mskNasc3.Mask = ""
mskNasc3.Text = ""
mskNasc4.Mask = ""
mskNasc4.Text = ""
txtIdade1.Text = ""
txtIdade2.Text = ""
txtIdade3.Text = ""
txtIdade4.Text = ""
txtNome1.Text = ""
txtNome2.Text = ""
txtNome3.Text = ""
txtNome4.Text = ""
mskAdmissao.Mask = ""
mskAdmissao.Text = ""
cboSetor.Text = ""
txtCargo.Text = ""
mskSalario.Text = Format(0, ocMONEY)
cboFormaPgto.Text = ""
cboDescanco.Text = ""
mskInicio.Mask = ""
mskInicio.Text = ""
mskIntervalo.Mask = ""
mskIntervalo.Text = ""
mskTermino.Mask = ""
mskTermino.Text = ""
mskCPF.Mask = ""
mskCPF.Text = ""
txtRG.Text = ""
txtRGorgao.Text = ""
txtCT.Text = ""
txtCTserie.Text = ""
txtTitulo.Text = ""
txtTitulozona.Text = ""
txtPIS.Text = ""
cboBanco.Text = ""
txtAgencia.Text = ""
txtConta.Text = ""
cboTipoConta.Text = ""
txtComVista1.Text = Format(0, ocMONEY)
txtComVista2.Text = Format(0, ocMONEY)
txtComVista3.Text = Format(0, ocMONEY)
txtComVistaAlvo1.Text = Format(0, ocMONEY)
txtComVistaAlvo2.Text = Format(0, ocMONEY)
txtComVistaAlvo3.Text = Format(0, ocMONEY)
txtComRecebidos1.Text = Format(0, ocMONEY)
txtComRecebidos2.Text = Format(0, ocMONEY)
txtComRecebidos3.Text = Format(0, ocMONEY)
txtComRECAlvo1.Text = Format(0, ocMONEY)
txtComRECAlvo2.Text = Format(0, ocMONEY)
txtComRECAlvo3.Text = Format(0, ocMONEY)
txtComPrazo1.Text = Format(0, ocMONEY)
txtComServicos1.Text = Format(0, ocMONEY)

If cmdAlterar.Enabled = False Then chkInativo.Value = 0
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
SSTab1.Tab = 0

If Not rTabela Is Nothing Then
txtCodigo.Text = ValidateNull(rTabela("codigo"))
txtCadastro.Text = ValidateNull(rTabela("data_cadastro"))
txtNome.Text = ValidateNull(rTabela("nome"))
txtSobreNome.Text = ValidateNull(rTabela("sobrenome"))
txtApelido.Text = ValidateNull(rTabela("apelido"))
txtEndereco.Text = ValidateNull(rTabela("endereco"))
txtBairro.Text = ValidateNull(rTabela("bairro"))
txtReferencia.Text = ValidateNull(rTabela("referencia"))
mskCEP.Text = ValidateNull(rTabela("cep"))
mskTelefone.Text = ValidateNull(rTabela("telefone"))
mskCelular.Text = ValidateNull(rTabela("celular"))
txtEmail.Text = ValidateNull(rTabela("email"))
cboCidade.Text = ValidateNull(rTabela("cidade"))
cboEstado.Text = ValidateNull(rTabela("estado"))
cboSexo.Text = ValidateNull(rTabela("sexo"))
mskNascimento.Text = Format$(rTabela("nascimento"), ocDATA)
txtIdade.Text = ValidateNull(rTabela("idade"))
cboEscolaridade.Text = ValidateNull(rTabela("escolaridade"))
cboEstadoCivil.Text = ValidateNull(rTabela("estado_civil"))
txtConjugue.Text = ValidateNull(rTabela("conjugue"))
txtNatural.Text = ValidateNull(rTabela("naturalidade"))
txtPai.Text = ValidateNull(rTabela("pai"))
txtMae.Text = ValidateNull(rTabela("mae"))
mskNasc1.Text = Format(rTabela("filho_nasc1"), ocDATA)
mskNasc2.Text = Format(rTabela("filho_nasc2"), ocDATA)
mskNasc3.Text = Format(rTabela("filho_nasc3"), ocDATA)
mskNasc4.Text = Format(rTabela("filho_nasc4"), ocDATA)
txtIdade1.Text = ValidateNull(rTabela("filho_idade1"))
txtIdade2.Text = ValidateNull(rTabela("filho_idade2"))
txtIdade3.Text = ValidateNull(rTabela("filho_idade3"))
txtIdade4.Text = ValidateNull(rTabela("filho_idade4"))
txtNome1.Text = ValidateNull(rTabela("filho_nome1"))
txtNome2.Text = ValidateNull(rTabela("filho_nome2"))
txtNome3.Text = ValidateNull(rTabela("filho_nome3"))
txtNome4.Text = ValidateNull(rTabela("filho_nome4"))
mskAdmissao.Text = Format(rTabela("data_admissao"), ocDATA)
cboSetor.Text = ValidateNull(rTabela("setor"))
txtCargo.Text = ValidateNull(rTabela("cargo"))
mskSalario.Text = Format(ValidateNull(rTabela("salario")), ocMONEY)
cboFormaPgto.Text = ValidateNull(rTabela("forma_pgto"))
cboDescanco.Text = ValidateNull(rTabela("descanco"))
mskInicio.Text = ValidateNull(rTabela("inicio"))
mskIntervalo.Text = ValidateNull(rTabela("intervalo"))
mskTermino.Text = ValidateNull(rTabela("termino"))
mskCPF.Text = ValidateNull(rTabela("cpf"))
txtRG.Text = ValidateNull(rTabela("rg"))
txtRGorgao.Text = ValidateNull(rTabela("rg_orgao"))
txtCT.Text = ValidateNull(rTabela("ct"))
txtCTserie.Text = ValidateNull(rTabela("ct_serie"))
txtTitulo.Text = ValidateNull(rTabela("titulo"))
txtTitulozona.Text = ValidateNull(rTabela("titulo_zona"))
txtPIS.Text = ValidateNull(rTabela("pis"))
cboBanco.Text = ValidateNull(rTabela("banco"))
txtAgencia.Text = ValidateNull(rTabela("agencia"))
txtConta.Text = ValidateNull(rTabela("conta"))
cboTipoConta.Text = ValidateNull(rTabela("tipo_conta"))
txtComVista1.Text = Format(ValidateNull(rTabela("Comissao_Avista1")), ocMONEY)
txtComVista2.Text = Format(ValidateNull(rTabela("Comissao_Avista2")), ocMONEY)
txtComVista3.Text = Format(ValidateNull(rTabela("Comissao_Avista3")), ocMONEY)
txtComVistaAlvo1.Text = Format(ValidateNull(rTabela("Valor_Comissao1")), ocMONEY)
txtComVistaAlvo2.Text = Format(ValidateNull(rTabela("Valor_Comissao2")), ocMONEY)
txtComVistaAlvo3.Text = Format(ValidateNull(rTabela("Valor_Comissao3")), ocMONEY)

txtComRecebidos1.Text = Format(ValidateNull(rTabela("Comissao_Recebido1")), ocMONEY)
txtComRecebidos2.Text = Format(ValidateNull(rTabela("Comissao_Recebido2")), ocMONEY)
txtComRecebidos3.Text = Format(ValidateNull(rTabela("Comissao_Recebido3")), ocMONEY)
txtComRECAlvo1.Text = Format(ValidateNull(rTabela("Valor_ComissaoRec1")), ocMONEY)
txtComRECAlvo2.Text = Format(ValidateNull(rTabela("Valor_ComissaoRec2")), ocMONEY)
txtComRECAlvo3.Text = Format(ValidateNull(rTabela("Valor_ComissaoRec3")), ocMONEY)

txtComPrazo1.Text = Format(ValidateNull(rTabela("Comissao_Prazo1")), ocMONEY)

txtComServicos1.Text = Format(ValidateNull(rTabela("Comissao_Servico1")), ocMONEY)

'If IsNull(rTabela("estornar")) = False Then
'    If Abs(CInt(rTabela("estornar"))) = 1 Then
'       chkEstornar.Value = Checked
'    ElseIf Abs(CInt(rTabela("estornar"))) = 0 Then
'       chkEstornar.Value = Unchecked
'    End If
'Else
'    chkEstornar.Value = Unchecked
'End If

If Abs(CInt(rTabela("ativo"))) = 1 Then
   chkInativo.Value = Checked
ElseIf Abs(CInt(rTabela("ativo"))) = 0 Then
   chkInativo.Value = Unchecked
End If

'      chkSistema.Value = IIf(CInt(rTabela("acesso")) = 1, 1, 0)
'      chkInativo.Value = Abs(CInt(rTabela("ativo")))


End If
   
End Sub

Private Sub MostrarGrid()
'INDICE PARA ORGANIZAR OS DADOS
Dim INDICE As String
Dim sSQL As String
Dim r As ADODB.Recordset

If optAdmissao.Value = True Then
   INDICE = "data_admissao;"
ElseIf optNome.Value = True Then
   INDICE = "nome;"
ElseIf optAtivo.Value = True Then
   INDICE = "ativo;"
End If

sSQL = "SELECT *, CASE ativo WHEN 0 THEN 'INATIVO' WHEN 1 THEN 'ATIVO' END AS var_ativo FROM funcionario ORDER BY " & INDICE
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cboBanco_GotFocus()
   cboBanco.Clear
   cboBanco.AddItem "BANCO DO BRASIL"
   cboBanco.AddItem "BANCO DO NORDESTE"
   cboBanco.AddItem "BRADESCO"
   cboBanco.AddItem "CAIXA ECONOMICA FEDERAL"
   cboBanco.AddItem "ITAU"
   moCombo.AttachTo cboBanco
End Sub

Private Sub cboBanco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboCidade_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboDescanco_GotFocus()
   cboDescanco.Clear
   cboDescanco.AddItem "DOMINGO"
   cboDescanco.AddItem "SEGUNDA-FEIRA"
   cboDescanco.AddItem "TERÇA-FEIRA"
   cboDescanco.AddItem "QUARTA-FEIRA"
   cboDescanco.AddItem "QUINTA-FEIRA"
   cboDescanco.AddItem "SEXTA-FEIRA"
   cboDescanco.AddItem "SÁBADO"
   moCombo.AttachTo cboDescanco
End Sub

Private Sub cboDescanco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboEscolaridade_GotFocus()
   cboEscolaridade.Clear
   cboEscolaridade.AddItem "ENSINO FUNDAMENTAL COMPLETO"
   cboEscolaridade.AddItem "ENSINO FUNDAMENTAL INCOMPLETO"
   cboEscolaridade.AddItem "ENSINO MÉDIO COMPLETO"
   cboEscolaridade.AddItem "ENSINO MÉDIO INCOMPLETO"
   cboEscolaridade.AddItem "ENSINO SUPERIOR COMPLETO"
   cboEscolaridade.AddItem "ENSINO SUPERIOR INCOMPLETO"
   moCombo.AttachTo cboEscolaridade
End Sub

Private Sub cboEscolaridade_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If Len(cboEstado) = 2 Then cboSexo.SetFocus
End Sub

Private Sub cboEstadoCivil_GotFocus()
   cboEstadoCivil.Clear
   cboEstadoCivil.AddItem "CASADO"
   cboEstadoCivil.AddItem "SOLTEIRO"
   cboEstadoCivil.AddItem "VIÚVO"
   cboEstadoCivil.AddItem "?"
   moCombo.AttachTo cboEstadoCivil
End Sub

Private Sub cboEstadoCivil_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFormaPgto_GotFocus()
   cboFormaPgto.Clear
   cboFormaPgto.AddItem "SEMANAL"
   cboFormaPgto.AddItem "QUINZENAL"
   cboFormaPgto.AddItem "MENSAL"
   moCombo.AttachTo cboFormaPgto
End Sub

Private Sub cboFormaPgto_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSetor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Limpa a lista
   cboSetor.Clear
   
   sSQL = "SELECT DISTINCT setor FROM setor ORDER BY setor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboSetor.AddItem r("setor")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboSetor
End Sub

Private Sub cboSexo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub





Private Sub chkInativo_Click()
If txtApelido.Text = "" Then Exit Sub
dbData.Execute "UPDATE funcionario SET ativo = " & Abs(chkInativo.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"
dbData.Execute "UPDATE Usuario SET Visivel = " & Abs(chkInativo.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"

End Sub

Private Sub cmdAlterar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub
If txtCodigo.Text = "1" Then Exit Sub

If mskInicio.Text = "__:__" Then mskInicio.Text = ""
If mskTermino.Text = "__:__" Then mskTermino.Text = ""
If Not IsDate(mskInicio) Then mskInicio = ""
If Not IsDate(mskTermino) Then mskTermino = ""

'Faz a atualização de forma direta e verifica se houve algum erro
If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

If Not Atualizar_Dados_Acesso Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

cmdAlterar.Enabled = False
Campos_Brancos
Form_Load
End Sub

Private Sub cmdCancelar_Click()
Campos_Brancos
frmDados.Enabled = False
frmEmpresa.Enabled = False
frmDocumento.Enabled = False
frmBancario.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
End Sub

Private Sub cmdExcluir_Click()
Dim sSQL As String
Dim bRet As Boolean

If txtNome.Text = "" Or txtEndereco.Text = "" Or cboCidade.Text = "" Then
   ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão cadastrados.", vbInformation
   Exit Sub
End If

If txtCodigo.Text = "" Then Exit Sub

'Solicita ao usuário confirmação da exclusão
If ShowMsg("Excluir esse funcionario?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
'Faz a exclusão usando o comando DELETE do SQL
sSQL = "DELETE FROM funcionario WHERE (codigo = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

'Faz a exclusão usando o comando DELETE do SQL
sSQL = "DELETE FROM usuario WHERE (codigo = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

'Faz a exclusão usando o comando DELETE do SQL
sSQL = "DELETE FROM Usuario_Acessos WHERE (Cod_Usuario = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

Campos_Brancos
Form_Load
End Sub

Private Sub cmdNovo_Click()
frmDados.Enabled = True
frmEmpresa.Enabled = True
frmDocumento.Enabled = True
frmBancario.Enabled = True
Campos_Brancos
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = False
txtCadastro.Text = Format(Date, "dd/mm/yyyy")
chkInativo.Value = Checked
txtNome.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvarSenha_Click()

End Sub

Private Sub Grid_DblClick()
frmDados.Enabled = True
frmEmpresa.Enabled = True
frmDocumento.Enabled = True
frmBancario.Enabled = True
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub mskAdmissao_KeyPress(KeyAscii As Integer)
   mskAdmissao.Mask = "##/##/####"
End Sub

Private Sub mskAdmissao_LostFocus()
   If mskAdmissao.Text = "" Or mskAdmissao.Text = "__/__/____" Then
      mskAdmissao.Mask = ""
      mskAdmissao.Text = ""
   Else
      If IsDate(mskAdmissao.Text) Then
         If cmdAlterar.Enabled = False Then txtCargo.SetFocus
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskAdmissao.SetFocus
      End If
   End If
End Sub

Private Sub mskCelular_KeyPress(KeyAscii As Integer)
   mskCelular.Mask = "(##) ####-####"
End Sub

Private Sub mskCelular_LostFocus()
   If mskCelular.Text = "(__) ____-____" Then
      mskCelular.Mask = ""
      mskCelular.Text = ""
   End If
End Sub

Private Sub mskCep_KeyPress(KeyAscii As Integer)
   mskCEP.Mask = "##.###-###"
End Sub

Private Sub mskCep_LostFocus()
   If mskCEP.Text = "__.___-__" Then
      mskCEP.Mask = ""
      mskCEP.Text = ""
   End If
End Sub

Private Sub mskCPF_KeyPress(KeyAscii As Integer)
   mskCPF.Mask = "###.###.###-##"
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
   mskInicio.Mask = "##:##"
End Sub

Private Sub mskIntervalo_KeyPress(KeyAscii As Integer)
   mskIntervalo.Mask = "##:## ás ##:##"
End Sub

Private Sub mskNasc1_Change()
   Dim iAnos As Integer, iMesAniv As Integer, iMesAtual As Integer
   
   If IsDate(mskNasc1.Text) Then
      iAnos = DateDiff("yyyy", CDate(mskNasc1.Text), Date)
      iMesAniv = Month(CDate(mskNasc1.Text))
      iMesAtual = Month(Date)
      
      If iMesAtual >= iMesAniv Then
         txtIdade1.Text = iAnos
      Else
         txtIdade1.Text = (iAnos - 1)
      End If
   End If
End Sub

Private Sub mskNasc1_GotFocus()
   SelectControl mskNasc1
End Sub

Private Sub mskNasc1_KeyPress(KeyAscii As Integer)
   mskNasc1.Mask = "##/##/####"
End Sub

Private Sub mskNasc1_LostFocus()
   Dim iAnos As Integer, iMesAniv As Integer, iMesAtual As Integer
   
   If mskNasc1.Text = "" Or mskNasc1.Text = "__/__/____" Then
      mskNasc1.Mask = ""
      mskNasc1.Text = ""
      txtIdade1.Text = ""
   Else
      If IsDate(mskNasc1.Text) Then
         iAnos = DateDiff("yyyy", CDate(mskNasc1.Text), Date)
         iMesAniv = Month(CDate(mskNasc1.Text))
         iMesAtual = Month(Date)
         
         If iMesAtual >= iMesAniv Then
            txtIdade1.Text = iAnos
         Else
            txtIdade1.Text = (iAnos - 1)
         End If
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskNasc1.SetFocus
      End If
   End If
End Sub

Private Sub mskNasc2_Change()
   Dim iAnos As Integer, iMesAniv As Integer, iMesAtual As Integer
   
   If IsDate(mskNasc2.Text) Then
      iAnos = DateDiff("yyyy", CDate(mskNasc2.Text), Date)
      iMesAniv = Month(CDate(mskNasc2.Text))
      iMesAtual = Month(Date)
      
      If iMesAtual >= iMesAniv Then
         txtIdade2.Text = iAnos
      Else
         txtIdade2.Text = (iAnos - 1)
      End If
   End If
End Sub

Private Sub mskNasc2_GotFocus()
   SelectControl mskNasc2
End Sub

Private Sub mskNasc2_KeyPress(KeyAscii As Integer)
   mskNasc2.Mask = "##/##/####"
End Sub

Private Sub mskNasc2_LostFocus()
   Dim iAnos As Integer, iMesAniv As Integer, iMesAtual As Integer
   
   If mskNasc2.Text = "" Or mskNasc2.Text = "__/__/____" Then
      mskNasc2.Mask = ""
      mskNasc2.Text = ""
      txtIdade2.Text = ""
   Else
      If IsDate(mskNasc2.Text) Then
         iAnos = DateDiff("yyyy", CDate(mskNasc2.Text), Date)
         iMesAniv = Month(CDate(mskNasc2.Text))
         iMesAtual = Month(Date)
         
         If iMesAtual >= iMesAniv Then
            txtIdade2.Text = iAnos
         Else
            txtIdade2.Text = (iAnos - 1)
         End If
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskNasc2.SetFocus
      End If
   End If
End Sub

Private Sub mskNasc3_Change()
   Dim iAnos As Integer, iMesAniv As Integer, iMesAtual As Integer
   
   If IsDate(mskNasc3.Text) Then
      iAnos = DateDiff("yyyy", CDate(mskNasc3.Text), Date)
      iMesAniv = Month(CDate(mskNasc3.Text))
      iMesAtual = Month(Date)
      
      If iMesAtual >= iMesAniv Then
         txtIdade3.Text = iAnos
      Else
         txtIdade3.Text = (iAnos - 1)
      End If
   End If
End Sub

Private Sub mskNasc3_GotFocus()
   SelectControl mskNasc3
End Sub

Private Sub mskNasc3_KeyPress(KeyAscii As Integer)
   mskNasc3.Mask = "##/##/####"
End Sub

Private Sub mskNasc3_LostFocus()
   Dim iAnos As Integer, iMesAniv As Integer, iMesAtual As Integer
   
   If mskNasc3.Text = "" Or mskNasc3.Text = "__/__/____" Then
      mskNasc3.Mask = ""
      mskNasc3.Text = ""
      txtIdade3.Text = ""
   Else
      If IsDate(mskNasc3.Text) Then
         iAnos = DateDiff("yyyy", CDate(mskNasc3.Text), Date)
         iMesAniv = Month(CDate(mskNasc3.Text))
         iMesAtual = Month(Date)
         
         If iMesAtual >= iMesAniv Then
            txtIdade3.Text = iAnos
         Else
            txtIdade3.Text = (iAnos - 1)
         End If
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskNasc3.SetFocus
      End If
   End If
End Sub

Private Sub mskNasc4_Change()
   Dim iAnos As Integer, iMesAniv As Integer, iMesAtual As Integer
   
   If IsDate(mskNasc4.Text) Then
      iAnos = DateDiff("yyyy", CDate(mskNasc4.Text), Date)
      iMesAniv = Month(CDate(mskNasc4.Text))
      iMesAtual = Month(Date)
      
      If iMesAtual >= iMesAniv Then
         txtIdade4.Text = iAnos
      Else
         txtIdade4.Text = (iAnos - 1)
      End If
   End If
End Sub

Private Sub mskNasc4_GotFocus()
   SelectControl mskNasc4
End Sub

Private Sub mskNasc4_KeyPress(KeyAscii As Integer)
   mskNasc4.Mask = "##/##/####"
End Sub

Private Sub mskNasc4_LostFocus()
   Dim iAnos As Integer, iMesAniv As Integer, iMesAtual As Integer
   
   If mskNasc4.Text = "" Or mskNasc4.Text = "__/__/____" Then
      mskNasc4.Mask = ""
      mskNasc4.Text = ""
      txtIdade4.Text = ""
   Else
      If IsDate(mskNasc4.Text) Then
         iAnos = DateDiff("yyyy", CDate(mskNasc4.Text), Date)
         iMesAniv = Month(CDate(mskNasc4.Text))
         iMesAtual = Month(Date)
         
         If iMesAtual >= iMesAniv Then
            txtIdade4.Text = iAnos
         Else
            txtIdade4.Text = (iAnos - 1)
         End If
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskNasc4.SetFocus
      End If
   End If
End Sub

Private Sub mskNascimento_Change()
   If IsDate(mskNascimento.Text) Then Calcular_Idade
End Sub

Private Sub mskNascimento_GotFocus()
   SelectControl mskNascimento
End Sub

Private Sub mskNascimento_KeyPress(KeyAscii As Integer)
   mskNascimento.Mask = "##/##/####"
End Sub

Private Sub cboSexo_GotFocus()
   cboSexo.Clear
   cboSexo.AddItem "MASCULINO"
   cboSexo.AddItem "FEMININO"
   moCombo.AttachTo cboSexo
End Sub

Private Sub cboCidade_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Limpa a lista
   cboCidade.Clear
   
   sSQL = "SELECT cidade FROM funcionario GROUP BY cidade;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCidade.AddItem ValidateNull(r("cidade"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboCidade
End Sub

Private Sub cboEstado_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Limpa a lista
   cboEstado.Clear
   
   sSQL = "SELECT estado FROM funcionario GROUP BY estado;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboEstado.AddItem ValidateNull(r("estado"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboEstado
End Sub

Private Sub cmdSalvar_Click()
   'On Error GoTo TrataErro
   
   If txtNome.Text = "" Or txtSobreNome.Text = "" Then Exit Sub
   
   If txtCargo.Text = "" Then
      ShowMsg "O campo CARGO não pode está vazio!", vbExclamation
      SSTab1.Tab = 1
      txtCargo.SetFocus
      Exit Sub
   End If
   
   If mskInicio.Text = "__:__" Then mskInicio.Text = ""
   If mskTermino.Text = "__:__" Then mskTermino.Text = ""
   If Not IsDate(mskInicio) Then mskInicio = ""
   If Not IsDate(mskTermino) Then mskTermino = ""

   AutoNumeracao
   
   If Not Inserir_Dados Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   If Not Inserir_Dados_Acesso Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   'chkInativo.Value = 1
   Campos_Brancos
   Form_Load

End Sub
Private Function Inserir_Dados_Acesso() As Boolean
   Dim sSQL As String
   
   sSQL = "INSERT INTO Usuario (Codigo, Login) VALUES (" & txtCodigo.Text & ", '" & txtApelido.Text & "');"
   
   Inserir_Dados_Acesso = dbData.Execute(sSQL)
End Function
Private Sub Form_Load()
frmDados.Enabled = False
frmEmpresa.Enabled = False
frmDocumento.Enabled = False
frmBancario.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
SSTab1.Tab = 0

MostrarGrid
txtCadastro.Text = Format(Date, "dd/mm/yyyy")
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
Set moCombo = New cComboHelper
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer, x As Integer

With Grid
   .Clear
   .Cols = 6
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 500
   .ColWidth(2) = 1100
   .ColWidth(3) = 3000
   .ColWidth(4) = 1500
   .ColWidth(5) = 1200
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "CÓD"
   .TextMatrix(0, 2) = "ADMISSÃO"
   .TextMatrix(0, 3) = "NOME"
   .TextMatrix(0, 4) = "CARGO"
   .TextMatrix(0, 5) = "SITUAÇÃO"
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = Format(rTabela("codigo"), "00")
         .TextMatrix(.rows - 1, 2) = Format(rTabela("data_admissao"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = rTabela("nome")
         .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("cargo"))
         .TextMatrix(.rows - 1, 5) = rTabela("var_ativo")
         
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 4
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
   
   .rows = .rows - 1
   .Redraw = True
End With
End Sub

Private Sub cboTipoConta_GotFocus()
   cboTipoConta.Clear
   cboTipoConta.AddItem "POUPANÇA"
   cboTipoConta.AddItem "CONTA CORRENTE"
   moCombo.AttachTo cboTipoConta
End Sub

'Substituir esta função pela função RemoverAcento que é mais completa
Public Function TiraAcentos(ByVal sTexto As String) As String
   Dim sAcentos(2, 9) As String
   Dim sCaracter As String
   Dim bAcentos As Boolean
   Dim i As Integer, j As Integer
   
   sAcentos(1, 1) = "Á"
   sAcentos(2, 1) = "A"
   sAcentos(1, 2) = "É"
   sAcentos(2, 2) = "E"
   sAcentos(1, 3) = "Í"
   sAcentos(2, 3) = "I"
   sAcentos(1, 4) = "Ó"
   sAcentos(2, 4) = "O"
   sAcentos(1, 5) = "Ú"
   sAcentos(2, 5) = "U"
   sAcentos(1, 6) = "Ê"
   sAcentos(2, 6) = "E"
   sAcentos(1, 7) = "Ô"
   sAcentos(2, 7) = "O"
   sAcentos(1, 8) = "Ã"
   sAcentos(2, 8) = "A"
   sAcentos(1, 9) = "Õ"
   sAcentos(2, 9) = "O"
   
   TiraAcentos = sTexto 'Coloca o texto original como retorno
   
   For i = 1 To Len(sTexto)
      sCaracter = Mid$(sTexto, i, 1) 'Testa cada caracter
      If Asc(sCaracter) >= 192 And Asc(sCaracter) <= 255 Then
         bAcentos = True 'Indica a presença de acentos
         Exit For
      End If
   Next
   
   If bAcentos = True Then
      'Comparamos cada caracter com os elementos da matriz
      For i = 1 To Len(sTexto)
         For j = 1 To 9
            sCaracter = Mid$(sTexto, i, 1)
            If Asc(sCaracter) >= 192 And Asc(sCaracter) <= 255 Then
               If sCaracter = sAcentos(1, j) Then
                  Mid$(sTexto, i, 1) = sAcentos(2, j)
                  TiraAcentos = sTexto
               End If
            End If
         Next
      Next
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub mskSalario_LostFocus()
If mskSalario.Text = "" Then
   mskSalario.Text = Format(0, ocMONEY)
Else
   mskSalario.Text = Format(mskSalario, ocMONEY)
End If
End Sub

Private Sub mskTelefone_KeyPress(KeyAscii As Integer)
   mskTelefone.Mask = "(##) ####-####"
End Sub

Private Sub mskNascimento_LostFocus()
   If mskNascimento.Text = "" Or mskNascimento.Text = "__/__/____" Then
      mskNascimento.Mask = ""
      mskNascimento.Text = ""
      txtIdade.Text = ""
   Else
      If IsDate(mskNascimento.Text) Then
         Calcular_Idade
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskNascimento.SetFocus
      End If
   End If
End Sub

Private Sub mskTelefone_LostFocus()
   If mskTelefone.Text = "(__) ____-____" Then
      mskTelefone.Mask = ""
      mskTelefone.Text = ""
   End If
End Sub
Private Sub mskTermino_KeyPress(KeyAscii As Integer)
   mskTermino.Mask = "##:##"
End Sub
Private Sub optAdmissao_Click()
   MostrarGrid
End Sub
Private Sub optAtivo_Click()
   MostrarGrid
End Sub
Private Sub optNome_Click()
   MostrarGrid
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      'txtNome.SetFocus
   ElseIf SSTab1.Tab = 1 Then
      'mskAdmissao.SetFocus
   ElseIf SSTab1.Tab = 2 Then
      'SSTab3.Tab = 0
   End If
End Sub

Private Sub txtApelido_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtApelido_LostFocus()
txtApelido.Text = UCase(txtApelido.Text)
End Sub


Private Sub txtBairro_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCargo_Validate(Cancel As Boolean)
txtCargo.Text = RemoverAcento(txtCargo.Text)
End Sub


Private Sub txtCodigo_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

sSQL = "SELECT * FROM funcionario WHERE (codigo = " & txtCodigo.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   If r.State <> 0 Then r.Close
   Set r = Nothing
   Exit Sub
End If

cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
Campos_Brancos
Mostrar_Dados r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub txtComPrazo1_LostFocus()
If txtComPrazo1.Text = "" Then
   txtComPrazo1.Text = Format(0, ocMONEY)
Else
   txtComPrazo1.Text = Format(txtComPrazo1, ocMONEY)
End If
End Sub


Private Sub txtComRECAlvo1_LostFocus()
If txtComRECAlvo1.Text = "" Then
   txtComRECAlvo1.Text = Format(0, ocMONEY)
Else
   txtComRECAlvo1.Text = Format(txtComRECAlvo1, ocMONEY)
End If
End Sub


Private Sub txtComRECAlvo2_LostFocus()
If txtComRECAlvo2.Text = "" Then
   txtComRECAlvo2.Text = Format(0, ocMONEY)
Else
   txtComRECAlvo2.Text = Format(txtComRECAlvo2, ocMONEY)
End If
End Sub


Private Sub txtComRECAlvo3_LostFocus()
If txtComRECAlvo3.Text = "" Then
   txtComRECAlvo3.Text = Format(0, ocMONEY)
Else
   txtComRECAlvo3.Text = Format(txtComRECAlvo3, ocMONEY)
End If
End Sub


Private Sub txtComRecebidos1_LostFocus()
If txtComRecebidos1.Text = "" Then
   txtComRecebidos1.Text = Format(0, ocMONEY)
Else
   txtComRecebidos1.Text = Format(txtComRecebidos1, ocMONEY)
End If
End Sub


Private Sub txtComRecebidos2_LostFocus()
If txtComRecebidos2.Text = "" Then
   txtComRecebidos2.Text = Format(0, ocMONEY)
Else
   txtComRecebidos2.Text = Format(txtComRecebidos2, ocMONEY)
End If
End Sub


Private Sub txtComRecebidos3_LostFocus()
If txtComRecebidos3.Text = "" Then
   txtComRecebidos3.Text = Format(0, ocMONEY)
Else
   txtComRecebidos3.Text = Format(txtComRecebidos3, ocMONEY)
End If
End Sub


Private Sub txtComServicos1_LostFocus()
If txtComServicos1.Text = "" Then
   txtComServicos1.Text = Format(0, ocMONEY)
Else
   txtComServicos1.Text = Format(txtComServicos1, ocMONEY)
End If
End Sub


Private Sub txtComVista1_LostFocus()
If txtComVista1.Text = "" Then
   txtComVista1.Text = Format(0, ocMONEY)
Else
   txtComVista1.Text = Format(txtComVista1, ocMONEY)
End If
End Sub


Private Sub txtComVista2_LostFocus()
If txtComVista2.Text = "" Then
   txtComVista2.Text = Format(0, ocMONEY)
Else
   txtComVista2.Text = Format(txtComVista2, ocMONEY)
End If
End Sub


Private Sub txtComVista3_LostFocus()
If txtComVista3.Text = "" Then
   txtComVista3.Text = Format(0, ocMONEY)
Else
   txtComVista3.Text = Format(txtComVista3, ocMONEY)
End If
End Sub


Private Sub txtComVistaAlvo1_LostFocus()
If txtComVistaAlvo1.Text = "" Then
   txtComVistaAlvo1.Text = Format(0, ocMONEY)
Else
   txtComVistaAlvo1.Text = Format(txtComVistaAlvo1, ocMONEY)
End If
End Sub


Private Sub txtComVistaAlvo2_LostFocus()
If txtComVistaAlvo2.Text = "" Then
   txtComVistaAlvo2.Text = Format(0, ocMONEY)
Else
   txtComVistaAlvo2.Text = Format(txtComVistaAlvo2, ocMONEY)
End If
End Sub


Private Sub txtComVistaAlvo3_LostFocus()
If txtComVistaAlvo3.Text = "" Then
   txtComVistaAlvo3.Text = Format(0, ocMONEY)
Else
   txtComVistaAlvo3.Text = Format(txtComVistaAlvo3, ocMONEY)
End If
End Sub


Private Sub txtConjugue_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMae_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNatural_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNome_LostFocus()
txtNome.Text = RemoverAcento(txtNome.Text)
txtApelido.Text = UCase(txtApelido.Text)
End Sub

Private Sub txtNome_Validate(Cancel As Boolean)
txtNome.Text = RemoverAcento(txtNome.Text)
End Sub

Private Sub txtNome4_LostFocus()
   SSTab1.Tab = 1
   mskAdmissao.SetFocus
End Sub

Private Sub txtPai_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRGorgao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSenha1_Change()

End Sub

Private Sub txtSobreNome_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSobreNome_LostFocus()
txtSobreNome.Text = UCase(txtSobreNome.Text)
End Sub


