VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmNFSe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NFS-e - Notas Fiscais de Serviço (padrão nacional)"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   13785
      TabIndex        =   4
      Top             =   0
      Width           =   13815
      Begin ChamaleonBtn.chameleonButton cmdRecebidas 
         Height          =   315
         Left            =   12060
         TabIndex        =   16
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "NFS-e &recebidas"
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
         MICON           =   "frmNFSe.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConfig 
         Height          =   315
         Left            =   10380
         TabIndex        =   17
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Con&figurar NFS-e"
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
         MICON           =   "frmNFSe.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   480
         Picture         =   "frmNFSe.frx":0038
         Top             =   60
         Width           =   1020
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NFSe - Nota Fiscal de Serviço Eletrônica"
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
         Left            =   1860
         TabIndex        =   5
         Top             =   180
         Width           =   6060
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Pesquisa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   13740
      Begin VB.ComboBox cboSituacao 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1030
         Width           =   2055
      End
      Begin VB.Frame frmFiltros 
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
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   1260
         TabIndex        =   23
         Top             =   180
         Width           =   7335
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   60
            TabIndex        =   33
            Top             =   420
            Visible         =   0   'False
            Width           =   5925
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   405
            Width           =   1000
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   405
            Width           =   1500
         End
         Begin ChamaleonBtn.chameleonButton cmdPesquisar 
            Height          =   315
            Left            =   6060
            TabIndex        =   24
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Pesquisar"
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
            MICON           =   "frmNFSe.frx":1105
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskDe 
            Height          =   315
            Left            =   60
            TabIndex        =   29
            Top             =   420
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton calDe 
            Height          =   315
            Left            =   1185
            TabIndex        =   30
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   420
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
            MICON           =   "frmNFSe.frx":1121
            PICN            =   "frmNFSe.frx":113D
            PICH            =   "frmNFSe.frx":3490
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskAte 
            Height          =   315
            Left            =   1680
            TabIndex        =   31
            Top             =   420
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton calAte 
            Height          =   315
            Left            =   2805
            TabIndex        =   32
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   420
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
            MICON           =   "frmNFSe.frx":57E3
            PICN            =   "frmNFSe.frx":57FF
            PICH            =   "frmNFSe.frx":7B52
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano:"
            Height          =   195
            Left            =   1680
            TabIndex        =   28
            Top             =   180
            Width           =   330
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês:"
            Height          =   195
            Left            =   60
            TabIndex        =   27
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.Frame frmCriterios 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1155
         Left            =   60
         TabIndex        =   18
         Top             =   180
         Width           =   1155
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   795
         End
         Begin VB.OptionButton optCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   795
         End
         Begin VB.OptionButton optPeriodo 
            Caption         =   "Período"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   640
            Width           =   915
         End
         Begin VB.OptionButton optMensal 
            Caption         =   "Mensal"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   440
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.Label lblSit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
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
         Left            =   1320
         TabIndex        =   34
         Top             =   1080
         Width           =   810
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdNotas 
      Height          =   4275
      Left            =   60
      TabIndex        =   1
      Top             =   2280
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   7541
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdCopiarChave 
      Height          =   315
      Left            =   12660
      TabIndex        =   6
      Top             =   7440
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmNFSe.frx":9EA5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   315
      Left            =   12660
      TabIndex        =   7
      Top             =   6600
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      MICON           =   "frmNFSe.frx":9EC1
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
      Height          =   315
      Left            =   7620
      TabIndex        =   8
      Top             =   6600
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "E&xcluir"
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
      MICON           =   "frmNFSe.frx":9EDD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelarNota 
      Height          =   315
      Left            =   6300
      TabIndex        =   9
      Top             =   6600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Cancelar NFS-e"
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
      MICON           =   "frmNFSe.frx":9EF9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEmail 
      Height          =   315
      Left            =   5040
      TabIndex        =   10
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Enviar e-&mail"
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
      MICON           =   "frmNFSe.frx":9F15
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
      Height          =   315
      Left            =   4140
      TabIndex        =   11
      Top             =   6600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir "
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
      MICON           =   "frmNFSe.frx":9F31
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
      Left            =   3120
      TabIndex        =   12
      Top             =   6600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "frmNFSe.frx":9F4D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdTransmitir 
      Height          =   315
      Left            =   2160
      TabIndex        =   13
      Top             =   6600
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Transmitir"
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
      MICON           =   "frmNFSe.frx":9F69
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEditar 
      Height          =   315
      Left            =   1260
      TabIndex        =   14
      Top             =   6600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmNFSe.frx":9F85
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdNova 
      Height          =   315
      Left            =   60
      TabIndex        =   15
      Top             =   6600
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Nova NFS-e"
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
      MICON           =   "frmNFSe.frx":9FA1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      TabIndex        =   2
      Top             =   6960
      Width           =   12120
   End
End
Attribute VB_Name = "frmNFSe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Lista das NFS-e (tabela NFSe, script 178). Gerar/transmitir/consultar/cancelar/DANFSe pelo NFSeEmissor.exe (modNFSe).
'Colunas do grid: 0 Id (oculta), 1 Nº NFS-e, 2 Competência, 3 Cliente, 4 OS, 5 Valor, 6 Situação, 7 Ambiente,
'8 Mensagem (oculta), 9 código da situação (oculta), 10 PDF (oculta), 11 cancelamento pedido à prefeitura (oculta, "1")
Private Const C_ID = 0, C_NUM = 1, C_COMP = 2, C_CLI = 3, C_OS = 4, C_VALOR = 5, C_SIT = 6, C_AMB = 7, C_MSG = 8, C_SITCOD = 9, C_PDF = 10, C_ANALISE = 11

Private bTemAnaliseFiscal As Boolean   'script 179 aplicado
Private bClientesCarregados As Boolean
Private bCarregando As Boolean
Private moCombo As cComboHelper        'auto preenchimento do cboCliente

'Quem abre esta tela confere antes NFSe_BancoPronto (Unload no Form_Load de tela modal da erro 364)
Private Sub Form_Load()
Dim i As Integer
Dim vMeses As Variant
bCarregando = True
Set moCombo = New cComboHelper
vMeses = Split("Janeiro|Fevereiro|Março|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro", "|")
For i = 0 To 11
   cboMes.AddItem vMeses(i)
Next
cboMes.ListIndex = Month(Date) - 1
For i = Year(Date) - 5 To Year(Date) + 1
   cboAno.AddItem CStr(i)
   If i = Year(Date) Then cboAno.ListIndex = cboAno.NewIndex
Next
mskDe.Text = Format$(DateSerial(Year(Date), Month(Date), 1), "dd/mm/yy")
mskAte.Text = Format$(Date, "dd/mm/yy")
cboSituacao.AddItem "TODAS"
cboSituacao.AddItem "EM DIGITAÇÃO"
cboSituacao.AddItem "AUTORIZADA"
cboSituacao.AddItem "REJEITADA"
cboSituacao.AddItem "CANCELADA"
cboSituacao.AddItem "PENDENTE"
cboSituacao.ListIndex = 0
bTemAnaliseFiscal = NFSe_TemAnaliseFiscal()
MontarGrid
optMensal.Value = True
AjustarCriterio
bCarregando = False
Pesquisar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

'Mostra so os campos do criterio marcado. lblMes/lblAno servem aos tres criterios (mudam o texto).
Private Sub AjustarCriterio()
Dim bMensal As Boolean, bPeriodo As Boolean, bCliente As Boolean
bMensal = optMensal.Value
bPeriodo = optPeriodo.Value
bCliente = optCliente.Value

lblMes.Visible = bMensal Or bPeriodo Or bCliente
lblAno.Visible = bMensal Or bPeriodo
If bMensal Then
   lblMes.Caption = "Mês:"
   lblAno.Caption = "Ano:"
ElseIf bPeriodo Then
   lblMes.Caption = "De:"
   lblAno.Caption = "Até:"
ElseIf bCliente Then
   lblMes.Caption = "Cliente:"
End If
cboMes.Visible = bMensal
cboAno.Visible = bMensal
mskDe.Visible = bPeriodo
calDe.Visible = bPeriodo
mskAte.Visible = bPeriodo
calAte.Visible = bPeriodo
cboCliente.Visible = bCliente
cmdPesquisar.Visible = Not optTodos.Value
If bCliente Then cmdPesquisar.Left = 6060 Else cmdPesquisar.Left = 3240
If bCliente Then CarregarClientes
End Sub

Private Sub CarregarClientes()
Dim r As ADODB.Recordset
If bClientesCarregados Then Exit Sub
cboCliente.Clear
Set r = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE (ISNULL(nome, '') <> '') ORDER BY nome")
Do While Not r.EOF
   cboCliente.AddItem Trim$(r("nome"))
   cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
   r.MoveNext
Loop
r.Close
Set r = Nothing
bClientesCarregados = True
moCombo.AttachTo cboCliente
End Sub

Private Sub optMensal_Click()
AjustarCriterio
If Not bCarregando Then cboMes.SetFocus
End Sub

Private Sub optPeriodo_Click()
AjustarCriterio
If Not bCarregando Then mskDe.SetFocus
End Sub

Private Sub optCliente_Click()
AjustarCriterio
If Not bCarregando Then cboCliente.SetFocus
End Sub

'Todos: sem botao Pesquisar, ja exibe tudo
Private Sub optTodos_Click()
AjustarCriterio
If Not bCarregando Then Pesquisar
End Sub

Private Sub cboSituacao_Click()
'no criterio Todos nao ha botao Pesquisar: a situacao filtra na hora
If Not bCarregando And optTodos.Value Then Pesquisar
End Sub

Private Function EscolherData() As Variant
Dim fCal As Calendario
Set fCal = New Calendario
fCal.Show vbModal
EscolherData = fCal.DateSelected
Unload fCal
Set fCal = Nothing
End Function

Private Sub calDe_Click()
Dim varData As Variant
varData = EscolherData()
If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub
mskDe.Text = Format$(varData, "dd/mm/yy")
End Sub

Private Sub calAte_Click()
Dim varData As Variant
varData = EscolherData()
If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub
mskAte.Text = Format$(varData, "dd/mm/yy")
End Sub

'MaskEdBox vazio fica "__/__/__"
Private Function TextoData(ByVal s As String) As String
If InStr(s, "_") > 0 Then TextoData = "" Else TextoData = s
End Function

Private Sub MontarGrid()
Dim i As Long
'MSFlexGrid exige pelo menos 1 linha alem do cabecalho: a linha 1 fica em branco quando nao ha notas
With grdNotas
   .Rows = 2
   .Cols = 12
   .FixedRows = 1
   .FixedCols = 0
   For i = 0 To .Cols - 1
      .TextMatrix(1, i) = ""
   Next
   .TextMatrix(0, C_ID) = "Id"
   .TextMatrix(0, C_NUM) = "Nº NFS-e"
   .TextMatrix(0, C_COMP) = "Competência"
   .TextMatrix(0, C_CLI) = "Cliente (tomador)"
   .TextMatrix(0, C_OS) = "OS"
   .TextMatrix(0, C_VALOR) = "Valor líquido"
   .TextMatrix(0, C_SIT) = "Situação"
   .TextMatrix(0, C_AMB) = "Ambiente"
   .ColWidth(C_ID) = 0
   .ColWidth(C_NUM) = 1000
   .ColWidth(C_COMP) = 1150
   .ColWidth(C_CLI) = 6400
   .ColWidth(C_OS) = 900
   .ColWidth(C_VALOR) = 1300
   .ColWidth(C_SIT) = 1700
   .ColWidth(C_AMB) = 950
   .ColWidth(C_MSG) = 0
   .ColWidth(C_SITCOD) = 0
   .ColWidth(C_PDF) = 0
   .ColWidth(C_ANALISE) = 0
   .ColAlignment(C_VALOR) = flexAlignRightCenter
End With
End Sub

Private Function DataSQL(ByVal s As String, ByVal bFim As Boolean) As String
If Not IsDate(s) Then
   DataSQL = ""
   Exit Function
End If
DataSQL = "'" & Format$(CDate(s), "yyyymmdd") & IIf(bFim, " 23:59:59", "") & "'"
End Function

Private Sub Pesquisar()
Dim r As ADODB.Recordset
Dim sSQL As String, sDe As String, sAte As String
Dim n As Long, nSit As Integer
Dim dIni As Date

If optMensal.Value Then
   If cboMes.ListIndex < 0 Or cboAno.ListIndex < 0 Then Exit Sub
   dIni = DateSerial(Val(cboAno.Text), cboMes.ListIndex + 1, 1)
   sDe = DataSQL(Format$(dIni, "dd/mm/yyyy"), False)
   sAte = DataSQL(Format$(DateAdd("m", 1, dIni) - 1, "dd/mm/yyyy"), True)
ElseIf optPeriodo.Value Then
   If TextoData(mskDe.Text) <> "" And Not IsDate(TextoData(mskDe.Text)) Then
      MsgBox "Data inicial inválida.", vbExclamation, "NFS-e"
      mskDe.SetFocus
      Exit Sub
   End If
   If TextoData(mskAte.Text) <> "" And Not IsDate(TextoData(mskAte.Text)) Then
      MsgBox "Data final inválida.", vbExclamation, "NFS-e"
      mskAte.SetFocus
      Exit Sub
   End If
   sDe = DataSQL(TextoData(mskDe.Text), False)
   sAte = DataSQL(TextoData(mskAte.Text), True)
End If
sSQL = "SELECT TOP 500 IdNFSe, NumeroNFSe, DataCompetencia, TomadorNome, NumeroPedido, " & _
       "CAST(STUFF((SELECT DISTINCT ', ' + RIGHT('00000' + CAST(i.CodOS AS varchar(10)), 5) FROM NFSeItens i " & _
       "WHERE i.IdNFSe = NFSe.IdNFSe FOR XML PATH('')), 1, 2, '') AS varchar(200)) AS NumerosOS, " & _
       "ValorServicos - DescontoIncondicionado AS Valor, Situacao, Ambiente, Mensagem, ArquivoPdf, " & _
       IIf(bTemAnaliseFiscal, "CASE WHEN AnaliseFiscalSolicitada IS NULL THEN 0 ELSE 1 END", "0") & " AS Analise " & _
       "FROM NFSe WHERE (1 = 1)"
'Mensal e Periodo pela competencia (finalizacao da OS = mes de apuracao do ISS/Simples), nao pela data de emissao.
If sDe <> "" Or sAte <> "" Then
   If sDe <> "" Then sSQL = sSQL & " AND (DataCompetencia >= " & sDe & ")"
   If sAte <> "" Then sSQL = sSQL & " AND (DataCompetencia <= " & sAte & ")"
End If
If cboSituacao.ListIndex > 0 Then sSQL = sSQL & " AND (Situacao = " & (cboSituacao.ListIndex - 1) & ")"
If optCliente.Value And Trim$(cboCliente.Text) <> "" Then
   'cliente escolhido da lista: pelo codigo; texto que nao esta na lista: parte do nome do tomador
   If cboCliente.ListIndex >= 0 Then
      sSQL = sSQL & " AND (CodCliente = " & cboCliente.ItemData(cboCliente.ListIndex) & ")"
   Else
      sSQL = sSQL & " AND (TomadorNome LIKE " & NFSe_Q("%" & Trim$(cboCliente.Text) & "%") & ")"
   End If
End If
sSQL = sSQL & " ORDER BY DataCompetencia DESC, IdNFSe DESC"

MontarGrid
Set r = dbData.OpenRecordset(sSQL)
With grdNotas
   .Redraw = False
   Do While Not r.EOF
      n = n + 1
      If n > 1 Then .Rows = n + 1
      nSit = Val(ValidateNull(r("Situacao")))
      .TextMatrix(n, C_ID) = r("IdNFSe")
      .TextMatrix(n, C_NUM) = ValidateNull(r("NumeroNFSe"))
      .TextMatrix(n, C_COMP) = Format$(r("DataCompetencia"), "dd/mm/yyyy")
      .TextMatrix(n, C_CLI) = ValidateNull(r("TomadorNome"))
      'OS pelos itens da nota (NumeroPedido pode ter a ordem de compra do cliente); nota antiga sem itens: pelo NumeroPedido
      .TextMatrix(n, C_OS) = ValidateNull(r("NumerosOS"))
      If .TextMatrix(n, C_OS) = "" Then .TextMatrix(n, C_OS) = NumerosOS(ValidateNull(r("NumeroPedido")))
      .TextMatrix(n, C_VALOR) = Format$(ValidateNull(r("Valor")), ocMONEY)
      .TextMatrix(n, C_SIT) = NFSe_SituacaoTexto(nSit)
      .TextMatrix(n, C_AMB) = IIf(Val(ValidateNull(r("Ambiente"))) = 1, "Produção", "Teste")
      .TextMatrix(n, C_MSG) = ValidateNull(r("Mensagem"))
      .TextMatrix(n, C_SITCOD) = nSit
      .TextMatrix(n, C_PDF) = ValidateNull(r("ArquivoPdf"))
      .TextMatrix(n, C_ANALISE) = IIf(nSit = NFSE_AUTORIZADA And Val(ValidateNull(r("Analise"))) = 1, "1", "")
      .Row = n
      .Col = C_SIT
      Select Case nSit
         Case NFSE_AUTORIZADA
            If .TextMatrix(n, C_ANALISE) = "1" Then
               'cancelamento pedido a prefeitura: a nota continua valida ate o deferimento
               .TextMatrix(n, C_SIT) = "CANC. SOLICITADO"
               .CellForeColor = &H80FF&
            Else
               .CellForeColor = &H8000&
            End If
         Case NFSE_REJEITADA: .CellForeColor = vbRed
         Case NFSE_CANCELADA: .CellForeColor = &H808080
         Case NFSE_PENDENTE: .CellForeColor = &H80FF&
         Case Else: .CellForeColor = vbBlack
      End Select
      r.MoveNext
   Loop
   .Redraw = True
   If .Rows > 1 Then .Row = 1: .Col = 0: .ColSel = .Cols - 1
End With
r.Close
Set r = Nothing
grdNotas_RowColChange
End Sub

'NumeroPedido da nota ("OS 10" ou "OS 9, 12") -> so os numeros com 5 digitos ("00010", "00009, 00012")
Private Function NumerosOS(ByVal s As String) As String
Dim v As Variant, i As Integer, t As String
s = Trim$(s)
If UCase$(Left$(s, 3)) = "OS " Then s = Mid$(s, 4)
v = Split(s, ",")
For i = 0 To UBound(v)
   t = Trim$(v(i))
   If t Like "#*" And IsNumeric(t) Then t = Format$(Val(t), "00000")
   NumerosOS = NumerosOS & IIf(i > 0, ", ", "") & t
Next
End Function

Private Function IdSelecionado() As Long
If grdNotas.Rows < 2 Or grdNotas.Row < 1 Then Exit Function
IdSelecionado = Val(grdNotas.TextMatrix(grdNotas.Row, C_ID))
End Function

Private Function SituacaoSelecionada() As Integer
SituacaoSelecionada = -1
If IdSelecionado() = 0 Then Exit Function
SituacaoSelecionada = Val(grdNotas.TextMatrix(grdNotas.Row, C_SITCOD))
End Function

Private Sub grdNotas_RowColChange()
Dim nSit As Integer
nSit = SituacaoSelecionada()
lblMensagem.Caption = ""
If nSit >= 0 Then lblMensagem.Caption = grdNotas.TextMatrix(grdNotas.Row, C_MSG)
cmdEditar.Enabled = (nSit = NFSE_DIGITACAO Or nSit = NFSE_REJEITADA)
cmdTransmitir.Enabled = (nSit = NFSE_DIGITACAO Or nSit = NFSE_REJEITADA Or nSit = NFSE_PENDENTE)
cmdConsultar.Enabled = (nSit = NFSE_PENDENTE Or nSit = NFSE_AUTORIZADA Or nSit = NFSE_REJEITADA)
cmdImprimir.Enabled = (nSit = NFSE_AUTORIZADA Or nSit = NFSE_CANCELADA)
cmdEmail.Enabled = cmdImprimir.Enabled
cmdCancelarNota.Enabled = (nSit = NFSE_AUTORIZADA)
If cmdCancelarNota.Enabled Then cmdCancelarNota.Enabled = (grdNotas.TextMatrix(grdNotas.Row, C_ANALISE) <> "1")
cmdExcluir.Enabled = (nSit = NFSE_DIGITACAO Or nSit = NFSE_REJEITADA)
End Sub

Private Sub grdNotas_DblClick()
Select Case SituacaoSelecionada()
   Case NFSE_DIGITACAO, NFSE_REJEITADA: cmdEditar_Click
   Case NFSE_AUTORIZADA, NFSE_CANCELADA: cmdImprimir_Click
End Select
End Sub

Private Sub cmdPesquisar_Click()
Pesquisar
End Sub

Private Sub cmdNova_Click()
Dim lId As Long
frmNFSe_OS.IdGerado = 0
frmNFSe_OS.Show vbModal
lId = frmNFSe_OS.IdGerado
Unload frmNFSe_OS
If lId = 0 Then Exit Sub
IrParaCompetencia lId
Pesquisar
SelecionarId lId
If frmNFSe_Editar.Editar(lId) Then TransmitirId lId
IrParaCompetencia lId
Pesquisar
SelecionarId lId
End Sub

'Nota nova (OS finalizada em outro mes) fora do Mensal/Periodo consultado: muda o Mensal para o mes da competencia dela
Private Sub IrParaCompetencia(ByVal lId As Long)
Dim sComp As String
Dim dComp As Date, dDataDe As Date, dDataAte As Date
Dim i As Integer, nAno As Integer
If Not (optMensal.Value Or optPeriodo.Value) Then Exit Sub
sComp = SQLExecutaRetorno("SELECT CONVERT(varchar(10), DataCompetencia, 103) AS Comp FROM NFSe WHERE (IdNFSe = " & lId & ")", "Comp", "")
If Not IsDate(sComp) Then Exit Sub
dComp = CDate(sComp)
If optMensal.Value Then
   If cboMes.ListIndex = Month(dComp) - 1 And Val(cboAno.Text) = Year(dComp) Then Exit Sub
Else
   If IsDate(TextoData(mskDe.Text)) Then dDataDe = CDate(TextoData(mskDe.Text)) Else dDataDe = DateSerial(1900, 1, 1)
   If IsDate(TextoData(mskAte.Text)) Then dDataAte = CDate(TextoData(mskAte.Text)) Else dDataAte = DateSerial(9999, 12, 31)
   If dComp >= dDataDe And dComp <= dDataAte Then Exit Sub
End If
'ano fora da lista: inclui na posicao certa (lista crescente)
nAno = Year(dComp)
For i = 0 To cboAno.ListCount - 1
   If Val(cboAno.List(i)) >= nAno Then Exit For
Next
If i >= cboAno.ListCount Then
   cboAno.AddItem CStr(nAno)
ElseIf Val(cboAno.List(i)) <> nAno Then
   cboAno.AddItem CStr(nAno), i
End If
bCarregando = True
optMensal.Value = True
AjustarCriterio
cboMes.ListIndex = Month(dComp) - 1
For i = 0 To cboAno.ListCount - 1
   If Val(cboAno.List(i)) = nAno Then cboAno.ListIndex = i: Exit For
Next
bCarregando = False
End Sub

Private Sub SelecionarId(ByVal lId As Long)
Dim i As Long
For i = 1 To grdNotas.Rows - 1
   If Val(grdNotas.TextMatrix(i, C_ID)) = lId Then
      grdNotas.Row = i
      grdNotas.Col = 0
      grdNotas.ColSel = grdNotas.Cols - 1
      If Not grdNotas.RowIsVisible(i) Then grdNotas.TopRow = i
      grdNotas_RowColChange
      Exit Sub
   End If
Next
End Sub

Private Sub cmdEditar_Click()
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
If frmNFSe_Editar.Editar(lId) Then TransmitirId lId
Pesquisar
SelecionarId lId
End Sub

Private Sub cmdTransmitir_Click()
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
If ShowMsg("Transmitir a NFS-e " & grdNotas.TextMatrix(grdNotas.Row, C_OS) & " de " & grdNotas.TextMatrix(grdNotas.Row, C_CLI) & _
           " (R$ " & grdNotas.TextMatrix(grdNotas.Row, C_VALOR) & ")?", vbQuestion + vbYesNo + vbDefaultButton1) <> vbYes Then Exit Sub
TransmitirId lId
Pesquisar
SelecionarId lId
End Sub

'Transmite e mostra o resultado (autorizada: oferece abrir a DANFSe)
Private Sub TransmitirId(ByVal lId As Long)
Dim oRes As Object
Dim sStatus As String
sStatus = NFSe_Executar("acao=transmitir" & vbCrLf & "id=" & lId, oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível transmitir:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
   Exit Sub
End If
Select Case Val(NFSe_Valor(oRes, "situacao"))
   Case NFSE_AUTORIZADA
      If MsgBox("NFS-e AUTORIZADA!" & vbCrLf & vbCrLf & "Número: " & NFSe_Valor(oRes, "numero") & vbCrLf & _
                "Chave: " & NFSe_Valor(oRes, "chave") & vbCrLf & vbCrLf & "Abrir a DANFSe (PDF) agora?", vbInformation + vbYesNo, "NFS-e") = vbYes Then
         NFSe_AbrirArquivo NFSe_Valor(oRes, "pdf")
      End If
   Case NFSE_REJEITADA
      MsgBox "A NFS-e foi REJEITADA pelo Sefin:" & vbCrLf & vbCrLf & NFSe_Valor(oRes, "mensagem") & vbCrLf & vbCrLf & _
             "Corrija em Editar (ou na Configuração / cadastro do cliente) e transmita de novo.", vbExclamation, "NFS-e"
   Case NFSE_PENDENTE
      MsgBox "Sem resposta do Sefin: a NFS-e ficou PENDENTE." & vbCrLf & vbCrLf & NFSe_Valor(oRes, "mensagem") & vbCrLf & vbCrLf & _
             "Use Consultar daqui a alguns minutos (não emita outra nota para as mesmas OS).", vbExclamation, "NFS-e"
   Case Else
      MsgBox NFSe_Valor(oRes, "mensagem"), vbInformation, "NFS-e"
End Select
End Sub

Private Sub cmdConsultar_Click()
Dim oRes As Object
Dim sStatus As String
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
sStatus = NFSe_Executar("acao=consultar" & vbCrLf & "id=" & lId, oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível consultar:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
Else
   MsgBox "Situação: " & NFSe_SituacaoTexto(Val(NFSe_Valor(oRes, "situacao"))) & _
          IIf(NFSe_Valor(oRes, "numero") <> "", vbCrLf & "Número: " & NFSe_Valor(oRes, "numero"), "") & _
          IIf(NFSe_Valor(oRes, "mensagem") <> "", vbCrLf & vbCrLf & NFSe_Valor(oRes, "mensagem"), ""), vbInformation, "NFS-e"
End If
Pesquisar
SelecionarId lId
End Sub

Private Sub cmdImprimir_Click()
Dim oRes As Object
Dim sStatus As String, sPdf As String
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
sPdf = grdNotas.TextMatrix(grdNotas.Row, C_PDF)
If sPdf <> "" Then
   If Dir$(sPdf) = "" Then sPdf = ""
End If
If sPdf = "" Then
   'PDF apagado ou de outro computador: gera de novo a partir do XML gravado no banco
   sStatus = NFSe_Executar("acao=pdf" & vbCrLf & "id=" & lId, oRes)
   If Left$(sStatus, 2) <> "OK" Then
      MsgBox "Não foi possível gerar a DANFSe:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
      Exit Sub
   End If
   sPdf = NFSe_Valor(oRes, "pdf")
End If
NFSe_AbrirArquivo sPdf
End Sub

Private Function ArquivoExiste(ByVal sArquivo As String) As Boolean
On Error Resume Next
If sArquivo <> "" Then ArquivoExiste = (Dir$(sArquivo) <> "")
End Function

'PDF (DANFSe) e XML da NFS-e num e-mail so, pela DLL da NF-e (mesma configuracao de e-mail da NF-e / OS de oleo vencido)
Private Sub cmdEmail_Click()
Dim r As ADODB.Recordset
Dim oRes As Object
Dim sistNFe As snfe.Util
Dim pathAnexo() As String, emailCC() As String
Dim sStatus As String, sXml As String, sPdf As String, sEmail As String, sNumero As String, sChave As String
Dim sTomador As String, sValor As String, sRazao As String, sFantasia As String, sCorpo As String, sAssunto As String
Dim bCancelada As Boolean
Dim lId As Long, lCli As Long
Dim vRet As Variant

lId = IdSelecionado()
If lId = 0 Then Exit Sub
Set r = dbData.OpenRecordset("SELECT NumeroNFSe, ChaveAcesso, Situacao, TomadorNome, TomadorEmail, CodCliente, ArquivoXml, ArquivoPdf, " & _
                             "ValorServicos - DescontoIncondicionado AS Valor FROM NFSe WHERE (IdNFSe = " & lId & ")")
If r.EOF Then
   r.Close
   Exit Sub
End If
sNumero = ValidateNull(r("NumeroNFSe"))
sChave = ValidateNull(r("ChaveAcesso"))
bCancelada = (Val(ValidateNull(r("Situacao"))) = NFSE_CANCELADA)
sTomador = ValidateNull(r("TomadorNome"))
sEmail = Trim$(ValidateNull(r("TomadorEmail")))
lCli = Val(ValidateNull(r("CodCliente")))
sXml = ValidateNull(r("ArquivoXml"))
sPdf = ValidateNull(r("ArquivoPdf"))
sValor = Format$(ValidateNull(r("Valor")), ocMONEY)
r.Close
Set r = Nothing

'e-mail da nota (copiado do cliente ao gerar) ou o do cadastro atual do cliente
If sEmail = "" And lCli > 0 Then sEmail = Trim$(SQLExecutaRetorno("SELECT Correio_eletronico FROM cliente WHERE (codigo = " & lCli & ")", "Correio_eletronico", ""))

'arquivos apagados ou de outro computador: o NFSeEmissor gera de novo o XML e a DANFSe a partir do XML gravado no banco
If Not ArquivoExiste(sXml) Or Not ArquivoExiste(sPdf) Then
   sStatus = NFSe_Executar("acao=pdf" & vbCrLf & "id=" & lId, oRes)
   If Left$(sStatus, 2) <> "OK" Then
      MsgBox "Não foi possível gerar os arquivos da NFS-e:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
      Exit Sub
   End If
   sXml = NFSe_Valor(oRes, "xml")
   sPdf = NFSe_Valor(oRes, "pdf")
End If
If Not ArquivoExiste(sXml) Or Not ArquivoExiste(sPdf) Then
   MsgBox "Arquivos da NFS-e não encontrados:" & vbCrLf & sXml & vbCrLf & sPdf, vbExclamation, "NFS-e"
   Exit Sub
End If

sEmail = Trim$(InputBox("E-mail para enviar a NFS-e nº " & sNumero & " de " & sTomador & vbCrLf & "(PDF e XML no mesmo e-mail):", _
                        "Enviar NFS-e por e-mail", sEmail))
If sEmail = "" Then Exit Sub
If InStr(1, sEmail, "@") = 0 Or InStr(1, sEmail, ".") = 0 Then
   MsgBox "E-mail inválido: " & sEmail, vbExclamation, "NFS-e"
   Exit Sub
End If

sRazao = Trim$(SQLExecutaRetorno("SELECT RAZAO FROM empresa", "RAZAO", ""))
sFantasia = Trim$(SQLExecutaRetorno("SELECT Fantasia FROM empresa", "Fantasia", ""))
If sFantasia = "" Then sFantasia = sRazao
sAssunto = "NFS-e nº " & sNumero & " - " & sFantasia & IIf(bCancelada, " (CANCELADA)", "")
sCorpo = "Olá" & IIf(sTomador <> "", ", " & sTomador, "") & ".<br><br>" & _
         "Segue em anexo a Nota Fiscal de Serviço eletrônica (NFS-e) nº <b>" & sNumero & "</b>, no valor de R$ " & sValor & _
         ": o PDF (DANFSe) e o arquivo XML.<br><br>" & _
         IIf(bCancelada, "<b>ATENÇÃO: esta NFS-e foi CANCELADA.</b><br><br>", "") & _
         "Chave de acesso: " & sChave & "<br>" & _
         "A autenticidade pode ser conferida no portal nacional da NFS-e (www.nfse.gov.br) ou pelo QR Code do PDF.<br><br>" & _
         "Atenciosamente,<br><br>" & sRazao

On Error GoTo deuErro
Set sistNFe = New snfe.Util
Call ConfiguraDLLNFeNFCe(55, "1", sistNFe)
ReDim pathAnexo(1)
pathAnexo(0) = sPdf
pathAnexo(1) = sXml
ReDim emailCC(0)
emailCC(0) = sEmail

Screen.MousePointer = vbHourglass
vRet = sistNFe.EmailEnviar(sEmail, sAssunto, sCorpo, pathAnexo, emailCC)
Screen.MousePointer = vbDefault
Set sistNFe = Nothing

If vRet Then
   MsgBox "E-mail enviado para " & sEmail & " com o PDF e o XML da NFS-e nº " & sNumero & ".", vbInformation, "NFS-e"
   'guarda o e-mail na nota quando ela ainda nao tinha (proximo envio ja vem preenchido)
   If Trim$(sEmail) <> "" Then dbData.Execute "UPDATE NFSe SET TomadorEmail = " & NFSe_Q(Left$(sEmail, 80)) & _
                                              " WHERE (IdNFSe = " & lId & ") AND ((TomadorEmail IS NULL) OR (TomadorEmail = ''))"
Else
   MsgBox "Falha ao enviar o e-mail. Verifique a configuração de envio de e-mail (a mesma da NF-e).", vbExclamation, "NFS-e"
End If
Exit Sub

deuErro:
Screen.MousePointer = vbDefault
MsgBox "Erro ao enviar o e-mail:" & vbCrLf & Err.Description, vbCritical, "NFS-e"
Err.Clear
Set sistNFe = Nothing
End Sub

Private Sub cmdCancelarNota_Click()
Dim oRes As Object
Dim sStatus As String, sEvento As String
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
frmNFSe_Cancelar.Confirmado = False
frmNFSe_Cancelar.lblNota.Caption = "NFS-e nº " & grdNotas.TextMatrix(grdNotas.Row, C_NUM) & " - " & grdNotas.TextMatrix(grdNotas.Row, C_CLI) & _
                                   " - R$ " & grdNotas.TextMatrix(grdNotas.Row, C_VALOR)
frmNFSe_Cancelar.Show vbModal
If Not frmNFSe_Cancelar.Confirmado Then
   Unload frmNFSe_Cancelar
   Exit Sub
End If
'motivo e justificativa servem tambem para o pedido de analise fiscal, se o municipio nao deixar cancelar direto
sEvento = "id=" & lId & vbCrLf & "motivo=" & frmNFSe_Cancelar.Motivo & vbCrLf & _
          "justificativa=" & Replace(Replace(frmNFSe_Cancelar.Justificativa, vbCr, " "), vbLf, " ")
Unload frmNFSe_Cancelar
sStatus = NFSe_Executar("acao=cancelar" & vbCrLf & sEvento, oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível cancelar:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
ElseIf NFSe_Valor(oRes, "analisefiscal") = "1" Then
   'E0822 prazo, E0823 valor, E0824 tomador: so a prefeitura pode cancelar (evento 101103)
   If MsgBox("O município não permite cancelar esta NFS-e direto pelo sistema:" & vbCrLf & vbCrLf & _
             NFSe_Valor(oRes, "mensagem") & vbCrLf & vbCrLf & _
             "Enviar um PEDIDO DE CANCELAMENTO para a prefeitura analisar (análise fiscal)?" & vbCrLf & vbCrLf & _
             "A NFS-e continua válida até a prefeitura deferir. Use Consultar nos próximos dias para acompanhar.", _
             vbQuestion + vbYesNo + vbDefaultButton1, "NFS-e") = vbYes Then
      sStatus = NFSe_Executar("acao=analisefiscal" & vbCrLf & sEvento, oRes)
      If Left$(sStatus, 2) <> "OK" Then
         MsgBox "Não foi possível enviar o pedido de cancelamento:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
      Else
         MsgBox "Pedido de cancelamento enviado à prefeitura." & vbCrLf & vbCrLf & NFSe_Valor(oRes, "mensagem"), vbInformation, "NFS-e"
      End If
   End If
ElseIf Val(NFSe_Valor(oRes, "situacao")) = NFSE_CANCELADA Then
   MsgBox "NFS-e cancelada." & vbCrLf & vbCrLf & "As OS dessa nota ficam livres para uma nova NFS-e.", vbInformation, "NFS-e"
Else
   MsgBox NFSe_Valor(oRes, "mensagem"), vbInformation, "NFS-e"
End If
Pesquisar
SelecionarId lId
End Sub

Private Sub cmdExcluir_Click()
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
If ShowMsg("Excluir o rascunho da NFS-e " & grdNotas.TextMatrix(grdNotas.Row, C_OS) & "?" & vbCrLf & vbCrLf & _
           "As OS voltam a ficar disponíveis para gerar outra NFS-e.", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
'so rascunho ou rejeitada (o Sefin nao gerou nota): NFSeItens sai junto (ON DELETE CASCADE)
dbData.Execute "DELETE FROM NFSe WHERE (IdNFSe = " & lId & ") AND (Situacao IN (0, 2));"
Pesquisar
End Sub

Private Sub cmdConfig_Click()
frmNFSe_Config.Show vbModal
End Sub

Private Sub cmdRecebidas_Click()
If Not NFSe_TemADN() Then
   MsgBox "Falta atualizar o banco de dados (script 180 - CreateTable_NFSeADN.sql) para baixar as NFS-e do Ambiente de Dados Nacional.", vbExclamation, "NFS-e"
   Exit Sub
End If
frmNFSe_ADN.Show vbModal
'a busca no ADN pode ter cancelado notas do sistema (cancelamento no portal / analise fiscal deferida)
Pesquisar
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub mskDe_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: mskAte.SetFocus
End Sub

Private Sub mskAte_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Pesquisar
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Pesquisar
End Sub

Private Sub mskDe_GotFocus()
SelectControl mskDe
End Sub

Private Sub mskAte_GotFocus()
SelectControl mskAte
End Sub
