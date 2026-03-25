VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Parcelas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PARCELAS"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9795
   Icon            =   "Parcelas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   9645
      TabIndex        =   50
      Top             =   60
      Width           =   9675
      Begin VB.TextBox txtCodFuncionario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8340
         TabIndex        =   106
         Top             =   180
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "Parcelas.frx":23D2
         Stretch         =   -1  'True
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS"
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
         TabIndex        =   51
         Top             =   120
         Width           =   1725
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   60
      TabIndex        =   26
      Top             =   720
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   15055
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
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
      TabCaption(0)   =   "À PAGAR"
      TabPicture(0)   =   "Parcelas.frx":D7CE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Picture2"
      Tab(0).Control(2)=   "frmPagamento"
      Tab(0).Control(3)=   "txtCodParc"
      Tab(0).Control(4)=   "frmParcela"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "HAVER"
      TabPicture(1)   =   "Parcelas.frx":D7EA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmHaver"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PAGAS"
      TabPicture(2)   =   "Parcelas.frx":D806
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblCliente"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblQuantHistorico"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "imgDesmarcadaPAGAS"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ImgMarcadaPAGAS"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label28"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label29"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label30"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblTotalSelecionadasQuitadas"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblTotalHistorico"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdImprimirParcQuitSel"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdImprimirParcelas"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "GridHaverPagas"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmdMostrarProdutosREATIVAR"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cmdMarcarTodasREATIVAR"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "cmdHabilitarREATIVAR"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtCONcodParc"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Picture3"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "frmReativar"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).ControlCount=   18
      Begin VB.Frame frmReativar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Parcela para reativar"
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
         ForeColor       =   &H00000080&
         Height          =   1035
         Left            =   120
         TabIndex        =   78
         Top             =   5520
         Width           =   9435
         Begin VB.TextBox txtCaixa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   180
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodCaixa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   180
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtConCodOS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   540
            Width           =   1215
         End
         Begin VB.TextBox txtConCodPedido 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtConNumParcela 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   540
            Width           =   435
         End
         Begin VB.TextBox txtConValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   540
            Width           =   1155
         End
         Begin MSMask.MaskEdBox mskConData 
            Height          =   315
            Left            =   2640
            TabIndex        =   83
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskConPgto 
            Height          =   315
            Left            =   3660
            TabIndex        =   84
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdReativar 
            Height          =   315
            Left            =   6000
            TabIndex        =   91
            Top             =   540
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "R&eativar"
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
            MICON           =   "Parcelas.frx":D822
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Origem"
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
            TabIndex        =   90
            Top             =   300
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pedido"
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
            TabIndex        =   89
            Top             =   300
            Width           =   600
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No."
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
            Left            =   2160
            TabIndex        =   88
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Venc."
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
            Left            =   2640
            TabIndex        =   87
            Top             =   300
            Width           =   510
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
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
            Left            =   4680
            TabIndex        =   86
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pgto."
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
            Left            =   3660
            TabIndex        =   85
            Top             =   300
            Width           =   465
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4065
         ScaleWidth      =   9405
         TabIndex        =   66
         Top             =   720
         Width           =   9435
         Begin VB.OptionButton optTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   3600
            TabIndex        =   128
            Top             =   120
            Width           =   1395
         End
         Begin VB.ComboBox cboAno 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8160
            Sorted          =   -1  'True
            TabIndex        =   127
            Top             =   60
            Width           =   1155
         End
         Begin VB.ComboBox cboMes 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6780
            TabIndex        =   126
            Top             =   60
            Width           =   1335
         End
         Begin VB.OptionButton optVenc 
            Caption         =   "Vencimento"
            Height          =   195
            Left            =   2340
            TabIndex        =   103
            Top             =   120
            Width           =   1395
         End
         Begin VB.OptionButton optPgto 
            Caption         =   "Pagamento"
            Height          =   195
            Left            =   1140
            TabIndex        =   102
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Historico 
            Height          =   3555
            Left            =   60
            TabIndex        =   74
            Top             =   420
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   6271
            _Version        =   393216
            FixedCols       =   0
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
         Begin VB.Label Label10 
            Caption         =   "Filtrar por:"
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
            Left            =   180
            TabIndex        =   101
            Top             =   120
            Width           =   1275
         End
      End
      Begin VB.Frame frmParcela 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Parcela Selecionada"
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
         Height          =   1815
         Left            =   -74280
         TabIndex        =   54
         Top             =   6120
         Visible         =   0   'False
         Width           =   4035
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   1500
            TabIndex        =   25
            Top             =   1260
            Width           =   1395
         End
         Begin VB.TextBox txtNumParcela 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox txtCodPedido 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   600
            Width           =   1275
         End
         Begin VB.TextBox txtOrigem 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   600
            Width           =   1215
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   1260
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632319
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
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
            Left            =   1500
            TabIndex        =   59
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimento"
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
            TabIndex        =   58
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parcela"
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
            Left            =   2700
            TabIndex        =   57
            Top             =   360
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1380
            TabIndex        =   56
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Origem"
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
            TabIndex        =   55
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.TextBox txtCONcodParc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3660
         TabIndex        =   49
         Top             =   4860
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCodParc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -70140
         TabIndex        =   46
         Top             =   6720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame frmPagamento 
         Caption         =   "Pagamento"
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
         Left            =   -68700
         TabIndex        =   30
         Top             =   5280
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox txtValorAutomatico 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1380
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   1380
            TabIndex        =   15
            Top             =   1980
            Width           =   1695
         End
         Begin ChamaleonBtn.chameleonButton cmdCal1 
            Height          =   315
            Left            =   2760
            TabIndex        =   7
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   600
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
            MICON           =   "Parcelas.frx":D83E
            PICN            =   "Parcelas.frx":D85A
            PICH            =   "Parcelas.frx":FBAD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboForma 
            Height          =   315
            Left            =   1380
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtTotalHaver 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1650
            Width           =   1695
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1695
         End
         Begin VB.TextBox txtDias 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Dias em atrazo"
            Top             =   960
            Width           =   435
         End
         Begin VB.TextBox txtJuros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1860
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Juros por dia"
            Top             =   960
            Width           =   435
         End
         Begin VB.TextBox txtMulta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1380
            TabIndex        =   13
            Top             =   1305
            Width           =   1695
         End
         Begin VB.TextBox txtTJuros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Valor do Juros"
            Top             =   960
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskPagamento 
            Height          =   315
            Left            =   1380
            TabIndex        =   6
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptChar      =   "_"
         End
         Begin VB.CheckBox chkJuros 
            Caption         =   "Juros/Dia (%):"
            Height          =   315
            Left            =   80
            TabIndex        =   8
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox chkMulta 
            Caption         =   "Multa (R$):"
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   1305
            Width           =   1095
         End
         Begin VB.Label lblValorAutomatico 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   960
            TabIndex        =   120
            Top             =   600
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblFormaPgto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pgto:"
            Height          =   195
            Left            =   240
            TabIndex        =   70
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lblHaverPgto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Haver:"
            Height          =   315
            Left            =   180
            TabIndex        =   47
            Top             =   1650
            Width           =   1110
         End
         Begin VB.Label lblPgto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Pagamento:"
            Height          =   315
            Left            =   240
            TabIndex        =   33
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label lblTotalPgto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Total:"
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
            Left            =   300
            TabIndex        =   32
            Top             =   2340
            Width           =   1005
         End
         Begin VB.Label lblJuros 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto (R$):"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   1995
            Width           =   1080
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   4695
         Left            =   -74880
         ScaleHeight     =   4635
         ScaleWidth      =   9375
         TabIndex        =   27
         Top             =   420
         Width           =   9435
         Begin VB.CheckBox chkCodPedido 
            Caption         =   "Cód. Pedido"
            Height          =   195
            Left            =   6120
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   60
            Width           =   1335
         End
         Begin VB.TextBox txtCodPed 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6120
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   300
            Width           =   1515
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Parcelas 
            Height          =   2535
            Left            =   60
            TabIndex        =   1
            Top             =   720
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4471
            _Version        =   393216
            FixedCols       =   0
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
         Begin VB.Frame Frame1 
            Caption         =   "Mostrar Juros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7680
            TabIndex        =   67
            Top             =   60
            Width           =   1635
            Begin VB.OptionButton optJurosNao 
               Caption         =   "Não"
               Height          =   195
               Left            =   840
               TabIndex        =   69
               Top             =   300
               Width           =   675
            End
            Begin VB.OptionButton optJurosSim 
               Caption         =   "Sim"
               Height          =   195
               Left            =   180
               TabIndex        =   68
               Top             =   300
               Value           =   -1  'True
               Width           =   795
            End
         End
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   60
            TabIndex        =   0
            Top             =   300
            Width           =   6015
         End
         Begin VB.TextBox txtCodCliente 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   5160
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   615
         End
         Begin ChamaleonBtn.chameleonButton cmdMostrarProdutos 
            Height          =   255
            Left            =   1500
            TabIndex        =   114
            Top             =   3300
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "MOSTRAR PRODUTOS"
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
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":11F00
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdMarcarCheck 
            Height          =   255
            Left            =   60
            TabIndex        =   115
            Top             =   3300
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "MARCAR TODAS"
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
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":11F1C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdMostrarHaveres 
            Height          =   255
            Left            =   3480
            TabIndex        =   124
            Top             =   3300
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "MOSTRAR HAVERES"
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
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":11F38
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "........Geral:........."
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
            TabIndex        =   118
            Top             =   3360
            Width           =   1545
         End
         Begin VB.Label lblQuantSel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6060
            TabIndex        =   117
            Top             =   3600
            Width           =   1005
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Quant.:"
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
            Left            =   5340
            TabIndex        =   116
            Top             =   3600
            Width           =   645
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Quant.:"
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
            Left            =   7560
            TabIndex        =   112
            Top             =   3600
            Width           =   645
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Devedor:"
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
            Left            =   5205
            TabIndex        =   111
            Top             =   4320
            Width           =   795
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
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
            Left            =   5220
            TabIndex        =   110
            Top             =   3840
            Width           =   780
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Haver(es):"
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
            Left            =   5100
            TabIndex        =   109
            Top             =   4080
            Width           =   900
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Devedor:"
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
            Left            =   7425
            TabIndex        =   108
            Top             =   4320
            Width           =   795
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
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
            Left            =   7440
            TabIndex        =   107
            Top             =   3840
            Width           =   780
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Haver(es):"
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
            Left            =   7320
            TabIndex        =   99
            Top             =   4080
            Width           =   900
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   ".....Selecionado:......"
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
            Left            =   5280
            TabIndex        =   98
            Top             =   3360
            Width           =   1785
         End
         Begin VB.Label lblHaverSel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00008000&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6060
            TabIndex        =   97
            ToolTipText     =   "Haveres"
            Top             =   4080
            Width           =   990
         End
         Begin VB.Label lblSubtotalSel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C00000&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6060
            TabIndex        =   96
            ToolTipText     =   "Sub-Total"
            Top             =   3840
            Width           =   990
         End
         Begin VB.Label lblTotalSel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000080&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6060
            TabIndex        =   93
            ToolTipText     =   "Sub-Total"
            Top             =   4320
            Width           =   990
         End
         Begin VB.Image imgDesmarcada 
            Height          =   195
            Left            =   2940
            Picture         =   "Parcelas.frx":11F54
            Top             =   3660
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image ImgMarcada 
            Height          =   195
            Left            =   2100
            Picture         =   "Parcelas.frx":142D0
            Top             =   3660
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblSubtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C00000&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   8280
            TabIndex        =   64
            ToolTipText     =   "Sub-Total"
            Top             =   3840
            Width           =   990
         End
         Begin VB.Label lblHaver 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00008000&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   8280
            TabIndex        =   63
            ToolTipText     =   "Haveres"
            Top             =   4080
            Width           =   990
         End
         Begin VB.Label lblQuantParc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   8280
            TabIndex        =   60
            Top             =   3600
            Width           =   1005
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000080&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   8280
            TabIndex        =   53
            ToolTipText     =   "Total"
            Top             =   4320
            Width           =   990
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do Cliente"
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
            Left            =   60
            TabIndex        =   29
            Top             =   60
            Width           =   1410
         End
      End
      Begin VB.Frame frmHaver 
         Caption         =   "Haver"
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
         ForeColor       =   &H00000000&
         Height          =   6255
         Left            =   -74880
         TabIndex        =   34
         Top             =   420
         Width           =   9435
         Begin VB.TextBox txtCodHaver 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            TabIndex        =   35
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox Picture7 
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   9135
            TabIndex        =   37
            Top             =   660
            Width           =   9195
            Begin VB.ComboBox cboFormaHaver 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   2340
               TabIndex        =   40
               Top             =   300
               Width           =   1875
            End
            Begin VB.TextBox txtValorHaver 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1140
               TabIndex        =   39
               Top             =   300
               Width           =   1155
            End
            Begin MSMask.MaskEdBox mskDataHaver 
               Height          =   315
               Left            =   60
               TabIndex        =   38
               Top             =   300
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12648447
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarHaver 
               Height          =   315
               Left            =   4320
               TabIndex        =   41
               Top             =   300
               Width           =   1335
               _ExtentX        =   2355
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
               MICON           =   "Parcelas.frx":166CF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverHaver 
               Height          =   315
               Left            =   5700
               TabIndex        =   43
               Top             =   300
               Width           =   1335
               _ExtentX        =   2355
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
               MICON           =   "Parcelas.frx":166EB
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdImprimirHaver 
               Height          =   315
               Left            =   7080
               TabIndex        =   123
               Top             =   300
               Width           =   1335
               _ExtentX        =   2355
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
               BCOL            =   12632256
               BCOLO           =   12632256
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Parcelas.frx":16707
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H80000009&
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de Pgto:"
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
               Left            =   2340
               TabIndex        =   92
               Top             =   60
               Width           =   1305
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data:"
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
               Left            =   60
               TabIndex        =   45
               Top             =   60
               Width           =   480
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor:"
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
               Left            =   1140
               TabIndex        =   44
               Top             =   60
               Width           =   510
            End
         End
         Begin VB.PictureBox Picture8 
            Height          =   4635
            Left            =   120
            ScaleHeight     =   4575
            ScaleWidth      =   9135
            TabIndex        =   36
            Top             =   1500
            Width           =   9195
            Begin MSFlexGridLib.MSFlexGrid Grid_Haver 
               Height          =   4215
               Left            =   60
               TabIndex        =   42
               Top             =   60
               Width           =   9015
               _ExtentX        =   15901
               _ExtentY        =   7435
               _Version        =   393216
               FixedCols       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin VB.Label lblQuantHaver 
               AutoSize        =   -1  'True
               BackColor       =   &H00000080&
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
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   60
               TabIndex        =   62
               Top             =   4320
               Width           =   225
            End
            Begin VB.Label lblTotalHaver 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "0,00"
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
               Left            =   8700
               TabIndex        =   61
               Top             =   4320
               Width           =   390
            End
         End
         Begin VB.Label lblClienteHaver 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Escolha o nome do cliente na aba BAIXA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   9195
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   73
         Top             =   5160
         Width           =   9435
         Begin VB.TextBox txtItem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4740
            TabIndex        =   122
            Top             =   1920
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Timer Timer1 
            Interval        =   600
            Left            =   0
            Top             =   1200
         End
         Begin ChamaleonBtn.chameleonButton cmdAlterar 
            Height          =   555
            Left            =   4260
            TabIndex        =   94
            Top             =   1200
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            MICON           =   "Parcelas.frx":16723
            PICN            =   "Parcelas.frx":1673F
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
            Height          =   555
            Left            =   4740
            TabIndex        =   95
            Top             =   2280
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            MICON           =   "Parcelas.frx":17019
            PICN            =   "Parcelas.frx":17035
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdQuitarUma 
            Height          =   435
            Left            =   60
            TabIndex        =   2
            Top             =   180
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   767
            BTYPE           =   3
            TX              =   "QUITAR A PARCELA"
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
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":1734F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdHabilitarHaver 
            Height          =   435
            Left            =   2400
            TabIndex        =   3
            Top             =   180
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   767
            BTYPE           =   3
            TX              =   "HAVER NA PARCELA"
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
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":1736B
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
            Height          =   255
            Left            =   6660
            TabIndex        =   17
            Top             =   2940
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "Salvar"
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
            MICON           =   "Parcelas.frx":17387
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCancelar 
            Height          =   255
            Left            =   8100
            TabIndex        =   20
            Top             =   2940
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
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
            MICON           =   "Parcelas.frx":173A3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdQuitarTodas 
            Height          =   255
            Left            =   6660
            TabIndex        =   19
            Top             =   2940
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "Salvar"
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
            MICON           =   "Parcelas.frx":173BF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdQuitarAutomatico 
            Height          =   435
            Left            =   60
            TabIndex        =   119
            Top             =   180
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   767
            BTYPE           =   3
            TX              =   "QUITAR AUTOMÁTICO"
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
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":173DB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdSalvarAutomatico 
            Height          =   255
            Left            =   6660
            TabIndex        =   18
            Top             =   2940
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "Salvar"
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
            MICON           =   "Parcelas.frx":173F7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAlterarVenc 
            Height          =   435
            Left            =   4740
            TabIndex        =   125
            Top             =   180
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   767
            BTYPE           =   3
            TX              =   "ALTERAR VENCIMENTO"
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
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":17413
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblHora 
            AutoSize        =   -1  'True
            Caption         =   "00:00"
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
            Left            =   -60
            TabIndex        =   113
            Top             =   1680
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Image imLogoCupom 
            Height          =   1125
            Left            =   360
            Picture         =   "Parcelas.frx":1742F
            Top             =   900
            Visible         =   0   'False
            Width           =   2850
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdHabilitarREATIVAR 
         Height          =   315
         Left            =   4200
         TabIndex        =   75
         Top             =   5160
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "REATIVAR"
         ENAB            =   0   'False
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Parcelas.frx":17F4C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMarcarTodasREATIVAR 
         Height          =   315
         Left            =   120
         TabIndex        =   76
         Top             =   5160
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "MARCAR TODAS"
         ENAB            =   0   'False
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Parcelas.frx":17F68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMostrarProdutosREATIVAR 
         Height          =   315
         Left            =   1920
         TabIndex        =   77
         Top             =   5160
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "MOSTRAR PRODUTOS"
         ENAB            =   0   'False
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Parcelas.frx":17F84
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridHaverPagas 
         Height          =   1455
         Left            =   120
         TabIndex        =   121
         Top             =   6900
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   2566
         _Version        =   393216
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimirParcelas 
         Height          =   315
         Left            =   5760
         TabIndex        =   130
         Top             =   5160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "RECIBO"
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
         MICON           =   "Parcelas.frx":17FA0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimirParcQuitSel 
         Height          =   315
         Left            =   5760
         TabIndex        =   135
         Top             =   5160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "RECIBO"
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
         MICON           =   "Parcelas.frx":17FBC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblTotalHistorico 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   8580
         TabIndex        =   134
         ToolTipText     =   "Sub-Total"
         Top             =   4860
         Width           =   990
      End
      Begin VB.Label lblTotalSelecionadasQuitadas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   8580
         TabIndex        =   133
         ToolTipText     =   "Haveres"
         Top             =   5100
         Width           =   990
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Selecionada(s):"
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
         Left            =   7185
         TabIndex        =   132
         Top             =   5100
         Width           =   1335
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Listada(s):"
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
         Left            =   7620
         TabIndex        =   131
         Top             =   4860
         Width           =   900
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HAVERES DA PARCELA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   129
         Top             =   6600
         Width           =   9435
      End
      Begin VB.Image ImgMarcadaPAGAS 
         Height          =   195
         Left            =   1320
         Picture         =   "Parcelas.frx":17FD8
         Top             =   4920
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgDesmarcadaPAGAS 
         Height          =   195
         Left            =   1620
         Picture         =   "Parcelas.frx":1A3D7
         Top             =   4920
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblQuantHistorico 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   4860
         Width           =   225
      End
      Begin VB.Label lblCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Escolha o nome do cliente na aba ABAIXO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   9435
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   100
      Top             =   9270
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10795
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2470
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18:19"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "Parcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Dim rEmpresa As ADODB.Recordset
Dim rPedidos As ADODB.Recordset
Dim rClientes As ADODB.Recordset
Dim rUsuario As ADODB.Recordset
Dim rParcelas As ADODB.Recordset

Dim var_Contador As Integer

Enum TipoOP
   MarcarTodos = 1
   DesmarcarTodos = 2
   contar = 3
End Enum                   'termina aqui

Private moCombo As cComboHelper

Dim CAIXA_FECHADO_BAIXA As Boolean      'verficar depois
Dim CAIXA_FECHADO_HAVER As Boolean      'verificar depois
Dim CAIXA_FECHADO_REATIVAR As Boolean   'verificar depois

Dim varCodCaixa As Long

Dim IMPRIMIR As Boolean
Dim var_ImpTermica As String
Dim var_ImpNormal As String
Dim varTipoRecPgto As String
Dim varTipoRecHaver As String
Dim oCfg As ConfigItem

Dim OP As TipoOP           'usado para o check das parcelas
Dim varPgtoAutomatico As Boolean
Dim i As Integer
Dim f As Integer
Dim iCodParc As Long
Dim vQuitarUma As Boolean
Dim vCodPed As Long
Dim vCodParc As Long
Dim vCidadeUF As String




Private Sub ConsultarCaixaAtual()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    varCodCaixa = ValidateNull(r("codcaixa"))
Else
    varCodCaixa = 0
End If
End Sub

Private Sub Limpar_Campos_Reativar()
txtConCodOS.Text = ""
txtConCodPedido.Text = ""
txtConNumParcela.Text = ""
mskConData.Text = ""
mskConPgto.Text = ""
txtCONValor.Text = ""
End Sub

Private Sub Reimprimir_HaverFolha()
i = Grid_Haver.Row
Dim vCodHaver As Long
If Grid_Haver.TextMatrix(i, 1) = "" Then Exit Sub
vCodHaver = Grid_Haver.TextMatrix(i, 1)


'Me.Show
'tabela empresa
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set rEmpresa = dbData.OpenRecordset(sSQL)
vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")

Set rPedidos = dbData.OpenRecordset("SELECT  pedidos.COD_CLIENTE as vCodCli, parcelas_haver.COD_FUNCIONARIO as vCodFunc, parcelas.CODIGO, parcelas.COD_PEDIDO as vCodPedido, parcelas.NUMERO as vNumParc, parcelas.VALOR_FINAL, parcelas_haver.CODIGO AS vCodHaver, parcelas_haver.NUMERO AS VnumHaver, parcelas_haver.VALOR_HAVER as vValorHaver, parcelas_haver.HAVER as vDataHaver, parcelas_haver.FORMA_PGTO as vForma " & _
               "FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN parcelas_haver ON parcelas.CODIGO = parcelas_haver.COD_PARCELA " & _
               "Where(parcelas_haver.CODIGO = " & vCodHaver & ");")

If Not rPedidos.EOF Then
    Dim vCodUsuario As Integer
    Dim vCodCliente As Integer
    vCodUsuario = rPedidos("vCodFunc")
    vCodCliente = rPedidos("vCodCli")
End If

'Set rParcelasHaver = dbData.OpenRecordset("SELECT FORMA_PGTO, VALOR_HAVER FROM parcelas_haver WHERE  (codigo = " & Grid_Haver.TextMatrix(i, 1) & ");")
'Set rPedidos = dbData.OpenRecordset("SELECT cod_pedido, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & vCodPed & ");")
Set rClientes = dbData.OpenRecordset("SELECT CODIGO, Nome FROM cliente WHERE  (CODIGO = " & vCodCliente & ");")
Set rUsuario = dbData.OpenRecordset("SELECT Codigo, Login FROM Usuario WHERE  (CODIGO = " & vCodUsuario & " );")
Me.Hide
With REL_Recibo
    '.txtUsuario.Caption = UCase(rUsuario("login"))
    If Not rUsuario.EOF Then
        .txtUsuario.Caption = UCase(rUsuario("login"))
    Else
        .txtUsuario.Caption = "Não Especificado"
    End If
    
    .txtFormaPgto.Caption = UCase(rPedidos("vForma"))
    .txtCliente.Caption = UCase(rClientes("Nome"))
    .txtValor.Caption = UCase(NumeroExtenso(rPedidos("vvalorhaver"), True))
    .txthead.Caption = "R$ " & Format(rPedidos("vvalorhaver"), ocMONEY)

   'If txtOrigem.Text <> "" And txtCodPedido.Text = "" Then
   '   .txtProveniente.Caption = "Haver da " & txtNumParcela.Text & "ª parcela da OS Nº " & Format(txtOrigem.Text, "000000")
   'ElseIf txtOrigem.Text = "" And txtCodPedido.Text <> "" Then
      .txtProveniente.Caption = "Haver da " & rPedidos("vnumparc") & "ª parcela do PEDIDO Nº " & Format(rPedidos("vCodPedido"), "000000")
   'End If

   .txtData.Caption = "" & vCidadeUF & ", " & Day(rPedidos("vDataHaver")) & " de " & MonthName(Month(rPedidos("vDataHaver"))) & " de " & Year(rPedidos("vDataHaver"))
   .Relatorio.NumeroRegistros = 1
   .Relatorio.NomeImpressora = var_ImpNormal
   .Relatorio.Ativar
End With
Unload REL_Recibo
Me.Show
End Sub
Private Sub ImprimirHaverFolha()


   'tabela empresa
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set rEmpresa = dbData.OpenRecordset(sSQL)
   vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")
   
   'Dim rUsuario As ADODB.Recordset
   Set rUsuario = dbData.OpenRecordset("SELECT codigo, login FROM Usuario WHERE  (codigo = " & txtCodFuncionario.Text & ");")
Me.Hide
With REL_Recibo
    '.txtUsuario.Caption = UCase(rUsuario("login"))
    If Not rUsuario.EOF Then
        .txtUsuario.Caption = UCase(rUsuario("login"))
    Else
        .txtUsuario.Caption = "Não Especificado"
    End If
    .txtFormaPgto.Caption = UCase(cboFormaHaver.Text)
    .txtCliente.Caption = UCase(cboCliente.Text)
    .txtValor.Caption = UCase(NumeroExtenso(txtValorHaver.Text, True))
    .txthead.Caption = "R$ " & Format(txtValorHaver.Text, ocMONEY)

   'If txtOrigem.Text <> "" And txtCodPedido.Text = "" Then
   '   .txtProveniente.Caption = "Haver da " & txtNumParcela.Text & "ª parcela da OS Nº " & Format(txtOrigem.Text, "000000")
   'ElseIf txtOrigem.Text = "" And txtCodPedido.Text <> "" Then
      .txtProveniente.Caption = "Haver da " & txtNumParcela.Text & "ª parcela do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
   'End If

   .txtData.Caption = "" & vCidadeUF & ", " & Day(mskDataHaver) & " de " & MonthName(Month(mskDataHaver)) & " de " & Year(mskDataHaver)
   .Relatorio.NumeroRegistros = 1
   .Relatorio.NomeImpressora = var_ImpNormal
   .Relatorio.Ativar
End With
Unload REL_Recibo
Me.Show
End Sub


Private Sub ImprimirPedidoFolhaSel()
If vQuitarUma = True Then
   If txtCodPedido.Text = "" Then
      ShowMsg "O código do PEDIDO está em branco !!", vbExclamation
      Exit Sub
   End If
End If

'tabela empresa
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set rEmpresa = dbData.OpenRecordset(sSQL)
vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")

Set rUsuario = dbData.OpenRecordset("SELECT codigo, login FROM Usuario WHERE  (codigo = " & txtCodFuncionario.Text & ");")

Me.Hide
With REL_Recibo
   .txtCliente.Caption = UCase(cboCliente.Text)
    If Not rUsuario.EOF Then
        .txtUsuario.Caption = UCase(rUsuario("login"))
    Else
        .txtUsuario.Caption = "Não Especificado"
    End If
   
   .txtFormaPgto.Caption = UCase(cboForma.Text)
   
   If vQuitarUma = True Then
      .txtValor.Caption = UCase(NumeroExtenso(txtTotal.Text, True))
      .txthead.Caption = "R$ " & Format(txtTotal.Text, "##,##0.00")
      .txtProveniente.Caption = "Pagamento da " & txtNumParcela.Text & "ª parcela do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
      .txtData.Caption = "" & vCidadeUF & ", " & Day(mskPagamento) & " de " & MonthName(Month(mskPagamento)) & " de " & Year(mskPagamento)
   Else
      Dim var_Parc As String
      Dim f As Integer

      var_Parc = ""
      
      With Grid_Historico
         For f = 1 To .rows - 1
            .Col = 0
            .Row = f

            If .CellPicture = ImgMarcadaPAGAS Then
               If f = 1 Then
                  var_Parc = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
               ElseIf Right(var_Parc, 8) = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00") Then
                  MsgBox "Tratar Repetido"
               Else
                  var_Parc = var_Parc & ", " & Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
                  If f = .rows - 1 Then Exit For
               End If
            End If
         Next f
      End With

      .txtValor.Caption = UCase(NumeroExtenso(lblTotalSelecionadasQuitadas.Caption, True))
      .txthead.Caption = "R$ " & Format(lblTotalSelecionadasQuitadas.Caption, "##,##0.00")
      .txtProveniente.Caption = "PEDIDO(S): " & var_Parc
      .txtData.Caption = "" & vCidadeUF & ", " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
   End If

   .Relatorio.NumeroRegistros = 1
   .Relatorio.NomeImpressora = var_ImpNormal
   .Relatorio.Ativar
End With
Unload REL_Recibo
Me.Show
End Sub

Private Sub ImprimirPedidoFolha()
If vQuitarUma = True Then
   If txtCodPedido.Text = "" Then
      ShowMsg "O código do PEDIDO está em branco !!", vbExclamation
      Exit Sub
   End If
End If

'tabela empresa
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set rEmpresa = dbData.OpenRecordset(sSQL)
vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")

Set rUsuario = dbData.OpenRecordset("SELECT codigo, login FROM Usuario WHERE  (codigo = " & txtCodFuncionario.Text & ");")

Me.Hide
With REL_Recibo
   .txtCliente.Caption = UCase(cboCliente.Text)
    If Not rUsuario.EOF Then
        .txtUsuario.Caption = UCase(rUsuario("login"))
    Else
        .txtUsuario.Caption = "Não Especificado"
    End If
   
   .txtFormaPgto.Caption = UCase(cboForma.Text)
   
   If vQuitarUma = True Then
      .txtValor.Caption = UCase(NumeroExtenso(txtTotal.Text, True))
      .txthead.Caption = "R$ " & Format(txtTotal.Text, "##,##0.00")
      .txtProveniente.Caption = "Pagamento da " & txtNumParcela.Text & "ª parcela do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
      .txtData.Caption = "" & vCidadeUF & ", " & Day(mskPagamento) & " de " & MonthName(Month(mskPagamento)) & " de " & Year(mskPagamento)
   Else
      Dim var_Parc As String
      Dim f As Integer

      var_Parc = ""
      
      With Grid_Parcelas
         For f = 1 To .rows - 1
            .Col = 0
            .Row = f

            If .CellPicture = ImgMarcada Then
               If f = 1 Then
                  var_Parc = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
               ElseIf Right(var_Parc, 8) = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00") Then
                  MsgBox "Tratar Repetido"
               Else
                  var_Parc = var_Parc & ", " & Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
                  If f = .rows - 1 Then Exit For
               End If
            End If
         Next f
      End With

      .txtValor.Caption = UCase(NumeroExtenso(lblTotalSel.Caption, True))
      .txthead.Caption = "R$ " & Format(lblTotalSel.Caption, "##,##0.00")
      .txtProveniente.Caption = "PEDIDO(S): " & var_Parc
      .txtData.Caption = "" & vCidadeUF & ", " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
   End If

   .Relatorio.NumeroRegistros = 1
   .Relatorio.NomeImpressora = var_ImpNormal
   .Relatorio.Ativar
End With
Unload REL_Recibo
Me.Show
End Sub


Public Sub LerConfiguracao()
   Dim sSQL As String            'Declara as variáveis
   Dim r As ADODB.Recordset
   Dim oCfg As ConfigItem
   
   Dim vValue As Variant
   Dim lDC As Long
   Dim cIni As Ini
   
   'Lê as configurações do banco de dados
   'Essas configurações são globais
   sSQL = "SELECT config_nome, config_valor FROM configuracao ORDER BY config_nome;"
   Set r = dbData.OpenRecordset(sSQL)
   'r.Open sSQL, dbData.ActiveConnection
   
   'Inicializa a coleção de configurações globais
   Set sysConfig = Nothing
   Set sysConfig = New Collection
   
   'Percorre a tabela até o fim
   Do While Not r.EOF
      'Cria um objeto ConfigItem e atribui os valores para cada configuração
      Set oCfg = New ConfigItem
      oCfg.SetValues r("config_nome"), r("config_valor")
      sysConfig.Add oCfg, oCfg.Name
      Set oCfg = Nothing
      r.MoveNext
   Loop
   
   'Fecha a tabela
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'Inicializa a coleção de configurações locais
   Set maqConfig = Nothing
   Set maqConfig = New Collection
   
   'Inicializa o objeto de controle de arquivos INI
   Set cIni = New Ini
   
   'Seta o nome do arquivo
   cIni.Arquivo = appPathIni
   
   'Recupera a configuração de atualização
   'vValue = cIni.LerTexto("GERAL", "URLAtualizacao", "\\HI-TECH02\PUBLICA\SOFTWARE\")
   'appURLUpdt = vValue
   
   'Destrói o objeto
   Set cIni = Nothing
End Sub

Private Sub FormatarGrid_HaverPagas(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridHaverPagas
      .Clear
      .Cols = 7
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1050
      .ColWidth(3) = 1050
      .ColWidth(4) = 2500
      .ColWidth(5) = 1200
      .ColWidth(6) = 1050
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "FORMA"
      .TextMatrix(0, 5) = "CÓD. CAIXA"
      .TextMatrix(0, 6) = "CAIXA"
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'Centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'mudar a cor da coluna
            For i = 1 To .rows - 1
               .Row = i
               .Col = 3
               .CellBackColor = &HC0C0FF
            Next
            
            .TextMatrix(.rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("haver"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("valor_haver"), ocMONEY)
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("forma_pgto"))
            .TextMatrix(.rows - 1, 5) = Format(rTabela("CODCAIXA"), "000000")
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("CAIXA"))
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   'lblTotalHaver.Caption = Format(SomaGrid(Grid_Haver, 5), ocMONEY)
   'txtTotalHaver.Text = Format(SomaGrid(Grid_Haver, 5), ocMONEY)
End Sub

Private Sub FormatarGrid_Haver(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Haver
      .Clear
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1200
      .ColWidth(3) = 1050
      .ColWidth(4) = 1050
      .ColWidth(5) = 1050
      .ColWidth(6) = 2500
      .ColWidth(7) = 1050
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "CÓD. CAIXA"
      .TextMatrix(0, 3) = "CAIXA"
      .TextMatrix(0, 4) = "DATA"
      .TextMatrix(0, 5) = "VALOR"
      .TextMatrix(0, 6) = "FORMA"
      .TextMatrix(0, 7) = "CÓD_PARCELA"

      
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'Centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'mudar a cor da coluna
            For i = 1 To .rows - 1
               .Row = i
               .Col = 4
               .CellBackColor = &HC0C0FF
            Next
            
            .TextMatrix(.rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("CODCAIXA"), "000000")
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("CAIXA"))
            .TextMatrix(.rows - 1, 4) = Format(rTabela("haver"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 5) = Format(rTabela("valor_haver"), ocMONEY)
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("forma_pgto"))
            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("COD_PARCELA"))
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   lblTotalHaver.Caption = Format(SomaGrid(Grid_Haver, 5), ocMONEY)
   txtTotalHaver.Text = Format(SomaGrid(Grid_Haver, 5), ocMONEY)
End Sub

Private Sub LimparGrid_Historico()
   Dim i As Integer
   
   With Grid_Historico
      .Visible = False
      
      .Clear
      .Cols = 9
      .rows = 2
      
      .ColWidth(0) = 300
      .ColWidth(1) = 500
      .ColWidth(2) = 900
      .ColWidth(3) = 900
      .ColWidth(4) = 500
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "PEDIDO"
      .TextMatrix(0, 4) = "No."
      .TextMatrix(0, 5) = "VENC."
      .TextMatrix(0, 6) = "PGTO"
      .TextMatrix(0, 7) = "VALOR"
      .TextMatrix(0, 8) = "TIPO"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      '.ColAlignment(2) = 1
      
      'Centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .rows = .rows + 1
      .Redraw = True
      .rows = .rows - 1
      .Visible = True
   End With
End Sub

Private Sub FormatarGrid_Historico(rTabela As ADODB.Recordset)
Dim i As Integer

With Grid_Historico
   .Visible = False
   
   .Clear
   .Cols = 15
   .rows = 2
   
   .ColWidth(0) = 300
   .ColWidth(1) = 0 '0
   .ColWidth(2) = 0 '0
   .ColWidth(3) = 750
   .ColWidth(4) = 450
   .ColWidth(5) = 900
   .ColWidth(6) = 900
   .ColWidth(7) = 800
   .ColWidth(8) = 900
   .ColWidth(9) = 750
   .ColWidth(10) = 850
   .ColWidth(11) = 900
   .ColWidth(12) = 800
   .ColWidth(13) = 800
   .ColWidth(14) = 850
   
   .TextMatrix(0, 1) = "CÓD"
   .TextMatrix(0, 2) = "ORIGEM"
   .TextMatrix(0, 3) = "PEDIDO"
   .TextMatrix(0, 4) = "No."
   .TextMatrix(0, 5) = "VENC."
   .TextMatrix(0, 6) = "VALOR"
   .TextMatrix(0, 7) = "JUROS"
   .TextMatrix(0, 8) = "SUBTOTAL"
   .TextMatrix(0, 9) = "DESC."
   .TextMatrix(0, 10) = "HAVER"
   .TextMatrix(0, 11) = "TOTAL"
   .TextMatrix(0, 12) = "PGTO"
   .TextMatrix(0, 13) = "CÓD."
   .TextMatrix(0, 14) = "CAIXA"
   
   .Redraw = False
   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   'Centralizar o titulo
   For i = 0 To .Cols - 1
      .Row = 0
      .Col = i
      .CellAlignment = flexAlignCenterCenter
   Next
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         'mudar a cor da coluna
         For i = 1 To .rows - 1
            .Row = i
            .Col = 11
            .CellBackColor = &HC0FFFF
         Next
         
         .TextMatrix(.rows - 1, 1) = rTabela("cod")
         .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("campo00"))
         .TextMatrix(.rows - 1, 3) = Format(rTabela("campo01"), "000000")
         .TextMatrix(.rows - 1, 4) = rTabela("campo02")
         .TextMatrix(.rows - 1, 5) = Format(rTabela("campo03"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 6) = FormatNumber(rTabela("campo04"), 2)
         .TextMatrix(.rows - 1, 7) = FormatNumber(rTabela("var_juros"), 2)
         .TextMatrix(.rows - 1, 8) = FormatNumber(rTabela("subtotal"), 2)
         .TextMatrix(.rows - 1, 9) = FormatNumber(rTabela("vardesc"), 2)
         .TextMatrix(.rows - 1, 10) = FormatNumber(rTabela("varSomaHaveres"), 2)
         .TextMatrix(.rows - 1, 11) = FormatNumber(rTabela("vValorFinal"), 2)
         .TextMatrix(.rows - 1, 12) = Format(rTabela("campo06"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 13) = Format(rTabela("VARCODCAIXAPARC"), "000000")
         .TextMatrix(.rows - 1, 14) = ValidateNull(rTabela("varCaixaParc"))
         
         
         '.TextMatrix(.rows - 1, 1) = rTabela("cod")
         '.TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("campo00"))
         '.TextMatrix(.rows - 1, 3) = Format(rTabela("campo01"), "000000")
         '.TextMatrix(.rows - 1, 4) = rTabela("campo02")
         '.TextMatrix(.rows - 1, 5) = Format(rTabela("campo03"), "dd/mm/yy")
         '.TextMatrix(.rows - 1, 6) = Format(rTabela("campo04"), ocMONEY)
'        ' .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_dias"))
         '.TextMatrix(.rows - 1, 8) = Format(rTabela("var_juros"), ocMONEY)
         '.TextMatrix(.rows - 1, 9) = Format(rTabela("vardesc"), ocMONEY)
         '.TextMatrix(.rows - 1, 10) = Format(rTabela("vValorFinal"), ocMONEY)
         '.TextMatrix(.rows - 1, 11) = Format(rTabela("VARCODCAIXAPARC"), "000000")
         '.TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("varCaixaParc"))
         '.TextMatrix(.rows - 1, 13) = Format(rTabela("campo06"), "dd/mm/yy")
         
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .Redraw = True
   .rows = .rows - 1
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 2
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
   
   For i = 1 To .rows - 1
      .Row = i
      .Col = 10
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
   
   .Col = 0
   For i = 1 To .rows - 1
      .Row = i
      Set .CellPicture = imgDesmarcadaPAGAS
      .CellPictureAlignment = 4
   Next
   
   .Redraw = True
   .Visible = True
End With

lblTotalHistorico.Caption = Format(SomaGrid(Grid_Historico, 11), ocMONEY)
End Sub

Private Sub FormatarGrid_Parcelas2(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Parcelas
      .Clear
      .Cols = 11
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1050
      .ColWidth(3) = 850
      .ColWidth(4) = 850
      .ColWidth(5) = 500
      .ColWidth(6) = 850
      .ColWidth(7) = 850
      .ColWidth(8) = 1150
      .ColWidth(9) = 850
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "TIPO"
      .TextMatrix(0, 3) = "OS"
      .TextMatrix(0, 4) = "PEDIDO"
      .TextMatrix(0, 5) = "No."
      .TextMatrix(0, 6) = "VENC."
      .TextMatrix(0, 7) = "VALOR"
      .TextMatrix(0, 8) = "HAVER(ES)"
      .TextMatrix(0, 9) = "TOTAL"
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      '.ColAlignment(2) = 1
      
      'Centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'mudar a cor da coluna
            For i = 1 To .rows - 1
               .Row = i
               .Col = 7:   .CellBackColor = &HC0FFFF
               .Col = 9:   .CellBackColor = &HC0C0FF
            Next
            
            .TextMatrix(.rows - 1, 1) = rTabela("cod")
            .TextMatrix(.rows - 1, 2) = rTabela("campo05")
            .TextMatrix(.rows - 1, 3) = rTabela("campo00")
            .TextMatrix(.rows - 1, 4) = Format(rTabela("campo01"), "000000")
            .TextMatrix(.rows - 1, 5) = rTabela("campo02")
            .TextMatrix(.rows - 1, 6) = Format(rTabela("campo03"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 7) = Format(rTabela("campo04"), ocMONEY)
            
            If Not IsNull(rTabela("campo06")) Then
               .TextMatrix(.rows - 1, 8) = Format(rTabela("campo06"), ocMONEY)
            Else
               .TextMatrix(.rows - 1, 8) = Format(0, ocMONEY)
            End If
            
            .TextMatrix(.rows - 1, 9) = Format(.TextMatrix(.rows - 1, 7) - .TextMatrix(.rows - 1, 8), ocMONEY)
            .TextMatrix(.rows - 1, 10) = rTabela("var_atrazo")
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .Redraw = True
      .rows = .rows - 1
   End With
   
   lblSubtotal.Caption = Format(SomaGrid(Grid_Parcelas, 7), ocMONEY)
   lblHaver.Caption = Format(SomaGrid(Grid_Parcelas, 8), ocMONEY)
   lblTotal.Caption = Format(SomaGrid(Grid_Parcelas, 9), ocMONEY)
End Sub

Private Sub FormatarGrid_Parcelas(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With Grid_Parcelas
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 13
      .rows = 2
      
      .ColWidth(0) = 300
      .ColWidth(1) = 0
      .ColWidth(2) = 1050
      .ColWidth(3) = 850
      .ColWidth(4) = 400
      .ColWidth(5) = 950
      .ColWidth(6) = 900
      .ColWidth(7) = 650
      .ColWidth(8) = 850
      .ColWidth(9) = 900
      .ColWidth(10) = 900
      .ColWidth(11) = 900
      .ColWidth(12) = 900
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "ORIGEM"
      .TextMatrix(0, 3) = "CÓDIGO"
      .TextMatrix(0, 4) = "No."
      .TextMatrix(0, 5) = "VENC"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "DIAS"
      .TextMatrix(0, 8) = "JUROS"
      .TextMatrix(0, 9) = "TOTAL"
      .TextMatrix(0, 10) = "HAVER"
      .TextMatrix(0, 11) = "DEVE"
      .TextMatrix(0, 12) = "ITEM"
     '
      

      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(2) = 3
      .ColAlignment(3) = 3
      .ColAlignment(5) = 3
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = ValidateNull(rTabela("codparcela"))
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("campo00"))
            .TextMatrix(.rows - 1, 3) = Format(rTabela("campo01"), "000000")
            .TextMatrix(.rows - 1, 4) = rTabela("campo02")
            .TextMatrix(.rows - 1, 5) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 6) = Format(rTabela("valor"), ocMONEY)
            
            If optJurosSim = True Then
               .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_atrazo"))
               .TextMatrix(.rows - 1, 8) = Format(rTabela("var_juros"), ocMONEY)
               .TextMatrix(.rows - 1, 9) = Format(rTabela("varTotalComJuros"), ocMONEY)
               .TextMatrix(.rows - 1, 10) = Format(rTabela("varsomahaveres"), ocMONEY)
               .TextMatrix(.rows - 1, 11) = Format(rTabela("varTotalDevedor"), ocMONEY)
            Else
               .TextMatrix(.rows - 1, 7) = Format(0, "0")
               .TextMatrix(.rows - 1, 8) = Format(0, ocMONEY)
               .TextMatrix(.rows - 1, 9) = Format(rTabela("valor"), ocMONEY)
               .TextMatrix(.rows - 1, 10) = Format(rTabela("varsomahaveres"), ocMONEY)
               .TextMatrix(.rows - 1, 11) = Format(rTabela("valor") - rTabela("varsomahaveres"), ocMONEY)
            End If
            .TextMatrix(.rows - 1, 12) = rTabela("vItem")

            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      .rows = .rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 5
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 6
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 9
         .CellForeColor = &HC00000
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 10
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 11
         .CellForeColor = &H80&
         .CellFontBold = True
      Next
      
      'Deixar negrito quando vencido
      For i = 1 To .rows - 1
         For j = 0 To .Cols - 1
            .Col = j
            .Row = i
            
            If CDate(.TextMatrix(i, 5)) < Date Then
               .CellForeColor = vbRed
            ElseIf CDate(.TextMatrix(i, 5)) = Date Then
               .CellForeColor = vbBlue
            End If
         Next
      Next
      
      'Grid_Parcelas.ColWidth(0) = 400
      'Grid_Parcelas.Rows = 11
      Grid_Parcelas.Col = 0
      
      For i = 1 To .rows - 1
         Grid_Parcelas.Row = i
         Set Grid_Parcelas.CellPicture = imgDesmarcada
         Grid_Parcelas.CellPictureAlignment = 4
      Next
      
      .Visible = True
      .Redraw = True
   End With
   
   lblSubtotal.Caption = Format(SomaGrid(Grid_Parcelas, 9), ocMONEY)
   lblHaver.Caption = Format(SomaGrid(Grid_Parcelas, 10), ocMONEY)
   lblTotal.Caption = Format(SomaGrid(Grid_Parcelas, 11), ocMONEY)
End Sub

Private Sub Calcular_Dias()
If DateDiff("d", mskData.Text, mskPagamento.Text) < 0 Then
   txtDias.Text = "0"
Else
   txtDias.Text = DateDiff("d", mskData.Text, mskPagamento.Text)
End If
End Sub

Private Sub Calcular_Juros()
If txtValor.Text = "" Or txtJuros.Text = "" Or txtDias.Text = "" Then Exit Sub

'Dim var_Dias As Integer
'Dim var_Juros As Currency

'var_Dias = txtDias.Text

'var_Juros = (CDbl(txtJuros.Text) / 100) * txtValor.Text
'var_Juros = var_Juros * var_Dias
'txtTJuros.Text = FormatNumber(var_Juros, 2)

Dim var_Dias As Integer
Dim var_VALOR As Currency
Dim var_JurosDia As Double

var_Dias = txtDias.Text
var_VALOR = txtValor.Text
var_JurosDia = txtJuros.Text

'txtTJuros.Text = Format((((var_VALOR * var_JurosDia) / 100) * var_Dias), ocMONEY)
End Sub

Private Sub LimparGrid_Haver()
   Dim i As Integer
   
   With Grid_Haver
      .Clear
      .Cols = 4
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 1050
      .ColWidth(2) = 1050
      .ColWidth(3) = 1050
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "VALOR"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      '.ColAlignment(2) = 1
      
      'Centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      .rows = .rows + 1
      
      .rows = .rows - 1
    .Redraw = True
   End With
   
   lblTotalHaver.Caption = Format(SomaGrid(Grid_Haver, 3), ocMONEY)
End Sub

Private Sub LimparGrid_Parcelas()
   Dim i As Integer
   
   With Grid_Parcelas
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 12
      .rows = 2
      
      .ColWidth(0) = 300
      .ColWidth(1) = 0
      .ColWidth(2) = 1050
      .ColWidth(3) = 850
      .ColWidth(4) = 400
      .ColWidth(5) = 800
      .ColWidth(6) = 1100
      .ColWidth(7) = 650
      .ColWidth(8) = 750
      .ColWidth(9) = 850
      .ColWidth(10) = 1150
      .ColWidth(11) = 1000
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "ORIGEM"
      .TextMatrix(0, 3) = "CÓDIGO"
      .TextMatrix(0, 4) = "No."
      .TextMatrix(0, 5) = "VENC"
      .TextMatrix(0, 6) = "SUBTOTAL"
      .TextMatrix(0, 7) = "DIAS"
      .TextMatrix(0, 8) = "JUROS"
      .TextMatrix(0, 9) = "TOTAL"
      .TextMatrix(0, 10) = "HAVER(ES)"
      .TextMatrix(0, 11) = "LIQUIDO"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(2) = 3
      .ColAlignment(3) = 3
      .ColAlignment(5) = 3
      .rows = .rows + 1
      
      i = i + 1
      .rows = .rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 9
         .CellForeColor = &HC00000
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 10
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 11
         .CellForeColor = &H80&
         .CellFontBold = True
      Next
      
      'Grid_Parcelas.ColWidth(0) = 400
      'Grid_Parcelas.Rows = 11
      Grid_Parcelas.Col = 0
      
      For i = 1 To .rows - 1
         Grid_Parcelas.Row = i
         Set Grid_Parcelas.CellPicture = imgDesmarcada
         Grid_Parcelas.CellPictureAlignment = 4
      Next
      
      .Visible = True
      .Redraw = True
   End With
   
   lblSubtotal.Caption = Format(SomaGrid(Grid_Parcelas, 9), ocMONEY)
   lblHaver.Caption = Format(SomaGrid(Grid_Parcelas, 10), ocMONEY)
   lblTotal.Caption = Format(SomaGrid(Grid_Parcelas, 11), ocMONEY)
End Sub

Private Sub LimparGridHaverPagas()
sSQL = "SELECT * FROM parcelas_haver WHERE 1 = 0;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_HaverPagas r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub LimparObjetos_GridParcelas()
lblQuantParc.Caption = 0
lblSubtotal.Caption = FormatNumber(0, 2)
lblHaver.Caption = FormatNumber(0, 2)
lblTotal.Caption = FormatNumber(0, 2)
End Sub

Private Sub LimparObjetos_Parcelas()
chkMulta.Value = 0
chkJuros.Value = 0
txtOrigem.Text = ""
txtCodPedido.Text = ""
txtNumParcela.Text = ""
mskData.Mask = ""
mskData.Text = ""
txtValor.Text = ""
mskPagamento.Mask = ""
mskPagamento.Text = ""
txtDias.Text = ""
'txtJuros.Text = ""
txtTJuros.Text = ""
txtMulta.Text = ""
txtTotalHaver.Text = ""
txtTotal.Text = ""
txtDesconto.Text = ""
mskDataHaver.Mask = ""
mskDataHaver.Text = ""
txtValorHaver.Text = ""
lblClienteHaver.Caption = ""
lblCliente.Caption = ""
cboForma.Text = ""
frmHaver.Enabled = False
'frmParcela.Visible = False
frmPagamento.Visible = False
cmdQuitarUma.Visible = False
cmdHabilitarHaver.Visible = False
cmdAlterarVenc.Visible = False
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = False
cmdExcluir.Visible = False
txtMulta.Text = FormatNumber(0, 2)
lblTotalHaver.Caption = FormatNumber(0, 2)
lblTotalHistorico.Caption = FormatNumber(0, 2)
lblQuantHaver.Caption = 0
lblQuantHistorico.Caption = 0
LimparGrid_Haver
'LimparGrid_Historico
End Sub

Private Sub LimparObjetos_Historico()
txtCONcodParc.Text = ""
txtConCodOS.Text = ""
txtConCodPedido.Text = ""
txtConNumParcela.Text = ""
mskConData.Text = ""
mskConPgto.Text = ""
txtCONValor.Text = ""
mskPagamento.Mask = ""
mskPagamento.Text = ""
txtCodCaixa.Text = ""
txtCaixa.Text = ""
End Sub

Private Sub MostraHaveresPagas()
If Grid_Historico.rows >= 2 Then
    i = Grid_Historico.Row
        
    sSQL = "SELECT * FROM parcelas_haver WHERE (cod_parcela = " & Grid_Historico.TextMatrix(i, 1) & ") ORDER BY haver, codigo;"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid_HaverPagas r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If
End Sub

Private Sub Mostrar_Juros()
Dim oCfg As ConfigItem

Set oCfg = sysConfig("JUROS_DIA")
txtJuros.Text = CCur(oCfg.Value)
Set oCfg = Nothing
End Sub

Private Sub Mostrar_Pedidos_Por_Codigo()
'Dim sSQL As String
Dim r As ADODB.Recordset
If txtCodPed.Text = "" Then Exit Sub

sSQL = "SELECT cliente.*, pedidos.* FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
   "WHERE (cod_pedido = " & txtCodPed.Text & ");"

Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtCodCliente.Text = r("cod_cliente")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub MostrarGrid_Historico()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Dim INDICE As String
Dim vWhere As String

 'indice
 If optPgto.Value = True Then
    INDICE = "parcelas.pagamento "
    vWhere = " AND (Month(parcelas.pagamento) = " & cboMES.ListIndex + 1 & ") And (Year(parcelas.pagamento) = " & cboAno & ") "
 ElseIf optVenc.Value = True Then
    INDICE = "parcelas.data "
    vWhere = " AND (Month(parcelas.data) = " & cboMES.ListIndex + 1 & ") And (Year(parcelas.data) = " & cboAno & ") "
 ElseIf optTodas.Value = True Then
    INDICE = "parcelas.pagamento "
    vWhere = " "
 End If

If txtCodCliente.Text = "" Then Exit Sub
If cboMES.Text = "" Then cboMES.Text = Format(Date, "mmmm")
If cboAno.Text = "" Then cboAno.Text = Year(Date)

sSQL = "SELECT parcelas.CODIGO AS cod, ISNULL(parcelas.CODCAIXA, 0) AS varCodCaixaParc, ISNULL(parcelas.CAIXA, 'CAIXA01') AS varCaixaParc, pedidos.TIPO_PEDIDO AS campo00, parcelas.COD_PEDIDO AS campo01, parcelas.NUMERO AS campo02, parcelas.DATA AS campo03,pedidos.PAGAMENTO AS campo05, parcelas.PAGAMENTO AS campo06, parcelas.STATUS, pedidos.COD_PEDIDO, pedidos.COD_CLIENTE, parcelas.VALOR AS campo04, ISNULL(parcelas.JUROS, 0) AS var_juros, parcelas.VALOR + ISNULL(parcelas.JUROS, 0) AS SubTotal, ISNULL(parcelas.DESCONTO, 0) AS varDesc, parcelas.VALOR_FINAL AS vValorFinal, " & _
            "(SELECT ISNULL(SUM(VALOR_HAVER), 0) FROM parcelas_haver WHERE (COD_PARCELA = parcelas.CODIGO)) AS varSomaHaveres " & _
            "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
            "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
            "WHERE (cliente.codigo = " & txtCodCliente.Text & ") AND (parcelas.status = 1) " & vWhere & "  ORDER BY  " & INDICE

'sSQL = "SELECT parcelas.codigo AS cod, isnull(parcelas.CODCAIXA, 0) AS varCodCaixaParc, isnull(parcelas.CAIXA, 'CAIXA01') AS varCaixaParc, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, isnull(parcelas.desconto,0) AS varDesc," & _
   "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, parcelas.pagamento AS campo06, " & _
   "isnull(parcelas.dias_atrazo,0) AS var_dias, isnull(parcelas.juros,0) AS var_juros, isnull(parcelas.MULTA,0), isnull(parcelas.DESCONTO,0), parcelas.VALOR_FINAL AS vValorFinal, parcelas.status, " & _
   "pedidos.cod_pedido, pedidos.cod_cliente FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
   "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
   "WHERE (cliente.codigo = " & txtCodCliente.Text & ") AND (parcelas.status = 1) " & vWhere & "  ORDER BY  " & INDICE
  
'Debug.Print sSQL

Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If txtCodCliente.Text <> "1" Then FormatarGrid_Historico r
If r.State <> 0 Then r.Close
Set r = Nothing

lblQuantHistorico.Caption = Format(totalRegistros, "00") & " parcela(s)"
cmdMarcarTodasREATIVAR.Enabled = Grid_Historico.rows > 1
End Sub

Private Sub MostrarGrid_Parcelas()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long

Dim oCfg As ConfigItem
Dim var_JurosDia As Double
Dim var_tipojuros As Integer

'MOSTRAR NO GRID AS PARCELAS À PAGAR
If txtCodCliente.Text = "" Then Exit Sub

Set oCfg = sysConfig("JUROS_DIA")
var_JurosDia = CCur(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("TIPO_JUROS")
var_tipojuros = oCfg.Value
Set oCfg = Nothing

Dim var_CampoFuros As String

If var_tipojuros = 0 Then              'Juros sobre o saldo restante
   var_CampoFuros = "(parcelas.valor - (SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)))"
ElseIf var_tipojuros = 1 Then          'Juros sobre o valor da parcela
   var_CampoFuros = "parcelas.valor"
End If

If txtCodCliente.Text <> "" Then
   If optJurosSim.Value = True Then
   
      sSQL = "SELECT CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END AS var_atrazo, (((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END) AS var_juros, " & _
         "(parcelas.valor + ((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * (CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END)) AS varTotalComJuros, cliente.codigo, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.juros AS varvalorjuros, cliente.nome, parcelas.codigo AS codparcela, ISNULL(parcelas.OS_ITEM, 0) as vItem, parcelas.data, parcelas.valor, parcelas.pagamento AS var_parcpgto, pedidos.pagamento AS var_tipopgto, " & _
         "(SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)) AS varSomaHaveres, ((parcelas.valor + (((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * (CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END))) - (SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) AS varTotalDevedor, " & _
         "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago, pedidos.*, CASE parcelas.cod_pedido WHEN 0 THEN 'OS' ELSE 'PEDIDO' END AS var_tipo, " & _
         "CASE parcelas.cod_pedido WHEN 0 THEN parcelas.cod_os ELSE parcelas.cod_pedido END AS var_codigo " & _
         "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (cliente.codigo = " & txtCodCliente.Text & ") AND (parcelas.status = 0) ORDER BY parcelas.data, parcelas.codigo;"
  
  Else
  
     sSQL = "SELECT CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END AS var_atrazo, (((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END) AS var_juros, " & _
         "(parcelas.valor + ((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * (CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END)) AS varTotalComJuros, cliente.codigo, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.juros AS varvalorjuros, cliente.nome, parcelas.codigo AS codparcela, ISNULL(parcelas.OS_ITEM, 0) as vItem, parcelas.data, parcelas.valor, parcelas.pagamento AS var_parcpgto, pedidos.pagamento AS var_tipopgto, " & _
         "(SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)) AS varSomaHaveres, ((parcelas.valor + (((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * (CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END))) - (SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) AS varTotalDevedor, " & _
         "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago, pedidos.*, CASE parcelas.cod_pedido WHEN 0 THEN 'OS' ELSE 'PEDIDO' END AS var_tipo, " & _
         "CASE parcelas.cod_pedido WHEN 0 THEN parcelas.cod_os ELSE parcelas.cod_pedido END AS var_codigo " & _
         "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (cliente.codigo = " & txtCodCliente.Text & ") AND (parcelas.status = 0) ORDER BY parcelas.data, parcelas.codigo;"
         'Debug.Print sSQL
  
     'esse é o codigo correto para não mostrar o juros (fazer depois)... o de cima mostra o juros
      'sSQL = "SELECT 0 AS var_atrazo, (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0) AS var_juros, " & _
         "(parcelas.valor + (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0)) AS var_total, cliente.codigo, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.juros AS varvalorjuros, cliente.nome, parcelas.codigo AS codparcela, parcelas.*, parcelas.pagamento AS var_parcpgto, pedidos.pagamento AS var_tipopgto, " & _
         "(SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)) AS varSomaHaveres, ((parcelas.valor + (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0)) - (SELECT SUM(valor_haver) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) AS campo07, " & _
         "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago, pedidos.*, CASE parcelas.cod_pedido WHEN 0 THEN 'OS' ELSE 'PEDIDO' END AS var_tipo, " & _
         "CASE parcelas.cod_pedido WHEN 0 THEN parcelas.cod_os ELSE parcelas.cod_pedido END AS var_codigo " & _
         "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (cliente.codigo = " & txtCodCliente.Text & ") AND (parcelas.status = 0) ORDER BY parcelas.data, parcelas.codigo;"
   
   End If
Else
   sSQL = "SELECT 0 AS var_atrazo, (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0) AS var_juros, " & _
      "(parcelas.valor + (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0)) AS var_total, cliente.codigo, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
      "parcelas.juros AS varvalorjuros, cliente.nome, parcelas.codigo AS codparcela, ISNULL(parcelas.OS_ITEM, 0) as vItem, parcelas.data, parcelas.valor, parcelas.pagamento AS var_parcpgto, pedidos.pagamento AS var_tipopgto, " & _
      "(SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)) AS varSomaHaveres, ((parcelas.valor + (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0)) - (SELECT SUM(valor_haver) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) AS campo07, " & _
      "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago, pedidos.*, CASE parcelas.cod_pedido WHEN 0 THEN 'OS' ELSE 'PEDIDO' END AS var_tipo, " & _
      "CASE parcelas.cod_pedido WHEN 0 THEN parcelas.cod_os ELSE parcelas.cod_pedido END AS var_codiog " & _
      "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
      "WHERE false ORDER BY parcelas.data, parcelas.codigo;"
   
End If
Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid_Parcelas r

If r.State <> 0 Then r.Close
Set r = Nothing

lblQuantParc.Caption = Format(totalRegistros, "00")

cmdMarcarCheck.Enabled = Grid_Parcelas.rows > 1

If Grid_Parcelas.rows > 1 Then
    cmdQuitarAutomatico.Visible = True
Else
    cmdQuitarAutomatico.Visible = False
End If
End Sub

Private Sub ReimprimirPedidoFolha()
Dim vValorTotal As Currency
Dim vNumParc As Integer
Dim vDataPgto As Date

For f = 0 To Grid_Historico.rows - 1
   Grid_Historico.Row = f
   Grid_Historico.Col = 0
   
   If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then
      vCodPed = Grid_Historico.TextMatrix(Grid_Historico.Row, 3)
      vValorTotal = Grid_Historico.TextMatrix(Grid_Historico.Row, 11)
      vNumParc = Grid_Historico.TextMatrix(Grid_Historico.Row, 4)
      vDataPgto = Grid_Historico.TextMatrix(Grid_Historico.Row, 12)
   End If
Next

'If vQuitarUma = True Then
'   If txtCodPedido.Text = "" Then
'      ShowMsg "O código do PEDIDO está em branco !!", vbExclamation
'      Exit Sub
'   End If
'End If

'tabela empresa
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set rEmpresa = dbData.OpenRecordset(sSQL)
vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")

Set rPedidos = dbData.OpenRecordset("SELECT cod_pedido, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & vCodPed & ");")
Set rParcelas = dbData.OpenRecordset("SELECT parcelas.FORMA_PGTO FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO WHERE  (pedidos.cod_pedido = " & vCodPed & ");")
Set rClientes = dbData.OpenRecordset("SELECT cliente.CODIGO, cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO WHERE  (pedidos.cod_pedido = " & vCodPed & ");")
Set rUsuario = dbData.OpenRecordset("SELECT Usuario.Codigo, Usuario.Login FROM pedidos INNER JOIN Usuario ON pedidos.COD_FUNCIONARIO = Usuario.Codigo WHERE  (pedidos.cod_pedido = " & vCodPed & ");")


'Set rUsuario = dbData.OpenRecordset("SELECT codigo, login FROM Usuario WHERE  (codigo = " & txtCodFuncionario.Text & ");")

Me.Hide
With REL_Recibo
   .txtCliente.Caption = UCase(rClientes("Nome"))
    If Not rUsuario.EOF Then
        .txtUsuario.Caption = UCase(rUsuario("login"))
    Else
        .txtUsuario.Caption = "Não Especificado"
    End If
   
   .txtFormaPgto.Caption = UCase(rParcelas("FORMA_PGTO"))
   
    .txtValor.Caption = UCase(NumeroExtenso(vValorTotal, True))
    .txthead.Caption = "R$ " & Format(vValorTotal, "##,##0.00")
    .txtProveniente.Caption = "Pagamento da " & vNumParc & "ª parcela do PEDIDO Nº " & Format(vCodPed, "000000")
    .txtData.Caption = "" & vCidadeUF & ", " & Day(vDataPgto) & " de " & MonthName(Month(vDataPgto)) & " de " & Year(vDataPgto)

   .Relatorio.NumeroRegistros = 1
   .Relatorio.NomeImpressora = var_ImpNormal
   .Relatorio.Ativar
End With
Unload REL_Recibo
Me.Show
End Sub

Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Currency
Dim i As Integer, Valor As Currency

Valor = 0
For i = 0 To Grid.rows - 1
   If IsNumeric(Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CCur(Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function

Private Sub Somar_Parcelas_SelecionadasQuitadas()
Dim Total As Currency
Dim i As Integer
Dim varContarSelecionadas As Integer

Total = 0
varContarSelecionadas = 0

With Grid_Historico
    If .TextMatrix(.Row, 10) = "" Then Exit Sub
   For i = 1 To .rows - 1
      .Col = 0
      .Row = i
      
      If .CellPicture = ImgMarcadaPAGAS Then
         .Col = 10
         Total = Total + .TextMatrix(.Row, 10)
      End If
   Next
   
   'lblQuantSel.Caption = Format(varContarSelecionadas, "00")
   lblTotalSelecionadasQuitadas.Caption = FormatNumber(Total, 2)
   'lblHaverSel.Caption = Format(HAVER, ocMONEY)
   'lblTotalSel.Caption = Format(Total, ocMONEY)
End With
End Sub

Private Sub Somar_Parcelas_Selecionadas()
Dim Total As Currency, SUBTOTAL As Currency, HAVER As Currency
Dim i As Integer
Dim varContarSelecionadas As Integer

SUBTOTAL = 0
HAVER = 0
Total = 0
varContarSelecionadas = 0

With Grid_Parcelas
    If .TextMatrix(.Row, 9) = "" Then Exit Sub
   For i = 1 To .rows - 1
      .Col = 0
      .Row = i
      
      If .CellPicture = ImgMarcada Then
         .Col = 6
         SUBTOTAL = SUBTOTAL + .TextMatrix(.Row, 9)
         varContarSelecionadas = varContarSelecionadas + 1
         .Col = 10
         HAVER = HAVER + .TextMatrix(.Row, 10)
         .Col = 11
         Total = Total + .TextMatrix(.Row, 11)
      End If
   Next
   
   lblQuantSel.Caption = Format(varContarSelecionadas, "00")
   lblSubtotalSel.Caption = Format(SUBTOTAL, ocMONEY)
   lblHaverSel.Caption = Format(HAVER, ocMONEY)
   lblTotalSel.Caption = Format(Total, ocMONEY)
End With
End Sub
Private Sub cboAno_Click()
cboAno_LostFocus
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

Private Sub cboAno_LostFocus()
MostrarGrid_Historico
End Sub


Private Sub cboMes_Click()
cboMes_LostFocus
End Sub

Private Sub cboMes_GotFocus()
cboMES.Clear

cboMES.AddItem "Janeiro"
cboMES.AddItem "Fevereiro"
cboMES.AddItem "Março"
cboMES.AddItem "Abril"
cboMES.AddItem "Maio"
cboMES.AddItem "Junho"
cboMES.AddItem "Julho"
cboMES.AddItem "Agosto"
cboMES.AddItem "Setembro"
cboMES.AddItem "Outubro"
cboMES.AddItem "Novembro"
cboMES.AddItem "Dezembro"

moCombo.AttachTo cboMES
End Sub

Private Sub cboCliente_Change()
'If chkCodPedido.Value = Unchecked Then CboCliente_LostFocus
End Sub

Private Sub cboCliente_Click()
If chkCodPedido.Value = Unchecked Then CboCliente_LostFocus
End Sub

Private Sub CboCliente_GotFocus()
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboCliente.Text
codAtual = txtCodCliente.Text
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

cboCliente.Text = itemAtual
txtCodCliente.Text = codAtual

SelectControl cboCliente

moCombo.AttachTo cboCliente
End Sub

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then cboCliente_Click
End Sub

Private Sub CboCliente_LostFocus()
If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
If cboCliente.Locked = True Then Exit Sub

On Error GoTo TrataErro

If vClienteEncontrado = False Then
    If cboCliente.ListIndex = -1 Then
        txtCodCliente.Text = ""
        cboCliente.Text = ""
        lblCliente.Caption = ""
        LimparObjetos_Parcelas
        LimparGrid_Parcelas
        LimparObjetos_GridParcelas
    Else
        If chkCodPedido.Value = Unchecked Then txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
        If chkCodPedido.Value = Unchecked Then Exit Sub
        lblCliente.Caption = cboCliente.Text
    End If
End If

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboForma_GotFocus()
cboForma.Clear
cboForma.AddItem "DINHEIRO"
cboForma.AddItem "CHEQUE"
cboForma.AddItem "CARTÃO - DÉBITO"
cboForma.AddItem "CARTÃO - CRÉDITO"
cboForma.AddItem "DEPOSITO"
cboForma.AddItem "TRANSFERÊNCIA"
cboForma.AddItem "BOLETO"
cboForma.AddItem "PIX"
moCombo.AttachTo cboForma
End Sub

Private Sub cboFormaHaver_GotFocus()
cboFormaHaver.Clear
cboFormaHaver.AddItem "DINHEIRO"
cboFormaHaver.AddItem "CHEQUE"
cboFormaHaver.AddItem "CARTÃO - DÉBITO"
cboFormaHaver.AddItem "CARTÃO - CRÉDITO"
cboFormaHaver.AddItem "PIX"
cboFormaHaver.AddItem "DEPOSITO"
cboFormaHaver.AddItem "TRANSFERÊNCIA"
   
If cboFormaHaver.ListCount <> 0 Then cboFormaHaver.ListIndex = 0
moCombo.AttachTo cboFormaHaver
End Sub

Private Sub cboMes_LostFocus()
MostrarGrid_Historico
End Sub

Private Sub chkCodPedido_Click()
   If chkCodPedido.Value = Checked Then txtCodPed.Enabled = True: txtCodPed.SetFocus Else txtCodPed.Enabled = False
End Sub

Private Sub chkJuros_Click()
If txtValor.Text = "" Or txtJuros.Text = "" Or txtDias.Text = "" Then Exit Sub

If chkJuros.Value = 1 Then
   lblJuros.Enabled = True
   txtJuros.Enabled = True
   txtTJuros.Enabled = True
   MostrarGrid_Haver
   txtJuros.SetFocus
Else
   lblJuros.Enabled = False
   txtJuros.Enabled = False
   txtTJuros.Enabled = False
   MostrarGrid_Haver
End If
End Sub

Private Sub chkMulta_Click()
   If chkMulta.Value = 1 Then
      txtMulta.Enabled = True
      txtMulta.SetFocus
      MostrarGrid_Haver
   Else
      txtMulta.Enabled = False
      MostrarGrid_Haver
   End If
End Sub

Private Sub cmdAdicionarHaver_Click()
ConsultarCaixaAtual

If varCodCaixa = 0 Then
    MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If cboFormaHaver.Text = "" Then MsgBox "Escolha uma forma de pagamento!", vbInformation, "Aviso do Sistema": Exit Sub

Dim varTipoCartao As String
varTipoCartao = "NULL"

If cboFormaHaver.Text = "CARTÃO - DÉBITO" Then
   varTipoCartao = "'D'"
ElseIf cboFormaHaver.Text = "CARTÃO - CRÉDITO" Then
   varTipoCartao = "'C'"
Else
    varTipoCartao = "NULL"
End If

Dim var_PAGAMENTO As String
If cboFormaHaver.Text = "DINHEIRO" Then
   var_PAGAMENTO = "DINHEIRO"
ElseIf cboFormaHaver.Text = "CARTÃO - DÉBITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboFormaHaver.Text = "CARTÃO - CRÉDITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboFormaHaver.Text = "CHEQUE" Then
   var_PAGAMENTO = "CHEQUE"
ElseIf cboFormaHaver.Text = "TRANSFERÊNCIA" Then
   var_PAGAMENTO = "TRANSFERENCIA"
ElseIf cboFormaHaver.Text = "DEPOSITO" Then
   var_PAGAMENTO = "DEPOSITO"
ElseIf cboFormaHaver.Text = "PIX" Then
   var_PAGAMENTO = "PIX"
End If

'calcular valor possivel do haver
Dim varValorParc As Currency
Dim varValorHaver As Currency

Dim f As Integer

For f = 0 To Grid_Parcelas.rows - 1
   Grid_Parcelas.Row = f
   Grid_Parcelas.Col = 0
   
   If Grid_Parcelas.CellPicture = ImgMarcada Then
        varValorParc = Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 11)
   End If
Next

varValorHaver = CCur(txtValorHaver.Text)

If varValorHaver > varValorParc Then MsgBox "O valor do haver não pode ultrapassar o valor da parcela!", vbInformation, "Aviso do Sistema": Exit Sub

'===================inicio
Dim lNovoCod As Long

If txtCodCliente.Text = "" Or txtValorHaver.Text = "" Or mskDataHaver.Text = "" Then Exit Sub

'ADICIONAR O HAVER NA TABELA HAVER
AutoNumeracao_Haver

dbData.Execute "INSERT INTO parcelas_haver (codigo, cod_parcela, numero, vencimento, haver, valor_parcela, valor_haver, hora, forma_pgto, caixa, CODCAIXA, tipo, tipo_cartao, COD_FUNCIONARIO) VALUES (" & _
   txtCodHaver.Text & ", " & txtCodParc.Text & ", " & txtNumParcela.Text & ", CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103), CONVERT(DATETIME, '" & Format(mskDataHaver.Text, ocDATA) & "', 103), " & _
   Replace(CCur(txtValor.Text), ",", ".") & ", " & Replace(CCur(txtValorHaver.Text), ",", ".") & ", '" & Format(lblHora, ocHRMN) & "', '" & var_PAGAMENTO & "', '" & StatusBar1.Panels(2).Text & "', " & varCodCaixa & ", 'PARCELA', " & varTipoCartao & ", " & txtCodFuncionario & ");"

'ADICIONAR NA TABELA CAIXA_ENTRADA
'lNovoCod = AutoNumeracao_Caixa

'dbData.Execute "INSERT INTO caixa_entrada (codigo, descricao, data, valor, setor, hora, cod_haver, forma_pgto, caixa) VALUES (" & _
'   lNovoCod & ", '" & "HAVER: " & cboCliente.Text & "', '" & Format(mskDataHaver.Text, ocDATA_EUA) & "', " & Replace(CCur(txtValorHaver.Text), ",", ".") & ", '" & _
'   IIf(txtOrigem.Text = "", "LOJA", "OFICINA") & "', '" & Format(lblHora, ocHRMN) & "', " & txtCodHaver.Text & ", '" & cboFormaHaver.Text & "', '" & StatusBar1.Panels(2).Text & "')"

'MARCAR O CAMPO HAVER DA TABELA PARCELAS
dbData.Execute "UPDATE parcelas SET haver = 1 WHERE (codigo = " & txtCodParc.Text & ");"

If varPgtoAutomatico = False Then
    If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
       IMPRIMIR = True
    Else
       IMPRIMIR = False
    End If
    
    If IMPRIMIR = True Then
       If IMPRIMIR = True Then
            'If vQuitarUma = True Then
                If varTipoRecHaver = "CUPOM" Then
                    Imprimir_HaverCupom
                Else
                    ImprimirHaverFolha
                End If
            'Else
            '    ImprimirHaverFolha
            'End If
       End If
    End If
Else
    IMPRIMIR = True
    If IMPRIMIR = True Then
       If IMPRIMIR = True Then
            'If vQuitarUma = True Then
                If varTipoRecHaver = "CUPOM" Then
                    Imprimir_HaverCupom
                Else
                    ImprimirHaverFolha
                End If
            'Else
            '    ImprimirHaverFolha
            'End If
       End If
    End If
End If
MostrarGrid_Parcelas
Calcular_Valores
MostrarGrid_Haver
'LimparObjetos_Parcelas_Haver
mskDataHaver.Text = Format(Date, "dd/mm/yy")
txtValorHaver.Text = ""
cboFormaHaver.Text = ""
Somar_Parcelas_Selecionadas
'txtValorHaver.SetFocus
SSTab1.Tab = 0
End Sub

Private Function AutoNumeracao_Caixa() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM caixa_entrada;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao_Caixa = lRet
End Function

Private Sub cmdAlterar_Click()
   If cboCliente.Text = "" And txtNumParcela.Text = "" And txtValor.Text = "" Then Exit Sub
   If txtCodPedido.Text = "" And txtOrigem.Text = "" Then Exit Sub
   
   If txtOrigem.Text <> "" And txtCodPedido.Text = "" Then
      If ShowMsg("Deseja alterar a parcela '" & txtNumParcela.Text & "' da OS No.'" & txtOrigem.Text & "' ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   ElseIf txtOrigem.Text = "" And txtCodPedido.Text <> "" Then
      If ShowMsg("Deseja alterar a parcela '" & txtNumParcela.Text & "' do Pedido No. '" & txtCodPedido.Text & "' ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   End If
   
   If txtOrigem.Text <> "" And txtCodPedido.Text = "" Then
      dbData.Execute "UPDATE parcelas SET valor = " & Replace(CCur(txtValor.Text), ",", ".") & ", data = CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103) WHERE (cod_os = " & txtOrigem.Text & ") AND (numero = " & txtNumParcela.Text & ");"
   ElseIf txtOrigem.Text = "" And txtCodPedido.Text <> "" Then
      dbData.Execute "UPDATE parcelas SET valor = " & Replace(CCur(txtValor.Text), ",", ".") & ", data = CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103) WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = " & txtNumParcela.Text & ");"
   End If
   
   MostrarGrid_Parcelas
   Calcular_Valores
   LimparGrid_Haver
   LimparObjetos_Parcelas
   cboCliente.SetFocus
End Sub

Private Sub cmdAlterarVenc_Click()
frmPagamento.Visible = True
mskPagamento.Visible = True
mskPagamento.SetFocus
cmdCal1.Visible = True
lblPgto.Visible = True
cmdSalvar.Visible = True
cmdCancelar.Visible = True
cmdQuitarUma.Visible = False
cmdHabilitarHaver.Visible = False
cmdAlterarVenc.Visible = False
lblFormaPgto.Visible = False
cboForma.Visible = False
lblValorAutomatico.Visible = False
txtValorAutomatico.Visible = False
cmdSalvarAutomatico.Visible = False
chkJuros.Visible = False
txtDias.Visible = False
txtJuros.Visible = False
txtTJuros.Visible = False
chkMulta.Visible = False
txtMulta.Visible = False
lblHaverPgto.Visible = False
txtTotalHaver.Visible = False
lblJuros.Visible = False
txtDesconto.Visible = False
lblTotalPgto.Visible = False
txtTotal.Visible = False
mskPagamento.Top = 240
lblPgto.Top = 240
cmdCal1.Top = 240
lblPgto.Caption = "Vencimento"
frmPagamento.Caption = "Vencimento"
cboCliente.Locked = True

For f = 0 To Grid_Parcelas.rows - 1
   Grid_Parcelas.Row = f
   Grid_Parcelas.Col = 0
   
   If Grid_Parcelas.CellPicture = ImgMarcada Then
      vCodParc = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 1))
   End If
Next
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

mskPagamento = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
mskPagamento.SetFocus
End Sub

Private Sub cmdCancelar_Click()
LimparObjetos_Parcelas
Dim i As Integer

If cmdQuitarTodas.Visible = True Then
    cmdQuitarTodas.Visible = False

    Grid_Parcelas.Col = 0
   
    For i = 1 To Grid_Parcelas.rows - 1
      Grid_Parcelas.Row = i
        Set Grid_Parcelas.CellPicture = imgDesmarcada
    Next
    Somar_Parcelas_Selecionadas
ElseIf cmdSalvarAutomatico.Visible = True Then
    frmPagamento.Visible = False
    lblFormaPgto.Visible = False
    cboForma.Visible = False
    lblValorAutomatico.Visible = False
    txtValorAutomatico.Visible = False
    cmdSalvarAutomatico.Visible = False
    cmdCancelar.Visible = False
End If


Grid_Parcelas.Col = 0

For i = 1 To Grid_Parcelas.rows - 1
  Grid_Parcelas.Row = i
    Set Grid_Parcelas.CellPicture = imgDesmarcada
Next
Somar_Parcelas_Selecionadas
cmdMarcarCheck.Caption = "MARCAR TODAS"
cmdQuitarAutomatico.Visible = True
frmPagamento.Visible = False
cmdCancelar.Visible = False
vQuitarUma = False
cboCliente.Locked = False
mskPagamento.Top = 600
lblPgto.Top = 600
cmdCal1.Top = 600
End Sub

Private Sub cmdHabilitarHaver_Click()
frmHaver.Enabled = True
mskDataHaver.Text = Format(Date, "dd/mm/yy")

For f = 0 To Grid_Parcelas.rows - 1
   Grid_Parcelas.Row = f
   Grid_Parcelas.Col = 0
   
   If Grid_Parcelas.CellPicture = ImgMarcada Then
      txtCodParc.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 1))
      txtOrigem.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 2))
      txtCodPedido.Text = Format((Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 3)), "000000")
      txtNumParcela.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 4))
      mskData.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 5))
      txtValor.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 6))
      MostrarGrid_Haver
   End If
Next

SSTab1.Tab = 1
Grid_Haver.Enabled = True
cmdAdicionarHaver.Enabled = True
cmdRemoverHaver.Enabled = True
cmdImprimirHaver.Enabled = False
txtValorHaver.SetFocus
End Sub

Private Sub cmdImprimirHaver_Click()
If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
   IMPRIMIR = True
Else
   IMPRIMIR = False
End If

If IMPRIMIR = True Then
   If IMPRIMIR = True Then
        If varTipoRecHaver = "CUPOM" Then
            Reimprimir_HaverCupom
        Else
            Reimprimir_HaverFolha
        End If
   End If
End If
End Sub

Private Sub cmdImprimirParcelas_Click()
If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo) = vbYes Then
    If varTipoRecPgto = "CUPOM" Then
        Reimprimir_ReciboCupom
    Else
        ReimprimirPedidoFolha
    End If
End If
End Sub

Private Sub cmdImprimirParcQuitSel_Click()
vQuitarUma = False
If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo) = vbYes Then
    If varTipoRecPgto = "CUPOM" Then
        Imprimir_ReciboCupomSel
    Else
        ImprimirPedidoFolhaSel
    End If
End If
vQuitarUma = True
End Sub


Private Sub cmdMostrarHaveres_Click()
For f = 0 To Grid_Parcelas.rows - 1
   Grid_Parcelas.Row = f
   Grid_Parcelas.Col = 0
   
   If Grid_Parcelas.CellPicture = ImgMarcada Then
      txtCodParc.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 1))
      txtOrigem.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 2))
      txtCodPedido.Text = Format((Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 3)), "000000")
      txtNumParcela.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 4))
      mskData.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 5))
      txtValor.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 6))
      MostrarGrid_Haver
   End If
Next

SSTab1.Tab = 1
frmHaver.Enabled = True
Grid_Haver.Enabled = True
cmdAdicionarHaver.Enabled = False
cmdRemoverHaver.Enabled = False
cmdImprimirHaver.Enabled = True
End Sub

Private Sub cmdQuitarAutomatico_Click()
frmPagamento.Visible = True
lblFormaPgto.Visible = True
cboForma.Visible = True
lblValorAutomatico.Visible = True
txtValorAutomatico.Visible = True
cmdSalvarAutomatico.Visible = True
cmdCancelar.Visible = True
lblPgto.Visible = False
mskPagamento.Visible = False
cmdCal1.Visible = False
chkJuros.Visible = False
txtDias.Visible = False
txtJuros.Visible = False
txtTJuros.Visible = False
chkMulta.Visible = False
txtMulta.Visible = False
lblHaverPgto.Visible = False
txtTotalHaver.Visible = False
lblJuros.Visible = False
txtDesconto.Visible = False
lblTotalPgto.Visible = False
txtTotal.Visible = False
cboForma.SetFocus
End Sub

Private Sub cmdQuitaruma_Click()
For f = 0 To Grid_Parcelas.rows - 1
   Grid_Parcelas.Row = f
   Grid_Parcelas.Col = 0
   
   If Grid_Parcelas.CellPicture = ImgMarcada Then
      txtCodParc.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 1))
      txtOrigem.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 2))
      txtCodPedido.Text = Format((Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 3)), "000000")
      txtNumParcela.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 4))
      mskData.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 5))
      txtValor.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 6))
      txtItem.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 12))
      txtTJuros.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 8))
      
      mskPagamento.Text = Format(Date, "dd/mm/yy")
      Mostrar_Juros
      chkJuros.Value = Checked
      Calcular_Dias
      MostrarGrid_Haver
   End If
Next

frmParcela.Visible = True
frmPagamento.Visible = True
cmdSalvar.Visible = True
cmdCancelar.Visible = True
cmdQuitarUma.Visible = False
cmdHabilitarHaver.Visible = False
cmdAlterarVenc.Visible = False
'cmdExcluir.Visible = True
vQuitarUma = True
cboForma.SetFocus
End Sub

Private Sub cmdHabilitarREATIVAR_Click()


For f = 0 To Grid_Historico.rows - 1
   Grid_Historico.Row = f
   Grid_Historico.Col = 0
   
   If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then
      frmReativar.Enabled = True
      cmdReativar.Visible = True
      txtCONcodParc.Text = (Grid_Historico.TextMatrix(Grid_Historico.Row, 1))
      txtConCodOS.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 2)), "000000")
      txtConCodPedido.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 3)), "000000")
      txtConNumParcela.Text = (Grid_Historico.TextMatrix(Grid_Historico.Row, 4))
      mskConData.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 5)), "dd/mm/yy")
      mskConPgto.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 12)), "dd/mm/yy")
      txtCONValor.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 6)), "##,##0.00")
      txtCodCaixa.Text = (Grid_Historico.TextMatrix(Grid_Historico.Row, 13))
      txtCaixa.Text = (Grid_Historico.TextMatrix(Grid_Historico.Row, 14))
   End If
Next
End Sub

Private Sub Reimprimir_HaverCupom()
    i = Grid_Haver.Row
    Dim vCodHaver As Long
    If Grid_Haver.TextMatrix(i, 1) = "" Then Exit Sub
    vCodHaver = Grid_Haver.TextMatrix(i, 1)
    
    'Me.Hide
    'Me.Show
    'tabela empresa
    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set rEmpresa = dbData.OpenRecordset(sSQL)
    vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")
    
    Set rPedidos = dbData.OpenRecordset("SELECT  pedidos.COD_CLIENTE as vCodCli, parcelas_haver.COD_FUNCIONARIO as vCodFunc, parcelas.CODIGO, parcelas.COD_PEDIDO as vCodPedido, parcelas.NUMERO as vNumParc, parcelas.VALOR_FINAL, parcelas_haver.CODIGO AS vCodHaver, parcelas_haver.NUMERO AS VnumHaver, parcelas_haver.VALOR_HAVER as vValorHaver, parcelas_haver.HAVER as vDataHaver, parcelas_haver.FORMA_PGTO as vForma " & _
                   "FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN parcelas_haver ON parcelas.CODIGO = parcelas_haver.COD_PARCELA " & _
                   "Where(parcelas_haver.CODIGO = " & vCodHaver & ");")
    
    If Not rPedidos.EOF Then
        Dim vCodUsuario As Integer
        Dim vCodCliente As Integer
        vCodUsuario = rPedidos("vCodFunc")
        vCodCliente = rPedidos("vCodCli")
    End If
    
    'Set rParcelasHaver = dbData.OpenRecordset("SELECT FORMA_PGTO, VALOR_HAVER FROM parcelas_haver WHERE  (codigo = " & Grid_Haver.TextMatrix(i, 1) & ");")
    'Set rPedidos = dbData.OpenRecordset("SELECT cod_pedido, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & vCodPed & ");")
    Set rClientes = dbData.OpenRecordset("SELECT CODIGO, Nome FROM cliente WHERE  (CODIGO = " & vCodCliente & ");")
    Set rUsuario = dbData.OpenRecordset("SELECT Codigo, Login FROM Usuario WHERE  (CODIGO = " & vCodUsuario & " );")

   'Recupera um número de arquivo disponível
   f = FreeFile()
      
   'pegar o nome da impressora no ini
   Dim oIni As Ini
   ''Dim var_ImpTermica As String
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
   Set oIni = Nothing
   
   Dim Prt As Printer
   Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   For Each Prt In Printers
      If Prt.DeviceName = var_ImpTermica Then
         Set Printer = Prt
         Exit For
      End If
   Next
   
   With Printer
      .ScaleMode = vbPixels
      .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
      
      For i = 1 To 6
         Printer.Print " "
      Next
      
      .ScaleMode = vbCentimeters
      .FontName = "courier new"
      '.PrintQuality = vbPRPQHigh

      Fonte 10, True, False
      Printer.Print Tab((35 - Len(rEmpresa("fantasia"))) / 2); rEmpresa("fantasia")   'Esse /2 é p/ centralizar
      Fonte 10, False, False
      Printer.Print Tab((35 - Len(rEmpresa("razao"))) / 2); rEmpresa("razao")
      Fonte 8, False, False
      Printer.Print rEmpresa("endereco") & ", " & rEmpresa("cidade") & "-" & rEmpresa("estado")
      Printer.Print "FONE: "; rEmpresa("telefone")                                        '& " - (89) 9986-3739"
      Fonte 8, False, False
      Printer.Print "CNPJ:"; rEmpresa("cnpj") & "  IE:" & rEmpresa("ie")
      Fonte 8, False, False
      Printer.Print String(40, "-")
      
       For i = 1 To 2
         Printer.Print " "
      Next
      
      Fonte 10, True, False
      Printer.Print Tab(10); "R E C I B O"
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Fonte 8, True, False
      Printer.Print Tab(28); "R$ " & Format(rPedidos("vvalorhaver"), "##,##0.00")
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Dim Line1 As String
      Dim Line2 As String
      
      Dim Texto As String
      Texto = UCase(NumeroExtenso(rPedidos("vvalorhaver"), True))
      Line1 = Mid(Texto, 1, 40)
      Line2 = Mid(Texto, 41, 80)
     
      Fonte 8, False, False
      Printer.Print Tab(2); "Recebi(emos) de: "
      Fonte 8, True, False
      Printer.Print Tab(2); rClientes("Nome")
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Fonte 8, False, False
      Printer.Print Tab(2); "A importância supra de: "
      Fonte 8, False, False
      Printer.Print Tab(2); Line1
      Printer.Print Tab(2); Line2
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Fonte 8, False, False
      Printer.Print Tab(2); "Proveniente do: "; "Haver da " & rPedidos("vnumparc") & "ª parcela "
      Printer.Print Tab(2); "do PEDIDO Nº " & Format(rPedidos("vCodPedido"), "000000")
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Printer.Print Tab(2); "Forma de Pgto: " & rPedidos("vForma")
      'Printer.Print Tab(2); "Funcionário: "; rUsuario("login")
    If Not rUsuario.EOF Then
        Printer.Print Tab(2); "Funcionário: "; rUsuario("login")
    Else
        Printer.Print Tab(2); "Funcionário: Não Especificado"
    End If

      For i = 1 To 2
         Printer.Print " "
      Next
     
      Fonte 8, False, False
      Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(rPedidos("vDataHaver")) & " de " & MonthName(Month(rPedidos("vDataHaver"))) & " de " & Year(rPedidos("vDataHaver"))
      
      For i = 1 To 3
            Printer.Print " "
      Next
      
      Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
      Printer.Print Tab((40 - Len("Assinatura")) / 2); "Assinatura"
      
       Close #f
       .EndDoc
    End With

 If Not rEmpresa Is Nothing Then If rEmpresa.State <> 0 Then rEmpresa.Close
 If Not rPedidos Is Nothing Then If rPedidos.State <> 0 Then rPedidos.Close
 If Not rClientes Is Nothing Then If rClientes.State <> 0 Then rClientes.Close
 If Not rUsuario Is Nothing Then If rUsuario.State <> 0 Then rUsuario.Close

Tratar_Erro:
   If Err.Number = 52 Then
     ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
     Printer.KillDoc
     Exit Sub
   End If
End Sub
Private Sub Imprimir_HaverCupom()
   'On Error GoTo Tratar_Erro
   Dim sSQL As String
   Dim rEmpresa As ADODB.Recordset
   
   Dim rPedidos As ADODB.Recordset
   Dim rClientes As ADODB.Recordset
   'Dim rUsuario As ADODB.Recordset
   
   Dim i As Integer
   Dim f As Integer
   
   If txtCodPedido.Text = "" Then Exit Sub
   
   'tabela empresa
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set rEmpresa = dbData.OpenRecordset(sSQL)
   vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")
   
   'tabela pedidos
   Set rPedidos = dbData.OpenRecordset("SELECT cod_pedido, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");")

   'tabela pedidos
   Set rClientes = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE  (codigo = " & txtCodCliente.Text & ");")
   
   If txtCodFuncionario.Text = "" Then txtCodFuncionario.Text = "1"
   Set rUsuario = dbData.OpenRecordset("SELECT codigo, login FROM Usuario WHERE  (codigo = " & txtCodFuncionario.Text & ");")

   'Recupera um número de arquivo disponível
   f = FreeFile()
      
   'pegar o nome da impressora no ini
   Dim oIni As Ini
   ''Dim var_ImpTermica As String
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
   Set oIni = Nothing
   
   Dim Prt As Printer
   Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   For Each Prt In Printers
      If Prt.DeviceName = var_ImpTermica Then
         Set Printer = Prt
         Exit For
      End If
   Next
   
   With Printer
      .ScaleMode = vbPixels
      .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
      
      For i = 1 To 6
         Printer.Print " "
      Next
      
      .ScaleMode = vbCentimeters
      .FontName = "courier new"
      '.PrintQuality = vbPRPQHigh

      Fonte 10, True, False
      Printer.Print Tab((35 - Len(rEmpresa("fantasia"))) / 2); rEmpresa("fantasia")   'Esse /2 é p/ centralizar
      Fonte 10, False, False
      Printer.Print Tab((35 - Len(rEmpresa("razao"))) / 2); rEmpresa("razao")
      Fonte 8, False, False
      Printer.Print rEmpresa("endereco") & ", " & rEmpresa("cidade") & "-" & rEmpresa("estado")
      Printer.Print "FONE: "; rEmpresa("telefone")                                        '& " - (89) 9986-3739"
      Fonte 8, False, False
      Printer.Print "CNPJ:"; rEmpresa("cnpj") & "  IE:" & rEmpresa("ie")
      Fonte 8, False, False
      Printer.Print String(40, "-")
      
       For i = 1 To 2
         Printer.Print " "
      Next
      
      Fonte 10, True, False
      Printer.Print Tab(10); "R E C I B O"
      
      For i = 1 To 1
         Printer.Print " "
      Next
               
      Fonte 8, True, False
      Printer.Print Tab(28); "R$ " & Format(txtValorHaver.Text, "##,##0.00")
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Dim Line1 As String
      Dim Line2 As String
      
      Dim Texto As String
      Texto = UCase(NumeroExtenso(txtValorHaver.Text, True))
      Line1 = Mid(Texto, 1, 40)
      Line2 = Mid(Texto, 41, 80)
     
      Fonte 8, False, False
      Printer.Print Tab(2); "Recebi(emos) de: "
      Fonte 8, True, False
      Printer.Print Tab(2); rClientes("NOME")
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Fonte 8, False, False
      Printer.Print Tab(2); "A importância supra de: "
      Fonte 8, False, False
      Printer.Print Tab(2); Line1
      Printer.Print Tab(2); Line2
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Fonte 8, False, False
      Printer.Print Tab(2); "Proveniente do: "; "Haver da " & txtNumParcela.Text & "ª parcela "
      Printer.Print Tab(2); "do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Printer.Print Tab(2); "Forma de Pgto: " & cboFormaHaver.Text
      If Not rUsuario.EOF Then
        Printer.Print Tab(2); "Funcionário: "; rUsuario("login")
      Else
        Printer.Print Tab(2); "Funcionário: Não Especificado"
      End If
      'Printer.Print Tab(2); "Funcionário: "; rUsuario("login")

      For i = 1 To 2
         Printer.Print " "
      Next
     
      Fonte 8, False, False
      Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(mskDataHaver) & " de " & MonthName(Month(mskDataHaver)) & " de " & Year(mskDataHaver)
      
      For i = 1 To 3
            Printer.Print " "
      Next
      
      Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
      Printer.Print Tab((40 - Len("Assinatura")) / 2); "Assinatura"
      
   Close #f
   .EndDoc
End With

 If Not rEmpresa Is Nothing Then If rEmpresa.State <> 0 Then rEmpresa.Close
 If Not rPedidos Is Nothing Then If rPedidos.State <> 0 Then rPedidos.Close
 If Not rClientes Is Nothing Then If rClientes.State <> 0 Then rClientes.Close

Tratar_Erro:
   If Err.Number = 52 Then
     ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
     Printer.KillDoc
     Exit Sub
   End If
End Sub

Private Sub Reimprimir_ReciboCupom()
Dim vValorTotal As Currency
Dim vNumParc As Integer
Dim vDataPgto As Date

For f = 0 To Grid_Historico.rows - 1
   Grid_Historico.Row = f
   Grid_Historico.Col = 0
   
   If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then
      vCodPed = Grid_Historico.TextMatrix(Grid_Historico.Row, 3)
      vValorTotal = Grid_Historico.TextMatrix(Grid_Historico.Row, 11)
      vNumParc = Grid_Historico.TextMatrix(Grid_Historico.Row, 4)
      vDataPgto = Grid_Historico.TextMatrix(Grid_Historico.Row, 12)
   End If
Next
   
   'tabela empresa
    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set rEmpresa = dbData.OpenRecordset(sSQL)
    vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")

    Set rPedidos = dbData.OpenRecordset("SELECT cod_pedido, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & vCodPed & ");")
    Set rParcelas = dbData.OpenRecordset("SELECT parcelas.FORMA_PGTO FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO WHERE  (pedidos.cod_pedido = " & vCodPed & ");")
    Set rClientes = dbData.OpenRecordset("SELECT cliente.CODIGO, cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO WHERE  (pedidos.cod_pedido = " & vCodPed & ");")
    Set rUsuario = dbData.OpenRecordset("SELECT Usuario.Codigo, Usuario.Login FROM pedidos INNER JOIN Usuario ON pedidos.COD_FUNCIONARIO = Usuario.Codigo WHERE  (pedidos.cod_pedido = " & vCodPed & ");")
  
  'Recupera um número de arquivo disponível
  f = FreeFile()

 'pegar o nome da impressora no ini
  Dim oIni As Ini
  
  Set oIni = New Ini
  oIni.Arquivo = appPathApp & "config.ini"
  var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
  Set oIni = Nothing
  
  Dim Prt As Printer
  Dim oldPrinter As String
  
  'Armazena o nome da impressora atual
  oldPrinter = Printer.DeviceName
  
  ' Find and use the printer just selected in the ListBox
  For Each Prt In Printers
     If Prt.DeviceName = var_ImpTermica Then
        Set Printer = Prt
        Exit For
     End If
  Next

     With Printer
        .ScaleMode = vbPixels
        .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
        
        For i = 1 To 6
           Printer.Print " "
        Next
        
        .ScaleMode = vbCentimeters
        .FontName = "courier new"
        '.PrintQuality = vbPRPQHigh
        

        Fonte 10, True, False
        Printer.Print Tab((35 - Len(rEmpresa("fantasia"))) / 2); rEmpresa("fantasia")   'Esse /2 é p/ centralizar
        Fonte 10, False, False
        Printer.Print Tab((35 - Len(rEmpresa("razao"))) / 2); rEmpresa("razao")
        Fonte 8, False, False
        Printer.Print rEmpresa("endereco") & ", " & rEmpresa("cidade") & "-" & rEmpresa("estado")
        Printer.Print "FONE: "; rEmpresa("telefone")                                        '& " - (89) 9986-3739"
        Fonte 8, False, False
        Printer.Print "CNPJ:"; rEmpresa("cnpj") & "  IE:" & rEmpresa("ie")
        Fonte 8, False, False
        Printer.Print String(40, "-")
        
         For i = 1 To 2
           Printer.Print " "
        Next
        
        Fonte 10, True, False
        Printer.Print Tab(10); "R E C I B O"
        
        
        For i = 1 To 1
           Printer.Print " "
        Next
                 
        Fonte 8, True, False
        'If vQuitarUma = True Then
           Printer.Print Tab(28); "R$ " & Format(vValorTotal, "##,##0.00")
        'Else
        '   Printer.Print Tab(28); "R$ " & Format(lblTotalSel.Caption, "##,##0.00")
        'End If
        
        For i = 1 To 1
           Printer.Print " "
        Next
        
        Dim Line1 As String
        Dim Line2 As String
        
        Dim Texto As String
        'If vQuitarUma = True Then
           Texto = UCase(NumeroExtenso(vValorTotal, True))
        'Else
           'Texto = UCase(NumeroExtenso(lblTotalSel.Caption, True))
        'End If
        Line1 = Mid(Texto, 1, 40)
        Line2 = Mid(Texto, 41, 80)
       
        Fonte 8, False, False
        Printer.Print Tab(2); "Recebi(emos) de: "
        Fonte 8, True, False
        Printer.Print Tab(2); rClientes("NOME")
        
        For i = 1 To 1
           Printer.Print " "
        Next
        
        Fonte 8, False, False
        Printer.Print Tab(2); "A importância supra de: "
        Fonte 8, False, False
        Printer.Print Tab(2); Line1
        Printer.Print Tab(2); Line2
        Fonte 8, False, False
       
        For i = 1 To 1
           Printer.Print " "
        Next
        
       'If vQuitarUma = True Then
           Printer.Print Tab(2); "Proveniente do: "; "Pagamento da " & vNumParc & "ª parcela "
           Printer.Print Tab(2); "do PEDIDO Nº " & Format(vCodPed, "000000")
       'Else
       'DIVIDINDO OS PEDIDOS POR LINHA
         '  Dim iMaxCol As Integer
         '  Dim var_Parc As String
         '  Dim strSeparator As String
  
  
         '  iMaxCol = 50
         '  strSeparator = ", "
         '  'var_Parc = "00001/00, 00001/01, 00001/02, 00001/03, 00001/04, 00001/05, 00001/06, 00001/07, 00001/08, 00001/09, 00001/10, 00001/11, 00001/12, 00001/13, 00001/14, 00001/15"
        
        ''INICIO DA IMPRESSÃO DE MULTIPLUS PEDIDOS
          '  'Dim var_Parc As String
          '  Dim y As Integer
   
         '   var_Parc = ""
            
         '   With Grid_Parcelas
         '      For y = 1 To .Rows - 1
         '         .Col = 0
         '         .Row = y
   
         '         If .CellPicture = ImgMarcada Then
         '            If y = 1 Then
         '               var_Parc = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
         '            ElseIf Right(var_Parc, 8) = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00") Then
         '               MsgBox "Tratar Repetido"
         '            Else
         '               var_Parc = var_Parc & ", " & Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
         '               If y = .Rows - 1 Then Exit For
         '            End If
         '         End If
         '      Next y
         '   End With
         '   'txtText1.Text = ReturnLargueToPrint(var_Parc, iMaxCol, strSeparator)
         '   Printer.Print Tab(2); "Proveniente do Pagamento dos pedidos: "
         '   Printer.Print Tab(2); ReturnLargueToPrint(var_Parc, iMaxCol, strSeparator)
        'End If
        
           'Printer.Print Tab(2); "Proveniente do: "; "Pagamento da " & txtNumParcela.Text & "ª parcela "
           

        
        '.txtProveniente.Caption = "PEDIDO(S): " & var_Parc
        
        '.txtValor.Caption = UCase(NumeroExtenso(lblTotalSel.Caption, True))
        '.txthead.Caption = "R$ " & Format(lblTotalSel.Caption, "##,##0.00")
        
        '.txtData.Caption = "" & vCidadeUF & ", " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
        '===================================================

        
        For i = 1 To 1
           Printer.Print " "
        Next
        
        
        Printer.Print Tab(2); "Forma de Pgto: " & rParcelas("FORMA_PGTO")
       If Not rUsuario.EOF Then
           Printer.Print Tab(2); "Funcionário: "; rUsuario("login");
       Else
           Printer.Print Tab(2); "Funcionário: Não Especificado"
       End If
        
        For i = 1 To 2
           Printer.Print " "
        Next
       
        Fonte 8, False, False
        'If vQuitarUma = True Then
           Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(vDataPgto) & " de " & MonthName(Month(vDataPgto)) & " de " & Year(vDataPgto)
       'Else
        '   Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
       'End If
        
        For i = 1 To 3
              Printer.Print " "
        Next
        
        Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
        Printer.Print Tab((40 - Len("Assinatura")) / 2); "Assinatura"
        

       
     Close #f
     .EndDoc
     'rsPedidos.Close
     'rsFunc.Close
     'RS.Close
     'BD.Close
  End With
   
Tratar_Erro:
   ' Atribui a impressora inicial
   'For Each Prt In Printers
   '   If Prt.DeviceName = oldPrinter Then
   '      Set Printer = Prt
   '      Exit For
   '   End If
   'Next
   
   If Not rEmpresa Is Nothing Then If rEmpresa.State <> 0 Then rEmpresa.Close
   If Not rPedidos Is Nothing Then If rPedidos.State <> 0 Then rPedidos.Close
   If Not rClientes Is Nothing Then If rClientes.State <> 0 Then rClientes.Close
   If Not rParcelas Is Nothing Then If rParcelas.State <> 0 Then rParcelas.Close
   
   'If Err.Number = 52 Then
    '  ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
    '  Printer.KillDoc
    '  Exit Sub
   'End If
End Sub
Private Sub Imprimir_ReciboCupomSel()
   'On Error GoTo Tratar_Erro
   Dim i As Integer
   Dim f As Integer
    
    'tabela empresa
    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set rEmpresa = dbData.OpenRecordset(sSQL)
    vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")
    
   If vQuitarUma = True Then
        If txtCodPedido.Text = "" Then Exit Sub
        'tabela pedidos
        Set rPedidos = dbData.OpenRecordset("SELECT cod_pedido, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");")
   End If
    
    'tabela pedidos
    Set rClientes = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE  (codigo = " & txtCodCliente.Text & ");")
   
   If txtCodFuncionario.Text = "" Then txtCodFuncionario.Text = "1"
   Set rUsuario = dbData.OpenRecordset("SELECT codigo, login FROM Usuario WHERE  (codigo = " & txtCodFuncionario.Text & ");")

   
   'Recupera um número de arquivo disponível
   f = FreeFile()
      
   'Dim Prt As Printer
   'Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   'oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   'For Each Prt In Printers
   '   If Prt.DeviceName = var_ImpTermica Then
   '      Set Printer = Prt
   '      Exit For
    '  End If
  ' Next
  
  'pegar o nome da impressora no ini
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
   Set oIni = Nothing
   
   Dim Prt As Printer
   Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   For Each Prt In Printers
      If Prt.DeviceName = var_ImpTermica Then
         Set Printer = Prt
         Exit For
      End If
   Next
 
      With Printer
         .ScaleMode = vbPixels
         .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
         
         For i = 1 To 6
            Printer.Print " "
         Next
         
         .ScaleMode = vbCentimeters
         .FontName = "courier new"
         '.PrintQuality = vbPRPQHigh
         

         Fonte 10, True, False
         Printer.Print Tab((35 - Len(rEmpresa("fantasia"))) / 2); rEmpresa("fantasia")   'Esse /2 é p/ centralizar
         Fonte 10, False, False
         Printer.Print Tab((35 - Len(rEmpresa("razao"))) / 2); rEmpresa("razao")
         Fonte 8, False, False
         Printer.Print rEmpresa("endereco") & ", " & rEmpresa("cidade") & "-" & rEmpresa("estado")
         Printer.Print "FONE: "; rEmpresa("telefone")                                        '& " - (89) 9986-3739"
         Fonte 8, False, False
         Printer.Print "CNPJ:"; rEmpresa("cnpj") & "  IE:" & rEmpresa("ie")
         Fonte 8, False, False
         Printer.Print String(40, "-")
         
          For i = 1 To 2
            Printer.Print " "
         Next
         
         Fonte 10, True, False
         Printer.Print Tab(10); "R E C I B O"
         
         
         For i = 1 To 1
            Printer.Print " "
         Next
                  
         Fonte 8, True, False
         If vQuitarUma = True Then
            Printer.Print Tab(28); "R$ " & Format(txtTotal.Text, "##,##0.00")
         Else
            Printer.Print Tab(28); "R$ " & Format(lblTotalSelecionadasQuitadas.Caption, "##,##0.00")
         End If
         
         For i = 1 To 1
            Printer.Print " "
         Next
         
         Dim Line1 As String
         Dim Line2 As String
         
         Dim Texto As String
         If vQuitarUma = True Then
            Texto = UCase(NumeroExtenso(txtTotal.Text, True))
         Else
            Texto = UCase(NumeroExtenso(lblTotalSelecionadasQuitadas.Caption, True))
         End If
         Line1 = Mid(Texto, 1, 40)
         Line2 = Mid(Texto, 41, 80)
        
         Fonte 8, False, False
         Printer.Print Tab(2); "Recebi(emos) de: "
         Fonte 8, True, False
         Printer.Print Tab(2); rClientes("NOME")
         
         For i = 1 To 1
            Printer.Print " "
         Next
         
         Fonte 8, False, False
         Printer.Print Tab(2); "A importância supra de: "
         Fonte 8, False, False
         Printer.Print Tab(2); Line1
         Printer.Print Tab(2); Line2
         Fonte 8, False, False
        
         For i = 1 To 1
            Printer.Print " "
         Next
         
        If vQuitarUma = True Then
            Printer.Print Tab(2); "Proveniente do: "; "Pagamento da " & txtNumParcela.Text & "ª parcela "
            Printer.Print Tab(2); "do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
        Else
        'DIVIDINDO OS PEDIDOS POR LINHA
            Dim iMaxCol As Integer
            Dim var_Parc As String
            Dim strSeparator As String
   
   
            iMaxCol = 50
            strSeparator = ", "
            'var_Parc = "00001/00, 00001/01, 00001/02, 00001/03, 00001/04, 00001/05, 00001/06, 00001/07, 00001/08, 00001/09, 00001/10, 00001/11, 00001/12, 00001/13, 00001/14, 00001/15"
         
         'INICIO DA IMPRESSÃO DE MULTIPLUS PEDIDOS
             'Dim var_Parc As String
             Dim y As Integer
    
             var_Parc = ""
             
             With Grid_Historico
                For y = 1 To .rows - 1
                   .Col = 0
                   .Row = y
    
                   If .CellPicture = ImgMarcadaPAGAS Then
                      If y = 1 Then
                         var_Parc = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
                      ElseIf Right(var_Parc, 8) = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00") Then
                         MsgBox "Tratar Repetido"
                      Else
                         var_Parc = var_Parc & ", " & Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
                         If y = .rows - 1 Then Exit For
                      End If
                   End If
                Next y
             End With
             'txtText1.Text = ReturnLargueToPrint(var_Parc, iMaxCol, strSeparator)
             Printer.Print Tab(2); "Proveniente do Pagamento dos pedidos: "
             Printer.Print Tab(2); ReturnLargueToPrint(var_Parc, iMaxCol, strSeparator)
         End If
         
            'Printer.Print Tab(2); "Proveniente do: "; "Pagamento da " & txtNumParcela.Text & "ª parcela "
            

         
         '.txtProveniente.Caption = "PEDIDO(S): " & var_Parc
         
         '.txtValor.Caption = UCase(NumeroExtenso(lblTotalSel.Caption, True))
         '.txthead.Caption = "R$ " & Format(lblTotalSel.Caption, "##,##0.00")
         
         '.txtData.Caption = "" & vCidadeUF & ", " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
         '===================================================

         
         For i = 1 To 1
            Printer.Print " "
         Next
         
         
         Printer.Print Tab(2); "Forma de Pgto: " & cboForma.Text
        If Not rUsuario.EOF Then
            Printer.Print Tab(2); "Funcionário: "; rUsuario("login")
        Else
            Printer.Print Tab(2); "Funcionário: Não Especificado"
        End If
         
         For i = 1 To 2
            Printer.Print " "
         Next
        
         Fonte 8, False, False
         If vQuitarUma = True Then
            Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(mskPagamento) & " de " & MonthName(Month(mskPagamento)) & " de " & Year(mskPagamento)
        Else
            Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
        End If
         
         For i = 1 To 3
               Printer.Print " "
         Next
         
         Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
         Printer.Print Tab((40 - Len("Assinatura")) / 2); "Assinatura"
         

        
      Close #f
      .EndDoc
      'rsPedidos.Close
      'rsFunc.Close
      'RS.Close
      'BD.Close
   End With
   
Tratar_Erro:
   ' Atribui a impressora inicial
   'For Each Prt In Printers
   '   If Prt.DeviceName = oldPrinter Then
   '      Set Printer = Prt
   '      Exit For
   '   End If
   'Next
   
   If Not rEmpresa Is Nothing Then If rEmpresa.State <> 0 Then rEmpresa.Close
   If Not rPedidos Is Nothing Then If rPedidos.State <> 0 Then rPedidos.Close
   If Not rClientes Is Nothing Then If rClientes.State <> 0 Then rClientes.Close
   
   'If Err.Number = 52 Then
    '  ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
    '  Printer.KillDoc
    '  Exit Sub
   'End If
End Sub
Private Sub Imprimir_ReciboCupom()
   'On Error GoTo Tratar_Erro
   Dim i As Integer
   Dim f As Integer
    
    'tabela empresa
    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set rEmpresa = dbData.OpenRecordset(sSQL)
    vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")
    
   If vQuitarUma = True Then
        If txtCodPedido.Text = "" Then Exit Sub
        'tabela pedidos
        Set rPedidos = dbData.OpenRecordset("SELECT cod_pedido, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");")
   End If
    
    'tabela pedidos
    Set rClientes = dbData.OpenRecordset("SELECT codigo, nome FROM cliente WHERE  (codigo = " & txtCodCliente.Text & ");")
   
   If txtCodFuncionario.Text = "" Then txtCodFuncionario.Text = "1"
   Set rUsuario = dbData.OpenRecordset("SELECT codigo, login FROM Usuario WHERE  (codigo = " & txtCodFuncionario.Text & ");")

   
   'Recupera um número de arquivo disponível
   f = FreeFile()
      
   'Dim Prt As Printer
   'Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   'oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   'For Each Prt In Printers
   '   If Prt.DeviceName = var_ImpTermica Then
   '      Set Printer = Prt
   '      Exit For
    '  End If
  ' Next
  
  'pegar o nome da impressora no ini
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
   Set oIni = Nothing
   
   Dim Prt As Printer
   Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   For Each Prt In Printers
      If Prt.DeviceName = var_ImpTermica Then
         Set Printer = Prt
         Exit For
      End If
   Next
 
      With Printer
         .ScaleMode = vbPixels
         .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
         
         For i = 1 To 6
            Printer.Print " "
         Next
         
         .ScaleMode = vbCentimeters
         .FontName = "courier new"
         '.PrintQuality = vbPRPQHigh
         

         Fonte 10, True, False
         Printer.Print Tab((35 - Len(rEmpresa("fantasia"))) / 2); rEmpresa("fantasia")   'Esse /2 é p/ centralizar
         Fonte 10, False, False
         Printer.Print Tab((35 - Len(rEmpresa("razao"))) / 2); rEmpresa("razao")
         Fonte 8, False, False
         Printer.Print rEmpresa("endereco") & ", " & rEmpresa("cidade") & "-" & rEmpresa("estado")
         Printer.Print "FONE: "; rEmpresa("telefone")                                        '& " - (89) 9986-3739"
         Fonte 8, False, False
         Printer.Print "CNPJ:"; rEmpresa("cnpj") & "  IE:" & rEmpresa("ie")
         Fonte 8, False, False
         Printer.Print String(40, "-")
         
          For i = 1 To 2
            Printer.Print " "
         Next
         
         Fonte 10, True, False
         Printer.Print Tab(10); "R E C I B O"
         
         
         For i = 1 To 1
            Printer.Print " "
         Next
                  
         Fonte 8, True, False
         If vQuitarUma = True Then
            Printer.Print Tab(28); "R$ " & Format(txtTotal.Text, "##,##0.00")
         Else
            Printer.Print Tab(28); "R$ " & Format(lblTotalSel.Caption, "##,##0.00")
         End If
         
         For i = 1 To 1
            Printer.Print " "
         Next
         
         Dim Line1 As String
         Dim Line2 As String
         
         Dim Texto As String
         If vQuitarUma = True Then
            Texto = UCase(NumeroExtenso(txtTotal.Text, True))
         Else
            Texto = UCase(NumeroExtenso(lblTotalSel.Caption, True))
         End If
         Line1 = Mid(Texto, 1, 40)
         Line2 = Mid(Texto, 41, 80)
        
         Fonte 8, False, False
         Printer.Print Tab(2); "Recebi(emos) de: "
         Fonte 8, True, False
         Printer.Print Tab(2); rClientes("NOME")
         
         For i = 1 To 1
            Printer.Print " "
         Next
         
         Fonte 8, False, False
         Printer.Print Tab(2); "A importância supra de: "
         Fonte 8, False, False
         Printer.Print Tab(2); Line1
         Printer.Print Tab(2); Line2
         Fonte 8, False, False
        
         For i = 1 To 1
            Printer.Print " "
         Next
         
        If vQuitarUma = True Then
            Printer.Print Tab(2); "Proveniente do: "; "Pagamento da " & txtNumParcela.Text & "ª parcela "
            Printer.Print Tab(2); "do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
        Else
        'DIVIDINDO OS PEDIDOS POR LINHA
            Dim iMaxCol As Integer
            Dim var_Parc As String
            Dim strSeparator As String
   
   
            iMaxCol = 50
            strSeparator = ", "
            'var_Parc = "00001/00, 00001/01, 00001/02, 00001/03, 00001/04, 00001/05, 00001/06, 00001/07, 00001/08, 00001/09, 00001/10, 00001/11, 00001/12, 00001/13, 00001/14, 00001/15"
         
         'INICIO DA IMPRESSÃO DE MULTIPLUS PEDIDOS
             'Dim var_Parc As String
             Dim y As Integer
    
             var_Parc = ""
             
             With Grid_Parcelas
                For y = 1 To .rows - 1
                   .Col = 0
                   .Row = y
    
                   If .CellPicture = ImgMarcada Then
                      If y = 1 Then
                         var_Parc = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
                      ElseIf Right(var_Parc, 8) = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00") Then
                         MsgBox "Tratar Repetido"
                      Else
                         var_Parc = var_Parc & ", " & Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
                         If y = .rows - 1 Then Exit For
                      End If
                   End If
                Next y
             End With
             'txtText1.Text = ReturnLargueToPrint(var_Parc, iMaxCol, strSeparator)
             Printer.Print Tab(2); "Proveniente do Pagamento dos pedidos: "
             Printer.Print Tab(2); ReturnLargueToPrint(var_Parc, iMaxCol, strSeparator)
         End If
         
            'Printer.Print Tab(2); "Proveniente do: "; "Pagamento da " & txtNumParcela.Text & "ª parcela "
            

         
         '.txtProveniente.Caption = "PEDIDO(S): " & var_Parc
         
         '.txtValor.Caption = UCase(NumeroExtenso(lblTotalSel.Caption, True))
         '.txthead.Caption = "R$ " & Format(lblTotalSel.Caption, "##,##0.00")
         
         '.txtData.Caption = "" & vCidadeUF & ", " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
         '===================================================

         
         For i = 1 To 1
            Printer.Print " "
         Next
         
         
         Printer.Print Tab(2); "Forma de Pgto: " & cboForma.Text
        If Not rUsuario.EOF Then
            Printer.Print Tab(2); "Funcionário: "; rUsuario("login")
        Else
            Printer.Print Tab(2); "Funcionário: Não Especificado"
        End If
         
         For i = 1 To 2
            Printer.Print " "
         Next
        
         Fonte 8, False, False
         If vQuitarUma = True Then
            Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(mskPagamento) & " de " & MonthName(Month(mskPagamento)) & " de " & Year(mskPagamento)
        Else
            Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
        End If
         
         For i = 1 To 3
               Printer.Print " "
         Next
         
         Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
         Printer.Print Tab((40 - Len("Assinatura")) / 2); "Assinatura"
         

        
      Close #f
      .EndDoc
      'rsPedidos.Close
      'rsFunc.Close
      'RS.Close
      'BD.Close
   End With
   
Tratar_Erro:
   ' Atribui a impressora inicial
   'For Each Prt In Printers
   '   If Prt.DeviceName = oldPrinter Then
   '      Set Printer = Prt
   '      Exit For
   '   End If
   'Next
   
   If Not rEmpresa Is Nothing Then If rEmpresa.State <> 0 Then rEmpresa.Close
   If Not rPedidos Is Nothing Then If rPedidos.State <> 0 Then rPedidos.Close
   If Not rClientes Is Nothing Then If rClientes.State <> 0 Then rClientes.Close
   
   'If Err.Number = 52 Then
    '  ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
    '  Printer.KillDoc
    '  Exit Sub
   'End If
End Sub


Private Sub Fonte(Tamanho As Byte, Negrito As Boolean, Italico As Boolean) 'Altera a fonte
   Printer.FontSize = Tamanho
   Printer.FontBold = Negrito
   Printer.FontItalic = Italico
End Sub

Private Sub cmdMarcarCheck_Click()
If cmdMarcarCheck.Caption = "MARCAR TODAS" Then
   OP = MarcarTodos
   AcaoGrid
   cmdQuitarTodas.Visible = True
   cmdCancelar.Visible = True
   cmdMarcarCheck.Caption = "DESMARCAR TODAS"
Else
   OP = DesmarcarTodos
   AcaoGrid
   cmdMarcarCheck.Caption = "MARCAR TODAS"
   cmdQuitarTodas.Visible = False
   cmdCancelar.Visible = False
End If

OP = contar
AcaoGrid
Somar_Parcelas_Selecionadas
End Sub

Private Sub cmdMarcarTodasREATIVAR_Click()
   If cmdMarcarTodasREATIVAR.Caption = "MARCAR TODAS" Then
      OP = MarcarTodos
      AcaoGridREATIVAR
      cmdMarcarTodasREATIVAR.Caption = "DESMARCAR TODAS"
   Else
      OP = DesmarcarTodos
      AcaoGridREATIVAR
      cmdMarcarTodasREATIVAR.Caption = "MARCAR TODAS"
   End If
   
   OP = contar
   AcaoGridREATIVAR
End Sub

Private Sub cmdMostrarProdutos_Click()
Dim f As Integer

For f = 0 To Grid_Parcelas.rows - 1
   Grid_Parcelas.Row = f
   Grid_Parcelas.Col = 0
   
   If Grid_Parcelas.CellPicture = ImgMarcada Then
      Parcelas_Consulta_Produtos.loadPedidos Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 3), Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 2)
      Parcelas_Consulta_Produtos.Show 1
   End If
Next
End Sub

Private Sub cmdMostrarProdutosREATIVAR_Click()
'Dim f As Integer
   
For f = 0 To Grid_Historico.rows - 1
   Grid_Historico.Row = f
   Grid_Historico.Col = 0
   
   If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then
      Parcelas_Consulta_Produtos.loadPedidos Grid_Historico.TextMatrix(Grid_Historico.Row, 3), ""
      Parcelas_Consulta_Produtos.Show 1
   End If
Next
End Sub

Private Sub cmdQuitarTodas_Click()
ConsultarCaixaAtual

If varCodCaixa = 0 Then
    MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim f As Integer
Dim sSQL As String

If cboForma.Text = "" Then
   ShowMsg "Faltou Escolher a forma de pagamento!", vbInformation
   cboForma.SetFocus
   Exit Sub
End If

Dim varTipoCartao As String
varTipoCartao = "NULL"

If cboForma.Text = "CARTÃO - DÉBITO" Then
   varTipoCartao = "'D'"
ElseIf cboForma.Text = "CARTÃO - CRÉDITO" Then
   varTipoCartao = "'C'"
Else
    varTipoCartao = "NULL"
End If

Dim var_PAGAMENTO As String
If cboForma.Text = "DINHEIRO" Then
   var_PAGAMENTO = "DINHEIRO"
ElseIf cboForma.Text = "CARTÃO - DÉBITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboForma.Text = "CARTÃO - CRÉDITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboForma.Text = "CHEQUE" Then
   var_PAGAMENTO = "CHEQUE"
ElseIf cboForma.Text = "TRANSFERÊNCIA" Then
   var_PAGAMENTO = "TRANSFERENCIA"
ElseIf cboForma.Text = "DEPOSITO" Then
   var_PAGAMENTO = "DEPOSITO"
ElseIf cboForma.Text = "BOLETO" Then
   var_PAGAMENTO = "BOLETO"
ElseIf cboForma.Text = "PIX" Then
   var_PAGAMENTO = "PIX"
End If

'MOSTRAR SE O CAIXA ESTÁ FECHADO
Verificar_Caixa_Baixa
If CAIXA_FECHADO_BAIXA = True Then Exit Sub

With Grid_Parcelas
   For f = 1 To .rows - 1
      .Col = 0
      .Row = f
      If .CellPicture = ImgMarcada Then
         sSQL = "UPDATE parcelas SET status = 1, valor_final = " & Replace(CCur(.TextMatrix(.Row, 11)), ",", ".") & ", desconto = 0, tipo = 'PARCELA', tipo_cartao = " & varTipoCartao & ", forma_pgto = '" & var_PAGAMENTO & "', COD_FUNCIONARIO = " & txtCodFuncionario & ", " & _
            "pagamento = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), hora = '" & Format(Now, ocHRMN) & "', " & _
            "juros = " & Replace(CCur(.TextMatrix(.Row, 8)), ",", ".") & ", dias_atrazo = " & .TextMatrix(.Row, 7) & ", " & _
            "caixa = '" & StatusBar1.Panels(2).Text & "', CODCAIXA = " & varCodCaixa & " " & _
            "WHERE (cod_pedido = " & .TextMatrix(.Row, 3) & ") AND (numero = " & .TextMatrix(.Row, 4) & ");"
         
         dbData.Execute sSQL
      End If
   Next f
End With

If varPgtoAutomatico = False Then
    If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
       IMPRIMIR = True
    Else
       IMPRIMIR = False
    End If
    
    If IMPRIMIR = True Then
         'If vQuitarUma = True Then
             If varTipoRecPgto = "CUPOM" Then
                 Imprimir_ReciboCupom
             Else
                 ImprimirPedidoFolha
             End If
         'Else
         '    ImprimirPedidoFolha
         'End If
    End If
        
    MostrarGrid_Haver
    MostrarGrid_Parcelas
    MostrarGrid_Historico
    OP = contar
    AcaoGrid
    Somar_Parcelas_Selecionadas
    cboForma.Text = ""
Else
    IMPRIMIR = True
    If IMPRIMIR = True Then
         'If vQuitarUma = True Then
             If varTipoRecPgto = "CUPOM" Then
                 Imprimir_ReciboCupom
             Else
                 ImprimirPedidoFolha
             End If
         'Else
         '    ImprimirPedidoFolha
         'End If
    End If
End If
End Sub

Private Sub cmdReativar_Click()
'verificar ser o caixa do haver selecionado está em aberto
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (codcaixa = " & txtCodCaixa.Text & ") and (caixa = '" & txtCaixa.Text & "') and caixa_dia.status = 1;"
Set r = dbData.OpenRecordset(sSQL)
Debug.Print sSQL

'If Not r.BOF Then
    If r.RecordCount > 0 Then
        MsgBox "O caixa onde essa parcela foi adicionado encontra-se fechado!", vbInformation, "Aviso do Sistema"
        r.Close
        Set r = Nothing
        Exit Sub
    End If
'Else

'End If

'INICIO DA ROTINA
If cboCliente.Text = "" Or txtConNumParcela.Text = "" Or txtCONValor.Text = "" Then Exit Sub
If txtConCodPedido.Text = "" And txtConCodOS.Text = "" Then Exit Sub

'If txtConCodOS.Text <> "" And txtConCodPedido.Text = "" Then
'   If MsgBox("Deseja reativar a parcela '" & txtConNumParcela.Text & "' da OS No.'" & txtConCodOS.Text & "' ??", vbInformation + vbYesNo, "Aviso do Sistema") = vbNo Then Exit Sub
'ElseIf txtConCodOS.Text = "" And txtConCodPedido.Text <> "" Then
   If ShowMsg("Deseja reativar a parcela '" & txtConNumParcela.Text & "' do Pedido No. '" & txtConCodPedido.Text & "' ??", vbInformation + vbYesNo) = vbNo Then Exit Sub
'End If

'If txtConCodOS.Text <> "" And txtConCodPedido.Text = "" Then
'   execSQL "UPDATE PARCELAS SET STATUS = false , VALOR_FINAL = '" & Format(0, "##,##0.00") & "', PAGAMENTO = #" & Format(MaskEdBox1.Text, "mm/dd/yyyy") & "# WHERE COD_OS = " & txtConCodOS.Text & " AND NUMERO = " & txtConNumParcela.Text & ""
'   execSQL "UPDATE CHEQUE SET STATUS = false WHERE COD_OS = " & txtConCodOS.Text & " AND PARCELA = " & txtConNumParcela.Text & ""
'   execSQL "UPDATE PARCELAS_HAVER SET STATUS = 0 WHERE (CODIGO = " & txtCONcodParc.Text & ")"
'ElseIf txtConCodOS.Text = "" And txtConCodPedido.Text <> "" Then
   dbData.Execute "UPDATE parcelas SET status = 0, valor_final = VALOR, JUROS = 0, DIAS_ATRAZO = 0, MULTA = 0, DESCONTO = 0, pagamento = Null, forma_pgto = '', caixa = '', codcaixa = '' WHERE (cod_pedido = " & txtConCodPedido.Text & ") AND (numero = " & txtConNumParcela.Text & ");"
   'dbData.Execute "UPDATE cheque SET status = 0 WHERE (cod_pedido = " & txtConCodPedido.Text & ") AND (parcela = " & txtConNumParcela.Text & ");"
   dbData.Execute "UPDATE parcelas_haver SET status = 0 WHERE (codigo = " & txtCONcodParc.Text & ");"
'End If

MostrarGrid_Parcelas
Calcular_Valores
MostrarGrid_Historico
LimparObjetos_Historico

OP = contar
AcaoGridREATIVAR
End Sub

Private Sub cmdRemoverHaver_Click()
On Error GoTo erro

'verificar ser o caixa do haver selecionado está em aberto
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (codcaixa = " & Grid_Haver.TextMatrix(Grid_Haver.Row, 2) & ") and (caixa = '" & Grid_Haver.TextMatrix(Grid_Haver.Row, 3) & "') and caixa_dia.status = 1;"
Set r = dbData.OpenRecordset(sSQL)

If r.RecordCount > 0 Then
    MsgBox "O caixa onde esse haver foi adicionado encontra-se fechado!", vbInformation, "Aviso do Sistema"
    r.Close
    Set r = Nothing
    Exit Sub
End If

'INICIO DA ROTINA
Verificar_Caixa_Haver
If CAIXA_FECHADO_HAVER = True Then Exit Sub

If Grid_Haver.TextMatrix(Grid_Haver.Row, 1) = "" Then GoSub erro
If ShowMsg("Deseja remover o haver de: " & Grid_Haver.TextMatrix(Grid_Haver.Row, 3) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

'If txtOrigem.Text <> "" And txtCodPedido.Text = "" Then
'   execSQL "DELETE FROM PARCELAS_HAVER WHERE CODIGO =" & Grid_Haver.TextMatrix(Grid_Haver.Row, 1) & ""
'   execSQL "DELETE FROM CAIXA_ENTRADA WHERE COD_HAVER = " & Grid_Haver.TextMatrix(Grid_Haver.Row, 1) & ""
'ElseIf txtOrigem.Text = "" And txtCodPedido.Text <> "" Then
   dbData.Execute "DELETE FROM parcelas_haver WHERE (codigo = " & Grid_Haver.TextMatrix(Grid_Haver.Row, 1) & ");"
   'dbData.Execute "DELETE FROM caixa_entrada WHERE (cod_haver = " & Grid_Haver.TextMatrix(Grid_Haver.Row, 1) & ");"
'End If

'se nao tiver nenhum haver, ele desmarca o campo HAVER
sSQL = "SELECT * FROM parcelas_haver WHERE (cod_parcela = " & txtCodParc.Text & ") AND (status = 0);"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   dbData.Execute "UPDATE parcelas SET haver = 0 WHERE (codigo = " & txtCodParc.Text & ");"
End If

If r.State <> 0 Then r.Close
Set r = Nothing

MostrarGrid_Parcelas
Calcular_Valores
MostrarGrid_Haver
LimparObjetos_Parcelas_Haver
mskDataHaver.Text = Format(Date, "dd/mm/yy")
Somar_Parcelas_Selecionadas
txtValorHaver.SetFocus
Exit Sub

erro:
   MsgBox "Não existe nenhum haver selecionado para ser removido!", vbExclamation
   Exit Sub
End Sub

Private Sub cmdSalvar_Click()
If txtDias.Visible = True Then
    ConsultarCaixaAtual
    
    If varCodCaixa = 0 Then
        MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
    
    If cboCliente.Text = "" And txtNumParcela.Text = "" And txtValor.Text = "" Then Exit Sub
    If txtCodPedido.Text = "" And txtOrigem.Text = "" Then Exit Sub
    
    If cboForma.Text = "" Then
       ShowMsg "Faltou Escolher a forma de pagamento!", vbInformation
       cboForma.SetFocus
       Exit Sub
    End If
    
    Dim varTipoCartao As String
    varTipoCartao = "NULL"
    
    If cboForma.Text = "CARTÃO - DÉBITO" Then
       varTipoCartao = "'D'"
    ElseIf cboForma.Text = "CARTÃO - CRÉDITO" Then
       varTipoCartao = "'C'"
    Else
        varTipoCartao = "NULL"
    End If
    
    Dim var_PAGAMENTO As String
    If cboForma.Text = "DINHEIRO" Then
       var_PAGAMENTO = "DINHEIRO"
    ElseIf cboForma.Text = "CARTÃO - DÉBITO" Then
       var_PAGAMENTO = "CARTAO"
    ElseIf cboForma.Text = "CARTÃO - CRÉDITO" Then
       var_PAGAMENTO = "CARTAO"
    ElseIf cboForma.Text = "CHEQUE" Then
       var_PAGAMENTO = "CHEQUE"
    ElseIf cboForma.Text = "TRANSFERÊNCIA" Then
       var_PAGAMENTO = "TRANSFERENCIA"
    ElseIf cboForma.Text = "DEPOSITO" Then
       var_PAGAMENTO = "DEPOSITO"
    ElseIf cboForma.Text = "BOLETO" Then
       var_PAGAMENTO = "BOLETO"
    ElseIf cboForma.Text = "PIX" Then
       var_PAGAMENTO = "PIX"
    End If
    
    If txtOrigem.Text <> "" And txtCodPedido.Text = "" Then
       If ShowMsg("Deseja quitar a parcela '" & txtNumParcela.Text & "' da OS No. '" & txtOrigem.Text & "' ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ElseIf txtOrigem.Text = "" And txtCodPedido.Text <> "" Then
       If ShowMsg("Deseja quitar a parcela '" & txtNumParcela.Text & "' do Pedido No. '" & txtCodPedido.Text & "' ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    'If txtOrigem.Text <> "" And txtCodPedido.Text = "" Then
    '   execSQL "UPDATE PARCELAS SET STATUS = 1 , VALOR_FINAL = '" & Format(txtTotal.Text, "##,##0.00") & "' , PAGAMENTO = #" & Format(mskPagamento.Text, "mm/dd/yyyy") & "#, HORA = #" & Format(lblHora.Caption, "hh:mm") & "#, FORMA_PGTO =  '" & cboForma.Text & "' WHERE COD_OS = " & txtOrigem.Text & " AND NUMERO = " & txtNumParcela.Text
    '   execSQL "UPDATE CHEQUE SET STATUS = 1 WHERE COD_OS = " & txtOrigem.Text & " AND PARCELA = " & txtNumParcela.Text
    '   execSQL "UPDATE PARCELAS_HAVER SET STATUS = 1 WHERE (COD_PARCELA = " & txtCodParc.Text & ")"
    'ElseIf txtOrigem.Text = "" And txtCodPedido.Text <> "" Then
    
    'verificar se o pedido foi criado no mesmo dia da parcela dada baixa
    If txtCodPedido.Text = "" Then Exit Sub

    sSQL = "SELECT COD_PEDIDO, DATA_COMPRA FROM pedidos " & _
       "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
    
    Dim vDataCompra As Date
    Dim vDataPgto As Date
    If Not r.BOF Then
        vDataCompra = r("DATA_COMPRA")
        vDataPgto = mskPagamento.Text
    Else
        vDataCompra = ""
        vDataPgto = ""
    End If
    
    If vDataCompra = vDataPgto Then
        dbData.Execute "UPDATE pedidos SET caixa = '" & StatusBar1.Panels(2).Text & "', CODCAIXA = " & varCodCaixa & " WHERE (cod_pedido = " & txtCodPedido.Text & ");"
    End If
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    Dim vOrigem As String
    If txtOrigem.Text = "ALUGUEL" Then
        vOrigem = "ALUGUEL"
    ElseIf txtOrigem.Text = "VENDA" Then
        vOrigem = "PARCELA"
    End If

    If txtItem.Text = "0" Then
        If chkJuros.Value = 1 Then
           dbData.Execute "UPDATE parcelas SET status = 1, valor_final = " & Replace(CCur(txtTotal.Text), ",", ".") & ", DESCONTO = " & Replace(CCur(txtDesconto.Text), ",", ".") & ", pagamento = CONVERT(DATETIME, '" & Format(mskPagamento.Text, ocDATA) & "', 103), hora = '" & Format(lblHora.Caption, ocHRMN) & "', juros = " & Replace(CCur(txtTJuros.Text), ",", ".") & ", dias_atrazo = " & txtDias.Text & ", forma_pgto = '" & var_PAGAMENTO & "', tipo = '" & vOrigem & "', tipo_cartao = " & varTipoCartao & ", caixa = '" & StatusBar1.Panels(2).Text & "', CODCAIXA = " & varCodCaixa & ", COD_FUNCIONARIO =" & txtCodFuncionario & " WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = " & txtNumParcela.Text & ");"
        Else
           dbData.Execute "UPDATE parcelas SET status = 1, valor_final = " & Replace(CCur(txtTotal.Text), ",", ".") & ", DESCONTO = " & Replace(CCur(txtDesconto.Text), ",", ".") & ", pagamento = CONVERT(DATETIME, '" & Format(mskPagamento.Text, ocDATA) & "', 103), hora = '" & Format(lblHora.Caption, ocHRMN) & "', juros = 0, dias_atrazo = 0, forma_pgto = '" & var_PAGAMENTO & "', tipo = '" & vOrigem & "', tipo_cartao = " & varTipoCartao & ", caixa = '" & StatusBar1.Panels(2).Text & "', CODCAIXA = " & varCodCaixa & ", COD_FUNCIONARIO =" & txtCodFuncionario & " WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = " & txtNumParcela.Text & ");"
        End If
    Else
        If chkJuros.Value = 1 Then
           dbData.Execute "UPDATE parcelas SET status = 1, valor_final = " & Replace(CCur(txtTotal.Text), ",", ".") & ", DESCONTO = " & Replace(CCur(txtDesconto.Text), ",", ".") & ", pagamento = CONVERT(DATETIME, '" & Format(mskPagamento.Text, ocDATA) & "', 103), hora = '" & Format(lblHora.Caption, ocHRMN) & "', juros = " & Replace(CCur(txtTJuros.Text), ",", ".") & ", dias_atrazo = " & txtDias.Text & ", forma_pgto = '" & var_PAGAMENTO & "', tipo = '" & vOrigem & "', tipo_cartao = " & varTipoCartao & ", caixa = '" & StatusBar1.Panels(2).Text & "', CODCAIXA = " & varCodCaixa & ", COD_FUNCIONARIO =" & txtCodFuncionario & " WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = " & txtNumParcela.Text & ") AND (OS_ITEM = " & txtItem.Text & ");"
        Else
           dbData.Execute "UPDATE parcelas SET status = 1, valor_final = " & Replace(CCur(txtTotal.Text), ",", ".") & ", DESCONTO = " & Replace(CCur(txtDesconto.Text), ",", ".") & ", pagamento = CONVERT(DATETIME, '" & Format(mskPagamento.Text, ocDATA) & "', 103), hora = '" & Format(lblHora.Caption, ocHRMN) & "', juros = 0, dias_atrazo = 0, forma_pgto = '" & var_PAGAMENTO & "', tipo = '" & vOrigem & "', tipo_cartao = " & varTipoCartao & ", caixa = '" & StatusBar1.Panels(2).Text & "', CODCAIXA = " & varCodCaixa & ", COD_FUNCIONARIO = " & txtCodFuncionario & " WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = " & txtNumParcela.Text & ") AND (OS_ITEM = " & txtItem.Text & ");"
        End If
    End If
    
       'ver depois dbData.Execute "UPDATE cheque SET status = 1 WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (parcela = " & txtNumParcela.Text & ");"
       dbData.Execute "UPDATE parcelas_haver SET status = 1 WHERE (cod_parcela = " & txtCodParc.Text & ");"
    'End If

    If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo) = vbYes Then
       IMPRIMIR = True
    Else
       IMPRIMIR = False
    End If
    
    If IMPRIMIR = True Then
         'If vQuitarUma = True Then
             If varTipoRecPgto = "CUPOM" Then
                 Imprimir_ReciboCupom
             Else
                 ImprimirPedidoFolha
             End If
         'Else
         '    ImprimirPedidoFolha
         'End If
    End If
Else
    dbData.Execute "UPDATE parcelas SET DATA = CONVERT(DATETIME, '" & Format(mskPagamento.Text, ocDATA) & "', 103) WHERE (codigo = " & vCodParc & ");"
End If

MostrarGrid_Haver
MostrarGrid_Parcelas
MostrarGrid_Historico
LimparObjetos_Parcelas
Somar_Parcelas_Selecionadas
frmParcela.Visible = False
vQuitarUma = False
txtCodParc.Text = ""
txtItem.Text = ""
cboCliente.Locked = False
mskPagamento.Top = 600
lblPgto.Top = 600
cmdCal1.Top = 600
End Sub
Private Sub cmdSalvarAutomatico_Click()
If cboForma.Text = "" Then MsgBox "Selecione uma forma de pagamento!", vbInformation, "Aviso do Sistema": cboForma.SetFocus: Exit Sub
If txtValorAutomatico.Text = "" Or txtValorAutomatico.Text = "0,00" Then MsgBox "Valor incorreto!", vbInformation, "Aviso do Sistema": txtValorAutomatico.SetFocus: Exit Sub

frmPagamento.Visible = True
varPgtoAutomatico = True

Dim varValorTotalParcelas As Currency
Dim varValorParaAbater As Currency

varValorTotalParcelas = lblTotal.Caption
varValorParaAbater = txtValorAutomatico

If varValorParaAbater > varValorTotalParcelas Then MsgBox "O valor adicionado ultrapassa a soma das parcelas", vbInformation, "Aviso do Sistema": Exit Sub

Dim varSomaParcelasSelecionas As Currency
Dim varSomaFutura As Currency
Dim varSobra As Currency

varSomaParcelasSelecionas = 0
varSomaFutura = 0
varSobra = 0

'Dim Total As Currency, SUBTOTAL As Currency, HAVER As Currency
Dim i As Integer

Dim vValorPrimeiraLinha As Currency
vValorPrimeiraLinha = Grid_Parcelas.TextMatrix(1, 11)

If varValorParaAbater < vValorPrimeiraLinha Then
    varSobra = varValorParaAbater
Else
    With Grid_Parcelas
        For i = 1 To .rows - 1
            .Col = 0
            .Row = i
            
            If varSomaParcelasSelecionas < varValorParaAbater Then
                varSomaFutura = varSomaParcelasSelecionas + Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 11)
                If varSomaFutura < varValorParaAbater Then
                    varSomaParcelasSelecionas = varSomaParcelasSelecionas + Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 11)
                    Set Grid_Parcelas.CellPicture = ImgMarcada
                    varSobra = varValorParaAbater - varSomaParcelasSelecionas
                End If
            End If
        Next
    End With

    Somar_Parcelas_Selecionadas
    
    cmdQuitarTodas_Click
    MostrarGrid_Parcelas
End If

Dim varFormaPgtoHaver As String
varFormaPgtoHaver = cboForma.Text
    
'colocando a sobra como haver
Dim varLinhaMarcada As Boolean
varLinhaMarcada = False

With Grid_Parcelas
    For i = 1 To .rows - 1
        .Col = 0
        .Row = i
        
        If varLinhaMarcada = False Then
            If varSobra < Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 11) Then
                Set Grid_Parcelas.CellPicture = ImgMarcada
                varLinhaMarcada = True
                cmdHabilitarHaver_Click
                cboFormaHaver.Text = varFormaPgtoHaver
                txtValorHaver = varSobra
                'Exit Sub
            End If
        End If
    Next
End With

cmdAdicionarHaver_Click
'MostrarGrid_Haver
MostrarGrid_Parcelas
MostrarGrid_Historico
varPgtoAutomatico = False
frmPagamento.Visible = False
lblFormaPgto.Visible = False
cboForma.Visible = False
lblValorAutomatico.Visible = False
txtValorAutomatico.Visible = False
cmdSalvarAutomatico.Visible = False
cmdCancelar.Visible = False
txtValorAutomatico.Text = ""
Exit Sub
End Sub
Private Function ReturnLargueToPrint(ByVal strText As String, iMaxCol As Integer, ByVal strSeparator As String) As String
Dim strCount() As String
Dim strJoin() As String
Dim i As Integer
Dim iLine As Integer

iLine = 0
strCount = Split(strText, strSeparator)

For i = 0 To UBound(strCount)
   If i = 0 Then
      ReDim Preserve strJoin(iLine)
      strJoin(iLine) = strCount(i)
   ElseIf i > 0 Then
      If Len(strJoin(iLine) & strSeparator & strCount(i)) > iMaxCol Then
         iLine = iLine + 1
         ReDim Preserve strJoin(iLine)
         strJoin(iLine) = strCount(i)
      Else
         strJoin(iLine) = strJoin(iLine) & strSeparator & strCount(i)
      End If
   End If
Next
ReturnLargueToPrint = Join(strJoin, vbCrLf)
End Function



Private Sub Form_Load()
SSTab1.Tab = 0
varPgtoAutomatico = False
vQuitarUma = True
vClienteEncontrado = False

cboMES.Text = Format(Date, "mmmm")
cboAno.Text = Year(Date)

'colocar o nome da maquina na barra de status
Dim var_Caixa As String
'Dim var_Maquina As String

'abre o ini
Dim oIni As Ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
'StatusBar1.Panels(2).Text = var_Caixa


'nome da caixa
var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
StatusBar1.Panels(2).Text = var_Caixa

'nome da caixa
'var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
'StatusBar1.Panels(1).Text = var_Maquina


StatusBar1.Panels(4).Text = Format(Date, "dd/mm/yy")

'abrindo arquivo .ini
'Set oIni = New Ini
'oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
Set oIni = Nothing  'fecha o ini

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

'tipo de recibo de pagamento
Set oCfg = sysConfig("TIPORECPGTO")
varTipoRecPgto = oCfg.Value
Set oCfg = Nothing

'logomarca impressa do cupom
Dim sLogo As String
Set oCfg = sysConfig("LOGO_CUPOM")
sLogo = oCfg.Value
Set oCfg = Nothing
If Not Existe(sLogo) Then
    'MsgBox "nao existe"
Else
    If Dir$(sLogo) <> "" Then Set imLogoCupom.Picture = LoadPicture(sLogo)
End If

'tipo de recibo de pagamento
Set oCfg = sysConfig("TIPORECHAVER")
varTipoRecHaver = oCfg.Value
Set oCfg = Nothing

If vCodFunc = 0 Then
    txtCodFuncionario.Text = "1"
Else
    txtCodFuncionario.Text = vCodFunc
End If

LimparGrid_Parcelas
LimparGrid_Haver
LimparGrid_Historico
Set moCombo = New cComboHelper
End Sub
Private Sub MostrarCodCaixa()
sSQL = "SELECT *, CASE status WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END AS varStatus " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    'varCodCaixa = ValidateNull(r("codcaixa"))
    'cmdImprimir.Enabled = True
    'cmdImprimirResumido.Enabled = True
    'cmdAbrirCaixa.Visible = False
    'cmdFecharCaixa.Visible = True
    'cmdTroco.Enabled = True
    'StatusBar1.Panels(3).Text = Format(ValidateNull(r("codcaixa")), "00000")
    'StatusBar1.Panels(4).Text = r("VARSTATUS")
    'vStatusCaixaAtual = r("VARSTATUS")
Else
    If varFluxoCaixa = False Then
        'varCodCaixa = 0
        'cmdImprimir.Enabled = False
        'cmdImprimirResumido.Enabled = False
        'cmdAbrirCaixa.Visible = True
        'cmdFecharCaixa.Visible = False
        'cmdTroco.Enabled = False
        'StatusBar1.Panels(3).Text = Format(0, "00000")
        'StatusBar1.Panels(4).Text = "FECHADO"
        'vStatusCaixaAtual = "FECHADO"
    Else
        'varCodCaixa = StatusBar1.Panels(3).Text
    End If
End If
End Sub

Sub AcaoGridREATIVAR()
'Grid_Historico.Col = 0
var_Contador = 0
For i = 1 To Grid_Historico.rows - 1
   Grid_Historico.Row = i
   If OP = MarcarTodos Then Set Grid_Historico.CellPicture = ImgMarcadaPAGAS
   If OP = DesmarcarTodos Then Set Grid_Historico.CellPicture = imgDesmarcadaPAGAS
   If OP = contar Then
      If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then var_Contador = var_Contador + 1
   End If
Next

If var_Contador = 1 Then
   frmReativar.Enabled = False
   cmdMostrarProdutosREATIVAR.Enabled = True
   cmdHabilitarREATIVAR.Enabled = True
   cmdImprimirParcelas.Visible = True
   cmdImprimirParcQuitSel.Visible = False
   
    'If Grid_Historico.Rows >= 2 Then
    '    sSQL = "SELECT * FROM parcelas_haver WHERE (cod_parcela = " & iCodParc & ") ORDER BY codigo;"
    '    Set r = dbData.OpenRecordset(sSQL)
        
    '    FormatarGrid_HaverPagas r
        
    '    If r.State <> 0 Then r.Close
    '    Set r = Nothing
    'End If
    
ElseIf var_Contador > 1 Then
   frmReativar.Enabled = False
   cmdMostrarProdutosREATIVAR.Enabled = True
   cmdHabilitarREATIVAR.Enabled = False
   cmdImprimirParcelas.Visible = False
   cmdImprimirParcQuitSel.Visible = True
  
   LimparGridHaverPagas
ElseIf var_Contador = 0 Then
   frmReativar.Enabled = False
   cmdMostrarProdutosREATIVAR.Enabled = False
   cmdHabilitarREATIVAR.Enabled = False
   cmdImprimirParcelas.Visible = False
   cmdImprimirParcQuitSel.Visible = True
   LimparGridHaverPagas
End If
   
   'If OP = Contar Then ShowMsg "Qtde de itens selecionados: " & var_Contador, , "Contador"
End Sub

Sub AcaoGrid()
Dim i As Integer
Dim var_Contador As Integer

Grid_Parcelas.Col = 0

For i = 1 To Grid_Parcelas.rows - 1
   Grid_Parcelas.Row = i
   If OP = MarcarTodos Then Set Grid_Parcelas.CellPicture = ImgMarcada
   If OP = DesmarcarTodos Then Set Grid_Parcelas.CellPicture = imgDesmarcada
   If OP = contar Then
      If Grid_Parcelas.CellPicture = ImgMarcada Then var_Contador = var_Contador + 1
   End If
Next

'If MSFlexgrid1.Rows = 0 then
'MsgBox Grid_Parcelas.Rows
If var_Contador = 1 Then
   frmPagamento.Visible = False
   'lblFormaPgto.Visible = False
   'cboForma.Visible = False
   cmdMostrarProdutos.Enabled = True
   cmdMostrarHaveres.Enabled = True
   cmdReativar.Enabled = True
   If Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 9) <> "" Then
    cmdQuitarUma.Visible = True
    cmdHabilitarHaver.Visible = True
    cmdAlterarVenc.Visible = True
   End If
   cmdQuitarTodas.Visible = False
   'mostrar dados do pagamento
   'frmPagamento.Visible = True
   lblFormaPgto.Visible = True
   cboForma.Visible = True
   lblPgto.Visible = True
   mskPagamento.Visible = True
   cmdCal1.Visible = True
   txtDias.Visible = True
   chkJuros.Visible = True
   txtJuros.Visible = True
   lblJuros.Visible = True
   txtTJuros.Visible = True
   chkMulta.Visible = True
   txtMulta.Visible = True
   lblHaverPgto.Visible = True
   txtTotalHaver.Visible = True
   lblTotalPgto.Visible = True
   txtTotal.Visible = True
   txtDesconto.Visible = True
   cmdQuitarAutomatico.Visible = False
   cmdCancelar.Visible = False
   lblValorAutomatico.Visible = False
   txtValorAutomatico.Visible = False
ElseIf var_Contador > 1 Then
   cmdQuitarTodas.Visible = True
   cmdMostrarProdutos.Enabled = False
   cmdMostrarHaveres.Enabled = False
   cmdReativar.Enabled = False
   cmdQuitarUma.Visible = False
   cmdHabilitarHaver.Visible = False
   cmdAlterarVenc.Visible = False
   frmParcela.Visible = False
   'frmPagamento.Visible = False
   cmdSalvar.Visible = False
   cmdCancelar.Visible = True
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   'mostrar dados do pagamento
   frmPagamento.Visible = True
   lblFormaPgto.Visible = True
   cboForma.Visible = True
   lblPgto.Visible = False
   mskPagamento.Visible = False
   cmdCal1.Visible = False
   txtDias.Visible = False
   chkJuros.Visible = False
   txtJuros.Visible = False
   lblJuros.Visible = False
   txtTJuros.Visible = False
   chkMulta.Visible = False
   txtMulta.Visible = False
   lblHaverPgto.Visible = False
   txtTotalHaver.Visible = False
   lblTotalPgto.Visible = False
   txtTotal.Visible = False
   txtDesconto.Visible = False
   cmdQuitarAutomatico.Visible = False
    lblValorAutomatico.Visible = False
    txtValorAutomatico.Visible = False
    Limpar_Campos_Reativar
ElseIf var_Contador = 0 Then
   cmdMostrarProdutos.Enabled = False
   cmdMostrarHaveres.Enabled = False
   cmdReativar.Enabled = False
   cmdQuitarUma.Visible = False
   cmdHabilitarHaver.Visible = False
   cmdQuitarTodas.Visible = False
   cmdAlterarVenc.Visible = False
   frmParcela.Visible = False
   frmPagamento.Visible = False
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   'mostra dados do pagamento
   frmPagamento.Visible = False
   frmPagamento.Visible = False
   lblFormaPgto.Visible = False
   cboForma.Visible = False
   lblPgto.Visible = False
   mskPagamento.Visible = False
   cmdCal1.Visible = False
   txtDias.Visible = False
   chkJuros.Visible = False
   txtJuros.Visible = False
   lblJuros.Visible = False
   txtTJuros.Visible = False
   chkMulta.Visible = False
   txtMulta.Visible = False
   lblHaverPgto.Visible = False
   txtTotalHaver.Visible = False
   lblTotalPgto.Visible = False
   txtTotal.Visible = False
   txtDesconto.Visible = False
   cmdQuitarAutomatico.Visible = True
    lblValorAutomatico.Visible = False
   txtValorAutomatico.Visible = False
   Limpar_Campos_Reativar
End If
End Sub
Private Function Verificar_Caixa_Reativar() As Integer
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim cxaStatus As Integer
  
   cxaStatus = -1   'Não foi aberto
   'If cmdAlterar.Enabled = True Then
      sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskConPgto.FormattedText, ocDATA) & "', 103)) AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
   'Else
      'sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = '" & Format(StatusBar1.Panels(4), ocDATA_EUA) & "') AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
   'End If
   
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cxaStatus = CInt(ValidateNull(r("status")))   '0 = aberto, 1 = fechado
   If r.State <> 0 Then r.Close
   Set r = Nothing
   Verificar_Caixa_Reativar = cxaStatus
End Function

Private Function Verificar_Caixa_Haver() As Integer
Dim sSQL As String
Dim r As ADODB.Recordset
Dim cxaStatus As Integer

cxaStatus = -1   'Não foi aberto
'If cmdAlterar.Enabled = True Then
   sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskDataHaver.FormattedText, ocDATA) & "', 103)) AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
'Else
   'sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = '" & Format(StatusBar1.Panels(4), ocDATA_EUA) & "') AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
'End If

Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then cxaStatus = CInt(ValidateNull(r("status")))   '0 = aberto, 1 = fechado
If r.State <> 0 Then r.Close
Set r = Nothing

Verificar_Caixa_Haver = cxaStatus
End Function

Private Function Verificar_Caixa_Baixa() As Integer
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim cxaStatus As Integer
  
   cxaStatus = -1   'Não foi aberto
   'If cmdAlterar.Enabled = True Then
      sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskPagamento.FormattedText, ocDATA) & "', 103)) AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
   'Else
      'sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = '" & Format(StatusBar1.Panels(4), ocDATA_EUA) & "') AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
   'End If
   
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cxaStatus = CInt(ValidateNull(r("status")))   '0 = aberto, 1 = fechado
   If r.State <> 0 Then r.Close
   Set r = Nothing
   Verificar_Caixa_Baixa = cxaStatus
End Function

Private Sub LimparObjetos_Parcelas_Haver()
   txtCodHaver.Text = ""
   mskDataHaver.Mask = ""
   mskDataHaver.Text = ""
   txtValorHaver.Text = ""
End Sub

Private Sub Calcular_Valores()
Dim Total As Currency, HAVER As Currency, TOTAL_GERAL As Currency, JUROS As Currency, var_MULTAS As Currency, varDesc As Currency

If txtMulta.Text = "" Then txtMulta.Text = FormatNumber(0, 2)
If txtDesconto.Text = "" Then txtDesconto.Text = FormatNumber(0, 2)
If txtJuros.Text = "" Then txtJuros.Text = FormatNumber("0,33", 2)

If txtValor.Text = "" Then Total = 0 Else Total = txtValor
If txtDesconto.Text = "" Then varDesc = 0 Else varDesc = txtDesconto.Text
If txtTotalHaver.Text = "" Then HAVER = 0 Else HAVER = txtTotalHaver

If chkJuros.Value = 1 Then JUROS = txtTJuros Else JUROS = 0
If chkMulta.Value = 1 Then var_MULTAS = txtMulta Else var_MULTAS = 0

TOTAL_GERAL = var_MULTAS + JUROS + Total
TOTAL_GERAL = TOTAL_GERAL - HAVER
TOTAL_GERAL = TOTAL_GERAL - varDesc
txtTotal = Format(TOTAL_GERAL, ocMONEY)
End Sub

Private Sub AutoNumeracao_Haver()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS codigo_haver FROM parcelas_haver;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodHaver.Text = r("codigo_haver") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub MostrarGrid_Haver()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long

'Mostrar no grid
If txtCodParc.Text = "" Then Exit Sub

sSQL = "SELECT * FROM parcelas_haver WHERE (cod_parcela = " & txtCodParc.Text & ") ORDER BY haver, codigo;"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)
FormatarGrid_Haver r
If r.State <> 0 Then r.Close
Set r = Nothing

lblQuantHaver.Caption = Format(totalRegistros, "00") & " haver(es)"
'Calcular_Juros
Calcular_Valores
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If vChamouCaixa = "PDV" Then
    Parcelas.Hide
    'PDV.Show 'desativei somente para geerar o online comerce
Else
    Parcelas.Hide
    'If FormExists("PDV") Then
     '   FormExists("PDV").Hide
    'End If
    'PDV.Show 1
End If

varFluxoCaixa = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'HabilitaObjetosVenda False
Set moCombo = Nothing
End Sub

Public Sub loadParcelas(Cliente As String)
   cboCliente.Text = Cliente
End Sub

Private Sub Grid_Haver_Click()
ConsultarCaixaAtual
i = Grid_Haver.Row
If Grid_Haver.TextMatrix(i, 1) <> "" Then cmdImprimirHaver.Enabled = True

If Grid_Haver.TextMatrix(i, 2) = varCodCaixa Then cmdRemoverHaver.Enabled = True Else:   cmdRemoverHaver.Enabled = False

End Sub

Private Sub Grid_Historico_Click()
'Dim i As Long
Dim vCodParc As Long
vCodParc = 0

'marcar a parcela
'i = Grid_Historico.Row
'If Grid_Historico.TextMatrix(i, 1) = "" Then Exit Sub
'iCodParc = Grid_Historico.TextMatrix(i, 1)

If Grid_Historico.Col <> 0 Then Exit Sub
'If Grid_Historico.TextMatrix(i, 1) = "" Then Exit Sub

If Grid_Historico.CellPicture = imgDesmarcadaPAGAS Then
   Set Grid_Historico.CellPicture = ImgMarcadaPAGAS
Else
   Set Grid_Historico.CellPicture = imgDesmarcadaPAGAS
End If

OP = contar
AcaoGridREATIVAR

If var_Contador = 1 Then
    'If Grid_Historico.Rows >= 2 Then
    'i = Grid_Historico.Row
    For f = 0 To Grid_Historico.rows - 1
       Grid_Historico.Row = f
       Grid_Historico.Col = 0
       
       If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then
          vCodParc = (Grid_Historico.TextMatrix(Grid_Historico.Row, 1))
       End If
    Next
        
    sSQL = "SELECT * FROM parcelas_haver WHERE (cod_parcela = " & vCodParc & ") ORDER BY haver, codigo;"
    Set r = dbData.OpenRecordset(sSQL)
    'Debug.Print sSQL
    FormatarGrid_HaverPagas r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If

Somar_Parcelas_SelecionadasQuitadas
End Sub

Private Sub Grid_Parcelas_Click()
'marcar a parcela
If Grid_Parcelas.Col <> 0 Then Exit Sub

If Grid_Parcelas.CellPicture = imgDesmarcada Then
   Set Grid_Parcelas.CellPicture = ImgMarcada
Else
   Set Grid_Parcelas.CellPicture = imgDesmarcada
End If

OP = contar
AcaoGrid
Somar_Parcelas_Selecionadas
End Sub


    

Private Sub mskConData_GotFocus()
   SelectControl mskConData
End Sub

Private Sub mskConData_KeyPress(KeyAscii As Integer)
   mskConData.Mask = "##/##/##"
End Sub

Private Sub mskConData_LostFocus()
   If mskConData.Text = "" Or mskConData.Text = "__/__/__" Then
      mskConData.Mask = ""
      mskConData.Text = ""
   Else
      If Not IsDate(mskConData.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskConData.SetFocus
      End If
   End If
End Sub

Private Sub mskConPgto_GotFocus()
   SelectControl mskConPgto
End Sub

Private Sub mskConPgto_KeyPress(KeyAscii As Integer)
   mskConPgto.Mask = "##/##/##"
End Sub

Private Sub mskConPgto_LostFocus()
   If mskConPgto.Text = "" Or mskConPgto.Text = "__/__/__" Then
      mskConPgto.Mask = ""
      mskConPgto.Text = ""
   Else
      If Not IsDate(mskConPgto.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskConPgto.SetFocus
      End If
   End If
End Sub

Private Sub mskData_GotFocus()
   SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
   If mskData.Text = "" Or mskData.Text = "__/__/__" Then
      mskData.Mask = ""
      mskData.Text = ""
   Else
      If Not IsDate(mskData.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskData.SetFocus
      End If
   End If
End Sub

Private Sub mskDataHaver_GotFocus()
   SelectControl mskDataHaver
End Sub

Private Sub mskDataHaver_KeyPress(KeyAscii As Integer)
   mskDataHaver.Mask = "##/##/##"
End Sub

Private Sub mskDataHaver_LostFocus()
   If mskDataHaver.Text = "" Or mskDataHaver.Text = "__/__/__" Then
      mskDataHaver.Mask = ""
      mskDataHaver.Text = ""
   Else
      If IsDate(mskDataHaver.Text) Then
         txtValorHaver.SetFocus
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskDataHaver.SetFocus
      End If
   End If
End Sub

Private Sub mskPagamento_GotFocus()
   SelectControl mskPagamento
End Sub

Private Sub mskPagamento_KeyPress(KeyAscii As Integer)
mskPagamento.Mask = "##/##/##"
End Sub

Private Sub mskPagamento_LostFocus()
   If mskPagamento.Text = "" Or mskPagamento.Text = "__/__/__" Then
      mskPagamento.Mask = ""
      mskPagamento.Text = ""
   Else
      If IsDate(mskPagamento.Text) Then
        If txtDias.Visible = True Then
            Calcular_Dias
        End If
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskPagamento.SetFocus
      End If
   End If
End Sub

Private Sub optJurosNao_Click()
   TxtCodCliente_Change
End Sub

Private Sub optJurosSim_Click()
   TxtCodCliente_Change
End Sub

Private Sub optPgto_Click()
MostrarGrid_Historico
End Sub

Private Sub optTodas_Click()
MostrarGrid_Historico
End Sub

Private Sub optVenc_Click()
MostrarGrid_Historico
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      Exit Sub
   ElseIf SSTab1.Tab = 1 Then
      If frmHaver.Enabled = True Then
        If txtValorHaver.Enabled = True Then
            'txtValorHaver.SetFocus
        End If
      End If
   ElseIf SSTab1.Tab = 2 Then
      Exit Sub
   End If
End Sub

Private Sub Timer1_Timer()
   lblHora.Caption = Format(Time, "hh:mm")
End Sub

Private Sub TxtCodCliente_Change()
If txtCodCliente.Text = "" Then Exit Sub
LimparObjetos_Parcelas

If chkCodPedido.Value = Checked Then
   cboCliente.Text = ""
   sSQL = "SELECT * FROM cliente WHERE (codigo= " & txtCodCliente.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cboCliente.Text = r("nome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If

If txtCodCliente.Text <> "" Then
   MostrarGrid_Parcelas
   MostrarGrid_Historico
   lblClienteHaver.Caption = cboCliente.Text
   lblCliente.Caption = cboCliente.Text
Else
   LimparObjetos_Parcelas
   LimparGrid_Parcelas
   LimparObjetos_GridParcelas
End If
vQuitarUma = False
End Sub

Private Sub txtCodPed_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then Mostrar_Pedidos_Por_Codigo
End Sub

Private Sub txtCodPed_LostFocus()
   Mostrar_Pedidos_Por_Codigo
End Sub

Private Sub txtConValor_GotFocus()
   SelectControl txtCONValor
End Sub

Private Sub txtConValor_LostFocus()
   If txtCONValor.Text = "" Then txtCONValor.Text = FormatNumber(0, 2) Else txtCONValor.Text = FormatNumber(txtCONValor.Text, 2)
End Sub

Private Sub txtDesconto_GotFocus()
SelectControl txtDesconto
End Sub


Private Sub txtDesconto_LostFocus()
If txtDesconto.Text = "" Then txtDesconto.Text = FormatNumber(0, 2) Else txtDesconto.Text = FormatNumber(txtDesconto.Text, 2)

'If CCur(txtDesconto.Text) > CCur(txtTJuros.Text) Then
'    MsgBox "O Desconto somente pode ser concedido em cima do valor do juros!", vbInformation, "Aviso do Sistema"
'    txtDesconto.Text = FormatNumber(0, 2)
'    Exit Sub
'Else
    Calcular_Valores
'End If
End Sub


Private Sub txtDias_Change()
   If txtDias.Text <> "" Then MostrarGrid_Haver
End Sub

Private Sub txtJuros_Change()
   MostrarGrid_Haver
End Sub

Private Sub txtJuros_GotFocus()
   SelectControl txtJuros
End Sub

Private Sub txtJuros_LostFocus()
   MostrarGrid_Haver
   If txtJuros.Text = "" Then txtJuros.Text = FormatNumber(0, 2) Else txtJuros.Text = FormatNumber(txtJuros.Text, 2)
End Sub

Private Sub txtMulta_Change()
   MostrarGrid_Haver
End Sub

Private Sub txtMulta_GotFocus()
   SelectControl txtMulta
End Sub

Private Sub txtMulta_LostFocus()
MostrarGrid_Haver
If txtMulta.Text = "" Then txtMulta.Text = FormatNumber(0, 2) Else txtMulta.Text = FormatNumber(txtMulta.Text, 2)
End Sub

Private Sub txtValor_Change()
   If txtNumParcela.Text <> "" And mskData.Text <> "" And txtValor.Text <> "" Then
      Label16.Enabled = True
      Label15.Enabled = True
      mskDataHaver.Enabled = True
      txtValorHaver.Enabled = True
      cmdAdicionarHaver.Enabled = True
      cmdRemoverHaver.Enabled = True
      mskDataHaver.Text = Format(Date, "dd/mm/yy")
   Else
      Label16.Enabled = False
      Label15.Enabled = False
      mskDataHaver.Enabled = False
      txtValorHaver.Enabled = False
      cmdAdicionarHaver.Enabled = False
      cmdRemoverHaver.Enabled = False
   End If
End Sub

Private Sub txtValor_GotFocus()
   SelectControl txtValor
End Sub

Private Sub txtValor_LostFocus()
   If txtValor.Text = "" Then txtValor.Text = FormatNumber(0, 2) Else txtValor.Text = FormatNumber(txtValor.Text, 2)
End Sub

Private Sub txtValorAutomatico_GotFocus()
SelectControl txtValorAutomatico
End Sub


Private Sub txtValorAutomatico_LostFocus()
If txtValorAutomatico.Text = "" Then txtValorAutomatico.Text = FormatNumber(0, 2) Else txtValorAutomatico.Text = FormatNumber(txtValorAutomatico.Text, 2)
End Sub


Private Sub txtValorHaver_GotFocus()
   SelectControl txtValorHaver
End Sub

Private Sub txtValorHaver_LostFocus()
   'On Error GoTo Erro
   
   If txtValorHaver.Text = "" Then txtValorHaver.Text = Format(0, "##,##0.00") Else txtValorHaver.Text = Format(txtValorHaver, "##,##0.00")
   If txtValor.Text = "" Or txtValorHaver.Text = "" Then Exit Sub
   
   'cmdAdicionarHaver.SetFocus
   
'Erro:
'   ShowMsg "O valor digitado é inválido!", vbExclamation
'   Exit Sub
End Sub
