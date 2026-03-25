VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Parcelas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PARCELAS"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "Parcelas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   495
      Left            =   8340
      TabIndex        =   45
      Top             =   240
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   873
      BTYPE           =   2
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Parcelas.frx":23D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   9645
      TabIndex        =   43
      Top             =   60
      Width           =   9675
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
         Left            =   1440
         TabIndex        =   44
         Top             =   240
         Width           =   1725
      End
      Begin VB.Image Image10 
         Height          =   900
         Left            =   240
         Picture         =   "Parcelas.frx":23EE
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   60
      Top             =   600
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   60
      TabIndex        =   17
      Top             =   1080
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   14420
      _Version        =   393216
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
      TabPicture(0)   =   "Parcelas.frx":90F1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdMostrarProdutos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdMarcarCheck"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdHabilitarHaver"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdHabilitarQuitar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Picture2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmPagamento"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCodParc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "frmParcela"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdQuitarTodas"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "HAVER"
      TabPicture(1)   =   "Parcelas.frx":910D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmHaver"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PAGAS"
      TabPicture(2)   =   "Parcelas.frx":9129
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmReativar"
      Tab(2).Control(1)=   "Picture3"
      Tab(2).Control(2)=   "txtCONcodParc"
      Tab(2).Control(3)=   "cmdTodasREATIVAR"
      Tab(2).Control(4)=   "cmdHabilitarREATIVAR"
      Tab(2).Control(5)=   "cmdMarcarTodasREATIVAR"
      Tab(2).Control(6)=   "cmdMostrarProdutosREATIVAR"
      Tab(2).Control(7)=   "ImgMarcadaPAGAS"
      Tab(2).Control(8)=   "imgDesmarcadaPAGAS"
      Tab(2).Control(9)=   "lblTotalHistorico"
      Tab(2).Control(10)=   "lblQuantHistorico"
      Tab(2).Control(11)=   "lblCliente"
      Tab(2).ControlCount=   12
      Begin VB.Frame frmReativar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Parcela para reativar"
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
         Height          =   2175
         Left            =   -73500
         TabIndex        =   79
         Top             =   5880
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtConCodOS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   300
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   780
            Width           =   1215
         End
         Begin VB.TextBox txtConCodPedido 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   780
            Width           =   735
         End
         Begin VB.TextBox txtConNumParcela 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   780
            Width           =   435
         End
         Begin VB.TextBox txtConValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   780
            Width           =   1155
         End
         Begin MSMask.MaskEdBox mskConData 
            Height          =   315
            Left            =   2820
            TabIndex        =   84
            Top             =   780
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskConPgto 
            Height          =   315
            Left            =   3840
            TabIndex        =   85
            Top             =   780
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdReativar 
            Height          =   675
            Left            =   2520
            TabIndex        =   92
            Top             =   1320
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1191
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
            MICON           =   "Parcelas.frx":9145
            PICN            =   "Parcelas.frx":9161
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
            Left            =   300
            TabIndex        =   91
            Top             =   540
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
            Left            =   1560
            TabIndex        =   90
            Top             =   540
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
            Left            =   2340
            TabIndex        =   89
            Top             =   540
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
            Left            =   2820
            TabIndex        =   88
            Top             =   540
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
            Left            =   4860
            TabIndex        =   87
            Top             =   540
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
            Left            =   3840
            TabIndex        =   86
            Top             =   540
            Width           =   465
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdQuitarTodas 
         Height          =   435
         Left            =   6180
         TabIndex        =   63
         Top             =   4380
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "QUITAR TODAS SELECIONADAS"
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
         MICON           =   "Parcelas.frx":947B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture3 
         Height          =   4095
         Left            =   -74880
         ScaleHeight     =   4035
         ScaleWidth      =   9375
         TabIndex        =   61
         Top             =   840
         Width           =   9435
         Begin VB.OptionButton optVenc 
            Caption         =   "Vencimento"
            Height          =   195
            Left            =   2760
            TabIndex        =   106
            Top             =   120
            Width           =   1395
         End
         Begin VB.OptionButton optPgto 
            Caption         =   "Pagamento"
            Height          =   195
            Left            =   1500
            TabIndex        =   105
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
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label Label10 
            Caption         =   "Organizar por:"
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
            TabIndex        =   104
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
         Height          =   2655
         Left            =   600
         TabIndex        =   48
         Top             =   5160
         Visible         =   0   'False
         Width           =   4035
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   1500
            TabIndex        =   7
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
            TabIndex        =   5
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
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   600
            Width           =   1275
         End
         Begin VB.TextBox txtCodOS 
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
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   600
            Width           =   1215
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   1260
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632319
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdSalvar 
            Height          =   555
            Left            =   180
            TabIndex        =   93
            Top             =   1860
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   979
            BTYPE           =   3
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
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":9497
            PICN            =   "Parcelas.frx":94B3
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
            Height          =   555
            Left            =   2100
            TabIndex        =   94
            Top             =   1860
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   979
            BTYPE           =   3
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
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas.frx":FD7D
            PICN            =   "Parcelas.frx":FD99
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.TextBox txtCONcodParc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70860
         TabIndex        =   42
         Top             =   5040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCodParc 
         Height          =   315
         Left            =   4800
         TabIndex        =   38
         Top             =   5160
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
         Height          =   3135
         Left            =   6420
         TabIndex        =   21
         Top             =   4860
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ComboBox cboForma 
            Height          =   315
            Left            =   1620
            TabIndex        =   68
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtTotalHaver 
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
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1455
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
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2700
            Width           =   1455
         End
         Begin VB.TextBox txtDias 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtJuros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1620
            TabIndex        =   11
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtMulta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1620
            TabIndex        =   14
            Top             =   1995
            Width           =   1455
         End
         Begin VB.TextBox txtTJuros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1650
            Width           =   1455
         End
         Begin MSMask.MaskEdBox mskPagamento 
            Height          =   315
            Left            =   1620
            TabIndex        =   8
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptChar      =   "_"
         End
         Begin VB.CheckBox chkJuros 
            Caption         =   "Juros/Dia (%):"
            Height          =   315
            Left            =   345
            TabIndex        =   10
            Top             =   1305
            Width           =   1335
         End
         Begin VB.CheckBox chkMulta 
            Caption         =   "Multa (R$):"
            Height          =   315
            Left            =   540
            TabIndex        =   13
            Top             =   1995
            Width           =   1095
         End
         Begin VB.Label lblFormaPgto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pgto:"
            Height          =   195
            Left            =   510
            TabIndex        =   69
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
            Left            =   480
            TabIndex        =   39
            Top             =   2340
            Width           =   1110
         End
         Begin VB.Label lblPgto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Pagamento:"
            Height          =   315
            Left            =   480
            TabIndex        =   25
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label lblAtrazo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Dias em Atrazo:"
            Height          =   315
            Left            =   480
            TabIndex        =   24
            Top             =   960
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
            Left            =   585
            TabIndex        =   23
            Top             =   2700
            Width           =   1005
         End
         Begin VB.Label lblJuros 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Juros (R$):"
            Enabled         =   0   'False
            Height          =   315
            Left            =   195
            TabIndex        =   22
            Top             =   1650
            Width           =   1395
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   3795
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   9375
         TabIndex        =   18
         Top             =   420
         Width           =   9435
         Begin VB.CheckBox chkCodPedido 
            Caption         =   "Cód. Pedido"
            Height          =   195
            Left            =   6120
            TabIndex        =   71
            Top             =   60
            Width           =   1335
         End
         Begin VB.TextBox txtCodPed 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6120
            TabIndex        =   70
            Top             =   300
            Width           =   1515
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Parcelas 
            Height          =   2475
            Left            =   60
            TabIndex        =   67
            Top             =   720
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4366
            _Version        =   393216
            FixedCols       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
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
            TabIndex        =   64
            Top             =   60
            Width           =   1635
            Begin VB.OptionButton optJurosNao 
               Caption         =   "Não"
               Height          =   195
               Left            =   840
               TabIndex        =   66
               Top             =   300
               Width           =   675
            End
            Begin VB.OptionButton optJurosSim 
               Caption         =   "Sim"
               Height          =   195
               Left            =   180
               TabIndex        =   65
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
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Todos:"
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
            Left            =   5625
            TabIndex        =   102
            Top             =   3480
            Width           =   600
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Selecionado:"
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
            TabIndex        =   101
            Top             =   3240
            Width           =   1125
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
            Left            =   7320
            TabIndex        =   100
            ToolTipText     =   "Haveres"
            Top             =   3240
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
            Left            =   6300
            TabIndex        =   99
            ToolTipText     =   "Sub-Total"
            Top             =   3240
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
            Left            =   8340
            TabIndex        =   96
            ToolTipText     =   "Sub-Total"
            Top             =   3240
            Width           =   990
         End
         Begin VB.Image imgDesmarcada 
            Height          =   195
            Left            =   4080
            Picture         =   "Parcelas.frx":1683D
            Top             =   3360
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image ImgMarcada 
            Height          =   195
            Left            =   3840
            Picture         =   "Parcelas.frx":18BB9
            Top             =   3360
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
            Left            =   6300
            TabIndex        =   58
            ToolTipText     =   "Sub-Total"
            Top             =   3480
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
            Left            =   7320
            TabIndex        =   57
            ToolTipText     =   "Haveres"
            Top             =   3480
            Width           =   990
         End
         Begin VB.Label lblQuantParc 
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
            TabIndex        =   54
            Top             =   3240
            Width           =   225
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
            Left            =   8340
            TabIndex        =   47
            ToolTipText     =   "Total"
            Top             =   3480
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
            TabIndex        =   20
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
         TabIndex        =   26
         Top             =   420
         Width           =   9435
         Begin VB.TextBox txtCodHaver 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            TabIndex        =   27
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox Picture7 
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   9135
            TabIndex        =   29
            Top             =   660
            Width           =   9195
            Begin VB.ComboBox cboFormaHaver 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   2340
               TabIndex        =   32
               Top             =   300
               Width           =   1875
            End
            Begin VB.TextBox txtValorHaver 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1140
               TabIndex        =   31
               Top             =   300
               Width           =   1155
            End
            Begin MSMask.MaskEdBox mskDataHaver 
               Height          =   315
               Left            =   60
               TabIndex        =   30
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
               Left            =   4260
               TabIndex        =   33
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
               MICON           =   "Parcelas.frx":1AFB8
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
               Left            =   5640
               TabIndex        =   35
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
               MICON           =   "Parcelas.frx":1AFD4
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
               TabIndex        =   95
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
               TabIndex        =   37
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
               TabIndex        =   36
               Top             =   60
               Width           =   510
            End
         End
         Begin VB.PictureBox Picture8 
            Height          =   4635
            Left            =   120
            ScaleHeight     =   4575
            ScaleWidth      =   9135
            TabIndex        =   28
            Top             =   1500
            Width           =   9195
            Begin MSFlexGridLib.MSFlexGrid Grid_Haver 
               Height          =   4215
               Left            =   60
               TabIndex        =   34
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
               TabIndex        =   56
               Top             =   4320
               Width           =   225
            End
            Begin VB.Label lblTotalHaver 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   8700
               TabIndex        =   55
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
               Name            =   "MS Sans Serif"
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
            TabIndex        =   46
            Top             =   240
            Width           =   9195
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdHabilitarQuitar 
         Height          =   435
         Left            =   4920
         TabIndex        =   1
         Top             =   4380
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
         MICON           =   "Parcelas.frx":1AFF0
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
         Left            =   7260
         TabIndex        =   2
         Top             =   4380
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
         MICON           =   "Parcelas.frx":1B00C
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
         Height          =   435
         Left            =   120
         TabIndex        =   62
         Top             =   4380
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   767
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
         MICON           =   "Parcelas.frx":1B028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   120
         TabIndex        =   72
         Top             =   4860
         Width           =   9435
         Begin ChamaleonBtn.chameleonButton cmdAlterar 
            Height          =   555
            Left            =   4800
            TabIndex        =   97
            Top             =   1500
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
            MICON           =   "Parcelas.frx":1B044
            PICN            =   "Parcelas.frx":1B060
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
            Left            =   4800
            TabIndex        =   98
            Top             =   2100
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
            MICON           =   "Parcelas.frx":1B93A
            PICN            =   "Parcelas.frx":1B956
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
      Begin ChamaleonBtn.chameleonButton cmdMostrarProdutos 
         Height          =   435
         Left            =   2460
         TabIndex        =   73
         Top             =   4380
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "MOSTRAR PRODUTOS"
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
         MICON           =   "Parcelas.frx":1BC70
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdTodasREATIVAR 
         Height          =   435
         Left            =   -68820
         TabIndex        =   75
         Top             =   5340
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "REATIVAR TODAS SELECIONADAS"
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
         MICON           =   "Parcelas.frx":1BC8C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdHabilitarREATIVAR 
         Height          =   435
         Left            =   -70380
         TabIndex        =   76
         Top             =   5340
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
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
         MICON           =   "Parcelas.frx":1BCA8
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
         Height          =   435
         Left            =   -74880
         TabIndex        =   77
         Top             =   5340
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   767
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
         MICON           =   "Parcelas.frx":1BCC4
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
         Height          =   435
         Left            =   -72660
         TabIndex        =   78
         Top             =   5340
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   767
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
         MICON           =   "Parcelas.frx":1BCE0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image ImgMarcadaPAGAS 
         Height          =   195
         Left            =   -73620
         Picture         =   "Parcelas.frx":1BCFC
         Top             =   5040
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgDesmarcadaPAGAS 
         Height          =   195
         Left            =   -73320
         Picture         =   "Parcelas.frx":1E0FB
         Top             =   5040
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblTotalHistorico 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   -65820
         TabIndex        =   60
         Top             =   5040
         Width           =   390
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
         Left            =   -74820
         TabIndex        =   59
         Top             =   5040
         Width           =   225
      End
      Begin VB.Label lblCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Escolha o nome do cliente na aba ABAIXO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   9435
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   103
      Top             =   9360
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10874
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
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
            TextSave        =   "10:39"
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
      Left            =   5220
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Parcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum TipoOP
   MarcarTodos = 1
   DesmarcarTodos = 2
   Contar = 3
End Enum                   'termina aqui

Private moCombo As cComboHelper

Dim CAIXA_FECHADO_BAIXA As Boolean
Dim CAIXA_FECHADO_HAVER As Boolean
Dim CAIXA_FECHADO_REATIVAR As Boolean
Dim Imprimir As Boolean

Dim OP As TipoOP           'usado para o check das parcelas

Private Sub FormatarGrid_Haver(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Haver
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1050
      .ColWidth(3) = 1050
      .ColWidth(4) = 2500
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "FORMA"
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
            For i = 1 To .Rows - 1
               .Row = i
               .Col = 3
               .CellBackColor = &HC0C0FF
            Next
            
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("haver"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("valor_haver"), ocMONEY)
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("forma_pgto"))
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblTotalHaver.Caption = Format(SomaGrid(Grid_Haver, 3), ocMONEY)
   txtTotalHaver.Text = Format(SomaGrid(Grid_Haver, 3), ocMONEY)
End Sub

Private Sub LimparGrid_Historico()
   Dim i As Integer
   
   With Grid_Historico
      .Visible = False
      
      .Clear
      .Cols = 9
      .Rows = 2
      
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
      
      .Rows = .Rows + 1
      .Redraw = True
      .Rows = .Rows - 1
      .Visible = True
   End With
End Sub

Private Sub FormatarGrid_Historico(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Historico
      .Visible = False
      
      .Clear
      .Cols = 11
      .Rows = 2
      
      .ColWidth(0) = 300
      .ColWidth(1) = 0
      .ColWidth(2) = 1050
      .ColWidth(3) = 900
      .ColWidth(4) = 500
      .ColWidth(5) = 1000
      .ColWidth(6) = 1100
      .ColWidth(7) = 750
      .ColWidth(8) = 1000
      .ColWidth(9) = 1000
      .ColWidth(10) = 1000
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "ORIGEM"
      .TextMatrix(0, 3) = "PEDIDO"
      .TextMatrix(0, 4) = "No."
      .TextMatrix(0, 5) = "VENC."
      .TextMatrix(0, 6) = "SUBTOTAL"
      .TextMatrix(0, 7) = "DIAS"
      .TextMatrix(0, 8) = "JUROS"
      .TextMatrix(0, 9) = "TOTAL"
      .TextMatrix(0, 10) = "PGTO"
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
            For i = 1 To .Rows - 1
               .Row = i
               .Col = 10
               .CellBackColor = &HC0FFFF
            Next
            
            .TextMatrix(.Rows - 1, 1) = rTabela("cod")
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("campo00"))
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("campo01"), "000000")
            .TextMatrix(.Rows - 1, 4) = rTabela("campo02")
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("campo03"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("campo04"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_dias"))
            .TextMatrix(.Rows - 1, 8) = Format(rTabela("var_juros"), ocMONEY)
            .TextMatrix(.Rows - 1, 9) = Format(rTabela("var_vfinal"), ocMONEY)
            .TextMatrix(.Rows - 1, 10) = Format(rTabela("campo06"), "dd/mm/yy")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Redraw = True
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 9
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         Set .CellPicture = imgDesmarcadaPAGAS
         .CellPictureAlignment = 4
      Next
      
      .Redraw = True
      .Visible = True
   End With
   
   lblTotalHistorico.Caption = Format(SomaGrid(Grid_Historico, 9), ocMONEY)
End Sub

Private Sub FormatarGrid_Parcelas2(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Parcelas
      .Clear
      .Cols = 11
      .Rows = 2
      
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
            For i = 1 To .Rows - 1
               .Row = i
               .Col = 7:   .CellBackColor = &HC0FFFF
               .Col = 9:   .CellBackColor = &HC0C0FF
            Next
            
            .TextMatrix(.Rows - 1, 1) = rTabela("cod")
            .TextMatrix(.Rows - 1, 2) = rTabela("campo05")
            .TextMatrix(.Rows - 1, 3) = rTabela("campo00")
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("campo01"), "000000")
            .TextMatrix(.Rows - 1, 5) = rTabela("campo02")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("campo03"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("campo04"), ocMONEY)
            
            If Not IsNull(rTabela("campo06")) Then
               .TextMatrix(.Rows - 1, 8) = Format(rTabela("campo06"), ocMONEY)
            Else
               .TextMatrix(.Rows - 1, 8) = Format(0, ocMONEY)
            End If
            
            .TextMatrix(.Rows - 1, 9) = Format(.TextMatrix(.Rows - 1, 7) - .TextMatrix(.Rows - 1, 8), ocMONEY)
            .TextMatrix(.Rows - 1, 10) = rTabela("var_atrazo")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Redraw = True
      .Rows = .Rows - 1
   End With
   
   lblSubTotal.Caption = Format(SomaGrid(Grid_Parcelas, 7), ocMONEY)
   lblHaver.Caption = Format(SomaGrid(Grid_Parcelas, 8), ocMONEY)
   lblTotal.Caption = Format(SomaGrid(Grid_Parcelas, 9), ocMONEY)
End Sub

Private Sub FormatarGrid_Parcelas(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With Grid_Parcelas
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 12
      .Rows = 2
      
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
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codparcela")
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("campo00"))
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("campo01"), "000000")
            .TextMatrix(.Rows - 1, 4) = rTabela("campo02")
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("valor"), ocMONEY)
            
            'If optJurosSim = True Then
               .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_atrazo"))
               .TextMatrix(.Rows - 1, 8) = Format(rTabela("var_juros"), ocMONEY)
               .TextMatrix(.Rows - 1, 9) = Format(rTabela("var_total"), ocMONEY)
            'Else
            '   .TextMatrix(.Rows - 1, 7) = 0
            '   .TextMatrix(.Rows - 1, 8) = Format(0, "##,##0.00")
            '   If Not IsNull(RS!Valor) Then .TextMatrix(.Rows - 1, 9) = Format(RS!Valor, "##,##0.00")
            'End If
            
            If Not IsNull(rTabela("campo06")) Then
               .TextMatrix(.Rows - 1, 10) = Format(rTabela("campo06"), ocMONEY)
               .TextMatrix(.Rows - 1, 11) = Format(rTabela("campo07"), ocMONEY)
            Else
               .TextMatrix(.Rows - 1, 10) = Format(0, ocMONEY)
               .TextMatrix(.Rows - 1, 11) = Format(rTabela("var_total"), ocMONEY)
            End If
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 9
         .CellForeColor = &HC00000
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 10
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 11
         .CellForeColor = &H80&
         .CellFontBold = True
      Next
      
      'Deixar negrito quando vencido
      For i = 1 To .Rows - 1
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
      
      For i = 1 To .Rows - 1
         Grid_Parcelas.Row = i
         Set Grid_Parcelas.CellPicture = imgDesmarcada
         Grid_Parcelas.CellPictureAlignment = 4
      Next
      
      .Visible = True
      .Redraw = True
   End With
   
   lblSubTotal.Caption = Format(SomaGrid(Grid_Parcelas, 9), ocMONEY)
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
   
   txtTJuros.Text = Format((((var_VALOR * var_JurosDia) / 100) * var_Dias), ocMONEY)
End Sub

Private Sub LimparGrid_Haver()
   Dim i As Integer
   
   With Grid_Haver
      .Clear
      .Cols = 4
      .Rows = 2
      
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
      .Rows = .Rows + 1
      
      .Rows = .Rows - 1
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
      .Rows = 2
      
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
      .Rows = .Rows + 1
      
      i = i + 1
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 9
         .CellForeColor = &HC00000
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 10
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 11
         .CellForeColor = &H80&
         .CellFontBold = True
      Next
      
      'Grid_Parcelas.ColWidth(0) = 400
      'Grid_Parcelas.Rows = 11
      Grid_Parcelas.Col = 0
      
      For i = 1 To .Rows - 1
         Grid_Parcelas.Row = i
         Set Grid_Parcelas.CellPicture = imgDesmarcada
         Grid_Parcelas.CellPictureAlignment = 4
      Next
      
      .Visible = True
      .Redraw = True
   End With
   
   lblSubTotal.Caption = Format(SomaGrid(Grid_Parcelas, 9), ocMONEY)
   lblHaver.Caption = Format(SomaGrid(Grid_Parcelas, 10), ocMONEY)
   lblTotal.Caption = Format(SomaGrid(Grid_Parcelas, 11), ocMONEY)
End Sub

Private Sub LimparObjetos_GridParcelas()
   lblQuantParc.Caption = 0
   lblSubTotal.Caption = FormatNumber(0, 2)
   lblHaver.Caption = FormatNumber(0, 2)
   lblTotal.Caption = FormatNumber(0, 2)
End Sub

Private Sub LimparObjetos_Parcelas()
   chkMulta.Value = 0
   chkJuros.Value = 0
   txtCodOS.Text = ""
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
   mskDataHaver.Mask = ""
   mskDataHaver.Text = ""
   txtValorHaver.Text = ""
   lblClienteHaver.Caption = ""
   lblCliente.Caption = ""
   cboForma.Text = ""
   frmHaver.Enabled = False
   frmParcela.Visible = False
   frmPagamento.Visible = False
   cmdHabilitarQuitar.Visible = False
   cmdHabilitarHaver.Visible = False
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
End Sub

Private Sub Mostrar_Juros()
   Dim oCfg As ConfigItem
   
   Set oCfg = sysConfig("JUROS_DIA")
   txtJuros.Text = CCur(oCfg.Value)
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_Pedidos_Por_Codigo()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodPed.Text = "" Then Exit Sub
   
   sSQL = "SELECT cliente.*, pedidos.* FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
      "WHERE (cod_pedido = " & txtCodPed.Text & ");"
   
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then TxtCodCliente.Text = r("cod_cliente")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub MostrarGrid_Historico()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   Dim INDICE As String
   
    'indice
    If optPgto.Value = True Then
      INDICE = "parcelas.pagamento "
    ElseIf optVenc.Value = True Then
      INDICE = "parcelas.data "
    End If
      
   
   If TxtCodCliente.Text = "" Then Exit Sub
   
   sSQL = "SELECT parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
      "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, parcelas.pagamento AS campo06, " & _
      "parcelas.dias_atrazo AS var_dias, parcelas.juros AS var_juros, parcelas.valor_final AS var_vfinal, parcelas.status, " & _
      "pedidos.cod_pedido, pedidos.cod_cliente FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
      "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
      "WHERE (cliente.codigo = " & TxtCodCliente.Text & ") AND (parcelas.status = 1)  ORDER BY  " & INDICE
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   If TxtCodCliente.Text <> "1" Then FormatarGrid_Historico r
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   lblQuantHistorico.Caption = Format(totalRegistros, "00") & " parcela(s)"
   cmdMarcarTodasREATIVAR.Enabled = Grid_Historico.Rows > 1
End Sub

Private Sub MostrarGrid_Parcelas()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   Dim oCfg As ConfigItem
   Dim var_JurosDia As Double
   Dim var_tipojuros As Integer
   
   'MOSTRAR NO GRID AS PARCELAS À PAGAR
   If TxtCodCliente.Text = "" Then Exit Sub
   
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
   
   If TxtCodCliente.Text <> "" Then
      If optJurosSim.Value = True Then
      
         sSQL = "SELECT CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END AS var_atrazo, (((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END) AS var_juros, " & _
            "(parcelas.valor + ((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * (CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END)) AS var_total, cliente.codigo, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
            "parcelas.juros AS varvalorjuros, cliente.nome, parcelas.codigo AS codparcela, parcelas.*, parcelas.pagamento AS var_parcpgto, pedidos.pagamento AS var_tipopgto, " & _
            "(SELECT SUM(valor_haver) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)) AS campo06, ((parcelas.valor + (((" & var_CampoFuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * (CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END))) - (SELECT SUM(valor_haver) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) AS campo07, " & _
            "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago, pedidos.*, CASE parcelas.cod_pedido WHEN 0 THEN 'OS' ELSE 'PEDIDO' END AS var_tipo, " & _
            "CASE parcelas.cod_pedido WHEN 0 THEN parcelas.cod_os ELSE parcelas.cod_pedido END AS var_codigo " & _
            "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
            "WHERE (cliente.codigo = " & TxtCodCliente.Text & ") AND (parcelas.status = 0) ORDER BY parcelas.data, parcelas.codigo;"
     
     Else
         sSQL = "SELECT 0 AS var_atrazo, (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0) AS var_juros, " & _
            "(parcelas.valor + (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0)) AS var_total, cliente.codigo, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
            "parcelas.juros AS varvalorjuros, cliente.nome, parcelas.codigo AS codparcela, parcelas.*, parcelas.pagamento AS var_parcpgto, pedidos.pagamento AS var_tipopgto, " & _
            "(SELECT SUM(valor_haver) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)) AS campo06, ((parcelas.valor + (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0)) - (SELECT SUM(valor_haver) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) AS campo07, " & _
            "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago, pedidos.*, CASE parcelas.cod_pedido WHEN 0 THEN 'OS' ELSE 'PEDIDO' END AS var_tipo, " & _
            "CASE parcelas.cod_pedido WHEN 0 THEN parcelas.cod_os ELSE parcelas.cod_pedido END AS var_codigo " & _
            "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
            "WHERE (cliente.codigo = " & TxtCodCliente.Text & ") AND (parcelas.status = 0) ORDER BY parcelas.data, parcelas.codigo;"
      
      End If
   Else
      sSQL = "SELECT 0 AS var_atrazo, (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0) AS var_juros, " & _
         "(parcelas.valor + (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0)) AS var_total, cliente.codigo, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.juros AS varvalorjuros, cliente.nome, parcelas.codigo AS codparcela, parcelas.*, parcelas.pagamento AS var_parcpgto, pedidos.pagamento AS var_tipopgto, " & _
         "(SELECT SUM(valor_haver) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)) AS campo06, ((parcelas.valor + (((parcelas.valor * " & Replace(var_JurosDia, ",", ".") & ") / 100) * 0)) - (SELECT SUM(valor_haver) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) AS campo07, " & _
         "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago, pedidos.*, CASE parcelas.cod_pedido WHEN 0 THEN 'OS' ELSE 'PEDIDO' END AS var_tipo, " & _
         "CASE parcelas.cod_pedido WHEN 0 THEN parcelas.cod_os ELSE parcelas.cod_pedido END AS var_codiog " & _
         "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
         "WHERE false ORDER BY parcelas.data, parcelas.codigo;"
      
   End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   FormatarGrid_Parcelas r
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   lblQuantParc.Caption = Format(totalRegistros, "00") & " parcela(s)"
   cmdMarcarCheck.Enabled = Grid_Parcelas.Rows > 1
End Sub

Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To Grid.Rows - 1
      If IsNumeric(Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Somar_Parcelas_Selecionadas()
   Dim Total As Currency, SUBTOTAL As Currency, HAVER As Currency
   Dim i As Integer
   
   SUBTOTAL = 0
   HAVER = 0
   Total = 0
   
   With Grid_Parcelas
      For i = 1 To .Rows - 1
         .Col = 0
         .Row = i
         
         If .CellPicture = ImgMarcada Then
            .Col = 6
            SUBTOTAL = SUBTOTAL + .TextMatrix(.Row, 9)
            .Col = 10
            HAVER = HAVER + .TextMatrix(.Row, 10)
            .Col = 11
            Total = Total + .TextMatrix(.Row, 11)
         End If
      Next
      
      lblSubtotalSel.Caption = Format(SUBTOTAL, ocMONEY)
      lblHaverSel.Caption = Format(HAVER, ocMONEY)
      lblTotalSel.Caption = Format(Total, ocMONEY)
   End With
End Sub

Private Sub cboCliente_Change()
If chkCodPedido.Value = Unchecked Then CboCliente_LostFocus
End Sub

Private Sub cboCliente_Click()
If chkCodPedido.Value = Unchecked Then CboCliente_LostFocus
End Sub

Private Sub CboCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim itemAtual As String
   Dim codAtual As String
   
   itemAtual = CboCliente.Text
   codAtual = TxtCodCliente.Text
   CboCliente.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      CboCliente.AddItem r("nome")
      CboCliente.ItemData(CboCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   CboCliente.Text = itemAtual
   TxtCodCliente.Text = codAtual
   moCombo.AttachTo CboCliente
End Sub

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then cboCliente_Click
End Sub

Private Sub CboCliente_LostFocus()
   On Error GoTo TrataErro
   If CboCliente.Text = "" Then TxtCodCliente.Text = "": Exit Sub
   
   If chkCodPedido.Value = Unchecked Then TxtCodCliente = CboCliente.ItemData(CboCliente.ListIndex)
   If chkCodPedido.Value = Unchecked Then Exit Sub
   lblCliente.Caption = CboCliente.Text
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboForma_GotFocus()
   cboForma.Clear
   cboForma.AddItem "DINHEIRO"
   cboForma.AddItem "CHEQUE"
   cboForma.AddItem "CARTAO"
   cboForma.AddItem "DEPOSITO"
   cboForma.AddItem "TRANSFERENCIA"
   moCombo.AttachTo cboForma
End Sub

Private Sub cboFormaHaver_GotFocus()
   cboFormaHaver.Clear
   cboFormaHaver.AddItem "DINHEIRO"
   cboFormaHaver.AddItem "CHEQUE"
   cboFormaHaver.AddItem "CARTAO"
   cboFormaHaver.AddItem "DEPOSITO"
   cboFormaHaver.AddItem "TRANSFERENCIA"
   
   If cboFormaHaver.ListCount <> 0 Then cboFormaHaver.ListIndex = 0
   moCombo.AttachTo cboFormaHaver
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
Dim lNovoCod As Long

'MOSTRAR SE O CAIXA ESTÁ FECHADO
'Dim cStatus As Integer
'cStatus = Verificar_Caixa_Haver
'Select Case cStatus
'   Case -1
'      ShowMsg "Este caixa ainda não foi aberto.", vbExclamation
'      Exit Sub
'   Case 1
'      ShowMsg "O caixa está fechado!", vbExclamation
'      Exit Sub
'   End Select

If TxtCodCliente.Text = "" Or txtValorHaver.Text = "" Or mskDataHaver.Text = "" Then Exit Sub

'ADICIONAR O HAVER NA TABELA HAVER
AutoNumeracao_Haver

dbData.Execute "INSERT INTO parcelas_haver (codigo, cod_parcela, numero, vencimento, haver, valor_parcela, valor_haver, hora, forma_pgto, maquina) VALUES (" & _
   txtCodHaver.Text & ", " & txtCodParc.Text & ", " & txtNumParcela.Text & ", CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103), CONVERT(DATETIME, '" & Format(mskDataHaver.Text, ocDATA) & "', 103), " & _
   Replace(CCur(txtValor.Text), ",", ".") & ", " & Replace(CCur(txtValorHaver.Text), ",", ".") & ", '" & Format(lblHora, ocHRMN) & "', '" & cboFormaHaver.Text & "','" & StatusBar1.Panels(2).Text & "');"

'ADICIONAR NA TABELA CAIXA_ENTRADA
'lNovoCod = AutoNumeracao_Caixa

'dbData.Execute "INSERT INTO caixa_entrada (codigo, descricao, data, valor, setor, hora, cod_haver, forma_pgto, maquina) VALUES (" & _
'   lNovoCod & ", '" & "HAVER: " & cboCliente.Text & "', '" & Format(mskDataHaver.Text, ocDATA_EUA) & "', " & Replace(CCur(txtValorHaver.Text), ",", ".") & ", '" & _
'   IIf(txtCodOS.Text = "", "LOJA", "OFICINA") & "', '" & Format(lblHora, ocHRMN) & "', " & txtCodHaver.Text & ", '" & cboFormaHaver.Text & "', '" & StatusBar1.Panels(2).Text & "')"

'MARCAR O CAMPO HAVER DA TABELA PARCELAS
dbData.Execute "UPDATE parcelas SET haver = 1 WHERE (codigo = " & txtCodParc.Text & ");"

'IMPRIMIR
If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
   Imprimir = True
Else
   Imprimir = False
End If

If Imprimir = True Then
   Me.Hide
   Me.Show
   
   With REL_Recibo
      .txtCliente.Caption = UCase(CboCliente.Text)
     .txtValor.Caption = UCase(NumeroExtenso(txtValorHaver.Text, True))
      .txthead.Caption = "R$ " & Format(txtValorHaver.Text, ocMONEY)

      'If txtCodOS.Text <> "" And txtCodPedido.Text = "" Then
      '   .txtProveniente.Caption = "Haver da " & txtNumParcela.Text & "ª parcela da OS Nº " & Format(txtCodOS.Text, "000000")
      'ElseIf txtCodOS.Text = "" And txtCodPedido.Text <> "" Then
         .txtProveniente.Caption = "Haver da " & txtNumParcela.Text & "ª parcela do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
      'End If

      .txtData.Caption = "Uruçuí-PI, " & Day(mskDataHaver) & " de " & MonthName(Month(mskDataHaver)) & " de " & Year(mskDataHaver)
      .Relatorio.NumeroRegistros = 1
      .Relatorio.Ativar
   End With

   Unload REL_Recibo
End If

MostrarGrid_Parcelas
Calcular_Valores
MostrarGrid_Haver
'LimparObjetos_Parcelas_Haver
mskDataHaver.Text = Format(Date, "dd/mm/yy")
txtValorHaver.Text = ""
cboFormaHaver.Text = ""
Somar_Parcelas_Selecionadas
txtValorHaver.SetFocus
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
   If CboCliente.Text = "" And txtNumParcela.Text = "" And txtValor.Text = "" Then Exit Sub
   If txtCodPedido.Text = "" And txtCodOS.Text = "" Then Exit Sub
   
   If txtCodOS.Text <> "" And txtCodPedido.Text = "" Then
      If ShowMsg("Deseja alterar a parcela '" & txtNumParcela.Text & "' da OS No.'" & txtCodOS.Text & "' ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   ElseIf txtCodOS.Text = "" And txtCodPedido.Text <> "" Then
      If ShowMsg("Deseja alterar a parcela '" & txtNumParcela.Text & "' do Pedido No. '" & txtCodPedido.Text & "' ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   End If
   
   If txtCodOS.Text <> "" And txtCodPedido.Text = "" Then
      dbData.Execute "UPDATE parcelas SET valor = " & Replace(CCur(txtValor.Text), ",", ".") & ", data = CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103) WHERE (cod_os = " & txtCodOS.Text & ") AND (numero = " & txtNumParcela.Text & ");"
   ElseIf txtCodOS.Text = "" And txtCodPedido.Text <> "" Then
      dbData.Execute "UPDATE parcelas SET valor = " & Replace(CCur(txtValor.Text), ",", ".") & ", data = CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103) WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = " & txtNumParcela.Text & ");"
   End If
   
   MostrarGrid_Parcelas
   Calcular_Valores
   LimparGrid_Haver
   LimparObjetos_Parcelas
   CboCliente.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    LimparObjetos_Parcelas
End Sub

Private Sub cmdHabilitarHaver_Click()
   Dim f As Integer
   
   frmHaver.Enabled = True
   mskDataHaver.Text = Format(Date, "dd/mm/yy")
   
   For f = 0 To Grid_Parcelas.Rows - 1
      Grid_Parcelas.Row = f
      Grid_Parcelas.Col = 0
      
      If Grid_Parcelas.CellPicture = ImgMarcada Then
         txtCodParc.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 1))
         txtCodOS.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 2))
         txtCodPedido.Text = Format((Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 3)), "000000")
         txtNumParcela.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 4))
         mskData.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 5))
         txtValor.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 6))
         MostrarGrid_Haver
      End If
   Next
   
   SSTab1.Tab = 1
End Sub

Private Sub cmdHabilitarQuitar_Click()
   Dim f As Integer
   
   For f = 0 To Grid_Parcelas.Rows - 1
      Grid_Parcelas.Row = f
      Grid_Parcelas.Col = 0
      
      If Grid_Parcelas.CellPicture = ImgMarcada Then
         txtCodParc.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 1))
         txtCodOS.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 2))
         txtCodPedido.Text = Format((Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 3)), "000000")
         txtNumParcela.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 4))
         mskData.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 5))
         txtValor.Text = (Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 6))
         
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
   'cmdAlterar.Visible = True
   'cmdExcluir.Visible = True
   cboForma.SetFocus
End Sub

Private Sub cmdHabilitarREATIVAR_Click()
   Dim f As Integer
   
   For f = 0 To Grid_Historico.Rows - 1
      Grid_Historico.Row = f
      Grid_Historico.Col = 0
      
      If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then
         frmReativar.Visible = True
         cmdReativar.Visible = True
         txtCONcodParc.Text = (Grid_Historico.TextMatrix(Grid_Historico.Row, 1))
         txtConCodOS.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 2)), "000000")
         txtConCodPedido.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 3)), "000000")
         txtConNumParcela.Text = (Grid_Historico.TextMatrix(Grid_Historico.Row, 4))
         mskConData.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 5)), "dd/mm/yy")
         mskConPgto.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 10)), "dd/mm/yy")
         txtCONValor.Text = Format((Grid_Historico.TextMatrix(Grid_Historico.Row, 6)), "##,##0.00")
      End If
   Next
End Sub

Private Sub cmdImprimir_Click()
   'colocar o nome da maquina na barra de status
   'Dim var_Impressora As String
   'var_Impressora = ReadINI(App.Path & "\config.ini", "DADOS_IMPRESSORA", "impressora")
   
   If cmdHabilitarQuitar.Visible = True Then
      If txtCodPedido.Text = "" Then
         ShowMsg "O código do PEDIDO está em branco !!", vbExclamation
         Exit Sub
      End If
   End If
   
   Me.Hide
   Me.Show
   
   With REL_Recibo
      .txtCliente.Caption = UCase(CboCliente.Text)
      
      If cmdHabilitarQuitar.Visible = True Then
         .txtValor.Caption = UCase(NumeroExtenso(txtTotal.Text, True))
         .txthead.Caption = "R$ " & Format(txtTotal.Text, "##,##0.00")
         .txtProveniente.Caption = "Pagamento da " & txtNumParcela.Text & "ª parcela do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
         .txtData.Caption = "Uruçuí-PI, " & Day(mskPagamento) & " de " & MonthName(Month(mskPagamento)) & " de " & Year(mskPagamento)
      Else
         Dim var_Parc As String
         Dim f As Integer

         var_Parc = ""
         With Grid_Parcelas
            For f = 1 To .Rows - 1
               .Col = 0
               .Row = f

               If .CellPicture = ImgMarcada Then
                  If f = 1 Then
                     var_Parc = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
                  ElseIf Right(var_Parc, 8) = Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00") Then
                     MsgBox "Tratar Repetido"
                  Else
                     var_Parc = var_Parc & ", " & Format(.TextMatrix(.Row, 3), "00000") & "/" & Format(.TextMatrix(.Row, 4), "00")
                     If f = .Rows - 1 Then Exit For
                  End If
               End If
            Next f
         End With

         .txtValor.Caption = UCase(NumeroExtenso(lblTotalSel.Caption, True))
         .txthead.Caption = "R$ " & Format(lblTotalSel.Caption, "##,##0.00")
         .txtProveniente.Caption = "PEDIDO(S): " & var_Parc
         .txtData.Caption = "Uruçuí-PI, " & Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date)
      End If

      .Relatorio.NumeroRegistros = 1
      '.Relatorio.NomeImpressora = var_Impressora
      .Relatorio.Ativar
   End With

   Unload REL_Recibo
End Sub

Private Sub cmdMarcarCheck_Click()
   If cmdMarcarCheck.Caption = "MARCAR TODAS" Then
      OP = MarcarTodos
      AcaoGrid
      cmdQuitarTodas.Visible = True
      cmdMarcarCheck.Caption = "DESMARCAR TODAS"
   Else
      OP = DesmarcarTodos
      AcaoGrid
      cmdMarcarCheck.Caption = "MARCAR TODAS"
      cmdQuitarTodas.Visible = False
   End If
   
   OP = Contar
   AcaoGrid
   Somar_Parcelas_Selecionadas
End Sub

Private Sub cmdMarcarTodasREATIVAR_Click()
   If cmdMarcarTodasREATIVAR.Caption = "MARCAR TODAS" Then
      OP = MarcarTodos
      AcaoGridREATIVAR
      cmdTodasREATIVAR.Enabled = True
      cmdMarcarTodasREATIVAR.Caption = "DESMARCAR TODAS"
   Else
      OP = DesmarcarTodos
      AcaoGridREATIVAR
      cmdMarcarTodasREATIVAR.Caption = "MARCAR TODAS"
      cmdTodasREATIVAR.Enabled = False
   End If
   
   OP = Contar
   AcaoGridREATIVAR
End Sub

Private Sub cmdMostrarProdutos_Click()
   Dim f As Integer
   
   For f = 0 To Grid_Parcelas.Rows - 1
      Grid_Parcelas.Row = f
      Grid_Parcelas.Col = 0
      
      If Grid_Parcelas.CellPicture = ImgMarcada Then
         Parcelas_Consulta_Produtos.loadPedidos Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 3), Grid_Parcelas.TextMatrix(Grid_Parcelas.Row, 2)
         Parcelas_Consulta_Produtos.Show 1
      End If
   Next
End Sub

Private Sub cmdMostrarProdutosREATIVAR_Click()
   Dim f As Integer
   
   For f = 0 To Grid_Historico.Rows - 1
      Grid_Historico.Row = f
      Grid_Historico.Col = 0
      
      If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then
         Parcelas_Consulta_Produtos.loadPedidos Grid_Historico.TextMatrix(Grid_Historico.Row, 3), ""
         Parcelas_Consulta_Produtos.Show 1
      End If
   Next
End Sub

Private Sub cmdQuitarTodas_Click()
   Dim f As Integer
   Dim sSQL As String
   
   If cboForma.Text = "" Then
      ShowMsg "Faltou Escolher a forma de pagamento!", vbInformation
      cboForma.SetFocus
      Exit Sub
   End If
   
   'MOSTRAR SE O CAIXA ESTÁ FECHADO
   Verificar_Caixa_Baixa
   If CAIXA_FECHADO_BAIXA = True Then Exit Sub
   
   With Grid_Parcelas
      For f = 1 To .Rows - 1
         .Col = 0
         .Row = f
         If .CellPicture = ImgMarcada Then
            sSQL = "UPDATE parcelas SET status = 1, valor_final = " & Replace(CCur(.TextMatrix(.Row, 11)), ",", ".") & ", " & _
               "pagamento = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), hora = '" & Format(Now, ocHRMN) & "', " & _
               "juros = " & Replace(CCur(.TextMatrix(.Row, 8)), ",", ".") & ", dias_atrazo = " & .TextMatrix(.Row, 7) & ", " & _
               "forma_pgto = '" & cboForma.Text & "', maquina = '" & StatusBar1.Panels(2).Text & "' " & _
               "WHERE (cod_pedido = " & .TextMatrix(.Row, 3) & ") AND (numero = " & .TextMatrix(.Row, 4) & ");"
            
            dbData.Execute sSQL
         End If
      Next f
   End With
   
   If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
      Imprimir = True
   Else
      Imprimir = False
   End If
   
   If Imprimir = True Then
      cmdImprimir_Click
   End If
   
   MostrarGrid_Haver
   MostrarGrid_Parcelas
   MostrarGrid_Historico
   OP = Contar
   AcaoGrid
   Somar_Parcelas_Selecionadas
   cboForma.Text = ""
End Sub

Private Sub cmdReativar_Click()
   If CboCliente.Text = "" Or txtConNumParcela.Text = "" Or txtCONValor.Text = "" Then Exit Sub
   If txtConCodPedido.Text = "" And txtConCodOS.Text = "" Then Exit Sub
   
'VERIFICAR O STATUS DO CAIXA
'Dim cStatus As Integer
'cStatus = Verificar_Caixa_Reativar
'Select Case cStatus
'   Case -1
'      ShowMsg "Este caixa ainda não foi aberto.", vbExclamation
'      Exit Sub
'   Case 1
'      ShowMsg "O caixa está fechado!", vbExclamation
'      Exit Sub
'   End Select
   
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
      dbData.Execute "UPDATE parcelas SET status = 0, valor_final = 0, pagamento = Null, forma_pgto = '', maquina = '' WHERE (cod_pedido = " & txtConCodPedido.Text & ") AND (numero = " & txtConNumParcela.Text & ");"
      dbData.Execute "UPDATE cheque SET status = 0 WHERE (cod_pedido = " & txtConCodPedido.Text & ") AND (parcela = " & txtConNumParcela.Text & ");"
      dbData.Execute "UPDATE parcelas_haver SET status = 0 WHERE (codigo = " & txtCONcodParc.Text & ");"
   'End If
   
   MostrarGrid_Parcelas
   Calcular_Valores
   MostrarGrid_Historico
   LimparObjetos_Historico
   
   OP = Contar
   AcaoGridREATIVAR
End Sub

Private Sub cmdRemoverHaver_Click()
   On Error GoTo erro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Verificar_Caixa_Haver
   If CAIXA_FECHADO_HAVER = True Then Exit Sub
   
   If Grid_Haver.TextMatrix(Grid_Haver.Row, 1) = "" Then GoSub erro
   If ShowMsg("Deseja remover o haver de: " & Grid_Haver.TextMatrix(Grid_Haver.Row, 3) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   'If txtCodOS.Text <> "" And txtCodPedido.Text = "" Then
   '   execSQL "DELETE FROM PARCELAS_HAVER WHERE CODIGO =" & Grid_Haver.TextMatrix(Grid_Haver.Row, 1) & ""
   '   execSQL "DELETE FROM CAIXA_ENTRADA WHERE COD_HAVER = " & Grid_Haver.TextMatrix(Grid_Haver.Row, 1) & ""
   'ElseIf txtCodOS.Text = "" And txtCodPedido.Text <> "" Then
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
   'VERIFICAR O STATUS DO CAIXA
   'Dim cStatus As Integer
   'cStatus = Verificar_Caixa_Baixa
   'Select Case cStatus
   '   Case -1
   '      ShowMsg "Este caixa ainda não foi aberto.", vbExclamation
   '      Exit Sub
   '   Case 1
   '      ShowMsg "O caixa está fechado!", vbExclamation
   '      Exit Sub
   '   End Select
      
   If CboCliente.Text = "" And txtNumParcela.Text = "" And txtValor.Text = "" Then Exit Sub
   If txtCodPedido.Text = "" And txtCodOS.Text = "" Then Exit Sub
   
   If cboForma.Text = "" Then
      ShowMsg "Faltou Escolher a forma de pagamento!", vbInformation
      cboForma.SetFocus
      Exit Sub
   End If
   
   If txtCodOS.Text <> "" And txtCodPedido.Text = "" Then
      If ShowMsg("Deseja quitar a parcela '" & txtNumParcela.Text & "' da OS No. '" & txtCodOS.Text & "' ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   ElseIf txtCodOS.Text = "" And txtCodPedido.Text <> "" Then
      If ShowMsg("Deseja quitar a parcela '" & txtNumParcela.Text & "' do Pedido No. '" & txtCodPedido.Text & "' ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   End If
   
   If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo) = vbYes Then
      Imprimir = True
   Else
      Imprimir = False
   End If

   'If txtCodOS.Text <> "" And txtCodPedido.Text = "" Then
   '   execSQL "UPDATE PARCELAS SET STATUS = 1 , VALOR_FINAL = '" & Format(txtTotal.Text, "##,##0.00") & "' , PAGAMENTO = #" & Format(mskPagamento.Text, "mm/dd/yyyy") & "#, HORA = #" & Format(lblHora.Caption, "hh:mm") & "#, FORMA_PGTO =  '" & cboForma.Text & "' WHERE COD_OS = " & txtCodOS.Text & " AND NUMERO = " & txtNumParcela.Text
   '   execSQL "UPDATE CHEQUE SET STATUS = 1 WHERE COD_OS = " & txtCodOS.Text & " AND PARCELA = " & txtNumParcela.Text
   '   execSQL "UPDATE PARCELAS_HAVER SET STATUS = 1 WHERE (COD_PARCELA = " & txtCodParc.Text & ")"
   'ElseIf txtCodOS.Text = "" And txtCodPedido.Text <> "" Then
      dbData.Execute "UPDATE parcelas SET status = 1, valor_final = " & Replace(CCur(txtTotal.Text), ",", ".") & ", pagamento = CONVERT(DATETIME, '" & Format(mskPagamento.Text, ocDATA) & "', 103), hora = '" & Format(lblHora.Caption, ocHRMN) & "', juros = " & Replace(CCur(txtTJuros.Text), ",", ".") & ", dias_atrazo = " & txtDias.Text & ", forma_pgto = '" & cboForma.Text & "', maquina = '" & StatusBar1.Panels(2).Text & "' WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = " & txtNumParcela.Text & ");"
      dbData.Execute "UPDATE cheque SET status = 1 WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (parcela = " & txtNumParcela.Text & ");"
      dbData.Execute "UPDATE parcelas_haver SET status = 1 WHERE (cod_parcela = " & txtCodParc.Text & ");"
   'End If
   
   If Imprimir = True Then
      cmdImprimir_Click
   End If
   
   MostrarGrid_Haver
   MostrarGrid_Parcelas
   MostrarGrid_Historico
   LimparObjetos_Parcelas
   Somar_Parcelas_Selecionadas
End Sub

Private Sub cmdTodasREATIVAR_Click()
   Dim f As Integer
   
   Verificar_Caixa_Reativar
   If CAIXA_FECHADO_REATIVAR = True Then Exit Sub
   
   With Grid_Historico
      For f = 1 To .Rows - 1
         .Col = 0
         .Row = f
         
         If .CellPicture = ImgMarcadaPAGAS Then
            dbData.Execute "UPDATE parcelas SET status = 0, forma_pgto = ' ', pagamento = Null WHERE (cod_pedido = " & .TextMatrix(.Row, 3) & ") AND (numero = " & .TextMatrix(.Row, 4) & ");"
         End If
      Next
   End With
   
   MostrarGrid_Haver
   MostrarGrid_Parcelas
   MostrarGrid_Historico
   OP = Contar
   AcaoGridREATIVAR
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0

'colocar o nome da maquina na barra de status
Dim var_Maquina As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
Set oIni = Nothing

StatusBar1.Panels(2).Text = var_Maquina
StatusBar1.Panels(4).Text = Format(Date, "dd/mm/yy")

LimparGrid_Parcelas
LimparGrid_Haver
LimparGrid_Historico
Set moCombo = New cComboHelper
End Sub

Sub AcaoGridREATIVAR()
   Dim i As Integer
   Dim var_Contador As Integer
   
   Grid_Historico.Col = 0
   
   For i = 1 To Grid_Historico.Rows - 1
      Grid_Historico.Row = i
      If OP = MarcarTodos Then Set Grid_Historico.CellPicture = ImgMarcadaPAGAS
      If OP = DesmarcarTodos Then Set Grid_Historico.CellPicture = imgDesmarcadaPAGAS
      If OP = Contar Then
         If Grid_Historico.CellPicture = ImgMarcadaPAGAS Then var_Contador = var_Contador + 1
      End If
   Next
   
   If var_Contador = 1 Then
      frmReativar.Visible = False
      cmdMostrarProdutosREATIVAR.Enabled = True
      cmdHabilitarREATIVAR.Enabled = True
      cmdTodasREATIVAR.Enabled = False
   ElseIf var_Contador > 1 Then
      frmReativar.Visible = False
      cmdMostrarProdutosREATIVAR.Enabled = True
      cmdHabilitarREATIVAR.Enabled = False
      cmdTodasREATIVAR.Enabled = True
   ElseIf var_Contador = 0 Then
      frmReativar.Visible = False
      cmdMostrarProdutosREATIVAR.Enabled = False
      cmdHabilitarREATIVAR.Enabled = False
      cmdTodasREATIVAR.Enabled = False
   End If
   
   'If OP = Contar Then ShowMsg "Qtde de itens selecionados: " & var_Contador, , "Contador"
End Sub

Sub AcaoGrid()
   Dim i As Integer
   Dim var_Contador As Integer
   
   Grid_Parcelas.Col = 0
   
   For i = 1 To Grid_Parcelas.Rows - 1
      Grid_Parcelas.Row = i
      If OP = MarcarTodos Then Set Grid_Parcelas.CellPicture = ImgMarcada
      If OP = DesmarcarTodos Then Set Grid_Parcelas.CellPicture = imgDesmarcada
      If OP = Contar Then
         If Grid_Parcelas.CellPicture = ImgMarcada Then var_Contador = var_Contador + 1
      End If
   Next
   
   If var_Contador = 1 Then
      frmPagamento.Visible = False
      'lblFormaPgto.Visible = False
      'cboForma.Visible = False
      cmdMostrarProdutos.Visible = True
      cmdHabilitarQuitar.Visible = True
      cmdHabilitarHaver.Visible = True
      cmdQuitarTodas.Visible = False
      'mostrar dados do pagamento
      'frmPagamento.Visible = True
      lblFormaPgto.Visible = True
      cboForma.Visible = True
      lblPgto.Visible = True
      mskPagamento.Visible = True
      lblAtrazo.Visible = True
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
   ElseIf var_Contador > 1 Then
      cmdQuitarTodas.Visible = True
      cmdMostrarProdutos.Visible = False
      cmdHabilitarQuitar.Visible = False
      cmdHabilitarHaver.Visible = False
      frmParcela.Visible = False
      'frmPagamento.Visible = False
      cmdSalvar.Visible = False
      cmdCancelar.Visible = False
      cmdAlterar.Visible = False
      cmdExcluir.Visible = False
      'mostrar dados do pagamento
      frmPagamento.Visible = True
      lblFormaPgto.Visible = True
      cboForma.Visible = True
      lblPgto.Visible = False
      mskPagamento.Visible = False
      lblAtrazo.Visible = False
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
   ElseIf var_Contador = 0 Then
      cmdMostrarProdutos.Visible = False
      cmdHabilitarQuitar.Visible = False
      cmdHabilitarHaver.Visible = False
      cmdQuitarTodas.Visible = False
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
      lblAtrazo.Visible = False
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
   End If
   
   'If OP = Contar Then MsgBox "Qtde de itens selecionados: " & var_Contador, , "Contador"
End Sub

Private Function Verificar_Caixa_Reativar() As Integer
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim cxaStatus As Integer
  
   cxaStatus = -1   'Não foi aberto
   'If cmdAlterar.Enabled = True Then
      sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskConPgto.FormattedText, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   'Else
      'sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = '" & Format(StatusBar1.Panels(4), ocDATA_EUA) & "') AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
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
      sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskDataHaver.FormattedText, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   'Else
      'sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = '" & Format(StatusBar1.Panels(4), ocDATA_EUA) & "') AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
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
      sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskPagamento.FormattedText, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   'Else
      'sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = '" & Format(StatusBar1.Panels(4), ocDATA_EUA) & "') AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
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
   Dim Total As Currency, HAVER As Currency, TOTAL_GERAL As Currency, JUROS As Currency, var_MULTAS As Currency
   
   If txtMulta.Text = "" Then txtMulta.Text = FormatNumber(0, 2)
   If txtJuros.Text = "" Then txtJuros.Text = FormatNumber("0,33", 2)
   
   If txtValor.Text = "" Then Total = 0 Else Total = txtValor
   If txtTotalHaver.Text = "" Then HAVER = 0 Else HAVER = txtTotalHaver
   
   If chkJuros.Value = 1 Then JUROS = txtTJuros Else JUROS = 0
   If chkMulta.Value = 1 Then var_MULTAS = txtMulta Else var_MULTAS = 0
   
   TOTAL_GERAL = var_MULTAS + JUROS + Total - HAVER
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
   Calcular_Juros
   Calcular_Valores
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Public Sub loadParcelas(Cliente As String)
   CboCliente.Text = Cliente
End Sub

Private Sub Grid_Historico_Click()
   'marcar a parcela
   If Grid_Historico.Col <> 0 Then Exit Sub
   
   If Grid_Historico.CellPicture = imgDesmarcadaPAGAS Then
      Set Grid_Historico.CellPicture = ImgMarcadaPAGAS
   Else
      Set Grid_Historico.CellPicture = imgDesmarcadaPAGAS
   End If
   
   OP = Contar
   AcaoGridREATIVAR
End Sub

Private Sub Grid_Parcelas_Click()
   'marcar a parcela
   If Grid_Parcelas.Col <> 0 Then Exit Sub
   
   If Grid_Parcelas.CellPicture = imgDesmarcada Then
      Set Grid_Parcelas.CellPicture = ImgMarcada
   Else
      Set Grid_Parcelas.CellPicture = imgDesmarcada
   End If
   
   OP = Contar
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
         Calcular_Dias
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

Private Sub optVenc_Click()
MostrarGrid_Historico
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      Exit Sub
   ElseIf SSTab1.Tab = 1 Then
      If frmHaver.Enabled = True Then txtValorHaver.SetFocus
   ElseIf SSTab1.Tab = 2 Then
      Exit Sub
   End If
End Sub

Private Sub Timer1_Timer()
   lblHora.Caption = Format(Time, "hh:mm")
End Sub

Private Sub TxtCodCliente_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If TxtCodCliente.Text = "" Then Exit Sub
   LimparObjetos_Parcelas
   
   If chkCodPedido.Value = Checked Then
      CboCliente.Text = ""
      sSQL = "SELECT * FROM cliente WHERE (codigo= " & TxtCodCliente.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then CboCliente.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
   
   If TxtCodCliente.Text <> "" Then
      MostrarGrid_Parcelas
      MostrarGrid_Historico
      lblClienteHaver.Caption = CboCliente.Text
      lblCliente.Caption = CboCliente.Text
   Else
      LimparObjetos_Parcelas
      LimparGrid_Parcelas
      LimparObjetos_GridParcelas
   End If
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
