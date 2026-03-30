VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Receber_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTAS À RECEBER"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "Receber_Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   11325
      TabIndex        =   34
      Top             =   60
      Width           =   11355
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   10020
         TabIndex        =   87
         Top             =   360
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   420
         Picture         =   "Receber_Cadastro.frx":23D2
         Top             =   180
         Width           =   600
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTAS RETROATIVAS"
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
         TabIndex        =   35
         Top             =   300
         Width           =   3630
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   780
      Top             =   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8355
      Left            =   60
      TabIndex        =   19
      Top             =   1140
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   14737
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "Receber_Cadastro.frx":2BBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalvar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExcluir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCadastrarCliente"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Receber_Cadastro.frx":2BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label23"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(4)=   "cmdImprimirConsulta"
      Tab(1).Control(5)=   "GridConsulta"
      Tab(1).Control(6)=   "txtCONHaver"
      Tab(1).Control(7)=   "txtCONValor"
      Tab(1).Control(8)=   "Picture3"
      Tab(1).Control(9)=   "txtCONquant"
      Tab(1).Control(10)=   "txtCONtotal"
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtCONtotal 
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
         Height          =   285
         Left            =   -65160
         TabIndex        =   41
         Top             =   7980
         Width           =   1335
      End
      Begin VB.TextBox txtCONquant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65160
         TabIndex        =   40
         Top             =   7080
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   -74940
         ScaleHeight     =   885
         ScaleWidth      =   11085
         TabIndex        =   39
         Top             =   420
         Width           =   11115
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Critérios"
            Height          =   675
            Left            =   4380
            TabIndex        =   58
            Top             =   60
            Width           =   4935
            Begin ChamaleonBtn.chameleonButton cmdFinal 
               Height          =   315
               Left            =   4260
               TabIndex        =   64
               Tag             =   "Calendario"
               Top             =   240
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
               MICON           =   "Receber_Cadastro.frx":2BF6
               PICN            =   "Receber_Cadastro.frx":2C12
               PICH            =   "Receber_Cadastro.frx":4F65
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdInicial 
               Height          =   315
               Left            =   2040
               TabIndex        =   63
               Tag             =   "Calendario"
               Top             =   240
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
               MICON           =   "Receber_Cadastro.frx":72B8
               PICN            =   "Receber_Cadastro.frx":72D4
               PICH            =   "Receber_Cadastro.frx":9627
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboNome 
               Height          =   315
               Left            =   720
               TabIndex        =   69
               Top             =   240
               Visible         =   0   'False
               Width           =   3855
            End
            Begin VB.TextBox txtCodClienteCons 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4080
               TabIndex        =   68
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   3120
               Sorted          =   -1  'True
               TabIndex        =   66
               Top             =   240
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMES 
               Height          =   315
               ItemData        =   "Receber_Cadastro.frx":B97A
               Left            =   1320
               List            =   "Receber_Cadastro.frx":B97C
               TabIndex        =   65
               Top             =   240
               Visible         =   0   'False
               Width           =   1755
            End
            Begin MSMask.MaskEdBox Mask2 
               Height          =   315
               Left            =   3300
               TabIndex        =   59
               Top             =   240
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Mask1 
               Height          =   315
               Left            =   1080
               TabIndex        =   60
               Top             =   240
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label lblCONnome 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nome:"
               Height          =   195
               Left            =   180
               TabIndex        =   70
               Top             =   240
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Label lblCONmes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E&scolha o mês:"
               Height          =   195
               Left            =   180
               TabIndex        =   67
               Top             =   300
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label lblCONint1 
               AutoSize        =   -1  'True
               Caption         =   "Da&ta Inicial:"
               Height          =   195
               Left            =   180
               TabIndex        =   62
               Top             =   300
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label lblCONint2 
               AutoSize        =   -1  'True
               Caption         =   "Data &Final:"
               Height          =   195
               Left            =   2460
               TabIndex        =   61
               Top             =   300
               Visible         =   0   'False
               Width           =   765
            End
         End
         Begin VB.Frame frm 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Status"
            Height          =   675
            Left            =   2820
            TabIndex        =   56
            Top             =   60
            Width           =   1515
            Begin VB.ComboBox cboCONStatus 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   60
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ordem"
            Height          =   675
            Left            =   1560
            TabIndex        =   54
            Top             =   60
            Width           =   1215
            Begin VB.ComboBox cboOrdem 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   60
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Filtro"
            Height          =   675
            Left            =   60
            TabIndex        =   52
            Top             =   60
            Width           =   1455
            Begin VB.ComboBox cboFiltro 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   60
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   240
               Width           =   1335
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirConsulta 
            Height          =   615
            Left            =   9480
            TabIndex        =   48
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
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
            BCOL            =   32768
            BCOLO           =   32768
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Receber_Cadastro.frx":B97E
            PICN            =   "Receber_Cadastro.frx":B99A
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
      Begin VB.TextBox txtCONValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65160
         TabIndex        =   38
         Top             =   7380
         Width           =   1335
      End
      Begin VB.TextBox txtCONHaver 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65160
         TabIndex        =   37
         Top             =   7680
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   7635
         Left            =   120
         ScaleHeight     =   7605
         ScaleWidth      =   9465
         TabIndex        =   20
         Top             =   480
         Width           =   9495
         Begin VB.Frame frmCliente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dados"
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
            Height          =   975
            Left            =   60
            TabIndex        =   31
            Top             =   60
            Width           =   9375
            Begin ChamaleonBtn.chameleonButton cmdCal1 
               Height          =   315
               Left            =   1080
               TabIndex        =   50
               Tag             =   "Calendario"
               Top             =   480
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
               MICON           =   "Receber_Cadastro.frx":C274
               PICN            =   "Receber_Cadastro.frx":C290
               PICH            =   "Receber_Cadastro.frx":E5E3
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox txtCliente 
               Height          =   315
               Left            =   1440
               TabIndex        =   2
               Top             =   480
               Width           =   7035
            End
            Begin VB.TextBox txtCodCliente 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7260
               TabIndex        =   32
               Top             =   180
               Visible         =   0   'False
               Width           =   615
            End
            Begin MSMask.MaskEdBox mskCompra 
               Height          =   315
               Left            =   120
               TabIndex        =   1
               Top             =   480
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label lblCompra 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Compra"
               Height          =   195
               Left            =   180
               TabIndex        =   47
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
               Height          =   195
               Left            =   1440
               TabIndex        =   33
               Top             =   240
               Width           =   525
            End
         End
         Begin VB.Frame frmParcelas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Parcelas"
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
            Height          =   1515
            Left            =   60
            TabIndex        =   25
            Top             =   4380
            Width           =   9375
            Begin VB.Frame Frame6 
               BackColor       =   &H00E0E0E0&
               Height          =   495
               Left            =   3540
               TabIndex        =   83
               Top             =   180
               Width           =   5775
               Begin VB.OptionButton optPorDatas 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Por Datas"
                  Height          =   195
                  Left            =   2100
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   85
                  TabStop         =   0   'False
                  Top             =   210
                  Width           =   1095
               End
               Begin VB.OptionButton optPorQuant 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Por Quant."
                  Height          =   195
                  Left            =   900
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   84
                  TabStop         =   0   'False
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1155
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Datas:"
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
                  TabIndex        =   86
                  Top             =   180
                  Width           =   570
               End
            End
            Begin ChamaleonBtn.chameleonButton chameleonButton2 
               Height          =   315
               Left            =   5460
               TabIndex        =   82
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
               MICON           =   "Receber_Cadastro.frx":10936
               PICN            =   "Receber_Cadastro.frx":10952
               PICH            =   "Receber_Cadastro.frx":12CA5
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboQuantParc 
               Height          =   315
               Left            =   1380
               TabIndex        =   11
               Top             =   1020
               Width           =   735
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E0E0E0&
               Height          =   495
               Left            =   120
               TabIndex        =   74
               Top             =   180
               Width           =   3375
               Begin VB.OptionButton optMultiplica 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Multiplica"
                  Height          =   195
                  Left            =   1080
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   77
                  TabStop         =   0   'False
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1035
               End
               Begin VB.OptionButton optDivide 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Divide"
                  Height          =   195
                  Left            =   2100
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   76
                  TabStop         =   0   'False
                  Top             =   210
                  Width           =   855
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Parcelas:"
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
                  TabIndex        =   78
                  Top             =   180
                  Width           =   810
               End
            End
            Begin ChamaleonBtn.chameleonButton chameleonButton1 
               Height          =   315
               Left            =   4260
               TabIndex        =   51
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
               MICON           =   "Receber_Cadastro.frx":14FF8
               PICN            =   "Receber_Cadastro.frx":15014
               PICH            =   "Receber_Cadastro.frx":17367
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
               Left            =   3300
               TabIndex        =   13
               Top             =   1020
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtParc 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   1020
               Width           =   1095
            End
            Begin VB.TextBox txtTotal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   10
               Top             =   1020
               Width           =   1215
            End
            Begin MSMask.MaskEdBox mskFinal 
               Height          =   315
               Left            =   4560
               TabIndex        =   14
               Top             =   1020
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data Final"
               Height          =   195
               Left            =   4620
               TabIndex        =   73
               Top             =   780
               Width           =   720
            End
            Begin VB.Label lbl5 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data Inicial"
               Height          =   195
               Left            =   3360
               TabIndex        =   29
               Top             =   780
               Width           =   795
            End
            Begin VB.Label lbl3 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Parc"
               Height          =   195
               Left            =   2160
               TabIndex        =   28
               Top             =   780
               Width           =   330
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Total"
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   780
               Width           =   360
            End
            Begin VB.Label lbl2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Qtde"
               Height          =   195
               Left            =   1380
               TabIndex        =   26
               Top             =   780
               Width           =   345
            End
         End
         Begin VB.Frame frmReferente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Referente"
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
            Height          =   3255
            Left            =   60
            TabIndex        =   22
            Top             =   1080
            Width           =   9375
            Begin VB.TextBox txtTotalProd 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   5820
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   1020
               Width           =   1095
            End
            Begin VB.TextBox txtQuant 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   5100
               TabIndex        =   5
               Top             =   1020
               Width           =   675
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E0E0E0&
               Height          =   495
               Left            =   120
               TabIndex        =   75
               Top             =   180
               Width           =   9195
               Begin VB.OptionButton optTipoAvulso 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Avulso"
                  Height          =   195
                  Left            =   660
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   80
                  Top             =   180
                  Value           =   -1  'True
                  Width           =   915
               End
               Begin VB.OptionButton optTipoCadastrado 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cadastrado"
                  Height          =   195
                  Left            =   1620
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   79
                  Top             =   180
                  Width           =   1275
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
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
                  Height          =   195
                  Left            =   120
                  TabIndex        =   81
                  Top             =   180
                  Width           =   450
               End
            End
            Begin VB.TextBox txtCodProduto 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3360
               TabIndex        =   72
               Top             =   720
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.ComboBox cboProdutos 
               Height          =   315
               Left            =   120
               TabIndex        =   3
               Top             =   1020
               Width           =   3915
            End
            Begin MSFlexGridLib.MSFlexGrid GridProdutos 
               Height          =   1395
               Left            =   120
               TabIndex        =   8
               Top             =   1440
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   2461
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin VB.TextBox txtValorProduto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4080
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   1020
               Width           =   975
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarProduto 
               Height          =   315
               Left            =   6960
               TabIndex        =   7
               ToolTipText     =   "Adiciona"
               Top             =   1020
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Adicionar"
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
               MICON           =   "Receber_Cadastro.frx":196BA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverProduto 
               Height          =   315
               Left            =   8100
               TabIndex        =   9
               ToolTipText     =   "Remove"
               Top             =   1020
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Remover"
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
               MICON           =   "Receber_Cadastro.frx":196D6
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Total"
               Height          =   195
               Left            =   5820
               TabIndex        =   89
               Top             =   780
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Quant."
               Height          =   195
               Left            =   5100
               TabIndex        =   88
               Top             =   780
               Width           =   480
            End
            Begin VB.Label lblSomaReferente 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   255
               Left            =   8340
               TabIndex        =   30
               Top             =   2880
               Width           =   915
            End
            Begin VB.Label lbl7 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor:"
               Height          =   195
               Left            =   4080
               TabIndex        =   24
               Top             =   780
               Width           =   405
            End
            Begin VB.Label lbl6 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Produto:"
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   780
               Width           =   600
            End
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdCadastrarCliente 
         Height          =   615
         Left            =   9660
         TabIndex        =   36
         Top             =   3780
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Clientes"
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
         MICON           =   "Receber_Cadastro.frx":196F2
         PICN            =   "Receber_Cadastro.frx":1970E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridConsulta 
         Height          =   5595
         Left            =   -74940
         TabIndex        =   42
         Top             =   1440
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   9869
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   9660
         TabIndex        =   16
         Top             =   1800
         Width           =   1515
         _ExtentX        =   2672
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
         MICON           =   "Receber_Cadastro.frx":19FE8
         PICN            =   "Receber_Cadastro.frx":1A004
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
         Left            =   9660
         TabIndex        =   17
         Top             =   2460
         Width           =   1515
         _ExtentX        =   2672
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
         MICON           =   "Receber_Cadastro.frx":1BD96
         PICN            =   "Receber_Cadastro.frx":1BDB2
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
         Left            =   9660
         TabIndex        =   18
         Top             =   3120
         Width           =   1515
         _ExtentX        =   2672
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
         MICON           =   "Receber_Cadastro.frx":1DB44
         PICN            =   "Receber_Cadastro.frx":1DB60
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
         Left            =   9660
         TabIndex        =   15
         Top             =   1140
         Width           =   1515
         _ExtentX        =   2672
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
         MICON           =   "Receber_Cadastro.frx":1F8F2
         PICN            =   "Receber_Cadastro.frx":1F90E
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
         Left            =   9660
         TabIndex        =   0
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
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
         MICON           =   "Receber_Cadastro.frx":216A0
         PICN            =   "Receber_Cadastro.frx":216BC
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
         Height          =   615
         Left            =   -74640
         TabIndex        =   71
         Top             =   7380
         Width           =   1515
         _ExtentX        =   2672
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
         MICON           =   "Receber_Cadastro.frx":2344E
         PICN            =   "Receber_Cadastro.frx":2346A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
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
         Left            =   -66960
         TabIndex        =   46
         Top             =   8040
         Width           =   1755
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
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
         Left            =   -66960
         TabIndex        =   45
         Top             =   7140
         Width           =   1755
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total:"
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
         Left            =   -66960
         TabIndex        =   44
         Top             =   7440
         Width           =   1755
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Haveres:"
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
         Left            =   -66960
         TabIndex        =   43
         Top             =   7740
         Width           =   1755
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   49
      Top             =   9450
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13282
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
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
      Left            =   180
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Receber_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Private printSQL As String

Private Sub CalcularParcelas()
Dim var_Total As Currency
Dim var_Quant As Integer
Dim var_Parc As Currency

If cboQuantParc.Text = "" Then cboQuantParc.Text = "1"
If txtTotal.Text = "" Then txtTotal.Text = "0"

If optMultiplica.Value = True Then
    txtParc.Text = Format(txtTotal, "##,##0.00")
ElseIf optDivide.Value = True Then
    var_Total = txtTotal.Text
    var_Quant = cboQuantParc.Text
    var_Parc = var_Total / var_Quant
    txtParc.Text = Format(var_Parc, "##,##0.00")
End If
End Sub

Private Sub CalcularQuantParcelas()
Dim date1 As Date
Dim date2 As Date
Dim Result As Integer

If Not IsDate(mskInicio) Then Exit Sub
If Not IsDate(mskFinal) Then Exit Sub

date1 = CDate(mskInicio.Text)
date2 = CDate(mskFinal.Text)

Result = DateDiff("m", date1, date2)
Result = Result + 1

cboQuantParc.Text = Result
End Sub

'Dim var_CodItem As Long
'Dim X As Long
'Dim Y As Long
'Dim f As Integer
'Dim var_STATUS As String
'Dim RESULTADO As Currency

'Dim Texto As String         'preencher os combo
'Dim i, Posicao As Integer   'preencher os combo
'Dim Posicionar As Boolean   'preencher os combo

Sub FormatarGridConsulta(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridConsulta
      .Clear
      .Cols = 13
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1000
      .ColWidth(3) = 800
      .ColWidth(4) = 2500
      .ColWidth(5) = 400
      .ColWidth(6) = 900
      .ColWidth(7) = 1100
      .ColWidth(8) = 750
      .ColWidth(9) = 750
      .ColWidth(10) = 900
      .ColWidth(11) = 750
      .ColWidth(12) = 900
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "ORIGEM"
      .TextMatrix(0, 3) = "PEDIDO"
      .TextMatrix(0, 4) = "CLIENTE"
      .TextMatrix(0, 5) = "No."
      .TextMatrix(0, 6) = "VENC."
      .TextMatrix(0, 7) = "SUBTOTAL"
      .TextMatrix(0, 8) = "DIAS"
      .TextMatrix(0, 9) = "JUROS"
      .TextMatrix(0, 10) = "TOTAL"
      .TextMatrix(0, 11) = "HAVER"
      .TextMatrix(0, 12) = "LIQUIDO"
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
            .TextMatrix(.rows - 1, 1) = ValidateNull(rTabela("cod"))
            .TextMatrix(.rows - 1, 2) = rTabela("campo00")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("campo01"), "000000")
            .TextMatrix(.rows - 1, 4) = rTabela("nome")
            .TextMatrix(.rows - 1, 5) = rTabela("campo02")
            .TextMatrix(.rows - 1, 6) = Format(rTabela("campo03"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 7) = Format(rTabela("campo04"), ocMONEY)
            .TextMatrix(.rows - 1, 8) = rTabela("var_atrazo")
            .TextMatrix(.rows - 1, 9) = Format(rTabela("var_juros"), ocMONEY)
            .TextMatrix(.rows - 1, 10) = Format(rTabela("var_total"), ocMONEY)
            '.TextMatrix(.Rows - 1, 11) = Format(RS!CAMPO06, "##,##0.00")
            '.TextMatrix(.Rows - 1, 12) = Format(RS!CAMPO07, "##,##0.00")
            
            If Not IsNull(rTabela("campo06")) Then
               .TextMatrix(.rows - 1, 11) = Format(rTabela("campo06"), ocMONEY)
               .TextMatrix(.rows - 1, 12) = Format(rTabela("campo07"), ocMONEY)
            Else
               .TextMatrix(.rows - 1, 11) = Format(0, ocMONEY)
               .TextMatrix(.rows - 1, 12) = Format(rTabela("var_total"), ocMONEY)
            End If
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .Redraw = True
      .rows = .rows - 1

        'mudar a cor da coluna
        For i = 1 To .rows - 1
           .Row = i
           .Col = 10
           .CellBackColor = &HC0FFFF
        Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      For i = 1 To .rows - 1
         .Row = i
         .Col = 9
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
   End With
   
   txtCONValor.Text = Format(SomaGrid(GridConsulta, 10), ocMONEY)
   txtCONHaver.Text = Format(SomaGrid(GridConsulta, 11), ocMONEY)
   txtCONtotal.Text = Format(SomaGrid(GridConsulta, 12), ocMONEY)
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub FormatarGridProdutos(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridProdutos
      .Clear
      .Cols = 6
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 0
      .ColWidth(2) = 5500
      .ColWidth(3) = 1100
      .ColWidth(4) = 900
      .ColWidth(5) = 1100
      
      .TextMatrix(0, 2) = "PRODUTO"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "QUANT."
      .TextMatrix(0, 5) = "TOTAL"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 2
      .ColAlignment(3) = 3
      .ColAlignment(4) = 4
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            For i = 1 To .rows - 1
               GridProdutos.Row = i
               'GridProdutos.Col = 1:   GridProdutos.CellBackColor = vbYellow
               'GridProdutos.Col = 5:   GridProdutos.CellBackColor = vbYellow
            Next
            
            .TextMatrix(.rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("descricao"))
            .TextMatrix(.rows - 1, 3) = FormatNumber(rTabela("PRECO"), 2)
            .TextMatrix(.rows - 1, 4) = rTabela("QUANTIDADE")
            .TextMatrix(.rows - 1, 5) = FormatNumber(rTabela("TOTAL"), 2)
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 5
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   lblSomaReferente.Caption = Format(SomaGrid(GridProdutos, 5), ocMONEY)
End Sub

Private Sub LimparGridProdutos()
   Dim i As Integer
   
   With GridProdutos
      .Clear
      .Cols = 4
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 0
      .ColWidth(2) = 3900
      .ColWidth(3) = 1100
      
      .TextMatrix(0, 2) = "PRODUTO"
      .TextMatrix(0, 3) = "VALOR"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 2
   End With
End Sub

Private Sub PreencherComboStatus()
   cboCONStatus.Clear
   cboCONStatus.AddItem "À RECEBER"
   cboCONStatus.AddItem "RECEBIDO"
   
   If cboCONStatus.Text = "" Then cboCONStatus.ListIndex = 0
   moCombo.AttachTo cboCONStatus
End Sub

Private Sub PreencherGridProdutos()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT codigo, DESCRICAO, preco, quantidade, total FROM a_receber_itens WHERE (cod_pedido = " & lblCodigo.Caption & ")" & _
   "UNION ALL "
sSQL = sSQL & "SELECT pedidos_itens.COD_PRODUTO, produtos.DESCRICAO, pedidos_itens.PRECO, pedidos_itens.quantidade, pedidos_itens.total FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO WHERE (cod_pedido = " & lblCodigo.Caption & ");"

Set r = dbData.OpenRecordset(sSQL)
FormatarGridProdutos r
If r.State <> 0 Then r.Close
End Sub

Private Function Atualizar_Dados() As Boolean
Dim sSQL As String

'ATUALIZA TABELA PEDIDOS
sSQL = "UPDATE pedidos SET " & _
        "cod_cliente = " & txtCodCliente.Text & ", " & _
        "data_compra = CONVERT(DATETIME, '" & Format(mskCompra.Text, ocDATA) & "', 103), " & _
        "subtotal = " & Replace(CCur(txtTotal.Text), ",", ".") & ", " & _
        "total = " & Replace(CCur(txtTotal.Text), ",", ".") & ""

sSQL = sSQL & "WHERE (cod_PEDIDO = " & lblCodigo.Caption & ");"

Atualizar_Dados = dbData.Execute(sSQL)

'ATUALIZA TABELA PARCELAS
sSQL = "UPDATE PARCELAS SET " & _
        "DATA = CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103), " & _
        "VALOR = " & Replace(CCur(txtTotal.Text), ",", ".") & ""

sSQL = sSQL & "WHERE (cod_PEDIDO = " & lblCodigo.Caption & ");"

Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Function Inserir_Dados() As Boolean
   'A inclusão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO a_receber (" & _
      "cod_receber, cod_cliente, cliente) VALUES ("
   
   sSQL = sSQL & _
      lblCodigo.Caption & ", " & txtCodCliente.Text & ", '" & txtCliente.Text & "')"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Sub AutoNumeracao_Pedidos()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(cod_pedido), 0) AS ultimo FROM pedidos;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lblCodigo.Caption = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Function AutonNumeracao_Caixa() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo) AS cod FROM caixa_entrada;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutonNumeracao_Caixa = lRet
End Function

Private Function AutoNumeracao_Detalhes() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo) AS cod_detalhe FROM a_receber_visitas;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao_Detalhes = lRet
End Function

Private Sub Limpar_Objetos()
'If cmdAlterar.Enabled = False Then lblCodigo.Caption = ""
txtCodCliente.Text = ""
txtCliente.Text = ""
'If cmdAlterar.ENABLED = False Then txtCodCobrador.Text = ""
txtTotal.Text = ""
lblSomaReferente.Caption = ""
txtTotal.Text = ""
cboQuantParc.Text = ""
txtParc.Text = ""
mskInicio.Mask = ""
mskInicio.Text = ""
mskFinal.Mask = ""
mskFinal.Text = ""
mskCompra.Mask = ""
mskCompra.Text = ""
lblCodigo.Caption = ""
optPorQuant.Value = True
optMultiplica.Value = True
frmCliente.Enabled = False
frmReferente.Enabled = False
frmParcelas.Enabled = False
LimparGridProdutos
End Sub

Private Sub cboAno_GotFocus()
   Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
   Dim i As Integer
     
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
   '   cboAno.AddItem i
   'Next
End Sub

Private Sub cboCONStatus_Change()
   cboCONStatus_Click
End Sub

Private Sub cboCONStatus_Click()
   'cmdExibirConsulta_Click 'desativei
End Sub

Private Sub cboCONStatus_GotFocus()
   PreencherComboStatus
End Sub

Private Sub cboCONStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cboFiltro_Change()
cboFiltro_LostFocus
End Sub

Private Sub cboFiltro_Click()
cboFiltro_LostFocus
End Sub


Private Sub cboFiltro_GotFocus()
cboFiltro.Clear
cboFiltro.AddItem "TODOS"
cboFiltro.AddItem "MÊS"
cboFiltro.AddItem "PERIODO"
cboFiltro.AddItem "CLIENTE"
moCombo.AttachTo cboFiltro
End Sub


Private Sub cboFiltro_LostFocus()
If cboFiltro.Text = "TODOS" Then
    lblCONmes.Visible = False
    cboMES.Visible = False
    cboAno.Visible = False
    lblCONint1.Visible = False
    Mask1.Visible = False
    lblCONint2.Visible = False
    Mask2.Visible = False
    cmdInicial.Visible = False
    cmdFinal.Visible = False
    lblCONnome.Visible = False
    cboNome.Visible = False
    'cmdExibirConsulta_Click
ElseIf cboFiltro.Text = "MÊS" Then
   lblCONmes.Visible = True
   cboMES.Visible = True
   cboAno.Visible = True
   lblCONint1.Visible = False
   Mask1.Visible = False
   lblCONint2.Visible = False
   Mask2.Visible = False
    cmdInicial.Visible = False
    cmdFinal.Visible = False
   lblCONnome.Visible = False
   cboNome.Visible = False
   cboMES.SetFocus
ElseIf cboFiltro.Text = "PERIODO" Then
   lblCONmes.Visible = False
   cboMES.Visible = False
   cboAno.Visible = False
   lblCONint1.Visible = True
   Mask1.Visible = True
   lblCONint2.Visible = True
   Mask2.Visible = True
   Mask1.SetFocus
    cmdInicial.Visible = True
    cmdFinal.Visible = True
   lblCONnome.Visible = False
   cboNome.Visible = False
ElseIf cboFiltro.Text = "CLIENTE" Then
   lblCONmes.Visible = False
   cboMES.Visible = False
   cboAno.Visible = False
   lblCONint1.Visible = False
   Mask1.Visible = False
   lblCONint2.Visible = False
   Mask2.Visible = False
    cmdInicial.Visible = False
    cmdFinal.Visible = False
   lblCONnome.Visible = True
   cboNome.Visible = True
   cboNome.SetFocus
End If
End Sub


Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMES.Clear
   
   For vMes = 1 To 12
      cboMES.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMES
End Sub

Private Sub cboNome_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboNome.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboNome.AddItem ValidateNull(r("nome"))
      cboNome.ItemData(cboNome.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboNome
End Sub

Private Sub cboNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'cmdExibirConsulta_Click
   End If
   
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboNome_LostFocus()
   On Error GoTo TrataErro
   
   If cboNome.Text = "" Then txtCodClienteCons.Text = "": Exit Sub
   txtCodClienteCons = cboNome.ItemData(cboNome.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then cboNome.Text = ""
End Sub

Private Sub cboOrdem_Change()
'cmdExibirConsulta_Click
End Sub

Private Sub cboOrdem_GotFocus()
cboOrdem.Clear
cboOrdem.AddItem "VENC."
cboOrdem.AddItem "CLIENTE"
moCombo.AttachTo cboFiltro
End Sub


Private Sub cboProdutos_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

cboProdutos.Clear

If optTipoCadastrado.Value = True Then
   sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboProdutos.AddItem ValidateNull(r("descricao"))
      cboProdutos.ItemData(cboProdutos.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
    If r.State <> 0 Then r.Close
    Set r = Nothing
ElseIf optTipoAvulso.Value = True Then
   ' sSQL = "SELECT DISTINCT DESCRICAO as varDesc, CODIGO, COD_PRODUTO as varCodProd, PRECO FROM a_receber_itens"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Do While Not r.EOF
   '   cboProdutos.AddItem ValidateNull(r("varDesc"))
   '   cboProdutos.ItemData(cboProdutos.NewIndex) = r("varCodProd")
   '   r.MoveNext
   'Loop
End If
   
   moCombo.AttachTo cboProdutos
End Sub

Private Sub cboProdutos_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboProdutos_LostFocus()
On Error GoTo TrataErro

Dim sSQL As String
Dim r As ADODB.Recordset

If cboProdutos.Text = "" Then txtCodProduto.Text = "": Exit Sub

If optTipoCadastrado.Value = True Then
    txtCodProduto = cboProdutos.ItemData(cboProdutos.ListIndex)
    
        sSQL = "SELECT pedidos_itens.COD_PRODUTO, pedidos_itens.PRECO, produtos.CODIGO, produtos.DESCRICAO FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO WHERE (COD_PRODUTO = " & txtCodProduto.Text & ");"
    Debug.Print sSQL
    Set r = dbData.OpenRecordset(sSQL)

    If Not r.BOF Then
        txtValorProduto.Text = Format(ValidateNull(r("PRECO")), "##,##0.00")
    End If
Else
    If cboProdutos.ListIndex = -1 Then
        Dim varCodProdNovo As Integer
        sSQL = "SELECT ISNULL(MAX(cod_produto), 0) AS ultimo_codigo FROM a_receber_itens;"
        Set r = dbData.OpenRecordset(sSQL)
        
        If Not r.BOF Then varCodProdNovo = r("ultimo_codigo") + 1
        If r.State <> 0 Then r.Close
        Set r = Nothing
    
        txtCodProduto = varCodProdNovo
    Else
        txtCodProduto = cboProdutos.ItemData(cboProdutos.ListIndex)
        
        sSQL = "SELECT COD_PRODUTO, descricao, PRECO FROM  a_receber_itens WHERE (COD_PRODUTO = " & txtCodProduto.Text & ");"
        Set r = dbData.OpenRecordset(sSQL)
    
        If Not r.BOF Then
            txtValorProduto.Text = Format(ValidateNull(r("PRECO")), "##,##0.00")
        End If
    End If
End If


TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboQuantParc_Change()
Call CalcularParcelas
Call Calcular_Prazo
End Sub

Private Sub cboQuantParc_GotFocus()
cboQuantParc.Clear
cboQuantParc.AddItem "1"
cboQuantParc.AddItem "2"
cboQuantParc.AddItem "3"
cboQuantParc.AddItem "4"
cboQuantParc.AddItem "5"
cboQuantParc.AddItem "6"
cboQuantParc.AddItem "7"
cboQuantParc.AddItem "8"
cboQuantParc.AddItem "9"
cboQuantParc.AddItem "10"
cboQuantParc.AddItem "11"
cboQuantParc.AddItem "12"
moCombo.AttachTo cboQuantParc
End Sub


Private Sub cboQuantParc_LostFocus()
If cboQuantParc.Text = "" Then cboQuantParc = 1
Call CalcularParcelas
Call Calcular_Prazo
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

mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
mskInicio_LostFocus
End Sub

Private Sub chameleonButton2_Click()
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

mskFinal = Format(varData, "dd/mm/yy")   'Exibe a data no campo
mskFinal_LostFocus
End Sub


Private Sub cmdAdicionarProduto_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim varCodItem As Long
   
   If txtCodCliente.Text = "" Then
      ShowMsg "Falta escolher o cliente!", vbInformation
      txtCliente.SetFocus
      Exit Sub
   End If
   
   If lblCodigo.Caption = "" Or cboProdutos.Text = "" Or txtValorProduto.Text = "" Then Exit Sub
   
   If mskCompra.Text = "" Then
      ShowMsg "Falta a data da compra!", vbInformation
      mskCompra.SetFocus
      Exit Sub
   End If
   
   varCodItem = 1
   
If optTipoAvulso.Value = True Then
    'indice do codigo
    sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_codigo FROM a_receber_itens;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then varCodItem = r("ultimo_codigo") + 1
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    'indice do item
    Dim IndiceItem As Integer
    
    sSQL = "SELECT ISNULL(MAX(item), 0) AS ultimo_item FROM a_receber_itens;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then varCodItem = r("ultimo_item") + 1
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    dbData.Execute "INSERT INTO a_receber_itens (codigo, cod_pedido, cod_produto, descricao, preco, quantidade, total, data, tipo_venda, item) VALUES (" & _
          varCodItem & ", " & lblCodigo.Caption & ", " & txtCodProduto.Text & ", '" & cboProdutos.Text & "', " & Replace(CCur(txtValorProduto.Text), ",", ".") & ", " & txtQuant.Text & ", " & Replace(CCur(txtTotalProd.Text), ",", ".") & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), 'BALCAO', " & varCodItem & ");"

Else
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_item FROM pedidos_itens;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then varCodItem = r("ultimo_item") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing

   dbData.Execute "INSERT INTO pedidos_itens (codigo, cod_pedido, cod_produto, preco, quantidade, subtotal, total, data, tipo_venda) VALUES (" & _
         varCodItem & ", " & lblCodigo.Caption & ", " & txtCodProduto.Text & ", " & Replace(CCur(txtValorProduto.Text), ",", ".") & ", " & txtQuant.Text & ", " & Replace(CCur(txtTotalProd.Text), ",", ".") & ", " & Replace(CCur(txtTotalProd.Text), ",", ".") & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), 'BALCAO');"
End If

LimparObjetos_Produto
cboProdutos.SetFocus
PreencherGridProdutos
End Sub



Private Sub LimparObjetos_Produto()
cboProdutos.Text = ""
txtValorProduto.Text = ""
txtCodProduto.Text = ""
txtQuant.Text = ""
txtTotalProd.Text = ""
End Sub



Private Sub cmdAlterar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If lblCodigo.Caption = "" Or txtCodCliente.Text = "" Then Exit Sub

sSQL = "SELECT codigo FROM parcelas WHERE (cod_pedido = " & lblCodigo.Caption & ") and status = 1 "
Set r = dbData.OpenRecordset(sSQL)

If r.RecordCount > 0 Then
    MsgBox "Existe uma(s) parcela(s) já quitada(s) para essa conta! Não será possivel modificar essa conta.", vbInformation, "Aviso do Sistema"
    Exit Sub
End If
If r.State <> 0 Then r.Close

If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
LimparGridProdutos
Form_Load
cmdExibirConsulta_Click
End Sub


Private Sub cmdCadastrarCliente_Click()
   Clientes_Cadastro.Show 1
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

mskCompra = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub
Private Sub cmdCancelar_Click()
dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & lblCodigo.Caption & ");"
dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & lblCodigo.Caption & ");"

frmCliente.Enabled = False
frmReferente.Enabled = False
frmParcelas.Enabled = False
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
'cmdImprimirConsulta.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdAdicionarProduto.Enabled = False
cmdRemoverProduto.Enabled = False

Limpar_Objetos
LimparGridProdutos
Form_Load
End Sub

Public Function Verifica_Dia(DIA, var_Mes)
   Dim diasDoMes As Variant
   
   DIA = Val(DIA)
   diasDoMes = Array(31, 28, 30, 30, 31, 30, 31, 30, 30, 31, 30, 31)
   
   If DIA = 31 Then
      Verifica_Dia = diasDoMes(var_Mes - 1)
   Else
      Verifica_Dia = DIA
   End If
End Function

Private Sub Gerar_Parcelas()
Dim i As Integer
Dim lNovoCod As Long
Dim var_Vencimento As Date
Dim Var_NumParc As Integer

var_Vencimento = CDate(mskInicio.Text)
Var_NumParc = 1

For i = 1 To CInt(cboQuantParc)
   lNovoCod = Autonumeracao_Parcelas
   
   dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, status, valor, VALOR_FINAL) VALUES (" & _
      lNovoCod & ", " & lblCodigo.Caption & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(var_Vencimento, ocDATA) & "', 103), 0, " & _
      Replace(CCur(txtParc.Text), ",", ".") & ", " & Replace(CCur(txtParc.Text), ",", ".") & ");"
   
   var_Vencimento = Format(DateAdd("m", Val(1), var_Vencimento), "dd/mm/yy")
   Var_NumParc = Var_NumParc + 1
Next
End Sub

Private Function Autonumeracao_Parcelas() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultima_parcela FROM parcelas;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("ultima_parcela") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   Autonumeracao_Parcelas = lRet
End Function

Private Sub cmdExcluir_Click()
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite essa operação!", vbInformation, "Aviso do Sistema": Exit Sub

If lblCodigo.Caption = "" Then Exit Sub

If ShowMsg("Excluir essa Conta?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT codigo FROM parcelas WHERE (cod_pedido = " & lblCodigo.Caption & ") and status = 1 "
Set r = dbData.OpenRecordset(sSQL)

If r.RecordCount > 0 Then
    MsgBox "Existe uma(s) parcela(s) já quitada(s) para essa conta! Não será possivel excluir essa conta.", vbInformation, "Aviso do Sistema"
    Exit Sub
End If
If r.State <> 0 Then r.Close

dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & lblCodigo.Caption & ");"
dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & lblCodigo.Caption & ");"

frmCliente.Enabled = False
frmReferente.Enabled = False
frmParcelas.Enabled = False
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
'cmdImprimirConsulta.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdAdicionarProduto.Enabled = False
cmdRemoverProduto.Enabled = False

Limpar_Objetos
LimparGridProdutos
Form_Load
cmdExibirConsulta_Click
End Sub

Private Sub cmdExibirConsulta_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   Dim oCfg As ConfigItem
   Dim var_JurosDia As Double
   
   Dim INDICE As String
   Dim var_STATUS As String
   Dim Tipo_Data As String
   
   'PEGAR O JUROS DO DIA
   Set oCfg = sysConfig("JUROS_DIA")
   var_JurosDia = CCur(oCfg.Value)
   Set oCfg = Nothing
   
   If cboFiltro.Text = "" Then Exit Sub
   
   'INDICE PARA ORGANIZAR OS DADOS
   If cboOrdem.Text = "CLIENTE" Then
      INDICE = "nome"
   ElseIf cboOrdem.Text = "VENC." Then
      INDICE = "campo03"
   Else
      cboOrdem.Text = "VENC."
      INDICE = "campo03"
   End If
   
   'STATUS
   If cboCONStatus.Text = "À RECEBER" Then
      var_STATUS = "AND (parcelas.status = 0) "
   Else
      var_STATUS = "AND (parcelas.status = 1) "
   End If
   
   
   '***********************************************
   'Variáveis para montagem da string
   Dim sql_ATRAZO As String
   Dim sql_JUROS As String
   Dim sql_HAVER As String
   
   sql_ATRAZO = "CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END"
   sql_JUROS = "(parcelas.valor * " & Replace(var_JurosDia / 100, ",", ".") & ")"
   sql_HAVER = "(SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE parcelas_haver.cod_parcela = parcelas.codigo)"
   '***********************************************
   '(SELECT ISNULL(SUM(valor_haver), 0) FROM
   
   If cboFiltro.Text = "TODOS" Then
      'MOSTRAR NO GRID
      sSQL = "SELECT " & sql_ATRAZO & " AS var_atrazo, " & _
         "(" & sql_JUROS & " * " & sql_ATRAZO & ") AS var_juros, " & _
         "(parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) AS var_total, " & sql_HAVER & " AS campo06, " & _
         "((parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) - " & sql_HAVER & ") AS campo07, cliente.nome, " & _
         "parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_STATUS, pedidos.cod_pedido, pedidos.cod_cliente " & _
         "FROM pedidos INNER JOIN cliente ON pedidos.cod_cliente = cliente.codigo " & _
         "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido WHERE (pedidos.tipo_pedido = 'RECEBER') " & _
         var_STATUS & " ORDER BY " & INDICE
      
   ElseIf cboFiltro.Text = "MÊS" Then
      If cboMES.Text = "" Or cboAno.Text = "" Then
         MsgBox "Selecione o MÊS e depois o ANO!", vbExclamation
         cboMES.SetFocus
         Exit Sub
      End If
      
      'DATA DE VENCIMENTO OU PAGAMENTO
      Tipo_Data = "AND (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(DATA) = " & cboAno & ") "
      
      'MOSTRAR NO GRID
      sSQL = "SELECT " & sql_ATRAZO & " AS var_atrazo, " & _
         "(" & sql_JUROS & " * " & sql_ATRAZO & ") AS var_juros, " & _
         "(parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) AS var_total, " & sql_HAVER & " AS campo06, " & _
         "((parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) - " & sql_HAVER & ") AS campo07, cliente.nome, " & _
         "parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05,  CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_STATUS, pedidos.cod_pedido, pedidos.cod_cliente " & _
         "FROM pedidos INNER JOIN cliente ON pedidos.cod_cliente = cliente.codigo " & _
         "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido WHERE (pedidos.tipo_pedido = 'RECEBER') " & _
         Tipo_Data & var_STATUS & " ORDER BY " & INDICE
      
   ElseIf cboFiltro.Text = "PERIODO" Then
      If Not IsDate(Mask1.Text) Or Not IsDate(Mask2.Text) Then Exit Sub
      If Mask1.Text = "" Or Mask2.Text = "" Then MsgBox "Digite a DATA INICIAL e DATA FINAL!", , "Aviso do Sistema": Mask1.SetFocus:      Exit Sub
      
      'MOSTRAR NO GRID
      sSQL = "SELECT " & sql_ATRAZO & " AS var_atrazo, " & _
         "(" & sql_JUROS & " * " & sql_ATRAZO & ") AS var_juros, " & _
         "(parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) AS var_total, " & sql_HAVER & " AS campo06, " & _
         "((parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) - " & sql_HAVER & ") AS campo07, cliente.nome, " & _
         "parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_STATUS, pedidos.cod_pedido, pedidos.cod_cliente " & _
         "FROM pedidos INNER JOIN cliente ON pedidos.cod_cliente = cliente.codigo " & _
         "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (pedidos.tipo_pedido = 'RECEBER') AND (parcelas.data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) " & _
         "AND (parcelas.data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103)) " & var_STATUS & " ORDER BY " & INDICE
        
   ElseIf cboFiltro.Text = "CLIENTE" Then
      If txtCodClienteCons.Text = "" Then MsgBox "Selecione um cliente!", , "Aviso do Sistema": Exit Sub
      If cboNome.Text = "" Then MsgBox "Digite o nome do cliente!", vbExclamation, "Aviso do Sistema": cboNome.SetFocus:     Exit Sub
      
      'MOSTRAR NO GRID
      sSQL = "SELECT " & sql_ATRAZO & " AS var_atrazo, " & _
         "(" & sql_JUROS & " * " & sql_ATRAZO & ") AS var_juros, " & _
         "(parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) AS var_total, " & sql_HAVER & " AS campo06, " & _
         "((parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) - " & sql_HAVER & ") AS campo07, cliente.nome, " & _
         "parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_STATUS, pedidos.cod_pedido, pedidos.cod_cliente " & _
         "FROM pedidos INNER JOIN cliente ON pedidos.cod_cliente = cliente.codigo " & _
         "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (pedidos.tipo_pedido = 'RECEBER') AND (cliente.codigo = " & txtCodClienteCons.Text & ") " & _
         var_STATUS & " ORDER BY " & INDICE
      
   End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   FormatarGridConsulta r
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'CONTAR QUANTOS REGISTROS COM UMA DATA
   txtCONquant.Text = Format(totalRegistros, "00")
   printSQL = sSQL
End Sub

Private Sub cmdFinal_Click()
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

Mask2 = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdImprimirConsulta_Click()
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

Set REL_Receber_Consulta.Relatorio.Recordset = r
REL_Receber_Consulta.dfQuant.Caption = "QUANTIDADE: " & txtCONquant.Text
REL_Receber_Consulta.dfTotal.Caption = "TOTAL: " & txtCONtotal.Text
REL_Receber_Consulta.lblTitulo.Caption = "RELATÓRIO - CONTAS À RECEBER/RECEBIDO"

If cboFiltro.Text = "TODOS" Then
   REL_Receber_Consulta.dfTipo.Caption = "Tipo: Todos os registros"
ElseIf cboFiltro.Text = "PERIODO" Then
   REL_Receber_Consulta.dfTipo.Caption = "Tipo: Intervalo de " & Mask1.Text & " à " & Mask2.Text
ElseIf cboFiltro.Text = "MÊS" Then
   REL_Receber_Consulta.dfTipo.Caption = "Tipo: Mês = " & cboMES.Text & "/" & cboAno.Text
ElseIf cboFiltro.Text = "CLIENTE" Then
   REL_Receber_Consulta.dfTipo.Caption = "Cliente = " & cboNome.Text
Else
   REL_Receber_Consulta.dfTipo.Caption = "Tipo:"
End If

REL_Receber_Consulta.Relatorio.NomeImpressora = var_Impressora
REL_Receber_Consulta.Relatorio.Ativar
Unload REL_Receber_Consulta

Me.Show
End Sub

Private Sub cmdInicial_Click()
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

Mask1 = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdNovo_Click()
Limpar_Objetos
LimparGridProdutos
SSTab1.Tab = 0

AutoNumeracao_Pedidos
dbData.Execute "INSERT INTO pedidos (cod_pedido, tipo_pedido, status_pedido) VALUES (" & lblCodigo.Caption & ", 'RECEBER', 0);"

frmCliente.Enabled = True
frmReferente.Enabled = True
frmParcelas.Enabled = True
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
'cmdImprimirConsulta.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdAdicionarProduto.Enabled = True
cmdRemoverProduto.Enabled = True
optPorQuant_Click
mskCompra.SetFocus
End Sub

Private Function AutoNumeracao_Itens() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_item FROM pedidos_itens;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("ultimo_item") + 1
   If Not r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao_Itens = lRet
End Function

Private Sub cmdRemoverProduto_Click()
Dim i As Integer

'Row retorna a linha selecionada
i = GridProdutos.Row

If GridProdutos.TextMatrix(i, 2) = "" Then Exit Sub
If ShowMsg("Deseja remover o produto: " & GridProdutos.TextMatrix(i, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

If optTipoAvulso.Value = True Then
    dbData.Execute "DELETE FROM a_receber_itens WHERE (codigo = " & GridProdutos.TextMatrix(i, 1) & ");"
End If

PreencherGridProdutos
End Sub

Private Sub cmdSalvar_Click()
   If lblCodigo.Caption = "" Or txtCliente.Text = "" Or txtParc.Text = "" Then Exit Sub
   
   If mskInicio.Text = "" Then
      ShowMsg "Falta a data de vencimento!", vbInformation
      mskInicio.SetFocus
      Exit Sub
   End If
   
      dbData.Execute "UPDATE pedidos SET cod_pedido = " & lblCodigo.Caption & ", cod_cliente = " & txtCodCliente.Text & ", " & _
      "data_compra = CONVERT(DATETIME, '" & Format(mskCompra.Text, ocDATA) & "', 103), tipo_desc = 'R', valor_desc = 0, " & _
      "subtotal = " & Replace(CCur(txtTotal.Text), ",", ".") & ", total = " & Replace(CCur(txtTotal.Text), ",", ".") & ", " & _
      "tipo_pagamento = 'À Prazo', pagamento = 'AVULSO', tipo_pedido = 'RECEBER', " & _
      "status_pedido = 1, " & _
      "maquina = 'CAIXA01' WHERE (cod_pedido = " & lblCodigo.Caption & ");"
      
   Gerar_Parcelas
   Limpar_Objetos
   Form_Load
   cmdExibirConsulta_Click
End Sub

Private Sub Calcular_Prazo()
'If cboPrazo.Text = "" Then Exit Sub
If Not IsDate(mskInicio.Text) Then Exit Sub
If mskCompra.Text = "" Then mskCompra.Text = Format(Date, "dd/mm/yy")

If optMultiplica.Value = True Then
   'mskInicio.Text = Format(mskCompra, "dd/mm/yy")
   mskFinal.Text = Format(DateAdd("m", Val(cboQuantParc.Text) - 1, mskInicio.Text), "dd/mm/yy")
ElseIf optDivide.Value = True Then
   'mskInicio.Text = Format(mskCompra, "dd/mm/yy")
   'mskFinal.Text = Format(DateAdd("m", Val(cboQuantParc.Text), mskInicio.Text), "dd/mm/yy")
   mskFinal.Text = Format(DateAdd("m", Val(cboQuantParc.Text) - 1, mskInicio.Text), "dd/mm/yy")
End If
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
'cmdImprimirConsulta.Enabled = False
cmdAdicionarProduto.Enabled = False
cmdRemoverProduto.Enabled = False
LimparGridProdutos
PreencherComboStatus
SSTab1.Tab = 0
cboFiltro.Text = "TODOS"
StatusBar1.Panels(4).Text = Format(Date, "dd/mm/yy")

'colocar o nome da maquina na barra de status
Dim oIni As Ini
Dim var_Maquina As String

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
Set oIni = Nothing

StatusBar1.Panels(2).Text = var_Maquina
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If vChamouCaixa = "PDV" Then
    Me.Hide
    'PDV.Show  'desativei somente para geerar o online comerce
Else
    Me.Hide
    'PDV.Show 1
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'HabilitaObjetosVenda False
Set moCombo = Nothing
End Sub

Private Sub gridConsulta_DblClick()
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
frmCliente.Enabled = True
frmReferente.Enabled = True
frmParcelas.Enabled = True
mskInicio.Enabled = True

lblCodigo.Caption = ""
lblCodigo.Caption = GridConsulta.TextMatrix(GridConsulta.RowSel, 3)
SSTab1.Tab = 0
End Sub

Private Sub lblCodigo_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If lblCodigo.Caption = "" Then Exit Sub
   
   If cmdExcluir.Enabled = True Then
      sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & lblCodigo.Caption & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If Not r.BOF Then
         mskCompra.Text = Format(r("data_compra"), "dd/mm/yy")
         txtTotal.Text = r("subtotal")
         cboQuantParc.Text = 1
         txtParc.Text = Format(r("subtotal"), ocMONEY)
         'mskInicio.Text = Format(r("vencimento"), "dd/mm/yy")
         txtCodCliente.Text = r("cod_cliente")
         mskInicio.Text = Format(DateAdd("m", 1, mskCompra), "dd/mm/yy")
      End If
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      PreencherGridProdutos
      cmdAdicionarProduto.Enabled = True
      cmdRemoverProduto.Enabled = True
   End If
End Sub

Private Sub lblSomaReferente_Change()
   txtTotal.Text = lblSomaReferente.Caption
   cboQuantParc = 1
   txtParc_GotFocus
End Sub

Private Sub MASK1_KeyPress(KeyAscii As Integer)
   Mask1.Mask = "##/##/##"
End Sub

Private Sub Mask1_LostFocus()
   If Mask1.Text = "__/__/__" Then
      Mask1.Mask = ""
      Mask1.Text = ""
      Exit Sub
   ElseIf Mask1.Text = "" Then
      Mask1.Mask = ""
      Mask1.Text = ""
      Exit Sub
   ElseIf Not IsDate(Mask1) Then
      ShowMsg "Data Inválida", vbExclamation
      Mask1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Mask2_KeyPress(KeyAscii As Integer)
   Mask2.Mask = "##/##/##"
End Sub

Private Sub Mask2_LostFocus()
   If Mask2.Text = "__/__/__" Then
      Mask2.Mask = ""
      Mask2.Text = ""
      Exit Sub
   ElseIf Mask2.Text = "" Then
      Mask2.Mask = ""
      Mask2.Text = ""
      Exit Sub
   ElseIf Not IsDate(Mask2) Then
      ShowMsg "Data Inválida", vbExclamation
      Mask2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub mskCompra_Change()
If IsDate(mskCompra.Text) Then
    mskInicio.Text = Format(mskCompra, "dd/mm/yy")
End If
End Sub

Private Sub mskCompra_GotFocus()
   SelectControl mskCompra
End Sub

Private Sub mskCompra_KeyPress(KeyAscii As Integer)
   mskCompra.Mask = "##/##/##"
End Sub

Private Sub mskCompra_LostFocus()
'If IsDate(mskCompra.Text) Then
'    mskInicio.Text = Format(DateAdd("m", 1, mskCompra), "dd/mm/yy")
'End If
End Sub


Private Sub mskFinal_Change()
'If mskFinal.Text <> "" Then mskFinal_LostFocus
End Sub

Private Sub mskFinal_GotFocus()
SelectControl mskFinal
End Sub


Private Sub mskFinal_KeyPress(KeyAscii As Integer)
mskFinal.Mask = "##/##/##"
End Sub


Private Sub mskFinal_LostFocus()
If mskFinal.Text = "" Or mskFinal.Text = "__/__/__" Then
   mskFinal.Mask = ""
   mskFinal.Text = ""
   Exit Sub
Else
    If Not IsDate(mskFinal.Text) Then
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      SelectControl mskFinal
   End If
End If

If optPorDatas.Value = True Then
    cboQuantParc.Enabled = False
    chameleonButton2.Enabled = True
    mskFinal.Enabled = True
    CalcularQuantParcelas
Else
    cboQuantParc.Enabled = True
    chameleonButton2.Enabled = False
    mskFinal.Enabled = False
End If
End Sub


Private Sub mskInicio_Change()
Calcular_Prazo
'If mskInicio.Text <> "" Then mskInicio_LostFocus
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
    If Not IsDate(mskInicio.Text) Then
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      SelectControl mskInicio
   End If
End If

If optPorDatas.Value = True Then
    cboQuantParc.Enabled = False
    chameleonButton2.Enabled = True
    mskFinal.Enabled = True
    CalcularQuantParcelas
Else
    cboQuantParc.Enabled = True
    chameleonButton2.Enabled = False
    mskFinal.Enabled = False
End If
End Sub

Private Sub optIntervalo_Click()

End Sub

Private Sub optMensal_Click()

End Sub

Private Sub optNome_Click()

End Sub

Private Sub optTodos_Click()
   
End Sub

Private Sub optDivide_Click()
Call CalcularParcelas
Call Calcular_Prazo
End Sub

Private Sub optMultiplica_Click()
Call CalcularParcelas
Call Calcular_Prazo
End Sub

Private Sub optPorDatas_Click()
If optPorDatas.Value = True Then
    lbl2.Enabled = False
    cboQuantParc.Enabled = False
    Label4.Enabled = True
    chameleonButton2.Enabled = True
    mskFinal.Enabled = True
    CalcularQuantParcelas
Else
    lbl2.Enabled = True
    cboQuantParc.Enabled = True
    Label4.Enabled = False
    chameleonButton2.Enabled = False
    mskFinal.Enabled = False
End If
End Sub

Private Sub optPorQuant_Click()
If optPorDatas.Value = True Then
    lbl2.Enabled = False
    cboQuantParc.Enabled = False
    Label4.Enabled = True
    chameleonButton2.Enabled = True
    mskFinal.Enabled = True
    CalcularQuantParcelas
Else
    lbl2.Enabled = True
    cboQuantParc.Enabled = True
    Label4.Enabled = False
    chameleonButton2.Enabled = False
    mskFinal.Enabled = False
End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
'   If SSTab1.Tab = 1 Then
'      cmdImprimirConsulta.Enabled = False
'      cmdSalvar.Enabled = False
'      cmdAlterar.Enabled = False
'      cmdExcluir.Enabled = False
'   ElseIf SSTab1.Tab = 3 Then
'      cmdImprimirConsulta.Enabled = True
'   End If
End Sub

Private Sub Timer1_Timer()
   lblHora.Caption = Format(Time, "hh:mm")
End Sub

Private Sub txtCliente_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = txtCliente.Text
codAtual = txtCodCliente.Text
txtCliente.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   txtCliente.AddItem r("nome")
   txtCliente.ItemData(txtCliente.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

txtCliente.Text = itemAtual
txtCodCliente.Text = codAtual
moCombo.AttachTo txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCliente_LostFocus()
   On Error GoTo TrataErro
   If txtCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   txtCodCliente = txtCliente.ItemData(txtCliente.ListIndex)
   Exit Sub
   
TrataErro:
  ' If Err.Number = 381 Then txtCodCliente.Text = ""
End Sub

Private Sub TxtCodCliente_Change()
Dim sSQL As String
Dim r As ADODB.Recordset
If txtCodCliente.Text = "" Then Exit Sub

If cmdExcluir.Enabled = True Then
   sSQL = "SELECT codigo, nome FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCliente.Text = r("nome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub

Private Sub txtParc_GotFocus()
SelectControl txtParc
End Sub

Private Sub txtParc_LostFocus()
If txtParc.Text = "" Then txtParc = Format(0, "##,##0.00") Else txtParc = Format(txtParc, "##,##0.00")
'Call CalcularParcelas
End Sub

Private Sub txtQuant_GotFocus()
SelectControl txtQuant
End Sub


Private Sub txtQuant_LostFocus()
Dim varValor As Currency
Dim varQuant As Integer
Dim varTotal As Currency

If txtValorProduto.Text = "" Then Exit Sub
If txtQuant.Text = "" Then txtQuant.Text = "1"

varValor = txtValorProduto.Text
varQuant = txtQuant.Text
varTotal = varValor * varQuant
txtTotalProd.Text = FormatNumber(varTotal, 2)
End Sub


Private Sub txtTotal_Change()
Call CalcularParcelas
Call Calcular_Prazo
End Sub

Private Sub txtTotal_LostFocus()
   If txtTotal.Text = "" Then txtTotal = Format(0, "##,##0.00") Else txtTotal = Format(txtTotal, "##,##0.00")
End Sub

Private Sub txtValorProduto_LostFocus()
   If txtValorProduto.Text = "" Then txtValorProduto = Format(0, "##,##0.00") Else txtValorProduto = Format(txtValorProduto, "##,##0.00")
End Sub
