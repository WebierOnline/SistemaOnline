VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Cadastro_Parcelas_Consultorio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO DE PARCELAS"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "Cadastro_Parcelas_Consultorio.frx":0000
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
      TabIndex        =   29
      Top             =   60
      Width           =   11355
      Begin VB.Image Image1 
         Height          =   645
         Left            =   420
         Picture         =   "Cadastro_Parcelas_Consultorio.frx":23D2
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CADASTRO DE PARCELAS"
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
         Left            =   1260
         TabIndex        =   30
         Top             =   300
         Width           =   4110
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
      TabIndex        =   14
      Top             =   1080
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
      TabPicture(0)   =   "Cadastro_Parcelas_Consultorio.frx":89C9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancelar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalvar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExcluir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCadastrarCliente"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNovo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Cadastro_Parcelas_Consultorio.frx":89E5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label23"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "GridConsulta"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtCONHaver"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtCONValor"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Picture3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtCONquant"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtCONtotal"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
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
         TabIndex        =   61
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
         TabIndex        =   60
         Top             =   7080
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         Height          =   1515
         Left            =   -74880
         ScaleHeight     =   1455
         ScaleWidth      =   10995
         TabIndex        =   36
         Top             =   480
         Width           =   11055
         Begin VB.Frame Frame4 
            Caption         =   "Ordem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   56
            Top             =   0
            Width           =   5535
            Begin VB.OptionButton optORDvencimento 
               Caption         =   "Vencimento"
               Height          =   195
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.OptionButton optORDnome 
               Caption         =   "Cliente"
               Height          =   195
               Left            =   1440
               TabIndex        =   57
               Top             =   240
               Width           =   1155
            End
         End
         Begin VB.PictureBox Picture6 
            Height          =   1335
            Left            =   60
            ScaleHeight     =   1275
            ScaleWidth      =   1035
            TabIndex        =   51
            Top             =   60
            Width           =   1095
            Begin VB.OptionButton optNome 
               Caption         =   "&Cliente"
               Height          =   195
               Left            =   60
               TabIndex        =   55
               Top             =   1020
               Width           =   915
            End
            Begin VB.OptionButton optIntervalo 
               Caption         =   "&Intervalo"
               Height          =   195
               Left            =   60
               TabIndex        =   54
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton optMensal 
               Caption         =   "&Mensal"
               Height          =   195
               Left            =   60
               TabIndex        =   53
               Top             =   420
               Width           =   1095
            End
            Begin VB.OptionButton optTodos 
               Caption         =   "&Todos"
               Height          =   195
               Left            =   60
               TabIndex        =   52
               Top             =   120
               Value           =   -1  'True
               Width           =   1155
            End
         End
         Begin VB.PictureBox Picture5 
            Height          =   1335
            Left            =   1200
            ScaleHeight     =   1275
            ScaleWidth      =   2535
            TabIndex        =   48
            Top             =   60
            Width           =   2595
            Begin VB.ComboBox cboCONStatus 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   660
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   60
               Width           =   1815
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Status:"
               Height          =   315
               Left            =   60
               TabIndex        =   50
               Top             =   60
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture4 
            Height          =   735
            Left            =   3840
            ScaleHeight     =   675
            ScaleWidth      =   5475
            TabIndex        =   37
            Top             =   660
            Width           =   5535
            Begin VB.TextBox txtCodClienteCons 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3960
               TabIndex        =   41
               Top             =   -60
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ComboBox cboNome 
               Height          =   315
               Left            =   720
               TabIndex        =   40
               Top             =   120
               Visible         =   0   'False
               Width           =   4695
            End
            Begin VB.ComboBox cboMES 
               Height          =   315
               ItemData        =   "Cadastro_Parcelas_Consultorio.frx":8A01
               Left            =   1200
               List            =   "Cadastro_Parcelas_Consultorio.frx":8A03
               TabIndex        =   39
               Top             =   120
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   3000
               Sorted          =   -1  'True
               TabIndex        =   38
               Top             =   120
               Visible         =   0   'False
               Width           =   1155
            End
            Begin MSMask.MaskEdBox Mask2 
               Height          =   315
               Left            =   3180
               TabIndex        =   42
               Top             =   120
               Visible         =   0   'False
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Mask1 
               Height          =   315
               Left            =   1200
               TabIndex        =   43
               Top             =   120
               Visible         =   0   'False
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label lblCONnome 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nome:"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   180
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Label lblCONint2 
               AutoSize        =   -1  'True
               Caption         =   "Data &Final:"
               Height          =   195
               Left            =   2340
               TabIndex        =   46
               Top             =   180
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.Label lblCONint1 
               AutoSize        =   -1  'True
               Caption         =   "Da&ta Inicial:"
               Height          =   195
               Left            =   300
               TabIndex        =   45
               Top             =   180
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label lblCONmes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E&scolha o mês:"
               Height          =   195
               Left            =   60
               TabIndex        =   44
               Top             =   180
               Visible         =   0   'False
               Width           =   1080
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdImprimirConsulta 
            Height          =   615
            Left            =   9420
            TabIndex        =   59
            Top             =   780
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
            MICON           =   "Cadastro_Parcelas_Consultorio.frx":8A05
            PICN            =   "Cadastro_Parcelas_Consultorio.frx":8A21
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirConsulta 
            Height          =   615
            Left            =   9420
            TabIndex        =   69
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
            MICON           =   "Cadastro_Parcelas_Consultorio.frx":8D3B
            PICN            =   "Cadastro_Parcelas_Consultorio.frx":8D57
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   7680
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Height          =   7635
         Left            =   120
         ScaleHeight     =   7575
         ScaleWidth      =   8115
         TabIndex        =   15
         Top             =   480
         Width           =   8175
         Begin VB.Frame frmCliente 
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
            TabIndex        =   26
            Top             =   60
            Width           =   7995
            Begin VB.ComboBox txtCliente 
               Height          =   315
               Left            =   1620
               TabIndex        =   2
               Top             =   480
               Width           =   6255
            End
            Begin VB.TextBox txtCodCliente 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6600
               TabIndex        =   27
               Top             =   180
               Visible         =   0   'False
               Width           =   615
            End
            Begin MSMask.MaskEdBox mskCompra 
               Height          =   315
               Left            =   120
               TabIndex        =   1
               Top             =   480
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton chameleonButton1 
               Height          =   315
               Left            =   1200
               TabIndex        =   72
               Top             =   480
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
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
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Cadastro_Parcelas_Consultorio.frx":9631
               PICN            =   "Cadastro_Parcelas_Consultorio.frx":964D
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
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
               Left            =   7140
               TabIndex        =   68
               Top             =   0
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblCompra 
               AutoSize        =   -1  'True
               Caption         =   "Inicio"
               Height          =   195
               Left            =   120
               TabIndex        =   67
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
               Height          =   195
               Left            =   1620
               TabIndex        =   28
               Top             =   240
               Width           =   525
            End
         End
         Begin VB.Frame frmParcelas 
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
            Height          =   1035
            Left            =   60
            TabIndex        =   20
            Top             =   3900
            Width           =   7995
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   315
               Left            =   3060
               TabIndex        =   11
               Top             =   540
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtParc 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1920
               TabIndex        =   10
               Top             =   540
               Width           =   1095
            End
            Begin VB.TextBox txtTotal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   8
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox txtQuantParc 
               Height          =   315
               Left            =   1380
               TabIndex        =   9
               Top             =   540
               Width           =   495
            End
            Begin ChamaleonBtn.chameleonButton cmdCalendario1 
               Height          =   315
               Left            =   4140
               TabIndex        =   71
               Top             =   540
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
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
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Cadastro_Parcelas_Consultorio.frx":BA2F
               PICN            =   "Cadastro_Parcelas_Consultorio.frx":BA4B
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lbl5 
               AutoSize        =   -1  'True
               Caption         =   "Vencimento"
               Height          =   195
               Left            =   3120
               TabIndex        =   24
               Top             =   300
               Width           =   840
            End
            Begin VB.Label lbl3 
               AutoSize        =   -1  'True
               Caption         =   "Parc"
               Height          =   195
               Left            =   1920
               TabIndex        =   23
               Top             =   300
               Width           =   330
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
               Height          =   195
               Left            =   120
               TabIndex        =   22
               Top             =   300
               Width           =   360
            End
            Begin VB.Label lbl2 
               AutoSize        =   -1  'True
               Caption         =   "Qtde"
               Height          =   195
               Left            =   1380
               TabIndex        =   21
               Top             =   300
               Width           =   345
            End
         End
         Begin VB.Frame frmReferente 
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
            Height          =   2775
            Left            =   60
            TabIndex        =   17
            Top             =   1080
            Width           =   7995
            Begin VB.ComboBox cboProdutos 
               Height          =   315
               Left            =   120
               TabIndex        =   3
               Text            =   "Tratamento Odontológico"
               Top             =   540
               Width           =   3915
            End
            Begin MSFlexGridLib.MSFlexGrid GridProdutos 
               Height          =   1395
               Left            =   120
               TabIndex        =   6
               Top             =   960
               Width           =   7755
               _ExtentX        =   13679
               _ExtentY        =   2461
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin VB.TextBox txtValorProduto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4080
               TabIndex        =   4
               Top             =   540
               Width           =   1215
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarProduto 
               Height          =   315
               Left            =   5340
               TabIndex        =   5
               ToolTipText     =   "Adiciona"
               Top             =   540
               Width           =   1215
               _ExtentX        =   2143
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
               MICON           =   "Cadastro_Parcelas_Consultorio.frx":DE2D
               PICN            =   "Cadastro_Parcelas_Consultorio.frx":DE49
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
               Left            =   6600
               TabIndex        =   7
               ToolTipText     =   "Remove"
               Top             =   540
               Width           =   1275
               _ExtentX        =   2249
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
               MICON           =   "Cadastro_Parcelas_Consultorio.frx":E1E3
               PICN            =   "Cadastro_Parcelas_Consultorio.frx":E1FF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
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
               Left            =   6960
               TabIndex        =   25
               Top             =   2400
               Width           =   915
            End
            Begin VB.Label lbl7 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
               Height          =   195
               Left            =   4080
               TabIndex        =   19
               Top             =   300
               Width           =   405
            End
            Begin VB.Label lbl6 
               AutoSize        =   -1  'True
               Caption         =   "Serviço:"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   300
               Width           =   585
            End
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   555
         Left            =   8400
         TabIndex        =   0
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "&Novo"
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
         MICON           =   "Cadastro_Parcelas_Consultorio.frx":E599
         PICN            =   "Cadastro_Parcelas_Consultorio.frx":E5B5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCadastrarCliente 
         Height          =   555
         Left            =   8400
         TabIndex        =   31
         Top             =   3480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   979
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
         MICON           =   "Cadastro_Parcelas_Consultorio.frx":F28F
         PICN            =   "Cadastro_Parcelas_Consultorio.frx":F2AB
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
         Height          =   555
         Left            =   8400
         TabIndex        =   32
         Top             =   2280
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   979
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
         MICON           =   "Cadastro_Parcelas_Consultorio.frx":FB85
         PICN            =   "Cadastro_Parcelas_Consultorio.frx":FBA1
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
         Left            =   8400
         TabIndex        =   33
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
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
         MICON           =   "Cadastro_Parcelas_Consultorio.frx":1047B
         PICN            =   "Cadastro_Parcelas_Consultorio.frx":10497
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
         Height          =   555
         Left            =   8400
         TabIndex        =   12
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Cadastro_Parcelas_Consultorio.frx":107B1
         PICN            =   "Cadastro_Parcelas_Consultorio.frx":107CD
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
         Left            =   8400
         TabIndex        =   13
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Cadastro_Parcelas_Consultorio.frx":17097
         PICN            =   "Cadastro_Parcelas_Consultorio.frx":170B3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridConsulta 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   62
         Top             =   2040
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8811
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total (LIQUIDO):"
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
         TabIndex        =   66
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
         TabIndex        =   65
         Top             =   7140
         Width           =   1755
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total (sem haveres):"
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
         TabIndex        =   64
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
         TabIndex        =   63
         Top             =   7740
         Width           =   1755
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   70
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
            TextSave        =   "11:50"
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
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Cadastro_Parcelas_Consultorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

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
      .Rows = 2
      
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
            'mudar a cor da coluna
            For i = 1 To .Rows - 1
               .Row = i
               .Col = 10
               .CellBackColor = &HC0FFFF
            Next
            
            .TextMatrix(.Rows - 1, 1) = rTabela("cod")
            .TextMatrix(.Rows - 1, 2) = rTabela("campo00")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("campo01"), "000000")
            .TextMatrix(.Rows - 1, 4) = rTabela("nome")
            .TextMatrix(.Rows - 1, 5) = rTabela("campo02")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("campo03"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("campo04"), ocMONEY)
            .TextMatrix(.Rows - 1, 8) = rTabela("var_atrazo")
            .TextMatrix(.Rows - 1, 9) = Format(rTabela("var_juros"), ocMONEY)
            .TextMatrix(.Rows - 1, 10) = Format(rTabela("var_total"), ocMONEY)
            '.TextMatrix(.Rows - 1, 11) = Format(RS!CAMPO06, "##,##0.00")
            '.TextMatrix(.Rows - 1, 12) = Format(RS!CAMPO07, "##,##0.00")
            
            If Not IsNull(rTabela("campo06")) Then
               .TextMatrix(.Rows - 1, 11) = Format(rTabela("campo06"), ocMONEY)
               .TextMatrix(.Rows - 1, 12) = Format(rTabela("campo07"), ocMONEY)
            Else
               .TextMatrix(.Rows - 1, 11) = Format(0, ocMONEY)
               .TextMatrix(.Rows - 1, 12) = Format(rTabela("var_total"), ocMONEY)
            End If
            
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
   End With
   
   txtCONValor.Text = Format(SomaGrid(GridConsulta, 10), ocMONEY)
   txtCONHaver.Text = Format(SomaGrid(GridConsulta, 11), ocMONEY)
   txtCONtotal.Text = Format(SomaGrid(GridConsulta, 12), ocMONEY)
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.Rows - 1
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
      .Cols = 4
      .Rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 0
      .ColWidth(2) = 5500
      .ColWidth(3) = 1100
      
      .TextMatrix(0, 2) = "PRODUTO"
      .TextMatrix(0, 3) = "VALOR"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 2
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            For i = 1 To .Rows - 1
               GridProdutos.Row = i
               'GridProdutos.Col = 1:   GridProdutos.CellBackColor = vbYellow
               'GridProdutos.Col = 5:   GridProdutos.CellBackColor = vbYellow
            Next
            
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_descricao")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("PRECO"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblSomaReferente.Caption = Format(SomaGrid(GridProdutos, 3), ocMONEY)
End Sub

Private Sub LimparGridProdutos()
   Dim i As Integer
   
   With GridProdutos
      .Clear
      .Cols = 4
      .Rows = 2
      
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
   
   sSQL = "SELECT codigo, 'TRATAMENTO ODONTOLOGICO' AS VAR_DESCRICAO, preco FROM pedidos_itens WHERE (cod_pedido = " & lblCodigo.Caption & ");"
   Set r = dbData.OpenRecordset(sSQL)
   FormatarGridProdutos r
   If r.State <> 0 Then r.Close
End Sub

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE a_receber SET " & _
      "cod_cliente = " & txtCodCliente.Text & ", " & _
      "cliente = '" & txtCliente.Text & "', "
   
   'Condição para atualização
   sSQL = sSQL & "WHERE (cod_receber = " & lblCodigo.Caption & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
   

'RS!PARC = IIf(txtParc.Text = "", Null, txtParc.Text)
'RS!QUANT = IIf(txtQuantParc.Text = "", Null, txtQuantParc.Text)
'RS!Valor = IIf(lblSomaReferente.Caption = "", Null, lblSomaReferente.Caption)
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

Private Sub AutoNumeracao_Receber()
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
   If cmdAlterar.Enabled = False Then lblCodigo.Caption = ""
   txtCodCliente.Text = ""
   txtCliente.Text = ""
   'If cmdAlterar.ENABLED = False Then txtCodCobrador.Text = ""
   txtTotal.Text = ""
   lblSomaReferente.Caption = ""
   txtTotal.Text = ""
   txtQuantParc.Text = ""
   txtParc.Text = ""
   mskInicio.Mask = ""
   mskInicio.Text = ""
   mskCompra.Mask = ""
   mskCompra.Text = ""
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
   cmdExibirConsulta_Click
End Sub

Private Sub cboCONStatus_GotFocus()
   PreencherComboStatus
End Sub

Private Sub cboCONStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMes.Clear
   
   For vMes = 1 To 12
      cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMes
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

Private Sub cboProdutos_GotFocus()
'   Dim sSQL As String
'   Dim r As ADODB.Recordset
   
'   cboProdutos.Clear
   
'   sSQL = "SELECT descricao FROM produtos GROUP BY descricao;"
'   Set r = dbData.OpenRecordset(sSQL)
   
'   Do While Not r.EOF
'      cboProdutos.AddItem ValidateNull(r("descricao"))
'      r.MoveNext
'   Loop
   
'   If r.State <> 0 Then r.Close
'   Set r = Nothing
   
'   moCombo.AttachTo cboProdutos
End Sub

Private Sub cboProdutos_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
  
   mskCompra = Format(varData, "dd/mm/yy")   'Exibe a data no campo
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
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_item FROM pedidos_itens;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then varCodItem = r("ultimo_item") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing

   dbData.Execute "INSERT INTO pedidos_itens (codigo, cod_pedido, cod_produto, preco, quantidade, data, maquina, tipo_venda) VALUES (" & _
         varCodItem & ", " & lblCodigo.Caption & ", 0, " & Replace(CCur(txtValorProduto.Text), ",", ".") & ", 1, CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), '" & _
         IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', 'BALCAO');"
         
   LimparObjetos_Produto
   
   cboProdutos.SetFocus
   PreencherGridProdutos
End Sub

Private Sub LimparObjetos_Produto()
   cboProdutos.Text = ""
   txtValorProduto.Text = ""
End Sub

Private Sub cmdAlterar_Click()
' On Error GoTo TrataErro
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite essa operação!", vbInformation, "Aviso do Sistema": Exit Sub


'If Not IsDate(mskDataPgto.Text) Then
'    MsgBox "Data Inválida!", vbInformation, "Aviso do Sistema"
'    txtVcto.SetFocus
'    Exit Sub
'End If

'If cboStatus.Text = "" Or txtCliente.Text = "" Or txtReferente.Text = "" Or txtValor.Text = "" Or cboSetor.Text = "" Then
'    MsgBox "Falta preencher alguns campos!", vbInformation, "Aviso do Sistema"
'    cboStatus.SetFocus
'    Exit Sub
'End If

'If cboStatus.Text = "PAGO" And mskPagamento.Text = "" Then
'    MsgBox "Falta a data de pagamento", vbInformation, "Aviso do Sistema"
'    mskPagamento.SetFocus
'    Exit Sub
'End If

'If cboStatus.Text = "PENDENTE" And IsDate(mskPagamento) Then
'    MsgBox "Falta escolher a opção no campos STATUS", vbInformation, "Aviso do Sistema"
'    cboStatus.SetFocus
'    Exit Sub
'End If

'If cboStatus.Text = "PAGO" And IsDate(mskPagamento) Then
    
'    Verificar_Caixa
'    If rs.RecordCount <> 0 Then
'        MsgBox "ESTE CAIXA JÁ ESTÁ FECHADO!", vbExclamation, "Aviso do Sistema"
'        Exit Sub
'    End If
    
    'CADASTRAR NA TABELA A_RECEBER
'    ABRIR_BD_com_Data Me.Data2
'    Data2.RecordSource = "SELECT * FROM A_RECEBER WHERE CODIGO = " & lblCodigo.Caption & ""
'    Data2.Refresh

'    If Not RS.EOF Then
'        RS.Edit
'        Atualizar_Dados
'        RS.Update
'    End If

    'CADASTRAR NA TABELA CAIXA_ENTRADA
'    Autonumeracao_Caixa
       
'    ABRIR_BD_com_Data Me.Data2
'    Data2.RecordSource = "SELECT * FROM CAIXA_ENTRADA"
'    Data2.Refresh
        
'    RS.AddNew
'    Atualizar_Dados_Caixa
'    RS.Update
'Else
    'CADASTRAR NA TABELA A_RECEBER
'    ABRIR_BD_com_Data Me.Data2
'    Data2.RecordSource = "SELECT * FROM A_RECEBER WHERE CODIGO = " & lblCodigo.Caption & ""
'    Data2.Refresh

'    If Not RS.EOF Then
'        RS.Edit
'        Atualizar_Dados
'        RS.Update
'    End If
'End If

'    Limpar_Objetos
'    Form_Load
'    Data4.Refresh

'TrataErro:
'If Err.Number = 3022 Then
'    MsgBox "DADOS DUPLICADO!" & vbCrLf & "Verifique se esta conta já está cadastrada.", vbInformation, "Aviso do Sistema"
'    Exit Sub
'End If
'If Err.Number = 3421 Then
'    MsgBox "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation, "Aviso do Sistema"
'    Exit Sub
'    cboStatus.SetFocus
'End If

End Sub

Private Sub cmdCadastrarCliente_Click()
   Clientes_Cadastro.Show 1
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

Private Sub cmdCancelar_Click()
   dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & lblCodigo.Caption & ");"
   dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & lblCodigo.Caption & ");"
   
   frmCliente.Enabled = False
   frmReferente.Enabled = False
   frmParcelas.Enabled = False
   cmdNovo.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdImprimirConsulta.Enabled = False
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
   
   For i = 1 To CInt(txtQuantParc)
      lNovoCod = Autonumeracao_Parcelas
      
      dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, status, valor) VALUES (" & _
         lNovoCod & ", " & lblCodigo.Caption & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(var_Vencimento, ocDATA) & "', 103), 0, " & _
         Replace(CCur(txtParc.Text), ",", ".") & ");"
      
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
   
   dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & lblCodigo.Caption & ");"
   dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & lblCodigo.Caption & ");"
   
   frmCliente.Enabled = False
   frmReferente.Enabled = False
   frmParcelas.Enabled = False
   cmdNovo.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdImprimirConsulta.Enabled = False
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
   
   'INDICE PARA ORGANIZAR OS DADOS
   If optORDnome.Value = True Then
      INDICE = "nome"
   ElseIf optORDvencimento.Value = True Then
      INDICE = "vencimento"
   Else
      optORDnome.Value = True
      INDICE = "nome"
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
   sql_HAVER = "(SELECT SUM(valor_haver) FROM parcelas_haver WHERE parcelas_haver.cod_parcela = parcelas.codigo)"
   '***********************************************
   
   If optTodos.Value = True Then
      lblCONmes.Visible = False
      cboMes.Visible = False
      cboAno.Visible = False
      lblCONint1.Visible = False
      Mask1.Visible = False
      lblCONint2.Visible = False
      Mask2.Visible = False
      lblCONnome.Visible = False
      cboNome.Visible = False
      
      'MOSTRAR NO GRID
      sSQL = "SELECT " & sql_ATRAZO & " AS var_atrazo, " & _
         "(" & sql_JUROS & " * " & sql_ATRAZO & ") AS var_juros, " & _
         "(parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) AS var_total, " & sql_HAVER & " AS campo06, " & _
         "((parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) - " & sql_HAVER & ") AS campo07, cliente.nome, " & _
         "parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, parcelas.status, pedidos.cod_pedido, pedidos.cod_cliente " & _
         "FROM pedidos INNER JOIN cliente ON pedidos.cod_cliente = cliente.codigo " & _
         "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido WHERE (pedidos.tipo_pedido = 'RECEBER') " & _
         var_STATUS & " ORDER BY " & INDICE
      
   ElseIf optMensal.Value = True Then
      If cboMes.Text = "" Or cboAno.Text = "" Then
         MsgBox "Selecione o MÊS e depois o ANO!", vbExclamation
         cboMes.SetFocus
         Exit Sub
      End If
      
      'DATA DE VENCIMENTO OU PAGAMENTO
      Tipo_Data = "AND (MONTH(data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(DATA) = " & cboAno & ") "
      
      'MOSTRAR NO GRID
      sSQL = "SELECT " & sql_ATRAZO & " AS var_atrazo, " & _
         "(" & sql_JUROS & " * " & sql_ATRAZO & ") AS var_juros, " & _
         "(parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) AS var_total, " & sql_HAVER & " AS campo06, " & _
         "((parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) - " & sql_HAVER & ") AS campo07, cliente.nome, " & _
         "parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, parcelas.status, pedidos.cod_pedido, pedidos.cod_cliente " & _
         "FROM pedidos INNER JOIN cliente ON pedidos.cod_cliente = cliente.codigo " & _
         "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido WHERE (pedidos.tipo_pedido = 'RECEBER') " & _
         Tipo_Data & var_STATUS & " ORDER BY " & INDICE
      
   ElseIf optIntervalo.Value = True Then
      If Not IsDate(Mask1.Text) Or Not IsDate(Mask2.Text) Then Exit Sub
      If Mask1.Text = "" Or Mask2.Text = "" Then MsgBox "Digite a DATA INICIAL e DATA FINAL!", , "Aviso do Sistema": Mask1.SetFocus:      Exit Sub
      
      'MOSTRAR NO GRID
      sSQL = "SELECT " & sql_ATRAZO & " AS var_atrazo, " & _
         "(" & sql_JUROS & " * " & sql_ATRAZO & ") AS var_juros, " & _
         "(parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) AS var_total, " & sql_HAVER & " AS campo06, " & _
         "((parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) - " & sql_HAVER & ") AS campo07, cliente.nome, " & _
         "parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, parcelas.status, pedidos.cod_pedido, pedidos.cod_cliente " & _
         "FROM pedidos INNER JOIN cliente ON pedidos.cod_cliente = cliente.codigo " & _
         "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (pedidos.tipo_pedido = 'RECEBER') AND (parcelas.data >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) " & _
         "AND (parcelas.data <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103)) " & var_STATUS & " ORDER BY " & INDICE
        
   ElseIf optNome.Value = True Then
      If txtCodClienteCons.Text = "" Then MsgBox "Selecione um cliente!", , "Aviso do Sistema": Exit Sub
      If cboNome.Text = "" Then MsgBox "Digite o nome do cliente!", vbExclamation, "Aviso do Sistema": cboNome.SetFocus:     Exit Sub
      
      'MOSTRAR NO GRID
      sSQL = "SELECT " & sql_ATRAZO & " AS var_atrazo, " & _
         "(" & sql_JUROS & " * " & sql_ATRAZO & ") AS var_juros, " & _
         "(parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) AS var_total, " & sql_HAVER & " AS campo06, " & _
         "((parcelas.valor + (" & sql_JUROS & " * " & sql_ATRAZO & ")) - " & sql_HAVER & ") AS campo07, cliente.nome, " & _
         "parcelas.codigo AS cod, pedidos.tipo_pedido AS campo00, parcelas.cod_pedido AS campo01, parcelas.numero AS campo02, " & _
         "parcelas.data AS campo03, parcelas.valor AS campo04, pedidos.pagamento AS campo05, parcelas.status, pedidos.cod_pedido, pedidos.cod_cliente " & _
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
   REL_Receber_Consulta.lblTitulo.Caption = "CONTAS À RECEBER"

   If optTodos.Value = True Then
      REL_Receber_Consulta.dfTipo.Caption = "Tipo: Todos os registros"
   ElseIf optIntervalo.Value = True Then
      REL_Receber_Consulta.dfTipo.Caption = "Tipo: Intervalo de " & Mask1.Text & " à " & Mask2.Text
   ElseIf optMensal.Value = True Then
      REL_Receber_Consulta.dfTipo.Caption = "Tipo: Mês = " & cboMes.Text & "/" & cboAno.Text
   ElseIf optNome.Value = True Then
      REL_Receber_Consulta.dfTipo.Caption = "Cliente = " & cboNome.Text
   Else
      REL_Receber_Consulta.dfTipo.Caption = "Tipo:"
   End If

   REL_Receber_Consulta.Relatorio.NomeImpressora = var_Impressora
   REL_Receber_Consulta.Relatorio.Ativar
   Unload REL_Receber_Consulta
   
   Me.Show 1
End Sub

Private Sub cmdNovo_Click()
   Limpar_Objetos
   LimparGridProdutos
   SSTab1.Tab = 0
   
   AutoNumeracao_Receber
   dbData.Execute "INSERT INTO pedidos (cod_pedido, tipo_pedido, status_pedido) VALUES (" & lblCodigo.Caption & ", 'RECEBER', 0);"
   
   frmCliente.Enabled = True
   frmReferente.Enabled = True
   frmParcelas.Enabled = True
   cmdNovo.Enabled = False
   cmdSalvar.Enabled = True
   cmdCancelar.Enabled = True
   cmdImprimirConsulta.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cmdAdicionarProduto.Enabled = True
   cmdRemoverProduto.Enabled = True
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
   
   dbData.Execute "DELETE FROM pedidos_itens WHERE (codigo = " & GridProdutos.TextMatrix(i, 1) & ");"
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

Private Sub Form_Load()
   Set moCombo = New cComboHelper
   cmdNovo.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cmdImprimirConsulta.Enabled = False
   cmdAdicionarProduto.Enabled = False
   cmdRemoverProduto.Enabled = False
   LimparGridProdutos
   PreencherComboStatus
   SSTab1.Tab = 0
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

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub gridConsulta_DblClick()
   cmdNovo.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   'cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   frmCliente.Enabled = True
   frmReferente.Enabled = True
   frmParcelas.Enabled = True
   
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
         txtQuantParc.Text = 1
         txtParc.Text = Format(r("subtotal"), ocMONEY)
         mskInicio.Text = Format(r("vencimento"), "dd/mm/yy")
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
   txtQuantParc = 1
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

Private Sub mskCompra_GotFocus()
   SelectControl mskCompra
End Sub

Private Sub mskCompra_KeyPress(KeyAscii As Integer)
   mskCompra.Mask = "##/##/##"
End Sub

Private Sub mskCompra_LostFocus()
   If IsDate(mskCompra.Text) Then
      mskInicio.Text = Format(DateAdd("m", 1, mskCompra), "dd/mm/yy")
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
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         SelectControl mskInicio
      End If
   End If
End Sub

Private Sub optIntervalo_Click()
   lblCONmes.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblCONint1.Visible = True
   Mask1.Visible = True
   lblCONint2.Visible = True
   Mask2.Visible = True
   Mask1.SetFocus
   lblCONnome.Visible = False
   cboNome.Visible = False
End Sub

Private Sub optMensal_Click()
   lblCONmes.Visible = True
   cboMes.Visible = True
   cboAno.Visible = True
   lblCONint1.Visible = False
   Mask1.Visible = False
   lblCONint2.Visible = False
   Mask2.Visible = False
   lblCONnome.Visible = False
   cboNome.Visible = False
   cboMes.SetFocus
End Sub

Private Sub optNome_Click()
   lblCONmes.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblCONint1.Visible = False
   Mask1.Visible = False
   lblCONint2.Visible = False
   Mask2.Visible = False
   lblCONnome.Visible = True
   cboNome.Visible = True
   cboNome.SetFocus
End Sub

Private Sub optORDnome_Click()
   cmdExibirConsulta_Click
End Sub

Private Sub optOrdVencimento_Click()
   cmdExibirConsulta_Click
End Sub

Private Sub optTodos_Click()
   cmdExibirConsulta_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 1 Then
      cmdImprimirConsulta.Enabled = False
      cmdSalvar.Enabled = False
      cmdAlterar.Enabled = False
      cmdExcluir.Enabled = False
   ElseIf SSTab1.Tab = 3 Then
      cmdImprimirConsulta.Enabled = True
   End If
End Sub

Private Sub Timer1_Timer()
   lblHora.Caption = Format(Time, "hh:mm")
End Sub

Private Sub txtCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   txtCliente.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      txtCliente.AddItem ValidateNull(r("nome"))
      txtCliente.ItemData(txtCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
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
   If Err.Number = 381 Then txtCodCliente.Text = ""
End Sub

Private Sub TxtCodCliente_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cmdExcluir.Enabled = True Then
      sSQL = "SELECT codigo, nome FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then txtCliente.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub

Private Sub txtParc_GotFocus()
   Dim var_Total As Currency
   Dim var_Quant As Integer
   Dim var_Parc As Currency
   
   SelectControl txtParc
   
   If txtQuantParc.Text = "" Then txtQuantParc.Text = "1"
   If txtTotal.Text = "" Then txtTotal.Text = "0"
   
   var_Total = txtTotal.Text
   var_Quant = txtQuantParc.Text
   var_Parc = var_Total / var_Quant
   txtParc.Text = Format(var_Parc, ocMONEY)
End Sub

Private Sub txtParc_LostFocus()
   If txtParc.Text = "" Then txtParc = Format(0, "##,##0.00") Else txtParc = Format(txtParc, "##,##0.00")
End Sub

Private Sub txtQuantParc_LostFocus()
   If txtQuantParc.Text = "" Then txtQuantParc = 1
End Sub

Private Sub txtTotal_Change()
   'lblSomaParc.Caption = txtTotal.Text
End Sub

Private Sub txtTotal_LostFocus()
   If txtTotal.Text = "" Then txtTotal = Format(0, "##,##0.00") Else txtTotal = Format(txtTotal, "##,##0.00")
End Sub

Private Sub txtValorProduto_LostFocus()
   If txtValorProduto.Text = "" Then txtValorProduto = Format(0, "##,##0.00") Else txtValorProduto = Format(txtValorProduto, "##,##0.00")
End Sub
