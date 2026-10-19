VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Produtos_Entrada 
   Caption         =   "Produtos_Entrada"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   15960
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Frm_NF 
      Height          =   9795
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   17277
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Nota Fiscal"
      TabPicture(0)   =   "Produtos_Entrada.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Tab_Entrada"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtCodPedido"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCodNota"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmNota"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "Produtos_Entrada.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblSomaNota"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblTotalNota"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label12"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label17"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label18"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label22"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label7"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label21"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label20"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label19"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label16"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label15"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label11"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label14"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "GridNotas"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Frame3"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin VB.Frame frmNota 
         Caption         =   "Nota Fiscal"
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
         Left            =   -74880
         TabIndex        =   78
         Top             =   360
         Width           =   15615
         Begin VB.ComboBox cboFornecedor 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1560
            TabIndex        =   87
            Top             =   540
            Width           =   6375
         End
         Begin VB.TextBox TxtCodFornecedor 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4740
            MaxLength       =   10
            TabIndex        =   86
            Top             =   300
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txtNumNota 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   85
            Top             =   540
            Width           =   1380
         End
         Begin VB.TextBox txtCodEntrada 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   12900
            TabIndex        =   83
            TabStop         =   0   'False
            ToolTipText     =   "Cód do Entrada"
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtItem 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   12000
            TabIndex        =   82
            TabStop         =   0   'False
            ToolTipText     =   "Item"
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtCodUsuario 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   11100
            TabIndex        =   81
            TabStop         =   0   'False
            ToolTipText     =   "Item"
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtTotaldaNota 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   9300
            TabIndex        =   80
            Top             =   540
            Width           =   1395
         End
         Begin VB.TextBox txtChave 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   10740
            TabIndex        =   79
            Top             =   540
            Width           =   4455
         End
         Begin MSMask.MaskEdBox mskEmissao 
            Height          =   315
            Left            =   7980
            TabIndex        =   84
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdCopiarChave 
            Height          =   315
            Left            =   15180
            TabIndex        =   88
            Top             =   540
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
            MICON           =   "Produtos_Entrada.frx":0038
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
            Height          =   195
            Left            =   1560
            TabIndex        =   93
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº NF"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   92
            Top             =   300
            Width           =   435
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. Emissão"
            Height          =   195
            Left            =   7980
            TabIndex        =   91
            Top             =   300
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor da Nota"
            Height          =   195
            Left            =   9300
            TabIndex        =   90
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chave"
            Height          =   195
            Left            =   10740
            TabIndex        =   89
            Top             =   300
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   54
         Top             =   8280
         Width           =   15615
         Begin VB.Frame Frame4 
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
            Height          =   735
            Left            =   5520
            TabIndex        =   60
            Top             =   240
            Width           =   6975
            Begin VB.ComboBox cboConNotaAno 
               Height          =   315
               Left            =   2640
               TabIndex        =   69
               Top             =   300
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.ComboBox cboConNotaMes 
               Height          =   315
               Left            =   600
               TabIndex        =   68
               Top             =   300
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txtConNotaNumNota 
               Height          =   315
               Left            =   960
               TabIndex        =   67
               Top             =   300
               Visible         =   0   'False
               Width           =   5895
            End
            Begin VB.ComboBox cboConNotaCliente 
               Height          =   315
               Left            =   1020
               TabIndex        =   66
               Top             =   300
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.TextBox txtConNotaCodCliente 
               Height          =   315
               Left            =   4860
               TabIndex        =   63
               Top             =   300
               Visible         =   0   'False
               Width           =   495
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal1 
               Height          =   315
               Left            =   1560
               TabIndex        =   61
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
               MICON           =   "Produtos_Entrada.frx":0054
               PICN            =   "Produtos_Entrada.frx":0070
               PICH            =   "Produtos_Entrada.frx":23C3
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal2 
               Height          =   315
               Left            =   3360
               TabIndex        =   62
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
               MICON           =   "Produtos_Entrada.frx":4716
               PICN            =   "Produtos_Entrada.frx":4732
               PICH            =   "Produtos_Entrada.frx":6A85
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskConNotaFinal 
               Height          =   315
               Left            =   2460
               TabIndex        =   64
               Top             =   300
               Visible         =   0   'False
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskConNotaInicial 
               Height          =   315
               Left            =   660
               TabIndex        =   65
               Top             =   300
               Visible         =   0   'False
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin VB.Label lblConNotaNumNota 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Num. Nota:"
               Height          =   195
               Left            =   120
               TabIndex        =   75
               Top             =   300
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.Label lblConNotaFinal 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Final:"
               Height          =   195
               Left            =   180
               TabIndex        =   74
               Top             =   300
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lblConNotaInicial 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicial:"
               Height          =   195
               Left            =   1980
               TabIndex        =   73
               Top             =   300
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label lblConNotaCliente 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fornecedor:"
               Height          =   195
               Left            =   120
               TabIndex        =   72
               Top             =   300
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblConNotaMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   180
               TabIndex        =   71
               Top             =   300
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblConNotaAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   2220
               TabIndex        =   70
               Top             =   300
               Visible         =   0   'False
               Width           =   330
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
            Height          =   735
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   5295
            Begin VB.ComboBox cboFiltroNota 
               Height          =   315
               Left            =   660
               TabIndex        =   57
               Top             =   300
               Width           =   1995
            End
            Begin VB.ComboBox cboOrdem 
               Height          =   315
               Left            =   3360
               TabIndex        =   56
               Top             =   300
               Width           =   1755
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Filtro:"
               Height          =   195
               Left            =   180
               TabIndex        =   59
               Top             =   300
               Width           =   375
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ordem:"
               Height          =   195
               Left            =   2760
               TabIndex        =   58
               Top             =   300
               Width           =   510
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirConNotas 
            Height          =   555
            Left            =   12540
            TabIndex        =   76
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   979
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
            MICON           =   "Produtos_Entrada.frx":8DD8
            PICN            =   "Produtos_Entrada.frx":8DF4
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
            Height          =   555
            Left            =   14040
            TabIndex        =   77
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   979
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
            MICON           =   "Produtos_Entrada.frx":96CE
            PICN            =   "Produtos_Entrada.frx":96EA
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
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   34
         Top             =   6120
         Width           =   15
         Begin VB.Frame frmCusto 
            Caption         =   "Custo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4000
            Left            =   200
            TabIndex        =   52
            Top             =   0
            Width           =   4000
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Custo Bruto"
               Height          =   195
               Left            =   60
               TabIndex        =   53
               Top             =   240
               Width           =   825
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Custo"
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
            Left            =   720
            TabIndex        =   47
            Top             =   180
            Width           =   1395
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   60
               TabIndex        =   48
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant."
               Height          =   195
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Varejo - À vista"
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
            Left            =   4260
            TabIndex        =   44
            Top             =   180
            Width           =   1995
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1080
               TabIndex        =   46
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Varejo - À Prazo"
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
            Left            =   6300
            TabIndex        =   41
            Top             =   180
            Width           =   1995
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1080
               TabIndex        =   42
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Atacado - À Vista"
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
            Left            =   8340
            TabIndex        =   38
            Top             =   180
            Width           =   1995
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1080
               TabIndex        =   40
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Atacado - À Prazo"
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
            Left            =   10380
            TabIndex        =   35
            Top             =   180
            Width           =   1995
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lucro %"
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   570
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1080
               TabIndex        =   36
               Top             =   240
               Width           =   360
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionar 
            Height          =   315
            Left            =   0
            TabIndex        =   50
            ToolTipText     =   "Adiciona"
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
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
            MICON           =   "Produtos_Entrada.frx":B47C
            PICN            =   "Produtos_Entrada.frx":B498
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdRemover 
            Height          =   315
            Left            =   10320
            TabIndex        =   51
            ToolTipText     =   "Remove"
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
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
            MICON           =   "Produtos_Entrada.frx":B832
            PICN            =   "Produtos_Entrada.frx":B84E
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
      Begin VB.TextBox txtCodNota 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -60180
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Cód da Nota"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtCodPedido 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -61080
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Cód do Pedido"
         Top             =   360
         Width           =   855
      End
      Begin TabDlg.SSTab Tab_Entrada 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   1
         Top             =   1380
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   14631
         _Version        =   393216
         TabHeight       =   520
         TabMaxWidth     =   4057
         TabCaption(0)   =   "Produtos"
         TabPicture(0)   =   "Produtos_Entrada.frx":BBE8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmItens"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Totais da nota"
         TabPicture(1)   =   "Produtos_Entrada.frx":BC04
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Parcelas e Transportes"
         TabPicture(2)   =   "Produtos_Entrada.frx":BC20
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).Control(1)=   "Frame18"
         Tab(2).Control(2)=   "Frame13"
         Tab(2).ControlCount=   3
         Begin VB.Frame frmItens 
            Caption         =   "Itens da Nota de Entrada (Compra)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7815
            Left            =   120
            TabIndex        =   113
            Top             =   360
            Width           =   15375
            Begin MSFlexGridLib.MSFlexGrid GridNotasItens 
               Height          =   7275
               Left            =   120
               TabIndex        =   114
               Top             =   240
               Width           =   15135
               _ExtentX        =   26696
               _ExtentY        =   12832
               _Version        =   393216
               BackColor       =   16777215
               SelectionMode   =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label lblValorNota 
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
               Left            =   15060
               TabIndex        =   115
               Top             =   7560
               Width           =   225
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Totais da Nota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2715
            Left            =   -74880
            TabIndex        =   30
            Top             =   360
            Width           =   15375
            Begin MSFlexGridLib.MSFlexGrid grdTotaisNota 
               Height          =   2355
               Left            =   60
               TabIndex        =   31
               Top             =   300
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   4154
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Transporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2715
            Left            =   -74940
            TabIndex        =   7
            Top             =   360
            Width           =   8355
            Begin VB.ComboBox cboModFrete 
               Height          =   315
               Left            =   120
               TabIndex        =   25
               Top             =   540
               Width           =   3375
            End
            Begin VB.ComboBox cboTransporte 
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   1140
               Width           =   6315
            End
            Begin VB.TextBox txtCodTransporte 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5280
               MaxLength       =   50
               TabIndex        =   23
               Top             =   840
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Frame Frame12 
               Caption         =   "Volume"
               Height          =   1095
               Left            =   60
               TabIndex        =   10
               Top             =   1500
               Width           =   8235
               Begin VB.TextBox txtVolPesoBruto 
                  Height          =   315
                  Left            =   4920
                  MaxLength       =   50
                  TabIndex        =   16
                  Top             =   480
                  Width           =   1545
               End
               Begin VB.TextBox txtVolQuant 
                  Height          =   315
                  Left            =   60
                  MaxLength       =   50
                  TabIndex        =   15
                  Top             =   480
                  Width           =   825
               End
               Begin VB.TextBox txtVolEspecie 
                  Height          =   315
                  Left            =   900
                  MaxLength       =   50
                  TabIndex        =   14
                  Top             =   480
                  Width           =   1245
               End
               Begin VB.TextBox txtVolMarca 
                  Height          =   315
                  Left            =   2160
                  MaxLength       =   50
                  TabIndex        =   13
                  Top             =   480
                  Width           =   1665
               End
               Begin VB.TextBox txtVolNumeracao 
                  Height          =   315
                  Left            =   3840
                  MaxLength       =   50
                  TabIndex        =   12
                  Top             =   480
                  Width           =   1065
               End
               Begin VB.TextBox txtVolPesoLiquido 
                  Height          =   315
                  Left            =   6480
                  MaxLength       =   50
                  TabIndex        =   11
                  Top             =   480
                  Width           =   1545
               End
               Begin VB.Label Label33 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Espécie"
                  Height          =   195
                  Left            =   900
                  TabIndex        =   22
                  Top             =   240
                  Width           =   570
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Peso liquido"
                  Height          =   195
                  Left            =   6480
                  TabIndex        =   21
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Peso bruto"
                  Height          =   195
                  Left            =   4920
                  TabIndex        =   20
                  Top             =   240
                  Width           =   765
               End
               Begin VB.Label Label36 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Qtde Vol."
                  Height          =   195
                  Left            =   60
                  TabIndex        =   19
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label Label40 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Marca"
                  Height          =   195
                  Left            =   2160
                  TabIndex        =   18
                  Top             =   240
                  Width           =   450
               End
               Begin VB.Label Label42 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Numeração"
                  Height          =   195
                  Left            =   3840
                  TabIndex        =   17
                  Top             =   240
                  Width           =   825
               End
            End
            Begin VB.TextBox txtPlacaUF 
               Height          =   315
               Left            =   7800
               MaxLength       =   2
               TabIndex        =   9
               Top             =   1140
               Width           =   465
            End
            Begin VB.TextBox txtPlaca 
               Height          =   315
               Left            =   6480
               MaxLength       =   8
               TabIndex        =   8
               Top             =   1140
               Width           =   1245
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Frete"
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   300
               Width           =   945
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transportadora"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   900
               Width           =   1080
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UF"
               Height          =   195
               Left            =   7800
               TabIndex        =   27
               Top             =   900
               Width           =   210
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Placa"
               Height          =   195
               Left            =   6480
               TabIndex        =   26
               Top             =   900
               Width           =   405
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Parcelas"
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
            Left            =   -66540
            TabIndex        =   5
            Top             =   360
            Width           =   7035
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
               Height          =   975
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   6795
               _ExtentX        =   11986
               _ExtentY        =   1720
               _Version        =   393216
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Informações"
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
            Left            =   -66540
            TabIndex        =   2
            Top             =   1740
            Width           =   7035
            Begin VB.TextBox txtInformacao 
               Alignment       =   1  'Right Justify
               Height          =   675
               Left            =   120
               MaxLength       =   50
               TabIndex        =   3
               Top             =   540
               Width           =   6780
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Informações Adicionais"
               Height          =   195
               Left            =   120
               TabIndex        =   4
               Top             =   300
               Width           =   1635
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridNotas 
         Height          =   7515
         Left            =   120
         TabIndex        =   94
         Top             =   480
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   13256
         _Version        =   393216
         TextStyleFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do Frete"
         Height          =   195
         Left            =   10560
         TabIndex        =   112
         Top             =   2040
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Especie"
         Height          =   195
         Left            =   2310
         TabIndex        =   111
         Top             =   2700
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         Height          =   195
         Left            =   2460
         TabIndex        =   110
         Top             =   1440
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Despesas"
         Height          =   195
         Left            =   6270
         TabIndex        =   109
         Top             =   1440
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seguro"
         Height          =   195
         Left            =   5130
         TabIndex        =   108
         Top             =   1440
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B. ICMS Sub."
         Height          =   195
         Left            =   8595
         TabIndex        =   107
         Top             =   1440
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B. ICMS"
         Height          =   195
         Left            =   7395
         TabIndex        =   106
         Top             =   1440
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fatura"
         Height          =   195
         Left            =   9720
         TabIndex        =   105
         Top             =   1200
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transportadora"
         Height          =   195
         Left            =   4680
         TabIndex        =   104
         Top             =   2040
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         Height          =   195
         Left            =   9555
         TabIndex        =   103
         Top             =   2040
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso liquido"
         Height          =   195
         Left            =   6375
         TabIndex        =   102
         Top             =   2700
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso bruto"
         Height          =   195
         Left            =   5430
         TabIndex        =   101
         Top             =   2700
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde Vol."
         Height          =   195
         Left            =   1440
         TabIndex        =   100
         Top             =   2700
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         Height          =   195
         Left            =   3315
         TabIndex        =   99
         Top             =   2700
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   195
         Left            =   4425
         TabIndex        =   98
         Top             =   2700
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa"
         Height          =   195
         Left            =   8700
         TabIndex        =   97
         Top             =   2040
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblTotalNota 
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
         Left            =   120
         TabIndex        =   96
         Top             =   8040
         Width           =   225
      End
      Begin VB.Label lblSomaNota 
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
         Left            =   15480
         TabIndex        =   95
         Top             =   8040
         Width           =   225
      End
   End
End
Attribute VB_Name = "Produtos_Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tela de CONSULTA das notas ja processadas (produtos_entrada / produtos_entrada_itens).
'Mesma formatacao/campos do Entrada_Estoque.Frm (que trabalha com EntradaEstoque/
'EntradaEstoqueItens - o XML "cru", antes de ser levado ao estoque) - aqui mostra os
'dados ja efetivados. So consulta - sem Importar XML/Adicionar/Excluir/etc.
'
'PENDENTE (nao fazem parte dessa 1a leva, flagueado pro usuario): aba "Parcelas e
'Transportes" -> grid de Parcelas (Frame18/MSFlexGrid1, sem fonte de dados clara - nem
'o Entrada_Estoque.Frm original popula isso) e o bloco de custo/margem (Frame6, calculo
'varejo/atacado - nao faz sentido numa tela so-consulta). txtCodPedido/txtItem/
'txtCodUsuario tambem ficaram sem popular (sem mapeamento obvio em produtos_entrada).

Private moCombo As cComboHelper
Private TbEntrada As ADODB.Recordset
Private TbConsulta As ADODB.Recordset

Private Sub Form_Load()
    Set moCombo = New cComboHelper
    Frm_NF.Tab = 1   'abre direto na aba de Consulta - essa tela nao cria nota nova

    cboFiltroNota_GotFocus
    cboFiltroNota.ListIndex = 5   'MENSAL

    cboOrdem_GotFocus
    cboOrdem.ListIndex = 0

    cboConNotaMes_GotFocus
    cboConNotaMes.ListIndex = Month(Date) - 1
    cboConNotaAno.Text = Year(Date)
    cmdExibirConNotas_Click
End Sub

'==============================================================
' FILTROS (identico ao Entrada_Estoque.Frm)
'==============================================================
Private Sub cboFiltroNota_Click()
    cboFiltroNota_LostFocus
End Sub

Private Sub cboFiltroNota_GotFocus()
    Dim vTexto As String
    vTexto = cboFiltroNota.Text
    cboFiltroNota.Clear
    cboFiltroNota.AddItem "TODAS"
    cboFiltroNota.AddItem "NUM. NOTA"
    cboFiltroNota.AddItem "CHAVE DE ACESSO"
    cboFiltroNota.AddItem "FORNECEDOR"
    cboFiltroNota.AddItem "PRODUTO"
    cboFiltroNota.AddItem "PERIODO"
    cboFiltroNota.AddItem "MENSAL"
    cboFiltroNota.Text = vTexto
    moCombo.AttachTo cboFiltroNota
End Sub

Private Sub cboFiltroNota_LostFocus()
    If cboFiltroNota.Text = "TODAS" Then
        lblConNotaNumNota.Visible = False
        txtConNotaNumNota.Visible = False
        lblConNotaCliente.Visible = False
        cboConNotaCliente.Visible = False
        txtConNotaCodCliente.Visible = False
        lblConNotaAno.Visible = False
        lblConNotaMes.Visible = False
        cboConNotaAno.Visible = False
        cboConNotaMes.Visible = False
        lblConNotaInicial.Visible = False
        lblConNotaFinal.Visible = False
        mskConNotaInicial.Visible = False
        mskConNotaFinal.Visible = False
        cmdConNotaCal1.Visible = False
        cmdConNotaCal2.Visible = False
    ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
        lblConNotaNumNota.Caption = "Numero da Nota"
        lblConNotaNumNota.Visible = True
        txtConNotaNumNota.Visible = True
        txtConNotaNumNota.Width = 1875
        lblConNotaCliente.Visible = False
        cboConNotaCliente.Visible = False
        txtConNotaCodCliente.Visible = False
        lblConNotaAno.Visible = False
        lblConNotaMes.Visible = False
        cboConNotaAno.Visible = False
        cboConNotaMes.Visible = False
        lblConNotaInicial.Visible = False
        lblConNotaFinal.Visible = False
        mskConNotaInicial.Visible = False
        mskConNotaFinal.Visible = False
        cmdConNotaCal1.Visible = False
        cmdConNotaCal2.Visible = False
        txtConNotaNumNota.SetFocus
    ElseIf cboFiltroNota.Text = "CHAVE DE ACESSO" Then
        lblConNotaNumNota.Caption = "Chave:"
        lblConNotaNumNota.Visible = True
        txtConNotaNumNota.Visible = True
        txtConNotaNumNota.Width = 5895
        lblConNotaCliente.Visible = False
        cboConNotaCliente.Visible = False
        txtConNotaCodCliente.Visible = False
        lblConNotaAno.Visible = False
        lblConNotaMes.Visible = False
        cboConNotaAno.Visible = False
        cboConNotaMes.Visible = False
        lblConNotaInicial.Visible = False
        lblConNotaFinal.Visible = False
        mskConNotaInicial.Visible = False
        mskConNotaFinal.Visible = False
        cmdConNotaCal1.Visible = False
        cmdConNotaCal2.Visible = False
        txtConNotaNumNota.SetFocus
    ElseIf cboFiltroNota.Text = "FORNECEDOR" Then
        lblConNotaCliente.Visible = True
        cboConNotaCliente.Visible = True
        lblConNotaNumNota.Visible = False
        txtConNotaNumNota.Visible = False
        txtConNotaCodCliente.Visible = False
        lblConNotaAno.Visible = False
        lblConNotaMes.Visible = False
        cboConNotaAno.Visible = False
        cboConNotaMes.Visible = False
        lblConNotaInicial.Visible = False
        lblConNotaFinal.Visible = False
        mskConNotaInicial.Visible = False
        mskConNotaFinal.Visible = False
        cmdConNotaCal1.Visible = False
        cmdConNotaCal2.Visible = False
        cboConNotaCliente.SetFocus
    ElseIf cboFiltroNota.Text = "PRODUTO" Then
        lblConNotaNumNota.Caption = "Codigo do Produto"
        lblConNotaNumNota.Visible = True
        txtConNotaNumNota.Visible = True
        txtConNotaNumNota.Width = 1875
        lblConNotaCliente.Visible = False
        cboConNotaCliente.Visible = False
        txtConNotaCodCliente.Visible = False
        lblConNotaAno.Visible = False
        lblConNotaMes.Visible = False
        cboConNotaAno.Visible = False
        cboConNotaMes.Visible = False
        lblConNotaInicial.Visible = False
        lblConNotaFinal.Visible = False
        mskConNotaInicial.Visible = False
        mskConNotaFinal.Visible = False
        cmdConNotaCal1.Visible = False
        cmdConNotaCal2.Visible = False
        'SEM SetFocus aqui (diferente dos outros ramos) - esse ramo tambem e usado
        'via FiltrarPorProduto, chamado de fora ANTES do .Show (form ainda nao esta
        'visivel nesse momento - SetFocus nessa hora da erro 5)
    ElseIf cboFiltroNota.Text = "PERIODO" Then
        lblConNotaInicial.Visible = True
        lblConNotaFinal.Visible = True
        mskConNotaInicial.Visible = True
        mskConNotaFinal.Visible = True
        cmdConNotaCal1.Visible = True
        cmdConNotaCal2.Visible = True
        lblConNotaNumNota.Visible = False
        txtConNotaNumNota.Visible = False
        lblConNotaCliente.Visible = False
        cboConNotaCliente.Visible = False
        txtConNotaCodCliente.Visible = False
        lblConNotaAno.Visible = False
        lblConNotaMes.Visible = False
        cboConNotaAno.Visible = False
        cboConNotaMes.Visible = False
        If mskConNotaInicial.Visible = True Then mskConNotaInicial.SetFocus
    ElseIf cboFiltroNota.Text = "MENSAL" Then
        lblConNotaAno.Visible = True
        lblConNotaMes.Visible = True
        cboConNotaAno.Visible = True
        cboConNotaMes.Visible = True
        lblConNotaNumNota.Visible = False
        txtConNotaNumNota.Visible = False
        lblConNotaCliente.Visible = False
        cboConNotaCliente.Visible = False
        txtConNotaCodCliente.Visible = False
        lblConNotaInicial.Visible = False
        lblConNotaFinal.Visible = False
        mskConNotaInicial.Visible = False
        mskConNotaFinal.Visible = False
        cmdConNotaCal1.Visible = False
        cmdConNotaCal2.Visible = False
        If cboConNotaMes.Visible = True Then cboConNotaMes.SetFocus
    End If
End Sub

Private Sub cboOrdem_GotFocus()
    Dim vTexto As String
    vTexto = cboOrdem.Text
    cboOrdem.Clear
    cboOrdem.AddItem "EMISSAO"
    cboOrdem.AddItem "NUM. NOTA"
    cboOrdem.AddItem "FORNECEDOR"
    cboOrdem.Text = vTexto
    moCombo.AttachTo cboOrdem
End Sub

Private Sub cboConNotaAno_GotFocus()
    Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
    Dim i As Integer
    cboConNotaAno.Clear
    iAno = Year(Date)
    FirstYear = iAno - 2
    LastYear = iAno + 2
    For i = FirstYear To LastYear
        cboConNotaAno.AddItem i
    Next
    moCombo.AttachTo cboConNotaAno
End Sub

Private Sub cboConNotaMes_GotFocus()
    If cboConNotaMes.ListCount = 0 Then
        cboConNotaMes.AddItem "Janeiro"
        cboConNotaMes.AddItem "Fevereiro"
        cboConNotaMes.AddItem "Marco"
        cboConNotaMes.AddItem "Abril"
        cboConNotaMes.AddItem "Maio"
        cboConNotaMes.AddItem "Junho"
        cboConNotaMes.AddItem "Julho"
        cboConNotaMes.AddItem "Agosto"
        cboConNotaMes.AddItem "Setembro"
        cboConNotaMes.AddItem "Outubro"
        cboConNotaMes.AddItem "Novembro"
        cboConNotaMes.AddItem "Dezembro"
    End If
    moCombo.AttachTo cboConNotaMes
End Sub

Private Sub cboConNotaCliente_GotFocus()
    Dim sSQL As String
    Dim r As ADODB.Recordset
    cboConNotaCliente.Clear
    sSQL = "SELECT DISTINCT RAZAO, codigo FROM FORNECEDOR ORDER BY RAZAO;"
    Set r = dbData.OpenRecordset(sSQL)
    Do While Not r.EOF
        cboConNotaCliente.AddItem r("RAZAO")
        cboConNotaCliente.ItemData(cboConNotaCliente.NewIndex) = r("codigo")
        r.MoveNext
    Loop
    If r.State <> 0 Then r.Close
    Set r = Nothing
    moCombo.AttachTo cboConNotaCliente
End Sub

Private Sub cboConNotaCliente_LostFocus()
    On Error GoTo TrataErro
    If cboConNotaCliente.Text = "" Then txtConNotaCodCliente.Text = "": Exit Sub
    If cboConNotaCliente.ListIndex = -1 Then txtConNotaCodCliente.Text = "": Exit Sub
    txtConNotaCodCliente = cboConNotaCliente.ItemData(cboConNotaCliente.ListIndex)
TrataErro:
    If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cmdConNotaCal1_Click()
    Dim varData As Variant
    Dim fCal As Calendario
    varData = Empty
    Set fCal = New Calendario
    fCal.Show vbModal
    varData = fCal.DateSelected
    Unload fCal
    Set fCal = Nothing
    If Not IsDate(varData) Then Exit Sub
    If varData = 0 Then Exit Sub
    mskConNotaInicial = Format(varData, "dd/mm/yy")
End Sub

Private Sub cmdConNotaCal2_Click()
    Dim varData As Variant
    Dim fCal As Calendario
    varData = Empty
    Set fCal = New Calendario
    fCal.Show vbModal
    varData = fCal.DateSelected
    Unload fCal
    Set fCal = Nothing
    If Not IsDate(varData) Then Exit Sub
    If varData = 0 Then Exit Sub
    mskConNotaFinal = Format(varData, "dd/mm/yy")
End Sub

Private Sub mskConNotaInicial_GotFocus()
    SelectControl mskConNotaInicial
End Sub

Private Sub mskConNotaInicial_KeyPress(KeyAscii As Integer)
    mskConNotaInicial.Mask = "##/##/##"
End Sub

Private Sub mskConNotaInicial_LostFocus()
    If mskConNotaInicial.Text = "" Or mskConNotaInicial.Text = "__/__/__" Then
        mskConNotaInicial.Mask = ""
        mskConNotaInicial.Text = ""
        Exit Sub
    Else
        If Not IsDate(mskConNotaInicial.Text) Then
            ShowMsg "DATA INVALIDA!" & vbCrLf & "A data digitada esta incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
            mskConNotaInicial.SetFocus
            SelectControl mskConNotaInicial
        End If
    End If
End Sub

Private Sub mskConNotaFinal_GotFocus()
    SelectControl mskConNotaFinal
End Sub

Private Sub mskConNotaFinal_KeyPress(KeyAscii As Integer)
    mskConNotaFinal.Mask = "##/##/##"
End Sub

Private Sub mskConNotaFinal_LostFocus()
    If mskConNotaFinal.Text = "" Or mskConNotaFinal.Text = "__/__/__" Then
        mskConNotaFinal.Mask = ""
        mskConNotaFinal.Text = ""
        Exit Sub
    Else
        If Not IsDate(mskConNotaFinal.Text) Then
            ShowMsg "DATA INVALIDA!" & vbCrLf & "A data digitada esta incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
            mskConNotaFinal.SetFocus
            SelectControl mskConNotaFinal
        End If
    End If
End Sub

Private Sub mskEmissao_KeyPress(KeyAscii As Integer)
    mskEmissao.Mask = "##/##/####"
End Sub

'==============================================================
' LISTAGEM DE NOTAS (produtos_entrada no lugar de EntradaEstoque - sem SITUACAO/
' PENDENCIAS, que nao fazem sentido aqui: todo item de produtos_entrada_itens ja
' foi efetivamente adicionado ao estoque, nao existe "pendente" nessa tabela)
'==============================================================
Private Sub cmdExibirConNotas_Click()
    Dim totalRegistros As Long
    On Error GoTo ErrLoad

    Dim INDICE As String
    If cboOrdem.Text = "EMISSAO" Then
        INDICE = "DataEmissao DESC;"
    ElseIf cboOrdem.Text = "NUM. NOTA" Then
        INDICE = "NumeroNota;"
    ElseIf cboOrdem.Text = "FORNECEDOR" Then
        INDICE = "RazaoSocial;"
    Else
        INDICE = "DataEmissao DESC;"
    End If

    Dim sBase As String
    sBase = "SELECT CODIGO, CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
            "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitacao' END) END) END) AS Status " & _
            "FROM produtos_entrada "

    If cboFiltroNota.Text = "TODAS" Then
        RsOpen TbConsulta, sBase & "ORDER BY " & INDICE
    ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
        If txtConNotaNumNota.Text = "" Then Exit Sub
        RsOpen TbConsulta, sBase & "WHERE NumeroNota = " & txtConNotaNumNota & " ORDER BY " & INDICE
    ElseIf cboFiltroNota.Text = "CHAVE DE ACESSO" Then
        If txtConNotaNumNota.Text = "" Then Exit Sub
        RsOpen TbConsulta, sBase & "WHERE ChavedeAcesso = '" & txtConNotaNumNota & "' ORDER BY " & INDICE
    ElseIf cboFiltroNota.Text = "FORNECEDOR" Then
        If txtConNotaCodCliente.Text = "" Then Exit Sub
        RsOpen TbConsulta, sBase & "WHERE CodigoCorrentista = " & txtConNotaCodCliente.Text & " ORDER BY " & INDICE
    ElseIf cboFiltroNota.Text = "PRODUTO" Then
        'Aqui precisa de JOIN com produtos_entrada_itens - nao da pra usar o sBase generico
        If txtConNotaNumNota.Text = "" Then Exit Sub
        RsOpen TbConsulta, "SELECT DISTINCT pe.CODIGO, pe.CodigoNota, pe.NumeroNota, pe.DataEmissao, pe.RazaoSocial, pe.ValorNota, " & _
                "(CASE WHEN pe.Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN pe.Enviada = 1 AND pe.Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN pe.Enviada = 1 AND pe.Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitacao' END) END) END) AS Status " & _
                "FROM produtos_entrada pe " & _
                "INNER JOIN produtos_entrada_itens pei ON pei.codigo_entrada = pe.CODIGO " & _
                "WHERE pei.CodigoProduto = " & txtConNotaNumNota & " ORDER BY " & INDICE
    ElseIf cboFiltroNota.Text = "PERIODO" Then
        If mskConNotaInicial.Text = "" Then Exit Sub
        RsOpen TbConsulta, sBase & "WHERE (DataEmissao >= CONVERT(DATETIME, '" & Format(mskConNotaInicial.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskConNotaFinal.Text, ocDATA) & "', 103)) ORDER BY " & INDICE
    ElseIf cboFiltroNota.Text = "MENSAL" Then
        If cboConNotaMes.Text = "" Then Exit Sub
        Dim vIndMes As Integer
        vIndMes = cboConNotaMes.ListIndex + 1
        RsOpen TbConsulta, sBase & "WHERE (MONTH(DataEmissao) = " & vIndMes & ") AND (YEAR(DataEmissao) = " & cboConNotaAno & ") ORDER BY CODIGO DESC"
    End If

    If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount
    lblTotalNota.Caption = Format(totalRegistros, "00")

    FormatarGridNotas TbConsulta

    Exit Sub

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbConsulta = Nothing
End Sub

Public Sub FormatarGridNotas(rTabela As ADODB.Recordset)
    Dim i As Integer

    With GridNotas
        .Visible = False
        .Redraw = False

        .Clear
        .Cols = 7
        .rows = 2

        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 5500
        .ColWidth(5) = 1350
        .ColWidth(6) = 1500

        .TextMatrix(0, 2) = "NUM NF"
        .TextMatrix(0, 3) = "EMISSAO"
        .TextMatrix(0, 4) = "FORNECEDOR"
        .TextMatrix(0, 5) = "VALOR"
        .TextMatrix(0, 6) = "STATUS"

        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellFontBold = True
            .CellAlignment = flexAlignCenterCenter
        Next

        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 1
        .ColAlignment(5) = 6
        .ColAlignment(6) = 4

        If Not rTabela Is Nothing Then
            Do While Not rTabela.EOF
                .TextMatrix(.rows - 1, 1) = rTabela("CODIGO")
                .TextMatrix(.rows - 1, 2) = Format(rTabela("NumeroNota"), "000000")
                .TextMatrix(.rows - 1, 3) = Format(rTabela("DataEmissao"), "dd/mm/yy")
                .TextMatrix(.rows - 1, 4) = rTabela("RazaoSocial")
                .TextMatrix(.rows - 1, 5) = Format(rTabela("ValorNota"), ocMONEY)
                .TextMatrix(.rows - 1, 6) = rTabela("Status")
                rTabela.MoveNext
                .rows = .rows + 1
            Loop
        End If

        .rows = .rows - 1
        .Visible = True
        .Redraw = True
    End With
End Sub

'==============================================================
' CARREGAR NOTA SELECIONADA (duplo-clique no GridNotas)
'==============================================================
Private Sub GridNotas_DblClick()
    If GridNotas.Row < 1 Then Exit Sub
    RsOpen TbEntrada, "SELECT * FROM produtos_entrada WHERE CODIGO = " & GridNotas.TextMatrix(GridNotas.Row, 1)
    If TbEntrada.EOF Then Exit Sub
    Load_Controls
    Frm_NF.Tab = 0
    Tab_Entrada.Tab = 0
End Sub

'Chamado de fora (Entrada_Consulta_PorProdutos.frm e similares) pra abrir ja
'filtrado pelas notas que tem esse produto - substitui o antigo txtCodigo.Text = ...
'do form Produtos_Entrada anterior. Chamar ANTES do .Show.
Public Sub FiltrarPorProduto(ByVal lCodProd As Long)
    cboFiltroNota_GotFocus
    cboFiltroNota.Text = "PRODUTO"
    cboFiltroNota_LostFocus
    txtConNotaNumNota.Text = CStr(lCodProd)
    Frm_NF.Tab = 1
    cmdExibirConNotas_Click
End Sub

'Chamado de fora (Entrada_Consulta_PorProdutos.frm e similares) pra abrir ja
'a nota exata, quando a origem do movimento foi FORMA = 'ENTRADA' (nota fiscal).
'Chamar ANTES do .Show.
Public Sub AbrirNotaDireto(ByVal lCodEntrada As Long)
    RsOpen TbEntrada, "SELECT * FROM produtos_entrada WHERE CODIGO = " & lCodEntrada
    If TbEntrada.EOF Then Exit Sub
    Load_Controls
    Frm_NF.Tab = 0
    Tab_Entrada.Tab = 0
End Sub

Public Sub Load_Controls()
    TxtCodFornecedor = Format(TbEntrada("CodigoCorrentista"), "@")
    cboFornecedor = Format(TbEntrada("RazaoSocial"), "@")
    txtNumNota = Format(TbEntrada("NumeroNota"), "@")
    txtCodNota = Format(TbEntrada("CodigoNota"), "@")
    mskEmissao = Format(TbEntrada("DataEmissao"), "@")
    txtTotaldaNota = Format(TbEntrada("ValorNota"), "##,##0.00")
    txtChave = Format(TbEntrada("ChavedeAcesso"), "@")
    txtInformacao = Format(TbEntrada("InformacoesAdicionais"), "@")

    'Transporte
    If TbEntrada("ModFrete") = 0 Then
        cboModFrete.Text = "0 - POR CONTA DO EMITENTE"
    ElseIf TbEntrada("ModFrete") = 1 Then
        cboModFrete.Text = "1 - POR CONTA DE AMBOS"
    ElseIf TbEntrada("ModFrete") = 2 Then
        cboModFrete.Text = "2 - POR CONTA DO TERCEIROS"
    ElseIf TbEntrada("ModFrete") = 9 Then
        cboModFrete.Text = "9 - SEM FRETE"
    End If
    txtCodTransporte = Format(TbEntrada("TranspCodigo"), "@")
    cboTransporte = Format(TbEntrada("TranspNome"), "@")
    txtPlacaUF = Format(TbEntrada("TranspPlacaUF"), "@")
    txtPlaca = Format(TbEntrada("TranspPlaca"), "@")
    txtVolPesoBruto = Format(TbEntrada("VolumePesoBruto"), "@")
    txtVolPesoLiquido = Format(TbEntrada("VolumePesoLiquido"), "@")
    txtVolNumeracao = Format(TbEntrada("VolumeNumeracao"), "@")
    txtVolQuant = Format(TbEntrada("VolumeQuantidade"), "@")
    txtVolEspecie = Format(TbEntrada("VolumeEspecie"), "@")
    txtVolMarca = Format(TbEntrada("VolumeMarca"), "@")

    CarregarTotaisNota TbEntrada

    'Dispara txtCodEntrada_Change -> Mostrar_Itens (carrega GridNotasItens)
    txtCodEntrada = Format(TbEntrada("CODIGO"), "@")
End Sub

Private Sub cmdCopiarChave_Click()
    If Vazio(txtChave.Text) Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText txtChave.Text
    MsgBox "Chave copiada: " & txtChave.Text, vbInformation
End Sub

'==============================================================
' ITENS DA NOTA (GridNotasItens <- produtos_entrada_itens) - mesma formatacao/
' cores/colunas do GridEntradaItens em Entrada_Estoque.Frm
'==============================================================
Private Sub txtCodEntrada_Change()
    Mostrar_Itens
End Sub

Private Sub Mostrar_Itens()
    Dim sSQL_Itens As String
    Dim r As ADODB.Recordset

    Dim sSel As String
    sSel = "SELECT pei.codigo AS varCod, pei.codigo_entrada, pei.CodigoProduto, " & _
           "pei.NomeProduto, ISNULL(pei.EAN,'') AS ean, " & _
           "ISNULL(pei.UnidadeTributavel,'') AS UnidadeTributavel, " & _
           "ISNULL(pei.NCM,'') AS NCM, ISNULL(pei.CFOP,0) AS CFOP, " & _
           "ISNULL(pei.CST,'') AS CST, ISNULL(pei.pICMS,0) AS pICMS, " & _
           "ISNULL(pei.PISCST,'') AS PISCST, ISNULL(pei.PISpPIS,0) AS PISpPIS, " & _
           "ISNULL(pei.COFINSCST,'') AS COFINSCST, " & _
           "ISNULL(pei.COFINSpCOFINS,0) AS COFINSpCOFINS, " & _
           "ISNULL(pei.QuantidadeTributavel,0) AS QuantidadeTributavel, " & _
           "ISNULL((SELECT TOP 1 pp.VALOR_VV FROM Produtos_Precos pp " & _
           "        WHERE pp.COD_PRODUTO = pei.CodigoProduto " & _
           "          AND pp.COD_ENTRADA = pei.codigo_entrada),0) AS VALOR_VV, " & _
           "pei.itemxml, p.COD_BARRA AS var_CodBarra, " & _
           "ISNULL(p.REF,'') AS var_REF, ISNULL(p.tamanho,'') AS tamanho " & _
           "FROM produtos p INNER JOIN produtos_entrada_itens pei " & _
           "ON p.codigo = pei.CodigoProduto"
    If txtCodEntrada.Text = "" Then
        sSQL_Itens = sSel & " WHERE 1=0"
        Set r = dbData.OpenRecordset(sSQL_Itens)
    Else
        sSQL_Itens = sSel & " WHERE pei.codigo_entrada = " & txtCodEntrada.Text & " ORDER BY varCod"
        Set r = dbData.OpenRecordset(sSQL_Itens)
    End If

    FormatarGrid_Itens r
    Dim dTotalGrid As Currency
    dTotalGrid = SomaGridItens(GridNotasItens, 17)
    lblValorNota.Caption = "Qtde Itens: " & (GridNotasItens.rows - 1) & _
                           "   Subtotal: " & FormatNumber(dTotalGrid, 2) & _
                           "   Total: " & FormatNumber(dTotalGrid, 2)

    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
    Dim i As Integer, j As Integer, x As Integer

    With GridNotasItens
        .Clear
        .Cols = 19
        .rows = 2

        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 1400
        .ColWidth(4) = 0
        .ColWidth(5) = 5300
        .ColWidth(6) = 500
        .ColWidth(7) = 850
        .ColWidth(8) = 640
        .ColWidth(9) = 500
        .ColWidth(10) = 530
        .ColWidth(11) = 480
        .ColWidth(12) = 530
        .ColWidth(13) = 750
        .ColWidth(14) = 530
        .ColWidth(15) = 900
        .ColWidth(16) = 700
        .ColWidth(17) = 900
        .ColWidth(18) = 0

        .TextMatrix(0, 1) = "COD"
        .TextMatrix(0, 2) = "COD_ENTRADA"
        .TextMatrix(0, 3) = "COD. BARRA"
        .TextMatrix(0, 4) = "COD."
        .TextMatrix(0, 5) = "DESCRICAO"
        .TextMatrix(0, 6) = "UND"
        .TextMatrix(0, 7) = "NCM"
        .TextMatrix(0, 8) = "CFOP"
        .TextMatrix(0, 9) = "CST"
        .TextMatrix(0, 10) = "ALIQ"
        .TextMatrix(0, 11) = "PIS"
        .TextMatrix(0, 12) = "ALIQ"
        .TextMatrix(0, 13) = "COFINS"
        .TextMatrix(0, 14) = "ALIQ"
        .TextMatrix(0, 15) = "VALOR"
        .TextMatrix(0, 16) = "QTDE"
        .TextMatrix(0, 17) = "TOTAL"
        .TextMatrix(0, 18) = "ITEMXML"

        For x = 0 To .Cols - 1
            .Col = x
            .Row = 0
            .CellFontBold = True
        Next

        For j = 0 To .Cols - 1
            .Row = 0
            .Col = j
            .CellAlignment = flexAlignCenterCenter
        Next

        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 1
        .ColAlignment(9) = 1
        .ColAlignment(10) = 1
        .ColAlignment(11) = 1
        .ColAlignment(12) = 1
        .ColAlignment(13) = 1
        .ColAlignment(14) = 1
        .ColAlignment(15) = 6
        .ColAlignment(16) = 6
        .ColAlignment(17) = 6
        .ColAlignment(18) = 1

        .Redraw = False

        If Not rTabela Is Nothing Then
            Do While Not rTabela.EOF
                .TextMatrix(.rows - 1, 1) = rTabela("varCod")
                .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("ean"))
                .TextMatrix(.rows - 1, 2) = rTabela("codigo_entrada")
                .TextMatrix(.rows - 1, 4) = Format$(rTabela("CodigoProduto"), "0000")
                If tipoEmpresa = 4 Then
                    .TextMatrix(.rows - 1, 5) = rTabela("NomeProduto") & " /  " & rTabela("tamanho") & " / " & rTabela("var_ref")
                Else
                    .TextMatrix(.rows - 1, 5) = rTabela("NomeProduto")
                End If
                .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("UnidadeTributavel"))
                .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("NCM"))
                .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("CFOP"))
                .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("CST"))
                .TextMatrix(.rows - 1, 10) = Format(ValidateNull(rTabela("pICMS")), ocMONEY)
                .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela("PISCST"))
                .TextMatrix(.rows - 1, 12) = Format(ValidateNull(rTabela("PISpPIS")), ocMONEY)
                .TextMatrix(.rows - 1, 13) = ValidateNull(rTabela("COFINSCST"))
                .TextMatrix(.rows - 1, 14) = Format(ValidateNull(rTabela("COFINSpCOFINS")), ocMONEY)
                .TextMatrix(.rows - 1, 15) = Format(ValidateNull(rTabela("VALOR_VV")), "##,##0.00")
                .TextMatrix(.rows - 1, 16) = ValidateNull(rTabela("QuantidadeTributavel"))
                .TextMatrix(.rows - 1, 17) = Format(CDbl(ValidateNull(rTabela("QuantidadeTributavel"))) * CDbl(ValidateNull(rTabela("VALOR_VV"))), "##,##0.00")
                .TextMatrix(.rows - 1, 18) = rTabela("itemxml")

                rTabela.MoveNext
                .rows = .rows + 1
            Loop
        End If

        For i = 1 To .rows - 1
            .Row = i
            .Col = 7:  .CellBackColor = &HE0E0E0
            .Col = 8:  .CellBackColor = &HE0E0E0
            .Col = 9:  .CellBackColor = &HE0E0E0
            .Col = 10: .CellBackColor = &HE0E0E0
            .Col = 11: .CellBackColor = &HE0E0E0
            .Col = 12: .CellBackColor = &HE0E0E0
            .Col = 13: .CellBackColor = &HE0E0E0
            .Col = 14: .CellBackColor = &HE0E0E0
        Next

        .Redraw = True
        .rows = .rows - 1
    End With
End Sub

Private Function SomaGridItens(var_Grid As MSFlexGrid, Col As Integer) As Currency
    Dim i As Integer, Valor As Currency
    Valor = 0
    For i = 0 To var_Grid.rows - 1
        If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
            Valor = Valor + CCur(var_Grid.TextMatrix(i, Col))
        End If
    Next
    SomaGridItens = Valor
End Function

'==============================================================
' TOTAIS DA NOTA (grdTotaisNota) - mesmos campos/helpers do Entrada_Estoque.Frm
'==============================================================
Private Sub CarregarTotaisNota(TbEnt As ADODB.Recordset)
    If TbEnt Is Nothing Then Exit Sub
    Dim f As String
    f = "##,##0.00"
    With grdTotaisNota
        .Redraw = False
        .FixedRows = 0
        .rows = 1
        .Cols = 3
        .FixedCols = 0
        .ColWidth(0) = 1500
        .ColWidth(1) = 3000
        .ColWidth(2) = 1500
        .TextMatrix(0, 0) = "Secao"
        .TextMatrix(0, 1) = "Descricao"
        .TextMatrix(0, 2) = "Valor"
    End With
    Call grdTotHeader(grdTotaisNota, "ICMS")
    Call grdTotAdd(grdTotaisNota, "ICMS", "Base de Calc. ICMS (vBC)", Format(TbEnt("BaseICMS"), f))
    Call grdTotAdd(grdTotaisNota, "ICMS", "Valor do ICMS", Format(TbEnt("ValorICMS"), f))
    Call grdTotAdd(grdTotaisNota, "ICMS", "Valor do FCP", Format(TbEnt("ValorFCP"), f))
    Call grdTotHeader(grdTotaisNota, "ICMS-ST")
    Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Base de Calc. ICMS-ST", Format(TbEnt("BaseICMSST"), f))
    Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Valor do ICMS-ST", Format(TbEnt("ValorICMSST"), f))
    Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Valor do FCP-ST", Format(TbEnt("ValorFCPST"), f))
    Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Valor FCP-ST Retido", Format(TbEnt("ValorFCPSTRet"), f))
    Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Base ICMS-ST Retido (Total)", Format(TbEnt("vBCSTRetTotal"), f))
    Call grdTotAdd(grdTotaisNota, "ICMS-ST", "ICMS-ST Retido (Total)", Format(TbEnt("vICMSSTRetTotal"), f))
    Call grdTotAdd(grdTotaisNota, "ICMS-ST", "ICMS Substituto (Total)", Format(TbEnt("vICMSSubstitutoTotal"), f))
    Call grdTotHeader(grdTotaisNota, "Produtos e Servicos")
    Call grdTotAdd(grdTotaisNota, "Produtos", "Total dos Produtos", Format(TbEnt("ValorProdutos"), f))
    Call grdTotAdd(grdTotaisNota, "Produtos", "Valor do Frete", Format(TbEnt("ValorFrete"), f))
    Call grdTotAdd(grdTotaisNota, "Produtos", "Valor do Seguro", Format(TbEnt("ValorSeguro"), f))
    Call grdTotAdd(grdTotaisNota, "Produtos", "Valor do Desconto", Format(TbEnt("ValorDesconto"), f))
    Call grdTotAdd(grdTotaisNota, "Produtos", "Outras Despesas", Format(TbEnt("ValorOutrasDespesas"), f))
    Call grdTotAdd(grdTotaisNota, "Produtos", "Valor do II (Importacao)", Format(TbEnt("ValorImportacao"), f))
    Call grdTotHeader(grdTotaisNota, "IPI / PIS / COFINS")
    Call grdTotAdd(grdTotaisNota, "IPI", "Valor do IPI", Format(TbEnt("ValorIPI"), f))
    Call grdTotAdd(grdTotaisNota, "IPI", "Valor IPI Devolvido", Format(TbEnt("ValorIPIDevol"), f))
    Call grdTotAdd(grdTotaisNota, "PIS", "Valor do PIS", Format(TbEnt("ValorPIS"), f))
    Call grdTotAdd(grdTotaisNota, "COFINS", "Valor do COFINS", Format(TbEnt("ValorCOFINS"), f))
    Call grdTotHeader(grdTotaisNota, "DIFAL")
    Call grdTotAdd(grdTotaisNota, "DIFAL", "FCP UF Destinataria", Format(TbEnt("vFCPUFDest"), f))
    Call grdTotAdd(grdTotaisNota, "DIFAL", "ICMS UF Destinataria", Format(TbEnt("vICMSUFDest"), f))
    Call grdTotAdd(grdTotaisNota, "DIFAL", "ICMS UF Remetente", Format(TbEnt("vICMSUFRemet"), f))
    Call grdTotTotal(grdTotaisNota, "VALOR TOTAL DA NOTA", Format(TbEnt("ValorNota"), f))
    Call grdTotHeader(grdTotaisNota, "IBS / CBS (Reforma Tributaria)")
    Call grdTotAdd(grdTotaisNota, "IBS/CBS", "Base de Calc. IBS/CBS (vBC)", Format(TbEnt("TotIBSCBSvBC"), f))
    Call grdTotAdd(grdTotaisNota, "IBS", "IBS Estado (UF)", Format(TbEnt("TotIBSUFvIBS"), f))
    Call grdTotAdd(grdTotaisNota, "IBS", "IBS Municipio", Format(TbEnt("TotIBSMunvIBS"), f))
    Call grdTotAdd(grdTotaisNota, "IBS", "Total IBS", Format(TbEnt("TotIBSvIBS"), f))
    Call grdTotAdd(grdTotaisNota, "IBS", "Credito Presumido IBS", Format(TbEnt("TotIBSvCredPres"), f))
    Call grdTotAdd(grdTotaisNota, "CBS", "Total CBS", Format(TbEnt("TotCBSvCBS"), f))
    Call grdTotAdd(grdTotaisNota, "CBS", "Credito Presumido CBS", Format(TbEnt("TotCBSvCredPres"), f))
    Call grdTotAdd(grdTotaisNota, "IBS/CBS", "Valor NF com IBS/CBS (vNFTot)", Format(TbEnt("vNFTot"), f))
    grdTotaisNota.FixedRows = 1
    grdTotaisNota.Redraw = True
End Sub

Private Sub grdTotHeader(g As MSFlexGrid, sTitle As String)
    g.rows = g.rows + 1
    Dim r As Long
    r = g.rows - 1
    g.TextMatrix(r, 0) = ""
    g.TextMatrix(r, 1) = sTitle
    g.TextMatrix(r, 2) = ""
    g.Row = r: g.Col = 0: g.CellBackColor = &HC0C0C0
    g.Row = r: g.Col = 1: g.CellBackColor = &HC0C0C0: g.CellFontBold = True
    g.Row = r: g.Col = 2: g.CellBackColor = &HC0C0C0
End Sub

Private Sub grdTotAdd(g As MSFlexGrid, sSec As String, sDesc As String, sVal As String)
    g.rows = g.rows + 1
    Dim r As Long
    r = g.rows - 1
    g.TextMatrix(r, 0) = sSec
    g.TextMatrix(r, 1) = sDesc
    g.TextMatrix(r, 2) = sVal
    g.Row = r: g.Col = 2: g.CellAlignment = flexAlignRightCenter
End Sub

Private Sub grdTotTotal(g As MSFlexGrid, sDesc As String, sVal As String)
    g.rows = g.rows + 1
    Dim r As Long
    r = g.rows - 1
    g.TextMatrix(r, 0) = "TOTAL"
    g.TextMatrix(r, 1) = sDesc
    g.TextMatrix(r, 2) = sVal
    Dim c As Integer
    For c = 0 To 2
        g.Row = r: g.Col = c
        g.CellFontBold = True
        g.CellBackColor = &HFFFFC0
    Next c
    g.Row = r: g.Col = 2: g.CellAlignment = flexAlignRightCenter
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    End If
End Sub
