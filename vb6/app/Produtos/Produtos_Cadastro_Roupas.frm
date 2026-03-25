VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Produtos_Cadastro_Roupas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRODUTOS"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13920
   Icon            =   "Produtos_Cadastro_Roupas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   13920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   13725
      TabIndex        =   26
      Top             =   60
      Width           =   13755
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Produtos_Cadastro_Roupas.frx":23D2
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUTOS"
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
         Left            =   1365
         TabIndex        =   27
         Top             =   240
         Width           =   1770
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8715
      Left            =   60
      TabIndex        =   19
      Top             =   1080
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   15372
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabMaxWidth     =   3175
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
      TabPicture(0)   =   "Produtos_Cadastro_Roupas.frx":7DA5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdSair"
      Tab(0).Control(1)=   "cmdNovo"
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(3)=   "cmdSalvar"
      Tab(0).Control(4)=   "cmdExcluir"
      Tab(0).Control(5)=   "cmdAlterar"
      Tab(0).Control(6)=   "frmCadastro"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Produtos_Cadastro_Roupas.frx":7DC1
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label25"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdImprimir"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdExibir"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Grid"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "frmCompra"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "frmEstoque"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "frmVenda"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame9"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame8"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "HISTÓRICO"
      TabPicture(2)   =   "Produtos_Cadastro_Roupas.frx":7DDD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid_Estoque"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame8 
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
         ForeColor       =   &H000000C0&
         Height          =   1635
         Left            =   3180
         TabIndex        =   91
         Top             =   6960
         Width           =   1515
         Begin VB.CheckBox ckkORDTam 
            Caption         =   "Tamanho"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   924
            Width           =   975
         End
         Begin VB.CheckBox ckkORDFab 
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   468
            Width           =   1335
         End
         Begin VB.CheckBox ckkORDRef 
            Caption         =   "Referência"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   696
            Width           =   1215
         End
         Begin VB.CheckBox ckkORDLinha 
            Caption         =   "Linha"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   1152
            Width           =   1215
         End
         Begin VB.CheckBox ckkORDDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox ckkORDQuant 
            Caption         =   "Quant."
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   1380
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filtro"
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
         Height          =   1635
         Left            =   4740
         TabIndex        =   78
         Top             =   6960
         Width           =   7335
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   900
            TabIndex        =   84
            Top             =   240
            Width           =   3015
         End
         Begin VB.ComboBox cboConsFab 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4980
            TabIndex        =   83
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox cboConsRef 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   82
            Top             =   660
            Width           =   1875
         End
         Begin VB.ComboBox cboConsTam 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3780
            TabIndex        =   81
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cboConsLinha 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5100
            TabIndex        =   80
            Top             =   660
            Width           =   2175
         End
         Begin VB.TextBox txtConsCodBarra 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   79
            Top             =   1080
            Width           =   2355
         End
         Begin VB.Label lblConsProduto 
            Caption         =   "Descrição:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   90
            Top             =   300
            Width           =   795
         End
         Begin VB.Label lblConsFab 
            Caption         =   "Fabricante:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4140
            TabIndex        =   89
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblConsRef 
            Caption         =   "Referência:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblConsTam 
            Caption         =   "Tamanho:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3000
            TabIndex        =   87
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lblConsLinha 
            Caption         =   "Linha:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4620
            TabIndex        =   86
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblConsCodBarra 
            Caption         =   "Cod. Barra:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   85
            Top             =   1140
            Width           =   855
         End
      End
      Begin VB.Frame Frame9 
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
         ForeColor       =   &H000000C0&
         Height          =   1635
         Left            =   120
         TabIndex        =   70
         Top             =   6960
         Width           =   3015
         Begin VB.CheckBox chkTam 
            Caption         =   "Tamanho"
            Height          =   195
            Left            =   1620
            TabIndex        =   77
            Top             =   300
            Width           =   975
         End
         Begin VB.CheckBox chkFab 
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CheckBox chkRef 
            Caption         =   "Referência"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   1260
            Width           =   1215
         End
         Begin VB.CheckBox chkTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   300
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.CheckBox chkCodBarra 
            Caption         =   "Cód. de Barra"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   780
            Width           =   1455
         End
         Begin VB.CheckBox chkProduto 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   540
            Width           =   1215
         End
         Begin VB.CheckBox chkLinha 
            Caption         =   "Linha"
            Height          =   195
            Left            =   1620
            TabIndex        =   71
            Top             =   540
            Width           =   975
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Estoque 
         Height          =   8055
         Left            =   -74880
         TabIndex        =   61
         Top             =   420
         Width           =   13515
         _ExtentX        =   23839
         _ExtentY        =   14208
         _Version        =   393216
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtro"
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
         Height          =   1635
         Left            =   0
         TabIndex        =   58
         Top             =   5280
         Width           =   5715
         Begin VB.ComboBox cboConsProduto2 
            Height          =   315
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Visible         =   0   'False
            Width           =   5475
         End
         Begin VB.Label lblNomeCombo 
            Caption         =   "Nome"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   480
            Visible         =   0   'False
            Width           =   1515
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   120
         TabIndex        =   57
         ToolTipText     =   "muda o cod_entrada dos nulos para 1"
         Top             =   5160
         Width           =   135
      End
      Begin VB.Frame frmVenda 
         Caption         =   "VENDA"
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
         Height          =   1575
         Left            =   11040
         TabIndex        =   48
         Top             =   5340
         Width           =   2595
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            Caption         =   "Imposto:"
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
            TabIndex        =   54
            Top             =   600
            Width           =   795
         End
         Begin VB.Label lblImpVenda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   53
            Top             =   600
            Width           =   1545
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            Caption         =   "Lucro:"
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
            Top             =   300
            Width           =   795
         End
         Begin VB.Label lblLucro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   51
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label lblVenda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1020
            TabIndex        =   50
            Top             =   1020
            Width           =   1545
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Venda:"
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
            TabIndex        =   49
            Top             =   900
            Width           =   615
         End
      End
      Begin VB.Frame frmEstoque 
         Caption         =   "ESTOQUE"
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
         Height          =   1575
         Left            =   5760
         TabIndex        =   34
         Top             =   5340
         Width           =   2595
         Begin VB.Label lblEstoque 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   38
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            Caption         =   "Itens:"
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
            TabIndex        =   37
            Top             =   540
            Width           =   795
         End
         Begin VB.Label lblProdutos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   36
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipos:"
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
            Left            =   405
            TabIndex        =   35
            Top             =   300
            Width           =   540
         End
      End
      Begin VB.Frame frmCompra 
         Caption         =   "COMPRA"
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
         Height          =   1575
         Left            =   8400
         TabIndex        =   29
         Top             =   5340
         Width           =   2595
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Custo:"
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
            Left            =   360
            TabIndex        =   47
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label lblCustoFinal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   960
            TabIndex        =   46
            Top             =   1200
            Width           =   1545
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            Caption         =   "Imposto:"
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
            TabIndex        =   45
            Top             =   900
            Width           =   795
         End
         Begin VB.Label lblImpCompra 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   44
            Top             =   900
            Width           =   1545
         End
         Begin VB.Label lblCusto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   33
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            Caption         =   "Compra:"
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
            TabIndex        =   32
            Top             =   300
            Width           =   795
         End
         Begin VB.Label lblFrete 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   31
            Top             =   600
            Width           =   1545
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            Caption         =   "Frete:"
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
            TabIndex        =   30
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.PictureBox frmCadastro 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   -74880
         ScaleHeight     =   1755
         ScaleWidth      =   13455
         TabIndex        =   20
         Top             =   480
         Width           =   13515
         Begin VB.ComboBox cboCategoria 
            Height          =   315
            Left            =   11760
            TabIndex        =   7
            Top             =   300
            Width           =   1635
         End
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   2160
            MaxLength       =   90
            TabIndex        =   2
            Top             =   300
            Width           =   3915
         End
         Begin VB.ComboBox cboUnidMedida 
            Height          =   315
            Left            =   10860
            TabIndex        =   6
            Top             =   300
            Width           =   855
         End
         Begin VB.ComboBox cboFabricante 
            Height          =   315
            Left            =   7680
            TabIndex        =   4
            Top             =   300
            Width           =   1635
         End
         Begin VB.TextBox txtTam 
            Height          =   315
            Left            =   9360
            MaxLength       =   20
            TabIndex        =   5
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtRef 
            Height          =   315
            Left            =   6120
            TabIndex        =   3
            Top             =   300
            Width           =   1515
         End
         Begin VB.TextBox txtValorAtual 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3720
            TabIndex        =   11
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   60
            MaxLength       =   90
            TabIndex        =   1
            Top             =   300
            Width           =   2055
         End
         Begin VB.CheckBox chkDestaque 
            Caption         =   "Destaque"
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
            Left            =   1080
            TabIndex        =   42
            Top             =   1380
            Width           =   1635
         End
         Begin VB.CheckBox chkAtivo 
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
            Left            =   60
            TabIndex        =   41
            Top             =   1380
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.TextBox txtObs 
            Height          =   315
            Left            =   6660
            MaxLength       =   90
            TabIndex        =   13
            Top             =   960
            Width           =   6735
         End
         Begin VB.TextBox txtCodigo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   12900
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   -60
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtQuant 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   10
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txtQuantMin 
            Height          =   315
            Left            =   1080
            TabIndex        =   9
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtUltCompra 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtPrateleira 
            Height          =   315
            Left            =   60
            MaxLength       =   4
            TabIndex        =   8
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
            Height          =   195
            Left            =   11760
            TabIndex        =   69
            Top             =   60
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            Height          =   195
            Left            =   2160
            TabIndex        =   68
            Top             =   60
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unid. Med."
            Height          =   195
            Left            =   10860
            TabIndex        =   67
            Top             =   60
            Width           =   780
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   7680
            TabIndex        =   66
            Top             =   60
            Width           =   750
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tam."
            Height          =   195
            Left            =   9360
            TabIndex        =   65
            Top             =   60
            Width           =   360
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            Height          =   195
            Left            =   6120
            TabIndex        =   64
            Top             =   60
            Width           =   300
         End
         Begin VB.Label lblValorAtual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Atual"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3720
            TabIndex        =   62
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. Barra"
            Height          =   195
            Left            =   60
            TabIndex        =   43
            Top             =   60
            Width           =   750
         End
         Begin VB.Label Observação 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observação"
            Height          =   195
            Left            =   6660
            TabIndex        =   28
            Top             =   720
            Width           =   870
         End
         Begin VB.Label lblQuantAtual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Atual"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2040
            TabIndex        =   25
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última Compra"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   24
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Min."
            Height          =   195
            Left            =   1080
            TabIndex        =   23
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local"
            Height          =   195
            Left            =   60
            TabIndex        =   22
            Top             =   720
            Width           =   390
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4635
         Left            =   120
         TabIndex        =   39
         Top             =   540
         Width           =   13515
         _ExtentX        =   23839
         _ExtentY        =   8176
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   555
         Left            =   -73140
         TabIndex        =   16
         Top             =   2340
         Width           =   1695
         _ExtentX        =   2990
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
         MICON           =   "Produtos_Cadastro_Roupas.frx":7DF9
         PICN            =   "Produtos_Cadastro_Roupas.frx":7E15
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
         Left            =   -71400
         TabIndex        =   17
         Top             =   2340
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Produtos_Cadastro_Roupas.frx":86EF
         PICN            =   "Produtos_Cadastro_Roupas.frx":870B
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
         Left            =   -73140
         TabIndex        =   14
         Top             =   2340
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
         MICON           =   "Produtos_Cadastro_Roupas.frx":8A25
         PICN            =   "Produtos_Cadastro_Roupas.frx":8A41
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
         Left            =   -71400
         TabIndex        =   15
         Top             =   2340
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
         MICON           =   "Produtos_Cadastro_Roupas.frx":F30B
         PICN            =   "Produtos_Cadastro_Roupas.frx":F327
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   555
         Left            =   -74880
         TabIndex        =   0
         Top             =   2340
         Width           =   1695
         _ExtentX        =   2990
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
         MICON           =   "Produtos_Cadastro_Roupas.frx":15DCB
         PICN            =   "Produtos_Cadastro_Roupas.frx":15DE7
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
         Height          =   555
         Left            =   -63060
         TabIndex        =   18
         Top             =   2340
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Produtos_Cadastro_Roupas.frx":16AC1
         PICN            =   "Produtos_Cadastro_Roupas.frx":16ADD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   735
         Left            =   12120
         TabIndex        =   55
         Top             =   7060
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1296
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
         MICON           =   "Produtos_Cadastro_Roupas.frx":16DF7
         PICN            =   "Produtos_Cadastro_Roupas.frx":16E13
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
         Height          =   735
         Left            =   12120
         TabIndex        =   56
         Top             =   7860
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1296
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
         MICON           =   "Produtos_Cadastro_Roupas.frx":176ED
         PICN            =   "Produtos_Cadastro_Roupas.frx":17709
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Dê um duplo-clique para ver mais informações"
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   300
         TabIndex        =   40
         Top             =   5160
         Width           =   3435
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   63
      Top             =   9870
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20214
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "09:57"
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
Attribute VB_Name = "Produtos_Cadastro_Roupas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

Dim var_cod_Preco As Long

Private Sub Entrada_Estoque()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'ENTRADA DO PRODUTO
   If cmdNovo.Visible = True Then
      Dim var_COD_Itens As Long
      
      'AUTONUMERAÇÃO
      sSQL = "SELECT IFNULL(MAX(codigo), 0) AS cod_itens FROM produtos_entrada_itens;"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then var_COD_Itens = r("cod_itens") + 1
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
      sSQL = "INSERT INTO produtos_entrada_itens (codigo, codigo_entrada, codigo_produto, descricao) VALUES (" & _
         var_COD_Itens & ", 0, " & txtCodigo.Text & ", '" & txtDescricao.Text & "');"
      dbData.Execute sSQL
   End If
   
   'COLOCAR O PREÇO
   'If cmdAlterar.Visible = True Then
   '   sSQL = "UPDATE produtos_entrada_itens SET VENDA = '" & txtValorAtual.Text & "' WHERE (codigo_produto = " & txtCodigo.Text & ") ORDER BY codigo DESC LIMIT 0, 1;"
   '   Set r = dbData.OpenRecordset(sSQL)
   'End If
   
   
End Sub

Private Sub FormatarGrid_Historico(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   Dim X As Integer
   
   With Grid_Estoque
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 500
      .ColWidth(2) = 1500
      .ColWidth(3) = 1500
      .ColWidth(4) = 7000
      .ColWidth(5) = 1500
      .ColWidth(6) = 1500
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "No. FISCAL"
      .TextMatrix(0, 4) = "FORNECEDOR"
      .TextMatrix(0, 5) = "QUANT"
      .TextMatrix(0, 6) = "COMPRA"
      
      'colocar os cabeçalho em negrito
      For X = 0 To .Cols - 1
         .Col = X
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For j = 0 To .Cols - 1
         .Row = 0
         .Col = j
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'mudar a cor da coluna
            'For i = 1 To .Rows - 1
            '   .Row = i
            '   .Col = 6
            '   .CellBackColor = &HC0FFFF
            'Next
            
            'ALINHAMENTO
            '.ColAlignment(2) = 1
            
            .TextMatrix(.Rows - 1, 1) = rTabela("var_codigo")
            .TextMatrix(.Rows - 1, 2) = Format$(rTabela("data_entrada"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = rTabela("notafiscal")
            .TextMatrix(.Rows - 1, 4) = rTabela("fornecedor")
            .TextMatrix(.Rows - 1, 5) = rTabela("quant")
            .TextMatrix(.Rows - 1, 6) = Format$(rTabela("custo"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Redraw = True
      .Rows = .Rows - 1
   End With
End Sub

Private Sub FormatarGrid_Produtos(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   Dim X As Integer
   
   With Grid
      .Clear
      .Cols = 19
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1600 '1600
      .ColWidth(3) = 4445 '4445
      .ColWidth(4) = 800
      .ColWidth(5) = 850
      .ColWidth(6) = 850
      .ColWidth(7) = 1000
      .ColWidth(8) = 850
      .ColWidth(9) = 850
      .ColWidth(10) = 1000
      .ColWidth(11) = 900
      .ColWidth(12) = 0
      .ColWidth(13) = 0
      .ColWidth(14) = 0
      .ColWidth(15) = 0
      .ColWidth(16) = 0
      .ColWidth(17) = 0
      .ColWidth(18) = 0
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "CÓD. BARRA"
      .TextMatrix(0, 3) = "PRODUTO"
      .TextMatrix(0, 4) = "QUANT"
      .TextMatrix(0, 5) = "CUSTO"
      .TextMatrix(0, 6) = "FRETE"
      .TextMatrix(0, 7) = "IMPOSTO"
      .TextMatrix(0, 8) = "VALOR"
      .TextMatrix(0, 9) = "LUCRO"
      .TextMatrix(0, 10) = "IMPOSTO"
      .TextMatrix(0, 11) = "VENDA"
      .TextMatrix(0, 12) = "T_venda"
      .TextMatrix(0, 13) = "T_IVenda"
      .TextMatrix(0, 14) = "T_Lucro"
      .TextMatrix(0, 15) = "T_Custo"
      .TextMatrix(0, 16) = "DIF"  '"T_Frete"
      .TextMatrix(0, 17) = "T_Compra"
      .TextMatrix(0, 18) = "T_Compra"
      
      'colocar os cabeçalho em negrito
      For X = 0 To .Cols - 1
         .Col = X
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For j = 0 To .Cols - 1
         .Row = 0
         .Col = j
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'mudar a cor da coluna
            'For i = 1 To .Rows - 1
            '   .Row = i
            '   .Col = 6
            '   .CellBackColor = &HC0FFFF
            'Next
            
            'ALINHAMENTO
            '.ColAlignment(2) = 1
    
            .TextMatrix(.Rows - 1, 1) = rTabela("var_codent")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_codbarra")
            .TextMatrix(.Rows - 1, 3) = rTabela("var_desc")
            .TextMatrix(.Rows - 1, 4) = Format$(rTabela("var_quant"), ocPESO)
            .TextMatrix(.Rows - 1, 5) = Format$(rTabela("var_custo"), ocMONEY)
            .TextMatrix(.Rows - 1, 6) = Format$(rTabela("var_frete"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = Format$(rTabela("var_impcompra"), ocMONEY)
            .TextMatrix(.Rows - 1, 8) = Format$(rTabela("var_vlrcompra"), ocMONEY)
            .TextMatrix(.Rows - 1, 9) = Format$(rTabela("var_lucro"), ocMONEY)
            .TextMatrix(.Rows - 1, 10) = Format$(rTabela("var_impvenda"), ocMONEY)
            .TextMatrix(.Rows - 1, 11) = Format$(rTabela("venda"), ocMONEY)
            .TextMatrix(.Rows - 1, 12) = Format$(rTabela("var_totalvenda"), ocMONEY)
            .TextMatrix(.Rows - 1, 13) = Format$(rTabela("var_totalimpvenda"), ocMONEY)
            .TextMatrix(.Rows - 1, 14) = Format$(rTabela("var_totallucro"), ocMONEY)
            .TextMatrix(.Rows - 1, 15) = Format$(rTabela("var_totalcusto"), ocMONEY)
            '.TextMatrix(.Rows - 1, 16) = Format$(Rtabela("var_diferenca"), ocMoney)
            .TextMatrix(.Rows - 1, 17) = Format(rTabela("var_totalicompra"), ocMONEY)
            .TextMatrix(.Rows - 1, 18) = Format(rTabela("var_totalcompra"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Redraw = True
      .Rows = .Rows - 1
   End With
   
   lblVenda.Caption = Format(SomaGrid(Grid, 12), ocMONEY)
   lblImpVenda.Caption = Format(SomaGrid(Grid, 13), ocMONEY)
   lblLucro.Caption = Format(SomaGrid(Grid, 14), ocMONEY)
   
   lblCusto.Caption = Format(SomaGrid(Grid, 15), ocMONEY)
   lblFrete.Caption = Format(SomaGrid(Grid, 16), ocMONEY)
   lblImpCompra.Caption = Format(SomaGrid(Grid, 17), ocMONEY)
   lblCustoFinal.Caption = Format(SomaGrid(Grid, 18), ocMONEY)
   lblEstoque.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
   lblProdutos.Caption = Grid.Rows - 1  'contar o numeros de linhas no grid
End Sub

Private Sub LimparGrid_Produtos()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT produtos.codigo AS var_codent, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, " & _
      "produtos_entrada_itens.custo AS var_custo, produtos_entrada_itens.frete AS var_frete, produtos_entrada_itens.imposto_valor_compra AS var_impcompra, " & _
      "produtos_entrada_itens.custo_compra AS var_vlrcompra, produtos_entrada_itens.lucro_valor AS var_lucro, produtos_entrada_itens.imposto_valor_venda AS var_impvenda, " & _
      "produtos_entrada_itens.venda AS var_vlrvenda, produtos.codigo, produtos_entrada_itens.codigo_produto " & _
      "FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.codigo_produto " & _
      "WHERE false;"
    
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid_Produtos r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

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

Private Sub MostrarDados_Produto(rTabela As ADODB.Recordset)
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim vrVenda As Currency
   
   'mostrar o ultimo preço de compra
   sSQL = "SELECT venda FROM produtos_entrada_itens WHERE (codigo_produto = " & rTabela("codigo") & ") ORDER BY codigo DESC LIMIT 0, 1;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.EOF Then vrVenda = r("venda")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   txtCodigo.Text = ValidateNull(rTabela("codigo"))
   txtCodBarra.Text = ValidateNull(rTabela("cod_barra"))
   txtDescricao.Text = ValidateNull(rTabela("descricao"))
   cboFabricante.Text = ValidateNull(rTabela("fabricante"))
   cboUnidMedida.Text = ValidateNull(rTabela("unid_medida"))
   cboCategoria.Text = ValidateNull(rTabela("categoria"))
   txtPrateleira.Text = ValidateNull(rTabela("prateleira"))
   txtQuant.Text = ValidateNull(rTabela("quant_estoque"))
   txtValorAtual.Text = Format(vrVenda, ocMONEY)
   txtQuantMin.Text = ValidateNull(rTabela("quant_min"))
   txtUltCompra.Text = Format$(rTabela("ult_compra"), "dd/mm/yy")
   txtObs.Text = ValidateNull(rTabela("observacao"))
   
   chkAtivo.Value = Abs(CBool(rTabela("ativo")))
   chkDestaque.Value = Abs(CBool(rTabela("destaque")))
End Sub

Private Sub Autonumeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT IFNULL(MAX(codigo), 0) AS cod_produto FROM produtos;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigo.Text = r("cod_produto") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub LimparObjetos_Produtos()
   If cmdAlterar.Visible = False Then txtCodigo.Text = ""
   txtCodBarra.Text = ""
   txtDescricao.Text = ""
   cboFabricante.Text = ""
   cboCategoria.Text = ""
   cboUnidMedida.Text = ""
   txtPrateleira.Text = ""
   txtQuant.Text = ""
   txtQuantMin.Text = ""
   txtUltCompra.Text = ""
   txtObs.Text = ""
   txtValorAtual.Text = ""
   chkAtivo.Value = Unchecked
   chkDestaque.Value = Unchecked
   
   cmdNovo.Enabled = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   'lblQuantAtual.Enabled = False
   'lblValorAtual.Enabled = False
   'txtQuant.Enabled = False
   'txtValorAtual.Enabled = False
End Sub

Private Sub cboCategoria_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Limpa a lista atual
   cboCategoria.Clear
   
   sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCategoria.AddItem ValidateNull(r("categoria"))
      r.MoveNext
   Loop
   
   moCombo.AttachTo cboCategoria
End Sub

Private Sub cboCategoria_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboConsProduto_Change()
   'If optCodBarra.Value = True And Len(cboConsProduto) = 13 Then
   '   cmdExibir_Click
   'End If
End Sub

Private Sub cboConsProduto_GotFocus()
 
   

End Sub

Private Sub cboConsProduto2_Change()
   If optCodBarra.Value = True Then
      cboConsProduto.Clear
   
   ElseIf optProduto.Value = True Then
     
      
   ElseIf optCategoria.Value = True Then
      cboConsProduto.Clear
      
      sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"
      Set r = dbData.OpenRecordset(sSQL)
      
      Do While Not r.EOF
         cboConsProduto.AddItem ValidateNull(r("categoria"))
         r.MoveNext
      Loop
   End If
   
   SelectControl cboConsProduto
   moCombo.AttachTo cboConsProduto
End Sub


Private Sub cboFabricante_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Limpa a lista
   cboFabricante.Clear
   
   sSQL = "SELECT DISTINCT fabricante FROM produtos ORDER BY fabricante;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboFabricante.AddItem ValidateNull(r("fabricante"))
      r.MoveNext
   Loop
   
   moCombo.AttachTo cboFabricante
End Sub

Private Sub cboFabricante_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboUnidMedida_GotFocus()
Dim var_Texto As String
var_Texto = cboUnidMedida.Text

   cboUnidMedida.Clear
   cboUnidMedida.AddItem "UNID"
   cboUnidMedida.AddItem "CX"
   cboUnidMedida.AddItem "M"
   cboUnidMedida.AddItem "M²"
   cboUnidMedida.AddItem "M³"
   cboUnidMedida.AddItem "ML"
   cboUnidMedida.AddItem "KG"
   cboUnidMedida.AddItem "G"
   moCombo.AttachTo cboUnidMedida
   
cboUnidMedida.Text = var_Texto
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

Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o produto na guia CONSULTA.", vbInformation
      Exit Sub
   End If
   
   'Não é necessário consulta o registro antes de atualiza-lo
   'sSQL = "SELECT * FROM produtos WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   'alterar o nome dos produtos da tabela de entrada de pedidos
   dbData.Execute "UPDATE produtos_entrada_itens SET descricao = '" & txtDescricao.Text & "' WHERE (codigo_produto = " & txtCodigo.Text & ");"
   'dbData.Execute "UPDATE pedidos_itens SET descricao = '" & txtDescricao.Text & "' WHERE (cod_produto = " & txtCodigo.Text & ");"
   
   'alterar o nome dos produtos da tabela de entrada de pedidos
   sSQL = "UPDATE produtos_entrada_itens SET VENDA = " & Replace(CCur(txtValorAtual.Text), ",", ".") & " WHERE (codigo = " & _
      "(SELECT codigo FROM (SELECT codigo FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodigo.Text & ") ORDER BY codigo DESC LIMIT 0, 1) as tempTabela));"

   dbData.Execute sSQL
   
   
   'alterar o valor da ultima entrada no estoque
   'sSQL = "SELECT * FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodigo.Text & ") ORDER BY codigo;"
   'Set r = dbData.OpenRecordset(sSQL)
   
   ' If Not RS.EOF Then
   ' RS.MoveLast
   ' RS.Edit
   ' RS.Fields!VENDA = IIf(txtValorAtual.Text = "", Null, txtValorAtual.Text)
   ' RS.Update
   ' End If
    
   cmdNovo.Enabled = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   'CmdHabilitar.Visible = False
   'lblQuantAtual.Enabled = False
   'lblValorAtual.Enabled = False
   'txtQuant.Enabled = False
   'txtValorAtual.Enabled = False
   frmCadastro.Enabled = False
   LimparGrid_Produtos
   Mostrar_Historico
End Sub

Private Function Inserir_Dados() As Boolean
   'A inclusão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   
   'Valida os campos
   If Trim(txtQuant.Text) = "" Then txtQuant.Text = 0
   If Trim(txtQuantMin.Text) = "" Then txtQuantMin.Text = 0
   
   'Comando de inclusão
   sSQL = "INSERT INTO produtos (" & _
      "codigo, ativo, destaque, cod_barra, descricao, fabricante, unid_medida, " & _
      "categoria, prateleira, quant_min, observacao, quant_estoque, ref, tamanho) VALUES (" & _
      txtCodigo.Text & ", " & Abs(chkAtivo.Value) & ", " & Abs(chkDestaque.Value) & ", '" & _
      IIf((txtCodBarra.Text = ""), txtCodigo.Text, txtCodBarra.Text) & "', '" & _
      txtDescricao.Text & "', '" & cboFabricante.Text & "', '" & cboUnidMedida.Text & "', '" & _
      cboCategoria.Text & "', '" & txtPrateleira.Text & "', " & Replace(CDbl(txtQuantMin.Text), ",", ".") & ", '" & _
      txtObs.Text & "', " & Replace(CDbl(txtQuant.Text), ",", ".") & ", '" & txtRef.Text & "', '" & txtTam.Text & "');"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE produtos SET " & _
      "ativo = " & Abs(chkAtivo.Value) & ", " & _
      "destaque = " & Abs(chkDestaque.Value) & ", " & _
      "cod_barra = '" & IIf((txtCodBarra.Text = ""), txtCodigo.Text, txtCodBarra.Text) & "', " & _
      "descricao = '" & txtDescricao.Text & "', " & _
      "fabricante = '" & cboFabricante.Text & "', " & _
      "unid_medida = '" & cboUnidMedida.Text & "', " & _
      "categoria = '" & cboCategoria.Text & "', " & _
      "tamanho = '" & txtTam.Text & "', " & _
      "ref = '" & txtRef.Text & "', " & _
      "prateleira = '" & txtPrateleira.Text & "', " & _
      "quant_min = " & Replace(CDbl(txtQuantMin.Text), ",", ".") & ", " & _
      "observacao = '" & txtObs.Text & "', " & _
      "quant_estoque = " & Replace(CDbl(txtQuant.Text), ",", ".")
   
   'Condição para atualização
   sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdCancelar_Click()
   LimparObjetos_Produtos
   frmCadastro.Enabled = False
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
   
   If txtCodigo.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o produto na guia CONSULTA", vbInformation
      Exit Sub
   End If
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir esse produto?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM produtos WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
      
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   LimparObjetos_Produtos
   cmdNovo.Enabled = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   'lblQuantAtual.Enabled = False
   'lblValorAtual.Enabled = False
   'txtQuant.Enabled = False
   'txtValorAtual.Enabled = False
   frmCadastro.Enabled = False
   LimparGrid_Produtos
   Mostrar_Historico
End Sub

Private Sub cmdImprimir_Click()
   'colocar o nome da maquina na barra de status
   Dim r As ADODB.Recordset
   Dim var_Impressora As String
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   
   Set oIni = Nothing
   
   Me.Hide
   
   Set r = dbData.OpenRecordset(printSQL)
   
   Set REL_Produtos.Relatorio.Recordset = r
   REL_Produtos.dfQuant.Caption = "Quant.: " & lblQuantAtual.Caption
   REL_Produtos.dfBruto.Caption = "Bruto: " & lblValorAtual.Caption
   
   'If optMensal.Value = True Then
   '   REL_Produtos.dfTipo.Caption = "Tipo: Mês = " & cboMes.Text & "/" & cboAno.Text
   'ElseIf optProduto.Value = True Then
   '   REL_Produtos.dfTipo.Caption = "Tipo: Produto = " & cboProduto.Text & ""
   'ElseIf optFornecedor.Value = True Then
   '   REL_Produtos.dfTipo.Caption = "Tipo: Fornecedor = " & cboConsFornecedor.Text & ""
   'ElseIf optNotaFiscal.Value = True Then
   '   REL_Produtos.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & txtConsNotaFiscal.Text & ""
   'Else
   '   REL_Produtos.dfTipo.Caption = "Tipo: Todas as notas"
   'End If
   
   REL_Produtos.Relatorio.NomeImpressora = var_Impressora
   REL_Produtos.Relatorio.Ativar
   Unload REL_Produtos
   
   Me.Show 1
End Sub

Private Sub cmdNovo_Click()
   frmCadastro.Enabled = True
   LimparObjetos_Produtos
   cmdNovo.Enabled = False
   cmdSalvar.Visible = True
   cmdCancelar.Visible = True
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   chkAtivo.Value = Checked
   Autonumeracao
   'lblQuantAtual.Enabled = True
   'lblValorAtual.Enabled = True
   'txtQuant.Enabled = True
   'txtValorAtual.Enabled = True
   cboUnidMedida.Text = "UNID"
   txtQuant.Text = "0"
   txtCodBarra.SetFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
   'Não foi informado a descricao do produto.
   If txtDescricao.Text = "" Then
      ShowMsg "Digite a Descrição do produto", vbInformation
      txtDescricao.SetFocus
      Exit Sub
   End If
   
   'Não é necessário consultar todos os registros antes de inserir um novo
   'sSQL = "SELECT * FROM produtos"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'A auto numeração do código deve ser utilizada no momento de salvar o registro
   'para evitar duplicidade de código para quando houver mais de um terminal operando
   'ao mesmo tempo
   'AutoNumeracao
   
   'Faz a inserção de forma direta e verifica se houve algum erro
   If Not Inserir_Dados Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   'Entrada_Estoque
   
   cmdNovo.Enabled = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   LimparGrid_Produtos
   frmCadastro.Enabled = False
   Mostrar_Historico
End Sub

Private Sub cmdExibir_Click()
   'Indice
   Dim INDICE As String
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If optCodigo.Value = True Then
      INDICE = "produtos.cod_barra;"
   ElseIf optDescricao.Value = True Then
      INDICE = "produtos.descricao;"
   ElseIf optVenda.Value = True Then
      INDICE = "produtos_entrada_itens.venda;"
   ElseIf optCusto.Value = True Then
      INDICE = "produtos_entrada_itens.custo;"
   End If
   
   'Monta a consulta básica para não repetir várias linhas
   sSQL = "SELECT produtos.codigo AS var_codent, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
      "produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant, " & _
      "produtos_entrada_itens.custo AS var_custo, produtos_entrada_itens.frete AS var_frete, produtos_entrada_itens.imposto_valor_compra AS var_impcompra, " & _
      "produtos_entrada_itens.custo_compra AS var_vlrcompra, produtos_entrada_itens.imposto_valor_venda AS var_impvenda, " & _
      "produtos_entrada_itens.lucro_valor AS var_lucro, IFNULL(produtos_entrada_itens.venda, 0) AS venda, " & _
      "(produtos_entrada_itens.custo_compra * produtos.quant_estoque) AS var_totalcompra, (produtos_entrada_itens.custo * produtos.quant_estoque) AS var_totalcusto, (produtos_entrada_itens.frete * produtos.quant_estoque) AS var_totalfrete, " & _
      "(produtos_entrada_itens.imposto_valor_venda * produtos.quant_estoque) AS var_totalicompra, (produtos_entrada_itens.lucro_valor * produtos.quant_estoque) AS var_totallucro, (produtos_entrada_itens.imposto_valor_venda * produtos.quant_estoque) AS var_totalimpvenda, " & _
      "(IFNULL(produtos_entrada_itens.venda, 0) * produtos.quant_estoque) AS var_totalvenda " & _
      "FROM produtos LEFT JOIN ultimas_entradas ON produtos.codigo  = ultimas_entradas.codigo_produto " & _
      "LEFT JOIN produtos_entrada_itens ON ultimas_entradas.codigo_produto = produtos_entrada_itens.codigo_produto " & _
      "AND ultimas_entradas.ultentrada = produtos_entrada_itens.codigo_entrada " & _
      "WHERE "
   
   If optCodBarra.Value = True Then
      sSQL = sSQL & "(produtos.cod_barra = '" & cboConsProduto.Text & "') AND (produtos.ativo = 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf optTodos.Value = True Then
      sSQL = sSQL & "(produtos.ativo = 1) ORDER BY " & INDICE
      Debug.Print sSQL
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf optCategoria.Value = True Then
      sSQL = sSQL & "(produtos.categoria = '" & cboConsProduto.Text & "') AND (produtos.ativo = 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf optProduto.Value = True Then
      sSQL = sSQL & "(produtos.descricao = '" & cboConsProduto.Text & "') AND (produtos.ativo = 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   End If
   
   If optTodos.Value = False Then
      SelectControl cboConsProduto
   End If
   
   printSQL = sSQL
End Sub

Private Sub Command1_Click()
   'execSQL "UPDATE PRODUTOS_ENTRADA_ITENS SET CODIGO_ENTRADA = 1 WHERE CODIGO_ENTRADA = NULL"
End Sub

Private Sub Form_Load()
   SSTab1.Tab = 0
   LimparGrid_Produtos
   Mostrar_Historico
   cmdNovo.Enabled = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   Set moCombo = New cComboHelper
   
   'If Tela_Principal.txtNivel.Text <> "1" Then chkAtivo.Enabled = False: Exit Sub
   
   If Tela_Principal.txtNivel.Text <> "1" Then
      frmEstoque.Visible = False
      frmCompra.Visible = False
      frmVenda.Visible = False
   Else
      frmEstoque.Visible = True
      frmCompra.Visible = True
      frmVenda.Visible = True
   End If
End Sub

Private Sub Mostrar_Historico()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Monta a consulta básica
   sSQL = "SELECT produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada.codigo AS var_codigo " & _
      "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada "
   
   'Define o filtro
   If txtCodigo.Text = "" Then
      sSQL = sSQL & "WHERE false "
      
   Else
      sSQL = sSQL & "WHERE (codigo_produto = " & txtCodigo.Text & ") "
   
   End If
   
   'Monta a ordem de exibição
   sSQL = sSQL & "ORDER BY produtos_entrada.data_entrada, produtos_entrada.hora_entrada;"
   
   Set r = dbData.OpenRecordset(sSQL)
   FormatarGrid_Historico r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
   SSTab1.Tab = 0
   cmdNovo.Enabled = False
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   txtCodigo.Text = ""
   txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub Grid_Estoque_DblClick()
   Me.Hide
   'PRODUTOS_ENTRADA.Show
   'PRODUTOS_ENTRADA.frmPrincipal.Enabled = True
   'PRODUTOS_ENTRADA.frmSecundario.Enabled = True
   'PRODUTOS_ENTRADA.cmdSalvar.Visible = False
   'PRODUTOS_ENTRADA.cmdCancelar.Visible = False
   'PRODUTOS_ENTRADA.cmdAlterar.Visible = True
   'PRODUTOS_ENTRADA.cmdExcluir.Visible = True
   'PRODUTOS_ENTRADA.cmdNovo.Enabled = True
   'PRODUTOS_ENTRADA.frmPrincipal.Enabled = False
   'PRODUTOS_ENTRADA.frmSecundario.Enabled = False
   'PRODUTOS_ENTRADA.cmdAdicionar.Enabled = False
   'PRODUTOS_ENTRADA.cmdRemover.Enabled = False
   'PRODUTOS_ENTRADA.txtCodigo.Text = ""
   'PRODUTOS_ENTRADA.txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub optCategoria_Click()
   cboConsProduto.Visible = True
   lblNomeCombo.Visible = True
   lblNomeCombo.Caption = "Categoria"
   cboConsProduto.SetFocus
End Sub

Private Sub optCodBarra_Click()
   cboConsProduto.Visible = True
   lblNomeCombo.Visible = True
   lblNomeCombo.Caption = "Cód. de Barra"
   cboConsProduto.SetFocus
End Sub

Private Sub optCodigo_Click()
   cmdExibir_Click
End Sub

Private Sub optCusto_Click()
   cmdExibir_Click
End Sub

Private Sub optDescricao_Click()
   cmdExibir_Click
End Sub

Private Sub optProduto_Click()
   cboConsProduto.Visible = True
   lblNomeCombo.Visible = True
   lblNomeCombo.Caption = "Nome do Produto"
   cboConsProduto.SetFocus
End Sub

Private Sub optTodos_Click()
   cboConsProduto.Visible = False
   cboConsProduto.Visible = False
   lblNomeCombo.Visible = False
End Sub

Private Sub optVenda_Click()
   cmdExibir_Click
End Sub

Private Sub txtCodBarra_GotFocus()
   SelectControl txtCodBarra
End Sub

Private Sub txtCodBarra_LostFocus() 'Trocar pelo evento Validate
   '
End Sub

Private Sub txtCodBarra_Validate(Cancel As Boolean)
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodBarra.Text = "" Then Exit Sub
   txtCodBarra.Text = Trim(txtCodBarra.Text)
   
   'Verifica se existe o código de barras cadastrado
   sSQL = "SELECT codigo FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "');"
   Set r = dbData.OpenRecordset(sSQL)
   
   If cmdAlterar.Visible = False Then
      If r.RecordCount > 0 Then
         ShowMsg "Já existe um produto cadastrado com esse cód. de barra!", vbInformation
         Cancel = True           'Cancela a entrada e permanece com o foco no campo
         txtCodBarra.Text = ""   'Limpa a entrada
         Exit Sub                'Evita a saída do campo
      End If
   End If
End Sub

Private Sub txtCodigo_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cmdSalvar.Visible = False Then
      If txtCodigo.Text = "" Then Exit Sub
            
      'A auto numeração do código deve ser utilizada no momento de salvar o registro
      'para evitar duplicidade de código para quando houver mais de um terminal operando
      'ao mesmo tempo
      'AutoNumeracao
      
      'Faz a inserção de forma direta e verifica se houve algum erro
      'If Not Inserir_Dados Then
      '   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      '   Exit Sub
      'End If
      
      sSQL = "SELECT * FROM produtos WHERE (codigo = " & txtCodigo.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      LimparObjetos_Produtos
      cmdSalvar.Visible = False
      cmdCancelar.Visible = False
      cmdAlterar.Visible = True
      cmdExcluir.Visible = True
      frmCadastro.Enabled = True
      MostrarDados_Produto r
      Mostrar_Historico
   End If
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtOBS_LostFocus()
   If cmdSalvar.Visible = True And cmdCancelar.Visible = True Then
      cmdSalvar.SetFocus
   ElseIf cmdAlterar.Visible = True Then
      cmdAlterar.SetFocus
   Else
      Exit Sub
   End If
End Sub

Private Sub txtPrateleira_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtQuant_GotFocus()
   SelectControl txtQuant
End Sub

Private Sub txtValorAtual_GotFocus()
   SelectControl txtValorAtual
End Sub

Private Sub txtValorAtual_LostFocus()
   txtValorAtual.Text = Format(txtValorAtual.Text, "##,##0.00")
End Sub

Private Sub txtValorAtual_Validate(Cancel As Boolean)
   If txtValorAtual.Text = "" Then txtValorAtual = 0
   txtValorAtual.Text = Format$(txtValorAtual.Text, ocMONEY)
End Sub
