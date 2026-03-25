VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Produtos_Cadastro_Sapataria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRODUTOS"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13920
   Icon            =   "Produtos_Cadastro_Sapataria.frx":0000
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
      TabIndex        =   22
      Top             =   60
      Width           =   13755
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Produtos_Cadastro_Sapataria.frx":23D2
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
         TabIndex        =   23
         Top             =   240
         Width           =   1770
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8715
      Left            =   60
      TabIndex        =   14
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
      TabPicture(0)   =   "Produtos_Cadastro_Sapataria.frx":7DA5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CmdHabilitar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSair"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNovo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSalvar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdExcluir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAlterar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frmCadastro"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Produtos_Cadastro_Sapataria.frx":7DC1
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label25"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdImprimir"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdExibir"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Grid"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame9"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "HISTÓRICO"
      TabPicture(2)   =   "Produtos_Cadastro_Sapataria.frx":7DDD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid_Estoque"
      Tab(2).ControlCount=   1
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
         TabIndex        =   73
         Top             =   6960
         Width           =   3015
         Begin VB.CheckBox chkLinha 
            Caption         =   "Linha"
            Height          =   195
            Left            =   1620
            TabIndex        =   80
            Top             =   540
            Width           =   975
         End
         Begin VB.CheckBox chkProduto 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   540
            Width           =   1215
         End
         Begin VB.CheckBox chkCodBarra 
            Caption         =   "Cód. de Barra"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   780
            Width           =   1455
         End
         Begin VB.CheckBox chkTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   300
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.CheckBox chkRef 
            Caption         =   "Referência"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   1260
            Width           =   1215
         End
         Begin VB.CheckBox chkFab 
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CheckBox chkTam 
            Caption         =   "Tamanho"
            Height          =   195
            Left            =   1620
            TabIndex        =   74
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   120
         TabIndex        =   50
         Top             =   5760
         Width           =   10875
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Estoque 
         Height          =   8055
         Left            =   -74880
         TabIndex        =   47
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
         Left            =   4740
         TabIndex        =   44
         Top             =   6960
         Width           =   7335
         Begin VB.TextBox txtConsCodBarra 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   68
            Top             =   1080
            Width           =   2355
         End
         Begin VB.ComboBox cboConsLinha 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5100
            TabIndex        =   65
            Top             =   660
            Width           =   2175
         End
         Begin VB.ComboBox cboConsTam 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3780
            TabIndex        =   63
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cboConsRef 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   61
            Top             =   660
            Width           =   1875
         End
         Begin VB.ComboBox cboConsFab 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4980
            TabIndex        =   59
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox cboConsProduto 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   45
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label lblConsCodBarra 
            Caption         =   "Cod. Barra:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   67
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label lblConsLinha 
            Caption         =   "Linha:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4620
            TabIndex        =   66
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblConsTam 
            Caption         =   "Tamanho:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3000
            TabIndex        =   64
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lblConsRef 
            Caption         =   "Referência:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblConsFab 
            Caption         =   "Fabricante:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4140
            TabIndex        =   60
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblConsProduto 
            Caption         =   "Descrição:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   46
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "muda o cod_entrada dos nulos para 1"
         Top             =   5460
         Width           =   135
      End
      Begin VB.Frame Frame1 
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
         Height          =   1215
         Left            =   11040
         TabIndex        =   36
         Top             =   5760
         Width           =   2595
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
            TabIndex        =   72
            Top             =   300
            Width           =   540
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
            TabIndex        =   71
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
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
            Left            =   120
            TabIndex        =   70
            Top             =   540
            Width           =   795
         End
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
            TabIndex        =   69
            Top             =   540
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
            Left            =   960
            TabIndex        =   38
            Top             =   840
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
            TabIndex        =   37
            Top             =   840
            Width           =   615
         End
      End
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
         TabIndex        =   25
         Top             =   6960
         Width           =   1515
         Begin VB.CheckBox ckkORDQuant 
            Caption         =   "Quant."
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   1380
            Width           =   1215
         End
         Begin VB.CheckBox ckkORDDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox ckkORDLinha 
            Caption         =   "Linha"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1152
            Width           =   1215
         End
         Begin VB.CheckBox ckkORDRef 
            Caption         =   "Referência"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   696
            Width           =   1215
         End
         Begin VB.CheckBox ckkORDFab 
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   468
            Width           =   1335
         End
         Begin VB.CheckBox ckkORDTam 
            Caption         =   "Tamanho"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   924
            Width           =   975
         End
      End
      Begin VB.PictureBox frmCadastro 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   -74880
         ScaleHeight     =   1755
         ScaleWidth      =   13455
         TabIndex        =   15
         Top             =   420
         Width           =   13515
         Begin VB.TextBox txtPrateleira 
            Height          =   315
            Left            =   3420
            MaxLength       =   4
            TabIndex        =   10
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtRef 
            Height          =   315
            Left            =   6120
            TabIndex        =   3
            Top             =   300
            Width           =   1515
         End
         Begin VB.TextBox txtTam 
            Height          =   315
            Left            =   9360
            MaxLength       =   20
            TabIndex        =   5
            Top             =   300
            Width           =   1455
         End
         Begin VB.ComboBox cboFabricante 
            Height          =   315
            Left            =   7680
            TabIndex        =   4
            Top             =   300
            Width           =   1635
         End
         Begin VB.TextBox txtValorAtual 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   1740
            TabIndex        =   9
            Top             =   960
            Width           =   1635
         End
         Begin VB.ComboBox cboUnidMedida 
            Height          =   315
            Left            =   10860
            TabIndex        =   6
            Top             =   300
            Width           =   855
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   1380
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.TextBox txtObs 
            Height          =   315
            Left            =   6420
            MaxLength       =   90
            TabIndex        =   13
            Top             =   960
            Width           =   6975
         End
         Begin VB.TextBox txtCodigo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   12900
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   -60
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   2160
            MaxLength       =   90
            TabIndex        =   2
            Top             =   300
            Width           =   3915
         End
         Begin VB.ComboBox cboCategoria 
            Height          =   315
            Left            =   11760
            TabIndex        =   7
            Top             =   300
            Width           =   1635
         End
         Begin VB.TextBox txtQuant 
            Enabled         =   0   'False
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txtQuantMin 
            Height          =   315
            Left            =   4200
            TabIndex        =   11
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtUltCompra 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local"
            Height          =   195
            Left            =   3420
            TabIndex        =   82
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            Height          =   195
            Left            =   6120
            TabIndex        =   53
            Top             =   60
            Width           =   300
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tam."
            Height          =   195
            Left            =   9360
            TabIndex        =   52
            Top             =   60
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   7680
            TabIndex        =   51
            Top             =   60
            Width           =   750
         End
         Begin VB.Label lblValorAtual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Atual"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1740
            TabIndex        =   48
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unid. Med."
            Height          =   195
            Left            =   10860
            TabIndex        =   40
            Top             =   60
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. Barra"
            Height          =   195
            Left            =   60
            TabIndex        =   34
            Top             =   60
            Width           =   750
         End
         Begin VB.Label Observação 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observação"
            Height          =   195
            Left            =   6420
            TabIndex        =   24
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            Height          =   195
            Left            =   2340
            TabIndex        =   21
            Top             =   60
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
            Height          =   195
            Left            =   11760
            TabIndex        =   20
            Top             =   60
            Width           =   675
         End
         Begin VB.Label lblQuantAtual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Atual"
            Enabled         =   0   'False
            Height          =   195
            Left            =   60
            TabIndex        =   19
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
            Left            =   5160
            TabIndex        =   18
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Min."
            Height          =   195
            Left            =   4200
            TabIndex        =   17
            Top             =   720
            Width           =   825
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4935
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   13515
         _ExtentX        =   23839
         _ExtentY        =   8705
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   555
         Left            =   -73140
         TabIndex        =   30
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":7DF9
         PICN            =   "Produtos_Cadastro_Sapataria.frx":7E15
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
         TabIndex        =   31
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":86EF
         PICN            =   "Produtos_Cadastro_Sapataria.frx":870B
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
         TabIndex        =   32
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":8A25
         PICN            =   "Produtos_Cadastro_Sapataria.frx":8A41
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
         TabIndex        =   33
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":F30B
         PICN            =   "Produtos_Cadastro_Sapataria.frx":F327
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":15DCB
         PICN            =   "Produtos_Cadastro_Sapataria.frx":15DE7
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
         TabIndex        =   39
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":16AC1
         PICN            =   "Produtos_Cadastro_Sapataria.frx":16ADD
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
         TabIndex        =   41
         Top             =   7080
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":16DF7
         PICN            =   "Produtos_Cadastro_Sapataria.frx":16E13
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
         TabIndex        =   42
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":176ED
         PICN            =   "Produtos_Cadastro_Sapataria.frx":17709
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton CmdHabilitar 
         Height          =   555
         Left            =   -69660
         TabIndex        =   49
         Top             =   2340
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "&Habilitar Estoque"
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
         MICON           =   "Produtos_Cadastro_Sapataria.frx":17A23
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
         TabIndex        =   27
         Top             =   5460
         Width           =   3435
         WordWrap        =   -1  'True
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   35
      Top             =   9855
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   20585
            Text            =   "Desenv.: OnLine.Info - Informática & Lan House - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: OnLine.Info - Informática & Lan House - Tel.: (89) 3544-2553"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "18:44"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            Key             =   ""
            Object.Tag             =   ""
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
Attribute VB_Name = "Produtos_Cadastro_Sapataria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moCombo As cComboHelper
Dim SQL As String
Dim RS As Recordset
Dim SQL2 As String
Dim Rs2 As Recordset
Dim var_cod_Preco As Long



Private Sub Desativa_Objetos()
If chkProduto.Value = Unchecked Then lblConsProduto.Enabled = False: cboConsProduto.Enabled = False Else lblConsProduto.Enabled = True: cboConsProduto.Enabled = True
If chkCodBarra.Value = Unchecked Then lblConsCodBarra.Enabled = False: txtConsCodBarra.Enabled = False Else lblConsCodBarra.Enabled = True: txtConsCodBarra.Enabled = True
If chkFab.Value = Unchecked Then lblConsFab.Enabled = False: cboConsFab.Enabled = False Else lblConsFab.Enabled = True: cboConsFab.Enabled = True
If chkRef.Value = Unchecked Then lblConsRef.Enabled = False: cboConsRef.Enabled = False Else lblConsRef.Enabled = True: cboConsRef.Enabled = True
If chkTam.Value = Unchecked Then lblConsTam.Enabled = False: cboConsTam.Enabled = False Else lblConsTam.Enabled = True: cboConsTam.Enabled = True
If chkLinha.Value = Unchecked Then lblConsLinha.Enabled = False: cboConsLinha.Enabled = False Else lblConsLinha.Enabled = True: cboConsLinha.Enabled = True
End Sub

Private Sub Entrada_Estoque()
        'AUTONUMERAÇÃO
        Call Abrir_BancodeDados
        SQL = "SELECT MAX(CODIGO) as COD_ITENS FROM PRODUTOS_ENTRADA_ITENS"
        Set RS = BD.OpenRecordset(SQL)
     
        Dim var_COD_Itens As Long
        var_COD_Itens = IIf(IsNull(RS!COD_ITENS) = True, 1, RS!COD_ITENS + 1)
        
        'ENTRADA DO PRODUTO
        Call Abrir_BancodeDados
        SQL = "SELECT * FROM PRODUTOS_ENTRADA_ITENS"
        Set RS = BD.OpenRecordset(SQL)
        
        RS.AddNew
        RS!Codigo = var_COD_Itens
        RS!CODIGO_ENTRADA = "0001"
        RS!CODIGO_PRODUTO = IIf(txtCodigo.Text = "", Null, txtCodigo.Text)
        RS!DESCRICAO = IIf(txtDescricao.Text = "", Null, txtDescricao.Text)
        RS!QUANT = IIf(txtQuant.Text = "", Null, txtQuant.Text)
        
        RS!CUSTO = IIf(txtValorAtual.Text = "", Null, txtValorAtual.Text)
        RS!IMPOSTO_VALOR_COMPRA = "0"
        RS!FRETE = "0"
        RS!CUSTO_COMPRA = IIf(txtValorAtual.Text = "", Null, txtValorAtual.Text)
       
        RS!LUCRO_VALOR = "0"
        RS!IMPOSTO_VALOR_VENDA = "0"
        RS!VENDA = IIf(txtValorAtual.Text = "", Null, txtValorAtual.Text)

        'IMPOSTO COMPRA
        RS!IMPOSTO_COMPRA = "0"
        
        'IMPOSTO_STATUS_COMPRA

        RS!IMPOSTO_STATUS_COMPRA = 1

        
        'LUCRO
        RS!LUCRO = "0"

        'LUCRO_STATUS
        RS!LUCRO_STATUS = 1
        
        'IMPOSTO_VENDA
        RS!IMPOSTO_VENDA = "0"

        'IMPOSTO_STATUS_VENDA
        RS!IMPOSTO_STATUS_VENDA = 1

        RS.Update
        
        execSQL "UPDATE PRODUTOS SET QUANT_ESTOQUE = " & Replace(txtQuant.Text, ",", ".") & " WHERE CODIGO = " & txtCodigo.Text
End Sub

Private Sub FormatarGrid_Historico()
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
    Dim X As Integer
    For X = 0 To .Cols - 1
    .Col = X
    .Row = 0
    .CellFontBold = True
    Next X
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until RS.EOF
    
    'mudar a cor da coluna
    'Dim i As Integer
    'For i = 1 To .Rows - 1
   '.Row = i
   '.Col = 6:   .CellBackColor = &HC0FFFF
   ' Next

    
    Grid.Redraw = False
    
    'ALINHAMENTO
    '.ColAlignment(2) = 1
    
    Grid.Redraw = True
    
    If Not IsNull(RS!VAR_CODIGO) Then .TextMatrix(.Rows - 1, 1) = RS!VAR_CODIGO
    If Not IsNull(RS!DATA_ENTRADA) Then .TextMatrix(.Rows - 1, 2) = Format(RS!DATA_ENTRADA, "dd/mm/yy")
    If Not IsNull(RS!NOTAFISCAL) Then .TextMatrix(.Rows - 1, 3) = RS!NOTAFISCAL
    If Not IsNull(RS!FORNECEDOR) Then .TextMatrix(.Rows - 1, 4) = RS!FORNECEDOR
    If Not IsNull(RS!QUANT) Then .TextMatrix(.Rows - 1, 5) = RS!QUANT
    If Not IsNull(RS!CUSTO) Then .TextMatrix(.Rows - 1, 6) = Format(RS!CUSTO, "##,##0.00")
    RS.MoveNext
    .Rows = .Rows + 1
        
    Loop
    
    .Rows = .Rows - 1
    
End With
End Sub
Private Sub FormatarGrid_Produtos()
With Grid
    
    .Clear
    .Cols = 23
    .Rows = 2
        
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 1450 '1600
    .ColWidth(3) = 1100
    .ColWidth(4) = 4245 '4445
    .ColWidth(5) = 1200
    .ColWidth(6) = 530
    .ColWidth(7) = 1800
    .ColWidth(8) = 1000
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = 1000
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(20) = 0
    .ColWidth(21) = 0
    .ColWidth(22) = 0
    
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "CÓD. BARRA"
    .TextMatrix(0, 3) = "REF."
    .TextMatrix(0, 4) = "PRODUTO"
    .TextMatrix(0, 5) = "FAB."
    .TextMatrix(0, 6) = "TAM."
    .TextMatrix(0, 7) = "LINHA"
    .TextMatrix(0, 8) = "QUANT"
    .TextMatrix(0, 9) = "CUSTO"
    .TextMatrix(0, 10) = "FRETE"
    .TextMatrix(0, 11) = "IMPOSTO"
    .TextMatrix(0, 12) = "VALOR"
    .TextMatrix(0, 13) = "LUCRO"
    .TextMatrix(0, 14) = "IMPOSTO"
    .TextMatrix(0, 15) = "VENDA"
    .TextMatrix(0, 16) = "T_venda"
    .TextMatrix(0, 17) = "T_IVenda"
    .TextMatrix(0, 18) = "T_Lucro"
    .TextMatrix(0, 19) = "T_Custo"
    .TextMatrix(0, 20) = "DIF"  '"T_Frete"
    .TextMatrix(0, 21) = "T_Compra"
    .TextMatrix(0, 22) = "T_Compra"

    
    'colocar os cabeçalho em negrito
    Dim X As Integer
    For X = 0 To .Cols - 1
    .Col = X
    .Row = 0
    .CellFontBold = True
    Next X
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until Rs2.EOF
    
    'mudar a cor da coluna
    'Dim i As Integer
    'For i = 1 To .Rows - 1
   '.Row = i
   '.Col = 6:   .CellBackColor = &HC0FFFF
   ' Next

    
    .Redraw = False
    
    'ALINHAMENTO
    '.ColAlignment(2) = 1
    
    
    
    If Not IsNull(Rs2!var_CodEnt) Then .TextMatrix(.Rows - 1, 1) = Rs2!var_CodEnt
    If Not IsNull(Rs2!var_CodBarra) Then .TextMatrix(.Rows - 1, 2) = Rs2!var_CodBarra
    If Not IsNull(Rs2!var_REF) Then .TextMatrix(.Rows - 1, 3) = Rs2!var_REF
    If Not IsNull(Rs2!var_desc) Then .TextMatrix(.Rows - 1, 4) = Rs2!var_desc
    If Not IsNull(Rs2!VAR_FAB) Then .TextMatrix(.Rows - 1, 5) = Rs2!VAR_FAB
    If Not IsNull(Rs2!VAR_TAM) Then .TextMatrix(.Rows - 1, 6) = Rs2!VAR_TAM
    If Not IsNull(Rs2!var_Linha) Then .TextMatrix(.Rows - 1, 7) = Rs2!var_Linha
    If Not IsNull(Rs2!var_Quant) Then .TextMatrix(.Rows - 1, 8) = Rs2!var_Quant
    If Not IsNull(Rs2!var_CUSTO) Then .TextMatrix(.Rows - 1, 9) = Format(Rs2!var_CUSTO, "##,##0.00")
    If Not IsNull(Rs2!var_FRETE) Then .TextMatrix(.Rows - 1, 10) = Format(Rs2!var_FRETE, "##,##0.00")
    If Not IsNull(Rs2!var_impcompra) Then .TextMatrix(.Rows - 1, 11) = Format(Rs2!var_impcompra, "##,##0.00")
    If Not IsNull(Rs2!var_vlrcompra) Then .TextMatrix(.Rows - 1, 12) = Format(Rs2!var_vlrcompra, "##,##0.00")
    If Not IsNull(Rs2!Var_Lucro) Then .TextMatrix(.Rows - 1, 13) = Format(Rs2!Var_Lucro, "##,##0.00")
    If Not IsNull(Rs2!var_ImpVenda) Then .TextMatrix(.Rows - 1, 14) = Format(Rs2!var_ImpVenda, "##,##0.00")
    If Not IsNull(Rs2!VENDA) Then .TextMatrix(.Rows - 1, 15) = Format(Rs2!VENDA, "##,##0.00")
    If Not IsNull(Rs2!VAR_TOTALVENDA) Then .TextMatrix(.Rows - 1, 16) = Format(Rs2!VAR_TOTALVENDA, "##,##0.00")
    If Not IsNull(Rs2!var_TotalImpVenda) Then .TextMatrix(.Rows - 1, 17) = Format(Rs2!var_TotalImpVenda, "##,##0.00")
    If Not IsNull(Rs2!var_TotalLucro) Then .TextMatrix(.Rows - 1, 18) = Format(Rs2!var_TotalLucro, "##,##0.00")
    If Not IsNull(Rs2!var_TotalCusto) Then .TextMatrix(.Rows - 1, 19) = Format(Rs2!var_TotalCusto, "##,##0.00")
'    If Not IsNull(Rs2!var_DIFERENCA) Then .TextMatrix(.Rows - 1, 20) = Format(Rs2!var_DIFERENCA, "##,##0.00")
    If Not IsNull(Rs2!var_TotalICompra) Then .TextMatrix(.Rows - 1, 21) = Format(Rs2!var_TotalICompra, "##,##0.00")
    If Not IsNull(Rs2!var_TotalCompra) Then .TextMatrix(.Rows - 1, 22) = Format(Rs2!var_TotalCompra, "##,##0.00")
    
    Rs2.MoveNext
    .Rows = .Rows + 1
    
    Loop
    
    .Rows = .Rows - 1
    .Redraw = True
End With

'Estoque
lblEstoque.Caption = SomaGrid(Grid, 8)
lblProdutos.Caption = Grid.Rows - 1  'contar o numeros de linhas no grid
lblVenda.Caption = Format(SomaGrid(Grid, 16), "##,##0.00")
End Sub

Private Sub LimparGrid_Produtos()
    Call Abrir_BancodeDados
    SQL2 = "SELECT (PRODUTOS.CODIGO) as var_CodEnt, (PRODUTOS.DESCRICAO) as var_Desc, (PRODUTOS.QUANT_ESTOQUE) as var_Quant, (PRODUTOS_ENTRADA_ITENS.CUSTO) as var_Custo, (PRODUTOS_ENTRADA_ITENS.FRETE) as var_Frete, (PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_COMPRA) as var_impcompra, (PRODUTOS_ENTRADA_ITENS.CUSTO_COMPRA) as var_vlrcompra, (PRODUTOS_ENTRADA_ITENS.LUCRO_VALOR) as var_lucro, (PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_VENDA) as var_impvenda, (PRODUTOS_ENTRADA_ITENS.VENDA) as var_vlrvenda, PRODUTOS.CODIGO, PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO FROM PRODUTOS INNER JOIN PRODUTOS_ENTRADA_ITENS ON PRODUTOS.CODIGO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO WHERE FALSE"
    Set Rs2 = BD.OpenRecordset(SQL2)
    
    FormatarGrid_Produtos
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
Private Sub MostrarDados_Produto()
If Not IsNull(Rs2!Codigo) Then txtCodigo.Text = Rs2!Codigo

    'mostrar o ultimo preço de compra
    Call Abrir_BancodeDados
    SQL = "Select * From PRODUTOS_ENTRADA_ITENS where CODIGO_PRODUTO = " & txtCodigo.Text & " ORDER BY CODIGO"
    Set RS = BD.OpenRecordset(SQL)
    
    If Not RS.EOF Then
    RS.MoveLast
    If Not IsNull(RS.Fields!VENDA) Then txtValorAtual.Text = Format(RS.Fields!VENDA, "##,##0.00")
    End If
    

If Not IsNull(Rs2!COD_BARRA) Then txtCodBarra.Text = Rs2!COD_BARRA
If Not IsNull(Rs2!DESCRICAO) Then txtDescricao.Text = Rs2!DESCRICAO
If Not IsNull(Rs2!UNID_MEDIDA) Then cboUnidMedida.Text = Rs2!UNID_MEDIDA
If Not IsNull(Rs2!CATEGORIA) Then cboCategoria.Text = Rs2!CATEGORIA
If Not IsNull(Rs2!Prateleira) Then txtPrateleira.Text = Rs2!Prateleira
If Not IsNull(Rs2!Quant_Estoque) Then txtQuant.Text = Rs2!Quant_Estoque
If Not IsNull(Rs2!QUANT_MIN) Then txtQuantMin.Text = Rs2!QUANT_MIN
If Not IsNull(Rs2!ULT_COMPRA) Then txtUltCompra.Text = Format(Rs2!ULT_COMPRA, "dd/mm/yy") Else txtUltCompra.Text = ""
If Not IsNull(Rs2!OBSERVACAO) Then txtObs.Text = Rs2!OBSERVACAO
If Not IsNull(Rs2!REF) Then txtRef.Text = Rs2!REF
If Not IsNull(Rs2!FABRICANTE) Then cboFabricante.Text = Rs2!FABRICANTE
If Not IsNull(Rs2!TAMANHO) Then txtTam.Text = Rs2!TAMANHO

If Rs2!ATIVO = True Then chkAtivo.Value = 1 Else chkAtivo.Value = 0
If Rs2!DESTAQUE = True Then chkDestaque.Value = 1 Else chkDestaque.Value = 0
End Sub
Private Sub Autonumeracao()
Call Abrir_BancodeDados
SQL = "SELECT MAX(CODIGO) as COD_PRODUTO FROM PRODUTOS"
Set RS = BD.OpenRecordset(SQL)
    
txtCodigo.Text = IIf(IsNull(RS!COD_PRODUTO) = True, 1, RS!COD_PRODUTO + 1)
End Sub
Private Sub LimparObjetos_Produtos()
If cmdAlterar.Visible = False Then txtCodigo.Text = ""
txtCodBarra.Text = ""
txtDescricao.Text = ""
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
txtRef.Text = ""
cboFabricante.Text = ""
txtTam.Text = ""

cmdNovo.Enabled = True
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = False
cmdExcluir.Visible = False
lblQuantAtual.Enabled = False
lblValorAtual.Enabled = False
txtQuant.Enabled = False
txtValorAtual.Enabled = False
End Sub

Private Sub cboCategoria_GotFocus()
cboCategoria.Clear
Call Abrir_BancodeDados
SQL = "SELECT DISTINCT CATEGORIA FROM PRODUTOS ORDER BY CATEGORIA"
Set RS = BD.OpenRecordset(SQL)

            
While Not RS.EOF
If Not IsNull(RS.Fields("CATEGORIA")) Then cboCategoria.AddItem RS.Fields("CATEGORIA")
RS.MoveNext
Wend

moCombo.AttachTo cboCategoria
End Sub
Private Sub cboCategoria_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboConsFab_GotFocus()
If cboConsFab.ListCount = 0 Then
Call Abrir_BancodeDados
SQL = "SELECT DISTINCT FABRICANTE FROM PRODUTOS ORDER BY FABRICANTE"
Set RS = BD.OpenRecordset(SQL)
    
    While Not RS.EOF
    cboConsFab.AddItem RS("FABRICANTE") & ""
    RS.MoveNext
    Wend
End If
    moCombo.AttachTo cboConsFab
End Sub


Private Sub cboConsLinha_GotFocus()
If cboConsLinha.ListCount = 0 Then
Call Abrir_BancodeDados
SQL = "SELECT DISTINCT CATEGORIA FROM PRODUTOS ORDER BY CATEGORIA"
Set RS = BD.OpenRecordset(SQL)
    
    While Not RS.EOF
    cboConsLinha.AddItem RS("CATEGORIA") & ""
    RS.MoveNext
    Wend
End If
    moCombo.AttachTo cboConsLinha
End Sub


Private Sub cboConsProduto_GotFocus()
If cboConsProduto.ListCount = 0 Then
Call Abrir_BancodeDados
SQL = "SELECT DISTINCT DESCRICAO FROM PRODUTOS ORDER BY DESCRICAO"
Set RS = BD.OpenRecordset(SQL)
    
    While Not RS.EOF
    cboConsProduto.AddItem RS("DESCRICAO") & ""
    RS.MoveNext
    Wend
End If
    moCombo.AttachTo cboConsProduto
End Sub


Private Sub cboConsRef_GotFocus()
If cboConsRef.ListCount = 0 Then
Call Abrir_BancodeDados
SQL = "SELECT DISTINCT REF FROM PRODUTOS ORDER BY REF"
Set RS = BD.OpenRecordset(SQL)
    
    While Not RS.EOF
    cboConsRef.AddItem RS("REF") & ""
    RS.MoveNext
    Wend
End If
    moCombo.AttachTo cboConsRef
End Sub


Private Sub cboConsTam_GotFocus()
If cboConsTam.ListCount = 0 Then
Call Abrir_BancodeDados
SQL = "SELECT DISTINCT TAMANHO FROM PRODUTOS ORDER BY TAMANHO"
Set RS = BD.OpenRecordset(SQL)
    
    While Not RS.EOF
    cboConsTam.AddItem RS("TAMANHO") & ""
    RS.MoveNext
    Wend
End If
    moCombo.AttachTo cboConsTam
End Sub


Private Sub cboFabricante_GotFocus()
cboFabricante.Clear
Call Abrir_BancodeDados
SQL = "SELECT DISTINCT FABRICANTE FROM PRODUTOS ORDER BY FABRICANTE"
Set RS = BD.OpenRecordset(SQL)

            
While Not RS.EOF
If Not IsNull(RS.Fields("FABRICANTE")) Then cboFabricante.AddItem RS.Fields("FABRICANTE")
RS.MoveNext
Wend

moCombo.AttachTo cboFabricante
End Sub


Private Sub cboFabricante_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboUnidMedida_GotFocus()
If cboUnidMedida.ListCount = 0 Then
    cboUnidMedida.AddItem "UNID"
    cboUnidMedida.AddItem "CX"
    cboUnidMedida.AddItem "M"
    cboUnidMedida.AddItem "M²"
    cboUnidMedida.AddItem "M³"
    cboUnidMedida.AddItem "ML"
    cboUnidMedida.AddItem "KG"
    cboUnidMedida.AddItem "G"
    cboUnidMedida.AddItem "PAR"
End If
moCombo.AttachTo cboUnidMedida
End Sub


Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Double
Dim i As Integer, Valor As Double
For i = 0 To Grid.Rows - 1
  If IsNumeric(Grid.TextMatrix(i, Col)) Then
    Valor = Valor + CDbl(Grid.TextMatrix(i, Col))
  End If
Next i
SomaGrid = Valor
End Function

Private Sub chkFab_Click()
chkTodos.Value = Unchecked
chkProduto.Value = Unchecked
chkCodBarra.Value = Unchecked
'chkFab.Value = Unchecked
'chkRef.Value = Unchecked
'chkTam.Value = Unchecked
'chkLinha.Value = Unchecked
Desativa_Objetos
If chkFab.Value = Checked Then cboConsFab.SetFocus
End Sub

Private Sub chkRef_Click()
chkTodos.Value = Unchecked
chkProduto.Value = Unchecked
chkCodBarra.Value = Unchecked
'chkFab.Value = Unchecked
'chkRef.Value = Unchecked
'chkTam.Value = Unchecked
'chkLinha.Value = Unchecked
Desativa_Objetos
If chkRef.Value = Checked Then cboConsRef.SetFocus
End Sub


Private Sub chkTam_Click()
chkTodos.Value = Unchecked
chkProduto.Value = Unchecked
chkCodBarra.Value = Unchecked
'chkFab.Value = Unchecked
'chkRef.Value = Unchecked
'chkTam.Value = Unchecked
'chkLinha.Value = Unchecked
Desativa_Objetos
If chkTam.Value = Checked Then cboConsTam.SetFocus
End Sub


Private Sub ckkORDDesc_Click()
cmdExibir_Click
End Sub

Private Sub ckkORDFab_Click()
cmdExibir_Click
End Sub


Private Sub ckkORDLinha_Click()
cmdExibir_Click
End Sub

Private Sub ckkORDQuant_Click()
cmdExibir_Click
End Sub


Private Sub ckkORDRef_Click()
cmdExibir_Click
End Sub


Private Sub ckkORDTam_Click()
cmdExibir_Click
End Sub


Private Sub cmdAlterar_Click()
If txtCodigo.Text = "" Then
    MsgBox "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o produto na guia CONSULTA.", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

    Call Abrir_BancodeDados
    SQL = "SELECT * FROM PRODUTOS WHERE (CODIGO = " & txtCodigo.Text & ")"
    Set RS = BD.OpenRecordset(SQL)

    If Not RS.EOF Then
        RS.Edit
        Atualizar_Dados
        RS.Update
    End If
    
    'alterar o nome dos produtos da tabela de entrada e pedidos
    execSQL "UPDATE PRODUTOS_ENTRADA_ITENS SET DESCRICAO = '" & txtDescricao.Text & "' WHERE CODIGO_PRODUTO = " & txtCodigo.Text
    execSQL "UPDATE PEDIDOS_ITENS SET DESCRICAO = '" & txtDescricao.Text & "' WHERE COD_PRODUTO = " & txtCodigo.Text
    
    'alterar o valor da ultima entrada no estoque
    Call Abrir_BancodeDados
    SQL = "Select * From PRODUTOS_ENTRADA_ITENS where CODIGO_PRODUTO = " & txtCodigo.Text & " ORDER BY CODIGO"
    Set RS = BD.OpenRecordset(SQL)
    
    If Not RS.EOF Then
    RS.MoveLast
    RS.Edit
    RS.Fields!VENDA = IIf(txtValorAtual.Text = "", Null, txtValorAtual.Text)
    RS.Update
    End If
    
    cmdNovo.Enabled = True
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    cmdAlterar.Visible = False
    cmdExcluir.Visible = False
    CmdHabilitar.Visible = False
    lblQuantAtual.Enabled = False
    lblValorAtual.Enabled = False
    txtQuant.Enabled = False
    txtValorAtual.Enabled = False
    frmCadastro.Enabled = False
    LimparGrid_Produtos
    Mostrar_Historico
End Sub
Private Sub Atualizar_Dados()
If chkAtivo.Value = 1 Then RS!ATIVO = True Else RS!ATIVO = False
If chkDestaque.Value = 1 Then RS!DESTAQUE = True Else RS!DESTAQUE = False

RS!Codigo = IIf(txtCodigo.Text = "", Null, txtCodigo.Text)
RS!COD_BARRA = IIf(txtCodBarra.Text = "", Null, txtCodBarra.Text)
RS!DESCRICAO = IIf(txtDescricao.Text = "", Null, txtDescricao.Text)
RS!UNID_MEDIDA = IIf(cboUnidMedida.Text = "", Null, cboUnidMedida.Text)
RS!CATEGORIA = IIf(cboCategoria.Text = "", Null, cboCategoria.Text)
RS!Prateleira = IIf(txtPrateleira.Text = "", Null, txtPrateleira.Text)
RS!QUANT_MIN = IIf(txtQuantMin.Text = "", Null, txtQuantMin.Text)
RS!OBSERVACAO = IIf(txtObs.Text = "", Null, txtObs.Text)
RS!Quant_Estoque = IIf(txtQuant.Text = "", Null, txtQuant.Text)
RS!REF = IIf(txtRef.Text = "", Null, txtRef.Text)
RS!FABRICANTE = IIf(cboFabricante.Text = "", Null, cboFabricante.Text)
RS!TAMANHO = IIf(txtTam.Text = "", Null, txtTam.Text)
End Sub

Private Sub cmdCancelar_Click()
LimparObjetos_Produtos
frmCadastro.Enabled = False
End Sub


Private Sub cmdExcluir_Click()
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub

If txtCodigo.Text = "" Then
    MsgBox "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o produto na guia CONSULTA", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If MsgBox("Excluir esse produto?", vbInformation + vbYesNo, "Aviso do Sistema") = vbYes Then
    
    execSQL "DELETE * FROM PRODUTOS WHERE CODIGO = " & txtCodigo.Text & ""
    
    LimparObjetos_Produtos
    cmdNovo.Enabled = True
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    cmdAlterar.Visible = False
    cmdExcluir.Visible = False
    CmdHabilitar.Visible = False
    lblQuantAtual.Enabled = False
    lblValorAtual.Enabled = False
    txtQuant.Enabled = False
    txtValorAtual.Enabled = False
    frmCadastro.Enabled = False
    LimparGrid_Produtos
    Mostrar_Historico
End If
End Sub

Private Sub CmdHabilitar_Click()
Call Abrir_BancodeDados
SQL = "Select * From PRODUTOS_ENTRADA_ITENS where CODIGO_PRODUTO = " & txtCodigo.Text & " ORDER BY CODIGO"
Set RS = BD.OpenRecordset(SQL)
    
If RS.RecordCount = 0 Then
    MsgBox "Não existe nenhuma entrada no estoque para esse produto!", vbInformation, "Aviso do Sistema"
    Exit Sub
Else
    frmCadastro.Enabled = True
    'frmComp.Enabled = True
    lblQuantAtual.Enabled = True
    lblValorAtual.Enabled = True
    txtQuant.Enabled = True
    txtValorAtual.Enabled = True
    txtQuant.SetFocus
End If
End Sub

Private Sub cmdImprimir_Click()
Me.Hide
Set REL_Prod_Cad_Imp.Relatorio.Recordset = Rs2
REL_Prod_Cad_Imp.rfTipo.Caption = lblProdutos.Caption
REL_Prod_Cad_Imp.rfITENS.Caption = lblEstoque.Caption
REL_Prod_Cad_Imp.rfVENDA.Caption = lblVenda.Caption

REL_Prod_Cad_Imp.Relatorio.Ativar
Unload REL_Prod_Cad_Imp
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
CmdHabilitar.Visible = False
chkAtivo.Value = Checked
Autonumeracao
'lblQuantAtual.Enabled = True
'lblValorAtual.Enabled = True
'txtQuant.Enabled = True
'txtValorAtual.Enabled = True
txtCodBarra.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub cmdSalvar_Click()
If txtDescricao.Text = "" Then MsgBox "Digite a Descrição do produto", vbInformation, "Aviso do Sistema": txtDescricao.SetFocus: Exit Sub

    Call Abrir_BancodeDados
    SQL = "SELECT * FROM PRODUTOS"
    Set RS = BD.OpenRecordset(SQL)
    
        RS.AddNew
        Atualizar_Dados
        RS.Update

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
    'INDICE
    Dim var_Criterio As String
    var_Criterio = ""
    var_Criterio = var_Criterio + IIf(ckkORDRef.Value, IIf(var_Criterio <> "", ", ", "") + "PRODUTOS.REF", "")
    var_Criterio = var_Criterio + IIf(ckkORDDesc.Value, IIf(var_Criterio <> "", ", ", "") + "PRODUTOS.DESCRICAO", "")
    var_Criterio = var_Criterio + IIf(ckkORDFab.Value, IIf(var_Criterio <> "", ", ", "") + "PRODUTOS.FABRICANTE", "")
    var_Criterio = var_Criterio + IIf(ckkORDTam.Value, IIf(var_Criterio <> "", ", ", "") + "PRODUTOS.TAMANHO", "")
    var_Criterio = var_Criterio + IIf(ckkORDLinha.Value, IIf(var_Criterio <> "", ", ", "") + "PRODUTOS.CATEGORIA", "")
    var_Criterio = var_Criterio + IIf(ckkORDQuant.Value, IIf(var_Criterio <> "", ", ", "") + "PRODUTOS.QUANT_ESTOQUE", "")
    
    If var_Criterio <> "" Then var_Criterio = " ORDER BY " + var_Criterio
    
    'Dim var_Indice As String
    'var_Indice = ""
    'var_Indice = var_Indice + IIf(ckkORDRef.Value, IIf(var_Indice <> "", ", ", "") + "PRODUTOS.REF", "")
    'var_Indice = var_Indice + IIf(ckkORDDesc.Value, IIf(var_Indice <> "", ", ", "") + "PRODUTOS.DESCRICAO", "")
    'var_Indice = var_Indice + IIf(ckkORDFab.Value, IIf(var_Indice <> "", ", ", "") + "PRODUTOS.FABRICANTE", "")
    'var_Indice = var_Indice + IIf(ckkORDTam.Value, IIf(var_Indice <> "", ", ", "") + "PRODUTOS.TAMANHO", "")
    'var_Indice = var_Indice + IIf(ckkORDLinha.Value, IIf(var_Indice <> "", ", ", "") + "PRODUTOS.CATEGORIA", "")
    'var_Indice = var_Indice + IIf(ckkORDQuant.Value, IIf(var_Indice <> "", ", ", "") + "PRODUTOS.QUANT_ESTOQUE", "")
    
    'If var_Indice <> "" Then var_Indice = " ORDER BY " + var_Indice
    
    Call Abrir_BancodeDados

If chkTodos.Value = Checked Then
    SQL2 = "SELECT (PRODUTOS.REF) as var_Ref,(PRODUTOS.FABRICANTE) as var_Fab,(PRODUTOS.TAMANHO) as var_Tam, (PRODUTOS.CATEGORIA) as var_linha, (PRODUTOS.CODIGO) as var_codEnt, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, " & _
    " (PRODUTOS.QUANT_ESTOQUE) as var_Quant, PRODUTOS_ENTRADA_ITENS.CUSTO AS var_Custo, PRODUTOS_ENTRADA_ITENS.FRETE AS var_Frete, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_COMPRA AS var_ImpCompra, " & _
    " PRODUTOS_ENTRADA_ITENS.CUSTO_COMPRA AS var_VlrCompra, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_VENDA AS var_ImpVenda, PRODUTOS_ENTRADA_ITENS.LUCRO_VALOR AS Var_Lucro, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA, " & _
    " (var_VlrCompra * var_Quant) as var_TotalCompra, (var_Custo * var_Quant) as var_TotalCusto, (var_Frete * var_Quant) as var_TotalFrete, (var_ImpCompra * var_Quant) as var_TotalICompra, (Var_Lucro * var_Quant) as var_TotalLucro, (var_ImpVenda * var_Quant) as var_TotalImpVenda, (VENDA * var_Quant) as var_TotalVenda, ( ( (VENDA * 100) / CUSTO_COMPRA) - 100 ) AS var_Diferenca  FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND " & _
    " (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.ATIVO = TRUE " & var_Criterio
    Set Rs2 = BD.OpenRecordset(SQL2)

    FormatarGrid_Produtos
    
ElseIf chkProduto.Value = Checked Then
    SQL2 = "SELECT (PRODUTOS.REF) as var_Ref,(PRODUTOS.FABRICANTE) as var_Fab,(PRODUTOS.TAMANHO) as var_Tam, (PRODUTOS.CATEGORIA) as var_linha, (PRODUTOS.CODIGO) as var_codEnt, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, " & _
    " (PRODUTOS.QUANT_ESTOQUE) as var_Quant, PRODUTOS_ENTRADA_ITENS.CUSTO AS var_Custo, PRODUTOS_ENTRADA_ITENS.FRETE AS var_Frete, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_COMPRA AS var_ImpCompra, " & _
    " PRODUTOS_ENTRADA_ITENS.CUSTO_COMPRA AS var_VlrCompra, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_VENDA AS var_ImpVenda, PRODUTOS_ENTRADA_ITENS.LUCRO_VALOR AS Var_Lucro, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA, " & _
    " (var_VlrCompra * var_Quant) as var_TotalCompra, (var_Custo * var_Quant) as var_TotalCusto, (var_Frete * var_Quant) as var_TotalFrete, (var_ImpCompra * var_Quant) as var_TotalICompra, (Var_Lucro * var_Quant) as var_TotalLucro, (var_ImpVenda * var_Quant) as var_TotalImpVenda, (VENDA * var_Quant) as var_TotalVenda   FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND " & _
    " (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.DESCRICAO = '" & cboConsProduto.Text & "' AND PRODUTOS.ATIVO = TRUE " & var_Criterio
    Set Rs2 = BD.OpenRecordset(SQL2)
 
    FormatarGrid_Produtos
    
ElseIf chkCodBarra.Value = Checked Then
    SQL2 = "SELECT (PRODUTOS.REF) as var_Ref,(PRODUTOS.FABRICANTE) as var_Fab,(PRODUTOS.TAMANHO) as var_Tam, (PRODUTOS.CATEGORIA) as var_linha, (PRODUTOS.CODIGO) as var_codEnt, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, " & _
    " (PRODUTOS.QUANT_ESTOQUE) as var_Quant, PRODUTOS_ENTRADA_ITENS.CUSTO AS var_Custo, PRODUTOS_ENTRADA_ITENS.FRETE AS var_Frete, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_COMPRA AS var_ImpCompra, " & _
    " PRODUTOS_ENTRADA_ITENS.CUSTO_COMPRA AS var_VlrCompra, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_VENDA AS var_ImpVenda, PRODUTOS_ENTRADA_ITENS.LUCRO_VALOR AS Var_Lucro, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA, " & _
    " (var_VlrCompra * var_Quant) as var_TotalCompra, (var_Custo * var_Quant) as var_TotalCusto, (var_Frete * var_Quant) as var_TotalFrete, (var_ImpCompra * var_Quant) as var_TotalICompra, (Var_Lucro * var_Quant) as var_TotalLucro, (var_ImpVenda * var_Quant) as var_TotalImpVenda, (VENDA * var_Quant) as var_TotalVenda   FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND " & _
    " (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.COD_BARRA = '" & txtConsCodBarra.Text & "' AND PRODUTOS.ATIVO = TRUE " & var_Criterio
    Set Rs2 = BD.OpenRecordset(SQL2)
    
    FormatarGrid_Produtos
ElseIf chkFab.Value = Checked Then
    SQL2 = "SELECT (PRODUTOS.REF) as var_Ref,(PRODUTOS.FABRICANTE) as var_Fab,(PRODUTOS.TAMANHO) as var_Tam, (PRODUTOS.CATEGORIA) as var_linha, (PRODUTOS.CODIGO) as var_codEnt, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, " & _
    " (PRODUTOS.QUANT_ESTOQUE) as var_Quant, PRODUTOS_ENTRADA_ITENS.CUSTO AS var_Custo, PRODUTOS_ENTRADA_ITENS.FRETE AS var_Frete, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_COMPRA AS var_ImpCompra, " & _
    " PRODUTOS_ENTRADA_ITENS.CUSTO_COMPRA AS var_VlrCompra, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_VENDA AS var_ImpVenda, PRODUTOS_ENTRADA_ITENS.LUCRO_VALOR AS Var_Lucro, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA, " & _
    " (var_VlrCompra * var_Quant) as var_TotalCompra, (var_Custo * var_Quant) as var_TotalCusto, (var_Frete * var_Quant) as var_TotalFrete, (var_ImpCompra * var_Quant) as var_TotalICompra, (Var_Lucro * var_Quant) as var_TotalLucro, (var_ImpVenda * var_Quant) as var_TotalImpVenda, (VENDA * var_Quant) as var_TotalVenda   FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND " & _
    " (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.FABRICANTE = '" & cboConsFab.Text & "' AND PRODUTOS.ATIVO = TRUE " & var_Criterio
    Set Rs2 = BD.OpenRecordset(SQL2)
 
    FormatarGrid_Produtos
ElseIf chkRef.Value = Checked Then
    SQL2 = "SELECT (PRODUTOS.REF) as var_Ref,(PRODUTOS.FABRICANTE) as var_Fab,(PRODUTOS.TAMANHO) as var_Tam, (PRODUTOS.CATEGORIA) as var_linha, (PRODUTOS.CODIGO) as var_codEnt, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, " & _
    " (PRODUTOS.QUANT_ESTOQUE) as var_Quant, PRODUTOS_ENTRADA_ITENS.CUSTO AS var_Custo, PRODUTOS_ENTRADA_ITENS.FRETE AS var_Frete, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_COMPRA AS var_ImpCompra, " & _
    " PRODUTOS_ENTRADA_ITENS.CUSTO_COMPRA AS var_VlrCompra, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_VENDA AS var_ImpVenda, PRODUTOS_ENTRADA_ITENS.LUCRO_VALOR AS Var_Lucro, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA, " & _
    " (var_VlrCompra * var_Quant) as var_TotalCompra, (var_Custo * var_Quant) as var_TotalCusto, (var_Frete * var_Quant) as var_TotalFrete, (var_ImpCompra * var_Quant) as var_TotalICompra, (Var_Lucro * var_Quant) as var_TotalLucro, (var_ImpVenda * var_Quant) as var_TotalImpVenda, (VENDA * var_Quant) as var_TotalVenda   FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND " & _
    " (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.REF = '" & cboConsRef.Text & "' AND PRODUTOS.ATIVO = TRUE " & var_Criterio
    Set Rs2 = BD.OpenRecordset(SQL2)
 
    FormatarGrid_Produtos
ElseIf chkTam.Value = Checked Then
    SQL2 = "SELECT (PRODUTOS.REF) as var_Ref,(PRODUTOS.FABRICANTE) as var_Fab,(PRODUTOS.TAMANHO) as var_Tam, (PRODUTOS.CATEGORIA) as var_linha, (PRODUTOS.CODIGO) as var_codEnt, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, " & _
    " (PRODUTOS.QUANT_ESTOQUE) as var_Quant, PRODUTOS_ENTRADA_ITENS.CUSTO AS var_Custo, PRODUTOS_ENTRADA_ITENS.FRETE AS var_Frete, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_COMPRA AS var_ImpCompra, " & _
    " PRODUTOS_ENTRADA_ITENS.CUSTO_COMPRA AS var_VlrCompra, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_VENDA AS var_ImpVenda, PRODUTOS_ENTRADA_ITENS.LUCRO_VALOR AS Var_Lucro, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA, " & _
    " (var_VlrCompra * var_Quant) as var_TotalCompra, (var_Custo * var_Quant) as var_TotalCusto, (var_Frete * var_Quant) as var_TotalFrete, (var_ImpCompra * var_Quant) as var_TotalICompra, (Var_Lucro * var_Quant) as var_TotalLucro, (var_ImpVenda * var_Quant) as var_TotalImpVenda, (VENDA * var_Quant) as var_TotalVenda   FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND " & _
    " (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.TAMANHO = '" & cboConsTam.Text & "' AND PRODUTOS.ATIVO = TRUE " & var_Criterio
    Set Rs2 = BD.OpenRecordset(SQL2)
 
    FormatarGrid_Produtos
ElseIf chkLinha.Value = Checked Then
    SQL2 = "SELECT (PRODUTOS.REF) as var_Ref,(PRODUTOS.FABRICANTE) as var_Fab,(PRODUTOS.TAMANHO) as var_Tam, (PRODUTOS.CATEGORIA) as var_linha, (PRODUTOS.CODIGO) as var_codEnt, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, " & _
    " (PRODUTOS.QUANT_ESTOQUE) as var_Quant, PRODUTOS_ENTRADA_ITENS.CUSTO AS var_Custo, PRODUTOS_ENTRADA_ITENS.FRETE AS var_Frete, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_COMPRA AS var_ImpCompra, " & _
    " PRODUTOS_ENTRADA_ITENS.CUSTO_COMPRA AS var_VlrCompra, PRODUTOS_ENTRADA_ITENS.IMPOSTO_VALOR_VENDA AS var_ImpVenda, PRODUTOS_ENTRADA_ITENS.LUCRO_VALOR AS Var_Lucro, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA, " & _
    " (var_VlrCompra * var_Quant) as var_TotalCompra, (var_Custo * var_Quant) as var_TotalCusto, (var_Frete * var_Quant) as var_TotalFrete, (var_ImpCompra * var_Quant) as var_TotalICompra, (Var_Lucro * var_Quant) as var_TotalLucro, (var_ImpVenda * var_Quant) as var_TotalImpVenda, (VENDA * var_Quant) as var_TotalVenda   FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND " & _
    " (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.CATEGORIA = '" & cboConsLinha.Text & "' AND PRODUTOS.ATIVO = TRUE " & var_Criterio
    Set Rs2 = BD.OpenRecordset(SQL2)
 
    FormatarGrid_Produtos
End If
If chkTodos.Value = False Then
cboConsProduto.SelStart = 0
cboConsProduto.SelLength = Len(cboConsProduto)
End If
End Sub





Private Sub Command1_Click()
execSQL "UPDATE PRODUTOS_ENTRADA_ITENS SET CODIGO_ENTRADA = 1 WHERE CODIGO_ENTRADA is NULL"
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
End Sub
Private Sub Mostrar_Historico()
    Call Abrir_BancodeDados
    If txtCodigo.Text = "" Then
        SQL = "SELECT PRODUTOS_ENTRADA.*, PRODUTOS_ENTRADA_ITENS.*, (PRODUTOS_ENTRADA.CODIGO) AS var_CODIGO FROM PRODUTOS_ENTRADA INNER JOIN PRODUTOS_ENTRADA_ITENS ON PRODUTOS_ENTRADA.CODIGO = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA WHERE FALSE ORDER BY PRODUTOS_ENTRADA.DATA_ENTRADA, PRODUTOS_ENTRADA.HORA_ENTRADA"
    Else
        SQL = "SELECT PRODUTOS_ENTRADA.*, PRODUTOS_ENTRADA_ITENS.*, (PRODUTOS_ENTRADA.CODIGO) AS var_CODIGO FROM PRODUTOS_ENTRADA INNER JOIN PRODUTOS_ENTRADA_ITENS ON PRODUTOS_ENTRADA.CODIGO = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA WHERE (CODIGO_PRODUTO = " & txtCodigo.Text & ") ORDER BY PRODUTOS_ENTRADA.DATA_ENTRADA, PRODUTOS_ENTRADA.HORA_ENTRADA"
    End If
    Set RS = BD.OpenRecordset(SQL)

    FormatarGrid_Historico
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
CmdHabilitar.Visible = True
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub


Private Sub Grid_Estoque_DblClick()
Me.Hide
Produtos_Entrada.Show
Produtos_Entrada.frmPrincipal.Enabled = True
Produtos_Entrada.frmSecundario.Enabled = True
Produtos_Entrada.cmdSalvar.Visible = False
Produtos_Entrada.cmdCancelar.Visible = False
Produtos_Entrada.cmdAlterar.Visible = True
Produtos_Entrada.cmdExcluir.Visible = True
Produtos_Entrada.cmdNovo.Enabled = True
Produtos_Entrada.frmPrincipal.Enabled = False
Produtos_Entrada.frmSecundario.Enabled = False
Produtos_Entrada.cmdAdicionar.Enabled = False
Produtos_Entrada.cmdRemover.Enabled = False
Produtos_Entrada.txtCodigo.Text = ""
Produtos_Entrada.txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub


Private Sub chkLinha_Click()
chkTodos.Value = Unchecked
chkProduto.Value = Unchecked
chkCodBarra.Value = Unchecked
'chkFab.Value = Unchecked
'chkRef.Value = Unchecked
'chkTam.Value = Unchecked
'chkLinha.Value = Unchecked
Desativa_Objetos
If chkLinha.Value = Checked Then cboConsLinha.SetFocus
End Sub

Private Sub chkCodBarra_Click()
chkTodos.Value = Unchecked
chkProduto.Value = Unchecked
'chkCodBarra.Value = Unchecked
chkFab.Value = Unchecked
chkRef.Value = Unchecked
chkTam.Value = Unchecked
chkLinha.Value = Unchecked
Desativa_Objetos
If chkCodBarra.Value = Checked Then txtConsCodBarra.SetFocus
End Sub

Private Sub chkProduto_Click()
chkTodos.Value = Unchecked
'chkProduto.Value = Unchecked
chkCodBarra.Value = Unchecked
chkFab.Value = Unchecked
chkRef.Value = Unchecked
chkTam.Value = Unchecked
chkLinha.Value = Unchecked
Desativa_Objetos
If chkProduto.Value = Checked Then cboConsProduto.SetFocus
End Sub

Private Sub chkTodos_Click()
'chkTodos.Value = Checked
chkProduto.Value = Unchecked
chkCodBarra.Value = Unchecked
chkFab.Value = Unchecked
chkRef.Value = Unchecked
chkTam.Value = Unchecked
chkLinha.Value = Unchecked
Desativa_Objetos
End Sub

Private Sub txtCodBarra_GotFocus()
txtCodBarra.SelStart = 0
txtCodBarra.SelLength = Len(txtCodBarra)
End Sub


Private Sub txtCodBarra_LostFocus()
    If txtCodBarra.Text = "" Then Exit Sub
    txtCodBarra.Text = Trim(txtCodBarra.Text)
    
    Call Abrir_BancodeDados
    SQL = "Select * From PRODUTOS where COD_BARRA = '" & txtCodBarra.Text & "'"
    Set RS = BD.OpenRecordset(SQL)
    
If cmdAlterar.Visible = False Then
    If RS.RecordCount > 0 Then
        MsgBox "Já existe um produto cadastrado com esse cód. de barra!", vbInformation, "Aviso do Sistema"
        txtCodBarra.SetFocus
    End If
End If
End Sub


Private Sub txtCodigo_Change()
    
    
If cmdSalvar.Visible = False Then
    If txtCodigo.Text = "" Then Exit Sub

    Call Abrir_BancodeDados
    SQL2 = "Select * from PRODUTOS where (CODIGO = " & txtCodigo.Text & ")"
    Set Rs2 = BD.OpenRecordset(SQL2)
    
    If Rs2.EOF Then
        Exit Sub
    End If
    
    LimparObjetos_Produtos
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    cmdAlterar.Visible = True
    cmdExcluir.Visible = True
    frmCadastro.Enabled = True
    MostrarDados_Produto
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



Private Sub txtQuant_GotFocus()
txtQuant.SelStart = 0
txtQuant.SelLength = Len(txtQuant)
End Sub


Private Sub txtValorAtual_GotFocus()
txtValorAtual.SelStart = 0
txtValorAtual.SelLength = Len(txtValorAtual)
End Sub


Private Sub txtValorAtual_LostFocus()
txtValorAtual.Text = Format(txtValorAtual.Text, "##,##0.00")
End Sub


