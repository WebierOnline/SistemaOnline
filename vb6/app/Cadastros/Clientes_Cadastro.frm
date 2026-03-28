VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Clientes_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CLIENTES"
   ClientHeight    =   8970
   ClientLeft      =   -870
   ClientTop       =   435
   ClientWidth     =   13020
   Icon            =   "Clientes_Cadastro.frx":0000
   LinkTopic       =   "Form49"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   106
      Top             =   8460
      Width           =   195
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   12825
      TabIndex        =   69
      Top             =   60
      Width           =   12855
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   180
         Width           =   1155
      End
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   8940
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTES"
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
         Left            =   1485
         TabIndex        =   71
         Top             =   240
         Width           =   1530
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "Clientes_Cadastro.frx":23D2
         Top             =   0
         Width           =   960
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   60
      TabIndex        =   37
      Top             =   1080
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   12938
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "Clientes_Cadastro.frx":372D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(4)=   "cmdExcluir"
      Tab(0).Control(5)=   "cmdSalvar"
      Tab(0).Control(6)=   "cmdNovo"
      Tab(0).Control(7)=   "cmdSair"
      Tab(0).Control(8)=   "cmdConsultarCNPJ"
      Tab(0).Control(9)=   "cmdConsultarIE"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "HISTÓRICO"
      TabPicture(1)   =   "Clientes_Cadastro.frx":3749
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid_Historico"
      Tab(1).Control(1)=   "cmdExibirPedidos"
      Tab(1).Control(2)=   "cmdExibirParcelas"
      Tab(1).Control(3)=   "cmdImprimirHistorico"
      Tab(1).Control(4)=   "cmdHistoricoFinanceiro"
      Tab(1).Control(5)=   "lblQuantHistorico"
      Tab(1).Control(6)=   "lblTotalHistorico"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "CONSULTA"
      TabPicture(2)   =   "Clientes_Cadastro.frx":3765
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label25"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label26"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblQuant"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdImprimir"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdExibir"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame4"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Grid_Consulta"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame5"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.Frame Frame5 
         Caption         =   "Tipo"
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
         Height          =   735
         Left            =   120
         TabIndex        =   76
         Top             =   4920
         Width           =   1455
         Begin VB.OptionButton optAtivos 
            Caption         =   "Ativos"
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
            TabIndex        =   78
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optInativos 
            Caption         =   "Inativos"
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
            TabIndex        =   77
            Top             =   480
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta 
         Height          =   4215
         Left            =   120
         TabIndex        =   73
         Top             =   420
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   7435
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
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
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   120
         TabIndex        =   65
         Top             =   5700
         Width           =   1455
         Begin VB.OptionButton optCidade 
            Caption         =   "Cidade"
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
            TabIndex        =   68
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optCPF 
            Caption         =   "CPF"
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
            TabIndex        =   67
            Top             =   420
            Width           =   1095
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
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
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
         Left            =   1620
         TabIndex        =   61
         Top             =   4920
         Width           =   11115
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   110
            Top             =   780
            Width           =   1755
            Begin VB.OptionButton chkCNPJ 
               Caption         =   "CNPJ:"
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
               Left            =   780
               TabIndex        =   112
               Top             =   180
               Width           =   855
            End
            Begin VB.OptionButton chkCPF 
               Caption         =   "CPF:"
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
               TabIndex        =   111
               Top             =   180
               Width           =   735
            End
         End
         Begin VB.CheckBox chkDiferente 
            Caption         =   "<>"
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
            Left            =   10440
            TabIndex        =   90
            Top             =   240
            Width           =   555
         End
         Begin VB.ComboBox cboNome 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   480
            Width           =   8115
         End
         Begin VB.CheckBox chkNome 
            Caption         =   "Nome:"
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
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkCidade 
            Caption         =   "Cidade:"
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
            Left            =   8280
            TabIndex        =   33
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboConsCidade 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8280
            TabIndex        =   34
            Top             =   480
            Width           =   2655
         End
         Begin MSMask.MaskEdBox mskConsCPF 
            Height          =   315
            Left            =   120
            TabIndex        =   89
            Top             =   1140
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   5595
         Left            =   -74880
         TabIndex        =   43
         Top             =   1620
         Width           =   10335
         Begin VB.TextBox txtComplemento 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   7320
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1260
            Width           =   675
         End
         Begin VB.ComboBox cboTipoCliente 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Clientes_Cadastro.frx":3781
            Left            =   120
            List            =   "Clientes_Cadastro.frx":3783
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   600
            Width           =   1875
         End
         Begin VB.TextBox txtObs 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   975
            Left            =   120
            MaxLength       =   100
            TabIndex        =   93
            Top             =   4500
            Width           =   10095
         End
         Begin VB.ComboBox cboSexo 
            Height          =   315
            ItemData        =   "Clientes_Cadastro.frx":3785
            Left            =   4500
            List            =   "Clientes_Cadastro.frx":3787
            TabIndex        =   20
            Top             =   2580
            Width           =   1515
         End
         Begin VB.TextBox txtIE 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   9060
            TabIndex        =   23
            Top             =   2580
            Width           =   1155
         End
         Begin VB.TextBox txtNum 
            BackColor       =   &H00C0FFFF&
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   6600
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1260
            Width           =   675
         End
         Begin VB.TextBox txtCodCid 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   2100
            TabIndex        =   83
            Top             =   2280
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtCodUF 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   420
            TabIndex        =   82
            Top             =   2220
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtCodigoIBGE 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   9000
            MaxLength       =   7
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   3900
            Width           =   1215
         End
         Begin VB.TextBox txtConjuge 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   2460
            MaxLength       =   100
            TabIndex        =   29
            Top             =   3900
            Width           =   6495
         End
         Begin VB.ComboBox cboCadBairro 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   8040
            TabIndex        =   12
            Top             =   1260
            Width           =   2175
         End
         Begin VB.ComboBox cboEstado 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Clientes_Cadastro.frx":3789
            Left            =   120
            List            =   "Clientes_Cadastro.frx":378B
            TabIndex        =   17
            Top             =   2580
            Width           =   615
         End
         Begin VB.ComboBox cboCidade 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Clientes_Cadastro.frx":378D
            Left            =   780
            List            =   "Clientes_Cadastro.frx":378F
            TabIndex        =   18
            Top             =   2580
            Width           =   2535
         End
         Begin VB.TextBox txtCI 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7800
            TabIndex        =   22
            Top             =   2580
            Width           =   1215
         End
         Begin VB.TextBox txtNome 
            BackColor       =   &H00C0FFFF&
            DataField       =   "CPF"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   2040
            MaxLength       =   60
            TabIndex        =   8
            Top             =   600
            Width           =   8145
         End
         Begin VB.TextBox txtEndereco 
            BackColor       =   &H00C0FFFF&
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1260
            Width           =   6435
         End
         Begin VB.TextBox txtReferencia 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1920
            Width           =   2955
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   6120
            MaxLength       =   30
            TabIndex        =   16
            Top             =   1920
            Width           =   4095
         End
         Begin VB.TextBox txtIdade 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   3240
            Width           =   435
         End
         Begin VB.ComboBox cboEstadoCivil 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   3900
            Width           =   2295
         End
         Begin VB.TextBox txtFiliacao 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   26
            Top             =   3240
            Width           =   5355
         End
         Begin MSMask.MaskEdBox mskCEP 
            Height          =   315
            Left            =   3360
            TabIndex        =   19
            Top             =   2580
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskNascimento 
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   3240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCPF 
            Height          =   315
            Left            =   6060
            TabIndex        =   21
            Top             =   2580
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCelular 
            Height          =   315
            Left            =   4620
            TabIndex        =   15
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtProfissao 
            Height          =   315
            Left            =   7320
            TabIndex        =   27
            Top             =   3240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskTelefone1 
            Height          =   315
            Left            =   3120
            TabIndex        =   14
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   7320
            TabIndex        =   109
            Top             =   1020
            Width           =   480
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "[COPIAR]"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6060
            TabIndex        =   104
            Top             =   2880
            Width           =   435
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cadastro"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Obs:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   4260
            Width           =   330
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Sexo:"
            Height          =   195
            Left            =   4500
            TabIndex        =   92
            Top             =   2340
            Width           =   405
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Insc. Estadual"
            Height          =   195
            Left            =   9060
            TabIndex        =   91
            Top             =   2340
            Width           =   1005
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Num.:"
            Height          =   195
            Left            =   6600
            TabIndex        =   84
            Top             =   1020
            Width           =   420
         End
         Begin VB.Label lblCódigoIBGE 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código IBGE"
            Height          =   195
            Left            =   9000
            TabIndex        =   81
            Top             =   3660
            Width           =   915
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Conjuge:"
            Height          =   195
            Left            =   2460
            TabIndex        =   80
            Top             =   3660
            Width           =   630
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Idade"
            Height          =   195
            Left            =   1440
            TabIndex        =   60
            Top             =   3000
            Width           =   405
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "RG"
            Height          =   195
            Left            =   7800
            TabIndex        =   59
            Top             =   2340
            Width           =   240
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Razão Social / Nome:"
            Height          =   195
            Left            =   2040
            TabIndex        =   58
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Ponto de Referência:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Correio Eletrônico:"
            Height          =   195
            Left            =   6120
            TabIndex        =   55
            Top             =   1680
            Width           =   1290
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   840
            TabIndex        =   54
            Top             =   2340
            Width           =   540
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   2340
            Width           =   255
         End
         Begin VB.Label lblCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF:"
            Height          =   195
            Left            =   6060
            TabIndex        =   52
            Top             =   2340
            Width           =   915
         End
         Begin VB.Label Label21 
            Caption         =   "Data de Nasc."
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   8040
            TabIndex        =   50
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Left            =   4620
            TabIndex        =   49
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Left            =   3360
            TabIndex        =   48
            Top             =   2340
            Width           =   330
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   3660
            Width           =   870
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Left            =   7320
            TabIndex        =   46
            Top             =   3000
            Width           =   690
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Filiação:"
            Height          =   195
            Left            =   1920
            TabIndex        =   45
            Top             =   3000
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fixo:"
            Height          =   195
            Left            =   3120
            TabIndex        =   44
            Top             =   1680
            Width           =   330
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Informações Extras"
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   10335
         Begin VB.ComboBox cboTipoPessoa 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2820
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox cboStatus 
            Height          =   315
            Left            =   60
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   600
            Width           =   1515
         End
         Begin VB.TextBox txtCadastro 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   600
            Width           =   1155
         End
         Begin VB.ComboBox cboTipo 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   4560
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtLimite 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7110
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtUltimaCompra 
            Height          =   315
            Left            =   8490
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Pessoa:"
            Height          =   195
            Left            =   2820
            TabIndex        =   108
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   60
            TabIndex        =   72
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastro:"
            Height          =   195
            Left            =   1620
            TabIndex        =   42
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Contribuinte:"
            Height          =   195
            Left            =   4560
            TabIndex        =   41
            Top             =   360
            Width           =   1470
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite de Crédito:"
            Height          =   195
            Left            =   7110
            TabIndex        =   40
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última Compra:"
            Height          =   195
            Left            =   8490
            TabIndex        =   39
            Top             =   360
            Width           =   1065
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   555
         Left            =   9240
         TabIndex        =   35
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
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
         MICON           =   "Clientes_Cadastro.frx":3791
         PICN            =   "Clientes_Cadastro.frx":37AD
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
         Left            =   10980
         TabIndex        =   36
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
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
         MICON           =   "Clientes_Cadastro.frx":553F
         PICN            =   "Clientes_Cadastro.frx":555B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Historico 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   74
         Top             =   420
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   10610
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   -64440
         TabIndex        =   85
         Top             =   1860
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
         MICON           =   "Clientes_Cadastro.frx":72ED
         PICN            =   "Clientes_Cadastro.frx":7309
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
         Left            =   -64440
         TabIndex        =   86
         Top             =   2520
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
         MICON           =   "Clientes_Cadastro.frx":909B
         PICN            =   "Clientes_Cadastro.frx":90B7
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
         Left            =   -64440
         TabIndex        =   87
         Top             =   3180
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
         MICON           =   "Clientes_Cadastro.frx":AE49
         PICN            =   "Clientes_Cadastro.frx":AE65
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
         Left            =   -64440
         TabIndex        =   88
         Top             =   1200
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
         MICON           =   "Clientes_Cadastro.frx":CBF7
         PICN            =   "Clientes_Cadastro.frx":CC13
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
         Left            =   -64440
         TabIndex        =   0
         Top             =   540
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
         MICON           =   "Clientes_Cadastro.frx":E9A5
         PICN            =   "Clientes_Cadastro.frx":E9C1
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
         Left            =   -64440
         TabIndex        =   95
         Top             =   3840
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
         MICON           =   "Clientes_Cadastro.frx":10753
         PICN            =   "Clientes_Cadastro.frx":1076F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
         Height          =   435
         Left            =   -74880
         TabIndex        =   96
         Top             =   6600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "EXIBIR PRODUTOS"
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
         MICON           =   "Clientes_Cadastro.frx":12501
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExibirParcelas 
         Height          =   435
         Left            =   -72420
         TabIndex        =   97
         Top             =   6600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "EXIBIR PARCELAS"
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
         MICON           =   "Clientes_Cadastro.frx":1251D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimirHistorico 
         Height          =   435
         Left            =   -69960
         TabIndex        =   99
         Top             =   6600
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "IMPRIMIR HISTÓRICO"
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
         MICON           =   "Clientes_Cadastro.frx":12539
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdHistoricoFinanceiro 
         Height          =   435
         Left            =   -67020
         TabIndex        =   101
         Top             =   6600
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "IMPRIMIR HISTÓRICO FINANCEIRO"
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
         MICON           =   "Clientes_Cadastro.frx":12555
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsultarCNPJ 
         Height          =   615
         Left            =   -64440
         TabIndex        =   102
         Top             =   4980
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Consultar Cadastro"
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
         MICON           =   "Clientes_Cadastro.frx":12571
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsultarIE 
         Height          =   615
         Left            =   -64440
         TabIndex        =   103
         Top             =   5640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Consultar IE"
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
         MICON           =   "Clientes_Cadastro.frx":1258D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblQuantHistorico 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
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
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   -62520
         TabIndex        =   100
         Top             =   6660
         Width           =   225
      End
      Begin VB.Label lblTotalHistorico 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
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
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   -62520
         TabIndex        =   75
         Top             =   6960
         Width           =   225
      End
      Begin VB.Label lblQuant 
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
         Left            =   11760
         TabIndex        =   64
         Top             =   4680
         Width           =   225
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registros"
         Height          =   195
         Left            =   12060
         TabIndex        =   63
         Top             =   4680
         Width           =   660
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Dê um duplo-clique para ver mais informações"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   4680
         Width           =   3255
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   79
      Top             =   8700
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18627
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "07:29"
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
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "Preenchimento obrigatório"
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
      Left            =   300
      TabIndex        =   107
      Top             =   8460
      Width           =   2220
   End
End
Attribute VB_Name = "Clientes_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim printSQL As String
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Private Sub AutoNumeracao()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_cliente FROM cliente;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtCodigo.Text = r("cod_cliente") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Campos_Brancos()
If cmdAlterar.Enabled = False Then txtCodigo.Text = ""
txtNome.Text = ""
cboStatus.Text = ""
txtEndereco.Text = ""
txtNum.Text = ""
txtReferencia.Text = ""
mskTelefone1.Mask = ""
mskTelefone1.Text = ""
txtComplemento.Text = ""
cboCidade.Clear
cboCidade.Text = ""
cboEstado.Clear
cboEstado.Text = ""
cboSexo.Text = ""
mskCPF.Mask = ""
mskCPF.Text = ""
txtEmail.Text = ""
mskNascimento.Mask = ""
mskNascimento.Text = ""
txtIdade.Text = ""
txtCI.Text = ""
txtIE.Text = ""
txtCadastro.Text = ""
cboTipo.Text = ""
txtLimite.Text = ""
txtUltimaCompra.Text = ""
mskCelular.Mask = ""
mskCelular.Text = ""
cboCadBairro.Text = ""
cboEstadoCivil.Text = ""
txtProfissao.Text = ""
txtFiliacao.Text = ""
mskCEP.Mask = ""
mskCEP.Text = ""
txtConjuge.Text = ""
txtCodigoIBGE.Text = ""
txtObs.Text = ""
End Sub

Private Sub Limite_Cliente()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim Limite As Currency
   
   'Inivializa as variáveis
   Limite = 0
   
   'Monta a consulta
   sSQL = "SELECT pedidos.cod_cliente, ISNULL(SUM(parcelas.valor_final), 0) AS limite " & _
      "FROM pedidos INNER JOIN parcelas ON pedidos.cod_pedido = parcelas.cod_pedido " & _
      "WHERE (parcelas.status = 0) AND (pedidos.cod_cliente = " & txtCodigo.Text & ") " & _
      "GROUP BY pedidos.cod_cliente;"
   
   Set r = dbData.OpenRecordset(sSQL)              'Abre a tabela
   If Not r.BOF Then Limite = CCur(r("limite"))    'Recupera o limtie de crédito se houver
   If r.State <> 0 Then r.Close                    'Fecha a tabela
   Set r = Nothing
   
   If CCur(txtLimite.Text) <> 0 Then
      txtLimite.Text = Format(CCur(txtLimite.Text) - Limite, ocMONEY)
   End If
End Sub

Private Sub LimparGrid_Historico()
   'Dim sSQL As String
   'Dim r As ADODB.Recordset
   
   'sSQL = "SELECT cliente.*, pedidos.* FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE false"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Não é necessário consultar nada para realizar a limpeza do grid
   'Basta passar como paramento Nothing
   
   FormatarGrid_Historico Nothing
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
Frame1.Enabled = True
Frame2.Enabled = True

If Not rTabela Is Nothing Then
   txtCodigo.Text = ValidateNull(rTabela("codigo"))
   cboStatus.Text = IIf(rTabela("status") = True, "ATIVO", "INATIVO")
   txtNome.Text = ValidateNull(rTabela("nome"))
   txtEndereco.Text = ValidateNull(rTabela("endereco"))
   txtNum.Text = ValidateNull(rTabela("numero"))
   txtReferencia.Text = ValidateNull(rTabela("ponto_de_referencia"))
   mskTelefone1.Text = ValidateNull(rTabela("telefone1"))
   txtComplemento.Text = ValidateNull(rTabela("Complemento"))
   cboCidade.Text = ValidateNull(rTabela("cidade"))
   cboEstado.Text = ValidateNull(rTabela("estado"))
   cboSexo.Text = ValidateNull(rTabela("sexo"))
   mskCPF.Text = ValidateNull(rTabela("cpf"))
   txtCI.Text = ValidateNull(rTabela("rg"))
   txtIE.Text = ValidateNull(rTabela("ie"))
   txtEmail.Text = ValidateNull(rTabela("correio_eletronico"))
   If Not IsNull(rTabela("data_de_nascimento")) Then mskNascimento.Text = Format$(ValidateNull(rTabela("data_de_nascimento")), ocDATA)
   txtCadastro.Text = Format$(ValidateNull(rTabela("data_cadastro")), ocDATA)
   cboTipoPessoa.Text = ValidateNull(rTabela("tipo"))
   txtLimite.Text = FormatNumber(rTabela("limite_credito"), 2)
   txtUltimaCompra.Text = IIf(IsNull(rTabela("ultima_compra")), "", Format$(rTabela("ultima_compra"), ocDATA))
   mskCelular.Text = ValidateNull(rTabela("celular"))
   cboCadBairro.Text = ValidateNull(rTabela("bairro"))
   mskCEP.Text = ValidateNull(rTabela("cep"))
   cboEstadoCivil.Text = ValidateNull(rTabela("estadocivil"))
   txtProfissao.Text = ValidateNull(rTabela("profissao"))
   txtFiliacao.Text = ValidateNull(rTabela("filiacao"))
   txtConjuge.Text = ValidateNull(rTabela("conjuge"))
   txtCodigoIBGE.Text = ValidateNull(rTabela("CodigoIBGE"))
   txtObs.Text = ValidateNull(rTabela("obs"))
   
    If rTabela("TipoContribuinte") = 1 Then
        cboTipo.Text = "1 - CONTRIBUINTE ICMS"
    ElseIf rTabela("TipoContribuinte") = 2 Then
        cboTipo.Text = "2 - CONTRIBUINTE ISENTO"
    ElseIf rTabela("TipoContribuinte") = 9 Then
        cboTipo.Text = "9 - NÃO CONTRIBUINTE"
    End If
End If
End Sub

Private Sub Mostrar_Historico()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

sSQL = "SELECT cliente.CODIGO, pedidos.COD_PEDIDO, pedidos.COD_CLIENTE, pedidos.TIPO_PAGAMENTO, pedidos.PAGAMENTO, pedidos.SUBTOTAL, pedidos.ValorDescReal, pedidos.ValorAcrescReal, pedidos.TOTAL, pedidos.TIPO_PEDIDO, cliente.Nome, pedidos.DATA_COMPRA " & _
       "FROM cliente INNER JOIN pedidos ON cliente.CODIGO = pedidos.COD_CLIENTE " & _
       "WHERE pedidos.COD_CLIENTE = " & txtCodigo.Text & ""
'Debug.Print sSQL
       
Set r = dbData.OpenRecordset(sSQL)

lblQuantHistorico.Caption = r.RecordCount

FormatarGrid_Historico r

printSQL = sSQL

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Preencher_TipoContribuinte()
'If cboTipo.ListCount = 0 Then 'Desnecessário a comparação bastar limpar a lista primeiro
cboTipo.Clear
cboTipo.AddItem "1 - CONTRIBUINTE ICMS"
cboTipo.AddItem "2 - CONTRIBUINTE ISENTO"
cboTipo.AddItem "9 - NÃO CONTRIBUINTE"
End Sub


Private Sub Preencher_TipoPessoa()
'If cboTipo.ListCount = 0 Then 'Desnecessário a comparação bastar limpar a lista primeiro
cboTipoPessoa.Clear
cboTipoPessoa.AddItem "FÍSICA"
cboTipoPessoa.AddItem "JURÍDICA"
cboTipoPessoa.AddItem "RURAL"
End Sub

Private Sub PreencherTipoCadastro()
cboTipoCliente.Clear
cboTipoCliente.AddItem "PRÉ-CADASTRO"
cboTipoCliente.AddItem "CADASTRO"
End Sub

Private Sub cboCadBairro_LostFocus()
cboCadBairro.Text = TirarEspaco(cboCadBairro.Text)
End Sub

Private Sub cboCidade_Click()
cboCidade_LostFocus
End Sub

Private Sub cboCidade_LostFocus()
On Error GoTo TrataErro

If cboCidade.Text = "" Then txtCodCid.Text = "": txtCodigoIBGE.Text = "": Exit Sub
If cboCidade.ListIndex = -1 Then txtCodCid.Text = "": Exit Sub

txtCodCid = cboCidade.ItemData(cboCidade.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboEstado_Click()
cboEstado_LostFocus
End Sub

Private Sub cboEstado_LostFocus()
On Error GoTo TrataErro

If cboEstado.Text = "" Then txtCodUF.Text = "": Exit Sub
If cboEstado.ListIndex = -1 Then txtCodUF.Text = "": Exit Sub

txtCodUF = cboEstado.ItemData(cboEstado.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboStatus_GotFocus()
cboStatus.Clear
cboStatus.AddItem "ATIVO"
cboStatus.AddItem "INATIVO"
moCombo.AttachTo cboStatus
End Sub

Private Sub cboCadBairro_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

'Limpa a lista atual
cboCadBairro.Clear

sSQL = "SELECT DISTINCT bairro FROM cliente ORDER BY bairro;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboCadBairro.AddItem ValidateNull(r("bairro"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboCadBairro
End Sub

Private Sub cboCidade_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboConsCidade_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT DISTINCT cidade FROM cliente ORDER BY cidade;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsCidade.AddItem ValidateNull(r("cidade"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboConsCidade
End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboEstadoCivil_GotFocus()
   cboEstadoCivil.Clear
   cboEstadoCivil.AddItem "CASADO"
   cboEstadoCivil.AddItem "SOLTEIRO"
   cboEstadoCivil.AddItem "VIÚVO"
   cboEstadoCivil.AddItem "DIVORCIADO"
   moCombo.AttachTo cboEstadoCivil
End Sub

Private Sub cboEstadoCivil_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboNome_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

'Limpa a lista atual
cboNome.Clear

sSQL = "SELECT DISTINCT nome FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboNome.AddItem r("nome")
   r.MoveNext
Loop

moCombo.AttachTo cboNome
End Sub

Private Sub cboSexo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboTipo_Change()
cboTipoCliente_Change
End Sub

Private Sub cboTipo_GotFocus()
moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboTipo_LostFocus()
If cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then
    If cboTipoPessoa.Text = "FÍSICA" Then
        lblCPF.Caption = "CPF"
    ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
        lblCPF.Caption = "CNPJ"
    End If
ElseIf cboTipo.Text = "2 - CONTRIBUINTE ISENTO" Then
    If cboTipoPessoa.Text = "FÍSICA" Then
        lblCPF.Caption = "CPF"
    ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
        lblCPF.Caption = "CNPJ"
    End If
ElseIf cboTipo.Text = "9 - NÃO CONTRIBUINTE" Then
    lblCPF.Caption = "CPF"
End If
End Sub

Private Sub cboTipo_Validate(Cancel As Boolean)
If cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then
    txtIE.BackColor = &HC0FFFF
Else
    txtIE.BackColor = &HFFFFFF
End If
End Sub

Private Sub cboTipoCliente_Change()
If cboTipoCliente.Text = "PRÉ-CADASTRO" Then
    txtNome.BackColor = &HC0FFFF
    txtEndereco.BackColor = &HFFFFFF
    txtNum.BackColor = &HFFFFFF
    cboCadBairro.BackColor = &HFFFFFF
    cboEstado.BackColor = &HFFFFFF
    cboCidade.BackColor = &HFFFFFF
    mskCEP.BackColor = &HFFFFFF
    mskCPF.BackColor = &HFFFFFF
    txtCodigoIBGE.BackColor = &HFFFFFF
    mskCelular.BackColor = &HC0FFFF
    txtIE.BackColor = &HFFFFFF
Else
    txtNome.BackColor = "&HC0FFFF"
    txtEndereco.BackColor = &HC0FFFF
    txtNum.BackColor = &HC0FFFF
    cboCadBairro.BackColor = &HC0FFFF
    cboEstado.BackColor = &HC0FFFF
    cboCidade.BackColor = &HC0FFFF
    mskCEP.BackColor = &HC0FFFF
    mskCPF.BackColor = &HC0FFFF
    txtCodigoIBGE.BackColor = &HC0FFFF
    mskCelular.BackColor = &HFFFFFF
    If cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then
        txtIE.BackColor = &HC0FFFF
    Else
        txtIE.BackColor = &HFFFFFF
    End If
End If
End Sub

Private Sub cboTipoCliente_Click()
cboTipoCliente_Change
End Sub


Private Sub cboTipoCliente_GotFocus()
moCombo.AttachTo cboTipoCliente
End Sub


Private Sub cboTipoCliente_LostFocus()
cboTipoCliente_Change
End Sub

Private Sub cboTipoPessoa_GotFocus()
moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipoPessoa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboTipoPessoa_LostFocus()
If cboTipoPessoa.Text = "FÍSICA" Then
    cboTipo.Text = "9 - NÃO CONTRIBUINTE"
ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
    cboTipo.Text = "1 - CONTRIBUINTE ICMS"
ElseIf cboTipoPessoa.Text = "RURAL" Then
    cboTipo.Text = "1 - CONTRIBUINTE ICMS"
End If

If cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then
    If cboTipoPessoa.Text = "RURAL" Then
        lblCPF.Caption = "CPF"
    ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
        lblCPF.Caption = "CNPJ"
    End If
ElseIf cboTipo.Text = "2 - CONTRIBUINTE ISENTO" Then
    If cboTipoPessoa.Text = "FÍSICA" Then
        lblCPF.Caption = "CPF"
    ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
        lblCPF.Caption = "CNPJ"
    End If
ElseIf cboTipo.Text = "9 - NÃO CONTRIBUINTE" Then
    lblCPF.Caption = "CPF"
End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkCidade_Click()
If chkCidade.Value = 1 Then
   cboConsCidade.Enabled = True
   cboConsCidade.SetFocus
   
    chkCPF.Value = Unchecked
    chkCNPJ.Value = Unchecked
    chkNome.Value = 0
Else
   cboConsCidade.Enabled = False
End If
End Sub

Private Sub chkCNPJ_Click()
If chkCNPJ.Value = True Then
   mskConsCPF.Enabled = True
   mskConsCPF.SetFocus
   
   chkNome.Value = 0
   chkCidade.Value = 0
   chkDiferente.Value = 0
Else
   mskConsCPF.Enabled = False
End If
End Sub

Private Sub chkCPF_Click()
If chkCPF.Value = True Then
   mskConsCPF.Enabled = True
   mskConsCPF.SetFocus
   
   chkNome.Value = 0
   chkCidade.Value = 0
   chkDiferente.Value = 0
Else
   mskConsCPF.Enabled = False
End If
End Sub

Private Sub chkNome_Click()
   If chkNome.Value = 1 Then
      cboNome.Enabled = True
      cboNome.Clear
      cboNome.SetFocus
      
      chkCPF.Value = Unchecked
      chkCNPJ.Value = Unchecked
      chkCidade.Value = 0
      chkDiferente.Value = 0
   Else
      cboNome.Enabled = False
   End If
End Sub

Private Sub cmdAlterar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

'Cliente CONSUMIDOR não permite alteração
If txtNome.Text = "CONSUMIDOR" Then Exit Sub

If cboTipoCliente.Text = "PRÉ-CADASTRO" Then
    If txtNome.Text = "" Then MsgBox "Campo NOME obrigatório!", vbCritical, "Online Commerce": txtNome.SetFocus: Exit Sub
    If mskCelular.Text = "" Then MsgBox "Campo CELULAR obrigatório!", vbCritical, "Online Commerce": mskCelular.SetFocus: Exit Sub
Else
    If txtNome.Text = "" Then MsgBox "Campo NOME obrigatório!", vbCritical, "Online Commerce": txtNome.SetFocus: Exit Sub
    If txtEndereco.Text = "" Then MsgBox "Campo ENDEREÇO obrigatório!", vbCritical, "Online Commerce": txtEndereco.SetFocus: Exit Sub
    If txtNum.Text = "" Then MsgBox "Campo NÚMERO obrigatório!", vbCritical, "Online Commerce": txtNum.SetFocus: Exit Sub
    If cboCadBairro.Text = "" Then MsgBox "Campo BAIRRO obrigatório!", vbCritical, "Online Commerce": cboCadBairro.SetFocus: Exit Sub
    If cboCidade.Text = "" Then MsgBox "Campo CIDADE obrigatório!", vbCritical, "Online Commerce": cboCidade.SetFocus: Exit Sub
    If cboEstado.Text = "" Then MsgBox "Campo ESTADO obrigatório!", vbCritical, "Online Commerce": cboEstado.SetFocus: Exit Sub
    If mskCEP.Text = "" Then MsgBox "Campo CEP obrigatório!", vbCritical, "Online Commerce": mskCEP.SetFocus: Exit Sub
    If mskCPF.Text = "" Then MsgBox "Campo CPF obrigatório!", vbCritical, "Online Commerce": mskCPF.SetFocus: Exit Sub
    If cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then
        If txtIE.Text = "" Then MsgBox "Campo INSCRIÇÃO ESTADUAL obrigatório!", vbCritical, "Online Commerce": txtIE.SetFocus: Exit Sub
    End If
    If txtCodigoIBGE.Text = "" Or txtCodigoIBGE.Text = "0" Or Len(txtCodigoIBGE.Text) <> 7 Then
        MsgBox "Campo CÓDIGO IBGE obrigatório!", vbCritical, "Online Commerce": txtCodigoIBGE.SetFocus: Exit Sub
    End If
End If

If cboTipoCliente.Text <> "PRÉ-CADASTRO" Then
    If cboTipoPessoa.Text = "FÍSICA" And cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!", vbCritical, "Online Commerce":  Exit Sub
    If cboTipoPessoa.Text = "JURÍDICA" And cboTipo.Text = "9 - NÃO CONTRIBUINTE" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!", vbCritical, "Online Commerce":  Exit Sub
    If cboTipoPessoa.Text = "RURAL" And cboTipo.Text = "9 - NÃO CONTRIBUINTE" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!", vbCritical, "Online Commerce":  Exit Sub
    If cboTipoPessoa.Text = "RURAL" And cboTipo.Text = "2 - CONTRIBUINTE ISENTO" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!", vbCritical, "Online Commerce":  Exit Sub
End If

'Não informou o código
If txtCodigo.Text = "" Then
   ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o cliente na guia CONSULTA.", vbInformation
   Exit Sub
End If

'Faz a atualização de forma direta e verifica se houve algum erro
If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
Campos_Brancos
Form_Load
End Sub

Private Function Inserir_Dados() As Boolean
Dim sSQL As String

'Validaçao dos dados
If Trim(txtLimite.Text) = "" Then txtLimite.Text = "0"
If Trim(txtCodigoIBGE.Text) = "" Then txtCodigoIBGE.Text = "0"

'Comando de inclusão
sSQL = "INSERT INTO cliente (" & _
   "codigo, status, nome, endereco, numero, ponto_de_referencia, bairro, cep, cidade, estado, " & _
   "telefone1, Complemento, celular, sexo, cpf, rg, ie, correio_eletronico, " & _
   "data_de_nascimento, data_cadastro, TipoContribuinte, limite_credito, tipo, " & _
   "estadocivil, profissao, filiacao, conjuge, CodigoIBGE, obs) VALUES ("
'
sSQL = sSQL & _
   txtCodigo.Text & ", " & IIf((cboStatus.Text = "ATIVO"), 1, 0) & ", '" & txtNome.Text & "', '" & txtEndereco.Text & "','" & IIf((txtNum.Text = ""), "0", txtNum.Text) & "', '" & _
   txtReferencia.Text & "', '" & cboCadBairro.Text & "', '" & mskCEP.Text & "', '" & cboCidade.Text & "', '" & _
   cboEstado.Text & "', '" & mskTelefone1.Text & "', '" & txtComplemento.Text & "', '" & _
   mskCelular.Text & "', '" & cboSexo.Text & "', '" & mskCPF.Text & "', '" & txtCI.Text & "', '" & txtIE.Text & "', '" & txtEmail.Text & "', " & _
   IIf((mskNascimento.Text = ""), "Null", "CONVERT(DATETIME, '" & Format$(mskNascimento.Text, ocDATA) & "', 103)") & ", " & _
   "CONVERT(DATETIME, '" & Format$(txtCadastro.Text, ocDATA) & "', 103), '" & IIf(IsNull(Format(Left(cboTipo.Text, 1), "@")) Or Vazio(Format(Left(cboTipo.Text, 1), "@")), 1, Format(Left(cboTipo.Text, 1), "@")) & "', " & Replace(CCur(txtLimite.Text), ",", ".") & ", '" & cboTipoPessoa.Text & "', '" & _
   cboEstadoCivil.Text & "', '" & txtProfissao.Text & "', '" & txtFiliacao.Text & "', '" & txtConjuge.Text & "', " & IIf((txtCodigoIBGE.Text = ""), "0", txtCodigoIBGE.Text) & ", '" & txtObs.Text & "')"
'" & IIf((txtCodigoIBGE.Text = ""), "0", "txtCodigoIBGE.Text") & "
Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
   Dim sSQL As String
   
   'Validaçao dos dados
   If Trim(txtLimite.Text) = "" Then txtLimite.Text = "0"
   If Trim(txtCodigoIBGE.Text) = "" Then txtCodigoIBGE.Text = "0"
   If Trim(txtNum.Text) = "" Then txtNum.Text = "0"
   
   'Comando de atualização
   sSQL = "UPDATE cliente SET " & _
      "status = " & IIf((cboStatus.Text = "ATIVO"), 1, 0) & ", " & _
      "nome = '" & txtNome.Text & "', " & _
      "endereco = '" & txtEndereco.Text & "', " & _
      "numero = '" & txtNum.Text & "', " & _
      "ponto_de_referencia = '" & txtReferencia.Text & "', " & _
      "bairro = '" & cboCadBairro.Text & "', " & _
      "cep = '" & mskCEP.Text & "', " & _
      "cidade = '" & cboCidade.Text & "', " & _
      "estado = '" & cboEstado.Text & "', " & _
      "telefone1 = '" & mskTelefone1.Text & "', " & _
      "Complemento = '" & txtComplemento.Text & "', " & _
      "celular = '" & mskCelular.Text & "', "

   sSQL = sSQL & _
      "sexo = '" & cboSexo.Text & "', " & _
      "cpf = '" & mskCPF.Text & "', " & _
      "rg = '" & txtCI.Text & "', " & _
      "ie = '" & txtIE.Text & "', " & _
      "correio_eletronico = '" & txtEmail.Text & "', " & _
      "data_de_nascimento = CONVERT(DATETIME, '" & Format$(mskNascimento.Text, ocDATA) & "', 103), " & _
      "TipoContribuinte = '" & IIf(IsNull(Format(Left(cboTipo.Text, 1), "@")) Or Vazio(Format(Left(cboTipo.Text, 1), "@")), 1, Format(Left(cboTipo.Text, 1), "@")) & "', " & _
      "limite_credito = " & Replace(CCur(txtLimite.Text), ",", ".") & ", " & _
      "tipo = '" & cboTipoPessoa.Text & "', " & _
      "estadocivil = '" & cboEstadoCivil.Text & "', " & _
      "profissao = '" & txtProfissao.Text & "', " & _
      "filiacao = '" & txtFiliacao.Text & "', " & _
      "conjuge = '" & txtConjuge.Text & "', " & _
      "obs = '" & txtObs.Text & "', " & _
      "CodigoIBGE = " & txtCodigoIBGE.Text
   'Debug.Print sSQL
   'Condição para atualização
   sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ")"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdCancelar_Click()
Campos_Brancos
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
Frame1.Enabled = False
Frame2.Enabled = False
cboStatus.Text = "ATIVO"
End Sub

Private Sub cmdConsultarCNPJ_Click()
ShellExecute hwnd, "open", "http://servicos.receita.fazenda.gov.br/Servicos/cnpjreva/Cnpjreva_Solicitacao.asp", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdConsultarIE_Click()
ShellExecute hwnd, "open", "https://dfe-portal.svrs.rs.gov.br/Nfe/Ccc", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdExcluir_Click()
Dim sSQL As String
Dim bRet As Boolean


If txtNome.Text = "CONSUMIDOR" Then Exit Sub

If txtCodigo.Text = "" Then
   ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o cliente na guia CONSULTA", vbInformation
   Exit Sub
End If

'Solicita ao usuário confirmação da exclusão
If ShowMsg("Excluir esse cliente?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

'sSQL = "SELECT COD_CLIENTE " & _
       "FROM pedidos " & _
       "WHERE (COD_CLIENTE = " & txtCodigo.Text & ");"

sSQL = "SELECT COD_CLIENTE FROM pedidos WHERE COD_CLIENTE = " & txtCodigo.Text & " " & _
        "Union " & _
        "SELECT CodigoCorrentista FROM NotaFiscal WHERE CodigoCorrentista = " & txtCodigo.Text & ""

Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    ShowMsg "Não é permitido excluir esse cliente!" & vbCrLf & "Esse cliente possui vendas efetuadas posteriores. Desative-o!", vbInformation
    Exit Sub
End If

'Faz a exclusão usando o comando DELETE do SQL
sSQL = "DELETE FROM cliente WHERE (codigo = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
Campos_Brancos
Form_Load
End Sub

Private Sub cmdExibir_Click()
   'If chkCodigo.Value = 1 And txtConsCodigo.Text = "" Then Exit Sub
   
   Dim RESULTADO As Long
   Dim INDICE As String       'INDICE PARA ORGANIZAR OS DADOS
   Dim SITUACAO As String
   
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Verifica qual o índice a ser utlizado
   If optNome.Value = True Then
      INDICE = "nome;"
   ElseIf optCPF.Value = True Then
      INDICE = "cpf;"
   ElseIf optCidade.Value = True Then
      INDICE = "cidade;"
   End If
   
   'Verifica a situacao do cliente
   If optAtivos.Value = True Then
      SITUACAO = "status = 1"
   ElseIf optInativos.Value = True Then
      SITUACAO = "status = 0"
   End If
   
   'Monta a consulta básica para não repetir várias linhas
   sSQL = "SELECT *, CASE status WHEN 0 THEN 'INATIVO' WHEN 1 THEN 'ATIVO' END AS var_status " & _
      "FROM cliente WHERE "
      
   If chkCPF.Value = True Or chkCNPJ.Value = True Then
      If mskConsCPF.Text = "" Then Exit Sub
      sSQL = sSQL & "(cpf = '" & mskConsCPF.Text & "') AND "
      
   ElseIf chkNome.Value = 1 Then
      If cboNome.Text = "" Then Exit Sub
      sSQL = sSQL & "(nome LIKE '%" & cboNome.Text & "%') AND "     'O sinal de % é o caracter curinga que no ACCESS era *
      
   ElseIf chkCidade.Value = 1 Then
      If cboConsCidade.Text = "" Then Exit Sub
      
      If chkDiferente.Value = 1 Then
        sSQL = sSQL & "(cidade <> '" & cboConsCidade.Text & "') AND "
      Else
        sSQL = sSQL & "(cidade = '" & cboConsCidade.Text & "') AND "
      End If
   
   End If
   
   'Finaliza a consulta com os critérios extras
   sSQL = sSQL & "(" & SITUACAO & ") ORDER BY " & INDICE
   
   'Abre os registros e conta quantos foram encontrados
   Set r = dbData.OpenRecordset(sSQL, RESULTADO)
   
   'Mostra os dados do grid
   FormatarGrid_Consulta r
   
   lblQuant.Caption = Format(RESULTADO, "00")
   
   printSQL = sSQL

   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdExibirParcelas_Click()
If Grid_Historico.Col = 0 Then Exit Sub
   If IsNumeric(Grid_Historico.TextMatrix(Grid_Historico.Row, 1)) = True Then
         Vendas_Consulta_Geral_Parcelas.loadInformacoes (Grid_Historico.TextMatrix(Grid_Historico.Row, 1))
         Vendas_Consulta_Geral_Parcelas.Show 1
   End If
End Sub

Private Sub cmdExibirPedidos_Click()
If Grid_Historico.Col = 0 Then Exit Sub
If IsNumeric(Grid_Historico.TextMatrix(Grid_Historico.Row, 1)) = True Then
   If Grid_Historico.Col = 1 Then
      If Grid_Historico.TextMatrix(Grid_Historico.Row, 1) = "" Then Exit Sub
      Parcelas_Consulta_Produtos.loadPedidos Grid_Historico.TextMatrix(Grid_Historico.Row, 1), Grid_Historico.TextMatrix(Grid_Historico.Row, 9)
      Parcelas_Consulta_Produtos.Show 1
   End If
End If
End Sub


Private Sub cmdHistoricoFinanceiro_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

sSQL = "SELECT parcelas.CODIGO, parcelas.COD_PEDIDO, parcelas.NUMERO, parcelas.DATA, parcelas.PAGAMENTO, parcelas.VALOR, parcelas.VALOR_FINAL, (SELECT SUM(VALOR_HAVER) FROM parcelas_haver WHERE (COD_PARCELA = parcelas.CODIGO)) AS varSomaHaveres, CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago " & _
       "FROM cliente INNER JOIN pedidos ON cliente.CODIGO = pedidos.COD_CLIENTE INNER JOIN parcelas ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "Where (Cliente.Codigo = " & txtCodigo.Text & ") " & _
       "ORDER BY parcelas.CODIGO, parcelas.NUMERO"
       
'Set r = dbData.OpenRecordset(sSQL)

'lblQuantHistorico.Caption = r.RecordCount

'FormatarGrid_Historico r

'PrintSQL = sSQL

'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(sSQL)

Set REL_Parcelas_PorCliente.Relatorio.Recordset = r

REL_Parcelas_PorCliente.dfQuant.Caption = r.RecordCount
'REL_Parcelas_PorCliente.dfBruto.Caption = lblTotalHistorico.Caption
REL_Parcelas_PorCliente.dfNome.Caption = txtNome.Text

REL_Parcelas_PorCliente.Relatorio.Ativar
Unload REL_Parcelas_PorCliente

Me.Show 1

If r.State <> 0 Then r.Close
Set r = Nothing


End Sub

Private Sub cmdImprimir_Click()
'Me.Hide
'Set Imp_ListaClientes.Relatorio.Recordset = Data4.Recordset
'Imp_ListaClientes.dfQuant.Caption = "Quant. de Registro(s): " & lblQuant.Caption
'Imp_ListaClientes.Relatorio.Ativar
Unload Imp_ListaClientes
'Me.Show 1

Dim var_ImpNormal As String

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

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


Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set Imp_ListaClientes.Relatorio.Recordset = r
'Imp_ListaClientes.dfQuant.Caption = lblTotalNota.Caption
'Imp_ListaClientes.dfBruto.Caption = lblSomaNota.Caption

'If cboFiltroNota.Text = "MENSAL" Then
'   Imp_ListaClientes.dfTipo.Caption = "Tipo: Mês = " & cboConNotaMes.Text & "/" & cboConNotaAno.Text
'ElseIf cboFiltroNota.Text = "DATAS" Then
'   Imp_ListaClientes.dfTipo.Caption = "Tipo: Datas = " & mskConNotaInicial.Text & " à " & mskConNotaFinal.Text
'ElseIf cboFiltroNota.Text = "FORNECEDOR" Then
'   Imp_ListaClientes.dfTipo.Caption = "Tipo: Fornecedor = " & cboConNotaCliente.Text & ""
'ElseIf cboFiltroNota.Text = "NOTA FISCAL" Then
'   Imp_ListaClientes.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & txtConNotaNumNota.Text & ""
'Else
'   Imp_ListaClientes.dfTipo.Caption = "Tipo: Todas as notas"
'End If

Imp_ListaClientes.Relatorio.Ativar
Unload Imp_ListaClientes
Me.Show 1

End Sub

Private Sub cmdImprimirHistorico_Click()
Dim r As ADODB.Recordset

'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Clientes_Historico.Relatorio.Recordset = r

REL_Clientes_Historico.dfQuant.Caption = lblQuantHistorico.Caption
REL_Clientes_Historico.dfBruto.Caption = lblTotalHistorico.Caption
REL_Clientes_Historico.dfNome.Caption = txtNome.Text

REL_Clientes_Historico.Relatorio.Ativar
Unload REL_Clientes_Historico

Me.Show 1
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label32.ForeColor = vbBlack
Label32.Font.Bold = False
End Sub


Private Sub Grid_Consulta_DblClick()
SSTab1.Tab = 0
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
txtCodigo.Text = ""
txtCodigo.Text = Format((Grid_Consulta.TextMatrix(Grid_Consulta.Row, 1)), "00000")
End Sub

Private Sub Label32_Click()
Clipboard.Clear
Clipboard.SetText mskCPF.Text
Label32.ForeColor = &HC0&
End Sub




Private Sub Label32_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label32.ForeColor = vbRed
Label32.Font.Bold = True
End Sub


Private Sub mskCelular_KeyPress(KeyAscii As Integer)
   mskCelular.Mask = "(##) #####-####"
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
If mskCEP.Text = "__.___-___" Then
   mskCEP.Mask = ""
   mskCEP.Text = ""
End If

If txtCodigoIBGE.Text = "" Then: Exit Sub
If mskCEP.Text = "" Then: Exit Sub

sSQL = "SELECT CodigoMunicipio, ID, CEP FROM CIDADE WHERE (CodigoMunicipio = " & txtCodigoIBGE.Text & ")"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    If IsNull(r("CEP")) Then
        dbData.Execute "UPDATE Cidade SET CEP = '" & mskCEP.Text & "' WHERE (CodigoMunicipio = " & txtCodigoIBGE.Text & ")"
    End If
Else

End If

'UPDATE Cidade set CEP = '65.800-000' WHERE CodigoMunicipio = '2101400'

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub mskConsCPF_KeyPress(KeyAscii As Integer)
If chkCPF.Value = True Then
    mskConsCPF.Mask = "###.###.###-##"
ElseIf chkCNPJ.Value = True Then
    mskConsCPF.Mask = "##.###.###/####-##"
End If
End Sub


Private Sub mskCPF_GotFocus()
SelectControl mskCPF
End Sub

Private Sub mskCPF_KeyPress(KeyAscii As Integer)
If cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then
    If cboTipoPessoa.Text = "FÍSICA" Then
        mskCPF.Mask = "###.###.###-##"
    ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
        mskCPF.Mask = "##.###.###/####-##"
    End If
ElseIf cboTipo.Text = "2 - CONTRIBUINTE ISENTO" Then
    If cboTipoPessoa.Text = "FÍSICA" Then
        mskCPF.Mask = "###.###.###-##"
    ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
        mskCPF.Mask = "##.###.###/####-##"
    End If
ElseIf cboTipo.Text = "9 - NÃO CONTRIBUINTE" Then
    mskCPF.Mask = "###.###.###-##"
End If
End Sub

Function calculaCPF(CPF As String) As Boolean
'Esta rotina foi adaptada da revista Fórum Access
On Error GoTo Err_CPF
Dim i As Integer 'utilizada nos FOR... NEXT
Dim strcampo As String 'armazena do CPF que será utilizada para o cálculo
Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
Dim intNumero As Integer 'armazena o digito separado para cálculo (uma a um)
Dim intMais As Integer 'armazena o digito específico multiplicado pela sua base
Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
Dim dblDivisao As Double 'armazena a divisão dos digitos*base por 11
Dim lngInteiro As Long 'armazena inteiro da divisão
Dim intResto As Integer 'armazena o resto
Dim intDig1 As Integer 'armazena o 1º digito verificador
Dim intDig2 As Integer 'armazena o 2º digito verificador
Dim strConf As String 'armazena o digito verificador

lngSoma = 0
intNumero = 0
intMais = 0
strcampo = RemoverFormato(CPF)
strcampo = Left(strcampo, 9)

'Inicia cálculos do 1º dígito
For i = 2 To 10
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11

lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If

strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
lngSoma = 0
intNumero = 0
intMais = 0
'Inicia cálculos do 2º dígito
For i = 2 To 11
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2

'Caso o CPF esteja errado dispara a mensagem
If strConf <> Right(CPF, 2) Then
    calculaCPF = False
Else
    calculaCPF = True
End If

If calculaCPF = False Then
    MsgBox "CPF inválido!", vbInformation, "Aviso do Sistema"
    mskCPF.SetFocus
End If

Exit Function

Exit_CPF:
    Exit Function
Err_CPF:
    'Merlin Error$
    Resume Exit_CPF
End Function



Public Function ValidaCNPJ(CGC As String) As Boolean
If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
   ValidaCGC = False
   Exit Function
End If

If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
    ValidaCGC = False
    MsgBox "CNPJ Inválido!", vbInformation, "Aviso do Sistema"
    mskCPF.SetFocus
    Exit Function
End If

ValidaCGC = True

End Function





Public Function CalculaCGC(Numero As String) As String
Dim i As Integer
Dim PROD As Integer
Dim Mult As Integer
Dim Digito As Integer

Numero = RemoverFormato(Numero) 'fica somente os numeros sem formatação

If Not IsNumeric(Numero) Then
   CalculaCGC = ""
   Exit Function
End If

Mult = 2
For i = Len(Numero) To 1 Step -1
  PROD = PROD + Val(Mid(Numero, i, 1)) * Mult
  Mult = IIf(Mult = 9, 2, Mult + 1)
Next

Digito = 11 - Int(PROD Mod 11)
Digito = IIf(Digito = 10 Or Digito = 11, 0, Digito)

CalculaCGC = Trim(str(Digito))
ValidaCNPJ (Numero)

End Function

Private Sub mskCPF_LostFocus()
If mskCPF.Text = "___.___.___-__" Or mskCPF.Text = "__.___.___/____-__" Then mskCPF.Mask = "": mskCPF.Text = "": Exit Sub

Dim vCPF As String
vCPF = RemoverFormato(mskCPF.Text)

If cboTipoCliente.Text = "CADASTRO" Then
     Select Case Len(vCPF)
            Case 0
                If Len(vCPF) = 0 Then
                    vCPF = Empty
                Else
                    mskCPF.SetFocus
                End If
                KeyCode = 0
            Case 14
                If Validar_CNPJ(vCPF) = False Then
                    MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                    mskCPF.SetFocus
                End If
            Case 11
                If Validar_CPF(vCPF) = False Then
                    MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                    mskCPF.SetFocus
                End If
            Case Is < 11
                MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                mskCNPJ.SetFocus
    End Select
End If
End Sub

Private Sub mskCPF_Validate(Cancel As Boolean)
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long

If mskCPF.Text <> "" Then
    If cmdAlterar.Enabled = False Then
        'sSQL = "SELECT cpf FROM cliente WHERE (cpf = '" & mskCPF.Text & "');"
        'Set r = dbData.OpenRecordset(sSQL, totalRegistros)
        'If r.State <> 0 Then r.Close
        'Set r = Nothing
       
       If totalRegistros > 0 Then
          ShowMsg "Já existe um cliente candastrado com esse CPF/CNPJ.", vbInformation
          mskCPF.Mask = ""
          mskCPF.Text = ""
          mskCPF.SetFocus
          Exit Sub
       End If
    End If
End If
End Sub


Private Sub mskNascimento_Change()
   'Calcular_Idade2 CDate(mskNascimento)
End Sub

Public Function Idade(dtNasc As Date, dtHoje As Date) As Integer
   '   Função que calcula a idade de uma pessoa
   If Month(Date) < Month(dtNasc) Or (Month(Date) = Month(dtNasc) And Day(Date) < Day(dtNasc)) Then Idade = Year(Date) - Year(dtNasc) - 1 Else Idade = Year(Date) - Year(dtNasc)
End Function

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

If txtCodUF.Text = "" Then Exit Sub

sSQL = "SELECT DISTINCT NOME, IdEstado, ID FROM CIDADE WHERE (IdEstado = " & txtCodUF.Text & ") ORDER BY NOME"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboCidade.AddItem ValidateNull(r("NOME"))
   cboCidade.ItemData(cboCidade.NewIndex) = r("ID")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboCidade
End Sub

Private Sub cboEstado_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim varTexto As String

'varTexto = cboEstado.Text

'Limpa a lista
cboEstado.Clear

sSQL = "SELECT DISTINCT UF, IdEstado FROM CIDADE UF ORDER BY UF"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboEstado.AddItem ValidateNull(r("UF"))
   cboEstado.ItemData(cboEstado.NewIndex) = r("IdEstado")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

'cboEstado.Text = varTexto

moCombo.AttachTo cboEstado
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Public Function TirarEspaco(ByVal Value As String) As String
Dim bRepete As Boolean
Value = Replace$(Value, "'", vbNullString)
Do
  Value = Replace$(Value, "  ", " ")
  bRepete = InStr(1, Value, "  ", vbTextCompare)
  Value = Trim(Value)
Loop Until Not bRepete

TirarEspaco = Value
End Function
Private Sub cmdSalvar_Click()
If cboTipoCliente.Text = "PRÉ-CADASTRO" Then
    If txtNome.Text = "" Then MsgBox "Campo NOME obrigatório!", vbCritical, "Online Commerce": txtNome.SetFocus: Exit Sub
    If mskCelular.Text = "" Then MsgBox "Campo CELULAR obrigatório!", vbCritical, "Online Commerce": mskCelular.SetFocus: Exit Sub
Else
    If txtNome.Text = "" Then MsgBox "Campo NOME obrigatório!", vbCritical, "Online Commerce": txtNome.SetFocus: Exit Sub
    If txtEndereco.Text = "" Then MsgBox "Campo ENDEREÇO obrigatório!", vbCritical, "Online Commerce": txtEndereco.SetFocus: Exit Sub
    If txtNum.Text = "" Then MsgBox "Campo NÚMERO obrigatório!", vbCritical, "Online Commerce": txtNum.SetFocus: Exit Sub
    If cboCadBairro.Text = "" Then MsgBox "Campo BAIRRO obrigatório!", vbCritical, "Online Commerce": cboCadBairro.SetFocus: Exit Sub
    If Len(cboCadBairro.Text) > 20 Then MsgBox "Campo BAIRRO possui mais caracteres que o permitido!", vbCritical, "Online Commerce": cboCadBairro.SetFocus: Exit Sub
    If cboCidade.Text = "" Then MsgBox "Campo CIDADE obrigatório!", vbCritical, "Online Commerce": cboCidade.SetFocus: Exit Sub
    If cboEstado.Text = "" Then MsgBox "Campo ESTADO obrigatório!", vbCritical, "Online Commerce": cboEstado.SetFocus: Exit Sub
    If mskCEP.Text = "" Then MsgBox "Campo CEP obrigatório!", vbCritical, "Online Commerce": mskCEP.SetFocus: Exit Sub
    If mskCPF.Text = "" Then MsgBox "Campo CPF obrigatório!", vbCritical, "Online Commerce": mskCPF.SetFocus: Exit Sub
    If cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then
        If txtIE.Text = "" Then MsgBox "Campo INSCRIÇÃO ESTADUAL obrigatório!", vbCritical, "Online Commerce": txtIE.SetFocus: Exit Sub
    End If
    If txtCodigoIBGE.Text = "" Or txtCodigoIBGE.Text = "0" Or Len(txtCodigoIBGE.Text) <> 7 Then
        MsgBox "Campo CÓDIGO IBGE obrigatório!", vbCritical, "Online Commerce": txtCodigoIBGE.SetFocus: Exit Sub
    End If
End If

If cboTipoCliente.Text <> "PRÉ-CADASTRO" Then
    If cboTipoPessoa.Text = "FÍSICA" And cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!", vbCritical, "Online Commerce":  Exit Sub
    If cboTipoPessoa.Text = "JURÍDICA" And cboTipo.Text = "9 - NÃO CONTRIBUINTE" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!", vbCritical, "Online Commerce":  Exit Sub
    If cboTipoPessoa.Text = "RURAL" And cboTipo.Text = "9 - NÃO CONTRIBUINTE" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!", vbCritical, "Online Commerce":  Exit Sub
    If cboTipoPessoa.Text = "RURAL" And cboTipo.Text = "2 - CONTRIBUINTE ISENTO" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!", vbCritical, "Online Commerce":  Exit Sub
End If

'Faz a inserção de forma direta e verifica se houve algum erro
If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Campos_Brancos
Form_Load
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
End Sub

Private Sub cmdNovo_Click()
Campos_Brancos
LimparGrid_Historico
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdNovo.Enabled = False
Frame1.Enabled = True
Frame2.Enabled = True
txtCadastro.Text = Format(Date, ocDATA)
cboStatus.Text = "ATIVO"
cboTipoCliente.Text = "PRÉ-CADASTRO"
AutoNumeracao
cboTipoPessoa.ListIndex = 0
cboTipo.ListIndex = 2
If cboTipo.Text = "1 - CONTRIBUINTE ICMS" Then
    lblCPF.Caption = "CNPJ"
ElseIf cboTipo.Text = "2 - CONTRIBUINTE ISENTO" Then
    lblCPF.Caption = "CNPJ"
ElseIf cboTipo.Text = "9 - NÃO CONTRIBUINTE" Then
    lblCPF.Caption = "CPF"
End If
txtNome.SetFocus
'Form_Load
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
txtCadastro.Text = Format(Date, ocDATA)

Frame1.Enabled = False
Frame2.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdExibir_Click
LimparGrid_Historico
Preencher_TipoPessoa
Preencher_TipoContribuinte
PreencherTipoCadastro
cboTipoCliente.ListIndex = 0
Label32.ForeColor = &H80000012

Set moCombo = New cComboHelper
End Sub

'Acrescentado o paramento rTabela para passa a consulta realizada
Private Sub FormatarGrid_Historico(rTabela As Recordset)
   Dim i As Integer, x As Integer
   
   With Grid_Historico
      .Clear
      .Cols = 10
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 1500
      .ColWidth(4) = 1600
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1000
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "FORMA"
      .TextMatrix(0, 4) = "TIPO"
      .TextMatrix(0, 5) = "VALOR"
      .TextMatrix(0, 6) = "DESC."
      .TextMatrix(0, 7) = "ACRESC."
      .TextMatrix(0, 8) = "TOTAL"
      .TextMatrix(0, 9) = "VENDA"
      
      'colocar os cabeçalho em negrito
      For x = 0 To .Cols - 1
         .Col = x
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         i = 1
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("cod_pedido"), "000000")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("tipo_pagamento"))
            If Not IsNull(rTabela("pagamento")) Then .TextMatrix(.rows - 1, 4) = rTabela("pagamento")
            .TextMatrix(.rows - 1, 5) = Format(rTabela("SUBTOTAL"), ocMONEY)
            .TextMatrix(.rows - 1, 6) = Format(rTabela("ValorDescReal"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = Format(rTabela("ValorAcrescReal"), ocMONEY)
            .TextMatrix(.rows - 1, 8) = Format(rTabela("total"), ocMONEY)
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("TIPO_PEDIDO"))
            
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If

      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 8
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   lblTotalHistorico.Caption = Format(SomaGrid(Grid_Historico, 8), ocMONEY)
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
Dim i As Integer, Valor As Currency

Valor = 0
For i = 0 To var_Grid.rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CCur(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function

'Acrescentado o paramento rTabela para passa a consulta realizada
Private Sub FormatarGrid_Consulta(rTabela As Recordset)
Dim i As Integer, x As Integer

With Grid_Consulta
   .Enabled = False
   
   .Clear
   .Cols = 11
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 3700
   .ColWidth(3) = 1150
   .ColWidth(4) = 2100
   .ColWidth(5) = 500
   .ColWidth(6) = 1100
   .ColWidth(7) = 1100
   .ColWidth(8) = 350
   .ColWidth(10) = 950
   .ColWidth(9) = 1300
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "NOME"
   .TextMatrix(0, 3) = "TELEFONE"
   .TextMatrix(0, 4) = "ENDEREÇO"
   .TextMatrix(0, 5) = "NUM"
   .TextMatrix(0, 6) = "BAIRRO"
   .TextMatrix(0, 7) = "CIDADE"
   .TextMatrix(0, 8) = "UF"
   .TextMatrix(0, 10) = "IE"
   .TextMatrix(0, 9) = "CPF/CNPJ"
   .Redraw = False
   
   If Not rTabela Is Nothing Then
      i = 1
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.rows - 1, 2) = rTabela("nome")
         .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("telefone1"))
         .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("endereco"))
         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("numero"))
         .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("bairro"))
         .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("cidade"))
         .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("estado"))
         .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("IE"))
         .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("cpf"))
         
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
   .rows = .rows - 1
   .Enabled = True
   .Redraw = True
End With
End Sub

Private Sub Calcular_Idade(wNascimento)
   Dim wMeses As Long, wIdade As Long
   Dim wMesesRestantes As Long
   
   If IsDate(wNascimento) = False Then Exit Sub
   
   wMeses = DateDiff("m", wNascimento, Now)
   wIdade = Int(wMeses / 12)
   wMesesRestantes = wMeses - (wIdade * 12)
   txtIdade = wIdade
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

Private Sub mskNascimento_LostFocus()
   If mskNascimento.Text = "" Or mskNascimento.Text = "__/__/____" Then
      mskNascimento.Mask = ""
      mskNascimento.Text = ""
      txtIdade.Text = ""
   Else
      If IsDate(mskNascimento.Text) Then
         Calcular_Idade CDate(mskNascimento)
      Else
          ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskNascimento.SetFocus
      End If
   End If
End Sub

Private Sub mskTelefone1_KeyPress(KeyAscii As Integer)
   mskTelefone1.Mask = "(##) ####-####"
End Sub

Private Sub mskTelefone1_LostFocus()
   If mskTelefone1.Text = "(__) ____-____" Then
      mskTelefone1.Mask = ""
      mskTelefone1.Text = ""
   End If
End Sub

Private Sub optAtivos_Click()
   cmdExibir_Click
End Sub

Private Sub optCidade_Click()
cmdExibir_Click
End Sub

Private Sub optCPF_Click()
cmdExibir_Click
End Sub


Private Sub optInativos_Click()
   cmdExibir_Click
End Sub

Private Sub optNome_Click()
   cmdExibir_Click
End Sub

Private Sub cboCadBairro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodCid_Change()
If txtCodCid.Text = "" Then txtCodigoIBGE.Text = "": Exit Sub
mskCEP.Mask = ""
mskCEP.Text = ""

sSQL = "SELECT CodigoMunicipio, ID, CEP FROM CIDADE WHERE (id = " & txtCodCid.Text & ")"
Set r = dbData.OpenRecordset(sSQL)

   If Not r.BOF Then
    txtCodigoIBGE.Text = ValidateNull(r("CodigoMunicipio"))
    mskCEP.Text = ValidateNull(r("CEP"))
   End If
   
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub txtCodigo_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If cmdSalvar.Enabled = False Then
   If txtCodigo.Text = "" Then Exit Sub
   
   sSQL = "SELECT * FROM cliente WHERE (codigo = " & txtCodigo.Text & ");"
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
   
   If Len(mskCPF) > 10 Then
    cboTipoCliente.Text = "CADASTRO"
   Else
    cboTipoCliente.Text = "PRÉ-CADASTRO"
   End If
   
   LimparGrid_Historico
   Mostrar_Historico
   'Limite_Cliente
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub

Private Sub txtCodigoIBGE_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "EkklesiaSoft": Exit Sub
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEndereco_LostFocus()
txtEndereco.Text = TirarEspaco(txtEndereco.Text)
End Sub


Private Sub txtFiliacao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLimite_LostFocus()
   If txtLimite.Text = "" Then
      txtLimite.Text = Format(0, "##,##0.00")
   Else
      txtLimite.Text = Format(txtLimite, "##,##0.00")
   End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNome_LostFocus()
txtNome.Text = TirarEspaco(txtNome.Text)
End Sub


Private Sub txtNum_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtProfissao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
