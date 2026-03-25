VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form Clientes_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CLIENTES"
   ClientHeight    =   8700
   ClientLeft      =   -870
   ClientTop       =   435
   ClientWidth     =   13020
   Icon            =   "Clientes_Cadastro.frx":0000
   LinkTopic       =   "Form49"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   12825
      TabIndex        =   73
      Top             =   60
      Width           =   12855
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
         TabIndex        =   74
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
         TabIndex        =   75
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
      Height          =   6675
      Left            =   60
      TabIndex        =   37
      Top             =   1080
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   11774
      _Version        =   393216
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
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "HISTÓRICO"
      TabPicture(1)   =   "Clientes_Cadastro.frx":3749
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid_Historico"
      Tab(1).Control(1)=   "lblTotalHistorico"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "CONSULTA"
      TabPicture(2)   =   "Clientes_Cadastro.frx":3765
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Grid_Consulta"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(4)=   "cmdExibir"
      Tab(2).Control(5)=   "cmdImprimir"
      Tab(2).Control(6)=   "lblQuant"
      Tab(2).Control(7)=   "Label26"
      Tab(2).Control(8)=   "Label25"
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
         Left            =   -74880
         TabIndex        =   80
         Top             =   4320
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
            TabIndex        =   82
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
            TabIndex        =   81
            Top             =   480
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   77
         Top             =   420
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   6376
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
         Left            =   -74880
         TabIndex        =   69
         Top             =   5100
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   70
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
         Left            =   -73380
         TabIndex        =   65
         Top             =   4320
         Width           =   11115
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
            TabIndex        =   95
            Top             =   240
            Width           =   555
         End
         Begin VB.ComboBox cboNome 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkCPF 
            Caption         =   "CPF/CNPJ:"
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
            Top             =   900
            Width           =   1575
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
            TabIndex        =   94
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
         Height          =   4935
         Left            =   120
         TabIndex        =   43
         Top             =   1620
         Width           =   10335
         Begin VB.TextBox txtIE 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   9060
            TabIndex        =   21
            Top             =   2580
            Width           =   1155
         End
         Begin VB.TextBox txtNum 
            BackColor       =   &H00C0FFFF&
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4500
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1260
            Width           =   675
         End
         Begin VB.TextBox txtCodCid 
            Height          =   345
            Left            =   2100
            TabIndex        =   87
            Top             =   2280
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtCodUF 
            Height          =   345
            Left            =   420
            TabIndex        =   86
            Top             =   2220
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtCodigoIBGE 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   9000
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   28
            Top             =   3900
            Width           =   1215
         End
         Begin VB.TextBox txtConjuge 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   2460
            MaxLength       =   100
            TabIndex        =   27
            Top             =   3900
            Width           =   6495
         End
         Begin VB.ComboBox cboCadBairro 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   5220
            TabIndex        =   8
            Top             =   1260
            Width           =   1875
         End
         Begin VB.ComboBox cboEstado 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Clientes_Cadastro.frx":3781
            Left            =   120
            List            =   "Clientes_Cadastro.frx":3783
            TabIndex        =   15
            Top             =   2580
            Width           =   615
         End
         Begin VB.ComboBox cboCidade 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Clientes_Cadastro.frx":3785
            Left            =   780
            List            =   "Clientes_Cadastro.frx":3787
            TabIndex        =   16
            Top             =   2580
            Width           =   2535
         End
         Begin VB.TextBox txtCI 
            Height          =   315
            Left            =   7800
            TabIndex        =   20
            Top             =   2580
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "Data_de_Nascimento"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   8160
            MaxLength       =   5
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtNome 
            BackColor       =   &H00C0FFFF&
            DataField       =   "CPF"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            MaxLength       =   80
            TabIndex        =   5
            Top             =   600
            Width           =   10065
         End
         Begin VB.TextBox txtEndereco 
            BackColor       =   &H00C0FFFF&
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1260
            Width           =   4335
         End
         Begin VB.TextBox txtReferencia 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   7140
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1260
            Width           =   3075
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   6120
            MaxLength       =   30
            TabIndex        =   14
            Top             =   1920
            Width           =   4095
         End
         Begin VB.TextBox txtIdade 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   3240
            Width           =   435
         End
         Begin VB.ComboBox cboSexo 
            Height          =   315
            ItemData        =   "Clientes_Cadastro.frx":3789
            Left            =   4500
            List            =   "Clientes_Cadastro.frx":378B
            TabIndex        =   18
            Top             =   2580
            Width           =   1515
         End
         Begin VB.ComboBox cboEstadoCivil 
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   3900
            Width           =   2295
         End
         Begin VB.TextBox txtFiliacao 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   24
            Top             =   3240
            Width           =   5355
         End
         Begin MSMask.MaskEdBox mskCEP 
            Height          =   315
            Left            =   3360
            TabIndex        =   17
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
            TabIndex        =   22
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
            TabIndex        =   19
            Top             =   2580
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskTelefone2 
            Height          =   315
            Left            =   1620
            TabIndex        =   11
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFax 
            Height          =   315
            Left            =   3120
            TabIndex        =   12
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCelular 
            Height          =   315
            Left            =   4620
            TabIndex        =   13
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
            TabIndex        =   25
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
            Left            =   120
            TabIndex        =   10
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Insc. Estadual"
            Height          =   195
            Left            =   9060
            TabIndex        =   96
            Top             =   2340
            Width           =   1005
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Num.:"
            Height          =   195
            Left            =   4500
            TabIndex        =   88
            Top             =   1020
            Width           =   420
         End
         Begin VB.Label lblCódigoIBGE 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código IBGE"
            Height          =   195
            Left            =   9000
            TabIndex        =   85
            Top             =   3660
            Width           =   915
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Conjuge:"
            Height          =   195
            Left            =   2460
            TabIndex        =   84
            Top             =   3660
            Width           =   630
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Idade"
            Height          =   195
            Left            =   1440
            TabIndex        =   64
            Top             =   3000
            Width           =   405
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "RG"
            Height          =   195
            Left            =   7800
            TabIndex        =   63
            Top             =   2340
            Width           =   240
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Razão Social / Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Ponto de Referência:"
            Height          =   195
            Left            =   7140
            TabIndex        =   60
            Top             =   1020
            Width           =   1515
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Correio Eletrônico:"
            Height          =   195
            Left            =   6120
            TabIndex        =   59
            Top             =   1680
            Width           =   1290
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   1620
            TabIndex        =   58
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   840
            TabIndex        =   57
            Top             =   2340
            Width           =   540
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   2340
            Width           =   255
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Sexo:"
            Height          =   195
            Left            =   4500
            TabIndex        =   55
            Top             =   2340
            Width           =   405
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF:"
            Height          =   195
            Left            =   6060
            TabIndex        =   54
            Top             =   2340
            Width           =   915
         End
         Begin VB.Label Label21 
            Caption         =   "Data de Nasc."
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   5220
            TabIndex        =   52
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   3120
            TabIndex        =   51
            Top             =   1680
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Left            =   4620
            TabIndex        =   50
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Left            =   3360
            TabIndex        =   49
            Top             =   2340
            Width           =   330
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   3660
            Width           =   870
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Left            =   7320
            TabIndex        =   47
            Top             =   3000
            Width           =   690
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Filiação:"
            Height          =   195
            Left            =   1920
            TabIndex        =   46
            Top             =   3000
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1680
            Width           =   675
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
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   10335
         Begin VB.ComboBox cboStatus 
            Height          =   315
            Left            =   60
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   600
            Width           =   1515
         End
         Begin VB.TextBox txtCadastro 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   600
            Width           =   1155
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   2880
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtLimite 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4830
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtUltimaCompra 
            Height          =   315
            Left            =   6210
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   60
            TabIndex        =   76
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
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   2880
            TabIndex        =   41
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite de Crédito:"
            Height          =   195
            Left            =   4830
            TabIndex        =   40
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última Compra:"
            Height          =   195
            Left            =   6210
            TabIndex        =   39
            Top             =   360
            Width           =   1065
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   555
         Left            =   -63960
         TabIndex        =   35
         Top             =   6000
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
         MICON           =   "Clientes_Cadastro.frx":378D
         PICN            =   "Clientes_Cadastro.frx":37A9
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
         Left            =   -65700
         TabIndex        =   36
         Top             =   6000
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
         MICON           =   "Clientes_Cadastro.frx":553B
         PICN            =   "Clientes_Cadastro.frx":5557
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
         Height          =   5895
         Left            =   -74880
         TabIndex        =   78
         Top             =   420
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   10398
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   10560
         TabIndex        =   89
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
         MICON           =   "Clientes_Cadastro.frx":72E9
         PICN            =   "Clientes_Cadastro.frx":7305
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
         Left            =   10560
         TabIndex        =   90
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
         MICON           =   "Clientes_Cadastro.frx":9097
         PICN            =   "Clientes_Cadastro.frx":90B3
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
         Left            =   10560
         TabIndex        =   91
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
         MICON           =   "Clientes_Cadastro.frx":AE45
         PICN            =   "Clientes_Cadastro.frx":AE61
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
         Left            =   10560
         TabIndex        =   92
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
         MICON           =   "Clientes_Cadastro.frx":CBF3
         PICN            =   "Clientes_Cadastro.frx":CC0F
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
         Left            =   10560
         TabIndex        =   93
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
         MICON           =   "Clientes_Cadastro.frx":E9A1
         PICN            =   "Clientes_Cadastro.frx":E9BD
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
         TabIndex        =   79
         Top             =   6360
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
         Left            =   -63240
         TabIndex        =   68
         Top             =   4080
         Width           =   225
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registros"
         Height          =   195
         Left            =   -62940
         TabIndex        =   67
         Top             =   4080
         Width           =   660
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Dê um duplo-clique para ver mais informações"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   66
         Top             =   4080
         Width           =   3255
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   615
      Left            =   11220
      TabIndex        =   29
      Top             =   7800
      Width           =   1695
      _ExtentX        =   2990
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
      MICON           =   "Clientes_Cadastro.frx":1074F
      PICN            =   "Clientes_Cadastro.frx":1076B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   83
      Top             =   8430
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18627
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "19:40"
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
Attribute VB_Name = "Clientes_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

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
   mskTelefone2.Mask = ""
   mskTelefone2.Text = ""
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
   mskFax.Mask = ""
   mskFax.Text = ""
   mskCelular.Mask = ""
   mskCelular.Text = ""
   cboCadBairro.Text = ""
   mskFax.Mask = ""
   mskFax.Text = ""
   cboEstadoCivil.Text = ""
   txtProfissao.Text = ""
   txtFiliacao.Text = ""
   mskCEP.Mask = ""
   mskCEP.Text = ""
   txtConjuge.Text = ""
   txtCodigoIBGE.Text = ""
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
   mskTelefone2.Text = ValidateNull(rTabela("telefone2"))
   cboCidade.Text = ValidateNull(rTabela("cidade"))
   cboEstado.Text = ValidateNull(rTabela("estado"))
   cboSexo.Text = ValidateNull(rTabela("sexo"))
   mskCPF.Text = ValidateNull(rTabela("cpf"))
   txtCI.Text = ValidateNull(rTabela("rg"))
   txtIE.Text = ValidateNull(rTabela("ie"))
   txtEmail.Text = ValidateNull(rTabela("correio_eletronico"))
   If Not IsNull(rTabela("data_de_nascimento")) Then mskNascimento.Text = Format$(ValidateNull(rTabela("data_de_nascimento")), ocDATA)
   txtCadastro.Text = Format$(ValidateNull(rTabela("data_cadastro")), ocDATA)
   cboTipo.Text = ValidateNull(rTabela("tipo"))
   txtLimite.Text = FormatNumber(rTabela("limite_credito"), 2)
   txtUltimaCompra.Text = IIf(IsNull(rTabela("ultima_compra")), "", Format$(rTabela("ultima_compra"), ocDATA))
   mskFax.Text = ValidateNull(rTabela("fax"))
   mskCelular.Text = ValidateNull(rTabela("celular"))
   cboCadBairro.Text = ValidateNull(rTabela("bairro"))
   mskCEP.Text = ValidateNull(rTabela("cep"))
   cboEstadoCivil.Text = ValidateNull(rTabela("estadocivil"))
   txtProfissao.Text = ValidateNull(rTabela("profissao"))
   txtFiliacao.Text = ValidateNull(rTabela("filiacao"))
   txtConjuge.Text = ValidateNull(rTabela("conjuge"))
   txtCodigoIBGE.Text = ValidateNull(rTabela("CodigoIBGE"))
End If
End Sub

Private Sub Mostrar_Historico()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Nenhum código foi informado
   If txtCodigo.Text = "" Then Exit Sub
   
   sSQL = "SELECT cliente.*, pedidos.* FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   'Mostra os dados no grid
   FormatarGrid_Historico r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Preencher_TipoPessoa()
   'If cboTipo.ListCount = 0 Then 'Desnecessário a comparação bastar limpar a lista primeiro
   cboTipo.Clear
   cboTipo.AddItem "FISICA"
   cboTipo.AddItem "JURIDICA"
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

Private Sub cboTipo_GotFocus()
   moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub chkCidade_Click()
If chkCidade.Value = 1 Then
   cboConsCidade.Enabled = True
   cboConsCidade.SetFocus
   
   chkCPF.Value = 0
   chkNome.Value = 0
Else
   cboConsCidade.Enabled = False
End If
End Sub

Private Sub chkCPF_Click()
If chkCPF.Value = 1 Then
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
      
      chkCPF.Value = 0
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

If txtNome.Text = "" Then MsgBox "Campo NOME obrigatório!", vbCritical, "Online Commerce": txtNome.SetFocus: Exit Sub
If txtEndereco.Text = "" Then MsgBox "Campo ENDEREÇO obrigatório!", vbCritical, "Online Commerce": txtEndereco.SetFocus: Exit Sub
If txtNum.Text = "" Then MsgBox "Campo NÚMERO obrigatório!", vbCritical, "Online Commerce": txtNum.SetFocus: Exit Sub
If cboCadBairro.Text = "" Then MsgBox "Campo BAIRRO obrigatório!", vbCritical, "Online Commerce": cboCadBairro.SetFocus: Exit Sub
If cboCidade.Text = "" Then MsgBox "Campo CIDADE obrigatório!", vbCritical, "Online Commerce": cboCidade.SetFocus: Exit Sub
If cboEstado.Text = "" Then MsgBox "Campo ESTADO obrigatório!", vbCritical, "Online Commerce": cboEstado.SetFocus: Exit Sub
If mskCEP.Text = "" Then MsgBox "Campo CEP obrigatório!", vbCritical, "Online Commerce": mskCEP.SetFocus: Exit Sub
If mskCPF.Text = "" Then MsgBox "Campo CPF obrigatório!", vbCritical, "Online Commerce": mskCPF.SetFocus: Exit Sub

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
Campos_Brancos
Form_Load
End Sub

Private Function Inserir_Dados() As Boolean
   'A inclusão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   
   'Validaçao dos dados
   If Trim(txtLimite.Text) = "" Then txtLimite.Text = "0"
   If Trim(txtCodigoIBGE.Text) = "" Then txtCodigoIBGE.Text = "0"
   
   'Comando de inclusão
   sSQL = "INSERT INTO cliente (" & _
      "codigo, status, nome, endereco, numero, ponto_de_referencia, bairro, cep, cidade, estado, " & _
      "telefone1, telefone2, fax, celular, sexo, cpf, rg, ie, correio_eletronico, " & _
      "data_de_nascimento, data_cadastro, tipo, limite_credito, " & _
      "estadocivil, profissao, filiacao, conjuge, CodigoIBGE) VALUES ("
   
   sSQL = sSQL & _
      txtCodigo.Text & ", " & IIf((cboStatus.Text = "ATIVO"), 1, 0) & ", '" & txtNome.Text & "', '" & txtEndereco.Text & "', " & txtNum.Text & ", '" & _
      txtReferencia.Text & "', '" & cboCadBairro.Text & "', '" & mskCEP.Text & "', '" & cboCidade.Text & "', '" & _
      cboEstado.Text & "', '" & mskTelefone1.Text & "', '" & mskTelefone2.Text & "', '" & mskFax.Text & "', '" & _
      mskCelular.Text & "', '" & cboSexo.Text & "', '" & mskCPF.Text & "', '" & txtCI.Text & "', '" & txtIE.Text & "', '" & txtEmail.Text & "', " & _
      IIf((mskNascimento.Text = ""), "Null", "CONVERT(DATETIME, '" & Format$(mskNascimento.Text, ocDATA) & "', 103)") & ", " & _
      "CONVERT(DATETIME, '" & Format$(txtCadastro.Text, ocDATA) & "', 103), '" & cboTipo.Text & "', " & Replace(CCur(txtLimite.Text), ",", ".") & ", '" & _
      cboEstadoCivil.Text & "', '" & txtProfissao.Text & "', '" & txtFiliacao.Text & "', '" & txtConjuge.Text & "', " & txtCodigoIBGE.Text & ")"
  'Debug.Print sSQL
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Validaçao dos dados
   If Trim(txtLimite.Text) = "" Then txtLimite.Text = "0"
   If Trim(txtCodigoIBGE.Text) = "" Then txtCodigoIBGE.Text = "0"
   
   'Comando de atualização
   sSQL = "UPDATE cliente SET " & _
      "status = " & IIf((cboStatus.Text = "ATIVO"), 1, 0) & ", " & _
      "nome = '" & txtNome.Text & "', " & _
      "endereco = '" & txtEndereco.Text & "', " & _
      "numero = " & txtNum.Text & ", " & _
      "ponto_de_referencia = '" & txtReferencia.Text & "', " & _
      "bairro = '" & cboCadBairro.Text & "', " & _
      "cep = '" & mskCEP.Text & "', " & _
      "cidade = '" & cboCidade.Text & "', " & _
      "estado = '" & cboEstado.Text & "', " & _
      "telefone1 = '" & mskTelefone1.Text & "', " & _
      "telefone2 = '" & mskTelefone2.Text & "', " & _
      "fax = '" & mskFax.Text & "', " & _
      "celular = '" & mskCelular.Text & "', "

   sSQL = sSQL & _
      "sexo = '" & cboSexo.Text & "', " & _
      "cpf = '" & mskCPF.Text & "', " & _
      "rg = '" & txtCI.Text & "', " & _
      "ie = '" & txtIE.Text & "', " & _
      "correio_eletronico = '" & txtEmail.Text & "', " & _
      "data_de_nascimento = CONVERT(DATETIME, '" & Format$(mskNascimento.Text, ocDATA) & "', 103), " & _
      "tipo = '" & cboTipo.Text & "', " & _
      "limite_credito = " & Replace(CCur(txtLimite.Text), ",", ".") & ", " & _
      "estadocivil = '" & cboEstadoCivil.Text & "', " & _
      "profissao = '" & txtProfissao.Text & "', " & _
      "filiacao = '" & txtFiliacao.Text & "', " & _
      "conjuge = '" & txtConjuge.Text & "', " & _
      "CodigoIBGE = " & txtCodigoIBGE.Text
   'Debug.Print sSQL
   'Condição para atualização
   sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdCancelar_Click()
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   Campos_Brancos
   Frame1.Enabled = False
   Frame2.Enabled = False
   cboStatus.Text = "ATIVO"
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
   
   If txtNome.Text = "CONSUMIDOR" Then Exit Sub
   
   If txtCodigo.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o cliente na guia CONSULTA", vbInformation
      Exit Sub
   End If
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir esse cliente?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Não é necessário consulta o registro antes de exclui-lo
   'sSQL = "SELECT * FROM cliente WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM cliente WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
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
      
   If chkCPF.Value = 1 Then
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
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdImprimir_Click()
   'Me.Hide
   'Set Imp_ListaClientes.Relatorio.Recordset = Data4.Recordset
   'Imp_ListaClientes.dfQuant.Caption = "Quant. de Registro(s): " & lblQuant.Caption
   'Imp_ListaClientes.Relatorio.Ativar
   'Unload Imp_ListaClientes
   'Me.Show 1
End Sub

Private Sub Grid_Consulta_DblClick()
   SSTab1.Tab = 0
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   txtCodigo.Text = ""
   txtCodigo.Text = (Grid_Consulta.TextMatrix(Grid_Consulta.Row, 1))
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
   If mskCEP.Text = "__.___-___" Then
      mskCEP.Mask = ""
      mskCEP.Text = ""
   End If
End Sub

Private Sub mskConsCPF_KeyPress(KeyAscii As Integer)
mskConsCPF.Mask = "###.###.###-##"
End Sub


Private Sub mskCPF_KeyPress(KeyAscii As Integer)
   If cboTipo.Text = "FISICA" Then
      mskCPF.Mask = "###.###.###-##"
   Else
      mskCPF.Mask = "##.###.###/####-##"
   End If
End Sub

Private Sub mskFax_KeyPress(KeyAscii As Integer)
   mskFax.Mask = "(##) ####-####"
End Sub

Private Sub mskFax_LostFocus()
   If mskFax.Text = "(__) ____-____" Then
      mskFax.Mask = ""
      mskFax.Text = ""
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

Private Sub cmdSalvar_Click()
If txtNome.Text = "" Then MsgBox "Campo NOME obrigatório!", vbCritical, "Online Commerce": txtNome.SetFocus: Exit Sub
If txtEndereco.Text = "" Then MsgBox "Campo ENDEREÇO obrigatório!", vbCritical, "Online Commerce": txtEndereco.SetFocus: Exit Sub
If txtNum.Text = "" Then MsgBox "Campo NÚMERO obrigatório!", vbCritical, "Online Commerce": txtNum.SetFocus: Exit Sub
If cboCadBairro.Text = "" Then MsgBox "Campo BAIRRO obrigatório!", vbCritical, "Online Commerce": cboCadBairro.SetFocus: Exit Sub
If cboCidade.Text = "" Then MsgBox "Campo CIDADE obrigatório!", vbCritical, "Online Commerce": cboCidade.SetFocus: Exit Sub
If cboEstado.Text = "" Then MsgBox "Campo ESTADO obrigatório!", vbCritical, "Online Commerce": cboEstado.SetFocus: Exit Sub
If mskCEP.Text = "" Then MsgBox "Campo CEP obrigatório!", vbCritical, "Online Commerce": mskCEP.SetFocus: Exit Sub
If mskCPF.Text = "" Then MsgBox "Campo CPF obrigatório!", vbCritical, "Online Commerce": mskCPF.SetFocus: Exit Sub

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
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
Campos_Brancos
Form_Load
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
cboStatus.Text = "ATIVO"
AutoNumeracao
LimparGrid_Historico
cboTipo.ListIndex = 0
txtNome.SetFocus
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
   
   Set moCombo = New cComboHelper
End Sub

'Acrescentado o paramento rTabela para passa a consulta realizada
Private Sub FormatarGrid_Historico(rTabela As Recordset)
   Dim i As Integer, x As Integer
   
   With Grid_Historico
      .Clear
      .Cols = 6
      .Rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 1200
      .ColWidth(4) = 1100
      .ColWidth(5) = 1500
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "FORMA"
      .TextMatrix(0, 5) = "TIPO"
      
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
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("cod_pedido"), "000000")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("total"), ocMONEY)
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("tipo_pagamento"))
            If Not IsNull(rTabela("pagamento")) Then .TextMatrix(.Rows - 1, 5) = rTabela("pagamento")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblTotalHistorico.Caption = Format(SomaGrid(Grid_Historico, 3), ocMONEY)
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
Dim i As Integer, Valor As Currency

Valor = 0
For i = 0 To var_Grid.Rows - 1
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
   .Rows = 2
   
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
         .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.Rows - 1, 2) = rTabela("nome")
         .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("telefone1"))
         .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("endereco"))
         .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("numero"))
         .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("bairro"))
         .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("cidade"))
         .TextMatrix(.Rows - 1, 8) = ValidateNull(rTabela("estado"))
         .TextMatrix(.Rows - 1, 10) = ValidateNull(rTabela("IE"))
         .TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("cpf"))
         
         rTabela.MoveNext
         .Rows = .Rows + 1
         i = i + 1
      Loop
   End If
   
   .Rows = .Rows - 1
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

Private Sub mskTelefone2_KeyPress(KeyAscii As Integer)
   mskTelefone2.Mask = "(##) ####-####"
End Sub

Private Sub mskTelefone2_LostFocus()
   If mskTelefone2.Text = "(__) ____-____" Then
      mskTelefone2.Mask = ""
      mskTelefone2.Text = ""
   End If
End Sub

Private Sub optAtivos_Click()
   cmdExibir_Click
End Sub

Private Sub optBairro_Click()
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
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodCid.Text = "" Then txtCodigoIBGE.Text = "": Exit Sub

sSQL = "SELECT CodigoMunicipio, ID FROM CIDADE WHERE (id = " & txtCodCid.Text & ")"
Set r = dbData.OpenRecordset(sSQL)

   If Not r.BOF Then
    txtCodigoIBGE.Text = ValidateNull(r("CodigoMunicipio"))
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

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
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

Private Sub txtNome_Validate(Cancel As Boolean)
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   txtNome.Text = RemoverAcento(txtNome.Text)
   
   If cmdAlterar.Enabled = False Then
      sSQL = "SELECT nome FROM cliente WHERE (nome = '" & txtNome.Text & "');"
      Set r = dbData.OpenRecordset(sSQL, totalRegistros)
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      If totalRegistros > 0 Then
         ShowMsg "Este cliente já está cadastrado!", vbInformation
         txtNome.Text = ""
         txtNome.SetFocus
         Exit Sub
      End If
   End If
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
