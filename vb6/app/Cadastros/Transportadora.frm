VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Transportadora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TRANSPORTADORAS"
   ClientHeight    =   8310
   ClientLeft      =   -870
   ClientTop       =   435
   ClientWidth     =   9345
   Icon            =   "Transportadora.frx":0000
   LinkTopic       =   "Form49"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   60
      TabIndex        =   29
      Top             =   1080
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11139
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   2293
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
      TabPicture(0)   =   "Transportadora.frx":23D2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frm_Principal"
      Tab(0).Control(1)=   "frm_Secundario"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Transportadora.frx":23EE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Grid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   8985
         TabIndex        =   54
         Top             =   5160
         Width           =   9015
         Begin VB.OptionButton optEstado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Estado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3720
            TabIndex        =   59
            Top             =   60
            Width           =   855
         End
         Begin VB.OptionButton optCidade 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2820
            TabIndex        =   58
            Top             =   60
            Width           =   855
         End
         Begin VB.OptionButton optFantasia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fantasia"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1800
            TabIndex        =   57
            Top             =   60
            Width           =   975
         End
         Begin VB.OptionButton optRazao 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Razão"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   960
            TabIndex        =   56
            Top             =   60
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ordem:"
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
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   60
            Width           =   615
         End
      End
      Begin VB.Frame frm_Principal 
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
         Height          =   4695
         Left            =   -74880
         TabIndex        =   35
         Top             =   480
         Width           =   9015
         Begin VB.ComboBox cboTipoPessoa 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   480
            Width           =   1995
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            ItemData        =   "Transportadora.frx":240A
            Left            =   1920
            List            =   "Transportadora.frx":240C
            TabIndex        =   13
            Top             =   3060
            Width           =   795
         End
         Begin VB.ComboBox cboCidade 
            Height          =   315
            ItemData        =   "Transportadora.frx":240E
            Left            =   120
            List            =   "Transportadora.frx":2410
            TabIndex        =   12
            Top             =   3060
            Width           =   1755
         End
         Begin VB.TextBox txtIE 
            Height          =   315
            Left            =   4560
            TabIndex        =   15
            Top             =   3060
            Width           =   1755
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "Data_de_Nascimento"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8160
            MaxLength       =   5
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtRazao 
            DataField       =   "CPF"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   8745
         End
         Begin VB.TextBox txtEndereco 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   1740
            Width           =   3375
         End
         Begin VB.TextBox txtReferencia 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   5280
            TabIndex        =   6
            Top             =   1740
            Width           =   2055
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   4620
            TabIndex        =   11
            Top             =   2400
            Width           =   4275
         End
         Begin VB.TextBox txtBairro 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   3540
            TabIndex        =   5
            Top             =   1740
            Width           =   1695
         End
         Begin VB.TextBox txtFantasia 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   2160
            MaxLength       =   80
            TabIndex        =   2
            Top             =   480
            Width           =   6495
         End
         Begin VB.TextBox txtContato 
            Height          =   315
            Left            =   6360
            TabIndex        =   16
            Top             =   3060
            Width           =   2535
         End
         Begin MSMask.MaskEdBox mskCEP 
            Height          =   315
            Left            =   7380
            TabIndex        =   7
            Top             =   1740
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCNPJ 
            Height          =   315
            Left            =   2760
            TabIndex        =   14
            Top             =   3060
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskTelefone 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFax 
            Height          =   315
            Left            =   1620
            TabIndex        =   9
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCelular 
            Height          =   315
            Left            =   3120
            TabIndex        =   10
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtObs 
            Height          =   675
            Left            =   120
            TabIndex        =   17
            Top             =   3720
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   1191
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton1 
            Height          =   315
            Left            =   8700
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   ">"
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
            MICON           =   "Transportadora.frx":2412
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
            Caption         =   "Tipo de Pessoa:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Estadual"
            Height          =   195
            Left            =   4560
            TabIndex        =   52
            Top             =   2820
            Width           =   1305
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razão Social*"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label9 
            Caption         =   "Endereço"
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Top             =   1500
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Ponto de Referência"
            Height          =   285
            Left            =   5280
            TabIndex        =   49
            Top             =   1500
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Correio Eletrônico"
            Height          =   285
            Left            =   4620
            TabIndex        =   48
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Telefone:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Cidade*"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   2820
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "UF"
            Height          =   255
            Left            =   1920
            TabIndex        =   45
            Top             =   2820
            Width           =   495
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   2760
            TabIndex        =   44
            Top             =   2820
            Width           =   405
         End
         Begin VB.Label Label1 
            Caption         =   "Bairro"
            Height          =   285
            Left            =   3540
            TabIndex        =   43
            Top             =   1500
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Fax"
            Height          =   255
            Left            =   1620
            TabIndex        =   42
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Celular"
            Height          =   255
            Left            =   3120
            TabIndex        =   41
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Cep"
            Height          =   285
            Left            =   7380
            TabIndex        =   40
            Top             =   1500
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Observação"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Fantasia*"
            Height          =   285
            Left            =   2160
            TabIndex        =   38
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Contato"
            Height          =   255
            Left            =   6360
            TabIndex        =   37
            Top             =   2820
            Width           =   1215
         End
      End
      Begin VB.Frame frm_Secundario 
         Caption         =   "Bancário"
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
         Left            =   -74880
         TabIndex        =   30
         Top             =   5220
         Width           =   9015
         Begin VB.TextBox txtConta 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   6120
            TabIndex        =   21
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtAgencia 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4440
            TabIndex        =   20
            Top             =   480
            Width           =   1635
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   2400
            TabIndex        =   19
            Top             =   480
            Width           =   1995
         End
         Begin VB.ComboBox cboTipoConta 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   2235
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
            Height          =   195
            Left            =   6120
            TabIndex        =   34
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
            Height          =   195
            Left            =   4440
            TabIndex        =   33
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   2400
            TabIndex        =   32
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   360
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4695
         Left            =   120
         TabIndex        =   53
         Top             =   420
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8281
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   9225
      TabIndex        =   27
      Top             =   120
      Width           =   9255
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSPORTADORAS"
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
         TabIndex        =   28
         Top             =   300
         Width           =   3285
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   300
         Picture         =   "Transportadora.frx":242E
         Top             =   0
         Width           =   900
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSalvar 
      Height          =   555
      Left            =   1800
      TabIndex        =   22
      Top             =   7440
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
      MICON           =   "Transportadora.frx":5222
      PICN            =   "Transportadora.frx":523E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAlterar 
      Height          =   555
      Left            =   1800
      TabIndex        =   24
      Top             =   7440
      Visible         =   0   'False
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
      MICON           =   "Transportadora.frx":BB08
      PICN            =   "Transportadora.frx":BB24
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
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   7440
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
      MICON           =   "Transportadora.frx":C3FE
      PICN            =   "Transportadora.frx":C41A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   555
      Left            =   3540
      TabIndex        =   23
      Top             =   7440
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
      MICON           =   "Transportadora.frx":D0F4
      PICN            =   "Transportadora.frx":D110
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExcluir 
      Height          =   555
      Left            =   3540
      TabIndex        =   25
      Top             =   7440
      Visible         =   0   'False
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
      MICON           =   "Transportadora.frx":13BB4
      PICN            =   "Transportadora.frx":13BD0
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
      Left            =   7620
      TabIndex        =   26
      Top             =   7440
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
      MICON           =   "Transportadora.frx":13EEA
      PICN            =   "Transportadora.frx":13F06
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
      TabIndex        =   60
      Top             =   8040
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12144
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "16:11"
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
Attribute VB_Name = "Transportadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private moCombo As cComboHelper

Private Sub Campos_Brancos()
   If cmdAlterar.Visible = False Then txtCodigo.Text = ""
   txtRazao.Text = ""
   cboTipoPessoa.Text = ""
   txtEndereco.Text = ""
   txtReferencia.Text = ""
   mskTelefone.Mask = ""
   mskTelefone.Text = ""
   cboCidade.Clear
   cboCidade.Text = ""
   cboEstado.Clear
   cboEstado.Text = ""
   mskCNPJ.Mask = ""
   mskCNPJ.Text = ""
   txtEmail.Text = ""
   txtIE.Text = ""
   mskFax.Mask = ""
   mskFax.Text = ""
   mskCelular.Mask = ""
   mskCelular.Text = ""
   txtBairro.Text = ""
   mskFax.Mask = ""
   mskFax.Text = ""
   mskCEP.Mask = ""
   mskCEP.Text = ""
   txtContato.Text = ""
   txtFantasia.Text = ""
   txtObs.Text = ""
   cboBanco.Text = ""
   txtAgencia.Text = ""
   txtConta.Text = ""
   cboTipoConta.Text = ""
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
   If Not rTabela Is Nothing Then
      txtCodigo.Text = ValidateNull(rTabela("codigo"))
      txtRazao.Text = ValidateNull(rTabela("razao"))
      cboTipoPessoa.Text = ValidateNull(rTabela("tipo"))
      txtEndereco.Text = ValidateNull(rTabela("endereco"))
      txtReferencia.Text = ValidateNull(rTabela("ponto_de_referencia"))
      mskTelefone.Text = ValidateNull(rTabela("telefone"))
      cboCidade.Text = ValidateNull(rTabela("cidade"))
      cboEstado.Text = ValidateNull(rTabela("estado"))
      mskCNPJ.Text = ValidateNull(rTabela("cnpj"))
      txtIE.Text = ValidateNull(rTabela("ie"))
      txtEmail.Text = ValidateNull(rTabela("correio_eletronico"))
      mskFax.Text = ValidateNull(rTabela("fax"))
      mskCelular.Text = ValidateNull(rTabela("celular"))
      txtBairro.Text = ValidateNull(rTabela("bairro"))
      mskCEP.Text = ValidateNull(rTabela("cep"))
      txtContato.Text = ValidateNull(rTabela("contato"))
      txtFantasia.Text = ValidateNull(rTabela("fantasia"))
      txtObs.Text = ValidateNull(rTabela("obs"))
      cboBanco.Text = ValidateNull(rTabela("banco"))
      txtAgencia.Text = ValidateNull(rTabela("agencia"))
      txtConta.Text = ValidateNull(rTabela("conta"))
      cboTipoConta.Text = ValidateNull(rTabela("tipo_conta"))
   End If
End Sub

Private Sub Mostrar_Fornecedores()
   'INDICE PARA ORGANIZAR OS DADOS
   Dim INDICE As String
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If optRazao.Value = True Then
      INDICE = "razao;"
   ElseIf optFantasia.Value = True Then
      INDICE = "fantasia;"
   ElseIf optCidade.Value = True Then
      INDICE = "cidade;"
   ElseIf optEstado.Value = True Then
      INDICE = "estado;"
   End If
   
   sSQL = "SELECT codigo, razao, fantasia, cidade, estado FROM transportadora ORDER BY " & INDICE
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cboBanco_GotFocus()
   cboBanco.Clear
   cboBanco.AddItem "BANCO DO BRASIL"
   cboBanco.AddItem "BANCO DO NORDESTE"
   cboBanco.AddItem "BRADESCO"
   cboBanco.AddItem "CAIXA ECONOMICA FEDERAL"
   cboBanco.AddItem "ITAU"
   moCombo.AttachTo cboBanco
End Sub

Private Sub cboBanco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboCidade_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
   'If Len(cboEstado) = 1 Then
   'mskCNPJ.SetFocus
   'End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboTipoPessoa_GotFocus()
cboTipoPessoa.Clear
cboTipoPessoa.AddItem "FÍSICA"
cboTipoPessoa.AddItem "JURÍDICA"
moCombo.AttachTo cboTipoPessoa
End Sub

Private Sub cboTipoPessoa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboTipoPessoa_LostFocus()
If cboTipoPessoa.Text = "FÍSICA" Then
    Label20.Caption = "CPF"
    Label17.Caption = "RG"
ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
    Label20.Caption = "CNPJ"
    Label17.Caption = "Inscrição Estadual"
End If
End Sub

Private Sub chameleonButton1_Click()
txtRazao.Text = txtFantasia.Text
End Sub

Private Sub cmdAlterar_Click()
   'On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtRazao.Text = "" Or txtFantasia.Text = "" Or cboCidade.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão cadastrados.", vbInformation
      Exit Sub
   End If
   
   'Não é necessário consulta o registro antes de atualiza-lo
   'sSQL = "SELECT * FROM transportadora WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   frm_Principal.Enabled = False
   frm_Secundario.Enabled = False
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdNovo.Enabled = True
   Campos_Brancos
   Mostrar_Fornecedores

'TrataErro:
   'If Err.Number = 3022 Then
   '   MsgBox "DADOS DUPLICADO!" & vbCrLf & "Verifique se esta O.S. já está cadastrada.", vbInformation, "Aviso do Sistema"
   '   Exit Sub
   'End If
   
   'If Err.Number = 3421 Then
   '   MsgBox "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation, "Aviso do Sistema"
   '   Exit Sub
   '   txtRazao.SetFocus
   'End If
End Sub

Private Function Inserir_Dados() As Boolean
Dim sSQL As String

sSQL = "INSERT INTO transportadora (" & _
   "codigo, tipo, razao, endereco, ponto_de_referencia, telefone, cidade, estado, cnpj, " & _
   "ie, correio_eletronico, fax, celular, bairro, cep, contato, fantasia, obs, " & _
   "banco, agencia, conta, tipo_conta) VALUES ("

sSQL = sSQL & _
   txtCodigo.Text & ", '" & cboTipoPessoa.Text & "', '" & txtRazao.Text & "', '" & txtEndereco.Text & "', '" & txtReferencia.Text & "', '" & _
   mskTelefone.Text & "', '" & cboCidade.Text & "', '" & cboEstado.Text & "', '" & mskCNPJ.Text & "', '" & _
   txtIE.Text & "', '" & txtEmail.Text & "', '" & mskFax.Text & "', '" & mskCelular.Text & "', '" & _
   txtBairro.Text & "', '" & mskCEP.Text & "', '" & txtContato.Text & "', '" & txtFantasia.Text & "', '" & _
   txtObs.Text & "', '" & cboBanco.Text & "', '" & txtAgencia.Text & "', '" & txtConta.Text & "', '" & cboTipoConta.Text & "')"

Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE transportadora SET " & _
      "razao = '" & txtRazao.Text & "', " & _
      "tipo = '" & cboTipoPessoa.Text & "', " & _
      "endereco = '" & txtEndereco.Text & "', " & _
      "ponto_de_referencia = '" & txtReferencia.Text & "', " & _
      "telefone = '" & mskTelefone.Text & "', " & _
      "cidade = '" & cboCidade.Text & "', " & _
      "estado = '" & cboEstado.Text & "', " & _
      "cnpj = '" & mskCNPJ.Text & "', " & _
      "ie = '" & txtIE.Text & "', " & _
      "correio_eletronico = '" & txtEmail.Text & "', " & _
      "fax = '" & mskFax.Text & "', " & _
      "celular = '" & mskCelular.Text & "', " & _
      "bairro = '" & txtBairro.Text & "', " & _
      "cep = '" & mskCEP.Text & "', " & _
      "contato = '" & txtContato.Text & "', " & _
      "fantasia = '" & txtFantasia.Text & "', " & _
      "obs = '" & txtObs.Text & "', " & _
      "banco = '" & cboBanco.Text & "', " & _
      "agencia = '" & txtAgencia.Text & "', " & _
      "conta = '" & txtConta.Text & "', " & _
      "tipo_conta = '" & cboTipoConta.Text & "' " & _
      "WHERE (codigo = " & Me.txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdCancelar_Click()
    Campos_Brancos
    frm_Principal.Enabled = False
    frm_Secundario.Enabled = False
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    cmdAlterar.Visible = False
    cmdExcluir.Visible = False
    cmdNovo.Enabled = True
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   If txtRazao.Text = "" Or txtFantasia.Text = "" Or cboCidade.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão cadastrados.", vbInformation
      Exit Sub
   End If
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Deseja excluir esse transportadora?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Não é necessário consulta o registro antes de exclui-lo
   'sSQL = "SELECT * FROM transportadora WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM transportadora WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
    
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   frm_Principal.Enabled = False
   frm_Secundario.Enabled = False
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdNovo.Enabled = True
   Campos_Brancos
   Mostrar_Fornecedores
End Sub

Private Sub cmdNovo_Click()
frm_Principal.Enabled = True
frm_Secundario.Enabled = True
Campos_Brancos
cmdSalvar.Visible = True
cmdCancelar.Visible = True
cmdAlterar.Visible = False
cmdExcluir.Visible = False
cmdNovo.Enabled = False
Auto_Numeracao
cboTipoPessoa.SetFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Grid_DblClick()
   SSTab1.Tab = 0
   frm_Principal.Enabled = True
   frm_Secundario.Enabled = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdNovo.Enabled = True
   txtCodigo.Text = ""
   txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
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
   If mskCEP.Text = "__.___-__" Then
      mskCEP.Mask = ""
      mskCEP.Text = ""
   End If
End Sub

Private Sub mskCNPJ_GotFocus()
SelectControl mskCNPJ
End Sub

Private Sub mskCNPJ_KeyPress(KeyAscii As Integer)
If cboTipoPessoa.Text = "FÍSICA" Then
    mskCNPJ.Mask = "###.###.###-##"
ElseIf cboTipoPessoa.Text = "JURÍDICA" Then
    mskCNPJ.Mask = "##.###.###/####-##"
End If
End Sub

Private Sub mskCNPJ_LostFocus()
If mskCNPJ.Text = "___.___.___-__" Or mskCNPJ.Text = "__.___.___/____-__" Then mskCNPJ.Mask = "": mskCNPJ.Text = "": Exit Sub

Dim vCPF As String
vCPF = RemoverFormato(mskCNPJ.Text)

'If cboTipoCliente.Text = "CADASTRO" Then
 Select Case Len(vCPF)
        Case 0
            If Len(vCPF) = 0 Then
                vCPF = Empty
            Else
                mskCNPJ.SetFocus
            End If
            KeyCode = 0
        Case 14
            If Validar_CNPJ(vCPF) = False Then
                MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                mskCNPJ.SetFocus
            End If
        Case 11
            If Validar_CPF(vCPF) = False Then
                MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                mskCNPJ.SetFocus
            End If
        Case Is < 11
            MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
            mskCNPJ.SetFocus
End Select
'End If

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

Private Sub cboCidade_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   'Limpa a lista
   cboCidade.Clear
   
   sSQL = "SELECT cidade FROM transportadora GROUP BY cidade;"
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   
   If totalRegistros = 0 Then
      cboCidade.AddItem "Nada"
   Else
      Do While Not r.EOF
         cboCidade.AddItem r("cidade")
         r.MoveNext
      Loop
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboCidade
End Sub

Private Sub cboEstado_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   'Limpa a lista
   cboEstado.Clear
   
   sSQL = "SELECT estado FROM transportadora GROUP BY estado;"
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   
   If totalRegistros = 0 Then
      cboEstado.AddItem "Nada"
   Else
      Do While Not r.EOF
         cboEstado.AddItem ValidateNull(r("estado"))
         r.MoveNext
      Loop
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboEstado
End Sub

Private Sub cmdSalvar_Click()
   'On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtRazao.Text = "" Or txtFantasia.Text = "" Or mskCNPJ.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!", vbExclamation
      txtRazao.SetFocus
      Exit Sub
   End If
   
   'Não é necessário consultar todos os registros antes de inserir um novo
   'sSQL = "SELECT * FROM transportadora;"
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
   
   Campos_Brancos
   frm_Principal.Enabled = False
   frm_Secundario.Enabled = False
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdNovo.Enabled = True
   Mostrar_Fornecedores
   
'TrataErro:
   'If Err.Number = 3315 Then
   '   MsgBox "FORMULÁRIO INCOMPLETO!" & vbCrLf & "É obrigatório o preenchimento todos os campos de Informações", vbInformation, "Aviso do Sistema"
   'End If
   
   'If Err.Number = 3421 Then
   '   MsgBox "FORMULÁRIO INCOMPLETO!" & vbCrLf & "É obrigatório o preenchimento todas as DATAS", vbInformation, "Aviso do Sistema"
   '   txtRazao.SetFocus
   'End If
   
   'If Err.Number = 3022 Then
   '   MsgBox "DADOS DUPLICADO!" & vbCrLf & "Verifique se este aluno ou responsável não está já cadastrado.", vbInformation, "Aviso do Sistema"
   '   txtRazao.SetFocus
   'End If
End Sub

Private Sub Auto_Numeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_transportadora FROM transportadora;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigo.Text = r("cod_transportadora") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Form_Load()
   SSTab1.Tab = 0
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdNovo.Enabled = True
   Mostrar_Fornecedores
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   Set moCombo = New cComboHelper
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   Dim x As Integer
   
   With Grid
      .Clear
      .Cols = 7
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 3100
      .ColWidth(3) = 2800
      .ColWidth(4) = 1300
      .ColWidth(5) = 500
      .ColWidth(6) = 1000
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "RAZÃO SOCIAL"
      .TextMatrix(0, 3) = "FANTASIA"
      .TextMatrix(0, 4) = "CIDADE"
      .TextMatrix(0, 5) = "UF"
      .TextMatrix(0, 6) = "COMPRAS"
      
      'colocar os cabeçalho em negrito
      For x = 0 To .Cols - 1
         .Col = x
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For x = 0 To .Cols - 1
         .Row = 0
         .Col = x
         .CellAlignment = flexAlignCenterCenter
      Next
      
      Grid.Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
         
            'mudar a cor da coluna
            For j = 1 To .rows - 1
               .Row = j
               .Col = 6
               .CellBackColor = &HC0FFFF
            Next
            
            'ALINHAMENTO
            .ColAlignment(2) = 1
            
            .TextMatrix(.rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.rows - 1, 2) = rTabela("razao")
            .TextMatrix(.rows - 1, 3) = rTabela("fantasia")
            .TextMatrix(.rows - 1, 4) = rTabela("cidade")
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("estado"))
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
End Sub

Private Sub cboTipoConta_GotFocus()
   cboTipoConta.Clear
   cboTipoConta.AddItem "POUPANÇA"
   cboTipoConta.AddItem "CONTA CORRENTE"
   moCombo.AttachTo cboTipoConta
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

Private Sub mskTelefone_KeyPress(KeyAscii As Integer)
   mskTelefone.Mask = "(##) ####-####"
End Sub

Private Sub mskTelefone_LostFocus()
   If mskTelefone.Text = "(__) ____-____" Then
      mskTelefone.Mask = ""
      mskTelefone.Text = ""
   End If
End Sub

Private Sub optCidade_Click()
   Mostrar_Fornecedores
End Sub

Private Sub optEstado_Click()
   Mostrar_Fornecedores
End Sub

Private Sub optFantasia_Click()
   Mostrar_Fornecedores
End Sub

Private Sub optRazao_Click()
   Mostrar_Fornecedores
End Sub

Private Sub txtBairro_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodigo_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cmdSalvar.Visible = False Then
      If txtCodigo.Text = "" Then Exit Sub
      
      sSQL = "SELECT * FROM transportadora WHERE (codigo = " & txtCodigo.Text & ")"
      Set r = dbData.OpenRecordset(sSQL)
    
      If r.BOF Then
         If r.State <> 0 Then r.Close
         Set r = Nothing
         Exit Sub
      End If
      
      Campos_Brancos
      Mostrar_Dados r
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub

Private Sub txtContato_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFantasia_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtIE_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRazao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRazao_LostFocus()
   txtRazao.Text = TiraAcentos(txtRazao.Text)
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
