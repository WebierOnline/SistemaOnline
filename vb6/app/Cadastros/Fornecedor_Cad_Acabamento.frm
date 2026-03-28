VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Fornecedor_Cad_Acabamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORNECEDORES"
   ClientHeight    =   7785
   ClientLeft      =   -870
   ClientTop       =   435
   ClientWidth     =   12090
   Icon            =   "Fornecedor_Cad_Acabamento.frx":0000
   LinkTopic       =   "Form49"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   60
      TabIndex        =   29
      Top             =   1080
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
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
      TabPicture(0)   =   "Fornecedor_Cad_Acabamento.frx":23D2
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
      Tab(0).Control(5)=   "frm_Secundario"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frm_Principal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Fornecedor_Cad_Acabamento.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(1)=   "Grid"
      Tab(1).ControlCount=   2
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74880
         ScaleHeight     =   345
         ScaleWidth      =   11625
         TabIndex        =   52
         Top             =   5880
         Width           =   11655
         Begin VB.OptionButton optEstado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Estado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3720
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
         Height          =   4755
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   9435
         Begin VB.TextBox txtNum 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   3540
            TabIndex        =   4
            Top             =   1260
            Width           =   555
         End
         Begin VB.ComboBox cboCidade 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Fornecedor_Cad_Acabamento.frx":240A
            Left            =   840
            List            =   "Fornecedor_Cad_Acabamento.frx":240C
            TabIndex        =   13
            Top             =   2580
            Width           =   2655
         End
         Begin VB.ComboBox cboEstado 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Fornecedor_Cad_Acabamento.frx":240E
            Left            =   180
            List            =   "Fornecedor_Cad_Acabamento.frx":2410
            TabIndex        =   12
            Top             =   2580
            Width           =   615
         End
         Begin VB.TextBox txtCodUF 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   480
            TabIndex        =   60
            Top             =   2220
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtCodCid 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   2160
            TabIndex        =   59
            Top             =   2280
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtIE 
            Height          =   315
            Left            =   7140
            TabIndex        =   16
            Top             =   2580
            Width           =   2115
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
            TabIndex        =   1
            Top             =   600
            Width           =   4425
         End
         Begin VB.TextBox txtEndereco 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   1260
            Width           =   3375
         End
         Begin VB.TextBox txtReferencia 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   5880
            TabIndex        =   6
            Top             =   1260
            Width           =   2055
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   4620
            TabIndex        =   11
            Top             =   1920
            Width           =   4635
         End
         Begin VB.TextBox txtBairro 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4140
            TabIndex        =   5
            Top             =   1260
            Width           =   1695
         End
         Begin VB.TextBox txtFantasia 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4620
            MaxLength       =   80
            TabIndex        =   2
            Top             =   600
            Width           =   4695
         End
         Begin VB.TextBox txtContato 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   3360
            Width           =   9135
         End
         Begin MSMask.MaskEdBox mskCEP 
            Height          =   315
            Left            =   7980
            TabIndex        =   7
            Top             =   1260
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCNPJ 
            Height          =   315
            Left            =   5340
            TabIndex        =   15
            Top             =   2580
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
            Top             =   1920
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
            Top             =   1920
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
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtObs 
            Height          =   675
            Left            =   120
            TabIndex        =   18
            Top             =   3960
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   1191
            _Version        =   393216
            MaxLength       =   40
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCodigoIBGE 
            Height          =   315
            Left            =   3540
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2580
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Num."
            Height          =   195
            Left            =   3540
            TabIndex        =   64
            Top             =   1020
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Cidade"
            Height          =   255
            Left            =   900
            TabIndex        =   63
            Top             =   2340
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Cód. IBGE"
            Height          =   255
            Left            =   3540
            TabIndex        =   62
            Top             =   2340
            Width           =   1215
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   180
            TabIndex        =   61
            Top             =   2340
            Width           =   255
         End
         Begin VB.Label Label17 
            Caption         =   "IE"
            Height          =   255
            Left            =   7140
            TabIndex        =   50
            Top             =   2340
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razão Social*"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label9 
            Caption         =   "Endereço"
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Top             =   1020
            Width           =   1575
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Ponto de Referência"
            Height          =   195
            Left            =   5880
            TabIndex        =   47
            Top             =   1020
            Width           =   1470
         End
         Begin VB.Label Label11 
            Caption         =   "Correio Eletrônico"
            Height          =   285
            Left            =   4620
            TabIndex        =   46
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Telefone:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "CNPJ"
            Height          =   255
            Left            =   5340
            TabIndex        =   44
            Top             =   2340
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   4140
            TabIndex        =   43
            Top             =   1020
            Width           =   405
         End
         Begin VB.Label Label2 
            Caption         =   "Fax"
            Height          =   255
            Left            =   1620
            TabIndex        =   42
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Celular"
            Height          =   255
            Left            =   3120
            TabIndex        =   41
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   7980
            TabIndex        =   40
            Top             =   1020
            Width           =   285
         End
         Begin VB.Label Label22 
            Caption         =   "Observação"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Fantasia*"
            Height          =   285
            Left            =   4620
            TabIndex        =   38
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Contato"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   3120
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
         Left            =   120
         TabIndex        =   30
         Top             =   5280
         Width           =   9435
         Begin VB.TextBox txtConta 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   6120
            TabIndex        =   22
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtAgencia 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4440
            TabIndex        =   21
            Top             =   480
            Width           =   1635
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   2400
            TabIndex        =   20
            Top             =   480
            Width           =   1995
         End
         Begin VB.ComboBox cboTipoConta 
            Height          =   315
            Left            =   120
            TabIndex        =   19
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
         Height          =   5355
         Left            =   -74880
         TabIndex        =   51
         Top             =   420
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   9446
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   9600
         TabIndex        =   24
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
         MICON           =   "Fornecedor_Cad_Acabamento.frx":2412
         PICN            =   "Fornecedor_Cad_Acabamento.frx":242E
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
         Left            =   9600
         TabIndex        =   25
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
         MICON           =   "Fornecedor_Cad_Acabamento.frx":41C0
         PICN            =   "Fornecedor_Cad_Acabamento.frx":41DC
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
         Left            =   9600
         TabIndex        =   26
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
         MICON           =   "Fornecedor_Cad_Acabamento.frx":5F6E
         PICN            =   "Fornecedor_Cad_Acabamento.frx":5F8A
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
         Left            =   9600
         TabIndex        =   23
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
         MICON           =   "Fornecedor_Cad_Acabamento.frx":7D1C
         PICN            =   "Fornecedor_Cad_Acabamento.frx":7D38
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
         Left            =   9600
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
         MICON           =   "Fornecedor_Cad_Acabamento.frx":9ACA
         PICN            =   "Fornecedor_Cad_Acabamento.frx":9AE6
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
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   11925
      TabIndex        =   27
      Top             =   120
      Width           =   11955
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FORNECEDORES - ACABAMENTOS"
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
         Left            =   1620
         TabIndex        =   28
         Top             =   300
         Width           =   5355
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   300
         Picture         =   "Fornecedor_Cad_Acabamento.frx":B878
         Top             =   0
         Width           =   1200
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   58
      Top             =   7515
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16986
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "23:11"
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
Attribute VB_Name = "Fornecedor_Cad_Acabamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper

Private Sub Campos_Brancos()
If cmdAlterar.Enabled = False Then txtCodigo.Text = ""
txtRazao.Text = ""
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
txtCodigoIBGE.Text = ""
txtNum.Text = ""
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
   If Not rTabela Is Nothing Then
      txtCodigo.Text = ValidateNull(rTabela("codigo"))
      txtRazao.Text = ValidateNull(rTabela("razao"))
      txtEndereco.Text = ValidateNull(rTabela("endereco"))
      txtReferencia.Text = ValidateNull(rTabela("ponto_de_referencia"))
      mskTelefone.Text = ValidateNull(rTabela("telefone"))
      cboCidade.Text = ValidateNull(rTabela("cidade"))
      cboEstado.Text = ValidateNull(rTabela("estado"))
      mskCNPJ.Text = ValidateNull(rTabela("cpf"))
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
      txtCodigoIBGE = ValidateNull(rTabela("CodigoIBGE"))
      txtNum = ValidateNull(rTabela("Numero"))
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
   
   sSQL = "SELECT codigo, razao, fantasia, cidade, estado FROM fornecedor_acabamentos ORDER BY " & INDICE
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
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboCidade_Click()
cboCidade_LostFocus
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
Private Sub cboEstado_LostFocus()
On Error GoTo TrataErro

If cboEstado.Text = "" Then txtCodUF.Text = "": Exit Sub
If cboEstado.ListIndex = -1 Then txtCodUF.Text = "": Exit Sub

txtCodUF = cboEstado.ItemData(cboEstado.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub chameleonButton1_Click()

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
   'sSQL = "SELECT * FROM fornecedor_acabamentos WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   frm_Principal.Enabled = False
   frm_Secundario.Enabled = False
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
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

 If Trim(txtCodigoIBGE.Text) = "" Then txtCodigoIBGE.Text = "0"
 If Trim(txtNum.Text) = "" Then txtNum.Text = "0"
   
   'Comando de inclusão
   sSQL = "INSERT INTO fornecedor_acabamentos (" & _
      "codigo, razao, endereco, ponto_de_referencia, telefone, cidade, estado, cpf, " & _
      "ie, correio_eletronico, fax, celular, bairro, cep, contato, fantasia, obs, " & _
      "banco, agencia, conta, tipo_conta, numero, codigoibge) VALUES ("
   
   sSQL = sSQL & _
      txtCodigo.Text & ", '" & txtRazao.Text & "', '" & txtEndereco.Text & "', '" & txtReferencia.Text & "', '" & _
      mskTelefone.Text & "', '" & cboCidade.Text & "', '" & cboEstado.Text & "', '" & mskCNPJ.Text & "', '" & _
      txtIE.Text & "', '" & txtEmail.Text & "', '" & mskFax.Text & "', '" & mskCelular.Text & "', '" & _
      txtBairro.Text & "', '" & mskCEP.Text & "', '" & txtContato.Text & "', '" & txtFantasia.Text & "', '" & _
      txtObs.Text & "', '" & cboBanco.Text & "', '" & txtAgencia.Text & "', '" & txtConta.Text & "', '" & cboTipoConta.Text & "', " & txtNum.Text & ", " & txtCodigoIBGE.Text & ")"
   
   'Retorna o resultado da atualização
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
Dim sSQL As String
 If Trim(txtCodigoIBGE.Text) = "" Then txtCodigoIBGE.Text = "0"
 If Trim(txtNum.Text) = "" Then txtNum.Text = "0"
   
   'Comando de atualização
   sSQL = "UPDATE fornecedor_acabamentos SET " & _
      "razao = '" & txtRazao.Text & "', " & _
      "endereco = '" & txtEndereco.Text & "', " & _
      "ponto_de_referencia = '" & txtReferencia.Text & "', " & _
      "telefone = '" & mskTelefone.Text & "', " & _
      "cidade = '" & cboCidade.Text & "', " & _
      "estado = '" & cboEstado.Text & "', " & _
      "cpf = '" & mskCNPJ.Text & "', " & _
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
      "tipo_conta = '" & cboTipoConta.Text & "', numero = " & txtNum.Text & ", codigoibge = " & txtCodigoIBGE.Text & " " & _
      "WHERE (codigo = " & Me.txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdCancelar_Click()
    Campos_Brancos
    frm_Principal.Enabled = False
    frm_Secundario.Enabled = False
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
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
   If ShowMsg("Deseja excluir esse fornecedor?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Não é necessário consulta o registro antes de exclui-lo
   'sSQL = "SELECT * FROM fornecedor_acabamentos WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM fornecedor_acabamentos WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
    
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   frm_Principal.Enabled = False
   frm_Secundario.Enabled = False
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cmdNovo.Enabled = True
   Campos_Brancos
   Mostrar_Fornecedores
End Sub

Private Sub cmdNovo_Click()
frm_Principal.Enabled = True
frm_Secundario.Enabled = True
Campos_Brancos
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = False
Auto_Numeracao
txtRazao.SetFocus
End Sub



Private Sub Grid_DblClick()
SSTab1.Tab = 0
frm_Principal.Enabled = True
frm_Secundario.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
cmdNovo.Enabled = True
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
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
   If mskCEP.Text = "__.___-__" Then
      mskCEP.Mask = ""
      mskCEP.Text = ""
   End If
End Sub

Private Sub mskCNPJ_KeyPress(KeyAscii As Integer)
   mskCNPJ.Mask = "##.###.###/####-##"
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

Private Sub cmdSalvar_Click()
   'On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtRazao.Text = "" Or txtFantasia.Text = "" Or cboCidade.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!", vbExclamation
      txtRazao.SetFocus
      Exit Sub
   End If
   
   'Não é necessário consultar todos os registros antes de inserir um novo
   'sSQL = "SELECT * FROM fornecedor_acabamentos;"
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
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
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
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_fornecedor FROM fornecedor_acabamentos;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigo.Text = r("cod_fornecedor") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
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
      .Rows = 2
      
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
            For j = 1 To .Rows - 1
               .Row = j
               .Col = 6
               .CellBackColor = &HC0FFFF
            Next
            
            'ALINHAMENTO
            .ColAlignment(2) = 1
            
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = rTabela("razao")
            .TextMatrix(.Rows - 1, 3) = rTabela("fantasia")
            .TextMatrix(.Rows - 1, 4) = rTabela("cidade")
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("estado"))
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
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
      
      sSQL = "SELECT * FROM fornecedor_acabamentos WHERE (codigo = " & txtCodigo.Text & ")"
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
