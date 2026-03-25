VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Contador_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORNECEDORES"
   ClientHeight    =   5475
   ClientLeft      =   -870
   ClientTop       =   435
   ClientWidth     =   12090
   Icon            =   "Contador_Cadastro.frx":0000
   LinkTopic       =   "Form49"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4035
      Left            =   60
      TabIndex        =   22
      Top             =   1080
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   7117
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
      TabPicture(0)   =   "Contador_Cadastro.frx":23D2
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
      Tab(0).Control(5)=   "frm_Principal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Contador_Cadastro.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdImprimir"
      Tab(1).Control(1)=   "Picture1"
      Tab(1).Control(2)=   "Grid"
      Tab(1).ControlCount=   3
      Begin ChamaleonBtn.chameleonButton cmdImprimir 
         Height          =   435
         Left            =   -65340
         TabIndex        =   47
         Top             =   3360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "Imprimir"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Contador_Cadastro.frx":240A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74880
         ScaleHeight     =   345
         ScaleWidth      =   3645
         TabIndex        =   35
         Top             =   3360
         Width           =   3675
         Begin VB.OptionButton optEstado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Estado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2700
            TabIndex        =   39
            Top             =   60
            Width           =   855
         End
         Begin VB.OptionButton optCidade 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1800
            TabIndex        =   38
            Top             =   60
            Width           =   855
         End
         Begin VB.OptionButton optRazao 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Razão"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   960
            TabIndex        =   37
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
            TabIndex        =   36
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
         Height          =   3315
         Left            =   180
         TabIndex        =   23
         Top             =   480
         Width           =   9435
         Begin VB.TextBox txtNum 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4320
            TabIndex        =   4
            Top             =   1260
            Width           =   555
         End
         Begin VB.ComboBox cboCidade 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Contador_Cadastro.frx":2426
            Left            =   840
            List            =   "Contador_Cadastro.frx":2428
            TabIndex        =   12
            Top             =   2580
            Width           =   2655
         End
         Begin VB.ComboBox cboEstado 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "Contador_Cadastro.frx":242A
            Left            =   180
            List            =   "Contador_Cadastro.frx":242C
            TabIndex        =   11
            Top             =   2580
            Width           =   615
         End
         Begin VB.TextBox txtCodUF 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   480
            TabIndex        =   42
            Top             =   2220
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtCodCid 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   2160
            TabIndex        =   41
            Top             =   2280
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtRazao 
            DataField       =   "CPF"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   600
            Width           =   7245
         End
         Begin VB.TextBox txtEndereco 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   1260
            Width           =   4155
         End
         Begin VB.TextBox txtComplemento 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   6900
            TabIndex        =   6
            Top             =   1260
            Width           =   1035
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   3120
            TabIndex        =   10
            Top             =   1920
            Width           =   6195
         End
         Begin VB.TextBox txtBairro 
            DataField       =   "Nickname"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   4920
            TabIndex        =   5
            Top             =   1260
            Width           =   1935
         End
         Begin VB.TextBox txtCRC 
            DataField       =   "Correio_eletronico"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   7380
            MaxLength       =   80
            TabIndex        =   2
            Top             =   600
            Width           =   1935
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
            TabIndex        =   14
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
         Begin MSMask.MaskEdBox txtCodigoIBGE 
            Height          =   315
            Left            =   3540
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   2580
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCPF 
            Height          =   315
            Left            =   7140
            TabIndex        =   15
            Top             =   2580
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "CPF"
            Height          =   195
            Left            =   7140
            TabIndex        =   48
            Top             =   2340
            Width           =   300
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Num."
            Height          =   195
            Left            =   4320
            TabIndex        =   46
            Top             =   1020
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Left            =   900
            TabIndex        =   45
            Top             =   2340
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cód. IBGE"
            Height          =   195
            Left            =   3540
            TabIndex        =   44
            Top             =   2340
            Width           =   750
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   180
            TabIndex        =   43
            Top             =   2340
            Width           =   255
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Contabilista:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label9 
            Caption         =   "Endereço"
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   1020
            Width           =   1575
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Complemento"
            Height          =   195
            Left            =   6900
            TabIndex        =   31
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Correio Eletrônico"
            Height          =   195
            Left            =   3120
            TabIndex        =   30
            Top             =   1680
            Width           =   1245
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   5340
            TabIndex        =   28
            Top             =   2340
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   4920
            TabIndex        =   27
            Top             =   1020
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Left            =   1620
            TabIndex        =   26
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   7980
            TabIndex        =   25
            Top             =   1020
            Width           =   285
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "CRC"
            Height          =   195
            Left            =   7440
            TabIndex        =   24
            Top             =   360
            Width           =   330
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   34
         Top             =   420
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5106
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   9660
         TabIndex        =   17
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
         MICON           =   "Contador_Cadastro.frx":242E
         PICN            =   "Contador_Cadastro.frx":244A
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
         TabIndex        =   18
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
         MICON           =   "Contador_Cadastro.frx":41DC
         PICN            =   "Contador_Cadastro.frx":41F8
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
         TabIndex        =   19
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
         MICON           =   "Contador_Cadastro.frx":5F8A
         PICN            =   "Contador_Cadastro.frx":5FA6
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
         TabIndex        =   16
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
         MICON           =   "Contador_Cadastro.frx":7D38
         PICN            =   "Contador_Cadastro.frx":7D54
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
         MICON           =   "Contador_Cadastro.frx":9AE6
         PICN            =   "Contador_Cadastro.frx":9B02
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
      TabIndex        =   20
      Top             =   120
      Width           =   11955
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONTABILIDADE"
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
         TabIndex        =   21
         Top             =   300
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   300
         Picture         =   "Contador_Cadastro.frx":B894
         Top             =   0
         Width           =   1200
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   40
      Top             =   5205
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
            TextSave        =   "21:55"
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
Attribute VB_Name = "Contador_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim printSQL As String
Dim sSQL As String
Dim r As ADODB.Recordset
Private Sub Campos_Brancos()
txtRazao.Text = ""
txtEndereco.Text = ""
txtComplemento.Text = ""
mskTelefone.Mask = ""
mskTelefone.Text = ""
cboCidade.Clear
cboCidade.Text = ""
cboEstado.Clear
cboEstado.Text = ""
mskCNPJ.Mask = ""
mskCNPJ.Text = ""
mskCPF.Mask = ""
mskCPF.Text = ""
txtEmail.Text = ""
mskFax.Mask = ""
mskFax.Text = ""
txtBairro.Text = ""
mskFax.Mask = ""
mskFax.Text = ""
mskCEP.Mask = ""
mskCEP.Text = ""
txtCRC.Text = ""
txtCodigoIBGE.Text = ""
txtNum.Text = ""
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
   If Not rTabela Is Nothing Then
      txtRazao.Text = ValidateNull(rTabela("NomeContabilista"))
      txtEndereco.Text = ValidateNull(rTabela("endereco"))
      txtComplemento.Text = ValidateNull(rTabela("Complemento"))
      mskTelefone.Text = ValidateNull(rTabela("Fone"))
      cboCidade.Text = ValidateNull(rTabela("cidade"))
      cboEstado.Text = ValidateNull(rTabela("UF"))
      mskCNPJ.Text = ValidateNull(rTabela("CNPJ"))
      mskCPF.Text = ValidateNull(rTabela("CPF"))
      txtEmail.Text = ValidateNull(rTabela("Email"))
      mskFax.Text = ValidateNull(rTabela("fax"))
      txtBairro.Text = ValidateNull(rTabela("bairro"))
      mskCEP.Text = ValidateNull(rTabela("cep"))
      txtCRC.Text = ValidateNull(rTabela("CRC"))
      txtCodigoIBGE = ValidateNull(rTabela("CodigoIBGE"))
      txtNum = ValidateNull(rTabela("Num"))
   End If
End Sub

Private Sub Mostrar_Contadores()
'INDICE PARA ORGANIZAR OS DADOS
Dim INDICE As String

If optRazao.Value = True Then
   INDICE = "NomeContabilista;"
ElseIf optCidade.Value = True Then
   INDICE = "cidade;"
ElseIf optEstado.Value = True Then
   INDICE = "UF;"
End If

sSQL = "SELECT NomeContabilista, CRC, cidade, UF, CNPJ, CPF, Fone FROM TbContabilista ORDER BY " & INDICE
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

printSQL = sSQL

If r.State <> 0 Then r.Close
Set r = Nothing
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

Private Sub cmdAlterar_Click()
If txtRazao.Text = "" Or txtCRC.Text = "" Or cboCidade.Text = "" Then
   ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão cadastrados.", vbInformation
   Exit Sub
End If

If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

frm_Principal.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
Campos_Brancos
Mostrar_Contadores
End Sub

Private Function Inserir_Dados() As Boolean
Dim sSQL As String

If Trim(txtCodigoIBGE.Text) = "" Then txtCodigoIBGE.Text = "0"
If Trim(txtNum.Text) = "" Then txtNum.Text = "0"
  
  'Comando de inclusão
  sSQL = "INSERT INTO TbContabilista (" & _
     "codigoibge, endereco, Complemento, Fone, cidade, UF, CNPJ, CPF,  " & _
     "Email, fax, bairro, cep, CRC, Num, NomeContabilista) VALUES ("
  
  sSQL = sSQL & _
     txtCodigoIBGE.Text & ", '" & txtEndereco.Text & "', '" & txtComplemento.Text & "', '" & _
     mskTelefone.Text & "', '" & cboCidade.Text & "', '" & cboEstado.Text & "', '" & mskCNPJ.Text & "', '" & mskCPF.Text & "', '" & _
     txtEmail.Text & "', '" & mskFax.Text & "', '" & txtBairro.Text & "', '" & mskCEP.Text & "', '" & txtCRC.Text & "', '" & _
     txtNum.Text & "', '" & txtRazao.Text & "')"
  
  Debug.Print sSQL
  'Retorna o resultado da atualização
  Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
If Trim(txtCodigoIBGE.Text) = "" Then txtCodigoIBGE.Text = "0"
If Trim(txtNum.Text) = "" Then txtNum.Text = "0"
  
'Comando de atualização
sSQL = "UPDATE TbContabilista SET " & _
   "NomeContabilista = '" & txtRazao.Text & "', " & _
   "endereco = '" & txtEndereco.Text & "', " & _
   "Complemento = '" & txtComplemento.Text & "', " & _
   "Fone = '" & mskTelefone.Text & "', " & _
   "cidade = '" & cboCidade.Text & "', " & _
   "UF = '" & cboEstado.Text & "', " & _
   "CNPJ = '" & mskCNPJ.Text & "', " & _
   "CPF = '" & mskCPF.Text & "', " & _
   "Email = '" & txtEmail.Text & "', " & _
   "fax = '" & mskFax.Text & "', " & _
   "bairro = '" & txtBairro.Text & "', " & _
   "cep = '" & mskCEP.Text & "', " & _
   "CRC = '" & txtCRC.Text & "', " & _
   "Num = " & txtNum.Text & ", codigoibge = " & txtCodigoIBGE.Text & " " & _
   "WHERE (CNPJ = '" & Me.mskCNPJ.Text & "');"

'Retorna o resultado da atualização
Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdCancelar_Click()
Campos_Brancos
frm_Principal.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
End Sub

Private Sub cmdExcluir_Click()
Dim bRet As Boolean

If txtRazao.Text = "" Or txtCRC.Text = "" Or cboCidade.Text = "" Then
   ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão cadastrados.", vbInformation
   Exit Sub
End If

If ShowMsg("Deseja excluir esse Contador?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

sSQL = "DELETE FROM TbContabilista WHERE (NomeContabilista = '" & txtRazao.Text & "');"
bRet = dbData.Execute(sSQL)
 
If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

frm_Principal.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
Campos_Brancos
Mostrar_Contadores
End Sub

Private Sub cmdImprimir_Click()
'Me.Hide
'Set Rel_Fornecedores.Relatorio.Recordset = Data4.Recordset
'Rel_Fornecedores.dfQuant.Caption = "Quant. de Registro(s): " & lblQuant.Caption
'Rel_Fornecedores.Relatorio.Ativar
Unload Rel_Fornecedores
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

Set Rel_Fornecedores.Relatorio.Recordset = r
'Rel_Fornecedores.dfQuant.Caption = lblTotalNota.Caption
'Rel_Fornecedores.dfBruto.Caption = lblSomaNota.Caption

'If cboFiltroNota.Text = "MENSAL" Then
'   Rel_Fornecedores.dfTipo.Caption = "Tipo: Mês = " & cboConNotaMes.Text & "/" & cboConNotaAno.Text
'ElseIf cboFiltroNota.Text = "DATAS" Then
'   Rel_Fornecedores.dfTipo.Caption = "Tipo: Datas = " & mskConNotaInicial.Text & " à " & mskConNotaFinal.Text
'ElseIf cboFiltroNota.Text = "TbContabilista" Then
'   Rel_Fornecedores.dfTipo.Caption = "Tipo: TbContabilista = " & cboConNotaCliente.Text & ""
'ElseIf cboFiltroNota.Text = "NOTA FISCAL" Then
'   Rel_Fornecedores.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & txtConNotaNumNota.Text & ""
'Else
'   Rel_Fornecedores.dfTipo.Caption = "Tipo: Todas as notas"
'End If

Rel_Fornecedores.Relatorio.Ativar
Unload Rel_Fornecedores
Me.Show 1
End Sub

Private Sub cmdNovo_Click()
frm_Principal.Enabled = True
Campos_Brancos
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = False
txtRazao.SetFocus
End Sub



Private Sub Grid_DblClick()
SSTab1.Tab = 0
frm_Principal.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
cmdNovo.Enabled = True

If cmdSalvar.Enabled = False Then
   sSQL = "SELECT * FROM TbContabilista WHERE (NomeContabilista LIKE '%" & (Grid.TextMatrix(Grid.Row, 1)) & "%') "
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

Private Sub mskCPF_KeyPress(KeyAscii As Integer)
mskCPF.Mask = "###.###.###-##"
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
If txtRazao.Text = "" Or txtCRC.Text = "" Or cboCidade.Text = "" Then
   ShowMsg "FORMULÁRIO INCOMPLETO!", vbExclamation
   txtRazao.SetFocus
   Exit Sub
End If

If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Campos_Brancos
frm_Principal.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
Mostrar_Contadores
End Sub


Private Sub Form_Load()
SSTab1.Tab = 0
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
Mostrar_Contadores
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
Set moCombo = New cComboHelper
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer, j As Integer
Dim x As Integer

With Grid
   .Clear
   .Cols = 5
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 3100
   .ColWidth(2) = 2800
   .ColWidth(3) = 1300
   .ColWidth(4) = 500

   .TextMatrix(0, 1) = "RAZÃO SOCIAL"
   .TextMatrix(0, 2) = "CRC"
   .TextMatrix(0, 3) = "CIDADE"
   .TextMatrix(0, 4) = "UF"
   
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
         'For j = 1 To .Rows - 1
         '   .Row = j
         '   .Col = 6
         '   .CellBackColor = &HC0FFFF
         'Next
         
         'ALINHAMENTO
         .ColAlignment(1) = 1
         .TextMatrix(.rows - 1, 1) = ValidateNull(rTabela("NomeContabilista"))
         .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("CRC"))
         .TextMatrix(.rows - 1, 3) = rTabela("cidade")
         .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("UF"))
         
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .rows = .rows - 1
   Grid.Redraw = True
End With
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
   Mostrar_Contadores
End Sub

Private Sub optEstado_Click()
   Mostrar_Contadores
End Sub

Private Sub optRazao_Click()
   Mostrar_Contadores
End Sub

Private Sub txtBairro_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBairro_LostFocus()
txtBairro.Text = UCase(txtBairro.Text)
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

Private Sub txtComplemento_LostFocus()
txtComplemento.Text = UCase(txtComplemento.Text)
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtcrc_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEndereco_LostFocus()
txtEndereco.Text = UCase(txtEndereco.Text)
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtRazao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRazao_LostFocus()
txtRazao.Text = TiraAcentos(txtRazao.Text)
txtRazao.Text = UCase(txtRazao.Text)
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
