VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRODUTOS"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14250
   Icon            =   "Produtos_Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   14085
      TabIndex        =   50
      Top             =   60
      Width           =   14115
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   12780
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   180
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Produtos_Cadastro.frx":23D2
         Top             =   30
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
         TabIndex        =   51
         Top             =   240
         Width           =   1770
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   52
      Top             =   9900
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20796
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21:38"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   60
      TabIndex        =   54
      Top             =   840
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabPicture(0)   =   "Produtos_Cadastro.frx":7DA5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSair"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdNovo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSalvar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdExcluir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAlterar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancelar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtUltCompra"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frmPrecos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frmDados"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "frmFiscal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "frmComp"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtTam"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "frmReferencia"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "frmGas"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkFracionado"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "frmFracionado"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Produtos_Cadastro.frx":7DC1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label25"
      Tab(1).Control(1)=   "cmdDesativar"
      Tab(1).Control(2)=   "cmdApagar"
      Tab(1).Control(3)=   "cmdEditar"
      Tab(1).Control(4)=   "ccmdDuplicar"
      Tab(1).Control(5)=   "cmdExibir"
      Tab(1).Control(6)=   "cmdImprimir"
      Tab(1).Control(7)=   "Grid"
      Tab(1).Control(8)=   "frmOrdemComum"
      Tab(1).Control(9)=   "frmCriterios"
      Tab(1).Control(10)=   "frmVenda"
      Tab(1).Control(11)=   "frmFiltroComum"
      Tab(1).Control(12)=   "frmSituacao"
      Tab(1).Control(13)=   "frmFiltro"
      Tab(1).Control(14)=   "Frame6"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "QUANTIDADES"
      TabPicture(2)   =   "Produtos_Cadastro.frx":7DDD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5(0)"
      Tab(2).Control(1)=   "Grid_Quant"
      Tab(2).Control(2)=   "Label40"
      Tab(2).Control(3)=   "lblEstoqueHoje"
      Tab(2).Control(4)=   "lblNomeProduto1"
      Tab(2).Control(5)=   "lblQuantAdicao"
      Tab(2).Control(6)=   "lblQuantRemocao"
      Tab(2).Control(7)=   "Label27"
      Tab(2).Control(8)=   "Label28"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "PREÇOS"
      TabPicture(3)   =   "Produtos_Cadastro.frx":7DF9
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5(1)"
      Tab(3).Control(1)=   "GridPrecos"
      Tab(3).Control(2)=   "lblNomeProduto2"
      Tab(3).ControlCount=   3
      Begin VB.Frame frmFracionado 
         Caption         =   "Relacionamento de Produtos Fracionados"
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
         Left            =   0
         TabIndex        =   201
         Top             =   6360
         Visible         =   0   'False
         Width           =   11535
         Begin VB.ComboBox cboProdutoFracionado 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   1140
            TabIndex        =   44
            Top             =   540
            Width           =   6615
         End
         Begin VB.TextBox txtCodProdFracionado 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtQuantFracionado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   7800
            TabIndex        =   45
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. Prod."
            Height          =   195
            Left            =   120
            TabIndex        =   204
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. por Embalagem"
            Height          =   195
            Left            =   7800
            TabIndex        =   203
            Top             =   300
            Width           =   1620
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição do Produto"
            Height          =   195
            Left            =   1140
            TabIndex        =   202
            Top             =   300
            Width           =   1545
         End
      End
      Begin VB.CheckBox chkFracionado 
         Caption         =   "Este produto será fracionado?"
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
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   5220
         Width           =   4155
      End
      Begin VB.Frame frmGas 
         Caption         =   "Combustíveis"
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
         TabIndex        =   187
         Top             =   5460
         Visible         =   0   'False
         Width           =   11535
         Begin VB.TextBox txtCODIF 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10620
            TabIndex        =   42
            Top             =   540
            Width           =   795
         End
         Begin VB.TextBox txtdescricaoANP 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1140
            MaxLength       =   8
            TabIndex        =   36
            Top             =   540
            Width           =   3975
         End
         Begin VB.TextBox txtpGLP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5160
            TabIndex        =   37
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtcProdANP 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtpMixGN 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8280
            TabIndex        =   40
            Top             =   540
            Width           =   1035
         End
         Begin VB.TextBox txtpGNn 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6180
            TabIndex        =   38
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtpGNi 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7200
            TabIndex        =   39
            Top             =   540
            Width           =   1035
         End
         Begin VB.TextBox txtValorPartida 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9360
            TabIndex        =   41
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CODIF"
            Height          =   195
            Left            =   10620
            TabIndex        =   195
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição ANP"
            Height          =   195
            Left            =   1140
            TabIndex        =   194
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GLP (%)"
            Height          =   195
            Left            =   5160
            TabIndex        =   193
            Top             =   300
            Width           =   570
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. ANP"
            Height          =   195
            Left            =   120
            TabIndex        =   192
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gás Mix (%)"
            Height          =   195
            Left            =   8280
            TabIndex        =   191
            Top             =   300
            Width           =   825
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gás Nat. (%)"
            Height          =   195
            Left            =   6180
            TabIndex        =   190
            Top             =   300
            Width           =   885
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gás Import.(%)"
            Height          =   195
            Left            =   7200
            TabIndex        =   189
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor de Partida"
            Height          =   195
            Left            =   9360
            TabIndex        =   188
            Top             =   300
            Width           =   1125
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Preço"
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
         Left            =   -73140
         TabIndex        =   179
         Top             =   6900
         Width           =   1275
         Begin VB.OptionButton optTodosPreco 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   182
            Top             =   720
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optComPreco 
            Caption         =   "Com Preço"
            Height          =   195
            Left            =   120
            TabIndex        =   181
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optSemPreco 
            Caption         =   "Sem Preço"
            Height          =   195
            Left            =   120
            TabIndex        =   180
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame frmReferencia 
         Caption         =   "Referência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   6780
         TabIndex        =   170
         Top             =   5220
         Width           =   4875
         Begin VB.TextBox txtReferencia 
            Height          =   315
            Left            =   60
            TabIndex        =   176
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtCodSituacao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   171
            Top             =   180
            Visible         =   0   'False
            Width           =   675
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Referencia 
            Height          =   1755
            Left            =   2640
            TabIndex        =   172
            Top             =   180
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   3096
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdRemoverReferencia 
            Height          =   315
            Left            =   2640
            TabIndex        =   173
            Top             =   3300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "R&emover"
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
            MICON           =   "Produtos_Cadastro.frx":7E15
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionarReferencia 
            Height          =   315
            Left            =   60
            TabIndex        =   174
            Top             =   900
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "A&dicionar"
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
            MICON           =   "Produtos_Cadastro.frx":7E31
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Referencia_Desc 
            Height          =   1275
            Left            =   2640
            TabIndex        =   177
            Top             =   1980
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   2249
            _Version        =   393216
            BackColor       =   12648447
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdDescReferencia 
            Height          =   315
            Left            =   3660
            TabIndex        =   178
            Top             =   3300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Descontinuar"
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
            MICON           =   "Produtos_Cadastro.frx":7E4D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência"
            Height          =   195
            Left            =   60
            TabIndex        =   175
            Top             =   240
            Width           =   780
         End
      End
      Begin VB.TextBox txtTam 
         Height          =   315
         Left            =   12420
         MaxLength       =   20
         TabIndex        =   153
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frmFiltro 
         Caption         =   "Quantidade"
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
         TabIndex        =   144
         Top             =   6900
         Width           =   1695
         Begin VB.OptionButton optMostrarTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   148
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optMostrarZerados 
            Caption         =   "Zerados"
            Height          =   195
            Left            =   120
            TabIndex        =   147
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optMostrarNegativos 
            Caption         =   "Negativos"
            Height          =   195
            Left            =   120
            TabIndex        =   146
            Top             =   540
            Width           =   1155
         End
         Begin VB.OptionButton optMostrarQuant 
            Caption         =   "Com quantidade"
            Height          =   195
            Left            =   120
            TabIndex        =   145
            Top             =   180
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame frmSituacao 
         Caption         =   "Situação"
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
         Left            =   -71820
         TabIndex        =   139
         Top             =   6900
         Width           =   1455
         Begin VB.OptionButton optDesabilitados 
            Caption         =   "Desabilitados"
            Height          =   195
            Left            =   120
            TabIndex        =   141
            Top             =   480
            Width           =   1275
         End
         Begin VB.OptionButton optHabilitado 
            Caption         =   "Habilitados"
            Height          =   195
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame frmFiltroComum 
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
         Height          =   915
         Left            =   -69660
         TabIndex        =   125
         Top             =   7920
         Visible         =   0   'False
         Width           =   5715
         Begin VB.OptionButton optPorPalavra 
            Caption         =   "Palavra"
            Height          =   195
            Left            =   3180
            TabIndex        =   186
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton optPalavrasDuplas 
            Caption         =   "Palavras Duplas"
            Height          =   195
            Left            =   4140
            TabIndex        =   185
            Top             =   240
            Width           =   1515
         End
         Begin VB.OptionButton optPorIniciais 
            Caption         =   "Iniciais"
            Height          =   195
            Left            =   2340
            TabIndex        =   184
            Top             =   240
            Width           =   795
         End
         Begin VB.OptionButton optCompleto 
            Caption         =   "Completo"
            Height          =   195
            Left            =   1320
            TabIndex        =   183
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.ComboBox cboConsProduto 
            Height          =   315
            Left            =   120
            TabIndex        =   126
            Top             =   480
            Visible         =   0   'False
            Width           =   5475
         End
         Begin VB.Label lblNomeCombo 
            Caption         =   "Nome"
            Height          =   195
            Left            =   120
            TabIndex        =   127
            Top             =   240
            Visible         =   0   'False
            Width           =   1515
         End
      End
      Begin VB.Frame frmVenda 
         Caption         =   "Totais"
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
         Height          =   975
         Left            =   -70320
         TabIndex        =   117
         Top             =   6900
         Width           =   9315
         Begin ChamaleonBtn.chameleonButton cmd 
            Height          =   675
            Left            =   7800
            TabIndex        =   199
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   1191
            BTYPE           =   3
            TX              =   "Gerar Arquivo"
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
            MICON           =   "Produtos_Cadastro.frx":7E69
            PICN            =   "Produtos_Cadastro.frx":7E85
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Custo/Total:"
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
            Left            =   3600
            TabIndex        =   143
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lblValorTotalCusto 
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
            Left            =   4860
            TabIndex        =   142
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Quant. de produtos:"
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
            Left            =   100
            TabIndex        =   123
            Top             =   480
            Width           =   1710
         End
         Begin VB.Label lblProdutos 
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
            Left            =   1920
            TabIndex        =   122
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label lblValorTotal 
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
            Left            =   4860
            TabIndex        =   121
            Top             =   480
            Width           =   1545
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Venda/Total:"
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
            Left            =   3600
            TabIndex        =   120
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label lblTipos 
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
            Left            =   1920
            TabIndex        =   119
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Quant. de tipos:"
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
            Left            =   100
            TabIndex        =   118
            Top             =   240
            Width           =   1380
         End
      End
      Begin VB.Frame frmComp 
         Caption         =   "Compartibilidade"
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
         Height          =   2835
         Left            =   120
         TabIndex        =   109
         Top             =   5460
         Width           =   5955
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   2115
            Left            =   60
            ScaleHeight     =   2085
            ScaleWidth      =   2145
            TabIndex        =   110
            Top             =   240
            Width           =   2175
            Begin VB.ComboBox cboFab 
               Height          =   315
               Left            =   60
               TabIndex        =   163
               Top             =   360
               Width           =   1995
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1380
               TabIndex        =   162
               Top             =   60
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox txtCodComp 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1380
               TabIndex        =   111
               Top             =   720
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.ComboBox cboModelo 
               Height          =   315
               Left            =   60
               TabIndex        =   165
               Top             =   1020
               Width           =   1695
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   60
               TabIndex        =   167
               Top             =   1620
               Width           =   1155
            End
            Begin ChamaleonBtn.chameleonButton cmdAddModelo 
               Height          =   315
               Left            =   1750
               TabIndex        =   166
               ToolTipText     =   "Salvar um novo modelo."
               Top             =   1020
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "+"
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
               MICON           =   "Produtos_Cadastro.frx":9C17
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fabricante"
               Height          =   195
               Left            =   60
               TabIndex        =   164
               Top             =   120
               Width           =   750
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Modelo"
               Height          =   195
               Left            =   60
               TabIndex        =   113
               Top             =   780
               Width           =   525
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano"
               Height          =   195
               Left            =   60
               TabIndex        =   112
               Top             =   1380
               Width           =   285
            End
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Comp 
            Height          =   2115
            Left            =   2280
            TabIndex        =   114
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   3731
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdRemoverComp 
            Height          =   315
            Left            =   1500
            TabIndex        =   115
            Top             =   2400
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "R&emover"
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
            MICON           =   "Produtos_Cadastro.frx":9C33
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionarComp 
            Height          =   315
            Left            =   60
            TabIndex        =   116
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "A&dicionar"
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
            MICON           =   "Produtos_Cadastro.frx":9C4F
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
      Begin VB.Frame frmFiscal 
         Caption         =   "Parametros Fiscais"
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
         Left            =   120
         TabIndex        =   100
         Top             =   3840
         Width           =   11535
         Begin VB.ComboBox cboCST 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   540
            Width           =   855
         End
         Begin VB.ComboBox cboCFOP 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   540
            Width           =   915
         End
         Begin VB.TextBox txtRedBCAliquota 
            Height          =   315
            Left            =   9060
            TabIndex        =   33
            Top             =   540
            Width           =   1035
         End
         Begin VB.TextBox txtIPIAliquota 
            Height          =   315
            Left            =   8160
            TabIndex        =   32
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtCofinsAliquota 
            Height          =   315
            Left            =   6420
            TabIndex        =   30
            Top             =   540
            Width           =   1035
         End
         Begin VB.TextBox txtPisAliquota 
            Height          =   315
            Left            =   4680
            TabIndex        =   28
            Top             =   540
            Width           =   675
         End
         Begin VB.TextBox txtPISCST 
            Height          =   315
            Left            =   3960
            TabIndex        =   27
            Top             =   540
            Width           =   675
         End
         Begin VB.TextBox txtCOFINSCST 
            Height          =   315
            Left            =   5400
            TabIndex        =   29
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtIPICST 
            Height          =   315
            Left            =   7500
            TabIndex        =   31
            Top             =   540
            Width           =   615
         End
         Begin VB.TextBox txtNCM 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txtCEST 
            Height          =   315
            Left            =   10140
            TabIndex        =   34
            Top             =   540
            Width           =   915
         End
         Begin VB.TextBox txtICMSAliquota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3060
            TabIndex        =   26
            Top             =   540
            Width           =   855
         End
         Begin ChamaleonBtn.chameleonButton cmdBuscarCEST 
            Height          =   315
            Left            =   9000
            TabIndex        =   159
            Top             =   900
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Consultar CEST pelo NCM"
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
            MICON           =   "Produtos_Cadastro.frx":9C6B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdConsultarNCMean 
            Height          =   315
            Left            =   120
            TabIndex        =   160
            Top             =   900
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Consultar NCM pelo EAN"
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
            MICON           =   "Produtos_Cadastro.frx":9C87
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdConsultarNCM 
            Height          =   315
            Left            =   2340
            TabIndex        =   161
            Top             =   900
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Consultar NCM pela Descrição"
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
            MICON           =   "Produtos_Cadastro.frx":9CA3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Red. BC Aliq."
            Height          =   195
            Left            =   9060
            TabIndex        =   209
            Top             =   300
            Width           =   945
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IPI Aliq."
            Height          =   195
            Left            =   8160
            TabIndex        =   158
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COFINS Aliq."
            Height          =   195
            Left            =   6420
            TabIndex        =   157
            Top             =   300
            Width           =   930
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PIS Aliq."
            Height          =   195
            Left            =   4680
            TabIndex        =   156
            Top             =   300
            Width           =   600
         End
         Begin VB.Label lblPISCST 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PIS CST"
            Height          =   195
            Left            =   3960
            TabIndex        =   108
            Top             =   300
            Width           =   615
         End
         Begin VB.Label lblCOFINSCST 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COFINS CST"
            Height          =   195
            Left            =   5400
            TabIndex        =   107
            Top             =   300
            Width           =   945
         End
         Begin VB.Label lblIPICST 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IPI CST"
            Height          =   195
            Left            =   7500
            TabIndex        =   106
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblNCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NCM"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblCEST 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CEST"
            Height          =   195
            Left            =   10140
            TabIndex        =   104
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lblICMSAliq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICMS Aliq."
            Height          =   195
            Left            =   3060
            TabIndex        =   103
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CFOP"
            Height          =   195
            Left            =   1260
            TabIndex        =   102
            Top             =   300
            Width           =   420
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICMS CST"
            Height          =   195
            Left            =   2220
            TabIndex        =   101
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.Frame frmDados 
         Caption         =   "Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   120
         TabIndex        =   88
         Top             =   420
         Width           =   11535
         Begin VB.CheckBox chkMateriaPrima 
            Caption         =   "Matéria Prima"
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
            Left            =   7020
            TabIndex        =   198
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1515
         End
         Begin VB.CheckBox ckkImobilizado 
            Caption         =   "Imobilizado"
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
            Left            =   8580
            TabIndex        =   197
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1275
         End
         Begin VB.CheckBox chkCombustivel 
            Caption         =   "Combustível"
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
            Left            =   5580
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtRef 
            Height          =   315
            Left            =   5580
            TabIndex        =   11
            Top             =   1140
            Width           =   1515
         End
         Begin VB.CheckBox chkPedirPeso 
            Caption         =   "Pedir Peso"
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
            TabIndex        =   152
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1275
         End
         Begin VB.CheckBox chkUsoConsumo 
            Caption         =   "Uso/Consumo"
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
            Left            =   9900
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1515
         End
         Begin VB.ComboBox cboCategoria 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1140
            Width           =   2235
         End
         Begin VB.TextBox txtDescricao 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3180
            MaxLength       =   90
            TabIndex        =   4
            Top             =   480
            Width           =   5415
         End
         Begin VB.ComboBox cboUnidMedida 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   10740
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   675
         End
         Begin VB.ComboBox cboFabricante 
            Height          =   315
            Left            =   8640
            TabIndex        =   5
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtCodBarra 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            MaxLength       =   90
            TabIndex        =   1
            Top             =   480
            Width           =   1395
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
            Left            =   1440
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtObs 
            Height          =   315
            Left            =   7140
            MaxLength       =   90
            TabIndex        =   12
            Top             =   1140
            Width           =   4275
         End
         Begin VB.TextBox txtQuant 
            Height          =   315
            Left            =   3480
            TabIndex        =   9
            Top             =   1140
            Width           =   1035
         End
         Begin VB.TextBox txtQuantMin 
            Height          =   315
            Left            =   2400
            TabIndex        =   8
            Top             =   1140
            Width           =   1035
         End
         Begin VB.TextBox txtPrateleira 
            Height          =   315
            Left            =   4560
            MaxLength       =   4
            TabIndex        =   10
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox txtEAN 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1740
            MaxLength       =   90
            TabIndex        =   3
            Top             =   480
            Width           =   1395
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton1 
            Height          =   315
            Left            =   1530
            TabIndex        =   2
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
            MICON           =   "Produtos_Cadastro.frx":9CBF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblRef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            Height          =   195
            Left            =   5580
            TabIndex        =   155
            Top             =   900
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   99
            Top             =   900
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            Height          =   195
            Left            =   3180
            TabIndex        =   98
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unid."
            Height          =   195
            Left            =   10740
            TabIndex        =   97
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   8640
            TabIndex        =   96
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. Barra"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Observação 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações Adicionais"
            Height          =   195
            Left            =   7140
            TabIndex        =   94
            Top             =   900
            Width           =   1635
         End
         Begin VB.Label lblQuantAtual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Atual"
            Height          =   195
            Left            =   3480
            TabIndex        =   93
            Top             =   900
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EAN"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1740
            TabIndex        =   92
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Min."
            Height          =   195
            Left            =   2400
            TabIndex        =   91
            Top             =   900
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local"
            Height          =   195
            Left            =   4560
            TabIndex        =   90
            Top             =   900
            Width           =   390
         End
      End
      Begin VB.Frame frmPrecos 
         Caption         =   "Preços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   120
         TabIndex        =   72
         Top             =   2400
         Width           =   11535
         Begin VB.Frame Frame1 
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
            Left            =   1800
            TabIndex        =   84
            Top             =   180
            Width           =   2415
            Begin VB.TextBox txtMargemVV 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   14
               Top             =   480
               Width           =   895
            End
            Begin VB.TextBox txtValorVV 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               TabIndex        =   15
               Top             =   480
               Width           =   975
            End
            Begin ChamaleonBtn.chameleonButton cmdRepetir 
               Height          =   315
               Left            =   2100
               TabIndex        =   16
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
               MICON           =   "Produtos_Cadastro.frx":9CDB
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   180
               TabIndex        =   86
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1140
               TabIndex        =   85
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.Frame Frame2 
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
            Left            =   4260
            TabIndex        =   81
            Top             =   180
            Width           =   2175
            Begin VB.TextBox txtValorVP 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               TabIndex        =   18
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtMargemVP 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   17
               Top             =   480
               Width           =   895
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1140
               TabIndex        =   83
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   180
               TabIndex        =   82
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame3 
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
            Left            =   6480
            TabIndex        =   78
            Top             =   180
            Width           =   2175
            Begin VB.TextBox txtMargemAV 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   19
               Top             =   480
               Width           =   895
            End
            Begin VB.TextBox txtValorAV 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               TabIndex        =   20
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   180
               TabIndex        =   80
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1140
               TabIndex        =   79
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
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
            Left            =   8700
            TabIndex        =   75
            Top             =   180
            Width           =   2175
            Begin VB.TextBox txtValorAP 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               TabIndex        =   22
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtMargemAP 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   480
               Width           =   895
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1140
               TabIndex        =   77
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   180
               TabIndex        =   76
               Top             =   240
               Width           =   735
            End
         End
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
            Height          =   915
            Left            =   60
            TabIndex        =   73
            Top             =   180
            Width           =   1695
            Begin VB.TextBox txtCusto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   60
               TabIndex        =   13
               Top             =   480
               Width           =   1515
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Último Vlr Custo"
               Height          =   195
               Left            =   120
               TabIndex        =   74
               Top             =   240
               Width           =   1110
            End
         End
         Begin VB.Label lblAviso 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pressione [ F2 ]  para obter o lucro estimado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1920
            TabIndex        =   87
            Top             =   1080
            Visible         =   0   'False
            Width           =   3960
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   -74880
         TabIndex        =   67
         Top             =   7920
         Width           =   13875
         Begin VB.ComboBox cboAnoCons 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   1740
            TabIndex        =   68
            Top             =   240
            Width           =   1695
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirQuant 
            Height          =   315
            Left            =   4980
            TabIndex        =   70
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
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
            MICON           =   "Produtos_Cadastro.frx":9CF7
            PICN            =   "Produtos_Cadastro.frx":9D13
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblCONmes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&scolha o mês/ano:"
            Height          =   195
            Left            =   180
            TabIndex        =   71
            Top             =   300
            Width           =   1425
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   -74880
         TabIndex        =   62
         Top             =   7920
         Width           =   13875
         Begin VB.ComboBox cboAnoPreco 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   64
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cboMesPreco 
            Height          =   315
            Left            =   1680
            TabIndex        =   63
            Top             =   240
            Width           =   1755
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirPreco 
            Height          =   315
            Left            =   4980
            TabIndex        =   65
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
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
            MICON           =   "Produtos_Cadastro.frx":BAA5
            PICN            =   "Produtos_Cadastro.frx":BAC1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&scolha o mês/ano:"
            Height          =   195
            Left            =   180
            TabIndex        =   66
            Top             =   300
            Width           =   1425
         End
      End
      Begin VB.TextBox txtUltCompra 
         Enabled         =   0   'False
         Height          =   315
         Left            =   12420
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   4680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frmCriterios 
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
         Height          =   915
         Left            =   -74940
         TabIndex        =   58
         Top             =   7920
         Width           =   2055
         Begin VB.ComboBox cboCriterios 
            Height          =   315
            Left            =   120
            TabIndex        =   59
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label32 
            Caption         =   "Escolha:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   795
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frmOrdemComum 
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
         Height          =   915
         Left            =   -72840
         TabIndex        =   55
         Top             =   7920
         Width           =   3135
         Begin VB.ComboBox cboOrdem2 
            Height          =   315
            Left            =   1920
            TabIndex        =   150
            Top             =   480
            Width           =   1155
         End
         Begin VB.ComboBox cboOrdem 
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Width           =   1755
         End
         Begin VB.Label Label34 
            Caption         =   "Escolha:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Visible         =   0   'False
            Width           =   795
            WordWrap        =   -1  'True
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Quant 
         Height          =   6315
         Left            =   -74880
         TabIndex        =   124
         Top             =   840
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   11139
         _Version        =   393216
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   128
         Top             =   420
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   10610
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid GridPrecos 
         Height          =   6915
         Left            =   -74880
         TabIndex        =   129
         Top             =   840
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   12197
         _Version        =   393216
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   11820
         TabIndex        =   47
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
         MICON           =   "Produtos_Cadastro.frx":D853
         PICN            =   "Produtos_Cadastro.frx":D86F
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
         Left            =   11820
         TabIndex        =   48
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "Produtos_Cadastro.frx":F601
         PICN            =   "Produtos_Cadastro.frx":F61D
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
         Left            =   11820
         TabIndex        =   49
         Top             =   3180
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "Produtos_Cadastro.frx":113AF
         PICN            =   "Produtos_Cadastro.frx":113CB
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
         Left            =   11820
         TabIndex        =   46
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
         MICON           =   "Produtos_Cadastro.frx":1315D
         PICN            =   "Produtos_Cadastro.frx":13179
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
         Left            =   11820
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
         MICON           =   "Produtos_Cadastro.frx":14F0B
         PICN            =   "Produtos_Cadastro.frx":14F27
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
         Left            =   11820
         TabIndex        =   130
         Top             =   8220
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
         MICON           =   "Produtos_Cadastro.frx":16CB9
         PICN            =   "Produtos_Cadastro.frx":16CD5
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
         Height          =   855
         Left            =   -62460
         TabIndex        =   131
         Top             =   7980
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1508
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
         MICON           =   "Produtos_Cadastro.frx":18A67
         PICN            =   "Produtos_Cadastro.frx":18A83
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
         Height          =   855
         Left            =   -63900
         TabIndex        =   132
         Top             =   7980
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1508
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
         MICON           =   "Produtos_Cadastro.frx":1A815
         PICN            =   "Produtos_Cadastro.frx":1A831
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton ccmdDuplicar 
         Height          =   315
         Left            =   -70560
         TabIndex        =   205
         Top             =   6480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Duplicar"
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
         MICON           =   "Produtos_Cadastro.frx":1C5C3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdEditar 
         Height          =   315
         Left            =   -74880
         TabIndex        =   206
         Top             =   6480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Editar"
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
         MICON           =   "Produtos_Cadastro.frx":1C5DF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdApagar 
         Height          =   315
         Left            =   -73440
         TabIndex        =   207
         Top             =   6480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Excluir"
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
         MICON           =   "Produtos_Cadastro.frx":1C5FB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdDesativar 
         Height          =   315
         Left            =   -72000
         TabIndex        =   208
         Top             =   6480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Desativar"
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
         MICON           =   "Produtos_Cadastro.frx":1C617
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "ESTOQUE DESSE PRODUTO HOJE:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   169
         Top             =   7200
         Width           =   2730
      End
      Begin VB.Label lblEstoqueHoje 
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
         Left            =   -72060
         TabIndex        =   168
         Top             =   7200
         Width           =   225
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tam."
         Height          =   195
         Left            =   12420
         TabIndex        =   154
         Top             =   4080
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Dê um duplo-clique para ver mais informações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -67320
         TabIndex        =   151
         Top             =   60
         Width           =   4035
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         X1              =   11700
         X2              =   11700
         Y1              =   420
         Y2              =   8580
      End
      Begin VB.Label lblNomeProduto1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   -74880
         TabIndex        =   138
         Top             =   420
         Width           =   13875
      End
      Begin VB.Label lblNomeProduto2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   -74880
         TabIndex        =   137
         Top             =   420
         Width           =   13875
      End
      Begin VB.Label lblQuantAdicao 
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
         Left            =   -61260
         TabIndex        =   136
         Top             =   7200
         Width           =   225
      End
      Begin VB.Label lblQuantRemocao 
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
         Left            =   -61260
         TabIndex        =   135
         Top             =   7440
         Width           =   225
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "ADIÇÃO:"
         Height          =   195
         Left            =   -62280
         TabIndex        =   134
         Top             =   7200
         Width           =   645
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "REMOÇÃO:"
         Height          =   195
         Left            =   -62460
         TabIndex        =   133
         Top             =   7440
         Width           =   855
      End
   End
End
Attribute VB_Name = "Produtos_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private moCombo As cComboHelper
Private printSQL As String
Dim var_cod_Preco As Long
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Dim VarIncluirPreco As Integer
Dim vTipoOS As String
Dim vMultiplasRef As String
Dim vUltimaReferencia As Integer
Dim sSQL As String
Dim r As ADODB.Recordset
'Dim vModoEdicao As Boolean
Public var_AliqInterna As Double    'buscar aliquota do estado para calcular icms

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Private Sub LimparObjeto_Gas()
txtcProdANP.Text = ""
txtdescricaoANP.Text = ""
txtpGLP.Text = ""
txtpGNn.Text = ""
txtpGNi.Text = ""
txtpMixGN.Text = ""
'txtValorPartida.Text = ""
txtCODIF.Text = ""
If txtpGLP.Text = "" Or txtpGLP.Text = "0" Then txtpGLP.Text = FormatNumber(0, 2) & "%"
If txtpGNn.Text = "" Or txtpGNn.Text = "0" Then txtpGNn.Text = FormatNumber(0, 2) & "%"
If txtpGNi.Text = "" Or txtpGNi.Text = "0" Then txtpGNi.Text = FormatNumber(0, 2) & "%"
If txtpMixGN.Text = "" Or txtpMixGN.Text = "0" Then txtpMixGN.Text = FormatNumber(0, 2) & "%"
txtValorPartida.Text = Format(0, ocMONEY)
frmGas.Visible = False
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
Public Sub CriarNovoProduto()
HabilitarFrames
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
vTipoEdicao = "Novo" 'desativei para teste
cmdExcluir.Enabled = False

LimparObjetos_Produtos

If frmComp.Visible = True Then LimparGrid_Comp

AutoNumeracao

cboUnidMedida.Text = "UN"
txtQuant.Text = "0"
'txtCodBarra.SetFocus
End Sub

Private Sub FormatarGrid_Historico(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   Dim x As Integer
   
   With Grid_Quant
      .Clear
      .Cols = 7
      .rows = 2
      
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
      For x = 0 To .Cols - 1
         .Col = x
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
            
            .TextMatrix(.rows - 1, 1) = rTabela("var_codigo")
            .TextMatrix(.rows - 1, 2) = Format$(rTabela("data_entrada"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = rTabela("notafiscal")
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("razao"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("quant"))
'            .TextMatrix(.Rows - 1, 6) = Format$(rTabela("custo"), ocMONEY)
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .Redraw = True
      .rows = .rows - 1
   End With
End Sub
Private Sub CalcularPrecos()
Dim varValorCusto As Currency
If txtCusto.Text = "" Then Exit Sub
varValorCusto = txtCusto.Text

'CALCULAR PREÇO - VAREJO A VISTA
Dim varMargemVV As Currency
Dim varValorVV As Currency

If txtMargemVV.Text = "" Then Exit Sub
If txtMargemVV.Text = "0" Then Exit Sub

varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)

varValorVV = (varValorCusto * varMargemVV) / 100
varValorVV = varValorCusto + varValorVV
txtValorVV.Text = Format(varValorVV, ocMONEY)

'CALCULAR PREÇO - VAREJO A PRAZO
Dim varMargemVP As Currency
Dim varValorVP As Currency

If txtMargemVP.Text = "" Then Exit Sub
If txtMargemVP.Text = "0" Then Exit Sub

varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)

varValorVP = (varValorCusto * varMargemVP) / 100
varValorVP = varValorCusto + varValorVP
txtValorVP.Text = Format(varValorVP, ocMONEY)

'CALCULAR PREÇO - ATACADO A VISTA
Dim varMargemAV As Currency
Dim varValorAV As Currency

If txtMargemAV.Text = "" Then Exit Sub
If txtMargemAV.Text = "0" Then Exit Sub

varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)

varValorAV = (varValorCusto * varMargemAV) / 100
varValorAV = varValorCusto + varValorAV
txtValorAV.Text = Format(varValorAV, ocMONEY)

'CALCULAR PREÇO - ATACADO A PRAZO
Dim varMargemAP As Currency
Dim varValorAP As Currency

If txtMargemAP.Text = "" Then Exit Sub
If txtMargemAP.Text = "0" Then Exit Sub

varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)

varValorAP = (varValorCusto * varMargemAP) / 100
varValorAP = varValorCusto + varValorAP
txtValorAP.Text = Format(varValorAP, ocMONEY)
End Sub

Private Sub DesabilitarBotoes()
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
vTipoEdicao = "Novo" 'desativei para teste
cmdExcluir.Enabled = False
End Sub

Private Sub DesabilitarFrames()
frmDados.Enabled = False
frmPrecos.Enabled = False
frmFiscal.Enabled = False
frmComp.Enabled = False
End Sub

Private Sub MostrarObjetosPrecos()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

sSQL = "SELECT TOP 1 * FROM Produtos_Precos WHERE (COD_PRODUTO = " & txtCodigo.Text & ") order by CODIGO desc;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    txtCusto.Text = Format$(r("CUSTO"), ocMONEY)
    txtValorVV.Text = Format$(r("VALOR_VV"), ocMONEY)
    txtValorVP.Text = Format$(r("VALOR_VP"), ocMONEY)
    txtValorAV.Text = Format$(r("VALOR_AV"), ocMONEY)
    txtValorAP.Text = Format$(r("VALOR_AP"), ocMONEY)
    txtMargemVV.Text = FormatNumber(r("MARGEM_VV"), 3) & "%"
    txtMargemVP.Text = FormatNumber(r("MARGEM_VP"), 3) & "%"
    txtMargemAV.Text = FormatNumber(r("MARGEM_AV"), 3) & "%"
    txtMargemAP.Text = FormatNumber(r("MARGEM_AP"), 3) & "%"
End If
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Quant_Entrada()
Dim sSQL As String
Dim r As ADODB.Recordset

'ENTRADA DO PRODUTO
If cmdSalvar.Enabled = True Then
   Dim AutoNumeracao As Long
   
   'AUTONUMERAÇÃO
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_quant;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing

   
   sSQL = "INSERT INTO produtos_quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO) VALUES (" & _
      AutoNumeracao & ", " & txtCodigo.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), 0, 'CADASTRO', " & Replace(CDbl(txtQuant.Text), ",", ".") & ", 'ADIÇÃO');"
   dbData.Execute sSQL
End If
End Sub

Private Sub Preco_Entrada()
Dim sSQL As String
Dim r As ADODB.Recordset

'ENTRADA DO PRODUTO
If cmdSalvar.Enabled = True Then
   Dim AutoNumeracao As Long
   
   'AUTONUMERAÇÃO
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_precos;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing

    Dim varMargemVV As Double
    Dim varMargemVP As Double
    Dim varMargemAV As Double
    Dim varMargemAP As Double
    
    varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)
    varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)
    varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)
    varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)
   
   sSQL = "INSERT INTO produtos_precos (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP, CUSTO) VALUES (" & _
      AutoNumeracao & ", " & txtCodigo.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & txtCodigo.Text & ", 'CADASTRO', " & Replace(CDbl(varMargemVV), ",", ".") & ", " & Replace(CCur(txtValorVV.Text), ",", ".") & ", " & Replace(CDbl(varMargemVP), ",", ".") & ", " & Replace(CCur(txtValorVP.Text), ",", ".") & ", " & Replace(CDbl(varMargemAV), ",", ".") & ", " & Replace(CCur(txtValorAV.Text), ",", ".") & ", " & Replace(CDbl(varMargemAP), ",", ".") & ", " & Replace(CCur(txtValorAP.Text), ",", ".") & ", " & Replace(CCur(txtCusto.Text), ",", ".") & "  );"
   dbData.Execute sSQL
End If
End Sub
Private Sub FormatarGrid_Comp(rTabela As ADODB.Recordset)
Dim i As Integer, j As Integer
Dim x As Integer

With Grid_Comp
   .Clear
   .Cols = 5
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 0
   .ColWidth(3) = 2300
   .ColWidth(4) = 900
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "COD_PRODUTO"
   .TextMatrix(0, 3) = "FAB/MODELO"
   .TextMatrix(0, 4) = "ANO"
   
   'colocar os cabeçalho em negrito
   For x = 0 To .Cols - 1
      .Col = x
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
         
         .TextMatrix(.rows - 1, 1) = rTabela("CODIGO")
         .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("COD_PRODUTO"))
         .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("MODELO"))
         .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("ANO"))
        
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .Redraw = True
   .rows = .rows - 1
End With
End Sub
Private Sub FormatarGrid_Quant(rTabela As ADODB.Recordset)
Dim x As Integer

With Grid_Quant
   .Clear
   .Cols = 11
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 1000
   .ColWidth(3) = 1000
   .ColWidth(4) = 0
   .ColWidth(5) = 1500
   .ColWidth(6) = 1800
   .ColWidth(7) = 1800
   .ColWidth(8) = 1000
   .ColWidth(9) = 1000
   .ColWidth(10) = 2000

   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "DATA"
   .TextMatrix(0, 3) = "HORA"
   .TextMatrix(0, 4) = "COD_PRODUTO"
   .TextMatrix(0, 5) = "NOTA FISCAL"
   .TextMatrix(0, 6) = "TIPO"
   .TextMatrix(0, 7) = "FORMA"
   .TextMatrix(0, 8) = "QUANT"
   .TextMatrix(0, 9) = "USUÁRIO"
   .TextMatrix(0, 10) = "ESTOQUE NO DIA"
   
   'colocar os cabeçalho em negrito
   For x = 0 To .Cols - 1
      .Col = x
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
         .TextMatrix(.rows - 1, 1) = ValidateNull(rTabela("Codigo"))
         .TextMatrix(.rows - 1, 2) = Format$(rTabela("Data"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = Format$(rTabela("HORA"), ocHORA)
         .TextMatrix(.rows - 1, 4) = rTabela("COD_PRODUTO")
         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("notafiscal"))
         .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("TIPO"))
         .TextMatrix(.rows - 1, 7) = rTabela("FORMA")
         .TextMatrix(.rows - 1, 8) = rTabela("QUANT")
         .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("COD_USUARIO"))
         .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("ESTOQUE"))

         
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   
    'Deixar negrito quando vencido
    For i = 1 To .rows - 1
       For j = 0 To .Cols - 1
          .Col = j
          .Row = i
          
          If .TextMatrix(i, 5) = "REMOÇÃO" Then
             .CellForeColor = vbRed
             .CellFontBold = True
          End If
       Next
    Next

      
    'MUDAR COR DE FONTE DA COLUNA
     For i = 1 To .rows - 1
        .Row = i
        .Col = 7
        .CellBackColor = &HC0FFFF
        .CellFontBold = True
     Next
   
   .Redraw = True
   .rows = .rows - 1
End With
End Sub

Private Sub FormatarGrid_Precos(rTabela As ADODB.Recordset)
Dim i As Integer, j As Integer
Dim x As Integer

With GridPrecos
   .Clear
   .Cols = 13
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 600
   .ColWidth(2) = 1000
   .ColWidth(3) = 1200
   .ColWidth(4) = 1000
   .ColWidth(5) = 900
   .ColWidth(6) = 1150
   .ColWidth(7) = 900
   .ColWidth(8) = 1150
   .ColWidth(9) = 900
   .ColWidth(10) = 1150
   .ColWidth(11) = 900
   .ColWidth(12) = 1150

   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "DATA"
   .TextMatrix(0, 3) = "FORMA"
   .TextMatrix(0, 4) = "CUSTO"
   .TextMatrix(0, 5) = "% VV"
   .TextMatrix(0, 6) = "VALOR VV"
   .TextMatrix(0, 7) = "% VP "
   .TextMatrix(0, 8) = "VALOR VP"
   .TextMatrix(0, 9) = "% AV"
   .TextMatrix(0, 10) = "VALOR AV"
   .TextMatrix(0, 11) = "% AP"
   .TextMatrix(0, 12) = "VALOR AP"
   
   'colocar os cabeçalho em negrito
   For x = 0 To .Cols - 1
      .Col = x
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
         .TextMatrix(.rows - 1, 1) = rTabela("Codigo")
         .TextMatrix(.rows - 1, 2) = Format$(rTabela("Data"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = rTabela("FORMA")
         .TextMatrix(.rows - 1, 4) = Format$(rTabela("custo"), ocMONEY)
         .TextMatrix(.rows - 1, 5) = FormatNumber(rTabela("MARGEM_VV"), 2) & "%"
         .TextMatrix(.rows - 1, 6) = Format$(rTabela("VALOR_VV"), ocMONEY)
         .TextMatrix(.rows - 1, 7) = FormatNumber(rTabela("MARGEM_VP"), 2) & "%"
         .TextMatrix(.rows - 1, 8) = Format$(rTabela("VALOR_VP"), ocMONEY)
         .TextMatrix(.rows - 1, 9) = FormatNumber(rTabela("MARGEM_AV"), 2) & "%"
         .TextMatrix(.rows - 1, 10) = Format$(rTabela("VALOR_AV"), ocMONEY)
         .TextMatrix(.rows - 1, 11) = FormatNumber(rTabela("MARGEM_AP"), 2) & "%"
         .TextMatrix(.rows - 1, 12) = Format$(rTabela("VALOR_AP"), ocMONEY)
         
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
        'MUDAR COR DE FONTE DA COLUNA
         For i = 1 To .rows - 1
            .Row = i
            .Col = 6
            .CellBackColor = &HC0FFFF
            .CellFontBold = True
         Next
         
        'COLUNA EM NEGRITO
         For i = 1 To .rows - 1
            .Row = i
            .Col = 5
            .CellFontBold = True
         Next
         
        'COLUNA EM NEGRITO
         For i = 1 To .rows - 1
            .Row = i
            .Col = 7
            .CellFontBold = True
         Next
         
        'COLUNA EM NEGRITO
         For i = 1 To .rows - 1
            .Row = i
            .Col = 9
            .CellFontBold = True
         Next
         
        'COLUNA EM NEGRITO
         For i = 1 To .rows - 1
            .Row = i
            .Col = 11
            .CellFontBold = True
         Next
   
   .Redraw = True
   .rows = .rows - 1
End With
End Sub



Private Sub FormatarGrid_Produtos(rTabela As ADODB.Recordset)
'Dim cCfg As ConfigItem
'Dim tipoEmpresa As Integer
            
'Set cCfg = sysConfig("TIPO_EMPRESA")
'tipoEmpresa = cCfg.Value
'Set cCfg = Nothing

Dim i As Integer, j As Integer
Dim x As Integer

 Dim VarTotalGrid As Currency
 VarTotalGrid = 0
 Dim VarTotalCustoGrid As Currency
 VarTotalCustoGrid = 0
   
'If tipoEmpresa = 4 Then
'   With Grid
'      .Clear
'      .Cols = 11
'      .Rows = 2
      
'      .ColWidth(0) = 0
'      .ColWidth(1) = 0
'      .ColWidth(2) = 1600 '1600
'      .ColWidth(3) = 4645 '4445
'      .ColWidth(4) = 850
'      .ColWidth(5) = 850
'      .ColWidth(6) = 1500
'      .ColWidth(7) = 850
'      .ColWidth(8) = 850
'      .ColWidth(9) = 850
'      .ColWidth(10) = 850
      
'      .TextMatrix(0, 1) = "COD"
'      .TextMatrix(0, 2) = "CÓD. BARRA"
'      .TextMatrix(0, 3) = "PRODUTO"
'      .TextMatrix(0, 4) = "TAM."
'      .TextMatrix(0, 5) = "REF."
'      .TextMatrix(0, 6) = "FABRICANTE"
'      .TextMatrix(0, 7) = "QUANT"
'      .TextMatrix(0, 8) = "MED."
'      .TextMatrix(0, 9) = "VENDA"
'      .TextMatrix(0, 10) = "TOTAL"
      
'      'colocar os cabeçalho em negrito
'      For x = 0 To .Cols - 1
'         .Col = x
'         .Row = 0
'         .CellFontBold = True
'      Next
      
'      'centralizar o titulo
'      For j = 0 To .Cols - 1
'         .Row = 0
'         .Col = j
'         .CellAlignment = flexAlignCenterCenter
'      Next
      
 '     .Redraw = False
      
'      If Not rTabela Is Nothing Then
'         Do While Not rTabela.EOF
'            'mudar a cor da coluna
'            'For i = 1 To .Rows - 1
'            '   .Row = i
'            '   .Col = 6
'            '   .CellBackColor = &HC0FFFF
'            'Next
            
'            'ALINHAMENTO
'            '.ColAlignment(2) = 1
'            VarTotalGrid = 0
'            .TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
'            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("var_codbarra"))
'            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("var_desc"))
'            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("var_TAM"))
'            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("var_REF"))
'            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("var_fab"))
'            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_quant"))
'            .TextMatrix(.Rows - 1, 8) = ValidateNull(rTabela("var_med"))
'            .TextMatrix(.Rows - 1, 9) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
'            VarTotalGrid = .TextMatrix(.Rows - 1, 9) * .TextMatrix(.Rows - 1, 7)
'            .TextMatrix(.Rows - 1, 10) = Format(VarTotalGrid, ocMONEY)
            
 '           rTabela.MoveNext
 '           .Rows = .Rows + 1
 '        Loop
  '    End If
  '
  '    .Redraw = True
  '    .Rows = .Rows - 1
  ' End With
   
  ' 'calcular quantidade
  ' Dim Col As Integer
  ' Dim Valor As Currency
   'Col = 7
  ' Valor = 0
   
 '  For i = 0 To Grid.Rows - 1
  '    If IsNumeric(Grid.TextMatrix(i, Col)) And Grid.TextMatrix(i, Col) <> "0" Then
 '        Valor = Valor + CCur(Grid.TextMatrix(i, Col))
 '     End If
 '  Next
   
'   lblProdutos.Caption = Format(Valor, ocMONEY)
   
'   lblValorTotal.Caption = Format(SomaGrid(Grid, 10), ocMONEY)
 '  lblProdutos.Caption = SomaGrid(Grid, 7)
'   lblTipos.Caption = Grid.Rows - 1  'contar o numeros de linhas no grid

'Else
   With Grid
      .Clear
      .Cols = 13
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1300 '1600
      .ColWidth(3) = 4745 '4445
      .ColWidth(4) = 1350
      .ColWidth(5) = 450 'MED
      .ColWidth(6) = 550 'LOC
      .ColWidth(7) = 800
      .ColWidth(8) = 850
      .ColWidth(9) = 1000
      .ColWidth(10) = 850
      .ColWidth(11) = 1000
      .ColWidth(12) = 1000
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "CÓD. BARRA"
      .TextMatrix(0, 3) = "PRODUTO / REF"
      .TextMatrix(0, 4) = "FABRICANTE"
      .TextMatrix(0, 5) = "UN"
      .TextMatrix(0, 6) = "LOC"
      .TextMatrix(0, 7) = "QTDE"
      .TextMatrix(0, 8) = "CUSTO"
      .TextMatrix(0, 9) = "TOTAL"
      .TextMatrix(0, 10) = "VENDA"
      .TextMatrix(0, 11) = "TOTAL"
      .TextMatrix(0, 12) = "SITUAÇÃO"

      
      'colocar os cabeçalho em negrito
      For x = 0 To .Cols - 1
         .Col = x
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
            '   .Col = 1
            '   .CellBackColor = &HC0FFFF
            'Next
            
            
'sSQL = "SELECT TOP 1 Produtos_Precos.VALOR_VV as varVenda, produtos.codigo as varCodProd, produtos.ref AS var_Ref, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.*, Produtos_Precos.* " & _
'    "produtos.fabricante AS var_fab, produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant  " & _

            'ALINHAMENTO
            '.ColAlignment(2) = 1
             VarTotalGrid = 0
            .TextMatrix(.rows - 1, 1) = Format$(ValidateNull(rTabela("varCodProd")), "000000")
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("var_desc")) & " / " & ValidateNull(rTabela("var_Ref"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_med"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_LOCAL"))
            

            .TextMatrix(.rows - 1, 7) = Format$(ValidateNull(rTabela("var_quant")), ocPESO)

            .TextMatrix(.rows - 1, 8) = Format$(ValidateNull(rTabela("CUSTO")), ocMONEY)
            VarTotalCustoGrid = .TextMatrix(.rows - 1, 8) * .TextMatrix(.rows - 1, 7)
            .TextMatrix(.rows - 1, 9) = Format(VarTotalCustoGrid, ocMONEY)
            
            .TextMatrix(.rows - 1, 10) = Format$(ValidateNull(rTabela("Venda")), ocMONEY)
            VarTotalGrid = .TextMatrix(.rows - 1, 10) * .TextMatrix(.rows - 1, 7)
            .TextMatrix(.rows - 1, 11) = Format(VarTotalGrid, ocMONEY)
            .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("vAtivo"))
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .Redraw = True
      .rows = .rows - 1
   End With
   
   lblValorTotalCusto.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
   lblValorTotal.Caption = Format(SomaGrid(Grid, 11), ocMONEY)
   lblProdutos.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
   lblTipos.Caption = Grid.rows - 1  'contar o numeros de linhas no grid
'End If
End Sub

Private Sub HabilitarFrames()
frmDados.Enabled = True
frmPrecos.Enabled = True
frmFiscal.Enabled = True
frmComp.Enabled = True
End Sub
Private Sub LimparGrid_Ref()
sSQL = "Select * FROM Produtos_Referencias WHERE 0 = 1;"

Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Referencia r
FormatarGrid_Referencia_Desc r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub LimparGrid_Comp()
sSQL = "Select * FROM PRODUTOS_COMP WHERE 0 = 1;"

Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Comp r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub LimparGrid_Produtos()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT TOP 1 Produtos_Precos.VALOR_VV as varVenda, produtos.*, Produtos_Precos.* " & _
  "FROM Produtos_Precos INNER JOIN produtos ON Produtos_Precos.cod_produto = produtos.codigo " & _
  "WHERE 0 = 1;"

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

Private Sub MostrarDados_Produto()
Dim vrCusto As Currency
Dim vrVenda As Currency

sSQL = "SELECT * FROM produtos WHERE (codigo = " & txtCodigo.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If r("FRACIONADO") = True Then
    frmFracionado.Visible = True
Else
    frmFracionado.Visible = False
End If

txtCodBarra.Text = ValidateNull(r("cod_barra"))
txtEAN.Text = ValidateNull(r("EAN"))
txtDescricao.Text = ValidateNull(r("descricao"))
chkFracionado.Value = Abs(CBool(r("FRACIONADO")))
txtQuantFracionado.Text = ValidateNull(r("QUANT_FRACAO"))
cboFabricante.Text = ValidateNull(r("fabricante"))
cboUnidMedida.Text = ValidateNull(r("unid_medida"))
cboCategoria.Text = ValidateNull(r("categoria"))
txtPrateleira.Text = ValidateNull(r("PRATELEIRA"))

If vTipoEdicao = "Duplicar" Then
    txtQuant.Text = 0
    lblEstoqueHoje.Caption = 0
Else
    txtQuant.Text = ValidateNull(r("quant_estoque"))
    lblEstoqueHoje.Caption = ValidateNull(r("quant_estoque"))
End If

txtCusto.Text = Format(vrCusto, ocMONEY)
txtQuantMin.Text = ValidateNull(r("quant_min"))
txtUltCompra.Text = Format$(r("ult_compra"), "dd/mm/yy")
txtObs.Text = ValidateNull(r("INF_ADICIONA"))
txtRef.Text = ValidateNull(r("ref"))
txtTam.Text = ValidateNull(r("tamanho"))
txtICMSAliquota.Text = Format(ValidateNull(r("ICMSAliq")), ocMONEY)
txtPisAliquota.Text = Format(ValidateNull(r("PISAliq")), ocMONEY)
txtRedBCAliquota.Text = Format(ValidateNull(r("pRedBc")), ocMONEY)
txtCofinsAliquota.Text = Format(ValidateNull(r("COFINSAliq")), ocMONEY)
txtIPIAliquota.Text = Format(ValidateNull(r("IPIAliq")), ocMONEY)
txtPISCST.Text = ValidateNull(r("PISCST"))
txtCOFINSCST.Text = ValidateNull(r("COFINSCST"))
txtIPICST.Text = ValidateNull(r("IPICST"))
txtNCM.Text = ValidateNull(r("NCM"))
cboCST.Text = ValidateNull(r("icmscst"))
cboCFOP.Text = ValidateNull(r("cfop"))
txtCEST.Text = ValidateNull(r("CEST"))
chkDestaque.Value = Abs(CBool(ValidateNull(r("destaque"))))
chkPedirPeso.Value = Abs(CBool(ValidateNull(r("PEDIRPESO"))))
chkUsoConsumo.Value = Abs(CBool(ValidateNull(r("USOCONSUMO"))))
chkCombustivel.Value = Abs(CBool(ValidateNull(r("COMBUSTIVEL"))))
chkMateriaPrima.Value = Abs(CBool(ValidateNull(r("MATERIAPRIMA"))))
ckkImobilizado.Value = Abs(CBool(ValidateNull(r("IMOBILIZADO"))))
txtCodProdFracionado.Text = ValidateNull(r("CODPROD_FRACAO"))
End Sub

Private Sub AutoNumeracao()
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_produto FROM produtos;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then txtCodigo.Text = r("cod_produto") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub LimparObjetos_Produtos()
If vTipoEdicao = "Novo" Or vTipoEdicao = "Cancela" Or vTipoEdicao = "Duplicar" Then txtCodigo.Text = ""
txtCodBarra.Text = ""
txtEAN.Text = ""
txtDescricao.Text = ""
cboFabricante.Text = ""
cboCategoria.Text = ""
'cboUnidMedida.Text = ""
txtPrateleira.Text = ""
txtQuant.Text = ""
txtQuantMin.Text = ""
txtUltCompra.Text = ""
txtObs.Text = ""
txtRef.Text = ""
txtTam.Text = ""
chkDestaque.Value = Unchecked
chkUsoConsumo.Value = Unchecked
chkCombustivel.Value = Unchecked
chkMateriaPrima.Value = Unchecked
ckkImobilizado.Value = Unchecked
txtICMSAliquota.Text = FormatNumber(0, 2)
txtPisAliquota.Text = FormatNumber(0, 2)
txtCofinsAliquota.Text = FormatNumber(0, 2)
txtIPIAliquota.Text = FormatNumber(0, 2)
txtRedBCAliquota.Text = FormatNumber(0, 2)
txtPISCST.Text = "04"
txtCOFINSCST.Text = "04"
txtIPICST.Text = "99"
txtNCM.Text = ""
'cboCFOP.Text = ""
'cboCST.Text = ""
txtCEST.Text = "0"
txtMargemVV.Text = Format(0, ocMONEY)
txtMargemVP.Text = Format(0, ocMONEY)
txtMargemAV.Text = Format(0, ocMONEY)
txtMargemAP.Text = Format(0, ocMONEY)
txtValorVV.Text = Format(0, ocMONEY)
txtValorVP.Text = Format(0, ocMONEY)
txtValorAV.Text = Format(0, ocMONEY)
txtValorAP.Text = Format(0, ocMONEY)
txtCusto.Text = Format(0, ocMONEY)
chkFracionado.Value = Unchecked
txtCodProdFracionado.Text = ""
cboProdutoFracionado.Text = ""
txtQuantFracionado.Text = ""
frmFracionado.Visible = False
If frmComp.Visible = True Then LimparGrid_Comp
End Sub

Private Sub cboAnoCons_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAnoCons.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = LastYear To FirstYear Step -1
   cboAnoCons.AddItem i
Next
End Sub


Private Sub cboAnoPreco_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAnoPreco.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = LastYear To FirstYear Step -1
   cboAnoPreco.AddItem i
Next
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

Private Sub cboCategoria_LostFocus()
cboCategoria.Text = TirarEspaco(cboCategoria.Text)
End Sub



Private Sub cboCFOP_Click()
    Dim sCFOP As String
    sCFOP = cboCFOP.Text

    ' 1. Limpa o combo para remover opções do CFOP anterior
    cboCST.Clear

    If var_RegimeEmpresa = 1 Then
        ' --- LÓGICA SIMPLES NACIONAL (CSOSN) ---
        Select Case sCFOP
            Case "5102", "5202"
                cboCST.AddItem "102": cboCST.AddItem "101"
                cboCST.ListIndex = 0
            Case "5405"
                cboCST.AddItem "500"
                cboCST.ListIndex = 0
            Case Else
                cboCST.AddItem "400"
        End Select

ElseIf var_RegimeEmpresa = 3 Then
    ' --- LÓGICA LUCRO PRESUMIDO (FILTRADO POR CFOP) ---
    Select Case sCFOP
        Case "5102" ' REVENDA DE MERCADORIA (O mais comum)
            cboCST.Clear
            cboCST.AddItem "000" ' Tributada integralmente (Padrão)
            cboCST.AddItem "020" ' Com redução de BC
            cboCST.AddItem "040" ' Isenta
            cboCST.AddItem "041" ' Não tributada
            cboCST.AddItem "070" ' Com redução de BC e cobrança de ST
            cboCST.AddItem "090" ' Outras
            cboCST.ListIndex = 0 ' Já sugere o 000

        Case "5101" ' FABRICAÇÃO PRÓPRIA (Indústria)
            cboCST.Clear
            cboCST.AddItem "000"
            cboCST.AddItem "010" ' Tributada com cobrança de ST
            cboCST.AddItem "051" ' Diferimento
            cboCST.AddItem "090"
            cboCST.ListIndex = 0

        Case "5401", "5403" ' VENDAS COM ST (Indústria/Importação)
            cboCST.Clear
            cboCST.AddItem "010" ' Tributada e com cobrança de ST
            cboCST.AddItem "030" ' Isenta/Não trib. com cobrança de ST
            cboCST.AddItem "070" ' Com redução de BC e cobrança de ST
            cboCST.AddItem "090"
            cboCST.ListIndex = 0

        Case "5405" ' REVENDA COM ST (Substituído)
            cboCST.Clear
            cboCST.AddItem "060" ' Padrão absoluto para 5405
            cboCST.AddItem "090"
            cboCST.ListIndex = 0

        Case "5949" ' OUTRAS SAÍDAS
            cboCST.Clear
            cboCST.AddItem "041": cboCST.AddItem "050": cboCST.AddItem "090"
            cboCST.ListIndex = 2 ' Sugere o 090
    End Select
End If
    ' 2. Seleciona automaticamente o primeiro item da lista filtrada
    'If cboCST.ListCount > 0 Then cboCST.ListIndex = 0
End Sub


Private Sub cboConsProduto_Change()
If cboCriterios.Text = "CÓD. BARRA" And Len(cboConsProduto) = 13 Then
  cmdExibir_Click
End If
End Sub

Private Sub cboConsProduto_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset

   If cboCriterios.Text = "CÓD. BARRA" Then
      cboConsProduto.Clear
   
   ElseIf cboCriterios.Text = "DESCRIÇÃO" Then
      cboConsProduto.Clear
      
      If optCompleto.Value = True Then
        sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsProduto.AddItem ValidateNull(r("descricao"))
           r.MoveNext
        Loop
      End If
      
   ElseIf cboCriterios.Text = "CATEGORIA" Then
      cboConsProduto.Clear
      
      sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"
      Set r = dbData.OpenRecordset(sSQL)
      
      Do While Not r.EOF
         cboConsProduto.AddItem ValidateNull(r("categoria"))
         r.MoveNext
      Loop
      
   ElseIf cboCriterios.Text = "REFERÊNCIA" Then
      cboConsProduto.Clear
      
      If vMultiplasRef = "SIM" Then
        sSQL = "SELECT DISTINCT referencia FROM produtos_referencias ORDER BY referencia;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsProduto.AddItem ValidateNull(r("referencia"))
           r.MoveNext
        Loop
      Else
        sSQL = "SELECT DISTINCT ref FROM produtos ORDER BY ref;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsProduto.AddItem ValidateNull(r("ref"))
           r.MoveNext
        Loop
      End If
      
   ElseIf cboCriterios.Text = "FABRICANTE" Then
      cboConsProduto.Clear
      
      sSQL = "SELECT DISTINCT fabricante FROM produtos ORDER BY fabricante;"
      Set r = dbData.OpenRecordset(sSQL)
      
      Do While Not r.EOF
         cboConsProduto.AddItem ValidateNull(r("fabricante"))
         r.MoveNext
      Loop
   End If
   
   
    SelectControl cboConsProduto
    moCombo.AttachTo cboConsProduto
End Sub

Private Sub cboConsProduto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboConsProduto_LostFocus()
If cboCriterios.Text = "CÓD. BARRA" Then
    If Len(cboConsProduto) < 13 And cboConsProduto.Text <> "" Then
        If Len(cboConsProduto) < 6 Then
            cboConsProduto.Text = Format(cboConsProduto.Text, "00000")
        Else
            cboConsProduto.Text = cboConsProduto.Text
        End If
    End If
End If
End Sub

Private Sub cboCriterios_Click()
cboCriterios_LostFocus
End Sub

Private Sub cboCriterios_GotFocus()
Dim var_Texto As String
var_Texto = cboCriterios.Text

cboCriterios.Clear
cboCriterios.AddItem "TODOS"
cboCriterios.AddItem "DESCRIÇÃO"
cboCriterios.AddItem "CÓD. BARRA"
cboCriterios.AddItem "CATEGORIA"
cboCriterios.AddItem "REFERÊNCIA"
cboCriterios.AddItem "FABRICANTE"
cboCriterios.AddItem "NCM"
moCombo.AttachTo cboCriterios
   
cboCriterios.Text = var_Texto
cboCriterios.SelStart = 0
cboCriterios.SelLength = Len(cboCriterios)
End Sub


Private Sub cboCriterios_LostFocus()
If cboCriterios.Text = "CÓD. BARRA" Then
   cboConsProduto.Visible = True
   lblNomeCombo.Visible = True
   lblNomeCombo.Caption = "Cód. de Barra"
   optTodosPreco.Value = True
   optMostrarTodos.Value = True
    optCompleto.Visible = False
    optPorIniciais.Visible = False
    optPorPalavra.Visible = False
    optPalavrasDuplas.Visible = False
   cboConsProduto.SetFocus
ElseIf cboCriterios.Text = "TODOS" Then
    cboConsProduto.Visible = False
    cboConsProduto.Visible = False
    lblNomeCombo.Visible = False
    optMostrarQuant.Value = True
    optCompleto.Visible = False
    optPorIniciais.Visible = False
    optPorPalavra.Visible = False
    optPalavrasDuplas.Visible = False
ElseIf cboCriterios.Text = "CATEGORIA" Then
   cboConsProduto.Visible = True
   lblNomeCombo.Visible = True
   lblNomeCombo.Caption = "Categoria"
    optTodosPreco.Value = True
    optMostrarTodos.Value = True
    optCompleto.Visible = False
    optPorIniciais.Visible = False
    optPorPalavra.Visible = False
    optPalavrasDuplas.Visible = False
   cboConsProduto.SetFocus
ElseIf cboCriterios.Text = "REFERÊNCIA" Then
    cboConsProduto.Visible = True
    lblNomeCombo.Visible = True
    lblNomeCombo.Caption = "Referência"
    optTodosPreco.Value = True
    optMostrarTodos.Value = True
    optCompleto.Visible = False
    optPorIniciais.Visible = False
    optPorPalavra.Visible = False
    optPalavrasDuplas.Visible = False
    cboConsProduto.SetFocus
ElseIf cboCriterios.Text = "DESCRIÇÃO" Then
   cboConsProduto.Visible = True
   lblNomeCombo.Visible = True
   lblNomeCombo.Caption = "Produto"
    optMostrarTodos.Value = True
    optTodosPreco.Value = True
    optCompleto.Visible = True
    optPorIniciais.Visible = True
    optPorPalavra.Visible = True
    optPalavrasDuplas.Visible = True
   cboConsProduto.SetFocus
ElseIf cboCriterios.Text = "FABRICANTE" Then
    cboConsProduto.Visible = True
    lblNomeCombo.Visible = True
    lblNomeCombo.Caption = "Fabricante"
    optTodosPreco.Value = True
    optMostrarTodos.Value = True
    optCompleto.Visible = False
    optPorIniciais.Visible = False
    optPorPalavra.Visible = False
    optPalavrasDuplas.Visible = False
    cboConsProduto.SetFocus
ElseIf cboCriterios.Text = "NCM" Then
    cboConsProduto.Visible = True
    lblNomeCombo.Visible = True
    lblNomeCombo.Caption = "NCM"
    optTodosPreco.Value = True
    optMostrarTodos.Value = True
    optCompleto.Visible = False
    optPorIniciais.Visible = False
    optPorPalavra.Visible = False
    optPalavrasDuplas.Visible = False
    cboConsProduto.SetFocus
End If
End Sub


Private Sub cboCST_Click()
    ' Limpa a redução por padrão ao trocar de CST
    txtRedBCAliquota.Text = "0,00"
    txtRedBCAliquota.Enabled = False
    txtRedBCAliquota.BackColor = &H8000000F ' Cinza (Desabilitado)

If var_RegimeEmpresa = 1 Then
    ' No Simples Nacional, a alíquota de saída no cadastro é SEMPRE 0,00
    txtICMSAliquota.Text = "0,00"
    txtICMSAliquota.Enabled = False
    txtRedBCAliquota.Text = "0,00"
    txtRedBCAliquota.Enabled = False
Else
    Select Case cboCST.Text
        Case "000", "010"
            ' Tributados Integral: Alíquota Interna (22,50)
            txtICMSAliquota.Text = Format(var_AliqInterna, "##0.00")
            
        Case "020", "070"
            ' Redução de Base: Alíquota Interna + Habilita Redução
            txtICMSAliquota.Text = Format(var_AliqInterna, "##0.00")
            txtRedBCAliquota.Enabled = True
            txtRedBCAliquota.BackColor = vbWhite ' Branco (Habilitado para edição)
            txtRedBCAliquota.SetFocus

        Case "040", "041", "060", "050"
            ' Isentos, ST e Não Tributados: Zera tudo
            txtICMSAliquota.Text = "0,00"
            
        Case "090"
            ' Outros: Deixa editar ambos por precaução
            txtICMSAliquota.Text = "0,00"
            txtRedBCAliquota.Enabled = True
            txtRedBCAliquota.BackColor = vbWhite
    End Select
End If
End Sub


Private Sub cboFab_GotFocus()
cboFab.Clear

If vTipoOS = "Automóveis" Then
    sSQL = "SELECT DISTINCT FABRICANTE FROM OS_Fabricantes_Carro ORDER BY FABRICANTE"
ElseIf vTipoOS = "Motocicletas" Then
    sSQL = "SELECT DISTINCT FABRICANTE FROM OS_Fabricante_Moto ORDER BY FABRICANTE"
End If
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFab.AddItem ValidateNull(r("FABRICANTE"))
   r.MoveNext
Loop

moCombo.AttachTo cboFab
End Sub


Private Sub cboFabricante_GotFocus()
Dim vTextoAntes As String

vTextoAntes = cboFabricante.Text
'Limpa a lista
cboFabricante.Clear

sSQL = "SELECT DISTINCT fabricante FROM produtos ORDER BY fabricante;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFabricante.AddItem ValidateNull(r("fabricante"))
   r.MoveNext
Loop

cboFabricante.Text = vTextoAntes

moCombo.AttachTo cboFabricante
End Sub

Private Sub cboFabricante_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFabricante_LostFocus()
cboFabricante.Text = TirarEspaco(cboFabricante.Text)
End Sub

Private Sub cboMes_GotFocus()
Dim vMes As Integer

cboMes.Clear

For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMes
End Sub


Private Sub cboMesPreco_GotFocus()
Dim vMes As Integer

cboMesPreco.Clear

For vMes = 1 To 12
   cboMesPreco.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMesPreco
End Sub


Private Sub cboModelo_GotFocus()
cboModelo.Clear
If vTipoOS = "Automóveis" Then
    sSQL = "SELECT DISTINCT MODELO FROM OS_Modelo_Carro ORDER BY MODELO"
ElseIf vTipoOS = "Motocicletas" Then
    sSQL = "SELECT DISTINCT MODELO FROM OS_Modelo_Moto ORDER BY MODELO"
End If

Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboModelo.AddItem ValidateNull(r("MODELO"))
   r.MoveNext
Loop

moCombo.AttachTo cboModelo
End Sub


Private Sub cboModelo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboOrdem_Click()
cmdExibir_Click
End Sub

Private Sub cboOrdem_GotFocus()
Dim var_Texto As String
var_Texto = cboOrdem.Text

cboOrdem.Clear
cboOrdem.AddItem "DESCRIÇÃO"
cboOrdem.AddItem "CÓD. BARRA"
cboOrdem.AddItem "FABRICANTE"
cboOrdem.AddItem "VENDA"
cboOrdem.AddItem "TOTAL VENDA"
cboOrdem.AddItem "CUSTO"
cboOrdem.AddItem "TOTAL VENDA"
moCombo.AttachTo cboOrdem
   
cboOrdem.Text = var_Texto
cboOrdem.SelStart = 0
cboOrdem.SelLength = Len(cboOrdem)
End Sub


Private Sub cboOrdem_LostFocus()
cmdExibir_Click
End Sub


Private Sub cboOrdem2_Click()
cmdExibir_Click
End Sub

Private Sub cboOrdem2_GotFocus()
Dim var_Texto As String
var_Texto = cboOrdem2.Text

cboOrdem2.Clear
cboOrdem2.AddItem "ASC"
cboOrdem2.AddItem "DESC"
moCombo.AttachTo cboOrdem2
   
cboOrdem2.Text = var_Texto
cboOrdem2.SelStart = 0
cboOrdem2.SelLength = Len(cboOrdem2)
End Sub


Private Sub cboOrdem2_LostFocus()
cmdExibir_Click
End Sub

Private Sub cboProdutoFracionado_GotFocus()
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboProdutoFracionado.Text
codAtual = txtCodProdFracionado.Text
cboProdutoFracionado.Clear

sSQL = "SELECT DISTINCT DESCRICAO, codigo FROM produtos ORDER BY DESCRICAO;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboProdutoFracionado.AddItem r("DESCRICAO")
   cboProdutoFracionado.ItemData(cboProdutoFracionado.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboProdutoFracionado.Text = itemAtual
txtCodProdFracionado.Text = codAtual

SelectControl cboProdutoFracionado

moCombo.AttachTo cboProdutoFracionado
End Sub


Private Sub cboProdutoFracionado_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboProdutoFracionado_LostFocus()
If cboProdutoFracionado.Text = "" Then txtCodProdFracionado.Text = "": Exit Sub
If cboProdutoFracionado.Locked = True Then Exit Sub

On Error GoTo TrataErro

    If cboProdutoFracionado.ListIndex = -1 Then
        txtCodProdFracionado.Text = ""
        cboProdutoFracionado.Text = ""
        Exit Sub
    Else
         txtCodProdFracionado = cboProdutoFracionado.ItemData(cboProdutoFracionado.ListIndex)
         Exit Sub
    End If

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboUnidMedida_GotFocus()
'feito por mim
'Dim var_Texto As String
'var_Texto = cboUnidMedida.Text

'   cboUnidMedida.Clear
'   cboUnidMedida.AddItem "UN"
'   cboUnidMedida.AddItem "CX"
'   cboUnidMedida.AddItem "M"
'   cboUnidMedida.AddItem "M2"
'   cboUnidMedida.AddItem "M3"
'   cboUnidMedida.AddItem "ML"
'   cboUnidMedida.AddItem "KG"
'   cboUnidMedida.AddItem "GR"
 '  cboUnidMedida.AddItem "CT"
 '  cboUnidMedida.AddItem "PO"
'   cboUnidMedida.AddItem "SC"
'   cboUnidMedida.AddItem "PA"
'   cboUnidMedida.AddItem "EX"
'   cboUnidMedida.AddItem "BJ"
'   cboUnidMedida.AddItem "DZ"
'   cboUnidMedida.AddItem "PC"
'   cboUnidMedida.AddItem "DI"
'   cboUnidMedida.AddItem "FD"
'   cboUnidMedida.AddItem "PT"

'   moCombo.AttachTo cboUnidMedida
   
'cboUnidMedida.Text = var_Texto
'SelectControl cboUnidMedida

    ' No foco, apenas anexa o componente de estilo e seleciona o texto (criado por IA)
    'moCombo.AttachTo cboUnidMedida
    'SelectControl cboUnidMedida
End Sub

Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To Grid.rows - 1
      If IsNumeric(Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub cboUnidMedida_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboUnidMedida_LostFocus()
cboUnidMedida.Text = TirarEspaco(cboUnidMedida.Text)
cboUnidMedida.Text = Left(cboUnidMedida.Text, 2)
End Sub

Private Sub ccmdDuplicar_Click()
If Grid.Row = 0 Then MsgBox "Selecione um produto na lista!", vbInformation, "Aviso do Sistema": Exit Sub
SSTab1.Tab = 0
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
vTipoEdicao = "Duplicar"
cmdExcluir.Enabled = False
LimparObjetos_Produtos
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
HabilitarFrames
MostrarDados_Produto
'cmdSalvar.Enabled = False
'cmdCancelar.Enabled = False
txtCodigo.Text = ""
txtCodBarra.Text = ""
txtEAN.Text = ""
txtCodBarra.SetFocus
'vModoEdicao = True
End Sub

Private Sub chameleonButton1_Click()
txtEAN.Text = txtCodBarra.Text
End Sub



Private Sub chkCombustivel_Click()
If txtDescricao.Text = "" Then Exit Sub
If vTipoEdicao = "Edicao" Then
    dbData.Execute "UPDATE Produtos SET COMBUSTIVEL = " & Abs(chkCombustivel.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"
    If chkCombustivel.Value = Checked Then
        frmGas.Visible = True
    Else
        frmGas.Visible = False
    End If
End If
End Sub

Private Sub chkDestaque_Click()
If txtDescricao.Text = "" Then Exit Sub
If vTipoEdicao = "Edicao" Then
    dbData.Execute "UPDATE Produtos SET DESTAQUE = " & Abs(chkDestaque.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"
End If
End Sub


Private Sub chkFracionado_Click()
If txtDescricao.Text = "" Then chkFracionado.Value = Unchecked: Exit Sub
If txtCodigo.Text = "" Or txtCodigo.Text = "0" Then chkFracionado.Value = Unchecked: Exit Sub
'If txtDescricao.Text = "" Then MsgBox "Precisa cadastrar o produto primeiro antes de fracionar!", vbInformation, "Aviso do Sistema": chkFracionado.Value = Unchecked: Exit Sub

'dbData.Execute "UPDATE Produtos SET FRACIONADO = " & Abs(chkFracionado.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"

If chkFracionado.Value = Checked Then
    frmFracionado.Visible = True
Else
    frmFracionado.Visible = False
End If
End Sub

Private Sub chkMateriaPrima_Click()
If txtDescricao.Text = "" Then Exit Sub
If vTipoEdicao = "Edicao" Then
    dbData.Execute "UPDATE Produtos SET MATERIAPRIMA = " & Abs(chkMateriaPrima.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"
End If
End Sub

Private Sub chkUsoConsumo_Click()
If txtDescricao.Text = "" Then Exit Sub
If vTipoEdicao = "Edicao" Then
    dbData.Execute "UPDATE Produtos SET USOCONSUMO = " & Abs(chkUsoConsumo.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"
End If
End Sub

Private Sub chkPedirPeso_Click()
If txtDescricao.Text = "" Then Exit Sub
If vTipoEdicao = "Edicao" Then
    dbData.Execute "UPDATE Produtos SET PEDIRPESO = " & Abs(chkPedirPeso.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"
End If
End Sub

Private Sub Desativa_Objetos()
   If chkProduto.Value = Unchecked Then lblConsProduto.Enabled = False: cboConsProdutoRoupas.Enabled = False Else lblConsProduto.Enabled = True: cboConsProdutoRoupas.Enabled = True
   If chkCodBarra.Value = Unchecked Then lblConsCodBarra.Enabled = False: txtConsCodBarra.Enabled = False Else lblConsCodBarra.Enabled = True: txtConsCodBarra.Enabled = True
   If chkFab.Value = Unchecked Then lblConsFab.Enabled = False: cboConsFab.Enabled = False Else lblConsFab.Enabled = True: cboConsFab.Enabled = True
   If chkRef.Value = Unchecked Then lblConsRef.Enabled = False: cboConsRef.Enabled = False Else lblConsRef.Enabled = True: cboConsRef.Enabled = True
   If chkTam.Value = Unchecked Then lblConsTam.Enabled = False: cboConsTam.Enabled = False Else lblConsTam.Enabled = True: cboConsTam.Enabled = True
   If chkLinha.Value = Unchecked Then lblConsLinha.Enabled = False: cboConsLinha.Enabled = False Else lblConsLinha.Enabled = True: cboConsLinha.Enabled = True
End Sub

Private Sub ckkImobilizado_Click()
If txtDescricao.Text = "" Then Exit Sub
If vTipoEdicao = "Edicao" Then
    dbData.Execute "UPDATE Produtos SET IMOBILIZADO = " & Abs(ckkImobilizado.Value) & " WHERE (codigo = " & txtCodigo.Text & ");"
End If
End Sub

Private Sub cmd_Click()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim Texto As String

sSQL = "SELECT COD_BARRA, DESCRICAO,  (SELECT  TOP (1) VALOR_VV FROM  Produtos_Precos WHERE (COD_PRODUTO = produtos.CODIGO) ORDER BY CODIGO DESC) AS venda FROM  produtos WHERE LEN(COD_BARRA) > 6 ORDER BY DESCRICAO"
Set r = dbData.OpenRecordset(sSQL)

Texto = ""

Open "C:\sistema\arquivo.txt" For Output As #1

Do While Not r.EOF
    Texto = r.GetString(, 1, "|")
    Print #1, Texto
 '  If Not r.EOF Then r.MoveNext
Loop

Close #1
End Sub

Private Sub cmdAddModelo_Click()
Dim vNovoCodigo As Long
If vTipoOS = "Motocicletas" Then
    sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS vCodigo FROM OS_Modelo_Moto;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then vNovoCodigo = r("vCodigo") + 1
    If r.State <> 0 Then r.Close
    Set r = Nothing
       
    sSQL = "INSERT INTO OS_Modelo_Moto (CODIGO, modelo) VALUES (" & vNovoCodigo & ", '" & cboModelo.Text & "');"
ElseIf vTipoOS = "Automóveis" Then
    sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS vCodigo FROM OS_Modelo_Carro;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then vNovoCodigo = r("vCodigo") + 1
    If r.State <> 0 Then r.Close
    Set r = Nothing
       
    sSQL = "INSERT INTO OS_Modelo_Carro (CODIGO, modelo) VALUES (" & vNovoCodigo & ", '" & cboModelo.Text & "');"
End If
dbData.Execute (sSQL)
MsgBox "Modelo Salvo!", vbInformation, "Aviso do Sistema"
End Sub

Private Sub cmdAdicionarComp_Click()
If cmdSalvar.Enabled = False And vTipoEdicao = "Novo" Then Exit Sub
AutoNumeracao_Comp

If txtCodComp.Text = "" Then Exit Sub

'ADICIONAR NA TABELA COMP
If Not Inserir_Dados_Comp Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Mostrar_Comp

txtCodComp.Text = ""
cboFab.Text = ""
cboModelo.Text = ""
cboAno.Text = ""
cboModelo.SetFocus
End Sub

Private Sub AutoNumeracao_Comp()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT MAX(CODIGO) as ULTIMO FROM PRODUTOS_COMP"
Set r = dbData.OpenRecordset(sSQL)

txtCodComp.Text = IIf(IsNull(r!ULTIMO) = True, 1, r!ULTIMO + 1)
End Sub

Private Sub cmdAdicionarReferencia_Click()
If txtReferencia.Text = "" Or txtCodBarra.Text = "" Then Exit Sub

'ADICIONAR NA TABELA OS SERVIÇOS
AutoNumeracao_Referencia
dbData.Execute "INSERT INTO Produtos_Referencias (codigo, cod_produto, referencia, descontinuado, COD_BARRA) VALUES(" & vUltimaReferencia & ", " & txtCodigo.Text & ", '" & txtReferencia.Text & "', 0, '" & txtCodBarra.Text & "')"

MostrarGrid_Referencia

txtReferencia.Text = ""
txtReferencia.SetFocus
End Sub


Private Sub AutoNumeracao_Referencia()
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo FROM Produtos_referenciaS;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then vUltimaReferencia = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub


Private Sub MostrarGrid_Referencia()
Dim vCodProdRef As String
If txtCodBarra.Text = "" Then
    vCodProdRef = 0
Else
    vCodProdRef = txtCodBarra.Text
End If

sSQL = "SELECT * FROM Produtos_Referencias WHERE (cod_barra = '" & vCodProdRef & "') AND DESCONTINUADO = 0;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Referencia r

If r.State <> 0 Then r.Close

sSQL = "SELECT * FROM Produtos_Referencias WHERE (cod_barra = '" & vCodProdRef & "') AND DESCONTINUADO = 1;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Referencia_Desc r

If r.State <> 0 Then r.Close
End Sub

Private Sub cmdAlterar_Click()


If txtEAN.Text <> "" Then
    If Len(txtEAN.Text) < 6 Then MsgBox "O EAN não pode ser um codigo criado", vbInformation, "Aviso do Sistema": txtEAN.Text = "": Exit Sub
End If

If txtCodigo.Text = "" Then
   ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o produto na guia CONSULTA.", vbInformation
   Exit Sub
End If

If txtCodBarra.Text = "" Then MsgBox "Não será permitido cadastrar produto sem código de barra", vbInformation, "Aviso do Sistema": Exit Sub

If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

If chkFracionado.Value = Checked And cboProdutoFracionado.Text = "" Then ShowMsg "Você selecionou a opção de fracionamento de produto!." & vbCr & "Escolha um produto ou desmarque a opção.", vbExclamation: Exit Sub

'alterar o nome dos produtos da tabela de entrada de pedidos
dbData.Execute "UPDATE produtos_entrada_itens SET descricao = '" & txtDescricao.Text & "' WHERE (codigo_produto = " & txtCodigo.Text & ");"

'alterar o nome dos produtos da tabela de entrada de pedidos
'sSQL = "UPDATE produtos_entrada_itens SET VENDA = " & Replace(CCur(txtValorAtual.Text), ",", ".") & " WHERE (codigo = " & _
'   "(SELECT codigo FROM (SELECT TOP 1 codigo FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodigo.Text & ") ORDER BY codigo DESC) as tempTabela));"
'dbData.Execute sSQL

'alterar alguns campos da tabela TbNFCe_Itens
dbData.Execute "UPDATE TbNFCe_Itens SET DescricaoProduto = '" & txtDescricao.Text & "', CodNcm = '" & txtNCM.Text & "', cfop = '" & cboCFOP.Text & "', ICMSCST = '" & cboCST.Text & "', CodBarras = '" & txtEAN.Text & "' WHERE (IDProduto = " & txtCodigo.Text & ");"

'alterar alguns campos da tabela NotaFiscalItens
'dbData.Execute "UPDATE NotaFiscalItens SET NomeProduto = '" & txtDescricao.Text & "', NCM = '" & txtNCM.Text & "', EAN = '" & txtEAN.Text & "'  WHERE (CodigoProduto = " & txtCodigo.Text & ");"
    'Mudar as referencias pra o codigo do produto criado
dbData.Execute "Update tb1 " & _
                "Set tb1.NomeProduto=tb2.DESCRICAO, tb1.EAN=tb2.EAN, tb1.NCM=tb2.NCM, tb1.CFOP=tb2.CFOP, tb1.UnidadeComercial=tb2.Unid_medida, tb1.CST=tb2.icmsCST, tb1.PISCST=tb2.PISCST, tb1.COFINSCST=tb2.COFINSCST " & _
                "FROM NotaFiscalItens as tb1 INNER JOIN NotaFiscal as tb0 ON tb1.CodigoNota = tb0 .CodigoNota INNER JOIN produtos as tb2 ON tb1.CodigoProduto = tb2 .CODIGO" & _
                "WHERE (tb0 .Enviada = 0)"
 
DesabilitarFrames
DesabilitarBotoes
LimparObjetos_Produtos
LimparObjeto_Gas
LimparGrid_Ref
LimparGrid_Comp

cmdExibir_Click
Mostrar_HistoricoQuant
End Sub

Private Function Inserir_Dados_Comp() As Boolean
Dim sSQL As String
Dim vJuncao As String

vJuncao = cboFab.Text & " / " & cboModelo.Text
'Comando de inclusão
sSQL = "INSERT INTO produtos_comp (" & _
   "Codigo, COD_PRODUTO, MODELO, ANO, COD_BARRA) VALUES (" & _
   txtCodComp.Text & ", " & txtCodigo.Text & ", '" & vJuncao & "', '" & cboAno.Text & "', '" & txtCodBarra.Text & "');"

'Retorna o resultado da inclusão
Inserir_Dados_Comp = dbData.Execute(sSQL)
End Function

Private Function Inserir_Dados_Gas() As Boolean
If Trim(txtValorPartida.Text) = "" Then txtValorPartida.Text = Format(0, ocMONEY)

Dim vpGLP As Double
Dim vpGNn As Double
Dim vpGNi As Double
Dim vpMixGN As Double

If txtpGLP.Text = "" Or txtpGLP.Text = "0" Then txtpGLP.Text = FormatNumber(0, 2) & "%"
If txtpGNn.Text = "" Or txtpGNn.Text = "0" Then txtpGNn.Text = FormatNumber(0, 2) & "%"
If txtpGNi.Text = "" Or txtpGNi.Text = "0" Then txtpGNi.Text = FormatNumber(0, 2) & "%"
If txtpMixGN.Text = "" Or txtpMixGN.Text = "0" Then txtpMixGN.Text = FormatNumber(0, 2) & "%"

vpGLP = Left$(txtpGLP.Text, Len(txtpGLP.Text) - 1)
vpGNn = Left$(txtpGNn.Text, Len(txtpGNn.Text) - 1)
vpGNi = Left$(txtpGNi.Text, Len(txtpGNi.Text) - 1)
vpMixGN = Left$(txtpMixGN.Text, Len(txtpMixGN.Text) - 1)

'Comando de inclusão
sSQL = "INSERT INTO Produtos_Gas (" & _
   "Cod_Produto, CODIF, cProdANP, descricaoANP, pGLP, pGNi, pGNn, pMixGN, ValorPartida) VALUES (" & _
   txtCodigo.Text & ",  " & txtCODIF.Text & ", " & txtcProdANP.Text & ", '" & txtdescricaoANP.Text & "', " & _
   Replace(CDbl(vpGLP), ",", ".") & ", " & Replace(CDbl(vpGNi), ",", ".") & ", " & Replace(CDbl(vpGNn), ",", ".") & ", " & Replace(CDbl(vpMixGN), ",", ".") & ", " & _
   Replace(CCur(txtValorPartida.Text), ",", ".") & ");"

'Retorna o resultado da inclusão
Inserir_Dados_Gas = dbData.Execute(sSQL)
End Function
Private Function Inserir_Dados() As Boolean
'Valida os campos
If Trim(txtQuant.Text) = "" Then txtQuant.Text = 0
If Trim(txtQuantMin.Text) = "" Then txtQuantMin.Text = 0
If Trim(txtICMSAliquota.Text) = "" Then txtICMSAliquota.Text = 0

Dim varMargemVV As Double
Dim varMargemVP As Double
Dim varMargemAV As Double
Dim varMargemAP As Double

If txtMargemVV.Text = "" Or txtMargemVV.Text = "0" Then txtMargemVV.Text = FormatNumber(0, 2) & "%"
If txtMargemVP.Text = "" Or txtMargemVP.Text = "0" Then txtMargemVP.Text = FormatNumber(0, 2) & "%"
If txtMargemAV.Text = "" Or txtMargemAV.Text = "0" Then txtMargemAV.Text = FormatNumber(0, 2) & "%"
If txtMargemAP.Text = "" Or txtMargemAP.Text = "0" Then txtMargemAP.Text = FormatNumber(0, 2) & "%"

varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)
varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)
varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)
varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)

'Comando de inclusão
sSQL = "INSERT INTO produtos (" & _
   "codigo, ativo, destaque, USOCONSUMO, COMBUSTIVEL, MATERIAPRIMA, IMOBILIZADO, FRACIONADO, cod_barra, ean, descricao, fabricante, unid_medida, " & _
   "categoria, PRATELEIRA, quant_min, INF_ADICIONA, quant_estoque, ref, tamanho, ICMSCST, ICMSAliq, PISCST, COFINSCST, IPICST, NCM, CEST, CFOP, Alterado, PedirPeso, IPIALIQ, COFINSALIQ, PISALIQ, pRedBc, CODPROD_FRACAO, QUANT_FRACAO) VALUES (" & _
   txtCodigo.Text & ", 1, " & Abs(chkDestaque.Value) & ", " & Abs(chkUsoConsumo.Value) & ", " & Abs(chkCombustivel.Value) & ", " & Abs(chkMateriaPrima.Value) & ", " & Abs(ckkImobilizado.Value) & ", " & Abs(chkFracionado.Value) & ", '" & _
   txtCodBarra.Text & "', '" & txtEAN.Text & "', '" & _
   txtDescricao.Text & "', '" & cboFabricante.Text & "', '" & cboUnidMedida.Text & "', '" & _
   cboCategoria.Text & "', '" & txtPrateleira.Text & "', " & Replace(CDbl(txtQuantMin.Text), ",", ".") & ", '" & _
   txtObs.Text & "', " & Replace(CDbl(txtQuant.Text), ",", ".") & ", '" & txtRef.Text & "', '" & txtTam.Text & "', '" & IIf((cboCST.Text = ""), 0, cboCST.Text) & "', " & Replace(CDbl(txtICMSAliquota.Text), ",", ".") & ", '" & txtPISCST.Text & "', '" & _
   txtCOFINSCST.Text & "', '" & txtIPICST.Text & "', '" & IIf((txtNCM.Text = ""), 0, txtNCM.Text) & "', '" & txtCEST.Text & "', '" & IIf((cboCFOP.Text = ""), 0, cboCFOP.Text) & "', 0, " & Abs(chkPedirPeso.Value) & ", " & Replace(CDbl(txtIPIAliquota.Text), ",", ".") & ", " & Replace(CDbl(txtCofinsAliquota.Text), ",", ".") & ", " & Replace(CDbl(txtPisAliquota.Text), ",", ".") & ", " & Replace(CDbl(txtRedBCAliquota.Text), ",", ".") & ", " & IIf((chkFracionado.Value = Checked), txtCodProdFracionado.Text, 0) & ", " & Replace(CDbl(IIf((chkFracionado.Value = Checked), txtQuantFracionado.Text, 0)), ",", ".") & ");"
'Debug.Print sSQL
'Retorna o resultado da inclusão
Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
Dim sSQL As String

'Comando de atualização
sSQL = "UPDATE produtos SET " & _
   "destaque = " & Abs(chkDestaque.Value) & ", " & _
   "USOCONSUMO = " & Abs(chkUsoConsumo.Value) & ", " & _
   "COMBUSTIVEL = " & Abs(chkCombustivel.Value) & ", " & _
   "MATERIAPRIMA = " & Abs(chkMateriaPrima.Value) & ", " & _
   "IMOBILIZADO = " & Abs(ckkImobilizado.Value) & ", " & _
   "FRACIONADO = " & Abs(chkFracionado.Value) & ", " & _
   "pedirpeso = " & Abs(chkPedirPeso.Value) & ", " & _
   "cod_barra = '" & txtCodBarra.Text & "', " & _
   "ean = '" & txtEAN.Text & "', " & _
   "descricao = '" & txtDescricao.Text & "', " & _
   "fabricante = '" & cboFabricante.Text & "', " & _
   "unid_medida = '" & cboUnidMedida.Text & "', " & _
   "categoria = '" & cboCategoria.Text & "', " & _
   "tamanho = '" & txtTam.Text & "', " & _
   "ref = '" & txtRef.Text & "', " & _
   "PRATELEIRA = '" & txtPrateleira.Text & "', " & _
   "quant_min = " & Replace(CDbl(txtQuantMin.Text), ",", ".") & ", " & _
   "INF_ADICIONA = '" & txtObs.Text & "', " & _
   "quant_estoque = " & Replace(CDbl(txtQuant.Text), ",", ".") & ", "
   sSQL = sSQL & _
   "ICMSAliq = " & Replace(CDbl(txtICMSAliquota.Text), ",", ".") & ", " & _
   "pisAliq = " & Replace(CDbl(txtPisAliquota.Text), ",", ".") & ", " & _
   "cOFINSAliq = " & Replace(CDbl(txtCofinsAliquota.Text), ",", ".") & ", " & _
   "ipiAliq = " & Replace(CDbl(txtIPIAliquota.Text), ",", ".") & ", " & _
   "pRedBc = " & Replace(CDbl(txtRedBCAliquota.Text), ",", ".") & ", " & _
   "PISCST = '" & txtPISCST.Text & "', " & _
   "COFINSCST = '" & txtCOFINSCST.Text & "', " & _
   "CODPROD_FRACAO = " & IIf((chkFracionado.Value = Checked), txtCodProdFracionado.Text, 0) & ", " & _
   "QUANT_FRACAO = " & Replace(CDbl(IIf((chkFracionado.Value = Checked), txtQuantFracionado.Text, 0)), ",", ".") & ", " & _
   "IPICST = '" & txtIPICST.Text & "', " & _
   "NCM = '" & txtNCM.Text & "', " & _
   "CFOP = " & cboCFOP.Text & ", " & _
   "ICMSCST = '" & cboCST.Text & "', " & _
   "CEST = '" & txtCEST.Text & "'"
   
'Condição para atualização
sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ");"

'Retorna o resultado da atualização
Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdApagar_Click()
If Grid.Row = 0 Then MsgBox "Selecione um produto na lista!", vbInformation, "Aviso do Sistema": Exit Sub
Dim bRet As Boolean

'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub

'If txtCodigo.Text = "" Then ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o produto na guia CONSULTA", vbInformation: Exit Sub
If Grid.TextMatrix(Grid.Row, 2) = "00001" Then ShowMsg "SEM PERMISSÃO!" & vbCrLf & "Você não tem permissão para excluir esse produto", vbInformation: Exit Sub

sSQL = "SELECT * " & _
       "FROM pedidos_itens " & _
       "WHERE (COD_PRODUTO = " & Grid.TextMatrix(Grid.Row, 1) & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    ShowMsg "Não é permitido excluir esse produto!" & vbCrLf & "Esse produto já foi usado em vendas anteriores!", vbInformation
    Exit Sub
End If

'Solicita ao usuário confirmação da exclusão
If ShowMsg("Deseja excluir o produto " & Grid.TextMatrix(Grid.Row, 3) & " ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
 
'Faz a exclusão usando o comando DELETE do SQL
sSQL = "DELETE FROM produtos WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 1) & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

cmdExibir_Click
End Sub

Private Sub cmdBuscarCEST_Click()
Dim varNomeProduto As String
varNomeProduto = txtNCM.Text
'ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
ShellExecute hwnd, "open", "http://www.buscacest.com.br/?utf8=" + Chr(95) + "&ncm=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo

'http://www.buscacest.com.br/?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdCancelar_Click()
vTipoEdicao = "Cancela"
LimparObjetos_Produtos
Mostrar_Precos
Mostrar_HistoricoQuant
LimparObjeto_Gas
DesabilitarFrames
DesabilitarBotoes
LimparGrid_Ref
LimparGrid_Comp
End Sub

Private Sub cmdConsultarNCM_Click()
Dim varNomeProduto As String
varNomeProduto = Replace(txtDescricao, " ", "+")
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo


End Sub

Private Sub cmdConsultarNCMean_Click()
Dim varNomeProduto As String
varNomeProduto = txtEAN.Text
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub


Private Sub cmdDesativar_Click()
i = Grid.Row
If Grid.Row = 0 Then MsgBox "Selecione um produto na lista!", vbInformation, "Aviso do Sistema": Exit Sub

If Grid.TextMatrix(i, 12) = "ATIVO" Then
    dbData.Execute "UPDATE Produtos SET ativo = 0 WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 1) & ");"
Else
    dbData.Execute "UPDATE Produtos SET ativo = 1 WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 1) & ");"
End If

cmdExibir_Click
End Sub

Private Sub cmdDescReferencia_Click()
On Error GoTo erro

If Not IsNumeric(Grid_Referencia.TextMatrix(Grid_Referencia.Row, 1)) = True Then GoSub erro
If ShowMsg("Deseja descontinuar a referência: " & Grid_Referencia.TextMatrix(Grid_Referencia.Row, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

dbData.Execute "UPDATE Produtos_Referencias SET DESCONTINUADO = 1 WHERE (codigo = " & Grid_Referencia.TextMatrix(Grid_Referencia.Row, 1) & ") AND (cod_produto = " & txtCodigo.Text & ");"

MostrarGrid_Referencia
Exit Sub
   
erro:
   ShowMsg "Não existe nenhuma referência para ser descontinuada!", vbExclamation
   Exit Sub
End Sub

Private Sub cmdEditar_Click()
If Grid.Row = 0 Then MsgBox "Selecione um produto na lista!", vbInformation, "Aviso do Sistema": Exit Sub
vTipoEdicao = "Edicao"
SSTab1.Tab = 0
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
'cmdAlterar.Enabled = True
'cmdExcluir.Enabled = True
'txtCodigo.Text = ""
'vModoEdicao = True
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub cmdExcluir_Click()
Dim sSQL As String
Dim bRet As Boolean

'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub

If txtCodigo.Text = "" Then ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o produto na guia CONSULTA", vbInformation: Exit Sub
If txtCodBarra.Text = "00001" Then ShowMsg "SEM PERMISSÃO!" & vbCrLf & "Você não tem permissão para excluir esse produto", vbInformation: Exit Sub

Dim r As ADODB.Recordset
sSQL = "SELECT * " & _
       "FROM pedidos_itens " & _
       "WHERE (COD_PRODUTO = " & txtCodigo.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    ShowMsg "Não é permitido excluir esse produto!" & vbCrLf & "Esse produto já foi usado em vendas posteriores", vbInformation
    Exit Sub
End If

'Solicita ao usuário confirmação da exclusão
If ShowMsg("Excluir esse produto?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
 
'dbData.Execute "UPDATE produtos SET desabilitado = 1 WHERE (codigo = " & txtCodigo.Text & ");"

'Faz a exclusão usando o comando DELETE do SQL
sSQL = "DELETE FROM produtos WHERE (codigo = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

'sSQL = "DELETE FROM produtos_comp WHERE (cod_produto = " & txtCodigo.Text & ");"
'dbData.Execute (sSQL)

LimparObjetos_Produtos

If frmComp.Visible = True Then LimparGrid_Comp

DesabilitarFrames
DesabilitarBotoes
LimparGrid_Produtos
LimparObjeto_Gas
LimparGrid_Ref
LimparGrid_Comp
Mostrar_HistoricoQuant
End Sub

Private Sub cmdExibirPreco_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

'Monta a consulta básica
sSQL = "SELECT * " & _
   "FROM produtos_precos "

'Define o filtro
If txtCodigo.Text = "" Then
   sSQL = sSQL & "WHERE 1 = 0 "
   
Else
   sSQL = sSQL & "WHERE (cod_produto = " & txtCodigo.Text & ") and (MONTH(data) = " & cboMesPreco.ListIndex + 1 & ") AND (YEAR(data) = " & cboAnoPreco & ")"
End If

'Monta a ordem de exibição
sSQL = sSQL & "ORDER BY codigo "

Set r = dbData.OpenRecordset(sSQL)
FormatarGrid_Precos r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cmdExibirQuant_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

'Monta a consulta básica
sSQL = "SELECT Produtos_Quant.*, produtos_entrada.* " & _
       "FROM Produtos_Quant LEFT JOIN produtos_entrada ON Produtos_Quant.cod_entrada = produtos_entrada.codigo "

'Define o filtro
If txtCodigo.Text = "" Then
   sSQL = sSQL & "WHERE 1 = 0 "
Else
   sSQL = sSQL & "WHERE (cod_produto = " & txtCodigo.Text & ") and (MONTH(data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(data) = " & cboAnoCons & ")"
End If

'Monta a ordem de exibição
sSQL = sSQL & "ORDER BY data "
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Quant r
SomaQuantAdicao
SomaQuantRemocao

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub



Private Sub cmdImprimir_Click()
   Dim r As ADODB.Recordset
   Dim var_Impressora As String
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   
   Set oIni = Nothing
   
   Me.Hide
   
   Set r = dbData.OpenRecordset(printSQL)
   
   Dim cCfg As ConfigItem
   Dim tipoEmpresa As Integer
   Set cCfg = sysConfig("TIPO_EMPRESA")
   tipoEmpresa = cCfg.Value
   Set cCfg = Nothing
   
If tipoEmpresa = 4 Then
   Set REL_Prod_Cad_Imp.Relatorio.Recordset = r
   REL_Prod_Cad_Imp.rfTipo.Caption = lblProdutos.Caption
   'REL_Prod_Cad_Imp.rfITENS.Caption = lblTotalUnid.Caption
   REL_Prod_Cad_Imp.rfVENDA.Caption = lblValorTotal.Caption
   
   'REL_Produtos.Relatorio.NomeImpressora = var_Impressora
   REL_Prod_Cad_Imp.Relatorio.Ativar
   Unload REL_Prod_Cad_Imp
Else
   Set REL_Produtos.Relatorio.Recordset = r
   REL_Produtos.rfTipo.Caption = lblTipos.Caption
   REL_Produtos.rfITENS.Caption = lblProdutos.Caption
   REL_Produtos.rfVENDA.Caption = lblValorTotal.Caption
   REL_Produtos.rfCUSTO.Caption = lblValorTotalCusto.Caption
   REL_Produtos.rfData.Caption = Format(Date, "dd/mm/yyyy")
   REL_Produtos.rfHora.Caption = Format(Time(), "HH:MM:ss")
   'REL_Produtos.Relatorio.NomeImpressora = var_Impressora
   REL_Produtos.Relatorio.Ativar
   Unload REL_Produtos
End If
   Me.Show 1
End Sub

Private Sub cmdNovo_Click()
HabilitarFrames
vTipoEdicao = "Novo" 'desativei para teste
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
'vTipoEdicao = "Novo"
cmdExcluir.Enabled = False
frmGas.Visible = False

LimparObjetos_Produtos
LimparObjeto_Gas
LimparGrid_Ref
LimparGrid_Comp
txtCodigo.Text = "0"
'AutoNumeracao

'If Not Inserir_Dados Then
'   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
'   Exit Sub
'End If

If frmComp.Visible = True Then LimparGrid_Comp

cboUnidMedida.Text = "UN"
txtQuant.Text = "0"
txtCodBarra.SetFocus
End Sub

Private Sub cmdRemoverComp_Click()
If cmdSalvar.Enabled = False And vTipoEdicao = "Novo" Then Exit Sub
On Error GoTo erro
Dim bRet As Boolean
Dim sSQL As String

If Not IsNumeric(Grid_Comp.TextMatrix(Grid_Comp.Row, 1)) = True Then GoSub erro
If MsgBox("Deseja remover o composto: " & Grid_Comp.TextMatrix(Grid_Comp.Row, 3) & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso do Sistema") = vbNo Then Exit Sub

'Faz a exclusão usando o comando DELETE do SQL
sSQL = "DELETE FROM PRODUTOS_COMP WHERE CODIGO = " & Grid_Comp.TextMatrix(Grid_Comp.Row, 1) & " AND COD_PRODUTO = " & txtCodigo.Text
bRet = dbData.Execute(sSQL)
   
If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If
   
Mostrar_Comp

Exit Sub

erro:
MsgBox "Não existe nenhum acessório para ser excluido!", vbExclamation, "Aviso do Sistema"
Exit Sub
End Sub

Private Sub Mostrar_Comp()
If txtCodComp.Text = "" Then txtCodComp.Text = 0
    
Dim sSQL As String
Dim Rs As ADODB.Recordset

'Monta a consulta básica
sSQL = "Select * FROM PRODUTOS_COMP WHERE COD_BARRA = '" & txtCodBarra.Text & "'"

Set Rs = dbData.OpenRecordset(sSQL)

FormatarGrid_Comp Rs

If Rs.State <> 0 Then Rs.Close
Set Rs = Nothing
End Sub

Private Sub cmdRemoverReferencia_Click()
On Error GoTo erro

If Not IsNumeric(Grid_Referencia.TextMatrix(Grid_Referencia.Row, 1)) = True Then GoSub erro
If ShowMsg("Deseja remover a referência: " & Grid_Referencia.TextMatrix(Grid_Referencia.Row, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

dbData.Execute "DELETE FROM Produtos_Referencias WHERE (codigo = " & Grid_Referencia.TextMatrix(Grid_Referencia.Row, 1) & ") AND (cod_produto = " & txtCodigo.Text & ");"

MostrarGrid_Referencia
Exit Sub
   
erro:
   ShowMsg "Não existe nenhuma referência para ser excluida!", vbExclamation
   Exit Sub
End Sub

Private Sub FormatarGrid_Referencia_Desc(rTabela As ADODB.Recordset)
Dim i As Integer, j As Integer
Dim x As Integer

With Grid_Referencia_Desc
   .Clear
   .Cols = 3
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 1600
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "REFERÊNCIA"
   
   'colocar os cabeçalho em negrito
   For x = 0 To .Cols - 1
      .Col = x
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
         
         .TextMatrix(.rows - 1, 1) = rTabela("CODIGO")
         .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("REFERENCIA"))
        
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .Redraw = True
   .rows = .rows - 1
End With
End Sub
Private Sub FormatarGrid_Referencia(rTabela As ADODB.Recordset)
Dim i As Integer, j As Integer
Dim x As Integer

With Grid_Referencia
   .Clear
   .Cols = 3
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 1600
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "REFERÊNCIA"
   
   'colocar os cabeçalho em negrito
   For x = 0 To .Cols - 1
      .Col = x
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
         
         .TextMatrix(.rows - 1, 1) = rTabela("CODIGO")
         .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("REFERENCIA"))
        
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .Redraw = True
   .rows = .rows - 1
End With
End Sub


Private Sub cmdRepetir_Click()
txtMargemVP.Text = txtMargemVV.Text
txtMargemAV = txtMargemVV.Text
txtMargemAP = txtMargemVV.Text
CalcularPrecos
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
If txtEAN.Text <> "" Then
    If Len(txtEAN.Text) < 6 Then MsgBox "O EAN não pode ser um codigo criado", vbInformation, "Aviso do Sistema": txtEAN.Text = "": Exit Sub
End If

If txtCodBarra.Text = "" Then MsgBox "Não será permitido cadastrar produto sem código de barra", vbInformation, "Aviso do Sistema": Exit Sub
If txtDescricao.Text = "" Then ShowMsg "Digite a Descrição do produto", vbInformation: txtDescricao.SetFocus: Exit Sub

If vTipoEdicao = "Novo" Or vTipoEdicao = "Duplicar" Then
    
    If txtMargemVV.Text = "" Or txtMargemVP.Text = "" Or txtMargemAV.Text = "" Or txtMargemAP.Text = "" Then
       ShowMsg "Produtos estão sem margens de vendas", vbInformation
       txtCusto.SetFocus
       Exit Sub
    End If
    
    'If txtCodBarra.Text = "" Then MsgBox "Não será permitido cadastrar produto sem código de barra", vbInformation, "Aviso do Sistema": Exit Sub
    
    AutoNumeracao
    
    'Faz a inserção de forma direta e verifica se houve algum erro
    If Not Inserir_Dados Then
       ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
       Exit Sub
    End If
    
    Preco_Entrada
    Quant_Entrada
    
    'Mudar as referencias pra o codigo do produto criado
    dbData.Execute "UPDATE Produtos_Referencias SET cod_produto = " & txtCodigo.Text & " WHERE (cod_barra = '" & txtCodBarra.Text & "');"
    
    'Mudar as referencias pra o codigo do produto criado
    'dbData.Execute "Update tb1 " & _
                    "Set tb1.EAN=tb2.EAN, tb1.NCM=tb2.NCM, tb1.CFOP=tb2.CFOP, tb1.UnidadeComercial=tb2.Unid_medida, tb1.CST=tb2.icmsCST, tb1.PISCST=tb2.PISCST, tb1.COFINSCST=tb2.COFINSCST " & _
                    "FROM NotaFiscalItens as tb1 INNER JOIN NotaFiscal as tb0 ON tb1.CodigoNota = tb0 .CodigoNota INNER JOIN produtos as tb2 ON tb1.CodigoProduto = tb2 .CODIGO" & _
                    "WHERE (tb0 .Enviada = 0)"

    'Mudar as referencias pra o codigo do produto criado
    dbData.Execute "UPDATE Produtos_COMP SET cod_produto = " & txtCodigo.Text & " WHERE (cod_barra = '" & txtCodBarra.Text & "');"
    
    If chkCombustivel.Value = Checked Then
        If Not Inserir_Dados_Gas Then
           ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
           Exit Sub
        End If
    End If
    
    LimparGrid_Produtos
    LimparObjeto_Gas
    Mostrar_HistoricoQuant
    LimparGrid_Ref
    LimparGrid_Comp
    
    DesabilitarFrames
    DesabilitarBotoes
    LimparObjetos_Produtos
ElseIf vTipoEdicao = "Edicao" Then
    'If txtEAN.Text <> "" Then
    '    If Len(txtEAN.Text) < 6 Then MsgBox "O EAN não pode ser um codigo criado", vbInformation, "Aviso do Sistema": txtEAN.Text = "": Exit Sub
    'End If
    
    If txtCodigo.Text = "" Then ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte o produto na guia CONSULTA.", vbInformation: Exit Sub
    
    If Not Atualizar_Dados Then
       ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
       Exit Sub
    End If
    
    If chkFracionado.Value = Checked And cboProdutoFracionado.Text = "" Then ShowMsg "Você selecionou a opção de fracionamento de produto!." & vbCr & "Escolha um produto ou desmarque a opção.", vbExclamation: Exit Sub
    
    'alterar o nome dos produtos da tabela de entrada de pedidos
    dbData.Execute "UPDATE produtos_entrada_itens SET descricao = '" & txtDescricao.Text & "' WHERE (codigo_produto = " & txtCodigo.Text & ");"
    
    'alterar alguns campos da tabela TbNFCe_Itens
    dbData.Execute "UPDATE TbNFCe_Itens SET DescricaoProduto = '" & txtDescricao.Text & "', CodNcm = '" & txtNCM.Text & "', cfop = '" & cboCFOP.Text & "', ICMSCST = '" & cboCST.Text & "', CodBarras = '" & txtEAN.Text & "' WHERE (IDProduto = " & txtCodigo.Text & ");"
    
    'alterar alguns campos da tabela NotaFiscalItens
    'dbData.Execute "UPDATE NotaFiscalItens SET NomeProduto = '" & txtDescricao.Text & "', NCM = '" & txtNCM.Text & "', EAN = '" & txtEAN.Text & "'  WHERE (CodigoProduto = " & txtCodigo.Text & ");"
    dbData.Execute "Update tb1 " & _
                "Set tb1.NomeProduto = tb2.DESCRICAO, tb1.EAN = tb2.EAN, tb1.NCM = tb2.NCM, tb1.CFOP = tb2.CFOP, tb1.UnidadeComercial = tb2.Unid_medida, tb1.CST = tb2.icmsCST, tb1.PISCST = tb2.PISCST, tb1.COFINSCST = tb2.COFINSCST " & _
                "FROM NotaFiscalItens as tb1 INNER JOIN NotaFiscal as tb0 ON tb1.CodigoNota = tb0 .CodigoNota INNER JOIN produtos as tb2 ON tb1.CodigoProduto = tb2.CODIGO " & _
                "WHERE (tb0.Enviada = 0)"
     
    DesabilitarFrames
    DesabilitarBotoes
    LimparObjetos_Produtos
    Mostrar_Precos
    Mostrar_HistoricoQuant
    LimparObjeto_Gas
    LimparGrid_Ref
    LimparGrid_Comp
    
    cmdExibir_Click
End If
End Sub

Private Sub cmdExibir_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

Dim tipoEmpresa As Integer
            
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

Dim varProdutoHabilitado As String

If cboCriterios.Text = "TODOS" Then
    If optDesabilitados.Value = True Then
        varProdutoHabilitado = "produtos.ativo = 0 and "
    ElseIf optHabilitado.Value = True Then
        varProdutoHabilitado = "produtos.ativo = 1 and "
    End If
Else
        varProdutoHabilitado = "produtos.codigo > 1 and "
End If

If cboCriterios.Text = "TODOS" Or cboCriterios.Text = "CATEGORIA" Then
    If optMostrarQuant.Value = True Then
        varTipoMostrar = "  produtos.quant_estoque > 0 AND "
    ElseIf optMostrarNegativos.Value = True Then
        varTipoMostrar = "  produtos.quant_estoque < 0 AND "
    ElseIf optMostrarZerados.Value = True Then
        varTipoMostrar = "  produtos.quant_estoque = 0 AND "
    ElseIf optMostrarTodos.Value = True Then
        varTipoMostrar = " "
    End If
Else
        varTipoMostrar = "produtos.codigo > 1 and "
End If

Dim vDescricao As String
vDescricao = ""

If optCompleto.Value = True Then
    vDescricao = "(produtos.descricao = '" & cboConsProduto.Text & "')"
ElseIf optPorIniciais.Value = True Then
    vDescricao = "(produtos.descricao  LIKE '%" & cboConsProduto.Text & "')"
ElseIf optPorPalavra.Value = True Then
    vDescricao = "(produtos.descricao  LIKE  '%" & cboConsProduto.Text & "%')"
ElseIf optPalavrasDuplas.Value = True Then
'    vDescricao = ?
End If

'If tipoEmpresa = 4 Then   'sapataria
'       'INDICE
'       Dim var_Criterio As String
'       var_Criterio = ""
'       var_Criterio = var_Criterio & IIf(ckkORDRef.Value, IIf(var_Criterio <> "", ", ", "") & "produtos.ref", "")
'       var_Criterio = var_Criterio & IIf(ckkORDDesc.Value, IIf(var_Criterio <> "", ", ", "") & "produtos.descricao", "")
'       var_Criterio = var_Criterio & IIf(ckkORDFab.Value, IIf(var_Criterio <> "", ", ", "") & "produtos.fabricante", "")
'       var_Criterio = var_Criterio & IIf(ckkORDTam.Value, IIf(var_Criterio <> "", ", ", "") & "produtos.tamanho", "")
'       var_Criterio = var_Criterio & IIf(ckkORDLinha.Value, IIf(var_Criterio <> "", ", ", "") & "produtos.categoria", "")
'       var_Criterio = var_Criterio & IIf(ckkORDQuant.Value, IIf(var_Criterio <> "", ", ", "") & "produtos.quant_estoque", "")
       
'       If var_Criterio <> "" Then var_Criterio = " ORDER BY " & var_Criterio
       
'   If chkTodos.Value = Checked Then
'       sSQL = "SELECT DISTINCT produtos.ref AS var_Ref, produtos.fabricante AS var_Fab, produtos.tamanho AS var_Tam, produtos.categoria AS var_linha, produtos.codigo AS var_codEnt, produtos.cod_barra AS var_CodBarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Prat, produtos.unid_medida AS var_Med, " & _
'         "produtos.quant_estoque AS var_Quant, (SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) AS venda " & _
'         "FROM produtos  " & _
'         "WHERE (produtos.codigo <> 1) " & varTipoMostrar & " " & var_Criterio
      
'      Set r = dbData.OpenRecordset(sSQL)
'      FormatarGrid_Produtos r
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
       
'   ElseIf chkProduto.Value = Checked Then
'       sSQL = "SELECT produtos.ref AS var_Ref, produtos.fabricante AS var_Fab, produtos.tamanho AS var_Tam, produtos.categoria AS var_linha, produtos.codigo AS var_codEnt, produtos.cod_barra AS var_CodBarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Prat, produtos.unid_medida AS var_Med, " & _
'         "produtos.quant_estoque AS var_Quant, ((SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) * produtos.quant_estoque) as var_Total, (SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) AS venda " & _
'         "FROM produtos  " & _
'         "WHERE (produtos.descricao = '" & cboConsProdutoRoupas.Text & "') AND (produtos.codigo <> 1) " & varTipoMostrar & " " & var_Criterio
      
'      Set r = dbData.OpenRecordset(sSQL)
'      FormatarGrid_Produtos r
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
       
'   ElseIf chkCodBarra.Value = Checked Then
'       sSQL = "SELECT produtos.ref AS var_Ref, produtos.fabricante AS var_Fab, produtos.tamanho AS var_Tam, produtos.categoria AS var_linha, produtos.codigo AS var_codEnt, produtos.cod_barra AS var_CodBarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Prat, produtos.unid_medida AS var_Med, " & _
'         "produtos.quant_estoque AS var_Quant, (SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) AS venda " & _
'         "FROM produtos  " & _
'         "WHERE (produtos.cod_barra = '" & txtConsCodBarra.Text & "') AND (produtos.codigo <> 1) " & var_Criterio
      
'      Set r = dbData.OpenRecordset(sSQL)
'      FormatarGrid_Produtos r
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
      
'   ElseIf chkFab.Value = Checked Then
'       sSQL = "SELECT produtos.ref AS var_Ref, produtos.fabricante AS var_Fab, produtos.tamanho AS var_Tam, produtos.categoria AS var_linha, produtos.codigo AS var_codEnt, produtos.cod_barra AS var_CodBarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Prat, produtos.unid_medida AS var_Med, " & _
'         "produtos.quant_estoque AS var_Quant, (SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) AS venda " & _
'         "FROM produtos  " & _
'         "WHERE (produtos.fabricante = '" & cboConsFab.Text & "') AND (produtos.codigo <> 1) " & var_Criterio
      
'      Set r = dbData.OpenRecordset(sSQL)
'      FormatarGrid_Produtos r
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
      
'   ElseIf chkRef.Value = Checked And chkTam.Value = Unchecked Then
'       sSQL = "SELECT produtos.ref AS var_Ref, produtos.fabricante AS var_Fab, produtos.tamanho AS var_Tam, produtos.categoria AS var_linha, produtos.codigo AS var_codEnt, produtos.cod_barra AS var_CodBarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Prat, produtos.unid_medida AS var_Med, " & _
'         "produtos.quant_estoque AS var_Quant, (SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) AS venda " & _
'         "FROM produtos  " & _
'         "WHERE (produtos.ref = '" & cboConsRef.Text & "') AND (produtos.codigo <> 1) " & var_Criterio
      
'      Set r = dbData.OpenRecordset(sSQL)
'      FormatarGrid_Produtos r
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
      
'   ElseIf chkTam.Value = Checked And chkRef.Value = Unchecked Then
'       sSQL = "SELECT produtos.ref AS var_Ref, produtos.fabricante AS var_Fab, produtos.tamanho AS var_Tam, produtos.categoria AS var_linha, produtos.codigo AS var_codEnt, produtos.cod_barra AS var_CodBarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Prat, produtos.unid_medida AS var_Med, " & _
'         "produtos.quant_estoque AS var_Quant, (SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) AS venda " & _
'         "FROM produtos  " & _
'         "WHERE (produtos.tamanho = '" & cboConsTam.Text & "') AND (produtos.codigo <> 1) " & var_Criterio
      
'      Set r = dbData.OpenRecordset(sSQL)
'      FormatarGrid_Produtos r
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
   
'   ElseIf chkTam.Value = Checked And chkRef.Value = Checked Then
'       sSQL = "SELECT produtos.ref AS var_Ref, produtos.fabricante AS var_Fab, produtos.tamanho AS var_Tam, produtos.categoria AS var_linha, produtos.codigo AS var_codEnt, produtos.cod_barra AS var_CodBarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Prat, produtos.unid_medida AS var_Med, " & _
'         "produtos.quant_estoque AS var_Quant, (SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) AS venda " & _
'         "FROM produtos  " & _
'         "WHERE (produtos.ref = '" & cboConsRef.Text & "') and (produtos.tamanho = '" & cboConsTam.Text & "') AND (produtos.codigo <> 1) " & var_Criterio
      
'      Set r = dbData.OpenRecordset(sSQL)
'      FormatarGrid_Produtos r
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
      
'   ElseIf chkLinha.Value = Checked Then
'       sSQL = "SELECT produtos.ref AS var_Ref, produtos.fabricante AS var_Fab, produtos.tamanho AS var_Tam, produtos.categoria AS var_linha, produtos.codigo AS var_codEnt, produtos.cod_barra AS var_CodBarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Prat, produtos.unid_medida AS var_Med, " & _
'         "produtos.quant_estoque AS var_Quant, (SELECT TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
'         "LEFT JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
'         "WHERE produtos_entrada_itens.codigo_produto = produtos.codigo ORDER BY " & _
'         "produtos_entrada_itens.CODIGO DESC) AS venda " & _
'         "FROM produtos  " & _
'         "WHERE (produtos.categoria = '" & cboConsLinha.Text & "') AND (produtos.codigo <> 1) " & var_Criterio
      
'      Set r = dbData.OpenRecordset(sSQL)
'      FormatarGrid_Produtos r
'      If r.State <> 0 Then r.Close
'      Set r = Nothing
'
'   End If
   
'   If chkTodos.Value = False Then
'      cboConsProduto.SelStart = 0
'      cboConsProduto.SelLength = Len(cboConsProduto)
'   End If

'Else  '..............outros tipos de empresa

   'Indice
   Dim INDICE As String
   If cboOrdem.Text = "CÓD. BARRA" Then
      INDICE = "produtos.cod_barra "
   ElseIf cboOrdem.Text = "DESCRIÇÃO" Then
      INDICE = "produtos.descricao "
   ElseIf cboOrdem.Text = "VENDA" Then
      INDICE = "venda "
   ElseIf cboOrdem.Text = "CUSTO" Then
      INDICE = "Custo "
   ElseIf cboOrdem.Text = "FABRICANTE" Then
      INDICE = "produtos.fabricante "
   Else
      INDICE = "produtos.descricao "
   End If

   'Indice2
   Dim INDICE2 As String
   If cboOrdem2.Text = "ASC" Then
      INDICE2 = "ASC;"
   ElseIf cboOrdem2.Text = "DESC" Then
      INDICE2 = "DESC;"
   Else
      INDICE2 = "ASC;"
   End If

    Dim vUltimoValorVenda As String     '===================TER QUE COLOCAR DEPOIS PARA TODOS OS TIPOS DE VENDAS
    
    If optComPreco.Value = True Then
        vUltimoValorVenda = " (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) > 0 "
    ElseIf optSemPreco.Value = True Then
        vUltimoValorVenda = " (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) = 0"
    ElseIf optTodosPreco.Value = True Then
        vUltimoValorVenda = " 1=1 "
    End If
   
   'Monta a consulta básica para não repetir várias linhas
   If vMultiplasRef = "SIM" Then
    sSQL = "SELECT DISTINCT produtos.codigo AS varCodProd, produtos.ref AS var_Ref, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Local, " & _
      "produtos.fabricante AS var_fab, produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant, " & _
      "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS Custo, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, (CASE WHEN ATIVO = 1 THEN 'ATIVO' ELSE 'DESATIVO' END) as vAtivo " & _
      "FROM  produtos  LEFT OUTER JOIN Produtos_Referencias ON produtos.CODIGO = Produtos_Referencias.COD_PRODUTO " & _
      "WHERE " & varProdutoHabilitado & " " & varTipoMostrar & " " & vUltimoValorVenda & "  and "
    Else
        sSQL = "SELECT DISTINCT produtos.codigo AS varCodProd, produtos.ref AS var_Ref, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.PRATELEIRA AS var_Local, " & _
      "produtos.fabricante AS var_fab, produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant, " & _
      "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS Custo, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, (CASE WHEN ATIVO = 1 THEN 'ATIVO' ELSE 'DESATIVO' END) as vAtivo " & _
      "FROM produtos " & _
      "WHERE " & varProdutoHabilitado & " " & varTipoMostrar & " " & vUltimoValorVenda & "  and "
    End If

   If cboCriterios.Text = "CÓD. BARRA" Then
        If Len(cboConsProduto) < 13 And cboConsProduto.Text <> "" Then
            If Len(cboConsProduto) < 6 Then
                cboConsProduto.Text = Format(cboConsProduto.Text, "00000")
            Else
                cboConsProduto.Text = cboConsProduto.Text
            End If
        End If
      sSQL = sSQL & "(produtos.cod_barra = '" & cboConsProduto.Text & "') AND (produtos.codigo <> 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf cboCriterios.Text = "TODOS" Then
      'sSQL = sSQL & "(produtos.codigo <> 1) ORDER BY produtos.codigo"
      sSQL = sSQL & " (produtos.codigo <> 1) ORDER BY " & INDICE & " " & INDICE2 & " "
      'Debug.Print sSQL
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf cboCriterios.Text = "CATEGORIA" Then
      sSQL = sSQL & "(produtos.categoria = '" & cboConsProduto.Text & "') AND (produtos.codigo <> 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing

   ElseIf cboCriterios.Text = "REFERÊNCIA" Then
    If vMultiplasRef = "NÃO" Then
      sSQL = sSQL & "(produtos.ref = '" & cboConsProduto.Text & "') AND (produtos.codigo <> 1) ORDER BY " & INDICE
    Else
      sSQL = sSQL & "(Produtos_Referencias.REFERENCIA = '" & cboConsProduto.Text & "') AND (produtos.codigo <> 1) ORDER BY " & INDICE
    End If
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf cboCriterios.Text = "DESCRIÇÃO" Then
      sSQL = sSQL & "(produtos.codigo <> 1) and " & vDescricao & " ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing

   ElseIf cboCriterios.Text = "FABRICANTE" Then
      sSQL = sSQL & "(produtos.fabricante = '" & cboConsProduto.Text & "') AND (produtos.codigo <> 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)

      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
   ElseIf cboCriterios.Text = "NCM" Then
      sSQL = sSQL & "(produtos.NCM = '" & cboConsProduto.Text & "') AND (produtos.codigo <> 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)

      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
   
   If cboCriterios.Text = "TODOS" Then
      SelectControl cboConsProduto
   End If
'End If

'Debug.Print sSQL
printSQL = sSQL
End Sub



Private Sub Form_Load()
SSTab1.Tab = 0
LimparGrid_Produtos
LimparObjeto_Gas
Mostrar_HistoricoQuant
Mostrar_Precos

Dim vUsarOS As Boolean
vTipoEdicao = ""

Set oCfg = sysConfig("OS")    'Recupera a config deseja
bStatus = CBool(oCfg.Value)   'Converte o valor para booleano
Set oCfg = Nothing            'Destroi o objeto

If bStatus = True Then
    Set oCfg = sysConfig("TIPO_OS")
    vTipoOS = oCfg.Value
    Set oCfg = Nothing
    
    Dim vTabelaServico As String
    If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
        frmComp.Visible = True
        frmReferencia.Visible = True
    Else
        frmComp.Visible = False
        frmReferencia.Visible = False
    End If
Else
    frmComp.Visible = False
    frmReferencia.Visible = False
End If

Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

'If tipoEmpresa = 4 Then
      'frmCriterioRoupas.Visible = True
      'frmOrdemRoupas.Visible = True
      'frmFiltroRoupas.Visible = True
      'Label3.Caption = "Linha"
      'frmComp.Visible = True
'ElseIf tipoEmpresa = 5 Then
'      frmCriterios.Visible = True
'      frmOrdemComum.Visible = True
'      frmFiltroComum.Visible = True
'      Label3.Caption = "Categoria"
'      'frmComp.Visible = True
'Else
      frmCriterios.Visible = True
      frmOrdemComum.Visible = True
      frmFiltroComum.Visible = True
      Label3.Caption = "Categoria"
'      'frmComp.Visible = False
'End If

'If Tela_Principal.txtNivel.Text <> "1" Then chkAtivo.Enabled = False: Exit Sub

'If Tela_Principal.txtNivel.Text <> "1" Then
'   frmEstoque.Visible = False
'   frmCompra.Visible = False
'   frmVenda.Visible = False
'Else
'   frmEstoque.Visible = True
'   frmCompra.Visible = True
'   frmVenda.Visible = True
'End If

Set cCfg = sysConfig("MULTIPLASREF")
vMultiplasRef = cCfg.Value
Set cCfg = Nothing

If vMultiplasRef = "SIM" Then
    frmReferencia.Visible = True
    lblRef.Visible = False
    txtRef.Visible = False
Else
    frmReferencia.Visible = False
    lblRef.Visible = True
    txtRef.Visible = True
End If

optCompleto.Visible = False
optPorIniciais.Visible = False
optPorPalavra.Visible = False
optPalavrasDuplas.Visible = False

    ' Preenche as unidades apenas no início
    With cboUnidMedida
        .Clear
        .AddItem "UN": .AddItem "CX": .AddItem "M": .AddItem "M2"
        .AddItem "M3": .AddItem "ML": .AddItem "KG": .AddItem "GR"
        .AddItem "CT": .AddItem "PO": .AddItem "SC": .AddItem "PA"
        .AddItem "EX": .AddItem "BJ": .AddItem "DZ": .AddItem "PC"
        .AddItem "DI": .AddItem "FD": .AddItem "PT"
    End With

DesabilitarFrames
DesabilitarBotoes
cboCriterios.Text = "TODOS"
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")

'ver o regime da empresa
sSQL = "SELECT CRT, pAliqUF FROM empresa"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    var_RegimeEmpresa = IIf(IsNull(r!CRT), 1, r!CRT)
    var_AliqInterna = IIf(IsNull(r!pAliqUF), 0, r!pAliqUF)
Else
    var_RegimeEmpresa = 1 ' Valor padrão caso a tabela esteja vazia
    var_AliqInterna = FormatNumber(0, 2)
End If

r.Close
Set r = Nothing

'AGORA CHAMA O PREENCHIMENTO DO COMBO
PreencherCFOP

'muda o label
If var_RegimeEmpresa = 1 Then
    Label31.Caption = "CSOSN"
ElseIf var_RegimeEmpresa = 3 Then
    Label31.Caption = "ICMS CST"
End If
Set moCombo = New cComboHelper
End Sub
Private Sub PreencherCFOP()
cboCFOP.Clear

If var_RegimeEmpresa = 1 Then ' SIMPLES NACIONAL
    cboCFOP.AddItem "5102"
    cboCFOP.AddItem "5405"
ElseIf var_RegimeEmpresa = 3 Then ' LUCRO PRESUMIDO
    cboCFOP.AddItem "5102"
    cboCFOP.AddItem "5405"
    cboCFOP.AddItem "5403"
    cboCFOP.AddItem "5401"
    cboCFOP.AddItem "5101"
    cboCFOP.AddItem "5949"
End If

' Opcional: Seleciona o primeiro item da lista automaticamente
If cboCFOP.ListCount > 0 Then cboCFOP.ListIndex = 0
End Sub
Private Sub SomaQuantRemocao()
   On Error GoTo errorhandeler
   Dim soma As Currency
   Dim i As Integer
   
   soma = 0
   With Grid_Quant
      For i = 1 To .rows - 1
         If .TextMatrix(i, 6) = "REMOÇÃO" And IsNumeric(.TextMatrix(i, 8)) Then
            soma = soma + CCur(.TextMatrix(i, 8))
         End If
      Next
   End With
   
   lblQuantRemocao.Caption = soma
   
errorhandeler:
End Sub
Private Sub SomaQuantAdicao()
   On Error GoTo errorhandeler
   Dim soma As Currency
   Dim i As Integer
   
   soma = 0
   With Grid_Quant
      For i = 1 To .rows - 1
         If .TextMatrix(i, 6) = "ADIÇÃO" And IsNumeric(.TextMatrix(i, 8)) Then
            soma = soma + CCur(.TextMatrix(i, 8))
         End If
      Next
   End With
   
   lblQuantAdicao.Caption = soma
   
errorhandeler:
End Sub

Private Sub Mostrar_Precos()
Dim sSQL As String
Dim r As ADODB.Recordset

'Monta a consulta básica
sSQL = "SELECT * " & _
   "FROM produtos_precos "

'Define o filtro
If txtCodigo.Text = "" Then
   sSQL = sSQL & "WHERE 1 = 0 "
   
Else
   sSQL = sSQL & "WHERE (cod_produto = " & txtCodigo.Text & ") "
End If

'Monta a ordem de exibição
sSQL = sSQL & "ORDER BY codigo "

Set r = dbData.OpenRecordset(sSQL)
FormatarGrid_Precos r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_HistoricoQuant()
Dim sSQL As String
Dim r As ADODB.Recordset

'Monta a consulta básica
sSQL = "SELECT Produtos_Quant.codigo, Produtos_Quant.Hora, Produtos_Quant.DATA, Produtos_Quant.QUANT, Produtos_Quant.cod_usuario, Produtos_Quant.Estoque, Produtos_Quant.TIPO, Produtos_Quant.FORMA, Produtos_Quant.COD_ENTRADA, Produtos_Quant.COD_PRODUTO, produtos_entrada.NOTAFISCAL " & _
       "FROM Produtos_Quant LEFT JOIN produtos_entrada ON Produtos_Quant.cod_entrada = produtos_entrada.codigo "

'Define o filtro
If txtCodigo.Text = "" Then
   sSQL = sSQL & "WHERE 1 = 0 "
Else
   sSQL = sSQL & "WHERE (cod_produto = " & txtCodigo.Text & ") "
End If

'Monta a ordem de exibição
sSQL = sSQL & "ORDER BY data, hora "

Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Quant r
SomaQuantAdicao
SomaQuantRemocao

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

Private Sub Grid_Click()
i = Grid.Row
If Grid.TextMatrix(i, 12) = "ATIVO" Then
    cmdDesativar.Caption = "Desativar"
Else
    cmdDesativar.Caption = "Ativar"
End If

End Sub

Private Sub Grid_DblClick()
'SSTab1.Tab = 0
'cmdNovo.Enabled = False
'cmdSalvar.Enabled = False
'cmdCancelar.Enabled = False
'cmdAlterar.Enabled = True
'cmdExcluir.Enabled = True
'txtCodigo.Text = ""
''vModoEdicao = True
'txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
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


Private Sub optDesabilitados_Click()
cmdExibir_Click
If optHabilitado.Value = True Then
    cmdDesativar.Caption = "Desativar"
ElseIf optDesabilitados.Value = True Then
    cmdDesativar.Caption = "Ativar"
Else
    cmdDesativar.Caption = "Desativar"
End If
End Sub

Private Sub optHabilitado_Click()
cmdExibir_Click
If optHabilitado.Value = True Then
    cmdDesativar.Caption = "Desativar"
ElseIf optDesabilitados.Value = True Then
    cmdDesativar.Caption = "Ativar"
Else
    cmdDesativar.Caption = "Desativar"
End If
End Sub

Private Sub txtCEST_GotFocus()
txtCEST.SelStart = 0
txtCEST.SelLength = Len(txtCEST)
End Sub

Private Sub txtCEST_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtCEST_LostFocus()
txtCEST = Replace(txtCEST, ".", "")
txtCEST = Trim(txtCEST.Text)
End Sub








Private Sub txtCodBarra_GotFocus()
SelectControl txtCodBarra
End Sub


Private Sub txtCodBarra_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtCodBarra_LostFocus()
If txtCodBarra.Text = "" Then
    Dim sSQL As String
    Dim r As ADODB.Recordset
    sSQL = "SELECT isnull(MAX(COD_BARRA), 0) as UltimoCodigo FROM produtos where len(COD_BARRA) = 5;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Val(Len(txtCodBarra)) < 13 Then
        Dim vCodBarraInt As String
        vCodBarraInt = Val(r("UltimoCodigo"))
    End If
    
    If Not r.BOF Then
        txtCodBarra.Text = Format(vCodBarraInt + 1, "00000")
    End If
Else
    If Len(txtCodBarra) < 13 And txtCodBarra.Text <> "" Then
        If Len(txtCodBarra) < 6 Then
            txtCodBarra.Text = Format(txtCodBarra.Text, "00000")
        Else
            txtCodBarra.Text = txtCodBarra.Text
        End If
    ElseIf Len(txtCodBarra) > 13 Then
        MsgBox "Esse Cód. de Barra possui mais números que o permitido", vbInformation, "Aviso do Sistema"
        txtCodBarra.SetFocus
        Exit Sub
    End If
End If
End Sub


Private Sub txtCodBarra_Validate(Cancel As Boolean)
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodBarra.Text = "" Then Exit Sub
txtCodBarra.Text = Trim(txtCodBarra.Text)

'Verifica se existe o código de barras cadastrado
sSQL = "SELECT codigo, ativo, cod_barra FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "');"
Set r = dbData.OpenRecordset(sSQL)

If vTipoEdicao = "Novo" Or vTipoEdicao = "Duplicar" Then
   If r.RecordCount > 0 Then
        If r("ativo") = True Then
            ShowMsg "Já existe um produto cadastrado com esse cód. de barra!", vbInformation
            Cancel = True           'Cancela a entrada e permanece com o foco no campo
            txtCodBarra.Text = ""   'Limpa a entrada
            txtCodBarra.SetFocus
            Exit Sub                'Evita a saída do campo
        ElseIf r("ativo") = False Then
            ShowMsg "Existe um produto DESABILITADO com esse cód. de barra!", vbInformation
            Cancel = True           'Cancela a entrada e permanece com o foco no campo
            txtCodBarra.Text = ""   'Limpa a entrada
            txtCodBarra.SetFocus
            Exit Sub
        End If
   End If
End If
End Sub

Private Sub txtCODIF_GotFocus()
SelectControl txtCODIF
End Sub

Private Sub txtCODIF_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtCodigo_Change()
If vTipoEdicao = "Edicao" Then
   If txtCodigo.Text = "" Then Exit Sub
   LimparObjetos_Produtos
   cmdSalvar.Enabled = True
   cmdCancelar.Enabled = True
   'cmdAlterar.Enabled = True
   'cmdExcluir.Enabled = True
   cmdNovo.Enabled = False
   HabilitarFrames
   MostrarDados_Produto
   MostrarObjetosPrecos
   Mostrar_HistoricoQuant
   Mostrar_Precos
   MostrarGrid_Referencia
   If vTipoEdicao = "Edicao" Then frmPrecos.Enabled = False
   If vTipoEdicao = "Edicao" Then txtQuant.Enabled = False: lblQuantAtual.Enabled = False
   If frmComp.Visible = True Then Mostrar_Comp
End If
End Sub

Private Sub txtCodProdFracionado_Change()
If txtCodProdFracionado.Text = "" Or txtCodProdFracionado.Text = "0" Then Exit Sub

If cboProdutoFracionado.Text = "" Then
   sSQL = "SELECT * FROM produtos WHERE (codigo= " & txtCodProdFracionado.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then cboProdutoFracionado.Text = r("DESCRICAO")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub

Private Sub txtCofinsAliquota_GotFocus()
txtCofinsAliquota.SelStart = 0
txtCofinsAliquota.SelLength = Len(txtCofinsAliquota.Text)
End Sub


Private Sub txtCofinsAliquota_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtCofinsAliquota_LostFocus()
If txtCofinsAliquota.Text = "" Then
    txtCofinsAliquota.Text = FormatNumber(0, 2)
Else
    txtCofinsAliquota.Text = FormatNumber(txtCofinsAliquota.Text, 2)
End If
End Sub

Private Sub txtCOFINSCST_GotFocus()
txtCOFINSCST.SelStart = 0
txtCOFINSCST.SelLength = Len(txtCOFINSCST)
End Sub

Private Sub txtCOFINSCST_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub


Private Sub txtcProdANP_GotFocus()
SelectControl txtcProdANP
End Sub


Private Sub txtcProdANP_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub











Private Sub txtCusto_GotFocus()
txtCusto.SelStart = 0
txtCusto.SelLength = Len(txtCusto)
End Sub


Private Sub txtCusto_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtCusto_LostFocus()
Dim varLucro As Currency

If txtCusto.Text = "" Then Exit Sub
varLucro = txtCusto.Text

txtCusto.Text = FormatNumber(varLucro, 2)

CalcularPrecos
End Sub


Private Sub txtDescricao_Change()
lblNomeProduto1.Caption = txtDescricao.Text
lblNomeProduto2.Caption = txtDescricao.Text
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



Private Sub txtDescricao_LostFocus()
txtDescricao.Text = TirarEspaco(txtDescricao.Text)
txtDescricao.Text = UCase(txtDescricao.Text)
End Sub

Private Sub txtdescricaoANP_GotFocus()
SelectControl txtdescricaoANP
End Sub

Private Sub txtdescricaoANP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtdescricaoANP_LostFocus()
txtdescricaoANP.Text = TirarEspaco(txtdescricaoANP.Text)
End Sub


Private Sub txtEAN_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtEAN_LostFocus()
If txtEAN.Text = "" Then txtEAN.Text = "SEM GTIN": Exit Sub
If txtEAN.Text <> "SEM GTIN" Then
    If Len(txtEAN.Text) < 6 Then MsgBox "O EAN não pode ser um codigo criado", vbInformation, "Aviso do Sistema": txtEAN.Text = "": Exit Sub
    If Len(txtEAN) > 13 Then
        MsgBox "Esse EAN possui mais números que o permitido", vbInformation, "Aviso do Sistema"
        txtCodBarra.SetFocus
        Exit Sub
    End If
End If
End Sub


Private Sub txtICMSAliquota_GotFocus()
txtICMSAliquota.SelStart = 0
txtICMSAliquota.SelLength = Len(txtICMSAliquota.Text)
End Sub

Private Sub txtICMSAliquota_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtICMSAliquota_LostFocus()
If txtICMSAliquota.Text = "" Then
    txtICMSAliquota.Text = FormatNumber(0, 2)
Else
    txtICMSAliquota.Text = FormatNumber(txtICMSAliquota.Text, 2)
End If
End Sub

Private Sub txtIPIAliquota_GotFocus()
txtIPIAliquota.SelStart = 0
txtIPIAliquota.SelLength = Len(txtIPIAliquota.Text)
End Sub


Private Sub txtIPIAliquota_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtIPIAliquota_LostFocus()
If txtIPIAliquota.Text = "" Then
    txtIPIAliquota.Text = FormatNumber(0, 2)
Else
    txtIPIAliquota.Text = FormatNumber(txtIPIAliquota.Text, 2)
End If
End Sub

Private Sub txtIPICST_GotFocus()
txtIPICST.SelStart = 0
txtIPICST.SelLength = Len(txtIPICST)
End Sub

Private Sub txtIPICST_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtMargemAP_GotFocus()
If txtMargemAP.Text = "" Then Exit Sub
Dim varMargemAP As Currency

If Right(txtMargemAP.Text, 1) = "%" Then
   varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)
Else
    varMargemAP = txtMargemAP.Text
End If

txtMargemAP.Text = varMargemAP

txtMargemAP.SelStart = 0
txtMargemAP.SelLength = Len(txtMargemAP.Text)
lblAviso.Visible = True
End Sub


Private Sub txtMargemAP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If txtCusto.Text = "" Then Exit Sub
    varValorEstimado = Empty
    varCustoEstimado = CCur(txtCusto)
    Produtos_ValorEstimado.Show vbModal
    Unload Produtos_ValorEstimado
    txtMargemAP.Text = varValorEstimado
End If
End Sub


Private Sub txtMargemAP_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtMargemAP_LostFocus()
Dim varMargemAP As Currency

If txtMargemAP.Text = "" Then txtMargemAP.Text = 0
varMargemAP = txtMargemAP.Text

txtMargemAP.Text = FormatNumber(varMargemAP, 3) & "%"

CalcularPrecos
lblAviso.Visible = False
End Sub


Private Sub txtMargemAV_GotFocus()
If txtMargemAV.Text = "" Then Exit Sub
Dim varMargemAV As Currency

If Right(txtMargemAV.Text, 1) = "%" Then
   varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)
Else
    varMargemAV = txtMargemAV.Text
End If

txtMargemAV.Text = varMargemAV

txtMargemAV.SelStart = 0
txtMargemAV.SelLength = Len(txtMargemAV.Text)
lblAviso.Visible = True
End Sub


Private Sub txtMargemAV_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If txtCusto.Text = "" Then Exit Sub
    varValorEstimado = Empty
    varCustoEstimado = CCur(txtCusto)
    Produtos_ValorEstimado.Show vbModal
    Unload Produtos_ValorEstimado
    txtMargemAV.Text = varValorEstimado
End If
End Sub


Private Sub txtMargemAV_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtMargemAV_LostFocus()
Dim varMargemAV As Currency

If txtMargemAV.Text = "" Then txtMargemAV.Text = 0
varMargemAV = txtMargemAV.Text

txtMargemAV.Text = FormatNumber(varMargemAV, 3) & "%"

CalcularPrecos
lblAviso.Visible = False
End Sub


Private Sub txtMargemVP_GotFocus()
If txtMargemVP.Text = "" Then Exit Sub
Dim varMargemVP As Currency

If Right(txtMargemVP.Text, 1) = "%" Then
   varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)
Else
    varMargemVP = txtMargemVP.Text
End If

txtMargemVP.Text = varMargemVP

txtMargemVP.SelStart = 0
txtMargemVP.SelLength = Len(txtMargemVP.Text)
lblAviso.Visible = True
End Sub


Private Sub txtMargemVP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If txtCusto.Text = "" Then Exit Sub
    varValorEstimado = Empty
    varCustoEstimado = CCur(txtCusto)
    Produtos_ValorEstimado.Show vbModal
    Unload Produtos_ValorEstimado
    txtMargemVP.Text = varValorEstimado
End If
End Sub


Private Sub txtMargemVP_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtMargemVP_LostFocus()
Dim varMargemVP As Currency

If txtMargemVP.Text = "" Then txtMargemVP.Text = 0
varMargemVP = txtMargemVP.Text

txtMargemVP.Text = FormatNumber(varMargemVP, 3) & "%"

CalcularPrecos
lblAviso.Visible = False
End Sub


Private Sub txtMargemVV_GotFocus()
If txtMargemVV.Text = "" Then Exit Sub
Dim varMargemVV As Currency

If Right(txtMargemVV.Text, 1) = "%" Then
   varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)
Else
    varMargemVV = txtMargemVV.Text
End If

txtMargemVV.Text = varMargemVV

txtMargemVV.SelStart = 0
txtMargemVV.SelLength = Len(txtMargemVV.Text)
lblAviso.Visible = True
End Sub

Private Sub txtMargemVV_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If txtCusto.Text = "" Then Exit Sub
    varValorEstimado = Empty
    varCustoEstimado = CCur(txtCusto)
    Produtos_ValorEstimado.Show vbModal
    Unload Produtos_ValorEstimado
    txtMargemVV.Text = varValorEstimado
End If
End Sub


Private Sub txtMargemVV_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtMargemVV_LostFocus()
Dim varMargemVV As Currency

If txtMargemVV.Text = "" Then txtMargemVV.Text = 0
varMargemVV = txtMargemVV.Text

txtMargemVV.Text = FormatNumber(varMargemVV, 3) & "%"
If txtMargemVP.Text = "" Then txtMargemVP.Text = txtMargemVV.Text
If txtMargemAV.Text = "" Then txtMargemAV.Text = txtMargemVV.Text
If txtMargemAP.Text = "" Then txtMargemAP.Text = txtMargemVV.Text
CalcularPrecos
lblAviso.Visible = False
End Sub


Private Sub txtNCM_GotFocus()
txtNCM.SelStart = 0
txtNCM.SelLength = Len(txtNCM)
End Sub

Private Sub txtNCM_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtNCM_LostFocus()
txtNCM = Replace(txtNCM, ".", "")
txtNCM = Trim(txtNCM.Text)

If txtNCM.Text <> "" Then
    If Len(txtNCM.Text) < 8 Or Len(txtNCM.Text) > 8 Then
        MsgBox "NCM Inválido!", vbInformation, "Aviso do Sistema"
        'txtNCM.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub txtOBS_LostFocus()
   If cmdSalvar.Enabled = True And cmdCancelar.Enabled = True Then
      'cmdSalvar.SetFocus
   ElseIf cmdAlterar.Enabled = True Then
      'cmdAlterar.SetFocus
   Else
      Exit Sub
   End If
End Sub

Private Sub txtpGLP_GotFocus()
If txtpGLP.Text = "" Then Exit Sub
Dim vpGLP As Currency

If Right(txtpGLP.Text, 1) = "%" Then
   vpGLP = Left$(txtpGLP.Text, Len(txtpGLP.Text) - 1)
Else
    vpGLP = txtpGLP.Text
End If

txtpGLP.Text = vpGLP

SelectControl txtpGLP
lblAviso.Visible = True
End Sub


Private Sub txtpGLP_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtpGLP_LostFocus()
Dim vpGLP As Currency

If txtpGLP.Text = "" Then txtpGLP.Text = 0
vpGLP = txtpGLP.Text

txtpGLP.Text = FormatNumber(vpGLP, 2) & "%"
End Sub

Private Sub txtpGNi_GotFocus()
If txtpGNi.Text = "" Then Exit Sub
Dim vpGNi As Currency

If Right(txtpGNi.Text, 1) = "%" Then
   vpGNi = Left$(txtpGNi.Text, Len(txtpGNi.Text) - 1)
Else
    vpGNi = txtpGNi.Text
End If

txtpGNi.Text = vpGNi

SelectControl txtpGNi
End Sub


Private Sub txtpGNi_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtpGNi_LostFocus()
Dim vpGNi As Currency

If txtpGNi.Text = "" Then txtpGNi.Text = 0
vpGNi = txtpGNi.Text

txtpGNi.Text = FormatNumber(vpGNi, 2) & "%"
End Sub

Private Sub txtpGNn_GotFocus()
If txtpGNn.Text = "" Then Exit Sub
Dim vpGNn As Currency

If Right(txtpGNn.Text, 1) = "%" Then
   vpGNn = Left$(txtpGNn.Text, Len(txtpGNn.Text) - 1)
Else
    vpGNn = txtpGNn.Text
End If

txtpGNn.Text = vpGNn

SelectControl txtpGNn
End Sub


Private Sub txtpGNn_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtpGNn_LostFocus()
Dim vpGNn As Currency

If txtpGNn.Text = "" Then txtpGNn.Text = 0
vpGNn = txtpGNn.Text

txtpGNn.Text = FormatNumber(vpGNn, 2) & "%"
End Sub

Private Sub txtPisAliquota_GotFocus()
txtPisAliquota.SelStart = 0
txtPisAliquota.SelLength = Len(txtPisAliquota.Text)
End Sub


Private Sub txtPisAliquota_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtPisAliquota_LostFocus()
If txtPisAliquota.Text = "" Then
    txtPisAliquota.Text = FormatNumber(0, 2)
Else
    txtPisAliquota.Text = FormatNumber(txtPisAliquota.Text, 2)
End If
End Sub

Private Sub txtPISCST_GotFocus()
txtPISCST.SelStart = 0
txtPISCST.SelLength = Len(txtPISCST)
End Sub

Private Sub txtPISCST_KeyPress(KeyAscii As Integer)
On Error GoTo erro
    If KeyAscii = 8 Then
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtpMixGN_GotFocus()
If txtpMixGN.Text = "" Then Exit Sub
Dim vpMixGN As Currency

If Right(txtpMixGN.Text, 1) = "%" Then
   vpMixGN = Left$(txtpMixGN.Text, Len(txtpMixGN.Text) - 1)
Else
    vpMixGN = txtpMixGN.Text
End If

txtpMixGN.Text = vpMixGN

SelectControl txtpMixGN
End Sub


Private Sub txtpMixGN_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtpMixGN_LostFocus()
Dim varMargemVP As Currency

If txtpMixGN.Text = "" Then txtpMixGN.Text = 0
varMargemVP = txtpMixGN.Text

txtpMixGN.Text = FormatNumber(varMargemVP, 2) & "%"
End Sub

Private Sub txtPrateleira_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtQuant_GotFocus()
   SelectControl txtQuant
End Sub

Private Sub txtQuantFracionado_GotFocus()
SelectControl txtQuantFracionado
End Sub


Private Sub txtQuantFracionado_LostFocus()
txtQuantFracionado.Text = Format(txtQuantFracionado.Text, ocPESO)
End Sub


Private Sub txtRedBCAliquota_GotFocus()
txtRedBCAliquota.SelStart = 0
txtRedBCAliquota.SelLength = Len(txtRedBCAliquota.Text)
End Sub


Private Sub txtRedBCAliquota_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtRedBCAliquota_LostFocus()
' Formata com duas casas decimais ao sair
If IsNumeric(txtRedBCAliquota.Text) Then
    txtRedBCAliquota.Text = Format(txtRedBCAliquota.Text, "##0.00")
Else
    txtRedBCAliquota.Text = "0,00"
End If
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtValorAP_Click()
SelectControl txtValorAP
End Sub

Private Sub txtValorAP_GotFocus()
SelectControl txtValorAP
End Sub


Private Sub txtValorAP_LostFocus()
If txtCusto.Text = "" Then Exit Sub
If txtValorAP.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency
Dim c As Currency

a = txtCusto.Text
B = txtValorAP.Text
If a = "0,00" Or B = "0,00" Then Exit Sub
c = ((B - a) / a) * 100

txtMargemAP.Text = FormatNumber(c, 3) & "%"
txtValorAP.Text = Format(txtValorAP.Text, ocMONEY)
End Sub


Private Sub txtValorAV_Click()
SelectControl txtValorAV
End Sub

Private Sub txtValorAV_GotFocus()
SelectControl txtValorAV
End Sub


Private Sub txtValorAV_LostFocus()
If txtCusto.Text = "" Then Exit Sub
If txtValorAV.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency
Dim c As Currency

a = txtCusto.Text
B = txtValorAV.Text
If a = "0,00" Or B = "0,00" Then Exit Sub
c = ((B - a) / a) * 100

txtMargemAV.Text = FormatNumber(c, 3) & "%"
txtValorAV.Text = Format(txtValorAV.Text, ocMONEY)
End Sub


Private Sub txtValorPartida_GotFocus()
SelectControl txtValorPartida
End Sub


Private Sub txtValorPartida_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtValorPartida_LostFocus()
txtValorPartida.Text = Format(txtValorPartida.Text, ocMONEY)
End Sub

Private Sub txtValorVP_Click()
SelectControl txtValorVP
End Sub

Private Sub txtValorVP_GotFocus()
SelectControl txtValorVP
End Sub


Private Sub txtValorVP_LostFocus()
If txtCusto.Text = "" Then Exit Sub
If txtValorVP.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency
Dim c As Currency

a = txtCusto.Text
B = txtValorVP.Text
If a = "0,00" Or B = "0,00" Then Exit Sub
c = ((B - a) / a) * 100

txtMargemVP.Text = FormatNumber(c, 3) & "%"
txtValorVP.Text = Format(txtValorVP.Text, ocMONEY)
End Sub


Private Sub txtValorVV_Click()
SelectControl txtValorVV
End Sub

Private Sub txtValorVV_GotFocus()
SelectControl txtValorVV
End Sub


Private Sub txtValorVV_LostFocus()
If txtCusto.Text = "" Then Exit Sub
If txtValorVV.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency
Dim c As Currency

a = txtCusto.Text
B = txtValorVV.Text
If a = "0,00" Or B = "0,00" Then Exit Sub
c = ((B - a) / a) * 100

txtMargemVV.Text = FormatNumber(c, 3) & "%"
txtValorVV.Text = Format(txtValorVV.Text, ocMONEY)
End Sub


