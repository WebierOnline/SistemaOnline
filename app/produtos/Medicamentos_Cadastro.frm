VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Medicamentos_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRODUTOS"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12375
   Icon            =   "Medicamentos_Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   12225
      TabIndex        =   18
      Top             =   60
      Width           =   12255
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Medicamentos_Cadastro.frx":23D2
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
         TabIndex        =   19
         Top             =   240
         Width           =   1770
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8715
      Left            =   60
      TabIndex        =   13
      Top             =   1080
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   15372
      _Version        =   393216
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
      TabPicture(0)   =   "Medicamentos_Cadastro.frx":7DA5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSalvar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdExcluir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAlterar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmCadastro"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Medicamentos_Cadastro.frx":7DC1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "frmFiltroComum"
      Tab(1).Control(2)=   "frmOrdemComum"
      Tab(1).Control(3)=   "frmCriterioComum"
      Tab(1).Control(4)=   "Grid"
      Tab(1).Control(5)=   "cmdExibir"
      Tab(1).Control(6)=   "cmdImprimir"
      Tab(1).Control(7)=   "Label20"
      Tab(1).Control(8)=   "lblProdutos"
      Tab(1).Control(9)=   "Label25"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "HISTÓRICO"
      TabPicture(2)   =   "Medicamentos_Cadastro.frx":7DDD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid_Estoque"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   -61380
         TabIndex        =   37
         Top             =   8580
         Width           =   135
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Estoque 
         Height          =   4395
         Left            =   -74940
         TabIndex        =   31
         Top             =   1620
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7752
         _Version        =   393216
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
         Height          =   735
         Left            =   -69000
         TabIndex        =   29
         Top             =   6180
         Width           =   4395
         Begin VB.ComboBox cboConsProduto 
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   300
            Visible         =   0   'False
            Width           =   4155
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
         Height          =   735
         Left            =   -71820
         TabIndex        =   22
         Top             =   6180
         Width           =   2775
         Begin VB.ComboBox cboOrdem 
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   300
            Width           =   2535
         End
      End
      Begin VB.Frame frmCriterioComum 
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   21
         Top             =   6180
         Width           =   3015
         Begin VB.ComboBox cboCriterio 
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   2715
         End
      End
      Begin VB.PictureBox frmCadastro 
         Enabled         =   0   'False
         Height          =   5055
         Left            =   180
         ScaleHeight     =   4995
         ScaleWidth      =   9855
         TabIndex        =   14
         Top             =   480
         Width           =   9915
         Begin VB.ComboBox cboCategoria 
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   2160
            MaxLength       =   90
            TabIndex        =   2
            Top             =   300
            Width           =   4755
         End
         Begin VB.ComboBox cboUnidMedida 
            Height          =   315
            Left            =   8640
            TabIndex        =   4
            Top             =   300
            Width           =   1095
         End
         Begin VB.ComboBox cboFabricante 
            Height          =   315
            Left            =   6960
            TabIndex        =   3
            Top             =   300
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
            TabIndex        =   25
            Top             =   1380
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.TextBox txtObs 
            Height          =   315
            Left            =   3900
            MaxLength       =   90
            TabIndex        =   8
            Top             =   960
            Width           =   5835
         End
         Begin VB.TextBox txtCodigo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9240
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtQuantMin 
            Height          =   315
            Left            =   2760
            TabIndex        =   7
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtPrateleira 
            Height          =   315
            Left            =   1740
            MaxLength       =   4
            TabIndex        =   6
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            Height          =   195
            Left            =   2160
            TabIndex        =   35
            Top             =   60
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unid. Med."
            Height          =   195
            Left            =   8640
            TabIndex        =   34
            Top             =   60
            Width           =   780
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   6960
            TabIndex        =   33
            Top             =   60
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. Barra"
            Height          =   195
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   750
         End
         Begin VB.Label Observação 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observação"
            Height          =   195
            Left            =   3900
            TabIndex        =   20
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Min."
            Height          =   195
            Left            =   2760
            TabIndex        =   17
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local"
            Height          =   195
            Left            =   1740
            TabIndex        =   16
            Top             =   720
            Width           =   390
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   23
         Top             =   420
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9340
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   555
         Left            =   9840
         TabIndex        =   11
         Top             =   1080
         Width           =   1875
         _ExtentX        =   3307
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
         MICON           =   "Medicamentos_Cadastro.frx":7DF9
         PICN            =   "Medicamentos_Cadastro.frx":7E15
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
         Left            =   10140
         TabIndex        =   12
         Top             =   1680
         Width           =   1875
         _ExtentX        =   3307
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
         MICON           =   "Medicamentos_Cadastro.frx":86EF
         PICN            =   "Medicamentos_Cadastro.frx":870B
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
         Left            =   10140
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
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
         MICON           =   "Medicamentos_Cadastro.frx":8A25
         PICN            =   "Medicamentos_Cadastro.frx":8A41
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
         Left            =   10140
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
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
         MICON           =   "Medicamentos_Cadastro.frx":F30B
         PICN            =   "Medicamentos_Cadastro.frx":F327
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
         Left            =   10140
         TabIndex        =   0
         Top             =   480
         Width           =   1875
         _ExtentX        =   3307
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
         MICON           =   "Medicamentos_Cadastro.frx":15DCB
         PICN            =   "Medicamentos_Cadastro.frx":15DE7
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
         Height          =   675
         Left            =   -64560
         TabIndex        =   27
         Top             =   6240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1191
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
         MICON           =   "Medicamentos_Cadastro.frx":16AC1
         PICN            =   "Medicamentos_Cadastro.frx":16ADD
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
         Left            =   -66780
         TabIndex        =   28
         Top             =   7080
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
         MICON           =   "Medicamentos_Cadastro.frx":173B7
         PICN            =   "Medicamentos_Cadastro.frx":173D3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   -64380
         TabIndex        =   41
         Top             =   5760
         Width           =   495
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
         Left            =   -63840
         TabIndex        =   40
         Top             =   5760
         Width           =   945
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Dê um duplo-clique para ver mais informações"
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   -74880
         TabIndex        =   24
         Top             =   5760
         Width           =   3435
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   32
      Top             =   9870
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17489
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21:27"
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
Attribute VB_Name = "Medicamentos_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

Dim var_cod_Preco As Long

Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Dim VarIncluirPreco As Integer


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
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("fornecedor"))
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("quant"))
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
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1600 '1600
      .ColWidth(3) = 4445 '4445
      .ColWidth(4) = 1500
      .ColWidth(5) = 850
      .ColWidth(6) = 800
      .ColWidth(7) = 2500
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "CÓD. BARRA"
      .TextMatrix(0, 3) = "PRODUTO"
      .TextMatrix(0, 4) = "FABRICANTE"
      .TextMatrix(0, 5) = "MED."
      .TextMatrix(0, 6) = "LOCAL"
      .TextMatrix(0, 7) = "CATEGORIA"
      
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
    
            .TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("var_desc"))
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("var_med"))
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("local"))
            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("categoria"))
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Redraw = True
      .Rows = .Rows - 1
   End With
   
   lblProdutos.Caption = Grid.Rows - 1  'contar o numeros de linhas no grid
End Sub

Private Sub LimparGrid_Produtos()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT produtos.codigo AS var_codent, produtos.descricao AS var_desc, " & _
      "produtos.codigo, produtos_entrada_itens.cod_produto " & _
      "FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.cod_produto " & _
      "WHERE 0 = 1;"
    
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid_Produtos r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub
Private Sub Preencher_ConsCriterio()
cboCriterio.Clear
cboCriterio.AddItem "TODOS"
cboCriterio.AddItem "COD. DE BARRA"
cboCriterio.AddItem "PRODUTO"
cboCriterio.AddItem "CATEGORIA"
   
cboCriterio.ListIndex = 0
End Sub

Private Sub Preencher_ConsOrdem()
cboOrdem.Clear
cboOrdem.AddItem "COD. DE BARRA"
cboOrdem.AddItem "DESCRIÇÃO"
cboOrdem.AddItem "CATEGORIA"
cboOrdem.ListIndex = 1
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

txtCodigo.Text = ValidateNull(rTabela("codigo"))
txtCodBarra.Text = ValidateNull(rTabela("cod_barra"))
txtDescricao.Text = ValidateNull(rTabela("descricao"))
cboFabricante.Text = ValidateNull(rTabela("fabricante"))
cboUnidMedida.Text = ValidateNull(rTabela("unid_medida"))
cboCategoria.Text = ValidateNull(rTabela("categoria"))
txtPrateleira.Text = ValidateNull(rTabela("local"))
txtQuantMin.Text = ValidateNull(rTabela("quant_min"))
txtObs.Text = ValidateNull(rTabela("observacao"))
chkAtivo.Value = Abs(CBool(rTabela("ativo")))
End Sub

Private Sub AutoNumeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_produto FROM produtos;"
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
   txtQuantMin.Text = ""
   txtObs.Text = ""
   chkAtivo.Value = Unchecked
   
   cmdNovo.Enabled = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
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
   If cboCriterio.Text = "COD. DE BARRA" And Len(cboConsProduto) = 13 Then
      cmdExibir_Click
   End If
End Sub

Private Sub cboConsProduto_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

If cboCriterio.Text = "COD. DE BARRA" Then
   cboConsProduto.Clear

ElseIf cboCriterio.Text = "PRODUTO" Then
   cboConsProduto.Clear
   
   sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsProduto.AddItem ValidateNull(r("descricao"))
      r.MoveNext
   Loop
   
ElseIf cboCriterio.Text = "CATEGORIA" Then
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

Private Sub cboCriterio_GotFocus()
Dim var_Texto As String
var_Texto = cboUnidMedida.Text

   cboCriterio.Clear
   cboCriterio.AddItem "TODOS"
   cboCriterio.AddItem "COD. DE BARRA"
   cboCriterio.AddItem "PRODUTO"
   cboCriterio.AddItem "CATEGORIA"
   
   moCombo.AttachTo cboCriterio
   
cboCriterio.Text = var_Texto
End Sub


Private Sub cboCriterio_Validate(Cancel As Boolean)
If cboCriterio.Text = "TODOS" Then
   cboConsProduto.Visible = False
   cboConsProduto.Visible = False
   'lblNomeCombo.Visible = False
ElseIf cboCriterio.Text = "COD. DE BARRA" Then
   cboConsProduto.Visible = True
   'lblNomeCombo.Visible = True
   'lblNomeCombo.Caption = "Cód. de Barra"
   cboConsProduto.SetFocus
ElseIf cboCriterio.Text = "PRODUTO" Then
   cboConsProduto.Visible = True
   'lblNomeCombo.Visible = True
   'lblNomeCombo.Caption = "Nome do Produto"
   cboConsProduto.SetFocus
ElseIf cboCriterio.Text = "CATEGORIA" Then
   cboConsProduto.Visible = True
   'lblNomeCombo.Visible = True
   'lblNomeCombo.Caption = "Categoria"
   cboConsProduto.SetFocus
End If
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

Private Sub cboOrdem_GotFocus()
Dim var_Texto As String
var_Texto = cboOrdem.Text

   cboOrdem.Clear
   cboOrdem.AddItem "COD. DE BARRA"
   cboOrdem.AddItem "DESCRIÇÃO"
   cboOrdem.AddItem "CATEGORIA"
   moCombo.AttachTo cboOrdem
   
cboOrdem.Text = var_Texto
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
   'dbData.Execute "UPDATE produtos_entrada_itens SET descricao = '" & txtDescricao.Text & "' WHERE (codigo_produto = " & txtCodigo.Text & ");"
   
   'alterar o nome dos produtos da tabela de entrada de pedidos
   'sSQL = "UPDATE produtos_entrada_itens SET VENDA = " & Replace(CCur(txtValorAtual.Text), ",", ".") & " WHERE (codigo = " & _
   '   "(SELECT codigo FROM (SELECT TOP 1 codigo FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodigo.Text & ") ORDER BY codigo DESC) as tempTabela));"

   'dbData.Execute sSQL
    
   cmdNovo.Enabled = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   frmCadastro.Enabled = False
   cmdExibir_Click
   Mostrar_Historico
End Sub

Private Function Inserir_Dados() As Boolean
   'A inclusão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   
   'Valida os campos
   If Trim(txtQuantMin.Text) = "" Then txtQuantMin.Text = 0
   
   'Comando de inclusão
   sSQL = "INSERT INTO produtos (" & _
      "codigo, ativo,  cod_barra, descricao, fabricante, unid_medida, " & _
      "categoria, local, quant_min, quant_estoque, observacao) VALUES (" & _
      txtCodigo.Text & ", " & Abs(chkAtivo.Value) & ", '" & _
      IIf((txtCodBarra.Text = ""), txtCodigo.Text, txtCodBarra.Text) & "', '" & _
      txtDescricao.Text & "', '" & cboFabricante.Text & "', '" & cboUnidMedida.Text & "', '" & _
      cboCategoria.Text & "', '" & txtPrateleira.Text & "', " & Replace(CDbl(txtQuantMin.Text), ",", ".") & ", 0, '" & _
      txtObs.Text & "');"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function
Private Function Atualizar_Dados() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE produtos SET " & _
      "ativo = " & Abs(chkAtivo.Value) & ", " & _
      "cod_barra = '" & IIf((txtCodBarra.Text = ""), txtCodigo.Text, txtCodBarra.Text) & "', " & _
      "descricao = '" & txtDescricao.Text & "', " & _
      "fabricante = '" & cboFabricante.Text & "', " & _
      "unid_medida = '" & cboUnidMedida.Text & "', " & _
      "categoria = '" & cboCategoria.Text & "', " & _
      "local = '" & txtPrateleira.Text & "', " & _
      "quant_min = " & Replace(CDbl(txtQuantMin.Text), ",", ".") & ", " & _
      "observacao = '" & txtObs.Text & "'"
   
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
   frmCadastro.Enabled = False
   LimparGrid_Produtos
   Mostrar_Historico
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
   REL_Prod_Cad_Imp.rfTIPO.Caption = lblProdutos.Caption
   REL_Prod_Cad_Imp.rfITENS.Caption = lblTotalUnid.Caption
   REL_Prod_Cad_Imp.rfVENDA.Caption = lblValorTotal.Caption
   
   'REL_Produtos.Relatorio.NomeImpressora = var_Impressora
   REL_Prod_Cad_Imp.Relatorio.Ativar
   Unload REL_Prod_Cad_Imp
Else
   Set REL_Produtos.Relatorio.Recordset = r
   REL_Produtos.dfQuant.Caption = "Quant.: " & lblQuantAtual.Caption
   REL_Produtos.dfBruto.Caption = "Bruto: " & lblValorAtual.Caption
   
   'REL_Produtos.Relatorio.NomeImpressora = var_Impressora
   REL_Produtos.Relatorio.Ativar
   Unload REL_Produtos
End If
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

AutoNumeracao

cboUnidMedida.Text = "UNID"
txtCodBarra.SetFocus
End Sub


Private Sub cmdSalvar_Click()
   'Não foi informado a descricao do produto.
   If txtDescricao.Text = "" Then
      ShowMsg "Digite a Descrição do produto", vbInformation
      txtDescricao.SetFocus
      Exit Sub
   End If
   
   'Faz a inserção de forma direta e verifica se houve algum erro
   If Not Inserir_Dados Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
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
   Dim sSQL As String
   Dim r As ADODB.Recordset

   'Indice
   Dim INDICE As String
   If cboOrdem.Text = "COD. DE BARRA" Then
      INDICE = "produtos.cod_barra;"
   ElseIf cboOrdem.Text = "DESCRIÇÃO" Then
      INDICE = "produtos.descricao;"
   ElseIf cboOrdem.Text = "CATEGORIA" Then
      INDICE = "produtos.CATEGORIA;"
   End If
   
   'Monta a consulta básica para não repetir várias linhas
   sSQL = "SELECT produtos.codigo AS var_codent, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
      "produtos.fabricante AS var_fab, local, categoria, produtos.unid_medida AS var_med " & _
      "FROM produtos  " & _
      "WHERE "
   
   If cboCriterio.Text = "COD. DE BARRA" Then
      sSQL = sSQL & "(produtos.cod_barra = '" & cboConsProduto.Text & "') AND (produtos.ativo = 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf cboCriterio.Text = "TODOS" Then
      sSQL = sSQL & "(produtos.ativo = 1) ORDER BY " & INDICE
      Debug.Print sSQL
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf cboCriterio.Text = "CATEGORIA" Then
      sSQL = sSQL & "(produtos.categoria = '" & cboConsProduto.Text & "') AND (produtos.ativo = 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   ElseIf cboCriterio.Text = "PRODUTO" Then
      sSQL = sSQL & "(produtos.descricao = '" & cboConsProduto.Text & "') AND (produtos.ativo = 1) ORDER BY " & INDICE
      Set r = dbData.OpenRecordset(sSQL)
      FormatarGrid_Produtos r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   End If
   
   If cboCriterio.Text <> "TODOS" Then
      SelectControl cboConsProduto
   End If

printSQL = sSQL
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

frmCriterioComum.Visible = True
frmOrdemComum.Visible = True
frmFiltroComum.Visible = True
Label3.Caption = "Categoria"
Preencher_ConsOrdem
Preencher_ConsCriterio
End Sub

Private Sub Mostrar_Historico()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Monta a consulta básica
   sSQL = "SELECT produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada.codigo AS var_codigo " & _
      "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.cod_entrada "
   
   'Define o filtro
   If txtCodigo.Text = "" Then
      sSQL = sSQL & "WHERE 1 = 0 "
      
   Else
      sSQL = sSQL & "WHERE (cod_produto = " & txtCodigo.Text & ") "
   
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

