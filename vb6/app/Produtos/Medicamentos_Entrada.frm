VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Medicamentos_Entrada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ENTRADA DE MEDICAMENTOS NO ESTOQUE"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "Medicamentos_Entrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   11625
      TabIndex        =   32
      Top             =   60
      Width           =   11655
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9480
         TabIndex        =   51
         Top             =   300
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTRADA DE MEDICAMENTOS NO ESTOQUE"
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
         TabIndex        =   33
         Top             =   240
         Width           =   6810
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Medicamentos_Entrada.frx":23D2
         Top             =   120
         Width           =   645
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7755
      Left            =   60
      TabIndex        =   16
      Top             =   1020
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13679
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
      TabPicture(0)   =   "Medicamentos_Entrada.frx":7DA5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdImprimirEntrada"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chameleonButton1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExcluir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNovo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAlterar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSalvar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frmPrincipal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "frmSecundario"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "HISTÓRICO"
      TabPicture(1)   =   "Medicamentos_Entrada.frx":7DC1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid_Historico"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CONSULTA"
      TabPicture(2)   =   "Medicamentos_Entrada.frx":7DDD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label25"
      Tab(2).Control(1)=   "cmdExibir"
      Tab(2).Control(2)=   "Frame10"
      Tab(2).Control(3)=   "Frame9"
      Tab(2).Control(4)=   "Data5"
      Tab(2).Control(5)=   "Data6"
      Tab(2).Control(6)=   "Grid"
      Tab(2).Control(7)=   "Frame1"
      Tab(2).Control(8)=   "Frame2"
      Tab(2).Control(9)=   "Frame3"
      Tab(2).ControlCount=   10
      Begin VB.Frame Frame3 
         Caption         =   "Critério"
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
         Left            =   -70500
         TabIndex        =   59
         Top             =   6720
         Width           =   2055
         Begin VB.ComboBox cboConsCriterio 
            Height          =   315
            Left            =   60
            TabIndex        =   60
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
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
         Left            =   -72420
         TabIndex        =   57
         Top             =   6720
         Width           =   1875
         Begin VB.ComboBox cboConsOrdem 
            Height          =   315
            Left            =   60
            TabIndex        =   58
            Top             =   480
            Width           =   1755
         End
      End
      Begin VB.PictureBox frmSecundario 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   5115
         Left            =   120
         ScaleHeight     =   5085
         ScaleWidth      =   11385
         TabIndex        =   43
         Top             =   1800
         Width           =   11415
         Begin VB.TextBox txtQuantAtual 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   285
            Width           =   1215
         End
         Begin VB.ComboBox cboDescricao 
            Height          =   315
            Left            =   2760
            TabIndex        =   7
            Top             =   285
            Width           =   6075
         End
         Begin VB.TextBox txtQuant 
            Height          =   315
            Left            =   10140
            TabIndex        =   9
            Top             =   285
            Width           =   1155
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   60
            TabIndex        =   6
            Top             =   300
            Width           =   2655
         End
         Begin VB.TextBox txtCodProduto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7920
            TabIndex        =   44
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Cadastro 
            Height          =   3915
            Left            =   60
            TabIndex        =   11
            Top             =   1080
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   6906
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionar 
            Height          =   315
            Left            =   8100
            TabIndex        =   10
            ToolTipText     =   "Adiciona"
            Top             =   720
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
            MICON           =   "Medicamentos_Entrada.frx":7DF9
            PICN            =   "Medicamentos_Entrada.frx":7E15
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
            Left            =   9720
            TabIndex        =   12
            ToolTipText     =   "Remove"
            Top             =   720
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
            MICON           =   "Medicamentos_Entrada.frx":81AF
            PICN            =   "Medicamentos_Entrada.frx":81CB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Atual"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   8880
            TabIndex        =   48
            Top             =   60
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            Height          =   195
            Left            =   2760
            TabIndex        =   47
            Top             =   60
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant."
            Height          =   195
            Index           =   0
            Left            =   10140
            TabIndex        =   46
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lblCodFabrica 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de Barra"
            Height          =   195
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   975
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Historico 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   42
         Top             =   420
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12726
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Frame Frame1 
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
         Height          =   915
         Left            =   -74940
         TabIndex        =   38
         Top             =   6720
         Width           =   2475
         Begin VB.ComboBox cboConsulta 
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   480
            Width           =   2355
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   5235
         Left            =   -74880
         TabIndex        =   37
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9234
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Data Data6 
         Caption         =   "Data6"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -73320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -73260
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
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
         Height          =   915
         Left            =   -68400
         TabIndex        =   30
         Top             =   6720
         Width           =   4095
         Begin VB.TextBox txtConsFornecedor 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   55
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1980
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboMES 
            Height          =   315
            ItemData        =   "Medicamentos_Entrada.frx":8565
            Left            =   120
            List            =   "Medicamentos_Entrada.frx":8567
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cboValor 
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Visible         =   0   'False
            Width           =   3915
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            Caption         =   "Ano"
            Height          =   195
            Left            =   1980
            TabIndex        =   63
            Top             =   240
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            Caption         =   "Mês"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblDiscriminacao 
            AutoSize        =   -1  'True
            Caption         =   "Descriminação"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Visible         =   0   'False
            Width           =   1050
         End
      End
      Begin VB.Frame Frame10 
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
         Left            =   -65880
         TabIndex        =   24
         Top             =   5580
         Width           =   2415
         Begin VB.Label lblValor 
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
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   900
            TabIndex        =   28
            Top             =   540
            Width           =   1365
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            Caption         =   "Valor:"
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
            TabIndex        =   27
            Top             =   540
            Width           =   795
         End
         Begin VB.Label lblQuant 
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
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   900
            TabIndex        =   26
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox frmPrincipal 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   120
         ScaleHeight     =   1245
         ScaleWidth      =   11385
         TabIndex        =   19
         Top             =   420
         Width           =   11415
         Begin VB.TextBox txtCodFuncionario 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10560
            TabIndex        =   54
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtCodFornecedor 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7680
            TabIndex        =   53
            Top             =   300
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboFuncionario 
            Height          =   315
            Left            =   8580
            TabIndex        =   5
            Top             =   540
            Width           =   2745
         End
         Begin VB.TextBox txtNotaFiscal 
            Height          =   315
            Left            =   1800
            TabIndex        =   3
            Top             =   540
            Width           =   1155
         End
         Begin VB.ComboBox cboFornecedor 
            Height          =   315
            Left            =   3000
            TabIndex        =   4
            Top             =   540
            Width           =   5565
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   315
            Left            =   60
            TabIndex        =   1
            Top             =   540
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskHora 
            Height          =   315
            Left            =   1140
            TabIndex        =   2
            Top             =   540
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastramento"
            Height          =   195
            Left            =   8580
            TabIndex        =   52
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label lblCadastrarFornecedores 
            AutoSize        =   -1  'True
            Caption         =   "Cadastrar fornecedores..."
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
            TabIndex        =   23
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota Fiscal"
            Height          =   195
            Left            =   1800
            TabIndex        =   22
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada"
            Height          =   195
            Left            =   75
            TabIndex        =   21
            Top             =   315
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
            Height          =   195
            Left            =   3000
            TabIndex        =   20
            Top             =   300
            Width           =   810
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvar 
         Height          =   555
         Left            =   1860
         TabIndex        =   13
         Top             =   7020
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
         MICON           =   "Medicamentos_Entrada.frx":8569
         PICN            =   "Medicamentos_Entrada.frx":8585
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
         Left            =   1860
         TabIndex        =   17
         Top             =   7020
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
         MICON           =   "Medicamentos_Entrada.frx":EE4F
         PICN            =   "Medicamentos_Entrada.frx":EE6B
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
         Left            =   120
         TabIndex        =   0
         Top             =   7020
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
         MICON           =   "Medicamentos_Entrada.frx":F745
         PICN            =   "Medicamentos_Entrada.frx":F761
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
         Left            =   3600
         TabIndex        =   14
         Top             =   7020
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
         MICON           =   "Medicamentos_Entrada.frx":1043B
         PICN            =   "Medicamentos_Entrada.frx":10457
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
         Left            =   3600
         TabIndex        =   18
         Top             =   7020
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
         MICON           =   "Medicamentos_Entrada.frx":16EFB
         PICN            =   "Medicamentos_Entrada.frx":16F17
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton chameleonButton1 
         Height          =   555
         Left            =   7080
         TabIndex        =   41
         Top             =   7020
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Es&tornar"
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
         MICON           =   "Medicamentos_Entrada.frx":17231
         PICN            =   "Medicamentos_Entrada.frx":1724D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimirEntrada 
         Height          =   555
         Left            =   5340
         TabIndex        =   50
         Top             =   7020
         Visible         =   0   'False
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
         MICON           =   "Medicamentos_Entrada.frx":17567
         PICN            =   "Medicamentos_Entrada.frx":17583
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
         Left            =   -64260
         TabIndex        =   64
         Top             =   6780
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
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
         MICON           =   "Medicamentos_Entrada.frx":1789D
         PICN            =   "Medicamentos_Entrada.frx":178B9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
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
         Height          =   195
         Left            =   -74880
         TabIndex        =   29
         Top             =   5640
         Width           =   3255
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   555
      Left            =   10020
      TabIndex        =   15
      Top             =   8820
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
      MICON           =   "Medicamentos_Entrada.frx":18193
      PICN            =   "Medicamentos_Entrada.frx":181AF
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
      Left            =   8280
      TabIndex        =   34
      Top             =   8820
      Visible         =   0   'False
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
      MICON           =   "Medicamentos_Entrada.frx":184C9
      PICN            =   "Medicamentos_Entrada.frx":184E5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCADProdutos 
      Height          =   555
      Left            =   1920
      TabIndex        =   39
      Top             =   8820
      Width           =   1815
      _ExtentX        =   3201
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
      MICON           =   "Medicamentos_Entrada.frx":187FF
      PICN            =   "Medicamentos_Entrada.frx":1881B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCADFornecedor 
      Height          =   555
      Left            =   60
      TabIndex        =   40
      Top             =   8820
      Width           =   1815
      _ExtentX        =   3201
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
      MICON           =   "Medicamentos_Entrada.frx":195AC
      PICN            =   "Medicamentos_Entrada.frx":195C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
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
      TabIndex        =   49
      Top             =   9450
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16431
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
Attribute VB_Name = "Medicamentos_Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer

Option Explicit

Private moCombo As cComboHelper
Private printSQL As String



Private Sub Calcular_Lucro_Venda()
   If txtCompraFinal.Text = "" Then txtCompraFinal.Text = 0
   If txtLucroReal.Text = "" Then txtLucroReal.Text = 0
   If txtImpostoRealVenda.Text = "" Then txtImpostoRealVenda.Text = 0
   If txtLucro.Text = "" Then txtLucro.Text = 0
   
   Dim COMPRA As Currency
   Dim LUCRO As Currency
   
   COMPRA = txtCompraFinal.Text
   
   If Right(txtLucro.Text, 1) = "%" Then
      LUCRO = Left$(txtLucro.Text, Len(txtLucro.Text) - 1)
   Else
      LUCRO = txtLucro.Text
   End If
   
   Dim Var_Lucro As Currency
   
   If chkLucroP.Value = 1 Then
      Var_Lucro = (COMPRA * LUCRO) / 100
   Else
      Var_Lucro = LUCRO
   End If
   
   txtLucroReal.Text = Format(Var_Lucro, ocMONEY)
End Sub

Private Sub Calcular_Imposto_Venda()
   If txtCompraFinal.Text = "" Then txtCompraFinal.Text = 0
   If txtLucroReal.Text = "" Then txtLucroReal.Text = 0
   If txtImpostoRealVenda.Text = "" Then txtImpostoRealVenda.Text = 0
   If txtImpostoVenda.Text = "" Then txtImpostoVenda.Text = 0
   
   Dim COMPRA As Currency
   Dim LUCRO As Currency
   Dim IMPOSTO As Currency
   Dim VALOR_VENDA As Currency
   
   COMPRA = txtCompraFinal.Text
   LUCRO = txtLucroReal.Text
   VALOR_VENDA = COMPRA
   
   If Right(txtImpostoVenda.Text, 1) = "%" Then
      IMPOSTO = Left$(txtImpostoVenda.Text, Len(txtImpostoVenda.Text) - 1)
   Else
      IMPOSTO = txtImpostoVenda.Text
   End If
   
   Dim Var_Imposto As Currency
   
   If chkImpostoVendaP.Value = 1 Then
      Var_Imposto = (VALOR_VENDA * IMPOSTO) / 100
   Else
      Var_Imposto = IMPOSTO
   End If
   
   txtImpostoRealVenda = Format(Var_Imposto, ocMONEY)
End Sub

Private Sub Calcular_Todos_Cadastrados()
   Dim sSQL As String
   If txtCodigo.Text = "" Then Exit Sub
   
   'Atualiza o custo de compra
   dbData.Execute "UPDATE produtos_entrada_itens SET custo_compra = custo + frete_valor_compra + CASE imposto_status_compra WHEN 1 THEN imposto_valor_compra ELSE ((custo * imposto_compra) / 100) END WHERE (codigo_entrada = " & txtCodigo.Text & ");"
   
   'Atualiza o lucro
   dbData.Execute "UPDATE produtos_entrada_itens SET lucro_valor = CASE lucro_status WHEN 1 THEN lucro_valor ELSE ((custo_compra * lucro) / 100) END WHERE (codigo_entrada = " & txtCodigo.Text & ");"
   
   'Atualiza o imposto de venda
   dbData.Execute "UPDATE produtos_entrada_itens SET imposto_valor_venda = CASE imposto_status_venda WHEN 1 THEN imposto_valor_venda ELSE (((custo_compra + lucro_valor) * imposto_venda) / 100) END WHERE (codigo_entrada = " & txtCodigo.Text & ");"
   
   'Atualiza o custo final
   dbData.Execute "UPDATE produtos_entrada_itens SET venda = custo_compra + lucro_valor + imposto_valor_venda WHERE (codigo_entrada = " & txtCodigo.Text & ");"
   
End Sub

Private Sub FormatarGrid_Historico(rTabela As ADODB.Recordset)
   Dim i As Integer, X As Integer
   
   With Grid_Historico
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 1000
      .ColWidth(2) = 1300
      .ColWidth(3) = 5300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      
      For X = 0 To .Cols - 1
         .Col = X
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "ENTRADA"
      .TextMatrix(0, 2) = "NOTA FISCAL"
      .TextMatrix(0, 3) = "FORNECEDOR"
      .TextMatrix(0, 4) = "ITENS"
      .TextMatrix(0, 5) = "FRETE"
      .TextMatrix(0, 6) = "VALOR"
      
      .Redraw = False
      i = 1
      
      If Not rTabela Is Nothing Then
        Do While Not rTabela.EOF
           .TextMatrix(.Rows - 1, 1) = Format(rTabela("data_entrada"), "dd/mm/yy")
           .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("notafiscal"))
           '.TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("COD_FORNECEDOR"))
           '.TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("itens"))
           '.TextMatrix(.Rows - 1, 5) = Format$(rTabela("frete"), ocMONEY)
           '.TextMatrix(.Rows - 1, 6) = Format$(rTabela("valor"), ocMONEY)
           
           rTabela.MoveNext
           .Rows = .Rows + 1
           i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 6
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub Limpar_Objetos()
If cmdAlterar.Visible = False Then txtCodigo.Text = ""
mskData.Mask = ""
mskData.Text = ""
mskHora.Mask = ""
mskHora.Text = ""
cboFornecedor.Text = ""
txtQuant.Text = ""
txtNotaFiscal.Text = ""
cboFuncionario.Text = ""
txtCodFornecedor.Text = ""
txtCodFuncionario.Text = ""
End Sub

Private Sub Limpar_SubDados()
   txtCodBarra.Text = ""
   cboDescricao.Text = ""
   txtCodProduto.Text = ""
   txtQuant.Text = ""
End Sub

Private Sub LimparGrid_Consulta()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT produtos_entrada.codigo AS var_codent, produtos_entrada.* FROM produtos_entrada WHERE 1 = 0;"

'Abre a consulta
Set r = dbData.OpenRecordset(sSQL)

'Exibe o resultado
FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Itens()
   Dim sSQL_Itens As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Then
      sSQL_Itens = "SELECT * FROM produtos_entrada_itens WHERE 1 = 0;"
      Set r = dbData.OpenRecordset(sSQL_Itens)
   Else
      sSQL_Itens = " SELECT produtos_entrada_itens.*, produtos_entrada_itens.codigo as varCod,  produtos.COD_BARRA as var_CodBarra, produtos.codigo, produtos.descricao as varDesc, produtos.fabricante  " & _
             " FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.COD_PRODUTO " & _
             " WHERE (cod_entrada = " & txtCodigo.Text & ") ORDER BY varDesc;"
      Set r = dbData.OpenRecordset(sSQL_Itens)
   End If
   
   printSQL = sSQL_Itens
   
   FormatarGrid_Itens r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
End Sub



Private Sub cboAno_GotFocus()
   Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
   Dim i As Integer
   
   'Calcula o intervalo de anos
   iAno = Year(Date)
   FirstYear = iAno - 2
   LastYear = iAno + 2
   
   'Limpa a lista
   cboAno.Clear
   
   'For x = iAno To FirstYear Step -1
   '   cboAno.AddItem x
   'Next
   '
   'iAno = iAno + 1
   'For x = iAno To LastYear
   '   cboAno.AddItem x
   'Next x
   
   For i = FirstYear To LastYear
      cboAno.AddItem i
   Next
End Sub

Private Sub cboConsCriterio_Change()
cboConsCriterio_LostFocus
End Sub

Private Sub cboConsCriterio_Click()
cboConsCriterio_LostFocus
End Sub


Private Sub cboConsCriterio_GotFocus()
Dim varTexto As String
varTexto = cboConsCriterio.Text

   cboConsCriterio.Clear
   cboConsCriterio.AddItem "TODOS"
   cboConsCriterio.AddItem "NOTA FISCAL"
   cboConsCriterio.AddItem "PRODUTO"
   cboConsCriterio.AddItem "FORNECEDOR"
   cboConsCriterio.AddItem "MENSAL"

cboConsCriterio.Text = varTexto
End Sub


Private Sub cboConsFornecedor_GotFocus()

End Sub

Private Sub Calcular_Valor_Compra()
   If txtCusto.Text = "" Then txtCusto.Text = 0
   If txtImpostoCompra.Text = "" Then Exit Sub
   If txtFrete.Text = "" Then Exit Sub
   
   Dim COMPRA As Currency
   Dim IMPOSTO As Currency
   Dim FRETE As Currency
   
   'Calcular_Frete
   
   If Right(txtFrete.Text, 1) = "%" Then
      FRETE = Left$(txtFrete.Text, Len(txtFrete.Text) - 1)
   Else
      FRETE = txtFrete.Text
   End If
   
   If Right(txtImpostoCompra.Text, 1) = "%" Then
      IMPOSTO = Left$(txtImpostoCompra.Text, Len(txtImpostoCompra.Text) - 1)
   Else
      IMPOSTO = txtImpostoCompra.Text
   End If
   
   COMPRA = txtCusto.Text
   
   Dim VALOR_COMPRA As Currency
   Dim Var_Imposto As Currency
   
   If chkImpostoCompraP.Value = 1 Then
      Var_Imposto = (COMPRA * IMPOSTO) / 100
   Else
      Var_Imposto = IMPOSTO
   End If
   
   Dim varFreteCompra As Double
   If chkImpostoFreteP.Value = 1 Then
      varFreteCompra = (COMPRA * FRETE) / 100
   Else
      varFreteCompra = FRETE
   End If
   
   VALOR_COMPRA = COMPRA + Var_Imposto + varFreteCompra
   
   txtFreteRealCompra = Format(varFreteCompra, ocMONEY)
   txtImpostoRealCompra = Format(Var_Imposto, ocMONEY)
   txtCompraFinal.Text = FormatCurrency(VALOR_COMPRA)
End Sub

Private Sub Calcular_Valor_Venda()
   If txtCompraFinal.Text = "" Then txtCompraFinal.Text = 0
   If txtLucroReal.Text = "" Then txtLucroReal.Text = 0
   If txtImpostoRealVenda.Text = "" Then txtImpostoRealVenda.Text = 0
   
   Dim COMPRA As Currency
   Dim LUCRO As Currency
   Dim IMPOSTO As Currency
   Dim VALOR_VENDA As Currency
   
   COMPRA = txtCompraFinal.Text
   LUCRO = txtLucroReal.Text
   IMPOSTO = txtImpostoRealVenda.Text
   
   VALOR_VENDA = COMPRA + LUCRO + IMPOSTO
   txtVenda.Text = FormatCurrency(VALOR_VENDA)
   'Calcular_Frete
End Sub

Private Sub cboConsFornecedor_LostFocus()
On Error GoTo TrataErro
   If cboConsFornecedor.Text = "" Then txtConsFornecedor.Text = "": Exit Sub
   'If cboFornecedor.ListIndex = -1 Then txtCodFornecedor.Text = "": Exit Sub
   
   txtConsFornecedor = cboConsFornecedor.ItemData(cboConsFornecedor.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboConsCriterio_LostFocus()
If cboConsCriterio.Text = "TODOS" Then
   lblDiscriminacao.Visible = False
   cboValor.Visible = False
   lblAno.Visible = False
   lblMes.Visible = False
   cboAno.Visible = False
   cboMES.Visible = False
ElseIf cboConsCriterio.Text = "NOTA FISCAL" Then
   lblDiscriminacao.Visible = True
   cboValor.Visible = True
   lblAno.Visible = False
   lblMes.Visible = False
   cboAno.Visible = False
   cboMES.Visible = False
ElseIf cboConsCriterio.Text = "PRODUTO" Then
   lblDiscriminacao.Visible = True
   cboValor.Visible = True
   lblAno.Visible = False
   lblMes.Visible = False
   cboAno.Visible = False
   cboMES.Visible = False
ElseIf cboConsCriterio.Text = "FORNECEDOR" Then
   lblDiscriminacao.Visible = True
   cboValor.Visible = True
   lblAno.Visible = False
   lblMes.Visible = False
   cboAno.Visible = False
   cboMES.Visible = False
ElseIf cboConsCriterio.Text = "MENSAL" Then
   lblDiscriminacao.Visible = False
   cboValor.Visible = False
   lblAno.Visible = True
   lblMes.Visible = True
   cboAno.Visible = True
   cboMES.Visible = True
End If
End Sub

Private Sub cboConsOrdem_GotFocus()
Dim varTexto As String
varTexto = cboConsOrdem.Text

   cboConsOrdem.Clear
   cboConsOrdem.AddItem "DATA"
   cboConsOrdem.AddItem "NO. DA NOTA"
   cboConsOrdem.AddItem "VALOR"
   cboConsOrdem.AddItem "FORNECEDOR"

cboConsOrdem.Text = varTexto
End Sub


Private Sub cboConsulta_Change()
cmdExibir_Click
End Sub

Private Sub cboConsulta_Click()
cmdExibir_Click
End Sub


Private Sub cboConsulta_GotFocus()
Dim varTexto As String
varTexto = cboConsulta.Text

   cboConsulta.Clear
   cboConsulta.AddItem "NOTA FISCAL"
   cboConsulta.AddItem "PRODUTO"
   cboConsulta.AddItem "ESTOQUE ANALITICO"

cboConsulta.Text = varTexto
End Sub


Private Sub cboDescricao_Click()
On Error GoTo TrataErro
Dim sSQL As String
Dim r As ADODB.Recordset

If cboDescricao.Text = "" Then txtCodProduto.Text = "": Exit Sub

txtCodProduto = cboDescricao.ItemData(cboDescricao.ListIndex)

'mostrar o codigo de BARRA
sSQL = "SELECT codigo, cod_barra, quant_estoque FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   txtCodBarra.Text = r("cod_barra")
   txtQuantAtual.Text = ValidateNull(r("quant_estoque"))
End If

If r.State <> 0 Then r.Close
Set r = Nothing
'Exit Sub
   
TrataErro:
  If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboDescricao_GotFocus()
   moCombo.AttachTo cboDescricao
End Sub

Private Sub cboDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboDescricao_LostFocus()
cboDescricao_Click
End Sub

Private Sub cboFornecedor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim var_Text As String
   Dim var_Cod As String
   
   var_Text = cboFornecedor.Text
   var_Cod = txtCodFornecedor.Text
   
   sSQL = "SELECT DISTINCT codigo, razao FROM fornecedor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   cboFornecedor.Clear
   
   Do While Not r.EOF
      cboFornecedor.AddItem r("razao")
      cboFornecedor.ItemData(cboFornecedor.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   cboFornecedor.Text = var_Text
   txtCodFornecedor.Text = var_Cod
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

   moCombo.AttachTo cboFornecedor
End Sub

Private Sub cboFornecedor_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFornecedor_LostFocus()
On Error GoTo TrataErro
   If cboFornecedor.Text = "" Then txtCodFornecedor.Text = "": Exit Sub
   'If cboFornecedor.ListIndex = -1 Then txtCodFornecedor.Text = "": Exit Sub
   
   txtCodFornecedor = cboFornecedor.ItemData(cboFornecedor.ListIndex)
   
Mostrar_Historico

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboFuncionario_GotFocus()
 Dim sSQL As String
   Dim r As ADODB.Recordset

   Dim var_Text As String
   Dim var_Cod As String
   
   var_Text = cboFuncionario.Text
   var_Cod = txtCodFuncionario.Text
   
   cboFuncionario.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM funcionario;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboFuncionario.AddItem r("nome")
      cboFuncionario.ItemData(cboFuncionario.NewIndex) = r("codigo")
      r.MoveNext
   Loop

   cboFuncionario.Text = var_Text
   txtCodFuncionario.Text = var_Cod
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFuncionario_LostFocus()
   On Error GoTo TrataErro
   If cboFuncionario.Text = "" Then txtCodFuncionario.Text = "": Exit Sub
   'If cboFuncionario.ListIndex = -1 Then txtCodFuncionario.Text = "": Exit Sub
   
   txtCodFuncionario = cboFuncionario.ItemData(cboFuncionario.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMES.Clear
   
   For vMes = 1 To 12
      cboMES.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMES
End Sub

Private Sub cboMes_LostFocus()
   If cboMES.Text = "" Then Exit Sub Else cboAno.SetFocus
End Sub

Private Sub cboProduto_GotFocus()

End Sub


Private Sub cboValor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   
If cboConsCriterio.Text = "TODOS" Then
   cboValor.Clear
ElseIf cboConsCriterio.Text = "NOTA FISCAL" Then
   cboValor.Clear
ElseIf cboConsCriterio.Text = "PRODUTO" Then
   'Limpa a lista
   cboValor.Clear
   
   sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboValor.AddItem ValidateNull(r("descricao"))
      cboValor.ItemData(cboValor.NewIndex) = r("codigo")
      r.MoveNext
   Loop
ElseIf cboConsCriterio.Text = "FORNECEDOR" Then
   Dim var_Text As String
   Dim var_Cod As String
   
   var_Text = cboValor.Text
   var_Cod = txtConsFornecedor.Text
   
   sSQL = "SELECT DISTINCT codigo, razao FROM fornecedor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   cboValor.Clear
   
   Do While Not r.EOF
      cboValor.AddItem r("razao")
      cboValor.ItemData(cboValor.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   cboValor.Text = var_Text
   txtConsFornecedor.Text = var_Cod
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

   moCombo.AttachTo cboValor
End If
   moCombo.AttachTo cboValor
End Sub


Private Sub cboValor_LostFocus()
On Error GoTo TrataErro

If cboConsCriterio.Text = "FORNECEDOR" Then
   If cboValor.Text = "" Then txtConsFornecedor.Text = "": Exit Sub
   'If cboFornecedor.ListIndex = -1 Then txtCodFornecedor.Text = "": Exit Sub
   
   txtConsFornecedor = cboValor.ItemData(cboValor.ListIndex)
End If

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cmdImprimirEntrada_Click()
   Dim r As ADODB.Recordset
   Dim var_Impressora As String
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   Set r = dbData.OpenRecordset(printSQL)
   
   Set REL_Prod_Entrada.Relatorio.Recordset = r
   REL_Prod_Entrada.dfData.Caption = mskData.Text & " - " & Format(mskHora.Text, "hh:mm")
   REL_Prod_Entrada.dfNota.Caption = txtNotaFiscal.Text
   REL_Prod_Entrada.dfFornecedor.Caption = cboFornecedor.Text
   'REL_Prod_Entrada.dfTotal.Caption = txtValor.Text
   REL_Prod_Entrada.Relatorio.Ativar
   Unload REL_Prod_Entrada
End Sub

Private Sub cmdAdicionar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Then Exit Sub
   
   If txtQuant.Text = "" Or txtQuant.Text = "0" Then
      ShowMsg "Insira uma quantidade válida!", vbExclamation
      txtQuant.SetFocus
      Exit Sub
   End If

   'VERIFICAR SE O PRODUTO FOI CADASTRADO
   If txtCodProduto.Text = "" Then Exit Sub
   
   sSQL = "SELECT * FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.BOF Then
      ShowMsg "O produto não consta no banco de dados. Cadastre-o!", vbExclamation
      Exit Sub
   End If
   
   'VERIFICAR A QUANTIDADE
   If txtCodigo.Text = "" Then Exit Sub
   
   Dim var_COD_ITENS As Long
   
   'AUTONUMERAÇÃO
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_entrada_itens;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then var_COD_ITENS = r("cod_itens") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'Inserir Medicamentos
sSQL = "INSERT INTO produtos_entrada_itens (" & _
      "codigo, " & _
      "cod_entrada, " & _
      "cod_produto, " & _
      "quant ) VALUES (" & _
      var_COD_ITENS & ", " & txtCodigo.Text & ", " & txtCodProduto.Text & ", " & _
      Replace(CDbl(txtQuant.Text), ",", ".") & ")"
      
   'Adiciona o registro
   dbData.Execute sSQL
   
   'Atualiza o estoque do produto
   dbData.Execute "UPDATE produtos SET quant_estoque =  quant_estoque + " & Replace(CDbl(txtQuant.Text), ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"
   
   Limpar_SubDados
   Mostrar_Itens
   On Local Error Resume Next
   txtCodBarra.SetFocus
End Sub

Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Then
      MsgBox "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte a NOTA FISCAL na guia CONSULTA.", vbInformation, "Aviso do Sistema"
      Exit Sub
   End If
   
   'Não é necessário consulta o registro antes de atualiza-lo
   sSQL = "SELECT * FROM produtos_entrada WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdImprimirEntrada.Visible = False
   cmdNovo.Enabled = True
   
   Limpar_Objetos
   Limpar_SubDados
   Limpar_Valores
   Mostrar_Itens
   cmdExibir_Click
   
   frmPrincipal.Enabled = False
   frmSecundario.Enabled = False
End Sub

Private Sub cmdCADFornecedor_Click()
   Fornecedor_Cadastro.Show 1
End Sub

Private Sub cmdCadProdutos_Click()
   'Dim oCfg As ConfigItem
   'Dim iOpcao As Integer
   
   'Substituiu a abertura da tabela de configuração
   
   'Set oCfg = sysConfig("PRODUTO")
   'iOpcao = oCfg.Value
   'Set oCfg = Nothing
   
   'Select Case iOpcao
      'Case 1
'         Produtos_Cadastro_ComEntrada.Show 1
      'Case 2
'         Produtos_Cadastro_SemEntrada.Show 1
      'Case 3
         Medicamentos_Cadastro.Show 1
   'End Select
End Sub

Private Sub cmdCancelar_Click()
   Dim i As Integer
   
   If txtCodigo.Text = "" Then Exit Sub
   
   If ShowMsg("Existe uma nota fiscal em aberto. Deseja sair e cancelar a entrada?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      'Cancel = 1
      Exit Sub
   End If
   
   With Grid_Cadastro
      For i = 1 To .Rows - 1
         dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo_entrada = " & txtCodigo.Text & ");"
         dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & CDbl(.TextMatrix(i, 6)) & ", ult_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) WHERE (codigo = " & CLng(.TextMatrix(i, 4)) & ");"
         dbData.Execute "DELETE FROM produtos_entrada WHERE (codigo = " & txtCodigo.Text & ");"
      Next
   End With
   
   Limpar_Objetos
   Limpar_SubDados
   Mostrar_Itens
   
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdNovo.Enabled = True
   frmPrincipal.Enabled = False
   frmSecundario.Enabled = False
End Sub

Private Sub cmdExcluir_Click()
   Dim i As Integer
   
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
   
   If txtCodigo.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Consulte Nota Fiscal na guia CONSULTA", vbInformation
      Exit Sub
   End If
   
   If ShowMsg("Excluir essa Nota Fiscal?", vbInformation + vbYesNo) = vbNo Then Exit Sub
   
   With Grid_Cadastro
      For i = 1 To .Rows - 1
         dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo_entrada = " & CLng(.TextMatrix(i, 2)) & ");"
         dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & CDbl(.TextMatrix(i, 6)) & ", ult_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) WHERE (codigo = " & CLng(.TextMatrix(i, 4)) & ");"
         dbData.Execute "DELETE FROM produtos_entrada WHERE (codigo = " & txtCodigo.Text & ");"
      Next
   End With
   
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdNovo.Enabled = True
   
   Limpar_Objetos
   Limpar_SubDados
   Limpar_Valores
   Mostrar_Itens
   cmdExibir_Click
   
   frmPrincipal.Enabled = False
   frmSecundario.Enabled = False
End Sub

Private Sub cmdExibir_Click()
If cboConsulta.Text = "" Or cboConsOrdem.Text = "" Or cboConsCriterio.Text = "" Then Exit Sub
   Dim INDICE As String
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   Dim fExibir As Integer
   
   'Seleciona a ordem dos registros
   If cboConsOrdem.Text = "NO. DA NOTA" Then
      INDICE = "notafiscal;"
   ElseIf cboConsOrdem.Text = "DATA" Then
      INDICE = "data_entrada;"
   ElseIf cboConsOrdem.Text = "VALOR" Then
      INDICE = "valor;"
   ElseIf cboConsOrdem.Text = "FORNECEDOR" Then
      INDICE = "fornecedor;"
   End If
   
   'Seleciona os registros
   If cboConsulta.Text = "NOTA FISCAL" Then
      fExibir = 0
          
      If cboConsCriterio.Text = "TODOS" Then
         sSQL = "SELECT produtos_entrada.codigo AS var_codent, produtos_entrada.*, fornecedor.codigo, fornecedor.razao as varFornecedor FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo ORDER BY " & INDICE
         
      ElseIf cboConsCriterio.Text = "NOTA FISCAL" Then
         If cboValor.Text = "" Then Exit Sub
         sSQL = "SELECT produtos_entrada.codigo AS var_codent, produtos_entrada.*, fornecedor.codigo, fornecedor.razao as varFornecedor FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo WHERE (notafiscal = " & cboValor.Text & ") ORDER BY " & INDICE
         
      ElseIf cboConsCriterio.Text = "FORNECEDOR" Then
         If cboValor.Text = "" Then Exit Sub
         sSQL = "SELECT produtos_entrada.codigo AS var_codent, produtos_entrada.*, fornecedor.codigo, fornecedor.razao as varFornecedor FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo WHERE (cod_fornecedor = " & txtConsFornecedor.Text & ") ORDER BY " & INDICE
         
      ElseIf cboConsCriterio.Text = "PRODUTO" Then
         If cboValor.Text = "" Then Exit Sub
         'sSQL = "SELECT produtos_entrada.codigo AS var_codent, produtos_entrada.*, produtos_entrada_itens.*, fornecedor.codigo, fornecedor.razao as varFornecedor FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo (INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada WHERE (descricao = '" & cboProduto.Text & "')) ORDER BY " & INDICE
         
      ElseIf cboConsCriterio.Text = "MENSAL" Then
         If Not ExistInList(cboMES) Then
            ShowMsg "Selecione o mês na lista.", vbExclamation
            Exit Sub
         End If
         
         If Not ExistInList(cboAno) Then
            ShowMsg "Selecione o ano na lista.", vbExclamation
            Exit Sub
         End If
      
         sSQL = "SELECT produtos_entrada.codigo AS var_codent, produtos_entrada.*, fornecedor.codigo, fornecedor.razao as varFornecedor FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo WHERE (MONTH(data_entrada) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_entrada) = " & cboAno & ") ORDER BY " & INDICE
         
      End If
      
      'Abre a consulta
      Set r = dbData.OpenRecordset(sSQL, totalRegistros)
      
      '===FUNÇÃO DE CONTAR REGISTROS
      lblQuant.Caption = Format(totalRegistros, "00")
      
      'Exibe o resultado
      FormatarGrid r
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
   ElseIf cboConsulta.Text = "PRODUTO" Then
      fExibir = 1
      
      If cboConsCriterio.Text = "TODOS" Then
         sSQL = "SELECT notafiscal, ref, produtos.fabricante as var_fab, produtos.tamanho as var_tam, produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada AS var_data, " & _
            "produtos_entrada_itens.descricao AS var_desc, produtos_entrada_itens.quant AS var_quant, " & _
            "produtos_entrada_itens.custo AS var_custo, produtos_entrada_itens.frete_compra AS var_frete, " & _
            "produtos_entrada_itens.imposto_valor_compra AS var_impcompra, produtos_entrada_itens.custo_compra AS var_vlrcompra, " & _
            "produtos_entrada_itens.lucro_valor AS var_lucro, produtos_entrada_itens.imposto_valor_venda AS var_impvenda, " & _
            "produtos_entrada_itens.venda AS var_vlrvenda, produtos_entrada.*, produtos_entrada_itens.* " & _
            "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
            "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
            "ORDER BY produtos_entrada_itens.descricao, " & INDICE
      ElseIf cboConsCriterio.Text = "MENSAL" Then
         If Not ExistInList(cboMES) Then
            ShowMsg "Selecione o mês na lista.", vbExclamation
            Exit Sub
         End If
         
         If Not ExistInList(cboAno) Then
            ShowMsg "Selecione o ano na lista.", vbExclamation
            Exit Sub
         End If
      
         sSQL = "SELECT notafiscal, ref, produtos.fabricante as var_fab, produtos.tamanho as var_tam, produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada AS var_data, " & _
            "produtos_entrada_itens.descricao AS var_desc, produtos_entrada_itens.quant AS var_quant, " & _
            "produtos_entrada_itens.custo AS var_custo, produtos_entrada_itens.frete_compra AS var_frete, " & _
            "produtos_entrada_itens.imposto_valor_compra AS var_impcompra, produtos_entrada_itens.custo_compra AS var_vlrcompra, " & _
            "produtos_entrada_itens.lucro_valor AS var_lucro, produtos_entrada_itens.imposto_valor_venda AS var_impvenda, " & _
            "produtos_entrada_itens.venda AS var_vlrvenda, produtos_entrada.*, produtos_entrada_itens.* " & _
            "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
            "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
            "WHERE (MONTH(data_entrada) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_entrada) = " & cboAno & ") " & _
            "ORDER BY produtos_entrada_itens.descricao, " & INDICE
         
      End If
      
      Set r = dbData.OpenRecordset(sSQL)
      
      'Exibe o resultado
      FormatarGrid r, True
      
      If r.State <> 0 Then r.Close
      Set r = Nothing

   ElseIf cboConsulta.Text = "ESTOQUE ANALITICO" Then
      Dim dIni As Date, dFim As Date
      Dim strData As String
      Dim DIA As Date
      Dim pInd As Integer
      
      Dim saldoInicial As Double
      Dim rEstoque() As String
      
      Dim saldoDia As Double
      Dim totalEntr As Double
      Dim totalSaida As Double
      
      If Not ExistInList(cboMES) Then
         ShowMsg "Selecione o mês na lista.", vbExclamation
         Exit Sub
      End If
      
      If Not ExistInList(cboAno) Then
         ShowMsg "Selecione o ano na lista.", vbExclamation
         Exit Sub
      End If
      
      'Período da pesquisa
      strData = "01/" & Format$(cboMES.ListIndex + 1, "00") & "/" & Format$(cboAno, "0000")
      dIni = CDate(strData)
      dFim = DateAdd("d", -1, DateAdd("m", 1, dIni))
      
      'Consulta o saldo inicial do perído
      sSQL = "SELECT codigo, descricao, " & _
         "(SELECT ISNULL(SUM(produtos_entrada_itens.quant), 0) FROM produtos_entrada_itens " & _
         "INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
         "WHERE (codigo_produto = produtos.codigo) AND (produtos_entrada.data_entrada < CONVERT(DATETIME, '" & Format(dIni, ocDATA) & "', 103))) - " & _
         "(SELECT ISNULL(SUM(quantidade), 0)  FROM produtos_entrada_itens WHERE cod_produto = produtos.Codigo " & _
         "AND (data < CONVERT(DATETIME, '" & Format$(dIni, ocDATA) & "', 103))) - " & _
         "(SELECT ISNULL(SUM(saida), 0) FROM produtos_saida WHERE (cod_produto = produtos.codigo) " & _
         "AND (data < CONVERT(DATETIME, '" & Format$(dIni, ocDATA) & "', 103))) AS estoque_inicial " & _
         "FROM produtos WHERE (codigo = " & cboValor.ItemData(cboValor.ListIndex) & ");"

      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then saldoInicial = r("estoque_inicial")
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      'Transfere o saldo inicial para o saldo do primeiro dia
      saldoDia = saldoInicial
      pInd = 1
      
      For DIA = dIni To dFim
         'Inicializa as variáveis
         totalEntr = 0
         totalSaida = 0
         
         'Consulta dia a dia
         sSQL = "SELECT codigo, descricao, " & _
            "(SELECT ISNULL(SUM(produtos_entrada_itens.quant), 0) FROM produtos_entrada_itens " & _
            "INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
            "WHERE (codigo_produto = produtos.codigo) AND (produtos_entrada.data_entrada = CONVERT(DATETIME, '" & Format$(DIA, ocDATA) & "', 103))) AS total_entrada, " & _
            "(SELECT ISNULL(SUM(quantidade), 0)  FROM produtos_entrada_itens WHERE (cod_produto = produtos.codigo) " & _
            "AND (data = CONVERT(DATETIME, '" & Format$(DIA, ocDATA) & "', 103))) + " & _
            "(SELECT ISNULL(SUM(saida ), 0) FROM produtos_saida WHERE (cod_produto = produtos.codigo) " & _
            "AND (data = CONVERT(DATETIME, '" & Format$(DIA, ocDATA) & "', 103))) AS total_saida " & _
            "FROM produtos WHERE (codigo = " & cboValor.ItemData(cboValor.ListIndex) & ");"

         Set r = dbData.OpenRecordset(sSQL)
         
         If Not r.BOF Then
            'Atribui os saldo para as variáveis
            totalEntr = r("total_entrada")
            totalSaida = r("total_saida")
         End If
         
         If r.State <> 0 Then r.Close
         Set r = Nothing
         
         'Calcula o saldo final do dia
         saldoDia = saldoDia + totalEntr - totalSaida
         
         'Monta a tabela com os valores
         ReDim Preserve rEstoque(1 To 4, 1 To pInd)
         rEstoque(1, pInd) = Format$(DIA, ocDATA)
         If totalEntr > 0 Then rEstoque(2, pInd) = Format$(totalEntr, ocPESO)
         If totalSaida > 0 Then rEstoque(3, pInd) = Format$(totalSaida, ocPESO)
         rEstoque(4, pInd) = Format$(saldoDia, ocPESO)
         
         'Incrementa o contador
         pInd = pInd + 1
      Next
      
      'Exibe o resultado
      FormatarGrid2 saldoInicial, rEstoque
End If
   
   printSQL = sSQL
End Sub
Private Sub cmdFechar_Click()
   If txtCodigo.Text <> "" And cmdSalvar.Visible = True Then
      ShowMsg "ENTRADA EM ABERTO!" & vbCrLf & "Clique no botão SALVAR ou no CANCELAR.", vbInformation
      Exit Sub
   Else
      Unload Me
   End If
   
End Sub

Private Sub cmdImprimir_Click()
   'colocar o nome da maquina na barra de status
   Dim oIni As Ini
   Dim var_Impressora As String
   Dim r As ADODB.Recordset
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   Me.Hide
   
   Set r = dbData.OpenRecordset(printSQL)
   
   If cboConsulta.Text = "NOTA FISCAL" Then
      Set REL_Prod_Entrada_Nota.Relatorio.Recordset = r
      REL_Prod_Entrada_Nota.dfQuant.Caption = lblQuant.Caption
      REL_Prod_Entrada_Nota.dfBruto.Caption = lblValor.Caption
      
      If cboConsCriterio.Text = "MENSAL" Then
         REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Mês = " & cboMES.Text & "/" & cboAno.Text
      ElseIf cboConsCriterio.Text = "FORNECEDOR" Then
         REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Produto = " & cboProduto.Text & ""
      ElseIf cboConsCriterio.Text = "FORNECEDOR" Then
         REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Fornecedor = " & cboConsFornecedor.Text & ""
      ElseIf cboConsCriterio.Text = "NOTA FISCAL" Then
         REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & txtConsNotaFiscal.Text & ""
      Else
         REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Todas as notas"
      End If
      
      REL_Prod_Entrada_Nota.Relatorio.Ativar
      Unload REL_Prod_Entrada_Nota
   
   ElseIf cboConsulta.Text = "PRODUTO" Then
      Set REL_Prod_Entrada_Produto.Relatorio.Recordset = r
      REL_Prod_Entrada_Produto.dfQuant.Caption = lblQuant.Caption
      REL_Prod_Entrada_Produto.dfBruto.Caption = lblValor.Caption
      
      If cboConsCriterio.Text = "MENSAL" Then
         REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Mês = " & cboMES.Text & "/" & cboAno.Text
      ElseIf cboConsCriterio.Text = "FORNECEDOR" Then
         REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Produto = " & cboProduto.Text & ""
      ElseIf cboConsCriterio.Text = "FORNECEDOR" Then
         REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Fornecedor = " & cboConsFornecedor.Text & ""
      ElseIf cboConsCriterio.Text = "NOTA FISCAL" Then
         REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & txtConsNotaFiscal.Text & ""
      Else
         REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Todas as notas"
      End If
      
      REL_Prod_Entrada_Produto.Relatorio.NomeImpressora = var_Impressora
      REL_Prod_Entrada_Produto.Relatorio.Ativar
      Unload REL_Prod_Entrada_Produto
   End If
   
   Me.Show 1
End Sub

Private Sub cmdNovo_Click()
frmPrincipal.Enabled = True
frmSecundario.Enabled = True
cmdSalvar.Visible = True
cmdCancelar.Visible = True
cmdNovo.Enabled = False
cmdAlterar.Visible = False
cmdExcluir.Visible = False
'cmdImprimirEntrada.Visible = True
Limpar_Objetos
Limpar_SubDados
mskData.Text = Format(Date, "dd/mm/yy")
mskHora.Text = Format(Now, "hh:mm")
Mostrar_Historico
Auto_Numeracao
Mostrar_Itens
mskData.SetFocus
End Sub

Private Sub cmdRemover_Click()
   If Grid_Cadastro.Rows <= 1 Then Exit Sub
   
   If Grid_Cadastro.TextMatrix(Grid_Cadastro.RowSel, 1) <> "" Then
      'dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo = " & Grid_Cadastro.TextMatrix(Grid_Cadastro.RowSel, 1) & ");"
      dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo = " & Grid_Cadastro.TextMatrix(Grid_Cadastro.Row, 1) & ") AND (cod_entrada = " & txtCodigo.Text & ");"
      dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Grid_Cadastro.TextMatrix(Grid_Cadastro.RowSel, 6) & " WHERE (codigo = " & Grid_Cadastro.TextMatrix(Grid_Cadastro.RowSel, 4) & ");"
   End If
   
   Mostrar_Itens
  End Sub

Private Sub cmdSalvar_Click()
   If txtCodigo.Text = "" Or cboFornecedor.Text = "" Or txtNotaFiscal.Text = "" Then
      ShowMsg "Dados Incompletos!", vbInformation
      txtNotaFiscal.SetFocus
      Exit Sub
   End If
   
   If Not Inserir_Dados Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos
   Limpar_SubDados
   Mostrar_Itens
   cmdExibir_Click
   
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdNovo.Enabled = True
   frmPrincipal.Enabled = False
   frmSecundario.Enabled = False
End Sub

Private Function Atualizar_Dados() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE produtos_entrada SET " & _
      "data_entrada = CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
      "hora_entrada = '" & Format$(mskHora.Text, ocHORA) & "', " & _
      "fornecedor = '" & cboFornecedor.Text & "', " & _
      "notafiscal = '" & txtNotaFiscal.Text & "', " & _
      "valor = " & Replace(CCur(txtValor.Text), ",", ".")
   
   'Condição para atualização
   sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Function Inserir_Dados() As Boolean
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO produtos_entrada (" & _
      "codigo, data_entrada, hora_entrada, cod_fornecedor, notafiscal, " & _
      "cod_funcionario) VALUES ("
   
   sSQL = sSQL & _
      txtCodigo & ", CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), '" & _
      Format$(mskHora.Text, ocHORA) & "', " & txtCodFornecedor.Text & ", '" & _
      txtNotaFiscal.Text & "', " & txtCodFuncionario & ");"
   
   'Retorna o resultado da atualização
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_entrada FROM produtos_entrada;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigo.Text = r("cod_entrada") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Form_Activate()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

Call PreencheProdutos

Mostrar_Itens
LimparGrid_Consulta
End Sub



Private Sub optFreteManual_Click()

End Sub

Private Sub txtCodFornecedor_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodFornecedor.Text = "" Then Exit Sub

If cmdAlterar.Visible = True Then
   sSQL = "SELECT codigo, razao FROM fornecedor WHERE (codigo = " & txtCodFornecedor.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cboFornecedor.Text = r("razao")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub

Private Sub txtCodFuncionario_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodFuncionario.Text = "" Then Exit Sub

If cmdAlterar.Visible = True Then
   sSQL = "SELECT codigo, nome FROM funcionario WHERE (codigo = " & txtCodFuncionario.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cboFuncionario.Text = r("nome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub


Private Sub txtCodProduto_Change()
If txtCodProduto.Text = "" Then
   txtQuantAtual.Text = ""
   cboDescricao.Text = ""
   txtCodBarra.Text = ""
End If


End Sub

Private Sub txtNotaFiscal_LostFocus()
   'Calcular_Frete
End Sub

Private Sub txtQuant_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset, Optional ByVal Agrupar As Boolean = False)
   Dim i As Integer, X As Integer
   
   Dim aux As String, iRow As Long
   Dim subtotalQtde As Double
   Dim bNovoGrupo As Boolean
   
   If cboConsulta.Text = "NOTA FISCAL" Then
      With Grid
         .Clear
         .Cols = 5
         .Rows = 2
         
         .ColWidth(0) = 0
         .ColWidth(1) = 0
         .ColWidth(2) = 900
         .ColWidth(3) = 9000
         .ColWidth(4) = 1200
         
         .TextMatrix(0, 1) = "COD"
         .TextMatrix(0, 2) = "DATA"
         .TextMatrix(0, 3) = "FORNECEDOR"
         .TextMatrix(0, 4) = "N. FISCAL"
         
         'colocar os cabeçalho em negrito
         For X = 0 To .Cols - 1
            .Col = X
            .Row = 0
            .CellFontBold = True
         Next
        
         'centralizar o titulo
         For X = 0 To .Cols - 1
            .Row = 0
            .Col = X
            .CellAlignment = flexAlignCenterCenter
         Next
         
         .Redraw = False
         i = 1
                  
         If Not rTabela Is Nothing Then
            Do While Not rTabela.EOF
               'ALINHAMENTO
               .ColAlignment(2) = 1
               
               .TextMatrix(.Rows - 1, 1) = rTabela("var_codent")
               .TextMatrix(.Rows - 1, 2) = Format$(rTabela("data_entrada"), "dd/mm/yy")
               .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("varFornecedor"))
               .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("notafiscal"))
               
               rTabela.MoveNext
               .Rows = .Rows + 1
               i = i + 1
            Loop
         End If
         
         .Rows = .Rows - 1
         .Redraw = True
      End With
   ElseIf cboConsulta.Text = "PRODUTO" Then
      With Grid
         .Clear
         .Cols = 12
         .Rows = 2
         
         .ColWidth(0) = 0
         .ColWidth(1) = 0
         .ColWidth(2) = 900
         .ColWidth(3) = 4250
         .ColWidth(4) = 750
         .ColWidth(5) = 750
         .ColWidth(6) = 750
         .ColWidth(7) = 750
         .ColWidth(8) = 750
         .ColWidth(9) = 750
         .ColWidth(10) = 750
         .ColWidth(11) = 750
         
         .TextMatrix(0, 1) = "COD"
         .TextMatrix(0, 2) = "DATA"
         .TextMatrix(0, 3) = "PRODUTO"
         .TextMatrix(0, 4) = "QUANT"
         .TextMatrix(0, 5) = "CUSTO"
         .TextMatrix(0, 6) = "FRETE"
         .TextMatrix(0, 7) = "IMP."
         .TextMatrix(0, 8) = "VALOR"
         .TextMatrix(0, 9) = "LUCRO"
         .TextMatrix(0, 10) = "IMP."
         .TextMatrix(0, 11) = "VENDA"
         
         'colocar os cabeçalho em negrito
         For X = 0 To .Cols - 1
            .Col = X
            .Row = 0
            .CellFontBold = True
         Next
         
         'ALINHAMENTO
         .ColAlignment(2) = 1
         
         'centralizar o titulo
         For X = 0 To .Cols - 1
            .Row = 0
            .Col = X
            .CellAlignment = flexAlignCenterCenter
         Next
         
         .Redraw = False
         i = 1
         
         'bNovoGrupo = True
         subtotalQtde = 0
         iRow = 1
         
         If Not rTabela Is Nothing Then
            'Atribui o nome do primeiro item do grupo
            'aux = rTabela("var_desc")
            
            Do While Not rTabela.EOF
               'mudar a cor da coluna
               'For i = 1 To .Rows - 1
               '   .Row = i
               '   .Col = 5:
               '   .CellBackColor = &HC0FFFF
               '   .Col = 11:
               '   .CellBackColor = &HC0C0FF
               'Next
               
               If Agrupar Then
                  If aux <> rTabela("var_desc") Then
                     .TextMatrix(iRow, 4) = Format$(subtotalQtde, ocPESO)
                     .TextMatrix(.Rows - 1, 3) = rTabela("var_desc")
                     
                     For i = 3 To 4
                        .Row = .Rows - 1
                        .Col = i
                        .CellFontBold = True
                     Next
                     
                     subtotalQtde = 0
                     iRow = .Rows - 1
                     .Rows = .Rows + 1
                  End If
               End If
               
               .TextMatrix(.Rows - 1, 1) = rTabela("var_codent")
               .TextMatrix(.Rows - 1, 2) = Format$(rTabela("var_data"), "dd/mm/yy")
            If tipoEmpresa = 4 Then
               .TextMatrix(.Rows - 1, 3) = "[" & Format$(rTabela("notafiscal"), "000,000") & "] " & rTabela("var_desc") & " /  " & rTabela("var_tam") & " / " & rTabela("var_fab")
            Else
               .TextMatrix(.Rows - 1, 3) = "[" & Format$(rTabela("notafiscal"), "000,000") & "] " & rTabela("var_desc")
            End If
               .TextMatrix(.Rows - 1, 4) = Format$(rTabela("var_quant"), ocMONEY)
               .TextMatrix(.Rows - 1, 5) = Format$(rTabela("var_custo"), ocMONEY)
               .TextMatrix(.Rows - 1, 6) = Format$(rTabela("var_frete"), ocMONEY)
               .TextMatrix(.Rows - 1, 7) = Format$(rTabela("var_impcompra"), ocMONEY)
               .TextMatrix(.Rows - 1, 8) = Format$(rTabela("var_vlrcompra"), ocMONEY)
               .TextMatrix(.Rows - 1, 9) = Format$(rTabela("var_lucro"), ocMONEY)
               .TextMatrix(.Rows - 1, 10) = Format$(rTabela("var_impvenda"), ocMONEY)
               .TextMatrix(.Rows - 1, 11) = Format$(rTabela("var_vlrvenda"), ocMONEY)
               
               aux = rTabela("var_Desc")
               'bNovoGrupo = False
               subtotalQtde = subtotalQtde + ValidateNull(rTabela("var_quant"))
               
               rTabela.MoveNext
               .Rows = .Rows + 1
               i = i + 1
            Loop
            
            .TextMatrix(iRow, 4) = Format$(subtotalQtde, ocPESO)
            
            For i = 3 To 4
               .Row = .Rows - 1
               .Col = i
               .CellFontBold = True
            Next
         End If
         
         'MUDAR COR DE FONTE DA COLUNA
         For X = 1 To .Rows - 1
            .Row = X
            .Col = 11
            .CellForeColor = &HC0&
            .CellFontBold = True
         Next
         
         .Rows = .Rows - 1
         .Redraw = True
      End With
   End If
End Sub

Private Sub FormatarGrid2(ByVal SaldoAnterior As Double, Movimento() As String)
   Dim i As Integer, X As Integer
   
   Dim aux As String, iRow As Long
   Dim subtotalQtde As Double
   Dim bNovoGrupo As Boolean
   
   With Grid
      .Clear
      .Cols = 6
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1200
      .ColWidth(3) = 1800
      .ColWidth(4) = 1800
      .ColWidth(5) = 1800
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "ENTRADAS"
      .TextMatrix(0, 4) = "SAÍDAS"
      .TextMatrix(0, 5) = "SALDO ATUAL"
      
      'colocar os cabeçalho em negrito
      For X = 0 To .Cols - 1
          .Col = X
          .Row = 0
          .CellFontBold = True
       Next
      
      'centralizar o titulo
      For X = 0 To .Cols - 1
         .Row = 0
         .Col = X
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      'Adiciona o saldo anterior
      .TextMatrix(.Rows - 1, 4) = "SALDO ANTERIOR"
      .TextMatrix(.Rows - 1, 5) = Format$(SaldoAnterior, ocPESO)
      
      .Row = .Rows - 1
      .Col = 5
      
      If CDbl(.TextMatrix(.Row, .Col)) > 0 Then
         .CellForeColor = RGB(0, 128, 0)
      ElseIf CDbl(.TextMatrix(.Row, .Col)) < 0 Then
         .CellForeColor = RGB(192, 0, 0)
      Else
         .CellForeColor = RGB(0, 0, 192)
      End If
      
      .CellFontBold = True
      .Rows = .Rows + 1
      
      For i = 1 To UBound(Movimento, 2)
         'ALINHAMENTO
         .ColAlignment(2) = 1
         
         '.TextMatrix(.Rows - 1, 1) = rTabela("var_codent")
         .TextMatrix(.Rows - 1, 2) = Movimento(1, i)
         .TextMatrix(.Rows - 1, 3) = Movimento(2, i)
         .TextMatrix(.Rows - 1, 4) = Movimento(3, i)
         .TextMatrix(.Rows - 1, 5) = Movimento(4, i)
         
         .Row = .Rows - 1
         .Col = 5
         
         If CDbl(Movimento(4, i)) > 0 Then
            .CellForeColor = RGB(0, 128, 0)
         ElseIf CDbl(Movimento(4, i)) < 0 Then
            .CellForeColor = RGB(192, 0, 0)
         Else
            .CellForeColor = RGB(0, 0, 192)
         End If
         
         .CellFontBold = True
         .Rows = .Rows + 1
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      'For X = 1 To .Rows - 1
      '   .Row = i
      '   .Col = 5
      '   .CellForeColor = &HC0&
      '   .CellFontBold = True
      'Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblValor.Caption = Format(SomaGrid(Grid, 5), ocMONEY)
End Sub

Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
Dim i As Integer, X As Integer

With Grid_Cadastro
   .Clear
   .Cols = 7        'numero de colunas
   .Rows = 2         'numero de linhas
   
   'largura da coluna
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 0
   .ColWidth(3) = 0
   .ColWidth(4) = 0
   .ColWidth(5) = 9000
   .ColWidth(6) = 1200
   
   'titulo das colunas
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "COD_ENTRADA"
   .TextMatrix(0, 3) = "COD_BARRA"
   .TextMatrix(0, 4) = "COD_PROD"
   .TextMatrix(0, 5) = "DESCRIÇÃO"
   .TextMatrix(0, 6) = "QTDE"
        
   'colocar os cabeçalho em negrito
   For X = 0 To .Cols - 1
      .Col = X
      .Row = 0
      .CellFontBold = True
   Next
   
   'centralizar o titulo
   For X = 0 To .Cols - 1
      .Row = 0
      .Col = X
      .CellAlignment = flexAlignCenterCenter
   Next

   .Redraw = False
   i = i + 1
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         'definir os campos para cada coluna
         .TextMatrix(.Rows - 1, 1) = rTabela("varCod")
         .TextMatrix(.Rows - 1, 2) = rTabela("cod_entrada")
         .TextMatrix(.Rows - 1, 4) = rTabela("cod_produto")
         .TextMatrix(.Rows - 1, 5) = rTabela("varDesc")
         .TextMatrix(.Rows - 1, 6) = rTabela("quant")
         
         rTabela.MoveNext
         .Rows = .Rows + 1
         i = i + 1
      Loop
   End If
   .Rows = .Rows - 1
   .Redraw = True
End With
End Sub
Private Sub Form_Load()
   SSTab1.Tab = 0
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   cmdNovo.Enabled = True
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdExibir.Visible = False
   cmdImprimir.Visible = False
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   
   PreencheProdutos
   
   Set moCombo = New cComboHelper
End Sub

Private Sub Mostrar_Historico()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cboFornecedor.Text = "" Then
      sSQL = "SELECT * FROM produtos_entrada WHERE 1 = 0"
   
   Else
      sSQL = "SELECT * FROM produtos_entrada WHERE (cod_fornecedor = '" & txtCodFornecedor.Text & "') ORDER BY data_entrada;"
   End If
   
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid_Historico r
   
   If Not r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Sub PreencheProdutos()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim var_cboTexto As String
   
   sSQL = "SELECT DISTINCT descricao, codigo, fabricante FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If cboDescricao.Text <> "" Then var_cboTexto = cboDescricao.Text
   cboDescricao.Clear
   cboDescricao.Text = var_cboTexto
   
   Do While Not r.EOF
      If tipoEmpresa = 4 Then
          cboDescricao.AddItem ValidateNull(r("descricao")) & " / " & ValidateNull(r("fabricante"))
      Else
         cboDescricao.AddItem ValidateNull(r("descricao"))
      End If
      cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'CHECAR SE O PEDIDO ESTÁ FECHADO
   If txtCodigo.Text = "" Then Exit Sub
   If Grid_Cadastro.Rows >= 1 And cmdNovo.Enabled = False Then cmdCancelar_Click
   Set moCombo = Nothing
End Sub

Private Sub frmPrincipal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblCadastrarFornecedores.FontBold = False
   lblCadastrarFornecedores.ForeColor = vbBlack
End Sub

Private Sub Grid_DblClick()
SSTab1.Tab = 0
frmPrincipal.Enabled = True
frmSecundario.Enabled = True
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = True
cmdExcluir.Visible = True
cmdImprimirEntrada.Visible = True
cmdNovo.Enabled = True
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub lblCadastrarFornecedores_Click()
   Fornecedor_Cadastro.Show
End Sub

Private Sub lblCadastrarFornecedores_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblCadastrarFornecedores.FontBold = True
   lblCadastrarFornecedores.ForeColor = vbRed
End Sub

Private Sub mskData_GotFocus()
   SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
   If mskData.Text = "" Or mskData.Text = "__/__/____" Then
      mskData.Mask = ""
      mskData.Text = ""
   Else
      If IsDate(mskData.Text) Then
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskData.SetFocus
      End If
   End If
End Sub

Private Sub mskHora_GotFocus()
   SelectControl mskHora
End Sub

Private Sub optExibNotas_Click()
   
End Sub

Private Sub OptFornecedor_Click()
   cboProduto.Enabled = False
   txtConsNotaFiscal.Enabled = False
   cboMES.Enabled = False
   cboAno.Enabled = False
   cboConsFornecedor.Enabled = True
   cboConsFornecedor.SetFocus
End Sub

Private Sub optMensal_Click()
   cboProduto.Enabled = False
   txtConsNotaFiscal.Enabled = False
   cboConsFornecedor.Enabled = False
   cboMES.Enabled = True
   cboAno.Enabled = True
   cboMES.SetFocus
End Sub

Private Sub optNotaFiscal_Click()
   cboProduto.Enabled = False
   cboConsFornecedor.Enabled = False
   cboMES.Enabled = False
   cboAno.Enabled = False
   txtConsNotaFiscal.Enabled = True
   txtConsNotaFiscal.SetFocus
End Sub

Private Sub optProduto_Click()
   cboProduto.Enabled = True
   txtNotaFiscal.Enabled = False
   cboFornecedor.Enabled = False
   cboMES.Enabled = False
   cboAno.Enabled = False
   cboProduto.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      cmdExibir.Visible = False
      cmdImprimir.Visible = False
   ElseIf SSTab1.Tab = 1 Then
      cmdExibir.Visible = False
      cmdImprimir.Visible = False
   ElseIf SSTab1.Tab = 2 Then
      cmdExibir.Visible = True
      cmdImprimir.Visible = True
   End If
End Sub

Private Sub txtCodBarra_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodBarra_LostFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodBarra.Text = "" Then txtCodProduto.Text = "": Exit Sub

sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, fabricante, quant_estoque FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "') AND (ativo = 1);"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   txtCodProduto.Text = r("var_codprod")
   cboDescricao.Text = ValidateNull(r("var_desc")) & " / " & ValidateNull(r("fabricante"))
   txtQuantAtual.Text = ValidateNull(r("quant_Estoque"))
Else
   ShowMsg "Produto Inexistente!", vbCritical
   txtCodBarra.Text = ""
   txtCodBarra.SetFocus
   Exit Sub
End If

On Local Error Resume Next
txtQuant.SetFocus
End Sub

Private Sub txtCodigo_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cmdSalvar.Visible = False Then
      If txtCodigo.Text = "" Then Exit Sub
      
      sSQL = "SELECT * FROM produtos_entrada WHERE (codigo = " & txtCodigo.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If r.EOF Then Exit Sub
      
      Limpar_Objetos
      Mostrar_Dados r
      Mostrar_Itens
      Mostrar_Historico
      mskData.SetFocus
   End If
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
   If Not rTabela Is Nothing Then
      txtCodigo.Text = ValidateNull(rTabela("codigo"))
      mskData.Text = Format$(rTabela("data_entrada"), "dd/mm/yy")
      mskHora.Text = Format$(rTabela("hora_entrada"), ocHORA)
      txtCodFornecedor.Text = ValidateNull(rTabela("cod_fornecedor"))
      txtCodFuncionario.Text = ValidateNull(rTabela("cod_funcionario"))
      txtNotaFiscal.Text = ValidateNull(rTabela("notafiscal"))
   End If
End Sub

Private Sub txtQuant_GotFocus()
   SelectControl txtQuant
End Sub

