VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_Estoque_Simples 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESTOQUE SIMPLES"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16275
   Icon            =   "Produtos_Estoque_Simples.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Produtos_Estoque_Simples.frx":1D82
   ScaleHeight     =   10035
   ScaleWidth      =   16275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmConsulta 
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   1815
      Left            =   60
      TabIndex        =   27
      Top             =   7920
      Width           =   10095
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
         ForeColor       =   &H00000080&
         Height          =   1515
         Left            =   1500
         TabIndex        =   55
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton optMostrarFiscal 
            Caption         =   "Fiscal"
            Height          =   195
            Left            =   60
            TabIndex        =   60
            Top             =   960
            Width           =   795
         End
         Begin VB.OptionButton optMostrarTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   60
            TabIndex        =   59
            Top             =   780
            Width           =   795
         End
         Begin VB.OptionButton optMostrarZerados 
            Caption         =   "Zerados"
            Height          =   195
            Left            =   60
            TabIndex        =   58
            Top             =   600
            Width           =   915
         End
         Begin VB.OptionButton optMostrarNegativos 
            Caption         =   "Negativos"
            Height          =   195
            Left            =   60
            TabIndex        =   57
            Top             =   420
            Width           =   1095
         End
         Begin VB.OptionButton optMostrarQuant 
            Caption         =   "Com quantidade"
            Height          =   195
            Left            =   60
            TabIndex        =   56
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
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
         ForeColor       =   &H00000080&
         Height          =   1515
         Left            =   3240
         TabIndex        =   44
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optORDDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   180
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optORDQuant 
            Caption         =   "Quant."
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   1035
         End
         Begin VB.OptionButton ORDQuantFiscal 
            Caption         =   "Quant. Fiscal"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.OptionButton optORDValor 
            Caption         =   "Valor Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   540
            Width           =   1275
         End
         Begin VB.OptionButton optORDLinha 
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   720
            Width           =   1035
         End
         Begin VB.OptionButton optORDValorCusto 
            Caption         =   "Valor Custo"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   900
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Frame Frame4 
            Caption         =   "Direção"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   675
            Left            =   1500
            TabIndex        =   46
            Top             =   180
            Width           =   975
            Begin VB.OptionButton optORDASC 
               Caption         =   "Asc"
               Height          =   195
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton optORDDescrescente 
               Caption         =   "Desc"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   420
               Width           =   675
            End
         End
         Begin VB.OptionButton optORDTFiscal 
            Caption         =   "Total Fiscal"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1260
            Visible         =   0   'False
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Busca Avançada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   5820
         TabIndex        =   35
         Top             =   240
         Width           =   4185
         Begin VB.ComboBox cboDesc 
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.OptionButton PorPalavraDupla 
            Caption         =   "Palavras Duplas"
            Height          =   195
            Left            =   2400
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton optPorPalavra 
            Caption         =   "Palavra"
            Height          =   195
            Left            =   1500
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   975
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizar 
            Height          =   315
            Left            =   3480
            TabIndex        =   40
            Top             =   480
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
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
            MICON           =   "Produtos_Estoque_Simples.frx":264C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   180
            TabIndex        =   43
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCategoria 
            Caption         =   "Categoria"
            Height          =   195
            Left            =   180
            TabIndex        =   42
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCodBarra 
            Caption         =   "Cód. de Barra"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
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
         ForeColor       =   &H00000080&
         Height          =   1515
         Left            =   60
         TabIndex        =   28
         Top             =   240
         Width           =   1395
         Begin VB.OptionButton optCodBarra 
            Caption         =   "Cód. Barra"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   420
            Width           =   1155
         End
         Begin VB.OptionButton optDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   1035
         End
         Begin VB.OptionButton optCategoria 
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   780
            Width           =   1035
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optTags 
            Caption         =   "Tags"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   1035
         End
         Begin VB.OptionButton optNCM 
            Caption         =   "NCM"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1140
            Width           =   1035
         End
      End
   End
   Begin VB.Frame frmAlterarGrupos 
      Caption         =   "Alterar em grupos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   10200
      TabIndex        =   15
      Top             =   7920
      Width           =   6015
      Begin VB.Frame frmEdicaoFiltros 
         Caption         =   "Alterar"
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
         Height          =   1515
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1395
         Begin VB.OptionButton optEICMS 
            Caption         =   "ICMS CST"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1155
         End
         Begin VB.OptionButton optENCM 
            Caption         =   "NCM"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1155
         End
         Begin VB.OptionButton optECFOP 
            Caption         =   "CFOP"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   420
            Width           =   1035
         End
         Begin VB.OptionButton optECategoria 
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   780
            Width           =   1035
         End
         Begin VB.OptionButton optETags 
            Caption         =   "Tags"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1035
         End
      End
      Begin VB.Frame frmEdicao 
         Caption         =   "Alterar em todos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   4395
         Begin VB.ComboBox cboEdicaoColetiva 
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.TextBox txtEdicaoColetiva 
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   3015
         End
         Begin ChamaleonBtn.chameleonButton cmdEdicaoColetiva 
            Height          =   315
            Left            =   3120
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Atualizar Todos"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Produtos_Estoque_Simples.frx":2668
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblEdicaoColetiva 
            AutoSize        =   -1  'True
            Caption         =   "Titulo"
            Height          =   195
            Left            =   60
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   390
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Left            =   1620
         TabIndex        =   63
         Top             =   1500
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Nenhum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   62
         Top             =   1500
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblCodUsuario 
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
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   61
         Top             =   1320
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6300
      Picture         =   "Produtos_Estoque_Simples.frx":2684
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   3540
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox cboEdit 
      Height          =   315
      Left            =   2520
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   16125
      TabIndex        =   0
      Top             =   60
      Width           =   16155
      Begin VB.Frame frmSenha 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   10020
         TabIndex        =   64
         Top             =   60
         Visible         =   0   'False
         Width           =   1935
         Begin VB.TextBox txtSenha 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   65
            Top             =   360
            Width           =   1335
         End
         Begin ChamaleonBtn.chameleonButton cmdSenha 
            Height          =   315
            Left            =   1500
            TabIndex        =   66
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "OK"
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
            MICON           =   "Produtos_Estoque_Simples.frx":36BC
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
      Begin VB.Frame frmTotalFiscal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Totais"
         Height          =   675
         Left            =   13200
         TabIndex        =   12
         Top             =   60
         Visible         =   0   'False
         Width           =   2895
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total Fiscal:"
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
            TabIndex        =   14
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label lblValorTotalFiscal 
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
            Left            =   1260
            TabIndex        =   13
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Produtos_Estoque_Simples.frx":36D8
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AJUSTE DE ESTOQUE"
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
         TabIndex        =   1
         Top             =   240
         Width           =   3360
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   315
      Left            =   14520
      TabIndex        =   5
      Top             =   7500
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
      MICON           =   "Produtos_Estoque_Simples.frx":90AB
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
      TabIndex        =   6
      Top             =   9765
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24368
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 9427-5280"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 9427-5280"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "14:53"
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6435
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   16155
      _ExtentX        =   28496
      _ExtentY        =   11351
      _Version        =   393216
      Cols            =   5
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizar 
      Height          =   315
      Left            =   12780
      TabIndex        =   7
      Top             =   7500
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Salvar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   128
      FCOLO           =   128
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Produtos_Estoque_Simples.frx":90C7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizarPreco 
      Height          =   315
      Left            =   9180
      TabIndex        =   8
      Top             =   7500
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Alterar Preço"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   128
      FCOLO           =   128
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Produtos_Estoque_Simples.frx":90E3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizarQuant 
      Height          =   315
      Left            =   10920
      TabIndex        =   9
      Top             =   7500
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Alterar Quantidade"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   128
      FCOLO           =   128
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Produtos_Estoque_Simples.frx":90FF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgDesmarcadaTODAS 
      Height          =   195
      Left            =   60
      Picture         =   "Produtos_Estoque_Simples.frx":911B
      Top             =   7500
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image ImgMarcada 
      Height          =   195
      Left            =   4440
      Picture         =   "Produtos_Estoque_Simples.frx":B497
      Top             =   7500
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDesmarcada 
      Height          =   195
      Left            =   4680
      Picture         =   "Produtos_Estoque_Simples.frx":D896
      Top             =   7500
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image ImgMarcadaTODAS 
      Height          =   195
      Left            =   60
      Picture         =   "Produtos_Estoque_Simples.frx":FC12
      Top             =   7500
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblMarcarTodas 
      AutoSize        =   -1  'True
      Caption         =   "Marcar Todas"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   7500
      Width           =   990
   End
End
Attribute VB_Name = "Produtos_Estoque_Simples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Private moCombo As cComboHelper
Private iRow As Long, iCol As Long
Private tipoEmpresa As Long
Dim varTipoValorVenda As String
Dim var_Indice As String
Dim var_Direcao As String

'arquivo .ini
Public cCfg As ConfigItem
'Public oIni As Ini


Private Sub LimparGrid2()
 
sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.TAGS AS var_tags, produtos.fabricante AS var_fab, produtos.PRATELEIRA AS var_Local, " & _
   "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, produtos.ESTOQUE_FISCAL AS var_EstoqueFiscal, produtos.UNID_MEDIDA AS var_UnidMed, " & _
   "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
   "FROM produtos " & _
   "WHERE 1 = 0"

Set r = dbData.OpenRecordset(sSQL)

Formatar_Grid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub LimparGrid()
   Dim i As Integer
   
   txtEdit.Text = ""
   
   With Grid
      .Clear
      .Cols = 9
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1500
      .ColWidth(4) = 4200
      .ColWidth(5) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 1000
      .ColWidth(8) = 2000
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "QUANT."
      .TextMatrix(0, 6) = "MIN."
      .TextMatrix(0, 7) = "VENDA"
      .TextMatrix(0, 8) = "CATEGORIA"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      'ALINHAMENTO
      .ColAlignment(2) = 1
      
      .rows = .rows + 1
      
      .rows = .rows - 1
      .Redraw = True
   End With
End Sub

Private Sub MostrarCriterios()
   Dim var_Criterio As String
   
   
   var_Criterio = ""
   
   If optDesc.Value Then
      Dim sW1 As String
      Dim sW2 As String
      Dim sPosEsp As Integer
      If optPorPalavra.Value Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao LIKE '%" & Replace(cboDesc.Text, "'", "''") & "%'"
      ElseIf PorPalavraDupla.Value Then
         sPosEsp = InStr(Trim(cboDesc.Text), " ")
         If sPosEsp > 0 Then
            sW1 = Left(Trim(cboDesc.Text), sPosEsp - 1)
            sW2 = Trim(Mid(cboDesc.Text, sPosEsp + 1))
         Else
            sW1 = Trim(cboDesc.Text)
            sW2 = ""
         End If
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao LIKE '%" & Replace(sW1, "'", "''") & "%'"
         If Len(sW2) > 0 Then
            var_Criterio = var_Criterio & " AND produtos.descricao LIKE '%" & Replace(sW2, "'", "''") & "%'"
         End If
      End If
   End If
   
   var_Criterio = var_Criterio & IIf(optCategoria.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = '" & cboDesc.Text & "'", "")
   var_Criterio = var_Criterio & IIf(optCodBarra.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = '" & txtCodBarra.Text & "'", "")
   var_Criterio = var_Criterio & IIf(optTags.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.TAGS = '" & cboDesc.Text & "'", "")
   var_Criterio = var_Criterio & IIf(optNCM.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.NCM = '" & txtCodBarra.Text & "'", "")
   
   If var_Criterio <> "" Then var_Criterio = " WHERE (produtos.ativo = 1) and " & var_Criterio
   
   var_Indice = ""
   var_Indice = var_Indice & IIf(optORDQuant.Value, IIf(var_Indice <> "", ", ", "") & "quant_estoque", "")
   var_Indice = var_Indice & IIf(optORDDesc.Value, IIf(var_Indice <> "", ", ", "") & "produtos.descricao", "")
   var_Indice = var_Indice & IIf(ORDQuantFiscal.Value, IIf(var_Indice <> "", ", ", "") & "quant_min", "")
   var_Indice = var_Indice & IIf(optORDValor.Value, IIf(var_Indice <> "", ", ", "") & "produtos_entrada_itens.venda", "")
   var_Indice = var_Indice & IIf(optORDLinha.Value, IIf(var_Indice <> "", ", ", "") & "produtos.categoria", "")
   
   If var_Indice <> "" Then var_Indice = " ORDER BY " & var_Indice
   
   sSQL = "SELECT produtos.NCM AS var_NCM, produtos.ICMSCST AS var_ICMS, produtos.CFOP AS var_CFOP, produtos.categoria AS var_cat, produtos.TAGS AS var_tags, produtos.fabricante AS var_fab, produtos.PRATELEIRA AS var_Local, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, produtos.ESTOQUE_FISCAL AS var_EstoqueFiscal, produtos.UNID_MEDIDA AS var_UnidMed, " & _
      "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS custo, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & var_Criterio & " " & var_Indice
   
   Set r = dbData.OpenRecordset(sSQL)
   'Debug.Print sSQL
   
       If r.RecordCount > 32000 Then
        MsgBox "A Consulta retornou um valor maior de registros que é permitido na grade!", vbInformation, "Aviso do sistema"
        LimparGrid2
        Exit Sub
    Else
        Formatar_Grid r
    End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
End Sub

Private Sub optORDASC_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDDesc_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDDescrescente_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDLinha_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDQuant_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDTFiscal_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDValorCusto_Click()
cmdLocalizar_Click
End Sub

Private Sub ORDQuantFiscal_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDValor_Click()
cmdLocalizar_Click
End Sub

Private Sub cmdAtualizar_Click()
Dim i As Integer

picAguarde.Visible = True
DoEvents

Dim iSemCatTag As Integer
Dim jv As Integer
For jv = 1 To Grid.rows - 1
   If Len(Trim(Grid.TextMatrix(jv, 8))) > 0 And Len(Trim(Grid.TextMatrix(jv, 7))) = 0 Then
      iSemCatTag = iSemCatTag + 1
   End If
Next jv
If iSemCatTag > 0 Then
   MsgBox iSemCatTag & " produto(s) com tag mas sem categoria. Defina a categoria antes de salvar.", vbExclamation, "Tag sem categoria"
   picAguarde.Visible = False
   Exit Sub
End If

    'txtDescricao.Text = TirarEspaco(txtDescricao.Text)
For i = 1 To Grid.rows - 1
   'Atualiza a tabela de produtos
   sSQL = "UPDATE produtos SET " & _
      "cod_barra = '" & Replace(Trim(Grid.TextMatrix(i, 3)), "'", "''") & "', " & _
      "descricao = '" & Replace(TirarEspaco(Grid.TextMatrix(i, 4)), "'", "''") & "', " & _
      "UNID_MEDIDA = '" & Replace(Grid.TextMatrix(i, 6), "'", "''") & "', " & _
      "categoria = '" & Replace(Grid.TextMatrix(i, 7), "'", "''") & "', " & _
      "TAGS = '" & Replace(Grid.TextMatrix(i, 8), "'", "''") & "', " & _
      "fabricante = '" & Replace(Grid.TextMatrix(i, 5), "'", "''") & "', " & _
      "PRATELEIRA = '" & Replace(Grid.TextMatrix(i, 9), "'", "''") & "'"
   If optMostrarFiscal.Value Then
      sSQL = sSQL & ", ESTOQUE_FISCAL = " & Replace(CStr(Val(Replace(Replace(Grid.TextMatrix(i, 10), ".", ""), ",", "."))), ",", ".") & ""
   Else
      sSQL = sSQL & ", quant_estoque = " & Replace(CStr(Val(Replace(Replace(Grid.TextMatrix(i, 10), ".", ""), ",", "."))), ",", ".") & ""
   End If
   sSQL = sSQL & " WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
   dbData.Execute sSQL
   If Len(Trim(Grid.TextMatrix(i, 8))) > 0 Then
      InserirTagSeNova Grid.TextMatrix(i, 7), Trim(Grid.TextMatrix(i, 8))
   End If
Next

picAguarde.Visible = False
cmdLocalizar_Click
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
Private Sub cmdAtualizarPreco_Click()
Me.Hide
'Load Produtos_AjustoPreco
Produtos_AjustoPreco.Show
Produtos_AjustoPreco.txtCodProduto.Text = (Grid.TextMatrix(Grid.Row, 2))
End Sub

Private Sub cmdAtualizarQuant_Click()
Me.Hide
Dim i As Integer
i = Grid.Row

If Grid.TextMatrix(i, 2) = "" Then Exit Sub

Produtos_AdicionarQuant.Show
Produtos_AdicionarQuant.txtCodProduto.Text = (Grid.TextMatrix(i, 2))
Produtos_AdicionarQuant.txtCodUsuario.Text = lblCodUsuario.Caption
Produtos_AdicionarQuant.txtQuantNova.SetFocus
End Sub


Private Sub cmdLocalizar_Click()
Dim varTipoMostrar As String


'criado pela IA
If optMostrarQuant.Value = True Then
    varTipoMostrar = " AND p.quant_estoque > 0" ' Mudei de produtos para p
ElseIf optMostrarNegativos.Value = True Then
    varTipoMostrar = " AND p.quant_estoque < 0"
ElseIf optMostrarZerados.Value = True Then
    varTipoMostrar = " AND p.quant_estoque = 0"
ElseIf optMostrarFiscal.Value = True Then
    varTipoMostrar = " AND p.ESTOQUE_FISCAL > 0"
ElseIf optMostrarTodos.Value = True Then
    varTipoMostrar = " "
End If

'meu código
'If optMostrarQuant.Value = True Then
'    varTipoMostrar = " AND produtos.quant_estoque > 0"
'ElseIf optMostrarNegativos.Value = True Then
'    varTipoMostrar = " AND produtos.quant_estoque < 0"
'ElseIf optMostrarZerados.Value = True Then
'    varTipoMostrar = " AND produtos.quant_estoque = 0"
'ElseIf optMostrarFiscal.Value = True Then
'    varTipoMostrar = " AND produtos.ESTOQUE_FISCAL > 0"
'ElseIf optMostrarTodos.Value = True Then
'    varTipoMostrar = " "
'End If

'criado pela IA
If optORDDesc.Value = True Then
   var_Indice = "p.descricao"
ElseIf optORDQuant.Value = True Then
   var_Indice = "p.quant_estoque"
ElseIf ORDQuantFiscal.Value = True Then
   var_Indice = "p.ESTOQUE_FISCAL"
ElseIf optORDValor.Value = True Then
   var_Indice = "precos.VALOR_VV" ' <--- Aqui a mágica do maior para o menor
ElseIf optORDValorCusto.Value = True Then
   var_Indice = "precos.CUSTO" ' Ordena pelo custo do maior para o menor
ElseIf optORDLinha.Value = True Then
   var_Indice = "p.categoria"
ElseIf optORDTFiscal.Value = True Then
   ' Realiza o cálculo (Custo * Estoque Fiscal) para ordenar
   var_Indice = "(precos.CUSTO * p.ESTOQUE_FISCAL)"
End If

If optORDASC.Value = True Then
   var_Direcao = " ASC"
Else
   var_Direcao = " DESC"
End If

'meu código
'If optORDDesc.Value = True Then
'   var_Indice = "produtos.descricao"
'ElseIf optORDQuant.Value = True Then
'   var_Indice = "produtos.quant_estoque"
'ElseIf ORDQuantFiscal.Value = True Then
'   var_Indice = "produtos.ESTOQUE_FISCAL"
'ElseIf optORDValor.Value = True Then
'    var_Indice = "(SELECT TOP 1 VALOR_VV FROM Produtos_Precos Where COD_PRODUTO = produtos.codigo order by CODIGO desc) DESC"
'ElseIf optORDLinha.Value = True Then
'   var_Indice = "produtos.categoria"
'End If


If optTodos.Value = True Then
    sSQL = "SELECT p.NCM AS var_NCM, p.CFOP AS var_CFOP, p.ICMSCST AS var_ICMS, " & _
           "p.categoria AS var_cat, p.TAGS AS var_tags, p.fabricante AS var_fab, p.PRATELEIRA AS var_Local, " & _
           "p.codigo AS var_cod, p.cod_barra AS var_codbarra, p.descricao AS var_desc, " & _
           "p.quant_estoque AS var_quant, p.ESTOQUE_FISCAL AS var_EstoqueFiscal, p.UNID_MEDIDA AS var_UnidMed, " & _
           "precos.CUSTO AS custo, precos.VALOR_VV AS venda " & _
           "FROM produtos p " & _
           "LEFT JOIN (" & _
           "   SELECT COD_PRODUTO, CUSTO, VALOR_VV, " & _
           "   ROW_NUMBER() OVER (PARTITION BY COD_PRODUTO ORDER BY CODIGO DESC) as RN " & _
           "   FROM Produtos_Precos" & _
           ") precos ON p.codigo = precos.COD_PRODUTO AND precos.RN = 1 " & _
           "WHERE (p.ativo = 1) " & varTipoMostrar & _
           " ORDER BY " & var_Indice & var_Direcao


    'meu código
    'sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.TAGS AS var_tags, produtos.fabricante AS var_fab, produtos.PRATELEIRA AS var_Local, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, produtos.ESTOQUE_FISCAL AS var_EstoqueFiscal, produtos.UNID_MEDIDA AS var_UnidMed, " & _
      "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS custo, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & _
      "WHERE (produtos.ativo = 1) " & varTipoMostrar & " ORDER BY " & var_Indice
   Set r = dbData.OpenRecordset(sSQL)
   
    If r.RecordCount > 32000 Then
        MsgBox "A Consulta retornou um valor maior de registros que é permitido na grade!", vbInformation, "Aviso do sistema"
        LimparGrid2
        Exit Sub
    Else
        If optMostrarFiscal.Value = True Then
            Formatar_Grid_Fiscal r
        Else
            Formatar_Grid r
        End If
        
    End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
Else
   MostrarCriterios
End If
'Debug.Print sSQL
If optCodBarra.Value = True Then txtCodBarra_GotFocus
If optNCM.Value = True Then txtCodBarra_GotFocus
AvaliarFrmEdicao
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSenha_Click()
sSQL = "SELECT * FROM usuario WHERE (password = '" & txtSenha.Text & "');"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblCodUsuario.Caption = ValidateNull(r("codigo"))
    lblUsuario.Caption = ValidateNull(r("login"))
    
    If lblCodUsuario.Caption = "" Then Exit Sub
    sSQL = "SELECT Usuario_permissoes.Codigo, Usuario_permissoes.permissao " & _
           "FROM Usuario_permissoes INNER JOIN Usuario_Acessos ON Usuario_permissoes.Codigo = Usuario_Acessos.Cod_Permissao " & _
           "WHERE (Usuario_permissoes.permissao = 'AJUSTE') AND (Usuario_Acessos.Cod_Usuario = " & lblCodUsuario.Caption & ")"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then
        cmdAtualizarPreco.Enabled = True
        cmdAtualizarQuant.Enabled = True
        txtSenha.Text = ""
        frmSenha.Visible = False
    Else
        cmdAtualizarPreco.Enabled = False
        cmdAtualizarQuant.Enabled = False
        ShowMsg "ACESSO NEGADO!" & vbCrLf & "Você não tem nivel de acesso a esse recurso", vbInformation
        lblCodUsuario.Caption = ""
        lblUsuario.Caption = ""
    End If
Else
    ShowMsg "ACESSO NEGADO!" & vbCrLf & "Senha Inválida!", vbInformation
    lblCodUsuario.Caption = ""
    lblUsuario.Caption = ""
End If
End Sub

Private Sub Form_Activate()
cmdLocalizar_Click
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
tipoEmpresa = CLng(sysConfig("TIPO_EMPRESA").Value)

'tipo de venda = 1 simples e 2 multiplus preços
Set cCfg = sysConfig("TIPOVALORVENDA")
varTipoValorVenda = cCfg.Value
Set cCfg = Nothing
End Sub

Private Sub cboDesc_Change()
   Dim p As Integer
   p = cboDesc.SelStart
   cboDesc.Text = UCase(cboDesc.Text)
   cboDesc.SelStart = p
End Sub

Private Sub cboDesc_GotFocus()
   cboDesc.Clear
   If optCategoria.Value Then
      sSQL = "SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;"
      Set r = dbData.OpenRecordset(sSQL)
      Do While Not r.EOF
         cboDesc.AddItem ValidateNull(r("Categoria"))
         r.MoveNext
      Loop
      If r.State <> 0 Then r.Close
      Set r = Nothing
   ElseIf optTags.Value Then
      sSQL = "SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Tipo_Empresa = " & tipoEmpresa & " ORDER BY c.Categoria, ct.Tags;"
      Set r = dbData.OpenRecordset(sSQL)
      Do While Not r.EOF
         cboDesc.AddItem ValidateNull(r("Tags"))
         r.MoveNext
      Loop
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
   moCombo.AttachTo cboDesc
End Sub

Private Sub cboDesc_LostFocus()
   'cboDesc_Click
End Sub

Private Sub Formatar_Grid_Fiscal(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim VarTotalGrid As Currency
   
    LimparGrid
    picAguarde.Visible = True
    DoEvents

    VarTotalGrid = 0

'If varTipoValorVenda = 2 Then
   With Grid
      .Clear
      .Cols = 15
      .rows = 2
      .FixedRows = 1
      .FixedCols = 0
      
      .ColWidth(0) = 300
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1500
      .ColWidth(4) = 4200
      .ColWidth(5) = 1600
      .ColWidth(6) = 800
      .ColWidth(7) = 1750
      .ColWidth(8) = 1200
      .ColWidth(9) = 800
      .ColWidth(10) = 1000
      .ColWidth(11) = 1000
      .ColWidth(12) = 1000
      .ColWidth(13) = 1100
      .ColWidth(14) = 1100
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FABRICANTE"
      .TextMatrix(0, 6) = "MED."
      .TextMatrix(0, 7) = "CATEGORIA"
      .TextMatrix(0, 8) = "TAG"
      .TextMatrix(0, 9) = "LOCAL"
      .TextMatrix(0, 10) = "FISCAL"
      .TextMatrix(0, 11) = "ESTOQUE"
      .TextMatrix(0, 12) = "VENDA"
      .TextMatrix(0, 13) = "CUSTO"
      .TextMatrix(0, 14) = "T.FISCAL"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'ALINHAMENTO
            .ColAlignment(2) = 1
            VarTotalGrid = 0
            '.TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("var_cod"))
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("var_desc"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_UnidMed"))
            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_cat"))
            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_tags"))
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_Local"))
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_EstoqueFiscal"))
            .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela("var_quant"))
            .TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
            .TextMatrix(.rows - 1, 13) = Format$(ValidateNull(rTabela("custo")), ocMONEY)
            
            VarTotalGrid = .TextMatrix(.rows - 1, 13) * .TextMatrix(.rows - 1, 10)
            .TextMatrix(.rows - 1, 14) = Format(VarTotalGrid, ocMONEY)
            '.TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("venda")), ocMONEY)

            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      lblValorTotalFiscal.Caption = Format(SomaGrid(Grid, 13), ocMONEY)
      .rows = .rows - 1
      Dim lRow As Long
      .FillStyle = 1
      For lRow = 1 To .rows - 1
         .Row = lRow
         .Col = 0
         .ColSel = .Cols - 1
         If lRow Mod 2 = 0 Then
            .CellBackColor = &HE0E0E0
         Else
            .CellBackColor = vbWhite
         End If
      Next lRow
      .FillStyle = 0
      .Redraw = True
      picAguarde.Visible = False
   End With
   Dim lChk As Long
   For lChk = 1 To Grid.rows - 1
      Grid.Row = lChk: Grid.Col = 0
      Set Grid.CellPicture = imgDesmarcada.Picture
      Grid.CellPictureAlignment = 4
   Next lChk
   ImgMarcadaTODAS.Visible = False
   imgDesmarcadaTODAS.Visible = True
   lblMarcarTodas.Caption = "Marcar todos"
   AvaliarFrmEdicao
'Else
'End If
End Sub
Private Sub Formatar_Grid(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim VarTotalGrid As Currency
   
    LimparGrid
    picAguarde.Visible = True
    DoEvents

    VarTotalGrid = 0

'If varTipoValorVenda = 2 Then
   With Grid
      .Clear
      .Cols = 12
      .rows = 2
      .FixedRows = 1
      .FixedCols = 0
      
      .ColWidth(0) = 300
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1500
      .ColWidth(4) = 5200
      .ColWidth(5) = 1600
      .ColWidth(6) = 800
      .ColWidth(7) = 1750
      .ColWidth(8) = 1200
      .ColWidth(9) = 800
      .ColWidth(10) = 1000
      .ColWidth(11) = 1000
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FABRICANTE"
      .TextMatrix(0, 6) = "MED."
      .TextMatrix(0, 7) = "CATEGORIA"
      .TextMatrix(0, 8) = "TAG"
      .TextMatrix(0, 9) = "LOCAL"
      .TextMatrix(0, 10) = "ESTOQUE"
      .TextMatrix(0, 11) = "VENDA"

      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'ALINHAMENTO
            .ColAlignment(2) = 1
            VarTotalGrid = 0
            '.TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("var_cod"))
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("var_desc"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_UnidMed"))
            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_cat"))
            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_tags"))
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_Local"))
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_quant"))
            .TextMatrix(.rows - 1, 11) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
            '.TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("custo")), ocMONEY)
            
            'VarTotalGrid = .TextMatrix(.rows - 1, 12) * .TextMatrix(.rows - 1, 9)
            '.TextMatrix(.rows - 1, 13) = Format(VarTotalGrid, ocMONEY)
            ''.TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("venda")), ocMONEY)

            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      'lblValorTotalFiscal.Caption = Format(SomaGrid(Grid, 13), ocMONEY)
      .rows = .rows - 1
      Dim lRow As Long
      .FillStyle = 1
      For lRow = 1 To .rows - 1
         .Row = lRow
         .Col = 0
         .ColSel = .Cols - 1
         If lRow Mod 2 = 0 Then
            .CellBackColor = &HE0E0E0
         Else
            .CellBackColor = vbWhite
         End If
      Next lRow
      .FillStyle = 0
      .Redraw = True
      picAguarde.Visible = False
   End With
   Dim lChk As Long
   For lChk = 1 To Grid.rows - 1
      Grid.Row = lChk: Grid.Col = 0
      Set Grid.CellPicture = imgDesmarcada.Picture
      Grid.CellPictureAlignment = 4
   Next lChk
   ImgMarcadaTODAS.Visible = False
   imgDesmarcadaTODAS.Visible = True
   lblMarcarTodas.Caption = "Marcar todos"
   AvaliarFrmEdicao
'Else
'End If
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
Private Sub optEICMS_Click()
   lblEdicaoColetiva.Caption = "ICMS CST"
   txtEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
   cboEdicaoColetiva.Visible = False
End Sub

Private Sub optENCM_Click()
   lblEdicaoColetiva.Caption = "NCM"
   txtEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
   cboEdicaoColetiva.Visible = False
End Sub

Private Sub optECFOP_Click()
   lblEdicaoColetiva.Caption = "CFOP"
   txtEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
   cboEdicaoColetiva.Visible = False
End Sub

Private Sub optECategoria_Click()
Dim rE As ADODB.Recordset
   lblEdicaoColetiva.Caption = "Categoria"
   txtEdicaoColetiva.Visible = False
   cboEdicaoColetiva.Clear
   Set rE = dbData.OpenRecordset("SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;")
   Do While Not rE.EOF: cboEdicaoColetiva.AddItem ValidateNull(rE("Categoria")): rE.MoveNext: Loop
   rE.Close: Set rE = Nothing
   cboEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
End Sub

Private Sub optETags_Click()
Dim rE As ADODB.Recordset
   lblEdicaoColetiva.Caption = "Tags"
   txtEdicaoColetiva.Visible = False
   cboEdicaoColetiva.Clear
   Set rE = dbData.OpenRecordset("SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Tipo_Empresa = " & tipoEmpresa & " ORDER BY c.Categoria, ct.Tags;")
   Do While Not rE.EOF: cboEdicaoColetiva.AddItem ValidateNull(rE("Tags")): rE.MoveNext: Loop
   rE.Close: Set rE = Nothing
   cboEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
End Sub

Private Sub cmdEdicaoColetiva_Click()
Dim i As Integer
Dim sVal As String
Dim iColE As Integer
Dim sSet As String

If txtEdicaoColetiva.Visible Then
   sVal = txtEdicaoColetiva.Text
Else
   sVal = cboEdicaoColetiva.Text
End If
If optETags.Value Or optECategoria.Value Then sVal = UCase(sVal)

If optEICMS.Value Then
   iColE = -1: sSet = "ICMSCST = '" & sVal & "'"
ElseIf optENCM.Value Then
   iColE = -1: sSet = "NCM = '" & sVal & "'"
ElseIf optECFOP.Value Then
   iColE = -1: sSet = "CFOP = " & sVal
ElseIf optECategoria.Value Then
   iColE = 7: sSet = "categoria = '" & sVal & "'"
ElseIf optETags.Value Then
   iColE = 8: sSet = "TAGS = '" & sVal & "'"
Else
   Exit Sub
End If

If Len(Trim(sSet)) = 0 Then Exit Sub

Dim iExact As Integer
iExact = 0
If optENCM.Value Then iExact = 8
If optECFOP.Value Then iExact = 4
If optEICMS.Value Then iExact = 3
If iExact > 0 Then
   Dim k As Integer
   Dim bNumOk As Boolean
   bNumOk = (Len(sVal) = iExact)
   If bNumOk Then
      For k = 1 To Len(sVal)
         If Mid(sVal, k, 1) < "0" Or Mid(sVal, k, 1) > "9" Then bNumOk = False: Exit For
      Next k
   End If
   If Not bNumOk Then
      MsgBox "Digite exatamente " & iExact & " dígitos numéricos.", vbExclamation, "Edição Coletiva"
      txtEdicaoColetiva.SetFocus
      Exit Sub
   End If
End If

Dim bTemMarcadas As Boolean
For i = 1 To Grid.rows - 1
   If Grid.TextMatrix(i, 0) = "1" Then bTemMarcadas = True: Exit For
Next i

If optETags.Value Then
   Dim iLinSemCat As Integer
   For i = 1 To Grid.rows - 1
      If bTemMarcadas And Grid.TextMatrix(i, 0) <> "1" Then GoTo ProxValCat
      If Len(Trim(Grid.TextMatrix(i, 7))) = 0 Then iLinSemCat = iLinSemCat + 1
ProxValCat:
   Next i
   If iLinSemCat > 0 Then
      MsgBox iLinSemCat & " produto(s) sem categoria. Defina a categoria antes de aplicar a tag.", vbExclamation, "Tag sem categoria"
      Exit Sub
   End If
End If

picAguarde.Visible = True
DoEvents

For i = 1 To Grid.rows - 1
   If bTemMarcadas And Grid.TextMatrix(i, 0) <> "1" Then GoTo ProximaLinha
   If iColE >= 0 Then Grid.TextMatrix(i, iColE) = sVal
   dbData.Execute "UPDATE produtos SET " & sSet & " WHERE codigo = " & Grid.TextMatrix(i, 2) & ";"
   If optETags.Value Then InserirTagSeNova Grid.TextMatrix(i, 7), sVal
ProximaLinha:
Next i

txtEdicaoColetiva.Text = ""
cboEdicaoColetiva.Text = ""
picAguarde.Visible = False
ResetarMarcas
optENCM.Value = False
optECFOP.Value = False
optEICMS.Value = False
optECategoria.Value = False
optETags.Value = False
txtEdicaoColetiva.Visible = False
cboEdicaoColetiva.Visible = False
cmdEdicaoColetiva.Visible = False
lblEdicaoColetiva.Visible = False
cmdAtualizar.Visible = True
End Sub

Private Sub txtEdicaoColetiva_LostFocus()
Dim s As String
s = Trim(txtEdicaoColetiva.Text)
If optENCM.Value Then
   s = Replace(s, ".", "")
   s = Replace(s, " ", "")
End If
txtEdicaoColetiva.Text = s
End Sub

Private Sub cboEdicaoColetiva_KeyPress(KeyAscii As Integer)
   If Not (optETags.Value Or optECategoria.Value) Then Exit Sub
   If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
   ElseIf optETags.Value Then
      If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32) Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub cboEdicaoColetiva_Change()
   If Not (optETags.Value Or optECategoria.Value) Then Exit Sub
   Dim pos As Integer
   pos = cboEdicaoColetiva.SelStart
   cboEdicaoColetiva.Text = UCase(cboEdicaoColetiva.Text)
   cboEdicaoColetiva.SelStart = pos
End Sub

Private Sub cboEdicaoColetiva_LostFocus()
   cboEdicaoColetiva.Text = Trim(cboEdicaoColetiva.Text)
End Sub

Private Sub ResetarMarcas()
Dim i As Integer
For i = 1 To Grid.rows - 1
   Grid.TextMatrix(i, 0) = ""
   Grid.Row = i: Grid.Col = 0
   Set Grid.CellPicture = imgDesmarcada.Picture
   Grid.CellPictureAlignment = 4
Next i
ImgMarcadaTODAS.Visible = False
imgDesmarcadaTODAS.Visible = True
lblMarcarTodas.Caption = "Marcar todos"
AvaliarFrmEdicao
End Sub

Private Sub AvaliarFrmEdicao()
Dim j As Integer
Dim bEnabled As Boolean
For j = 1 To Grid.rows - 1
   If Grid.TextMatrix(j, 0) = "1" Then bEnabled = True: Exit For
Next j
If Not bEnabled Then
   If Not (optTodos.Value Or optCodBarra.Value) Then
      If Grid.rows > 1 Then bEnabled = True
   End If
End If
frmAlterarGrupos.Enabled = bEnabled
frmEdicao.Enabled = bEnabled
frmEdicaoFiltros.Enabled = bEnabled
End Sub

Private Sub imgDesmarcadaTODAS_Click()
Dim i As Integer
imgDesmarcadaTODAS.Visible = False
ImgMarcadaTODAS.Visible = True
lblMarcarTodas.Caption = "Desmarcar todos"
For i = 1 To Grid.rows - 1
   Grid.TextMatrix(i, 0) = "1"
   Grid.Row = i: Grid.Col = 0
   Set Grid.CellPicture = ImgMarcada.Picture
   Grid.CellPictureAlignment = 4
Next i
AvaliarFrmEdicao
End Sub

Private Sub ImgMarcadaTODAS_Click()
Dim i As Integer
ImgMarcadaTODAS.Visible = False
imgDesmarcadaTODAS.Visible = True
lblMarcarTodas.Caption = "Marcar todos"
For i = 1 To Grid.rows - 1
   Grid.TextMatrix(i, 0) = ""
   Grid.Row = i: Grid.Col = 0
   Set Grid.CellPicture = imgDesmarcada.Picture
   Grid.CellPictureAlignment = 4
Next i
AvaliarFrmEdicao
End Sub

Private Sub lblMarcarTodas_Click()
Dim i As Integer
If ImgMarcadaTODAS.Visible Then
   ImgMarcadaTODAS.Visible = False
   imgDesmarcadaTODAS.Visible = True
   lblMarcarTodas.Caption = "Marcar todos"
   For i = 1 To Grid.rows - 1
      Grid.TextMatrix(i, 0) = ""
      Grid.Row = i: Grid.Col = 0
      Set Grid.CellPicture = imgDesmarcada.Picture
      Grid.CellPictureAlignment = 4
   Next i
Else
   ImgMarcadaTODAS.Visible = True
   imgDesmarcadaTODAS.Visible = False
   lblMarcarTodas.Caption = "Desmarcar todos"
   For i = 1 To Grid.rows - 1
      Grid.TextMatrix(i, 0) = "1"
      Grid.Row = i: Grid.Col = 0
      Set Grid.CellPicture = ImgMarcada.Picture
      Grid.CellPictureAlignment = 4
   Next i
End If
AvaliarFrmEdicao
End Sub

Private Sub InserirTagSeNova(sCat As String, sTag As String)
Dim sTagU As String
Dim nID As Long
Dim nExiste As Long
sTagU = UCase(Trim(sTag))
If Len(sTagU) = 0 Or Len(Trim(sCat)) = 0 Then Exit Sub
nID = Val(SQLExecutaRetorno("SELECT ID_Categoria FROM Categorias WHERE Categoria = '" & Replace(sCat, "'", "''") & "' AND Tipo_Empresa = " & tipoEmpresa, "ID_Categoria"))
If nID = 0 Then Exit Sub
nExiste = Val(SQLExecutaRetorno("SELECT COUNT(*) AS n FROM Categorias_Tags WHERE ID_Categoria = " & nID & " AND Tags = '" & Replace(sTagU, "'", "''") & "'", "n"))
If nExiste = 0 Then
   dbData.Execute "INSERT INTO Categorias_Tags (ID_Categoria, Tags) VALUES (" & nID & ", '" & Replace(sTagU, "'", "''") & "');"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_Click()
Dim ColLimite As Integer
Dim rCbo As ADODB.Recordset

If optMostrarFiscal.Value = True Then
    ColLimite = 10
Else
    ColLimite = 9
End If

iRow = Grid.Row
iCol = Grid.Col

If iCol = 0 Then
   If iRow < 1 Then Exit Sub
   If Grid.TextMatrix(iRow, 0) = "1" Then
      Grid.TextMatrix(iRow, 0) = ""
      Grid.Row = iRow: Grid.Col = 0
      Set Grid.CellPicture = imgDesmarcada.Picture
      Grid.CellPictureAlignment = 4
   Else
      Grid.TextMatrix(iRow, 0) = "1"
      Grid.Row = iRow: Grid.Col = 0
      Set Grid.CellPicture = ImgMarcada.Picture
      Grid.CellPictureAlignment = 4
   End If
   AvaliarFrmEdicao
   Exit Sub
End If
If iCol < 3 Or iCol > ColLimite Then Exit Sub

Select Case iCol
   Case 7
      cboEdit.Clear
      Set rCbo = dbData.OpenRecordset("SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;")
      Do While Not rCbo.EOF
         cboEdit.AddItem ValidateNull(rCbo("Categoria"))
         rCbo.MoveNext
      Loop
      If rCbo.State <> 0 Then rCbo.Close
      Set rCbo = Nothing
      cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth
      cboEdit.Text = Grid.TextMatrix(iRow, iCol)
      cboEdit.Visible = True
      cboEdit.SetFocus
   Case 8
      If Len(Trim(Grid.TextMatrix(iRow, 7))) = 0 Then
         MsgBox "Defina a categoria do produto antes de editar a tag.", vbExclamation, "Tag sem categoria"
         Exit Sub
      End If
      cboEdit.Clear
      Set rCbo = dbData.OpenRecordset("SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Tipo_Empresa = " & tipoEmpresa & " ORDER BY c.Categoria, ct.Tags;")
      Do While Not rCbo.EOF
         cboEdit.AddItem ValidateNull(rCbo("Tags"))
         rCbo.MoveNext
      Loop
      If rCbo.State <> 0 Then rCbo.Close
      Set rCbo = Nothing
      cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth
      cboEdit.Text = Grid.TextMatrix(iRow, iCol)
      cboEdit.Visible = True
      cboEdit.SetFocus
   Case 6
      cboEdit.Clear
      With cboEdit
         .AddItem "UN": .AddItem "CX": .AddItem "M":   .AddItem "M2"
         .AddItem "M3": .AddItem "ML": .AddItem "KG":  .AddItem "GR"
         .AddItem "CT": .AddItem "PO": .AddItem "SC":  .AddItem "PA"
         .AddItem "EX": .AddItem "BJ": .AddItem "DZ":  .AddItem "PC"
         .AddItem "DI": .AddItem "FD": .AddItem "PT"
      End With
      cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth
      cboEdit.Text = Grid.TextMatrix(iRow, iCol)
      cboEdit.Visible = True
      cboEdit.SetFocus
   Case Else
      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
      txtEdit.Text = Grid.TextMatrix(iRow, iCol)
      txtEdit.Visible = True
      txtEdit.SetFocus
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(txtEdit.Text)
End Select
End Sub

Private Sub optCategoria_Click()
   lblCategoria.Visible = False
   lblDesc.Visible = True
   cboDesc.Visible = True
   optPorPalavra.Visible = False
   PorPalavraDupla.Visible = False
   lblCodBarra.Caption = "Categoria"
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = True
   cboDesc.SetFocus
End Sub

Private Sub optCodBarra_Click()
   lblCategoria.Visible = False
   lblDesc.Visible = False
   cboDesc.Visible = False
   cboDesc.Visible = False
   optPorPalavra.Visible = False
   PorPalavraDupla.Visible = False
   lblCodBarra.Caption = "Cód. Barra"
   lblCodBarra.Visible = True
   txtCodBarra.Visible = True
   cmdLocalizar.Visible = True
   txtCodBarra.SetFocus
End Sub

Private Sub optDesc_Click()
   lblCategoria.Visible = False
   lblDesc.Visible = True
   cboDesc.Visible = True
   optPorPalavra.Visible = True
   PorPalavraDupla.Visible = True
   lblCodBarra.Caption = "Descrição"
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = True
   cboDesc.SetFocus
End Sub

Private Sub optTags_Click()
   lblCategoria.Visible = False
   lblDesc.Visible = True
   cboDesc.Visible = True
   optPorPalavra.Visible = False
   PorPalavraDupla.Visible = False
   lblCodBarra.Caption = "Tags"
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = True
   cboDesc.SetFocus
End Sub

Private Sub optNCM_Click()
   lblCategoria.Visible = False
   lblDesc.Visible = False
   cboDesc.Visible = False
   optPorPalavra.Visible = False
   PorPalavraDupla.Visible = False
   lblCodBarra.Caption = "NCM"
   lblCodBarra.Visible = True
   txtCodBarra.Visible = True
   cmdLocalizar.Visible = True
   txtCodBarra.SetFocus
End Sub

Private Sub optMostrarFiscal_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = True
optORDValorCusto.Visible = True
ORDQuantFiscal.Visible = True
optORDTFiscal.Visible = True
End Sub

Private Sub optMostrarNegativos_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = False
optORDValorCusto.Visible = False
ORDQuantFiscal.Visible = False
optORDTFiscal.Visible = False
End Sub

Private Sub optMostrarQuant_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = False
optORDValorCusto.Visible = False
ORDQuantFiscal.Visible = False
optORDTFiscal.Visible = False
End Sub

Private Sub optMostrarTodos_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = False
optORDValorCusto.Visible = False
ORDQuantFiscal.Visible = False
optORDTFiscal.Visible = False
End Sub

Private Sub optMostrarZerados_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = False
optORDValorCusto.Visible = False
ORDQuantFiscal.Visible = False
optORDTFiscal.Visible = False
End Sub

Private Sub optTodos_Click()
   lblCategoria.Visible = False
   lblDesc.Visible = False
   cboDesc.Visible = False
   cboDesc.Visible = False
   optPorPalavra.Visible = False
   PorPalavraDupla.Visible = False
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = False
   cmdLocalizar_Click
End Sub

Private Sub txtCodBarra_Change()
   If Len(txtCodBarra.Text) = 13 Then cmdLocalizar_Click
End Sub

Private Sub txtCodBarra_GotFocus()
   SelectControl txtCodBarra
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
   Select Case iCol
   Case 3
      If KeyAscii <> 8 Then
         If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
      End If
   Case 4, 5
      If KeyAscii = 8 Then Exit Sub
      If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32: Exit Sub
      If (KeyAscii >= 65 And KeyAscii <= 90) Then Exit Sub
      If KeyAscii = 32 Then Exit Sub
      If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
      If KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 47 Then Exit Sub
      KeyAscii = 0
   End Select
End Sub

Private Sub txtEdit_Change()
   If iCol = 4 Or iCol = 5 Then
      Dim p As Integer
      p = txtEdit.SelStart
      txtEdit.Text = UCase(txtEdit.Text)
      txtEdit.SelStart = p
   End If
End Sub

Private Sub txtEdit_GotFocus()
'criado por IA
' Registra as coordenadas ONDE o editor nasceu.
' O LostFocus usará essas variáveis para saber onde salvar o valor.
iRow = Grid.Row
iCol = Grid.Col
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
'criado pela IA
' Captura a linha e coluna ATUAIS onde o txtEdit está posicionado
  Dim r As Long, c As Long
  r = Grid.Row
  c = Grid.Col

  If KeyCode = 38 Then ' Seta para CIMA
     If r > 1 Then ' Evita subir além do cabeçalho
        ' 1. Salva o valor atual do texto na célula antes de sair dela
        Grid.TextMatrix(r, c) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
        
        ' 2. Move para a linha de cima e reativa o editor
        Grid.Row = r - 1
        Grid_Click
     Else
        MsgBox "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation
     End If
  
  ElseIf KeyCode = 40 Then ' Seta para BAIXO
     If r < Grid.rows - 1 Then
        ' 1. Salva o valor atual do texto na célula
        Grid.TextMatrix(r, c) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
        
        ' 2. Move para a linha de baixo e reativa o editor
        Grid.Row = r + 1
        Grid_Click
     Else
        MsgBox "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation
     End If
  End If
'criado por mim
 'If KeyCode = 38 Then
 '  If Grid.Row - 1 = 0 Then ShowMsg "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation: Exit Sub
 '  Grid.Row = iRow - 1
 '  Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
 '  Grid_Click

'ElseIf KeyCode = 40 Then
'   If Grid.rows = Grid.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
'   Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
'   Grid.Row = iRow + 1
'   Grid_Click
'End If
End Sub

Private Sub txtEdit_LostFocus()
If iCol = 3 Then
   txtEdit.Text = Replace(txtEdit.Text, " ", "")
ElseIf iCol = 4 Or iCol = 5 Then
   txtEdit.Text = Trim(txtEdit.Text)
End If

If iCol >= 3 And iCol <= 5 Then
   Grid.TextMatrix(iRow, iCol) = txtEdit.Text
Else
   Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
End If
txtEdit.Visible = False
End Sub

Private Sub cboEdit_LostFocus()
   If iRow > 0 Then
      Dim sVal As String
      sVal = Trim(cboEdit.Text)
      If iCol = 8 Then sVal = UCase(sVal)
      If iCol = 6 Then
         Dim sMed As String
         sMed = ",UN,CX,M,M2,M3,ML,KG,GR,CT,PO,SC,PA,EX,BJ,DZ,PC,DI,FD,PT,"
         If InStr(sMed, "," & UCase(sVal) & ",") = 0 Then sVal = Grid.TextMatrix(iRow, iCol)
      End If
      Grid.TextMatrix(iRow, iCol) = sVal
   End If
   cboEdit.Visible = False
End Sub

Private Sub cboEdit_KeyPress(KeyAscii As Integer)
   If iCol = 6 Then
      KeyAscii = 0
   ElseIf iCol = 8 Then
      If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   End If
End Sub

Private Sub cboEdit_Change()
   If iCol = 7 Or iCol = 8 Then
      Dim pos As Integer
      pos = cboEdit.SelStart
      cboEdit.Text = UCase(cboEdit.Text)
      cboEdit.SelStart = pos
   End If
End Sub



Private Sub txtSenha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSenha_Click
End Sub


