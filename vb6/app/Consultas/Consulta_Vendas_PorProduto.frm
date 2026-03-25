VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Consulta_Vendas_PorProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE VENDAS"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "Consulta_Vendas_PorProduto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10995
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
      Height          =   255
      Left            =   60
      TabIndex        =   53
      Top             =   9300
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Exibir pedidos deste produto"
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
      MICON           =   "Consulta_Vendas_PorProduto.frx":23D2
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
      Height          =   3435
      Left            =   60
      ScaleHeight     =   3375
      ScaleWidth      =   9795
      TabIndex        =   20
      ToolTipText     =   "Imprimir"
      Top             =   1080
      Width           =   9855
      Begin VB.ComboBox cboTipoPgtoSec 
         Height          =   315
         Left            =   2760
         TabIndex        =   7
         Top             =   3000
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.ComboBox CboFormaPgtoSec 
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.ComboBox cboTipoPgto 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   2595
      End
      Begin VB.Frame Frame8 
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
         Height          =   2715
         Left            =   5400
         TabIndex        =   21
         Top             =   60
         Width           =   4335
         Begin VB.Frame frmFiltro1 
            Height          =   2340
            Left            =   60
            TabIndex        =   22
            Top             =   180
            Width           =   4155
            Begin VB.ComboBox cboCliente 
               Height          =   315
               Left            =   180
               TabIndex        =   32
               Top             =   420
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.ComboBox cboVendedor 
               Height          =   315
               Left            =   180
               TabIndex        =   31
               Top             =   420
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.TextBox txtCodigo 
               Height          =   315
               Left            =   180
               TabIndex        =   30
               Top             =   420
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   1560
               Sorted          =   -1  'True
               TabIndex        =   29
               Top             =   1740
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMes 
               Height          =   315
               Left            =   180
               TabIndex        =   28
               Top             =   1740
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtCodCliente 
               Appearance      =   0  'Flat
               Height          =   195
               Left            =   3420
               TabIndex        =   27
               Top             =   180
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtCodFunc 
               Appearance      =   0  'Flat
               Height          =   195
               Left            =   2760
               TabIndex        =   26
               Top             =   180
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.ComboBox cboCategoria 
               Height          =   315
               Left            =   180
               TabIndex        =   25
               Top             =   1080
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.ComboBox cboDescricao 
               Height          =   315
               Left            =   180
               TabIndex        =   24
               Top             =   420
               Visible         =   0   'False
               Width           =   3855
            End
            Begin VB.TextBox txtCodBarra 
               Height          =   315
               Left            =   180
               TabIndex        =   23
               Top             =   420
               Visible         =   0   'False
               Width           =   2355
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   315
               Left            =   180
               TabIndex        =   33
               Top             =   420
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFim 
               Height          =   315
               Left            =   2040
               TabIndex        =   34
               Top             =   420
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdCalendario1 
               Height          =   315
               Left            =   1200
               TabIndex        =   57
               Top             =   420
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
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
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Consulta_Vendas_PorProduto.frx":23EE
               PICN            =   "Consulta_Vendas_PorProduto.frx":240A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdCalendario2 
               Height          =   315
               Left            =   3060
               TabIndex        =   58
               Top             =   420
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
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
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Consulta_Vendas_PorProduto.frx":47EC
               PICN            =   "Consulta_Vendas_PorProduto.frx":4808
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblClientes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes:"
               Height          =   195
               Left            =   180
               TabIndex        =   45
               Top             =   180
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblVendedor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vendedor(a):"
               Height          =   195
               Left            =   180
               TabIndex        =   44
               Top             =   180
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblAte 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "até"
               Height          =   195
               Left            =   1680
               TabIndex        =   43
               Top             =   480
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label lblFim 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data final:"
               Height          =   195
               Left            =   2040
               TabIndex        =   42
               Top             =   180
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label lblInicio 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data inicial:"
               Height          =   195
               Left            =   180
               TabIndex        =   41
               Top             =   180
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label lblCodigo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Código:"
               Height          =   195
               Left            =   180
               TabIndex        =   40
               Top             =   180
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label lblAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   1560
               TabIndex        =   39
               Top             =   1500
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lblMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   180
               TabIndex        =   38
               Top             =   1500
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblCategoria 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Categoria:"
               Height          =   195
               Left            =   180
               TabIndex        =   37
               Top             =   840
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label lblDescricao 
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   180
               TabIndex        =   36
               Top             =   180
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblCodBarra 
               Caption         =   "Cod. Barra:"
               Height          =   195
               Left            =   180
               TabIndex        =   35
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
         End
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2595
      End
      Begin VB.ComboBox cboCriterioPrinc 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   1020
         Width           =   2595
      End
      Begin VB.ComboBox cboCriterioSec 
         Height          =   315
         Left            =   2820
         TabIndex        =   2
         Top             =   1020
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.ComboBox cboIndice 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1620
         Width           =   2595
      End
      Begin VB.ComboBox CboFormaPgto 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   2595
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   495
         Left            =   5460
         TabIndex        =   8
         Top             =   2820
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
         MICON           =   "Consulta_Vendas_PorProduto.frx":6BEA
         PICN            =   "Consulta_Vendas_PorProduto.frx":6C06
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
         Height          =   495
         Left            =   8400
         TabIndex        =   46
         Top             =   2820
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         MICON           =   "Consulta_Vendas_PorProduto.frx":74E0
         PICN            =   "Consulta_Vendas_PorProduto.frx":74FC
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
         Height          =   495
         Left            =   6960
         TabIndex        =   47
         Top             =   2820
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Imprimir"
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
         MICON           =   "Consulta_Vendas_PorProduto.frx":7816
         PICN            =   "Consulta_Vendas_PorProduto.frx":7832
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pgto"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Criterio"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Organizar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1380
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pgto"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   1035
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3600
      Picture         =   "Consulta_Vendas_PorProduto.frx":7B4C
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   12
      Top             =   6660
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   9825
      TabIndex        =   9
      Top             =   60
      Width           =   9855
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE VENDAS - POR PRODUTOS"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   6405
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Consulta_Vendas_PorProduto.frx":8B84
         Top             =   0
         Width           =   1140
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   19
      Top             =   10725
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13309
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "16:20"
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
      Height          =   4695
      Left            =   60
      TabIndex        =   54
      Top             =   4560
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8281
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirProdutos 
      Height          =   255
      Left            =   60
      TabIndex        =   55
      Top             =   9300
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Exibir produtos"
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
      MICON           =   "Consulta_Vendas_PorProduto.frx":F3CA
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
      Height          =   255
      Left            =   2760
      TabIndex        =   56
      Top             =   9300
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Exibir Parcelas"
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
      MICON           =   "Consulta_Vendas_PorProduto.frx":F3E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
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
      Left            =   7440
      TabIndex        =   18
      Top             =   10200
      Width           =   675
   End
   Begin VB.Label lblEntrada 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   8160
      TabIndex        =   17
      Top             =   9840
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENT.:"
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
      Left            =   7560
      TabIndex        =   16
      Top             =   9840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   8160
      TabIndex        =   15
      Top             =   10200
      Width           =   1635
   End
   Begin VB.Label lblQtda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   8160
      TabIndex        =   14
      Top             =   9480
      Width           =   1635
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANT.:"
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
      Left            =   7320
      TabIndex        =   13
      Top             =   9480
      Width           =   780
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1275
      Left            =   7200
      Top             =   9360
      Width           =   2715
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   9900
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Consulta_Vendas_PorProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

Dim posX As Single

Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer

Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Private Sub FormatarGrid_ProdutosLucros(rTabela As ADODB.Recordset)
   Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 6760
      .ColWidth(2) = 1000
      .ColWidth(3) = 800
      .ColWidth(4) = 1000
      
      .TextMatrix(0, 1) = "DESCRIÇÃO"
      .TextMatrix(0, 2) = "PREÇO"
      .TextMatrix(0, 3) = "QTDE"
      .TextMatrix(0, 4) = "TOTAL"
      
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      .ColAlignment(1) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("descricao")
            .TextMatrix(.Rows - 1, 2) = Format$(rTabela("preco"), ocMONEY)
            .TextMatrix(.Rows - 1, 3) = rTabela("var_qtde")
            .TextMatrix(.Rows - 1, 4) = Format$(rTabela("var_total"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Redraw = True
      .Rows = .Rows - 1
   End With
   
   lblQtda.Caption = Format(SomaGrid(Grid, 3), ocPESO)
   lblTotal.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub FormatarGrid_Produtos(rTabela As ADODB.Recordset)
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 6660
      .ColWidth(2) = 1000
      .ColWidth(3) = 900
      .ColWidth(4) = 1000
      
      .TextMatrix(0, 1) = "DESCRIÇÃO"
      .TextMatrix(0, 2) = "PREÇO"
      .TextMatrix(0, 3) = "QTDE"
      .TextMatrix(0, 4) = "TOTAL"
      
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      .ColAlignment(1) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
      
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 0) = rTabela("cod_produto")
            
            If tipoEmpresa = 4 Then
            .TextMatrix(.Rows - 1, 1) = rTabela("var_desc") & " /  " & rTabela("var_tam") & " / " & rTabela("var_fab") & " /  " & rTabela("ref")
            Else
            .TextMatrix(.Rows - 1, 1) = rTabela("var_desc")
            End If
            
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.Rows - 1, 3) = rTabela("var_qtde")
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("var_total"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblQtda.Caption = Format(SomaGrid(Grid, 3), ocPESO)
   lblTotal.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
   lblEntrada.Caption = Format(0, ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub Limpar_Grid_Venda()
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1220
      .ColWidth(7) = 0
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "TIPO"
      .TextMatrix(0, 7) = "TIPO"
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      i = 1
      
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
         .Col = 4
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      Grid.Redraw = True
   End With
   
   lblQtda.Caption = Format(0, ocMONEY)
   lblTotal.Caption = Format(0, ocMONEY)
   lblEntrada.Caption = Format(0, ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
cboCategoria.Text = ""
cboVendedor.Text = ""
txtCodigo.Text = ""
cboCliente.Text = ""
mskFim.Mask = ""
mskFim.Text = ""
mskInicio.Mask = ""
mskInicio.Text = ""
txtCodFunc.Text = ""
txtCodCliente.Text = ""
End Sub

Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAno.AddItem i
Next

moCombo.AttachTo cboAno
End Sub

Private Sub cboAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboCategoria_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboCategoria.Clear
   
   sSQL = "SELECT categoria FROM produtos GROUP BY categoria;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCategoria.AddItem ValidateNull(r("categoria"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboCategoria
End Sub

Private Sub cboCliente_Click()
   On Error GoTo TrataErro
   
   If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   If cboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
   txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub CboCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboCliente.Clear
   
   sSQL = "SELECT * FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCliente.AddItem r("nome")
      cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboCliente
End Sub

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub CboCliente_LostFocus()
   cboCliente_Click
End Sub

Private Sub cboCriterioPrinc_Change()
   LimparObjetos_Consulta
   
   If cboCriterioPrinc.Text = "TODOS" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboCategoria.Visible = False
      lblCategoria.Visible = False
      
      lblCodBarra.Visible = False
      txtCodBarra.Visible = False
      lblDescricao.Visible = False
      cboDescricao.Visible = False
      
      cboCriterioSec.Visible = False
      
   ElseIf cboCriterioPrinc.Text = "VENDEDOR" Then
      lblVendedor.Visible = True
      cboVendedor.Visible = True
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboCategoria.Visible = False
      lblCategoria.Visible = False
      
      lblCodBarra.Visible = False
      txtCodBarra.Visible = False
      lblDescricao.Visible = False
      cboDescricao.Visible = False
      
      cboCriterioSec.Visible = True

      cboVendedor.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "CLIENTE" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario2.Visible = False
      cmdCalendario1.Visible = False
      
      lblClientes.Visible = True
      cboCliente.Visible = True
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboCategoria.Visible = False
      lblCategoria.Visible = False
      
      lblCodBarra.Visible = False
      txtCodBarra.Visible = False
      lblDescricao.Visible = False
      cboDescricao.Visible = False
      
      cboCriterioSec.Visible = True
      
      cboCliente.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "PERIODO" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = True
      mskInicio.Visible = True
      lblFim.Visible = True
      mskFim.Visible = True
      lblAte.Visible = True
      cmdCalendario1.Visible = True
      cmdCalendario2.Visible = True
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboCategoria.Visible = False
      lblCategoria.Visible = False
      
      lblCodBarra.Visible = False
      txtCodBarra.Visible = False
      lblDescricao.Visible = False
      cboDescricao.Visible = False
      
      cboCriterioSec.Visible = False
      
      mskInicio.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "CÓDIGO" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = True
      txtCodigo.Visible = True
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboCategoria.Visible = False
      lblCategoria.Visible = False
      
      lblCodBarra.Visible = False
      txtCodBarra.Visible = False
      lblDescricao.Visible = False
      cboDescricao.Visible = False
      
      cboCriterioSec.Visible = False
      
      txtCodigo.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "MENSAL" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      cboCategoria.Visible = False
      lblCategoria.Visible = False
      
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
      
      lblCodBarra.Visible = False
      txtCodBarra.Visible = False
      lblDescricao.Visible = False
      cboDescricao.Visible = False
      
      cboCriterioSec.Visible = False
      
      cboMes.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "CATEGORIA" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboCategoria.Visible = True
      lblCategoria.Visible = True
      
      lblCodBarra.Visible = False
      txtCodBarra.Visible = False
      lblDescricao.Visible = False
      cboDescricao.Visible = False

      cboCriterioSec.Visible = True
      
   ElseIf cboCriterioPrinc.Text = "ESPECIFICO" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboCategoria.Visible = False
      lblCategoria.Visible = False

      cboCriterioSec.Visible = True

      cboCriterioSec.Text = "DESCRIÇÃO"
   ElseIf cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      cboCategoria.Visible = False
      lblCategoria.Visible = False

      cboCriterioSec.Visible = True

      cboCriterioSec.Text = "DESCRIÇÃO"
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
   Else
      Exit Sub
   End If
   
   cboCriterioSec.Clear
   cboIndice.Clear
   CboFormaPgto.Clear
   cboTipoPgto.Clear
   CboFormaPgtoSec.Clear
   cboTipoPgtoSec.Clear
   CboFormaPgtoSec.Visible = False
   cboTipoPgtoSec.Visible = False
   
   If cboCriterioPrinc.Text = "VENDEDOR" Or cboCriterioPrinc.Text = "CLIENTE" Then
      cboCriterioSec.Text = "TODOS"
   Else
      cboCriterioSec.Text = ""
   End If
End Sub

Private Sub cboCriterioPrinc_Click()
   cboCriterioPrinc_Change
End Sub

Private Sub cboCriterioPrinc_GotFocus()
   cboCriterioPrinc.Clear
   
   If cboTipo.Text = "PEDIDOS" Or cboTipo.Text = "ORÇAMENTOS" Then
      cboCriterioPrinc.AddItem "TODOS"
      cboCriterioPrinc.AddItem "VENDEDOR"
      cboCriterioPrinc.AddItem "CLIENTE"
      cboCriterioPrinc.AddItem "PERIODO"
      cboCriterioPrinc.AddItem "CÓDIGO"
      cboCriterioPrinc.AddItem "MENSAL"
   ElseIf cboTipo.Text = "OFICINA" Then
      cboCriterioPrinc.AddItem "TODOS"
      cboCriterioPrinc.AddItem "VENDEDOR"
      cboCriterioPrinc.AddItem "TECNICO"
      cboCriterioPrinc.AddItem "CLIENTE"
      cboCriterioPrinc.AddItem "PERIODO"
      cboCriterioPrinc.AddItem "CÓDIGO"
      cboCriterioPrinc.AddItem "MENSAL"
   ElseIf cboTipo.Text = "POR PRODUTOS" Then
      cboCriterioPrinc.AddItem "TODOS"
      cboCriterioPrinc.AddItem "MENSAL"
      cboCriterioPrinc.AddItem "ESPECIFICO/MENSAL"
      cboCriterioPrinc.AddItem "ESPECIFICO"
   ElseIf cboTipo.Text = "POR SERVIÇOS" Then
      cboCriterioPrinc.AddItem "TODOS"
      cboCriterioPrinc.AddItem "MENSAL"
      cboCriterioPrinc.AddItem "ESPECIFICO/MENSAL"
      cboCriterioPrinc.AddItem "ESPECIFICO"
   ElseIf cboTipo.Text = "PRODUTOS COM LUCRO" Then
      cboCriterioPrinc.AddItem "TODOS"
      cboCriterioPrinc.AddItem "MENSAL"
   
   End If
   
   moCombo.AttachTo cboCriterioPrinc
End Sub

Private Sub cboCriterioPrinc_Validate(Cancel As Boolean)
If cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
   lblMes.Visible = True
   cboMes.Visible = True
   lblAno.Visible = True
   cboAno.Visible = True
   cboDescricao.Visible = True
   lblDescricao.Visible = True
End If
End Sub

Private Sub cboCriterioSec_Change()
   If cboCriterioSec.Text = "DESCRIÇÃO" Or cboCriterioSec.Text = "REFERÊNCIA" Or cboCriterioSec.Text = "FABRICANTE" Or cboCriterioSec.Text = "TECNICO" Or cboCriterioSec.Text = "SERVIÇO" Then
      cboDescricao.Visible = True
      lblDescricao.Visible = True
      cboDescricao.SetFocus
   Else
      cboDescricao.Visible = False
      lblDescricao.Visible = False
   End If
   
   If cboCriterioSec.Text = "CÓD. BARRA" Then
      txtCodBarra.Visible = True
      lblCodBarra.Visible = True
      txtCodBarra.SetFocus
   Else
      txtCodBarra.Visible = False
      lblCodBarra.Visible = False
   End If
   
   If cboCriterioSec.Text = "CATEGORIA" Then
      cboCategoria.Visible = True
      lblCategoria.Visible = True
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
   Else
      cboCategoria.Visible = False
      lblCategoria.Visible = False
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
   End If

   If cboCriterioSec.Text = "MENSAL + CAT." Then
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
      cboCategoria.Visible = True
      lblCategoria.Visible = True
   End If
   
   If cboCriterioSec.Text = "MENSAL" Then
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
      cboCategoria.Visible = False
      lblCategoria.Visible = False
   End If
   
If cboCriterioSec.Text = "CÓD. BARRA" Then
   lblDescricao.Caption = "Cód. Barra"
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
   lblDescricao.Caption = "Descrição"
ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
   lblDescricao.Caption = "Referência"
ElseIf cboCriterioSec.Text = "FABRICANTE" Then
   lblDescricao.Caption = "Fabricante"
ElseIf cboCriterioSec.Text = "CATEGORIA" Then
   lblDescricao.Caption = "Categoria"
End If

End Sub

Private Sub cboCriterioSec_Click()
   cboCriterioSec_Change
End Sub

Private Sub cboCriterioSec_GotFocus()
cboCriterioSec.Clear

If cboCriterioPrinc.Text = "VENDEDOR" Then
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "CATEGORIA"
   cboCriterioSec.AddItem "MENSAL"
   cboCriterioSec.AddItem "MENSAL + CAT."
ElseIf cboCriterioPrinc.Text = "TECNICO" Then
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "SERVIÇO MENSAL"
ElseIf cboCriterioPrinc.Text = "CLIENTE" Then
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "MENSAL"
ElseIf cboCriterioPrinc.Text = "CATEGORIA" Then
   cboCriterioSec.AddItem "MENSAL"
ElseIf cboCriterioPrinc.Text = "ESPECIFICO" Or cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
   If cboTipo.Text = "POR PRODUTOS" Then
      'cboCriterioSec.AddItem "TODOS"
      cboCriterioSec.AddItem "DESCRIÇÃO"
      cboCriterioSec.AddItem "CÓD. BARRA"
      'cboCriterioSec.AddItem "MENSAL"
      cboCriterioSec.AddItem "REFERÊNCIA"
      cboCriterioSec.AddItem "FABRICANTE"
   ElseIf cboTipo.Text = "POR SERVIÇOS" Then
      cboCriterioSec.AddItem "TECNICO"
      cboCriterioSec.AddItem "SERVIÇO"
      'cboCriterioSec.AddItem "DESCRIÇÃO"
      'cboCriterioSec.AddItem "DESCRIÇÃO"
   End If
End If

moCombo.AttachTo cboCriterioSec
End Sub

Private Sub cboCriterioSec_LostFocus()
   If cboCriterioSec.Text = "" Then cboCriterioSec.Text = "TODOS"
End Sub

Private Sub cboDescricao_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboDescricao.Clear
   
If cboCriterioSec.Text = "DESCRIÇÃO" Then
   sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem r("descricao")
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
   sSQL = "SELECT DISTINCT REF FROM produtos ORDER BY REF;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem r("REF")
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "FABRICANTE" Then
   sSQL = "SELECT DISTINCT FABRICANTE FROM produtos ORDER BY FABRICANTE;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem ValidateNull(r("FABRICANTE"))
      r.MoveNext
   Loop
Else
   Exit Sub
End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboDescricao
End Sub

Private Sub CboFormaPgto_Change()
CboFormaPgtoSec.Clear
If CboFormaPgto.Text <> "TODOS" Then
   CboFormaPgtoSec.Visible = True
   If CboFormaPgtoSec.Text = "" Then
      CboFormaPgtoSec.Text = "TODOS"
   End If
Else
   CboFormaPgtoSec.Visible = False
End If
End Sub

Private Sub CboFormaPgto_Click()
   CboFormaPgto_Change
End Sub


Private Sub cboFormaPgto_GotFocus()
   CboFormaPgto.Clear
   CboFormaPgto.AddItem "TODOS"
   CboFormaPgto.AddItem "À VISTA"
   CboFormaPgto.AddItem "À PRAZO"
End Sub


Private Sub CboFormaPgtoSec_GotFocus()
   CboFormaPgtoSec.Clear
   If CboFormaPgto.Text = "À VISTA" Then
      CboFormaPgtoSec.AddItem "TODOS"
      If cboTipoPgto.Text = "TODOS" Or cboTipoPgto.Text = "" Then
         CboFormaPgtoSec.AddItem "COM CARTÃO"
         CboFormaPgtoSec.AddItem "SEM CARTÃO"
      End If
   ElseIf CboFormaPgto.Text = "À PRAZO" Then
      If cboTipo.Text <> "POR PRODUTOS" Then
         CboFormaPgtoSec.AddItem "TODOS"
         CboFormaPgtoSec.AddItem "COM ENTRADA"
         CboFormaPgtoSec.AddItem "SEM ENTRADA"
      Else
         CboFormaPgtoSec.AddItem "TODOS"
      End If
   End If
   If CboFormaPgtoSec.ListCount <> 0 Then CboFormaPgtoSec.ListIndex = 0
End Sub

Private Sub cboIndice_GotFocus()
cboIndice.Clear
cboIndice.AddItem "QUANT."
cboIndice.AddItem "PRODUTO"
moCombo.AttachTo cboIndice
End Sub

Private Sub cboMes_GotFocus()
cboMes.Clear

cboMes.AddItem "Janeiro"
cboMes.AddItem "Fevereiro"
cboMes.AddItem "Março"
cboMes.AddItem "Abril"
cboMes.AddItem "Maio"
cboMes.AddItem "Junho"
cboMes.AddItem "Julho"
cboMes.AddItem "Agosto"
cboMes.AddItem "Setembro"
cboMes.AddItem "Outubro"
cboMes.AddItem "Novembro"
cboMes.AddItem "Dezembro"

moCombo.AttachTo cboMes
End Sub

Private Sub cboMes_LostFocus()
   cboAno.SetFocus
End Sub

Private Sub cboTipo_Change()
If cboTipo.Text = "POR PRODUTOS" Then
   cmdExibirProdutos.Visible = False
   cmdExibirParcelas.Visible = False
   cmdExibirPedidos.Visible = True
ElseIf cboTipo.Text = "POR SERVIÇOS" Then
   cmdExibirProdutos.Visible = False
   cmdExibirParcelas.Visible = False
   cmdExibirPedidos.Visible = True
Else
   Exit Sub
End If
   
cboCriterioPrinc.Clear
cboCriterioSec.Clear
cboIndice.Clear
CboFormaPgto.Clear
cboTipoPgto.Clear
CboFormaPgtoSec.Clear
cboTipoPgtoSec.Clear
cboCriterioSec.Visible = False
CboFormaPgtoSec.Visible = False
cboTipoPgtoSec.Visible = False
End Sub

Private Sub cboTipo_Click()
cboTipo_Change
End Sub

Private Sub cboTipo_GotFocus()
cboTipo.Clear
cboTipo.AddItem "POR PRODUTOS"
cboTipo.AddItem "POR SERVIÇOS"
moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipoPgto_Change()
   If cboTipoPgto.Text = "CARTÃO" Then
         cboTipoPgtoSec.Visible = True
         If cboTipoPgtoSec.Text = "" Then
            cboTipoPgtoSec.Text = "TODOS"
         End If
   Else
      cboTipoPgtoSec.Visible = False
   End If
End Sub

Private Sub cboTipoPgto_Click()
   cboTipoPgto_Change
End Sub

Private Sub cboTipoPgto_GotFocus()
   cboTipoPgto.Clear
   cboTipoPgto.AddItem "TODOS"
   cboTipoPgto.AddItem "AVULSO"
   cboTipoPgto.AddItem "DINHEIRO"
   cboTipoPgto.AddItem "CARTÃO"
   cboTipoPgto.AddItem "CHEQUE"
   cboTipoPgto.AddItem "PROMISSORIA"
End Sub

Private Sub cboTipoPgtoSec_GotFocus()
   cboTipoPgtoSec.Clear
   cboTipoPgtoSec.AddItem "TODOS"
   cboTipoPgtoSec.AddItem "DEBITO"
   cboTipoPgtoSec.AddItem "CREDITO"
   
   If cboTipoPgtoSec.ListCount <> 0 Then cboTipoPgtoSec.ListIndex = 0
End Sub

Private Sub cboVendedor_Click()
   cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboVendedor.Clear
   
   sSQL = "SELECT codigo, nome, cargo FROM funcionario WHERE (CARGO IN ('vendedor', 'vendedora')) ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboVendedor.AddItem r("nome")
      cboVendedor.ItemData(cboVendedor.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboVendedor
End Sub

Private Sub cboVendedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboVendedor_LostFocus()
   On Error GoTo TrataErro
   
   If cboVendedor.Text = "" Then txtCodFunc.Text = "": Exit Sub
   If cboVendedor.ListIndex = -1 Then txtCodFunc.Text = "": Exit Sub
   txtCodFunc = cboVendedor.ItemData(cboVendedor.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cmdCalendario1_Click()
   Dim varData As Variant
   Dim fCal As Calendario
  
   varData = Empty                    'Inicializa a variável

   Set fCal = New Calendario      'Cria o form de calendário
   fCal.Show vbModal
  
   varData = fCal.DateSelected    'Recupera a data selecionada

   Unload fCal                           'Fecha o form
   Set fCal = Nothing                   'Destrói a variável
  
   If Not IsDate(varData) Then Exit Sub   'Valida a data
   If varData = 0 Then Exit Sub
  
   mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCalendario2_Click()
   Dim varData As Variant
   Dim fCal As Calendario
  
   varData = Empty                    'Inicializa a variável

   Set fCal = New Calendario      'Cria o form de calendário
   fCal.Show vbModal
  
   varData = fCal.DateSelected    'Recupera a data selecionada

   Unload fCal                           'Fecha o form
   Set fCal = Nothing                   'Destrói a variável
  
   If Not IsDate(varData) Then Exit Sub   'Valida a data
   If varData = 0 Then Exit Sub
  
   mskFim = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdExibirParcelas_Click()
If Grid.Col = 0 Then Exit Sub
   If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
         Vendas_Consulta_Geral_Parcelas.loadInformacoes (Grid.TextMatrix(Grid.Row, 1))
         Vendas_Consulta_Geral_Parcelas.Show 1
   End If
End Sub

Private Sub cmdExibirPedidos_Click()
If cboTipo.Text = "POR PRODUTOS" Then
   If Grid.Col = 0 Then Exit Sub
   If IsNumeric(Grid.TextMatrix(Grid.Row, 0)) = True Then
      If Grid.TextMatrix(Grid.Row, 0) = "" Then Exit Sub
      Vendas_Consulta_Pedidos.loadPedidos Grid.TextMatrix(Grid.Row, 0)
      Vendas_Consulta_Pedidos.Show 1
   End If
End If
End Sub

Private Sub cmdExibirProdutos_Click()
If Grid.Col = 0 Then Exit Sub
If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
   If Grid.Col = 1 Then
      If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), Grid.TextMatrix(Grid.Row, 7)
      Parcelas_Consulta_Produtos.Show 1
   End If
End If
End Sub

Private Sub cmdImprimir_Click()
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
   
   If cboTipo.Text = "PEDIDOS" Or cboTipo.Text = "ORÇAMENTOS" Or cboTipo.Text = "OFICINA" Then
      Set REL_Cons_Venda.Relatorio.Recordset = r
      
      REL_Cons_Venda.dfQuant.Caption = lblQtda.Caption
      REL_Cons_Venda.dfTotal.Caption = lblTotal.Caption

      If cboTipo.Text = "PEDIDOS" Then
         REL_Cons_Venda.lblTitulo.Caption = "RELATÓRIO DE VENDAS"
      ElseIf cboTipo.Text = "ORÇAMENTOS" Then
         REL_Cons_Venda.lblTitulo.Caption = "RELATÓRIO DE ORÇAMENTOS"
      End If

      If CboFormaPgto.Text = "TODOS" Then
         REL_Cons_Venda.rfForma.Caption = "TODAS"
      ElseIf CboFormaPgto.Text = "À VISTA" Then
         REL_Cons_Venda.rfForma.Caption = "À VISTA"
      ElseIf CboFormaPgto.Text = "À PRAZO" Then
         REL_Cons_Venda.rfForma.Caption = "À PRAZO"
      Else
         REL_Cons_Venda.rfForma.Caption = "TODAS"
      End If

      If cboTipoPgto.Text = "TODOS" Then
         REL_Cons_Venda.rfTIPO.Caption = "TODAS"
      ElseIf cboTipoPgto.Text = "AVULSO" Then
         REL_Cons_Venda.rfTIPO.Caption = "AVULSO"
      ElseIf cboTipoPgto.Text = "DINHEIRO" Then
         REL_Cons_Venda.rfTIPO.Caption = "DINHEIRO"
      ElseIf cboTipoPgto.Text = "CARTÃO" Then
         REL_Cons_Venda.rfTIPO.Caption = "CARTÃO"
      ElseIf cboTipoPgto.Text = "CHEQUE" Then
         REL_Cons_Venda.rfTIPO.Caption = "CHEQUE"
      ElseIf cboTipoPgto.Text = "PROMISSORIA" Then
         REL_Cons_Venda.rfTIPO.Caption = "PROMISSÓRIA"
      Else
         REL_Cons_Venda.rfTIPO.Caption = "TODAS"
      End If

      If cboCriterioPrinc.Text = "VENDEDOR" Then
         REL_Cons_Venda.rfCons1.Caption = "Vendedor = " & cboVendedor.Text & ""
      ElseIf cboCriterioPrinc.Text = "CLIENTE" Then
         REL_Cons_Venda.rfCons1.Caption = "Cliente = " & cboCliente.Text & ""
      ElseIf cboCriterioPrinc.Text = "PERIODO" Then
         REL_Cons_Venda.rfCons1.Caption = "Intervalo de " & mskInicio.Text & " à " & mskFim.Text
      ElseIf cboCriterioPrinc.Text = "CÓDIGO" Then
         REL_Cons_Venda.rfCons1.Caption = "Código = " & txtCodigo.Text & ""
      ElseIf cboCriterioPrinc.Text = "CATEGORIA" Then
         REL_Cons_Venda.rfCons1.Caption = "Categoria = " & cboCategoria.Text & ""
      ElseIf cboCriterioPrinc.Text = "MENSAL" Then
         REL_Cons_Venda.rfCons1.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
      ElseIf cboCriterioPrinc.Text = "TODOS" Then
         REL_Cons_Venda.rfCons1.Caption = "TODOS"
      Else
         REL_Cons_Venda.rfCons1.Caption = "TODOS"
      End If

      If cboCriterioSec.Text = "MENSAL" Then
         REL_Cons_Venda.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
      End If
      
      If cboCriterioSec.Text = "CATEGORIA" Then
         REL_Cons_Venda.rfCons3.Caption = "Categoria = " & cboCategoria.Text & ""
      End If

      REL_Cons_Venda.Relatorio.Ativar
      Unload REL_Cons_Venda

   ElseIf cboTipo.Text = "POR PRODUTOS" Then
      Set REL_Cons_Venda_Prod.Relatorio.Recordset = r

      REL_Cons_Venda_Prod.dfQuant.Caption = lblQtda.Caption
      REL_Cons_Venda_Prod.dfTotal.Caption = Format(lblTotal.Caption, "##,##0.00")

      If CboFormaPgto.Text = "TODOS" Then
         REL_Cons_Venda_Prod.rfForma.Caption = "TODAS"
      ElseIf CboFormaPgto.Text = "À VISTA" Then
         REL_Cons_Venda_Prod.rfForma.Caption = "À VISTA"
      ElseIf CboFormaPgto.Text = "À PRAZO" Then
         REL_Cons_Venda_Prod.rfForma.Caption = "À PRAZO"
      End If

      If cboTipoPgto.Text = "TODOS" Then
         REL_Cons_Venda_Prod.rfTIPO.Caption = "TODAS"
      ElseIf cboTipoPgto.Text = "AVULSO" Then
         REL_Cons_Venda_Prod.rfTIPO.Caption = "AVULSO"
      ElseIf cboTipoPgto.Text = "DINHEIRO" Then
         REL_Cons_Venda_Prod.rfTIPO.Caption = "DINHEIRO"
      ElseIf cboTipoPgto.Text = "CARTÃO" Then
         REL_Cons_Venda_Prod.rfTIPO.Caption = "CARTÃO"
      ElseIf cboTipoPgto.Text = "CHEQUE" Then
         REL_Cons_Venda_Prod.rfTIPO.Caption = "CHEQUE"
      ElseIf cboTipoPgto.Text = "PROMISSORIA" Then
         REL_Cons_Venda_Prod.rfTIPO.Caption = "PROMISSÓRIA"
      End If

      If cboCriterioPrinc.Text = "TODOS" Then
         REL_Cons_Venda_Prod.rfCons1.Caption = "TODOS"
      ElseIf cboCriterioPrinc.Text = "ESPECIFICO" Or cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
         If cboCriterioSec.Text = "DESCRIÇÃO" Then
            REL_Cons_Venda_Prod.rfCons1.Caption = "PRODUTO = " & cboDescricao.Text & ""
         ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
            REL_Cons_Venda_Prod.rfCons1.Caption = "CÓD. BARRA = " & cboDescricao.Text & ""
         ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
            REL_Cons_Venda_Prod.rfCons1.Caption = "REF.: = " & cboDescricao.Text & ""
         ElseIf cboCriterioSec.Text = "FABRICANTE" Then
            REL_Cons_Venda_Prod.rfCons1.Caption = "FABRICANTE = " & cboDescricao.Text & ""
         End If
         
         If cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
         Else
            REL_Cons_Venda_Prod.rfCons2.Caption = ""
         End If
         
      ElseIf cboCriterioPrinc.Text = "MENSAL" Then
         REL_Cons_Venda_Prod.rfCons1.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
      Else
         REL_Cons_Venda_Prod.rfCons1.Caption = "TODOS"
      End If

      If cboCriterioSec.Text = "MENSAL" Then
         REL_Cons_Venda_Prod.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
      End If

      If cboCriterioSec.Text = "CATEGORIA" Then
         REL_Cons_Venda_Prod.rfCons3.Caption = "Categoria = " & cboCategoria.Text & ""
      End If

      'REL_Cons_Venda_Prod.Relatorio.NomeImpressora = var_Impressora
      REL_Cons_Venda_Prod.Relatorio.Ativar
      Unload REL_Cons_Venda_Prod

   ElseIf cboTipo.Text = "PRODUTOS COM LUCRO" Then
   
   End If
   
   Me.Show 1
End Sub

Public Sub cmdLocalizar_Click()
'VARIAVEIS

Dim INDICE As String       'INDICE PARA ORGANIZAR OS DADOS
Dim Tipo As String         'FORMA DE PAGAMENTO
Dim var_VCartao As String  'VENDAS COM CARTÃO
Dim tipo_Cartao As String  'TIPO DE CARTÃO
Dim var_TipoPgto As String 'TIPO DE PAGAMENTO
Dim varFormaPagto As String

totalRegistros = "0"

'INDICE
If cboTipo.Text = "POR PRODUTOS" Then
   If cboIndice.Text = "QUANT." Then
      INDICE = "SUM(quantidade) DESC;"
   ElseIf cboIndice.Text = "PRODUTO" Then
      INDICE = "produtos.descricao, produtos.tamanho, produtos.ref;"
   Else
      INDICE = "produtos.descricao, produtos.tamanho, produtos.ref;"
   End If
End If

'FORMA DE PAGAMENTO
If CboFormaPgto.Text = "TODOS" Then
   Tipo = ""
ElseIf CboFormaPgto.Text = "À VISTA" And CboFormaPgtoSec.Text = "TODOS" Then
   Tipo = " AND (pedidos.tipo_pagamento = 'À Vista')"
ElseIf CboFormaPgto.Text = "À VISTA" And CboFormaPgtoSec.Text = "COM CARTÃO" Then
   Tipo = " AND (pedidos.tipo_pagamento = 'À Vista') AND (pedidos.pagamento = 'Cartao')"
ElseIf CboFormaPgto.Text = "À VISTA" And CboFormaPgtoSec.Text = "SEM CARTÃO" Then
   Tipo = " AND (pedidos.tipo_pagamento = 'À Vista') AND (pedidos.pagamento <> 'Cartao')"
ElseIf CboFormaPgto.Text = "À PRAZO" And CboFormaPgtoSec.Text = "TODOS" Then
   Tipo = " AND (pedidos.tipo_pagamento = 'À Prazo')"
   
ElseIf CboFormaPgto.Text = "À PRAZO" And CboFormaPgtoSec.Text = "COM ENTRADA" Then
   Tipo = " AND (pedidos.tipo_pagamento = 'à prazo') AND (pedidos.data_compra = parcelas.pagamento) AND (parcelas.numero = 1)"    'VER COMO SE FAZ ISSO SEM USAR CAMPO ENTRADA
   
ElseIf CboFormaPgto.Text = "À PRAZO" And CboFormaPgtoSec.Text = "SEM ENTRADA" Then
   Tipo = " AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.data_compra <> parcelas.pagamento)"  'VER COMO SE FAZ ISSO SEM USAR CAMPO ENTRADA
End If
   
'TIPO DE PAGAMENTO
If cboTipoPgto.Text = "TODOS" Then
   var_TipoPgto = ""
ElseIf cboTipoPgto.Text = "AVULSO" Then
   var_TipoPgto = " AND (pedidos.pagamento = 'Avulso')"
ElseIf cboTipoPgto.Text = "DINHEIRO" Then
   var_TipoPgto = " AND (pedidos.pagamento = 'Dinheiro')"
ElseIf cboTipoPgto.Text = "CARTÃO" And cboTipoPgtoSec.Text = "TODOS" Then
   var_TipoPgto = " AND (pedidos.pagamento = 'Cartao')"
ElseIf cboTipoPgto.Text = "CARTÃO" And cboTipoPgtoSec.Text = "DEBITO" Then
   var_TipoPgto = " AND (pedidos.tipo_cartao = 'D')"
ElseIf cboTipoPgto.Text = "CARTÃO" And cboTipoPgtoSec.Text = "CREDITO" Then
   var_TipoPgto = " AND (pedidos.tipo_cartao = 'C')"
ElseIf cboTipoPgto.Text = "CHEQUE" Then
   var_TipoPgto = " AND (pedidos.pagamento = 'Cheque')"
ElseIf cboTipoPgto.Text = "CARNÊ" Then
   var_TipoPgto = " AND (pedidos.pagamento = 'Carne')"
ElseIf cboTipoPgto.Text = "PROMISSORIA" Then
   var_TipoPgto = " AND (pedidos.pagamento = 'Promissoria')"
End If

    
'TIPO DE CONSULTA
Dim varTipoConsulta As String
If cboTipo.Text = "PEDIDOS" Then               'ver depois
   varTipoConsulta = "BALCAO"
ElseIf cboTipo.Text = "ORÇAMENTOS" Then
   varTipoConsulta = "ORÇAMENTO"
ElseIf cboTipo.Text = "OFICINA" Then
   varTipoConsulta = "OFICINA"
End If
   
If cboTipo.Text = "POR PRODUTOS" Then
   
      'TODOS
      If cboCriterioPrinc.Text = "TODOS" And cboCriterioSec.Text = "" Then
         sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao as var_desc, produtos.ref, produtos.tamanho as var_Tam, produtos.fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
            "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
            "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
            "WHERE (pedidos.tipo_pedido = 'BALCAO' or pedidos.tipo_pedido = 'OFICINA')  " & Tipo & tipo_Cartao & var_TipoPgto & _
            "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, produtos.ref, pedidos_itens.preco ORDER BY " & INDICE
            
      ElseIf cboCriterioPrinc.Text = "MENSAL" And cboCriterioSec.Text = "" Then
         If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
         
         sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao as var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
            "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
            "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
            "WHERE (MONTH(data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ") " & _
            "AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & Tipo & tipo_Cartao & var_TipoPgto & _
            "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco ORDER BY " & INDICE
         
      'ElseIf cboCriterioPrinc.Text = "PERIODO" And cboCriterioSec.Text = "" Then
       '  If Not IsDate(mskInicio) Or Not IsDate(mskFim) Then Exit Sub
        
        ' sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
        '    "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
        '    "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
        '    "WHERE (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
        '    "AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & Tipo & tipo_Cartao & var_TipoPgto & _
        '    "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, pedidos_itens.preco ORDER BY " & INDICE
         
      'ElseIf cboCriterioPrinc.Text = "CATEGORIA" And cboCriterioSec.Text = "" Then
       '  If cboCategoria.Text = "" Then Exit Sub
         
       '  sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
       '     "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
       '    "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
       '     "WHERE (produtos.categoria = '" & cboCategoria.Text & "') AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
       '     Tipo & tipo_Cartao & var_TipoPgto & _
       '     "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, pedidos_itens.preco ORDER BY " & INDICE
         
      'ElseIf cboCriterioPrinc.Text = "CATEGORIA" And cboCriterioSec.Text = "MENSAL" Then
      '   If cboCategoria.Text = "" Then Exit Sub
      '   If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
         
      '   sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
      '      "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
      '      "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      '      "WHERE (produtos.categoria = '" & cboCategoria.Text & "') AND (MONTH(data) = " & cboMes.ListIndex + 1 & ") " & _
      '      "AND (YEAR(data) = " & cboAno & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & Tipo & tipo_Cartao & var_TipoPgto & _
      '      "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, pedidos_itens.preco ORDER BY " & INDICE
         
      'ElseIf cboCriterioPrinc.Text = "VENDEDOR" And cboCriterioSec.Text = "TODOS" Then
      '   If cboVendedor.Text = "" Then Exit Sub
         
      '   sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
      '      "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
      '      "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      '      "WHERE (pedidos.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
      '      Tipo & tipo_Cartao & var_TipoPgto & _
      '      "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, pedidos_itens.preco ORDER BY " & INDICE
         
      'ElseIf cboCriterioPrinc.Text = "VENDEDOR" And cboCriterioSec.Text = "MENSAL" Then
      '   If cboVendedor.Text = "" Then Exit Sub
      '   If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
         
      '   sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
      '      "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
      '      "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      '      "WHERE (pedidos.cod_funcionario = " & txtCodFunc.Text & ") AND (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") " & _
      '      "AND (YEAR(pedidos.data_compra) = " & cboAno & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
      '      Tipo & tipo_Cartao & var_TipoPgto & _
      '      "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, pedidos_itens.preco ORDER BY " & INDICE
         
      'ElseIf cboCriterioPrinc.Text = "VENDEDOR" And cboCriterioSec.Text = "CATEGORIA" Then
      '   If cboVendedor.Text = "" Then Exit Sub
      '   If cboCategoria.Text = "" Then Exit Sub
         
      '   sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
      '      "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
      '      "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      '      "WHERE (produtos.categoria = '" & cboCategoria.Text & "') AND (pedidos.cod_funcioanrio = " & txtCodFunc.Text & ") " & _
      '      "AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & Tipo & tipo_Cartao & var_TipoPgto & _
      '      "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, pedidos_itens.preco ORDER BY " & INDICE
         
      'ElseIf cboCriterioPrinc.Text = "VENDEDOR" And cboCriterioSec.Text = "MENSAL" And cboCriterioSec.Text = "CATEGORIA" Then
      '   If cboVendedor.Text = "" Then Exit Sub
      '   If cboCategoria.Text = "" Then Exit Sub
      '   If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
         
      '   sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
      '      "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
      '      "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      '      "WHERE (produtos.categoria = '" & cboCategoria.Text & "') AND (pedidos.cod_funcionario = " & txtCodFunc.Text & ") " & _
      '      "AND (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos.data_compra) = " & cboAno & ") " & _
      '      "AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & Tipo & tipo_Cartao & var_TipoPgto & _
      '      "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, pedidos_itens.preco ORDER BY " & INDICE
         
      ElseIf cboCriterioPrinc.Text = "ESPECIFICO" Then
             If cboCriterioSec.Text = "DESCRIÇÃO" Then
                If cboDescricao.Text = "" Then Exit Sub
                sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                   "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                   "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                   "WHERE (produtos.descricao = '" & cboDescricao.Text & "') AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                   Tipo & tipo_Cartao & var_TipoPgto & _
                   "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco " & _
                   "ORDER BY " & INDICE
    
             ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
                If cboDescricao.Text = "" Then Exit Sub
                sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                   "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                   "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                   "WHERE (produtos.REF = '" & cboDescricao.Text & "') AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                   Tipo & tipo_Cartao & var_TipoPgto & _
                   "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco " & _
                   "ORDER BY " & INDICE
    
             ElseIf cboCriterioSec.Text = "FABRICANTE" Then
                If cboDescricao.Text = "" Then Exit Sub
                sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                   "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                   "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                   "WHERE (produtos.FABRICANTE = '" & cboDescricao.Text & "') AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                   Tipo & tipo_Cartao & var_TipoPgto & _
                   "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco " & _
                   "ORDER BY " & INDICE
    
             ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
                'If cboDescricao.Text = "" Then Exit Sub
                sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                   "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                   "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                   "WHERE (produtos.cod_barra = '" & txtCodBarra.Text & "') AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                   Tipo & tipo_Cartao & var_TipoPgto & _
                   "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco " & _
                   "ORDER BY " & INDICE
                
             ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioSec.Text = "MENSAL" Then
                If cboDescricao.Text = "" Then Exit Sub
                sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                   "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                   "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                   "WHERE (produtos.descricao = '" & cboDescricao.Text & "') AND (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") " & _
                   "AND (YEAR(pedidos.data_compra) = " & cboAno & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                   Tipo & tipo_Cartao & var_TipoPgto & _
                   "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco " & _
                   "ORDER BY " & INDICE
                
             ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioSec.Text = "MENSAL" Then
                If cboDescricao.Text = "" Then Exit Sub
                sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                   "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                   "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                   "WHERE (produtos.cod_barra = '" & txtCodBarra.Text & "') AND (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") " & _
                   "AND (YEAR(pedidos.data_compra) = " & cboAno & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                   Tipo & " " & tipo_Cartao & " " & var_TipoPgto & _
                   "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco " & _
                   "ORDER BY " & INDICE
            End If
      
      ElseIf cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
            
            If cboCriterioSec.Text = "DESCRIÇÃO" Then
               If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
               If cboDescricao.Text = "" Then Exit Sub
               sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                  "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                  "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                  "WHERE (produtos.descricao = '" & cboDescricao.Text & "') AND (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") " & _
                  "AND (YEAR(pedidos.data_compra) = " & cboAno & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                  Tipo & tipo_Cartao & var_TipoPgto & _
                  "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco ORDER BY " & INDICE
               
            ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
               If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
               If cboDescricao.Text = "" Then Exit Sub
               sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                  "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                  "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                  "WHERE (produtos.cod_barra = '" & txtCodBarra.Text & "') AND (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") " & _
                  "AND (YEAR(pedidos.data_compra) = " & cboAno & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                  Tipo & " " & tipo_Cartao & " " & var_TipoPgto & _
                  "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco ORDER BY " & INDICE
            
            ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
               If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
               If cboDescricao.Text = "" Then Exit Sub
               sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                  "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                  "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                  "WHERE (produtos.REF = '" & cboDescricao.Text & "') AND (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") " & _
                  "AND (YEAR(pedidos.data_compra) = " & cboAno & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                  Tipo & " " & tipo_Cartao & " " & var_TipoPgto & _
                  "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco ORDER BY " & INDICE
            
            ElseIf cboCriterioSec.Text = "FABRICANTE" Then
               If cboDescricao.Text = "" Then Exit Sub
               If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
               sSQL = "SELECT pedidos_itens.cod_produto, produtos.descricao AS var_desc, ref, tamanho as var_Tam, fabricante as var_Fab, SUM(pedidos_itens.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
                  "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                  "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                  "WHERE (produtos.FABRICANTE = '" & cboDescricao.Text & "') AND (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") " & _
                  "AND (YEAR(pedidos.data_compra) = " & cboAno & ") AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
                  Tipo & " " & tipo_Cartao & " " & var_TipoPgto & _
                  "GROUP BY pedidos_itens.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante,  produtos.ref, pedidos_itens.preco ORDER BY " & INDICE
            End If
    End If
        
ElseIf cboTipo.Text = "POR SERVIÇOS" Then


         'TODOS
         If cboCriterioPrinc.Text = "TODOS" And cboCriterioSec.Text = "" Then
            sSQL = "SELECT os_servicos.cod_produto, os_servicos.descricao as var_desc, SUM(os_servicos.quantidade) AS var_qtde, preco, SUM(preco * quantidade) AS var_total " & _
               "FROM produtos LEFT JOIN os_servicos ON produtos.codigo = os_servicos.cod_produto " & _
               "LEFT JOIN pedidos ON os_servicos.cod_pedido = pedidos.cod_pedido " & _
               "WHERE (pedidos.tipo_pedido = 'BALCAO' or pedidos.tipo_pedido = 'OFICINA')  " & Tipo & tipo_Cartao & var_TipoPgto & _
               "GROUP BY os_servicos.cod_produto, produtos.descricao, produtos.tamanho, produtos.fabricante, produtos.ref, os_servicos.preco ORDER BY " & INDICE
         End If
    

End If
      
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid_Produtos r

If r.State <> 0 Then r.Close
Set r = Nothing
   
printSQL = sSQL
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Double
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   
   For i = 0 To var_Grid.Rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub cmdSair_Click()
   Unload Me
End Sub



Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
   
'FORMATAR O GRID
With Grid
   .Clear
   .Cols = 7
   .Rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 0
   .ColWidth(3) = 0
   .ColWidth(4) = 0
   .ColWidth(5) = 0
   .ColWidth(6) = 0
End With

StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   posX = x
   Label3 = posX
   If Label3.Caption > 0 And Label3.Caption < 149 Then Grid.ToolTipText = ""
   If Label3.Caption > 150 And Label3.Caption < 930 Then Grid.ToolTipText = "Dê um duplo-clique para exibir os itens do Pedido."
   If Label3.Caption > 931 And Label3.Caption < 7230 Then Grid.ToolTipText = ""
   If Label3.Caption > 7231 And Label3.Caption < 8355 Then Grid.ToolTipText = "Dê um duplo-clique para exibir a forma de pgto."
   If Label3.Caption > 8356 And Label3.Caption < 9555 Then Grid.ToolTipText = ""
End Sub

Private Sub txtcodigo_GotFocus()
   SelectControl txtCodigo
End Sub

Private Sub mskFim_GotFocus()
SelectControl mskFim
End Sub

Private Sub mskFim_KeyPress(KeyAscii As Integer)
mskFim.Mask = "##/##/##"
End Sub

Private Sub mskFim_LostFocus()
If mskFim.Text = "" Or mskFim.Text = "__/__/__" Then
   mskFim.Mask = ""
   mskFim.Text = ""
   Exit Sub
Else
   If IsDate(mskFim.Text) Then
      cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskFim.SetFocus
      SelectControl mskFim
   End If
End If
End Sub

Private Sub mskInicio_GotFocus()
   SelectControl mskInicio
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
   mskInicio.Mask = "##/##/##"
End Sub

Sub FormatarGrid_Vendas(rTabela As ADODB.Recordset)
   Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1220
      .ColWidth(7) = 0
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "TIPO"
      .TextMatrix(0, 7) = "TIPO"
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = UCase(rTabela("nome"))
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("var_total"), ocMONEY)
            .TextMatrix(.Rows - 1, 5) = rTabela("tipo_pagamento")
            .TextMatrix(.Rows - 1, 6) = rTabela("pagamento")
            .TextMatrix(.Rows - 1, 7) = rTabela("tipo_pedido")
            
            
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
         .Col = 4
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      Grid.Redraw = True
   End With
   
   lblTotal.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
   lblEntrada.Caption = Format(0, ocMONEY)
picAguarde.Visible = False
End Sub

Sub FormatarGrid_VendasComEntrada(rTabela As ADODB.Recordset)
   Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 3600
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 800
      .ColWidth(7) = 1100
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "ENTRADA"
      .TextMatrix(0, 5) = "VALOR"
      .TextMatrix(0, 6) = "FORMA"
      .TextMatrix(0, 7) = "TIPO"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      .Redraw = False
      
      i = 1
      
            '.TextMatrix(.Rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            '.TextMatrix(.Rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            '.TextMatrix(.Rows - 1, 3) = UCase(rTabela("nome"))
            '.TextMatrix(.Rows - 1, 4) = Format(rTabela("var_total"), ocMONEY)
            '.TextMatrix(.Rows - 1, 5) = rTabela("tipo_pagamento")
            '.TextMatrix(.Rows - 1, 6) = rTabela("pagamento")
            '.TextMatrix(.Rows - 1, 7) = rTabela("tipo_pedido")

      
      
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = UCase(rTabela("nome"))
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("valor_final"), ocMONEY)
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("var_total"), ocMONEY)
            .TextMatrix(.Rows - 1, 6) = rTabela("tipo_pagamento")
            .TextMatrix(.Rows - 1, 7) = rTabela("pagamento")
            
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
         .Col = 4
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 5
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      Grid.Redraw = True
   End With
   
   lblTotal.Caption = Format(SomaGrid(Grid, 5), "##,##0.00")
   lblEntrada.Caption = Format(SomaGrid(Grid, 4), "##,##0.00")
picAguarde.Visible = False
End Sub

Private Sub mskInicio_LostFocus()
   If mskInicio.Text = "" Or mskInicio.Text = "__/__/__" Then
      mskInicio.Mask = ""
      mskInicio.Text = ""
      Exit Sub
   Else
      If IsDate(mskInicio.Text) Then
         If mskFim.Visible = True Then mskFim.SetFocus
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskInicio.SetFocus
         SelectControl mskInicio
      End If
   End If
End Sub
