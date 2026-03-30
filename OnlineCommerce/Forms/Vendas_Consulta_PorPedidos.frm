VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Vendas_Consulta_PorPedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE VENDAS"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "Vendas_Consulta_PorPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   13215
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
      Height          =   255
      Left            =   60
      TabIndex        =   41
      Top             =   8340
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
      MICON           =   "Vendas_Consulta_PorPedidos.frx":23D2
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
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   60
      ScaleHeight     =   1905
      ScaleWidth      =   13065
      TabIndex        =   15
      ToolTipText     =   "Imprimir"
      Top             =   780
      Width           =   13095
      Begin VB.ComboBox cboTipoPgto 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   1500
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
         Height          =   1755
         Left            =   5400
         TabIndex        =   16
         Top             =   60
         Width           =   7635
         Begin VB.Frame frmFiltro1 
            Height          =   1485
            Left            =   60
            TabIndex        =   17
            Top             =   180
            Width           =   7455
            Begin ChamaleonBtn.chameleonButton cmdCalendario2 
               Height          =   315
               Left            =   2700
               TabIndex        =   50
               Tag             =   "Calendario"
               Top             =   420
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               BTYPE           =   8
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
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Vendas_Consulta_PorPedidos.frx":23EE
               PICN            =   "Vendas_Consulta_PorPedidos.frx":240A
               PICH            =   "Vendas_Consulta_PorPedidos.frx":475D
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdCalendario1 
               Height          =   315
               Left            =   1080
               TabIndex        =   49
               Tag             =   "Calendario"
               Top             =   420
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               BTYPE           =   8
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
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Vendas_Consulta_PorPedidos.frx":6AB0
               PICN            =   "Vendas_Consulta_PorPedidos.frx":6ACC
               PICH            =   "Vendas_Consulta_PorPedidos.frx":8E1F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboCliente 
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   1080
               Visible         =   0   'False
               Width           =   4965
            End
            Begin VB.ComboBox cboVendedor 
               Height          =   315
               Left            =   120
               TabIndex        =   23
               Top             =   1080
               Visible         =   0   'False
               Width           =   4965
            End
            Begin VB.TextBox txtCodigo 
               Height          =   315
               Left            =   120
               TabIndex        =   22
               Top             =   420
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   1500
               Sorted          =   -1  'True
               TabIndex        =   21
               Top             =   420
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMes 
               Height          =   315
               Left            =   120
               TabIndex        =   20
               Top             =   420
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtCodCliente 
               Appearance      =   0  'Flat
               Height          =   195
               Left            =   4380
               TabIndex        =   19
               Top             =   780
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtCodFunc 
               Appearance      =   0  'Flat
               Height          =   195
               Left            =   3720
               TabIndex        =   18
               Top             =   780
               Visible         =   0   'False
               Width           =   615
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   315
               Left            =   120
               TabIndex        =   25
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
               Left            =   1740
               TabIndex        =   26
               Top             =   420
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdCal1 
               Height          =   315
               Left            =   1320
               TabIndex        =   47
               Tag             =   "Calendario"
               Top             =   420
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               BTYPE           =   8
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
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Vendas_Consulta_PorPedidos.frx":B172
               PICN            =   "Vendas_Consulta_PorPedidos.frx":B18E
               PICH            =   "Vendas_Consulta_PorPedidos.frx":D4E1
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   315
               Left            =   120
               TabIndex        =   48
               Top             =   420
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdLocalizar 
               Height          =   315
               Left            =   5160
               TabIndex        =   51
               Top             =   1080
               Width           =   1035
               _ExtentX        =   1826
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
               MICON           =   "Vendas_Consulta_PorPedidos.frx":F834
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
               Height          =   315
               Left            =   6240
               TabIndex        =   52
               Top             =   1080
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
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
               MICON           =   "Vendas_Consulta_PorPedidos.frx":F850
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
               TabIndex        =   35
               Top             =   840
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblVendedor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vendedor(a):"
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   840
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblAte 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "até"
               Height          =   195
               Left            =   1440
               TabIndex        =   33
               Top             =   480
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label lblFim 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data final:"
               Height          =   195
               Left            =   1740
               TabIndex        =   32
               Top             =   180
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label lblInicio 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data inicial:"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   180
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label lblCodigo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido:"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   180
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label lblAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   1500
               TabIndex        =   29
               Top             =   180
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lblMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   180
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblData 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data::"
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   180
               Visible         =   0   'False
               Width           =   435
            End
         End
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   2595
      End
      Begin VB.ComboBox cboCriterioPrinc 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   900
         Width           =   2595
      End
      Begin VB.ComboBox cboCriterioSec 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   900
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.ComboBox cboIndice 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   300
         Width           =   2595
      End
      Begin VB.ComboBox cboFormaPgto 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1500
         Width           =   2595
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pgto"
         Height          =   195
         Left            =   2760
         TabIndex        =   40
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   60
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Criterio"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Organizar por:"
         Height          =   195
         Left            =   2760
         TabIndex        =   37
         Top             =   60
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pgto"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1260
         Width           =   1035
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5580
      Picture         =   "Vendas_Consulta_PorPedidos.frx":F86C
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   13065
      TabIndex        =   5
      Top             =   60
      Width           =   13095
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE VENDAS - POR PEDIDOS"
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
         TabIndex        =   7
         Top             =   180
         Width           =   6045
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   240
         Picture         =   "Vendas_Consulta_PorPedidos.frx":108A4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   14
      Top             =   9375
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18971
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21:13"
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
      Height          =   4875
      Left            =   60
      TabIndex        =   42
      Top             =   2760
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8599
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirProdutos 
      Height          =   255
      Left            =   60
      TabIndex        =   43
      Top             =   8340
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
      MICON           =   "Vendas_Consulta_PorPedidos.frx":170EA
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
      TabIndex        =   44
      Top             =   8340
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
      MICON           =   "Vendas_Consulta_PorPedidos.frx":17106
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
      Caption         =   "ACRÉSCIMOS:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10380
      TabIndex        =   58
      Top             =   8640
      Width           =   1005
   End
   Begin VB.Label lblTotalAcrescimo 
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
      Left            =   11460
      TabIndex        =   57
      Top             =   8640
      Width           =   1635
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCONTOS:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10500
      TabIndex        =   56
      Top             =   8340
      Width           =   915
   End
   Begin VB.Label lblTotalDesconto 
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
      Left            =   11460
      TabIndex        =   55
      Top             =   8340
      Width           =   1635
   End
   Begin VB.Label lblSubtotalBruto 
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
      Left            =   11460
      TabIndex        =   54
      Top             =   8040
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DE VENDAS - BRUTO:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9420
      TabIndex        =   53
      Top             =   8040
      Width           =   1995
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DE VENDAS - LIQUIDO:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9300
      TabIndex        =   46
      Top             =   8940
      Width           =   2100
   End
   Begin VB.Label lblSubtotal 
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
      Left            =   11460
      TabIndex        =   45
      Top             =   8940
      Width           =   1635
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
      Left            =   6180
      TabIndex        =   13
      Top             =   8520
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
      Left            =   7740
      TabIndex        =   12
      Top             =   8940
      Visible         =   0   'False
      Width           =   510
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
      Left            =   11460
      TabIndex        =   11
      Top             =   7740
      Width           =   1635
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QTDE DE VENDAS:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10140
      TabIndex        =   10
      Top             =   7740
      Width           =   1290
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1635
      Left            =   8580
      Top             =   7680
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   8940
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Vendas_Consulta_PorPedidos"
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
Private Sub Limpar_Grid_Venda()
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 8
      .rows = 2
      
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
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
   lblQtda.Caption = Format(0, ocMONEY)
   lblSubtotal.Caption = Format(0, ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
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

Private Sub PreencherPrincipal()
cboCriterioPrinc.Clear

cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "VENDEDOR"
cboCriterioPrinc.AddItem "CLIENTE"
cboCriterioPrinc.AddItem "PERIODO"
cboCriterioPrinc.AddItem "CÓDIGO"
cboCriterioPrinc.AddItem "MENSAL"
End Sub

Private Sub PreencherTipoPgto()
If cboFormaPgto.Text = "À VISTA" Then
    cboTipoPgto.Clear
    cboTipoPgto.AddItem "TODOS"
    cboTipoPgto.AddItem "DINHEIRO"
    cboTipoPgto.AddItem "CARTÃO DÉBITO"
    cboTipoPgto.AddItem "CARTÃO CRÉDITO"
    cboTipoPgto.AddItem "TRANSFERÊNCIA"
    cboTipoPgto.AddItem "DEPOSITO"
    cboTipoPgto.AddItem "FINANCEIRA"
    cboTipoPgto.AddItem "CHEQUE"
    cboTipoPgto.AddItem "PIX"
ElseIf cboFormaPgto.Text = "À PRAZO" Then
    cboTipoPgto.Clear
    cboTipoPgto.AddItem "TODOS"
    cboTipoPgto.AddItem "CARTÃO DÉBITO"
    cboTipoPgto.AddItem "CARTÃO CRÉDITO"
    cboTipoPgto.AddItem "TRANSFERÊNCIA"
    cboTipoPgto.AddItem "DEPOSITO"
    cboTipoPgto.AddItem "FINANCEIRA"
    cboTipoPgto.AddItem "PROMISSÓRIA"
    cboTipoPgto.AddItem "CHEQUE"
    cboTipoPgto.AddItem "BOLETO"
ElseIf cboFormaPgto.Text = "TODOS" Then
    cboTipoPgto.Clear
    cboTipoPgto.AddItem "TODOS"
    cboTipoPgto.AddItem "DINHEIRO"
    cboTipoPgto.AddItem "CARTÃO DÉBITO"
    cboTipoPgto.AddItem "CARTÃO CRÉDITO"
    cboTipoPgto.AddItem "TRANSFERÊNCIA"
    cboTipoPgto.AddItem "DEPOSITO"
    cboTipoPgto.AddItem "FINANCEIRA"
    cboTipoPgto.AddItem "PROMISSÓRIA"
    cboTipoPgto.AddItem "CHEQUE"
    cboTipoPgto.AddItem "BOLETO"
    cboTipoPgto.AddItem "PIX"
End If
End Sub

Private Sub PreencherIndice()
cboIndice.Clear
cboIndice.AddItem "PEDIDO"
cboIndice.AddItem "DATA"
cboIndice.AddItem "POR NOME"
cboIndice.AddItem "FORMA PGTO"
cboIndice.AddItem "VALOR"
End Sub

Private Sub PreencherFormaPgto()
cboFormaPgto.Clear
cboFormaPgto.AddItem "À VISTA"
cboFormaPgto.AddItem "À PRAZO"
cboFormaPgto.AddItem "TODOS"
End Sub

Private Sub PreencherTipoConsulta()
cboTipo.Clear
cboTipo.AddItem "VENDA"
cboTipo.AddItem "ORÇAMENTOS"
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
      
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
      
      cboCriterioSec.Visible = False
      
      If cboMes.Visible = True Then cboMes.SetFocus
      
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
      
      cboCriterioSec.Visible = True

      cboCriterioSec.Text = "DESCRIÇÃO"
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
   Else
      Exit Sub
   End If
   
   'cboCriterioSec.Clear
   'cboIndice.Clear
   'cboFormaPgto.Clear
   'cboTipoPgto.Clear
   'CboFormaPgtoSec.Clear
   'cboTipoPgtoSec.Clear
  
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

moCombo.AttachTo cboCriterioPrinc
End Sub

Private Sub cboCriterioPrinc_Validate(Cancel As Boolean)
If cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
   lblMes.Visible = True
   cboMes.Visible = True
   lblAno.Visible = True
   cboAno.Visible = True
End If
End Sub

Private Sub cboCriterioSec_Change()
If cboCriterioSec.Text = "TODOS" Then
   lblMes.Visible = False
   cboMes.Visible = False
   lblAno.Visible = False
   cboAno.Visible = False
   cmdCal1.Visible = False
   mskData.Visible = False
   lblData.Visible = False
End If

If cboCriterioSec.Text = "MENSAL" Then
   lblMes.Visible = True
   cboMes.Visible = True
   lblAno.Visible = True
   cboAno.Visible = True
   cmdCal1.Visible = False
   mskData.Visible = False
   lblData.Visible = False
End If

If cboCriterioSec.Text = "DATA" Then
   lblMes.Visible = False
   cboMes.Visible = False
   lblAno.Visible = False
   cboAno.Visible = False
   cmdCal1.Visible = True
   mskData.Visible = True
   lblData.Visible = True
End If
End Sub

Private Sub cboCriterioSec_Click()
   cboCriterioSec_Change
End Sub

Private Sub cboCriterioSec_GotFocus()
cboCriterioSec.Clear

If cboCriterioPrinc.Text = "VENDEDOR" Then
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "MENSAL"
   cboCriterioSec.AddItem "DATA"
ElseIf cboCriterioPrinc.Text = "CLIENTE" Then
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "MENSAL"
End If

moCombo.AttachTo cboCriterioSec
End Sub

Private Sub cboCriterioSec_LostFocus()
   If cboCriterioSec.Text = "" Then cboCriterioSec.Text = "TODOS"
End Sub


Private Sub cboFormaPgto_LostFocus()
PreencherTipoPgto
End Sub


Private Sub cboIndice_GotFocus()

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
If cboTipo.Text = "VENDA" Then
   cmdExibirProdutos.Visible = True
   cmdExibirParcelas.Visible = True
   cmdExibirPedidos.Visible = False
ElseIf cboTipo.Text = "OFICINA" Then
   cmdExibirProdutos.Visible = True
   cmdExibirParcelas.Visible = True
   cmdExibirPedidos.Visible = False
ElseIf cboTipo.Text = "ORÇAMENTOS" Then
   cmdExibirProdutos.Visible = True
   cmdExibirParcelas.Visible = True
   cmdExibirPedidos.Visible = False
Else
   Exit Sub
End If
   
'cboCriterioPrinc.Clear
'cboCriterioSec.Clear
'cboIndice.Clear
'cboFormaPgto.Clear
'cboTipoPgto.Clear
'CboFormaPgtoSec.Clear
'cboTipoPgtoSec.Clear
cboCriterioSec.Visible = False
'cboTipoPgtoSec.Visible = False
End Sub

Private Sub cboTipo_Click()
cboTipo_Change
End Sub

Private Sub cboTipo_GotFocus()
moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipoPgto_Click()
   'cboTipoPgto_Change
End Sub

Private Sub cboVendedor_Click()
   cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboVendedor.Clear
   
   sSQL = "SELECT codigo, nome, cargo FROM funcionario ORDER BY nome;"
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

Private Sub cmdCal1_Click()
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

mskData = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
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

If cboTipo.Text = "VENDA" Or cboTipo.Text = "ORÇAMENTOS" Or cboTipo.Text = "OFICINA" Then
   Set REL_Cons_Venda.Relatorio.Recordset = r
   
   REL_Cons_Venda.dfQuant.Caption = lblQtda.Caption
   'REL_Cons_Venda.dfTotal.Caption = lblTotal.Caption
   REL_Cons_Venda.dfSubtotal.Caption = lblSubtotal.Caption
   REL_Cons_Venda.dfSubtotalBruto.Caption = lblSubtotalBruto.Caption
   
   
   'REL_Cons_Venda.dfDesc.Caption = lblDesc.Caption
   'REL_Cons_Venda.dfAcresc.Caption = lblAcresc.Caption

   If cboTipo.Text = "VENDA" Then
      REL_Cons_Venda.lblTitulo.Caption = "RELATÓRIO DE VENDAS POR PEDIDOS"
   ElseIf cboTipo.Text = "ORÇAMENTOS" Then
      REL_Cons_Venda.lblTitulo.Caption = "RELATÓRIO DE ORÇAMENTOS"
   End If

   If cboFormaPgto.Text = "TODOS" Then
      REL_Cons_Venda.rfForma.Caption = "TODAS"
   ElseIf cboFormaPgto.Text = "À VISTA" Then
      REL_Cons_Venda.rfForma.Caption = "À VISTA"
   ElseIf cboFormaPgto.Text = "À PRAZO" Then
      REL_Cons_Venda.rfForma.Caption = "À PRAZO"
   Else
      REL_Cons_Venda.rfForma.Caption = "TODAS"
   End If

   'If cboTipoPgto.Text = "TODOS" Then
   '   REL_Cons_Venda.rfTipo.Caption = "TODAS"
   'ElseIf cboTipoPgto.Text = "AVULSO" Then
   '   REL_Cons_Venda.rfTipo.Caption = "AVULSO"
   'ElseIf cboTipoPgto.Text = "DINHEIRO" Then
   '   REL_Cons_Venda.rfTipo.Caption = "DINHEIRO"
   'ElseIf cboTipoPgto.Text = "CARTÃO" Then
   '   REL_Cons_Venda.rfTipo.Caption = "CARTÃO"
   'ElseIf cboTipoPgto.Text = "CHEQUE" Then
   '   REL_Cons_Venda.rfTipo.Caption = "CHEQUE"
   'ElseIf cboTipoPgto.Text = "PROMISSORIA" Then
   '   REL_Cons_Venda.rfTipo.Caption = "PROMISSÓRIA"
   'ElseIf cboTipoPgto.Text = "PIX" Then
   '   REL_Cons_Venda.rfTipo.Caption = "PIX"
   'Else
   '   REL_Cons_Venda.rfTipo.Caption = "TODAS"
   'End If

   If cboCriterioPrinc.Text = "VENDEDOR" Then
      REL_Cons_Venda.rfCons1.Caption = "Vendedor = " & cboVendedor.Text & ""
   ElseIf cboCriterioPrinc.Text = "CLIENTE" Then
      REL_Cons_Venda.rfCons1.Caption = "Cliente = " & cboCliente.Text & ""
   ElseIf cboCriterioPrinc.Text = "PERIODO" Then
      REL_Cons_Venda.rfCons1.Caption = "Intervalo de " & mskInicio.Text & " à " & mskFim.Text
   ElseIf cboCriterioPrinc.Text = "CÓDIGO" Then
      REL_Cons_Venda.rfCons1.Caption = "Código = " & txtCodigo.Text & ""
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
   
   REL_Cons_Venda.Relatorio.Ativar
   Unload REL_Cons_Venda

End If

Me.Show 1
End Sub

Public Sub cmdLocalizar_Click()
totalRegistros = "0"
Dim vTABELAS As String
Dim SQLCount As String
Dim SQLWhere As String
Dim vWhere As String
Dim vPegarCliente As String

'INDICE
Dim INDICE As String
If cboIndice.Text = "DATA" Then
   INDICE = "data_compra;"
ElseIf cboIndice.Text = "POR NOME" Then
   INDICE = "nome;"
ElseIf cboIndice.Text = "FORMA PGTO" Then
   INDICE = "tipo_pagamento;"
ElseIf cboIndice.Text = "VALOR" Then
   INDICE = "pedidos_1.total;"
ElseIf cboIndice.Text = "PEDIDO" Then
   INDICE = "pedidos_1.cod_pedido;"
Else
   INDICE = "pedidos_1.cod_pedido;"
End If

'FORMA DE PAGAMENTO ============= ver depois
Dim TipoPgto As String

If cboFormaPgto.Text = "À VISTA" Then
   TipoPgto = " AND (pedidos_1.tipo_pagamento = 'À Vista')"
ElseIf cboFormaPgto.Text = "À PRAZO" Then
   TipoPgto = " AND (pedidos_1.tipo_pagamento = 'À prazo')"
Else
    TipoPgto = " AND (pedidos_1.tipo_pagamento IN ('À Vista', 'À prazo'))"
End If

'TIPO DE PAGAMENTO (PARCELAS)
Dim vTipoPgtoParcelas As String
'If cboFormaPgto.Text = "À VISTA" Then
'    If cboTipoPgto.Text = "TODOS" Then
'    ElseIf cboTipoPgto.Text = "DINHEIRO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'DINHEIRO')"
'    ElseIf cboTipoPgto.Text = "CARTÃO DÉBITO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CARTAO') and (parcelas.TIPO_CARTAO = 'D')"""
'    ElseIf cboTipoPgto.Text = "CARTÃO CRÉDITO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CARTAO') and (parcelas.TIPO_CARTAO = 'C')"""
'    ElseIf cboTipoPgto.Text = "TRANSFERÊNCIA" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'TRANSFERENCIA')"
'    ElseIf cboTipoPgto.Text = "DEPOSITO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'DEPOSITO')"
'    ElseIf cboTipoPgto.Text = "FINANCEIRA" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'FINANCEIRA')"
'    ElseIf cboTipoPgto.Text = "PROMISSÓRIA" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'PROMISSORIA')"
'    ElseIf cboTipoPgto.Text = "CHEQUE" Then
 '       vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CHEQUE')"
 '   ElseIf cboTipoPgto.Text = "BOLETO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'BOLETO')"
'    ElseIf cboTipoPgto.Text = "PIX" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'PIX')"
'    End If
'ElseIf cboFormaPgto.Text = "À PRAZO" Then
'    If cboTipoPgto.Text = "TODOS" Then
'    ElseIf cboTipoPgto.Text = "DINHEIRO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'DINHEIRO')"
'    ElseIf cboTipoPgto.Text = "CARTÃO DÉBITO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CARTAO') and (parcelas.TIPO_CARTAO = 'D')"""
'    ElseIf cboTipoPgto.Text = "CARTÃO CRÉDITO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CARTAO') and (parcelas.TIPO_CARTAO = 'C')"""
'    ElseIf cboTipoPgto.Text = "TRANSFERÊNCIA" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'TRANSFERENCIA')"
'    ElseIf cboTipoPgto.Text = "DEPOSITO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'DEPOSITO')"
'    ElseIf cboTipoPgto.Text = "FINANCEIRA" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'FINANCEIRA')"
'    ElseIf cboTipoPgto.Text = "PROMISSÓRIA" Then
 '       vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'PROMISSORIA')"
'    ElseIf cboTipoPgto.Text = "CHEQUE" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CHEQUE')"
'    ElseIf cboTipoPgto.Text = "BOLETO" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'BOLETO')"
'    ElseIf cboTipoPgto.Text = "PIX" Then
'        vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'PIX')"
'    End If
'Else
'End If

If cboFormaPgto.Text = "À PRAZO" Then
    If cboTipoPgto.Text = "TODOS" Then
       vTipoPgtoParcelas = ""
    ElseIf cboTipoPgto.Text = "DINHEIRO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'DINHEIRO')"
    ElseIf cboTipoPgto.Text = "CARTÃO DÉBITO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'CARTAO') and (parcelas.TIPO_CARTAO = 'D')"
    ElseIf cboTipoPgto.Text = "CARTÃO CRÉDITO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'CARTAO') and (parcelas.TIPO_CARTAO = 'C')"
    ElseIf cboTipoPgto.Text = "TRANSFERÊNCIA" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'TRANSFERENCIA')"
    ElseIf cboTipoPgto.Text = "DEPOSITO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'DEPOSITO')"
    ElseIf cboTipoPgto.Text = "FINANCEIRA" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'FINANCEIRA')"
    ElseIf cboTipoPgto.Text = "PROMISSÓRIA" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'PROMISSORIA')"
    ElseIf cboTipoPgto.Text = "CHEQUE" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'Cheque')"
    ElseIf cboTipoPgto.Text = "BOLETO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'BOLETO')"
    ElseIf cboTipoPgto.Text = "PIX" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) = 'PIX')"
    Else
        vTipoPgtoParcelas = ""
    End If
ElseIf cboFormaPgto.Text = "À VISTA" Then
    If cboTipoPgto.Text = "TODOS" Then
       vTipoPgtoParcelas = ""
    ElseIf cboTipoPgto.Text = "DINHEIRO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'DINHEIRO')"
    ElseIf cboTipoPgto.Text = "CARTÃO DÉBITO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'CARTAO') and (parcelas.TIPO_CARTAO = 'D')"
    ElseIf cboTipoPgto.Text = "CARTÃO CRÉDITO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'CARTAO') and (parcelas.TIPO_CARTAO = 'C')"
    ElseIf cboTipoPgto.Text = "TRANSFERÊNCIA" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'TRANSFERENCIA')"
    ElseIf cboTipoPgto.Text = "DEPOSITO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'DEPOSITO')"
    ElseIf cboTipoPgto.Text = "FINANCEIRA" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'FINANCEIRA')"
    ElseIf cboTipoPgto.Text = "PROMISSÓRIA" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'PROMISSORIA')"
    ElseIf cboTipoPgto.Text = "CHEQUE" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'Cheque')"
    ElseIf cboTipoPgto.Text = "BOLETO" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'BOLETO')"
    ElseIf cboTipoPgto.Text = "PIX" Then
       vTipoPgtoParcelas = " AND ((CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Vista' THEN parcelas.FORMA_PGTO else pedidos_1.PAGAMENTO END) = 'PIX')"
    Else
        vTipoPgtoParcelas = ""
    End If
Else
End If

'TIPO DE CONSULTA
Dim varTipoConsulta As String
If cboTipo.Text = "VENDA" Then
   varTipoConsulta = "VENDA"
ElseIf cboTipo.Text = "ORÇAMENTOS" Then
   varTipoConsulta = "ORÇAMENTO"
End If


'consulta principal
'sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO as pFormaPgto, parcelas.VALOR_FINAL as pValorPgto, parcelas.TIPO_CARTAO, " & _
       "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto, "
 

'vTABELAS = " FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO"

If cboCriterioPrinc.Text = "" Then Exit Sub
   
   'VENDEDOR - TODOS
    If cboCriterioPrinc.Text = "VENDEDOR" And cboCriterioSec.Text = "TODOS" Then
        If cboVendedor.Text = "" Then Limpar_Grid_Venda: Exit Sub
      
        'sSQL = sSQL & vTABELAS & " Where (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        'SQLWhere = " Where (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "

        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO AS pFormaPgto, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
            "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto, " & _
            "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO WHERE (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = '0') " & TipoPgto & ") AS Nome " & _
            "FROM pedidos AS pedidos_1 INNER JOIN  parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO " & _
            "WHERE (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = 0) " & TipoPgto & " " & _
            "ORDER BY var_codped"
        
   'VENDEDOR - MENSAL
    ElseIf cboCriterioPrinc.Text = "VENDEDOR" And cboCriterioSec.Text = "MENSAL" Then
        If cboVendedor.Text = "" Then Limpar_Grid_Venda: Exit Sub
        If cboMes.Text = "" Or cboAno.Text = "" Then Limpar_Grid_Venda: Exit Sub
      
        'sSQL = sSQL & vTABELAS & " Where (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (Month(pedidos_1.data_compra) = " & cboMES.ListIndex + 1 & ") And (Year(pedidos_1.data_compra) = " & cboAno & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        'SQLWhere = " Where (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (Month(pedidos_1.data_compra) = " & cboMES.ListIndex + 1 & ") And (Year(pedidos_1.data_compra) = " & cboAno & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "

        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO AS pFormaPgto, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
            "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto, " & _
            "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO WHERE (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (Month(pedidos_1.data_compra) = " & cboMes.ListIndex + 1 & ") And (Year(pedidos_1.data_compra) = " & cboAno & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = '0') " & TipoPgto & ") AS Nome " & _
            "FROM pedidos AS pedidos_1 INNER JOIN  parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO " & _
            "WHERE (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (Month(pedidos_1.data_compra) = " & cboMes.ListIndex + 1 & ") And (Year(pedidos_1.data_compra) = " & cboAno & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = 0) " & TipoPgto & " " & _
            "ORDER BY var_codped"

   'VENDEDOR - DATA
    ElseIf cboCriterioPrinc.Text = "VENDEDOR" And cboCriterioSec.Text = "DATA" Then
        If cboVendedor.Text = "" Then Limpar_Grid_Venda: Exit Sub
        If mskData.Text = "" Then Exit Sub

        'sSQL = sSQL & vTABELAS & " Where (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos_1.data_compra = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        'SQLWhere = " Where (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos_1.data_compra = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "

        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO AS pFormaPgto, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
            "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto, " & _
            "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO WHERE (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos_1.data_compra = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = '0') " & TipoPgto & ") AS Nome " & _
            "FROM pedidos AS pedidos_1 INNER JOIN  parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO " & _
            "WHERE (pedidos_1.cod_funcionario = " & txtCodFunc.Text & ") AND (pedidos_1.data_compra = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = 0) " & TipoPgto & " " & _
            "ORDER BY var_codped"

   'CLIENTE - TODOS
    ElseIf cboCriterioPrinc.Text = "CLIENTE" And cboCriterioSec.Text = "TODOS" Then
        If txtCodCliente.Text = "" Then Limpar_Grid_Venda: Exit Sub

        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO as pFormaPgto, parcelas.VALOR_FINAL as pValorPgto, parcelas.TIPO_CARTAO, cliente.nome, " & _
                "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto "

        vTABELAS = " FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO"

        sSQL = sSQL & vTABELAS & " Where (cliente.codigo = " & txtCodCliente.Text & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        SQLWhere = " Where (cliente.codigo = " & txtCodCliente.Text & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "

   'CLIENTE - MENSAL
    ElseIf cboCriterioPrinc.Text = "CLIENTE" And cboCriterioSec.Text = "MENSAL" Then
        If txtCodCliente.Text = "" Then Limpar_Grid_Venda: Exit Sub
        If cboMes.Text = "" Or cboAno.Text = "" Then Limpar_Grid_Venda: Exit Sub

        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO as pFormaPgto, parcelas.VALOR_FINAL as pValorPgto, parcelas.TIPO_CARTAO, cliente.nome, " & _
                "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto "

        vTABELAS = " FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO"

        sSQL = sSQL & vTABELAS & " Where (cliente.codigo = " & txtCodCliente.Text & ") and (Month(pedidos_1.data_compra) = " & cboMes.ListIndex + 1 & ") And (Year(pedidos_1.data_compra) = " & cboAno & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        SQLWhere = " Where (cliente.codigo = " & txtCodCliente.Text & ") and (Month(pedidos_1.data_compra) = " & cboMes.ListIndex + 1 & ") And (Year(pedidos_1.data_compra) = " & cboAno & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
   
   'PERIODO
    ElseIf cboCriterioPrinc.Text = "PERIODO" And cboCriterioSec.Text = "" Then
        If Not IsDate(mskInicio) Or Not IsDate(mskFim) Then Limpar_Grid_Venda: Exit Sub

        sSQL = sSQL & vTABELAS & " Where (pedidos_1.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_1.data_compra <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        SQLWhere = " Where (pedidos_1.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_1.data_compra <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        
        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO AS pFormaPgto, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
            "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto, " & _
            "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO WHERE (pedidos_1.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_1.data_compra <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = '0') " & TipoPgto & ") AS Nome " & _
            "FROM pedidos AS pedidos_1 INNER JOIN  parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO " & _
            "WHERE (pedidos_1.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_1.data_compra <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = 0) " & TipoPgto & " " & _
            "ORDER BY var_codped"

   'CODIGO
    ElseIf cboCriterioPrinc.Text = "CÓDIGO" And cboCriterioSec.Text = "" Then
        If txtCodigo.Text = "" Then Limpar_Grid_Venda: Exit Sub
        
        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO as pFormaPgto, parcelas.VALOR_FINAL as pValorPgto, parcelas.TIPO_CARTAO, cliente.nome, " & _
        "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto "

        vTABELAS = " FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO"

        sSQL = sSQL & vTABELAS & " Where (pedidos_1.cod_pedido = " & txtCodigo.Text & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        SQLWhere = " Where (pedidos_1.cod_pedido = " & txtCodigo.Text & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "

   'MENSAL
    ElseIf cboCriterioPrinc.Text = "MENSAL" And cboCriterioSec.Text = "" Then
        If cboMes.Text = "" Or cboAno.Text = "" Then Limpar_Grid_Venda: Exit Sub
        
        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
            "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, parcelas.VALOR_FINAL AS pValorPgto, " & _
            "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto, " & _
            "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO WHERE (Month(pedidos_1.data_compra) = " & cboMes.ListIndex + 1 & ") And (Year(pedidos_1.data_compra) = " & cboAno & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = '0') " & TipoPgto & ") AS Nome " & _
            "FROM pedidos AS pedidos_1 INNER JOIN  parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO " & _
            "WHERE (Month(pedidos_1.data_compra) = " & cboMes.ListIndex + 1 & ") And (Year(pedidos_1.data_compra) = " & cboAno & ") AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = 0) " & TipoPgto & " " & vTipoPgtoParcelas & "      " & _
            "ORDER BY var_codped"
            
            'Debug.Print sSQL
   
   'TODOS
    ElseIf cboCriterioPrinc.Text = "TODOS" And cboCriterioSec.Text = "" Then

        sSQL = sSQL & vTABELAS & " Where (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        SQLWhere = " Where (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & " " & TipoPgto & " "
        
        sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
            "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, parcelas.VALOR_FINAL AS pValorPgto, " & _
            "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto, " & _
            "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO WHERE (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = '0') " & TipoPgto & ") AS Nome " & _
            "FROM pedidos AS pedidos_1 INNER JOIN  parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO " & _
            "WHERE (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = 0) " & TipoPgto & " " & _
            "ORDER BY var_codped"
   End If
   
 'SQLCount = "SELECT COUNT (DISTINCT pedidos_1.COD_PEDIDO) AS vQuant " & vTABELAS & " " & SQLWhere & ""
' sSQL = sSQL & " ORDER BY " & INDICE
 
'Debug.Print sSQL

Set r = dbData.OpenRecordset(sSQL, totalRegistros)
lblQtda.Caption = Format(totalRegistros, "00")

FormatarGrid_Vendas r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Double
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
   
'limpar o grid

PreencherPrincipal
cboCriterioPrinc.ListIndex = 5

PreencherIndice
cboIndice.ListIndex = 0

PreencherTipoConsulta
cboTipo.ListIndex = 0

PreencherFormaPgto
cboFormaPgto.ListIndex = 0

PreencherTipoPgto
cboTipoPgto.ListIndex = 0

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

Private Sub mskData_KeyPress(KeyAscii As Integer)
mskData.Mask = "##/##/##"
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
      .Cols = 11
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 900
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 1000
      .ColWidth(8) = 850
      .ColWidth(9) = 850
      .ColWidth(10) = 1300
     
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "SUBTOTAL"
      .TextMatrix(0, 5) = "DESC."
      .TextMatrix(0, 6) = "ACRE."
      .TextMatrix(0, 7) = "VALOR"
      .TextMatrix(0, 8) = "TIPO"
      .TextMatrix(0, 9) = "PARC."
      .TextMatrix(0, 10) = "FORMA"
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
            .TextMatrix(.rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("nome"))
            .TextMatrix(.rows - 1, 4) = Format(rTabela("subtotal"), ocMONEY)
            .TextMatrix(.rows - 1, 5) = Format(rTabela("ValorDescReal"), ocMONEY)
            .TextMatrix(.rows - 1, 6) = Format(rTabela("ValorAcrescReal"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = Format(rTabela("var_total"), ocMONEY)
            .TextMatrix(.rows - 1, 8) = rTabela("tipo_pagamento")
            .TextMatrix(.rows - 1, 9) = Format(rTabela("pvalorPgto"), ocMONEY)
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("pFormaPgto"))
            'ValidateNull(UCase(rTabela("Nome")))
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
         .Col = 7
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
   
    lblSubtotal.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
    lblSubtotalBruto.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
    lblTotalAcrescimo.Caption = Format(SomaGrid(Grid, 6), ocMONEY)
    lblTotalDesconto.Caption = Format(SomaGrid(Grid, 5), ocMONEY)
    'lblEntrada.Caption = Format(0, ocMONEY)
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
