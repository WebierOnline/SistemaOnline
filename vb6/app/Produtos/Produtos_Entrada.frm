VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_Entrada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ENTRADA DE PRODUTOS"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15435
   Icon            =   "Produtos_Entrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   15435
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   60
      TabIndex        =   38
      Top             =   900
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16536
      _Version        =   393216
      Tab             =   1
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
      TabPicture(0)   =   "Produtos_Entrada.frx":23D2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmTransporte"
      Tab(0).Control(1)=   "frmItens"
      Tab(0).Control(2)=   "frmNota"
      Tab(0).Control(3)=   "cmdCancelar"
      Tab(0).Control(4)=   "cmdAlterar"
      Tab(0).Control(5)=   "cmdExcluir"
      Tab(0).Control(6)=   "cmdSalvar"
      Tab(0).Control(7)=   "cmdNovo"
      Tab(0).Control(8)=   "cmdFechar"
      Tab(0).Control(9)=   "cmdImprimirEntrada"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "HISTÓRICO"
      TabPicture(1)   =   "Produtos_Entrada.frx":23EE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Grid_Historico"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CONSULTA"
      TabPicture(2)   =   "Produtos_Entrada.frx":240A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtAliqICMSProd"
      Tab(2).Control(1)=   "txtValorICMSProd"
      Tab(2).Control(2)=   "txtAliqIPIProd"
      Tab(2).Control(3)=   "txtValorIPIProd"
      Tab(2).Control(4)=   "txtAliqICMSSTProd"
      Tab(2).Control(5)=   "txtValorICMSSTProd"
      Tab(2).Control(6)=   "txtCustoLiquido"
      Tab(2).Control(7)=   "Grid"
      Tab(2).Control(8)=   "Data6"
      Tab(2).Control(9)=   "Data5"
      Tab(2).Control(10)=   "Frame9"
      Tab(2).Control(11)=   "cmdImprimir"
      Tab(2).Control(12)=   "cmdExibir"
      Tab(2).Control(13)=   "Label5"
      Tab(2).Control(14)=   "Label6"
      Tab(2).Control(15)=   "Label15"
      Tab(2).Control(16)=   "Label16"
      Tab(2).Control(17)=   "Label18"
      Tab(2).Control(18)=   "Label21"
      Tab(2).Control(19)=   "Label27"
      Tab(2).Control(20)=   "Label26"
      Tab(2).Control(21)=   "lblQuant"
      Tab(2).Control(22)=   "Label9"
      Tab(2).Control(23)=   "lblValor"
      Tab(2).Control(24)=   "Label25"
      Tab(2).ControlCount=   25
      Begin VB.TextBox txtAliqICMSProd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73920
         TabIndex        =   106
         Top             =   8760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtValorICMSProd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -72600
         TabIndex        =   105
         Top             =   8760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtAliqIPIProd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -71280
         TabIndex        =   104
         Top             =   8760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtValorIPIProd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -69960
         TabIndex        =   103
         Top             =   8760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtAliqICMSSTProd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -68640
         TabIndex        =   102
         Top             =   8760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtValorICMSSTProd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67320
         TabIndex        =   101
         Top             =   8760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtCustoLiquido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66000
         TabIndex        =   100
         Top             =   8760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Frame frmTransporte 
         Caption         =   "Transporte:"
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
         Left            =   -74880
         TabIndex        =   65
         Top             =   1380
         Width           =   12735
         Begin VB.TextBox txtCodTransportadora 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6480
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.ComboBox cboTransportadora 
            Height          =   315
            Left            =   2760
            TabIndex        =   11
            Top             =   480
            Width           =   4425
         End
         Begin VB.ComboBox cboTipoFrete 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   2625
         End
         Begin VB.TextBox txtFreteTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7200
            TabIndex        =   12
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label lblTransportadora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transportadora"
            Height          =   195
            Left            =   2760
            TabIndex        =   92
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Frete"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblFreteTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   7200
            TabIndex        =   66
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Frame frmItens 
         Caption         =   "Produtos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6915
         Left            =   -74880
         TabIndex        =   57
         Top             =   2340
         Width           =   12735
         Begin VB.Frame frmPrecos 
            Caption         =   "Preços / Quantidade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5835
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   12555
            Begin VB.TextBox txtValor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   10860
               Locked          =   -1  'True
               TabIndex        =   86
               TabStop         =   0   'False
               Top             =   5400
               Width           =   1635
            End
            Begin VB.Frame Frame5 
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
               Left            =   120
               TabIndex        =   82
               Top             =   240
               Width           =   1395
               Begin VB.TextBox txtQuant 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFC0&
                  Height          =   315
                  Left            =   60
                  TabIndex        =   15
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Quant."
                  Height          =   195
                  Left            =   120
                  TabIndex        =   83
                  Top             =   240
                  Width           =   480
               End
            End
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
               Left            =   3000
               TabIndex        =   79
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtMargemVV 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  TabIndex        =   17
                  Top             =   480
                  Width           =   895
               End
               Begin VB.TextBox txtValorVV 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   18
                  Top             =   480
                  Width           =   795
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton3 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   114
                  Top             =   480
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   ">>"
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
                  MICON           =   "Produtos_Entrada.frx":2426
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
                  Left            =   120
                  TabIndex        =   81
                  Top             =   240
                  Width           =   735
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   80
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
               Left            =   5340
               TabIndex        =   76
               Top             =   240
               Width           =   1995
               Begin VB.TextBox txtValorVP 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   20
                  Top             =   480
                  Width           =   795
               End
               Begin VB.TextBox txtMargemVP 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  TabIndex        =   19
                  Top             =   480
                  Width           =   895
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   78
                  Top             =   240
                  Width           =   360
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Margem %"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   77
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
               Left            =   7380
               TabIndex        =   73
               Top             =   240
               Width           =   1995
               Begin VB.TextBox txtMargemAV 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  TabIndex        =   21
                  Top             =   480
                  Width           =   895
               End
               Begin VB.TextBox txtValorAV 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   22
                  Top             =   480
                  Width           =   795
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Margem %"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   75
                  Top             =   240
                  Width           =   735
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   74
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
               Left            =   9420
               TabIndex        =   70
               Top             =   240
               Width           =   1995
               Begin VB.TextBox txtValorAP 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   24
                  Top             =   480
                  Width           =   795
               End
               Begin VB.TextBox txtMargemAP 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  TabIndex        =   23
                  Top             =   480
                  Width           =   895
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   72
                  Top             =   240
                  Width           =   360
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Margem %"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   71
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
               Left            =   1560
               TabIndex        =   68
               Top             =   240
               Width           =   1395
               Begin VB.TextBox txtCusto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   60
                  TabIndex        =   16
                  Top             =   480
                  Width           =   1275
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Custo Bruto"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   69
                  Top             =   240
                  Width           =   825
               End
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionar 
               Height          =   315
               Left            =   8100
               TabIndex        =   25
               ToolTipText     =   "Adiciona"
               Top             =   1260
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
               MICON           =   "Produtos_Entrada.frx":2442
               PICN            =   "Produtos_Entrada.frx":245E
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
               TabIndex        =   27
               ToolTipText     =   "Remove"
               Top             =   1260
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
               MICON           =   "Produtos_Entrada.frx":27F8
               PICN            =   "Produtos_Entrada.frx":2814
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Cadastro 
               Height          =   3675
               Left            =   60
               TabIndex        =   26
               Top             =   1680
               Width           =   12435
               _ExtentX        =   21934
               _ExtentY        =   6482
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin ChamaleonBtn.chameleonButton cmdCADProdutos 
               Height          =   315
               Left            =   1920
               TabIndex        =   87
               Top             =   5400
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Produtos"
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
               MICON           =   "Produtos_Entrada.frx":2BAE
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
               Height          =   315
               Left            =   60
               TabIndex        =   88
               Top             =   5400
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Fornecedor"
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
               MICON           =   "Produtos_Entrada.frx":2BCA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   4
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
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
               Left            =   3060
               TabIndex        =   84
               Top             =   1140
               Visible         =   0   'False
               Width           =   3300
               WordWrap        =   -1  'True
            End
         End
         Begin VB.TextBox txtCodProduto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7620
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   540
            Width           =   1935
         End
         Begin VB.ComboBox cboDescricao 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2100
            TabIndex        =   14
            Top             =   525
            Width           =   6975
         End
         Begin VB.TextBox txtQuantAtual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   525
            Width           =   1215
         End
         Begin VB.TextBox txtValorAtual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   10380
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   525
            Width           =   1155
         End
         Begin VB.Label lblTipoConsulta 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            Height          =   195
            Left            =   5340
            TabIndex        =   90
            Top             =   240
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblCodFabrica 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de Barra"
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
            TabIndex        =   64
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
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
            Left            =   2100
            TabIndex        =   63
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant. Atual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   9120
            TabIndex        =   62
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Atual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   10380
            TabIndex        =   61
            Top             =   300
            Width           =   945
         End
      End
      Begin VB.Frame frmNota 
         Caption         =   "Dados da Nota"
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
         Left            =   -74880
         TabIndex        =   54
         Top             =   420
         Width           =   12735
         Begin VB.Frame Frame8 
            Caption         =   "Saída"
            Height          =   615
            Left            =   3540
            TabIndex        =   97
            Top             =   240
            Width           =   2055
            Begin ChamaleonBtn.chameleonButton chameleonButton2 
               Height          =   315
               Left            =   960
               TabIndex        =   98
               Tag             =   "Calendario"
               Top             =   240
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
               MICON           =   "Produtos_Entrada.frx":2BE6
               PICN            =   "Produtos_Entrada.frx":2C02
               PICH            =   "Produtos_Entrada.frx":4F55
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskDataSaida 
               Height          =   315
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskHoraSaida 
               Height          =   315
               Left            =   1320
               TabIndex        =   7
               Top             =   240
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Emissão"
            Height          =   615
            Left            =   2160
            TabIndex        =   96
            Top             =   240
            Width           =   1335
            Begin ChamaleonBtn.chameleonButton chameleonButton1 
               Height          =   315
               Left            =   960
               TabIndex        =   5
               Tag             =   "Calendario"
               Top             =   240
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
               MICON           =   "Produtos_Entrada.frx":72A8
               PICN            =   "Produtos_Entrada.frx":72C4
               PICH            =   "Produtos_Entrada.frx":9617
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskDataEmissao 
               Height          =   315
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Cadastro"
            Height          =   615
            Left            =   60
            TabIndex        =   95
            Top             =   240
            Width           =   2055
            Begin ChamaleonBtn.chameleonButton cmdCal1 
               Height          =   315
               Left            =   960
               TabIndex        =   2
               Tag             =   "Calendario"
               Top             =   240
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
               MICON           =   "Produtos_Entrada.frx":B96A
               PICN            =   "Produtos_Entrada.frx":B986
               PICH            =   "Produtos_Entrada.frx":DCD9
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
               TabIndex        =   1
               Top             =   240
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskHora 
               Height          =   315
               Left            =   1320
               TabIndex        =   3
               Top             =   240
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
         End
         Begin VB.TextBox txtCodFornecedor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   10620
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.ComboBox cboFornecedor 
            Height          =   315
            Left            =   6900
            TabIndex        =   9
            Top             =   480
            Width           =   4725
         End
         Begin VB.TextBox txtNotaFiscal 
            Height          =   315
            Left            =   5700
            TabIndex        =   8
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
            Height          =   195
            Left            =   6900
            TabIndex        =   56
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota Fiscal"
            Height          =   195
            Left            =   5700
            TabIndex        =   55
            Top             =   255
            Width           =   795
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Historico 
         Height          =   7215
         Left            =   120
         TabIndex        =   44
         Top             =   420
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   12726
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   43
         Top             =   420
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   9763
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
         ForeColor       =   &H000000C0&
         Height          =   1275
         Left            =   -74880
         TabIndex        =   40
         Top             =   6780
         Width           =   12615
         Begin VB.ComboBox cboOrdem 
            Height          =   315
            ItemData        =   "Produtos_Entrada.frx":1002C
            Left            =   2280
            List            =   "Produtos_Entrada.frx":1002E
            TabIndex        =   34
            Top             =   480
            Width           =   1935
         End
         Begin VB.ComboBox cboConsulta 
            Height          =   315
            ItemData        =   "Produtos_Entrada.frx":10030
            Left            =   120
            List            =   "Produtos_Entrada.frx":10032
            TabIndex        =   35
            Top             =   480
            Width           =   2115
         End
         Begin VB.ComboBox cboConsDescricao 
            Height          =   315
            ItemData        =   "Produtos_Entrada.frx":10034
            Left            =   4260
            List            =   "Produtos_Entrada.frx":10036
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cboConsAno 
            Height          =   315
            Left            =   6120
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   480
            Visible         =   0   'False
            Width           =   2115
         End
         Begin ChamaleonBtn.chameleonButton cmdConNotaCal1 
            Height          =   315
            Left            =   5280
            TabIndex        =   116
            Tag             =   "Calendario"
            Top             =   480
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
            MICON           =   "Produtos_Entrada.frx":10038
            PICN            =   "Produtos_Entrada.frx":10054
            PICH            =   "Produtos_Entrada.frx":123A7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdConNotaCal2 
            Height          =   315
            Left            =   6540
            TabIndex        =   117
            Tag             =   "Calendario"
            Top             =   480
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
            MICON           =   "Produtos_Entrada.frx":146FA
            PICN            =   "Produtos_Entrada.frx":14716
            PICH            =   "Produtos_Entrada.frx":16A69
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskFinal 
            Height          =   315
            Left            =   5640
            TabIndex        =   118
            Top             =   480
            Visible         =   0   'False
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskInicial 
            Height          =   315
            Left            =   4260
            TabIndex        =   119
            Top             =   480
            Visible         =   0   'False
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin VB.Label lblFinal 
            AutoSize        =   -1  'True
            Caption         =   "Final"
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
            Left            =   5640
            TabIndex        =   120
            Top             =   240
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Organizar:"
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
            Left            =   2280
            TabIndex        =   48
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Critério:"
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
            TabIndex        =   47
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblConsDescricao 
            Caption         =   "Descrição"
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
            Left            =   4320
            TabIndex        =   46
            Top             =   240
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   -62040
         TabIndex        =   29
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
         MICON           =   "Produtos_Entrada.frx":18DBC
         PICN            =   "Produtos_Entrada.frx":18DD8
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
         Left            =   -62040
         TabIndex        =   30
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
         MICON           =   "Produtos_Entrada.frx":1AB6A
         PICN            =   "Produtos_Entrada.frx":1AB86
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
         Left            =   -62040
         TabIndex        =   31
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
         MICON           =   "Produtos_Entrada.frx":1C918
         PICN            =   "Produtos_Entrada.frx":1C934
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
         Left            =   -62040
         TabIndex        =   28
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
         MICON           =   "Produtos_Entrada.frx":1E6C6
         PICN            =   "Produtos_Entrada.frx":1E6E2
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
         Left            =   -62040
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
         MICON           =   "Produtos_Entrada.frx":20474
         PICN            =   "Produtos_Entrada.frx":20490
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdFechar 
         Height          =   615
         Left            =   -62040
         TabIndex        =   33
         Top             =   8580
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
         MICON           =   "Produtos_Entrada.frx":22222
         PICN            =   "Produtos_Entrada.frx":2223E
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
         Height          =   615
         Left            =   -62040
         TabIndex        =   32
         Top             =   3840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "Produtos_Entrada.frx":23FD0
         PICN            =   "Produtos_Entrada.frx":23FEC
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
         Left            =   -62220
         TabIndex        =   89
         Top             =   7440
         Width           =   2355
         _ExtentX        =   4154
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
         MICON           =   "Produtos_Entrada.frx":25D7E
         PICN            =   "Produtos_Entrada.frx":25D9A
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
         Height          =   555
         Left            =   -62220
         TabIndex        =   99
         Top             =   6840
         Width           =   2355
         _ExtentX        =   4154
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
         MICON           =   "Produtos_Entrada.frx":27B2C
         PICN            =   "Produtos_Entrada.frx":27B48
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aliq. ICMS"
         Height          =   195
         Left            =   -73920
         TabIndex        =   113
         Top             =   8520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor ICMS"
         Height          =   195
         Left            =   -72600
         TabIndex        =   112
         Top             =   8520
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aliq. IPI"
         Height          =   195
         Left            =   -71280
         TabIndex        =   111
         Top             =   8520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor IPI"
         Height          =   195
         Left            =   -69960
         TabIndex        =   110
         Top             =   8520
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aliq. ICMS ST"
         Height          =   195
         Left            =   -68580
         TabIndex        =   109
         Top             =   8520
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor ICMS ST"
         Height          =   195
         Left            =   -67260
         TabIndex        =   108
         Top             =   8520
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo Liquido"
         Height          =   195
         Left            =   -66000
         TabIndex        =   107
         Top             =   8520
         Visible         =   0   'False
         Width           =   960
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
         Left            =   -64740
         TabIndex        =   52
         Top             =   6000
         Width           =   735
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
         Left            =   -63960
         TabIndex        =   51
         Top             =   6000
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
         Left            =   -62400
         TabIndex        =   50
         Top             =   6000
         Width           =   795
      End
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
         Left            =   -61500
         TabIndex        =   49
         Top             =   6000
         Width           =   1605
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Dê um duplo-clique para ver mais informações"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74820
         TabIndex        =   39
         Top             =   6000
         Width           =   3255
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   45
      Top             =   10320
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22886
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "19:26"
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
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   15285
      TabIndex        =   41
      Top             =   60
      Width           =   15315
      Begin VB.TextBox txtCodUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9000
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   6540
         TabIndex        =   85
         Top             =   240
         Visible         =   0   'False
         Width           =   1155
      End
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
         Left            =   13620
         TabIndex        =   53
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTRADA DE PRODUTOS"
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
         TabIndex        =   42
         Top             =   240
         Width           =   3885
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Produtos_Entrada.frx":298DA
         Top             =   60
         Width           =   645
      End
   End
End
Attribute VB_Name = "Produtos_Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Private moCombo As cComboHelper
Private printSQL As String
Public Campo As Integer

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

Private Sub DesativarBotoes()
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdImprimirEntrada.Enabled = False
frmNota.Enabled = False
frmTransporte.Enabled = False
frmItens.Enabled = False
End Sub

Private Sub FormatarGrid_Historico(rTabela As ADODB.Recordset)
   Dim i As Integer, x As Integer
   
   With Grid_Historico
      .Clear
      .Cols = 6
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 1000
      .ColWidth(2) = 1300
      .ColWidth(3) = 5300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      
      For x = 0 To .Cols - 1
         .Col = x
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "CADASTRO"
      .TextMatrix(0, 2) = "NOTA FISCAL"
      .TextMatrix(0, 3) = "FORNECEDOR"
      .TextMatrix(0, 4) = "FRETE"
      .TextMatrix(0, 5) = "VALOR"
      
      .Redraw = False
      i = 1
      
      If Not rTabela Is Nothing Then
        Do While Not rTabela.EOF
           .TextMatrix(.rows - 1, 1) = Format(rTabela("data_entrada"), "dd/mm/yy")
           .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("notafiscal"))
           .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("varfornecedor"))
           .TextMatrix(.rows - 1, 4) = Format$(rTabela("valor_frete"), ocMONEY)
           .TextMatrix(.rows - 1, 5) = Format$(rTabela("valor"), ocMONEY)
           
           rTabela.MoveNext
           .rows = .rows + 1
           i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 5
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      .Redraw = True
   End With
End Sub


Private Sub Limpar_Objetos()
If cmdAlterar.Enabled = False Then txtCodigo.Text = ""
mskData.Mask = ""
mskData.Text = ""
mskHora.Mask = ""
mskHora.Text = ""
cboFornecedor.Text = ""
TxtCodFornecedor.Text = ""
txtQuant.Text = ""
txtNotaFiscal.Text = ""
txtValor.Text = ""
txtFreteTotal.Text = ""
mskDataEmissao.Mask = ""
mskDataEmissao.Text = ""
mskDataSaida.Mask = ""
mskDataSaida.Text = ""
mskHoraSaida.Mask = ""
mskHoraSaida.Text = ""
cboTipoFrete.Text = ""
cboTransportadora.Text = ""
txtCodTransportadora.Text = ""
txtFreteTotal.Text = ""
End Sub

Private Sub Limpar_SubDados()
txtCodBarra.Text = ""
cboDescricao.Text = ""
'cboDescricao2.Text = ""
txtCodProduto.Text = ""
txtQuant.Text = ""
txtQuantAtual.Text = ""
txtValorAtual.Text = ""
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
   sSQL_Itens = "SELECT produtos_entrada_itens.*, produtos_entrada_itens.codigo as varCod,  produtos.COD_BARRA as var_CodBarra, produtos.REF as var_REF, produtos.codigo, produtos.descricao as varDesc, produtos.tamanho, produtos.fabricante, produtos_entrada_itens.CUSTO as var_custo, (produtos_entrada_itens.CUSTO * produtos_entrada_itens.QUANT) as varTotalCustoItem, produtos_entrada_itens.VALOR_VV as var_venda  " & _
          " FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CODIGO_PRODUTO " & _
          " WHERE (codigo_entrada = " & txtCodigo.Text & ") ORDER BY varDesc, TAMANHO, REF;"
   Set r = dbData.OpenRecordset(sSQL_Itens)
End If

printSQL = sSQL_Itens

FormatarGrid_Itens r
txtValor.Text = FormatNumber(SomaGrid(Grid_Cadastro, 16), 2)

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Total()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vlrSoma As Currency

If txtCodigo.Text = "" Then Exit Sub
vlrSoma = 0

'somar custo
sSQL = "SELECT ISNULL(SUM(custo_compra * quant), 0) AS var_soma_custo FROM produtos_entrada_itens WHERE (codigo_entrada = " & txtCodigo.Text & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then vlrSoma = r("var_soma_custo")
If r.State <> 0 Then r.Close
Set r = Nothing

txtValor.Text = Format$(vlrSoma, ocMONEY)
'txt_NumItens.Text = SomaGrid(Grid_Cadastro, 6)   'SOMAR A QUANTIDADE
End Sub

Private Sub MostrarValorVenda()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vrVenda As Currency
If txtCodProduto.Text = "" Then Exit Sub

'mostrar o ultimo preço de compra
sSQL = "SELECT TOP 1 VALOR_VV FROM Produtos_Precos WHERE (cod_produto = " & txtCodProduto & ") ORDER BY codigo DESC;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then vrVenda = r("VALOR_VV")
If r.State <> 0 Then r.Close
Set r = Nothing

txtValorAtual.Text = Format(vrVenda, ocMONEY)
End Sub



Private Sub cboConsAno_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
        


If cboConsulta.Text = "DETALHADO" Then
    cboConsAno.Clear
    
    If cboConsDescricao.Text = "FABRICANTE" Then
        sSQL = "SELECT DISTINCT fabricante FROM produtos ORDER BY fabricante;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("fabricante"))
           r.MoveNext
        Loop
    ElseIf cboConsDescricao.Text = "REFERÊNCIA" Then
        sSQL = "SELECT DISTINCT REF FROM produtos ORDER BY REF;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("REF"))
           r.MoveNext
        Loop
    ElseIf cboConsDescricao.Text = "LINHA" Then
        sSQL = "SELECT DISTINCT CATEGORIA FROM produtos ORDER BY CATEGORIA;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("CATEGORIA"))
           r.MoveNext
        Loop
    ElseIf cboConsDescricao.Text = "TAMANHO" Then
        sSQL = "SELECT DISTINCT TAMANHO FROM produtos ORDER BY TAMANHO;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("TAMANHO"))
           r.MoveNext
        Loop
    ElseIf cboConsDescricao.Text = "COD. BARRA" Then
        sSQL = "SELECT DISTINCT COD_BARRA FROM produtos ORDER BY COD_BARRA;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("COD_BARRA"))
           r.MoveNext
        Loop
    End If
End If

    If cboConsulta.Text = "MENSAL" Or cboConsulta.Text = "DETALHADO + MENSAL" Then
        Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
        Dim i As Integer
        
        'Calcula o intervalo de anos
        iAno = Year(Date)
        FirstYear = iAno - 2
        LastYear = iAno + 2
        
        cboConsAno.Clear
        
        For i = FirstYear To LastYear
           cboConsAno.AddItem i
        Next
    End If


moCombo.AttachTo cboConsAno
End Sub


Private Sub cboConsDescricao_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vMes As Integer

If cboConsulta.Text = "PRODUTO" Then
   cboConsDescricao.Clear
   
   sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsDescricao.AddItem ValidateNull(r("descricao"))
      cboConsDescricao.ItemData(cboConsDescricao.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   moCombo.AttachTo cboConsDescricao

ElseIf cboConsulta.Text = "FABRICANTE" Then
   cboConsDescricao.Clear
   
   sSQL = "SELECT DISTINCT fabricante FROM produtos ORDER BY fabricante;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsDescricao.AddItem ValidateNull(r("fabricante"))
      r.MoveNext
   Loop
   
   moCombo.AttachTo cboConsDescricao

ElseIf cboConsulta.Text = "REFERÊNCIA" Then
   cboConsDescricao.Clear
   
   sSQL = "SELECT DISTINCT ref FROM produtos ORDER BY ref;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsDescricao.AddItem ValidateNull(r("ref"))
      r.MoveNext
   Loop
   
   moCombo.AttachTo cboConsDescricao

   
ElseIf cboConsulta.Text = "FORNECEDOR" Then
   cboConsDescricao.Clear
   
   sSQL = "SELECT DISTINCT razao FROM fornecedor ORDER BY razao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsDescricao.AddItem r("razao")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   
   moCombo.AttachTo cboConsDescricao
ElseIf cboConsulta.Text = "MENSAL" Then

   
   cboConsDescricao.Clear
   
   For vMes = 1 To 12
      cboConsDescricao.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboConsDescricao
ElseIf cboConsulta.Text = "DETALHADO" Then
    cboConsDescricao.Clear
    cboConsDescricao.AddItem "FABRICANTE"
    cboConsDescricao.AddItem "REFERENCIA"
    cboConsDescricao.AddItem "LINHA"
    cboConsDescricao.AddItem "TAMANHO"
    cboConsDescricao.AddItem "COD. BARRA"

ElseIf cboConsulta.Text = "DETALHADO + MENSAL" Then
   cboConsDescricao.Clear
   
   For vMes = 1 To 12
      cboConsDescricao.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboConsDescricao
End If
End Sub
Private Sub cboConsDescricao_LostFocus()
If cboConsulta.Text = "MENSAL" Then
    If cboConsDescricao.Text = "" Then Exit Sub Else cboConsAno.SetFocus
End If
End Sub


Private Sub cboConsulta_Click()
cboConsulta_LostFocus
End Sub

Private Sub cboConsulta_GotFocus()
cboConsulta.Clear
cboConsulta.AddItem "TODOS"
cboConsulta.AddItem "MENSAL"
cboConsulta.AddItem "PERÍODO"
cboConsulta.AddItem "NOTA FISCAL"
cboConsulta.AddItem "FORNECEDOR"
cboConsulta.AddItem "PRODUTO"
cboConsulta.AddItem "CÓD. BARRA"
cboConsulta.AddItem "REFERÊNCIA"
cboConsulta.AddItem "FABRICANTE"

moCombo.AttachTo cboConsulta
End Sub


Private Sub cboConsulta_LostFocus()
If cboConsulta.Text = "TODOS" Then
    cboConsDescricao.Visible = False
    cboConsAno.Visible = False
    lblConsDescricao.Visible = False
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "PERÍODO" Then
    cboConsDescricao.Visible = False
    cboConsAno.Visible = False
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Inicial"
    mskInicial.Visible = True
    mskFinal.Visible = True
    cmdConNotaCal1.Visible = True
    cmdConNotaCal2.Visible = True
    lblFinal.Visible = True
ElseIf cboConsulta.Text = "NOTA FISCAL" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = False
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Nota Fiscal"
    cboConsDescricao.Width = 4095
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "FORNECEDOR" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = False
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Fornecedor"
    cboConsDescricao.Width = 4095
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "MENSAL" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = True
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Mês/Ano"
    cboConsDescricao.Width = 1815
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "PRODUTO" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = False
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Produto"
    cboConsDescricao.Width = 4095
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "FABRICANTE" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = False
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Fabricante"
    cboConsDescricao.Width = 4095
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "REFERÊNCIA" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = False
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Referência"
    cboConsDescricao.Width = 4095
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "CÓD. BARRA" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = False
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Cód. Barra"
    cboConsDescricao.Width = 4095
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "DETALHADO" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = True
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Critério"
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
ElseIf cboConsulta.Text = "DETALHADO + MENSAL" Then
    cboConsDescricao.Visible = True
    cboConsAno.Visible = True
    lblConsDescricao.Visible = True
    lblConsDescricao.Caption = "Mês/Ano"
    cboConsDescricao.Width = 1815
    mskInicial.Visible = False
    mskFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    lblFinal.Visible = False
End If

cboConsDescricao.Clear
End Sub


Private Sub cboDescricao_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

'Limpa a lista
If lblTipoConsulta.Caption = "0" Or lblTipoConsulta.Caption = "2" Then
    
    If cboDescricao.ListIndex = -1 Then
        cboDescricao.Clear
        
        sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboDescricao.AddItem ValidateNull(r("descricao"))
            cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
           r.MoveNext
        Loop
    End If
End If
moCombo.AttachTo cboDescricao
End Sub

Private Sub cboDescricao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboDescricao_LostFocus()
If lblTipoConsulta.Caption = "0" Or lblTipoConsulta.Caption = "2" Then
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cboDescricao.Text = "" Then txtCodProduto.Text = "": lblTipoConsulta.Caption = "0": txtCodBarra.Locked = False: txtCodBarra.Text = "": Exit Sub
   If cboDescricao.ListIndex = -1 Then txtCodProduto.Text = "": lblTipoConsulta.Caption = "0": txtCodBarra.Locked = False: cboDescricao.Text = "": txtCodBarra.Text = "": Exit Sub

   txtCodProduto = cboDescricao.ItemData(cboDescricao.ListIndex)
   
   If txtCodProduto.Text = "" Then lblTipoConsulta.Caption = "0": txtCodBarra.Locked = False: cboDescricao.Text = "": txtCodBarra.Text = "": Exit Sub
   
   sSQL = "SELECT codigo, descricao, cod_barra, quant_estoque FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      'cboDescricao2.Text = ValidateNull(r("descricao"))
      If txtCodBarra.Text = "" Then txtCodBarra.Text = r("cod_barra")
      txtQuantAtual.Text = ValidateNull(r("quant_estoque"))
      lblTipoConsulta.Caption = "2"
      txtCodBarra.Locked = True
   Else
       ShowMsg "Produto não cadastrado.", vbExclamation
       lblTipoConsulta.Caption = "0"
       cboDescricao.Text = ""
       txtCodBarra.Text = ""
       txtCodBarra.Locked = False
   End If

    MostrarValorVenda
    txtQuant.SetFocus

   If r.BOF Then ShowMsg "Produto não cadastrado.", vbExclamation
   If r.State <> 0 Then r.Close
End If
End Sub



Private Sub cboFornecedor_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim var_cboTexto As String

sSQL = "SELECT DISTINCT razao, codigo FROM fornecedor;"
Set r = dbData.OpenRecordset(sSQL)

If cboFornecedor.Text <> "" Then var_cboTexto = cboFornecedor.Text
cboFornecedor.Clear
cboFornecedor.Text = var_cboTexto

Do While Not r.EOF
   cboFornecedor.AddItem r("razao")
   cboFornecedor.ItemData(cboFornecedor.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

SelectControl cboFornecedor
moCombo.AttachTo cboFornecedor
End Sub

Private Sub cboFornecedor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFornecedor_LostFocus()
If cboFornecedor.Text = "" Then TxtCodFornecedor.Text = "": Exit Sub
If cboFornecedor.ListIndex = -1 Then TxtCodFornecedor.Text = "": MsgBox "Fornecedor inexistente!!", vbCritical: cboFornecedor.Text = "": Exit Sub
TxtCodFornecedor = cboFornecedor.ItemData(cboFornecedor.ListIndex)
Exit Sub
End Sub


Private Sub cboOrdem_GotFocus()
    cboOrdem.Clear
    cboOrdem.AddItem "DATA"
    cboOrdem.AddItem "NUM DA NOTA"
    cboOrdem.AddItem "VALOR"
    cboOrdem.AddItem "FORNECEDOR"
    moCombo.AttachTo cboOrdem
End Sub





Private Sub cboTipoFrete_Change()
cboTipoFrete_LostFocus
End Sub

Private Sub cboTipoFrete_GotFocus()
cboTipoFrete.Clear
cboTipoFrete.AddItem "0 - EMITENTE"
cboTipoFrete.AddItem "1 - DESTINATARIO"
moCombo.AttachTo cboTipoFrete
End Sub


Private Sub cboTipoFrete_LostFocus()
If cboTipoFrete.Text = "0 - EMITENTE" Then
    lblTransportadora.Enabled = False
    cboTransportadora.Enabled = False
    lblFreteTotal.Enabled = False
    txtFreteTotal.Enabled = False
ElseIf cboTipoFrete.Text = "1 - DESTINATARIO" Then
    lblTransportadora.Enabled = True
    cboTransportadora.Enabled = True
    lblFreteTotal.Enabled = True
    txtFreteTotal.Enabled = True
    cboTransportadora.SetFocus
End If
End Sub


Private Sub cboTipoFrete_Validate(Cancel As Boolean)
cboTipoFrete_LostFocus
End Sub

Private Sub cboTransportadora_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim var_cboTexto As String

sSQL = "SELECT DISTINCT razao, codigo FROM transportadora;"
Set r = dbData.OpenRecordset(sSQL)

If cboTransportadora.Text <> "" Then var_cboTexto = cboTransportadora.Text
cboTransportadora.Clear
cboTransportadora.Text = var_cboTexto

Do While Not r.EOF
   cboTransportadora.AddItem r("razao")
   cboTransportadora.ItemData(cboTransportadora.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

SelectControl cboTransportadora
moCombo.AttachTo cboTransportadora
End Sub

Private Sub cboTransportadora_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboTransportadora_LostFocus()
If cboTransportadora.Text = "" Then txtCodTransportadora.Text = "": Exit Sub
If cboTransportadora.ListIndex = -1 Then txtCodTransportadora.Text = "": MsgBox "Transportadora inexistente!!", vbCritical: cboTransportadora.Text = "": Exit Sub
txtCodTransportadora = cboTransportadora.ItemData(cboTransportadora.ListIndex)
Exit Sub
End Sub


Private Sub chameleonButton1_Click()
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

mskDataEmissao = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub chameleonButton2_Click()
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

mskDataSaida = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub chameleonButton3_Click()
txtMargemVP.Text = txtMargemVV.Text
txtMargemAV = txtMargemVV.Text
txtMargemAP = txtMargemVV.Text
CalcularPrecos
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

mskData = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdConNotaCal1_Click()
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

mskInicial = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdConNotaCal2_Click()
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

mskFinal = Format(varData, "dd/mm/yy")   'Exibe a data no campo
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
REL_Prod_Entrada.dfTotal.Caption = txtValor.Text
REL_Prod_Entrada.Relatorio.Ativar
Unload REL_Prod_Entrada
End Sub



 
Private Sub Command1_Click()

  Dim sSQL As String
   Dim r As ADODB.Recordset

      Dim dIni As Date, dFim As Date
      Dim strData As String
      Dim DIA As Date
      Dim pInd As Integer
      
      Dim saldoInicial As Double
      Dim rEstoque() As String
      
      Dim saldoDia As Double
      Dim totalEntr As Double
      Dim totalSaida As Double
      
      If Not ExistInList(cboConsDescricao) Then
         ShowMsg "Selecione o mês na lista.", vbExclamation
         Exit Sub
      End If
      
      If Not ExistInList(cboConsAno) Then
         ShowMsg "Selecione o ano na lista.", vbExclamation
         Exit Sub
      End If
      
      'Período da pesquisa
      strData = "01/" & Format$(cboConsDescricao.ListIndex + 1, "00") & "/" & Format$(cboConsAno, "0000")
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
         "FROM produtos WHERE (codigo = " & cboConsDescricao.ItemData(cboConsDescricao.ListIndex) & ");"

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
            "FROM produtos WHERE (codigo = " & cboConsDescricao.ItemData(cboConsDescricao.ListIndex) & ");"

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
      
      lblQuant.Caption = SomaGrid(Grid, 4)
      lblValor.Caption = Format(SomaGrid(Grid, 5), "##,##0.00")
End Sub







Private Sub lblTipoConsulta_Change()
If lblTipoConsulta.Caption = "1" Then
    txtCodBarra.BackColor = &HC0FFFF
    cboDescricao.BackColor = &HFFFFFF
ElseIf lblTipoConsulta.Caption = "2" Then
    txtCodBarra.BackColor = &HFFFFFF
    cboDescricao.BackColor = &HC0FFFF
Else
    txtCodBarra.BackColor = &HFFFFFF
    cboDescricao.BackColor = &HFFFFFF
End If
End Sub

Private Sub cmdAdicionar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub
If txtCodProduto.Text = "" Then Exit Sub

If txtQuant.Text = "" Or txtQuant.Text = "0" Then
   ShowMsg "Insira uma quantidade válida!", vbExclamation
   txtQuant.SetFocus
   Exit Sub
End If

If txtCusto.Text = "" Or txtCusto.Text = "0" Then
   ShowMsg "Insira um custo de produto válido!", vbExclamation
   txtCusto.SetFocus
   Exit Sub
End If

Dim var_COD_ITENS As Long

'AUTONUMERAÇÃO
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_entrada_itens;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then var_COD_ITENS = r("cod_itens") + 1
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

sSQL = "INSERT INTO produtos_entrada_itens (" & _
   "codigo, " & _
   "codigo_entrada, " & _
   "codigo_produto, " & _
   "descricao, " & _
   "quant, " & _
   "custo, " & _
   "MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP ) VALUES (" & _
   var_COD_ITENS & ", " & txtCodigo.Text & ", " & txtCodProduto.Text & ", '" & cboDescricao.Text & "', " & _
   Replace(CDbl(txtQuant.Text), ",", ".") & ", " & Replace(CCur(txtCusto.Text), ",", ".") & ", " & _
   Replace(CDbl(varMargemVV), ",", ".") & ", " & Replace(CCur(txtValorVV.Text), ",", ".") & ", " & Replace(CDbl(varMargemVP), ",", ".") & ", " & Replace(CCur(txtValorVP.Text), ",", ".") & ", " & Replace(CDbl(varMargemAV), ",", ".") & ", " & Replace(CCur(txtValorAV.Text), ",", ".") & ", " & Replace(CDbl(varMargemAP), ",", ".") & ", " & Replace(CCur(txtValorAP.Text), ",", ".") & ")"

'Adiciona o registro
dbData.Execute sSQL

Preco_Entrada
Quant_Entrada

'Atualiza o estoque do produto
dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(txtQuant.Text, ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"

Limpar_SubDados
Limpar_Valores
Mostrar_Itens
txtValor.Text = Format(SomaGrid(Grid_Cadastro, 16), "##,##0.000")

On Local Error Resume Next
txtCodBarra.SetFocus
End Sub
Private Sub Quant_Entrada()
Dim sSQL As String
Dim r As ADODB.Recordset

'ENTRADA DO PRODUTO
'If cmdSalvar.Enabled = True Then
   Dim AutoNumeracao As Long
   
   'AUTONUMERAÇÃO
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_quant;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   sSQL = "INSERT INTO produtos_quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO, HORA, COD_USUARIO, ESTOQUE) VALUES (" & _
      AutoNumeracao & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & txtCodigo.Text & ", 'ENTRADA', " & Replace(CDbl(txtQuant.Text), ",", ".") & ", 'ADIÇÃO', '" & Format(Now, ocHRMN) & "', " & txtCodUsuario.Text & ", " & Replace(CDbl(txtQuantAtual.Text), ",", ".") & ");"
   dbData.Execute sSQL
'End If
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

Limpar_Objetos
Limpar_SubDados
Limpar_Valores
txtCodigo.Text = ""
Mostrar_Itens
cmdExibir_Click
DesativarBotoes
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
         Produtos_Cadastro.Show 1
   'End Select
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer

If txtCodigo.Text = "" Then Exit Sub

If cmdAlterar.Enabled = True Then
    MsgBox "Entrada já gravada não pode ser cancelada!", vbInformation, "Aviso do Sistema"
Else
    If ShowMsg("Existe uma nota fiscal em aberto. Deseja sair e cancelar a entrada?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
       'Cancel = 1
       Exit Sub
    End If
    
    With Grid_Cadastro
       For i = 1 To .rows - 1
          dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo_entrada = " & txtCodigo.Text & ");"
          dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & CDbl(.TextMatrix(i, 6)) & " WHERE (codigo = " & CLng(.TextMatrix(i, 4)) & ");"
          dbData.Execute "DELETE FROM produtos_entrada WHERE (codigo = " & txtCodigo.Text & ");"
          dbData.Execute "DELETE FROM Produtos_Precos WHERE (cod_entrada = " & txtCodigo.Text & ");"
          dbData.Execute "DELETE FROM Produtos_quant WHERE (cod_entrada = " & txtCodigo.Text & ");"
       Next
    End With
    
    Limpar_Objetos
    Limpar_SubDados
    Limpar_Valores
    Mostrar_Itens
    
    
    DesativarBotoes
    lblTipoConsulta.Caption = "0"
End If
End Sub

Private Sub Preco_Entrada()
Dim sSQL As String
Dim r As ADODB.Recordset

'ENTRADA DO PRODUTO
'If cmdSalvar.Enabled = True Then
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
   
   sSQL = "INSERT INTO produtos_precos (Codigo, COD_PRODUTO, Data, FORMA, MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP, CUSTO, COD_ENTRADA) VALUES (" & _
      AutoNumeracao & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103), 'ENTRADA', " & Replace(CDbl(varMargemVV), ",", ".") & ", " & Replace(CCur(txtValorVV.Text), ",", ".") & ", " & Replace(CDbl(varMargemVP), ",", ".") & ", " & Replace(CCur(txtValorVP.Text), ",", ".") & ", " & Replace(CDbl(varMargemAV), ",", ".") & ", " & Replace(CCur(txtValorAV.Text), ",", ".") & ", " & Replace(CDbl(varMargemAP), ",", ".") & ", " & Replace(CCur(txtValorAP.Text), ",", ".") & ", " & Replace(CCur(txtCusto.Text), ",", ".") & ", " & txtCodigo.Text & ");"
   dbData.Execute sSQL
'End If
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
   For i = 1 To .rows - 1
        dbData.Execute "DELETE FROM Produtos_Quant WHERE (COD_PRODUTO = " & CLng(.TextMatrix(i, 4)) & ") AND (COD_ENTRADA = " & txtCodigo.Text & ");"
        dbData.Execute "DELETE FROM Produtos_Precos WHERE (COD_PRODUTO = " & CLng(.TextMatrix(i, 4)) & ") AND (COD_ENTRADA = " & txtCodigo.Text & ");"
        dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo_entrada = " & CLng(.TextMatrix(i, 2)) & ");"
        dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & CDbl(.TextMatrix(i, 6)) & ", ult_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) WHERE (codigo = " & CLng(.TextMatrix(i, 4)) & ");"
        dbData.Execute "DELETE FROM produtos_entrada WHERE (codigo = " & txtCodigo.Text & ");"
   Next
End With

Limpar_Objetos
Limpar_SubDados
Limpar_Valores
Mostrar_Itens
cmdExibir_Click
   
DesativarBotoes
End Sub

Private Sub cmdExibir_Click()
'INDICE PARA ORGANIZAR OS DADOS
Dim INDICE As String
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Dim fExibir As Integer

If cboConsulta.Text = "" Then Exit Sub

'Seleciona a ordem dos registros
If cboOrdem.Text = "NUM DA NOTA" Then
    INDICE = "notafiscal;"
ElseIf cboOrdem.Text = "DATA" Then
    INDICE = "data_entrada;"
ElseIf cboOrdem.Text = "VALOR" Then
    INDICE = "valor;"
ElseIf cboOrdem.Text = "FORNECEDOR" Then
    INDICE = "fornecedor;"
Else
    INDICE = "notafiscal;"
End If
   
'Seleciona os registros
fExibir = 0

If cboConsulta.Text = "TODOS" Then
   sSQL = "SELECT fornecedor.codigo, fornecedor.razao, produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML " & _
   "FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo " & _
   "ORDER BY " & INDICE
ElseIf cboConsulta.Text = "NOTA FISCAL" Then
  If cboConsDescricao.Text = "" Then Exit Sub
   sSQL = "SELECT fornecedor.codigo, fornecedor.razao, produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML " & _
          "FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo " & _
          "WHERE (notafiscal = " & cboConsDescricao.Text & ") " & _
          "ORDER BY " & INDICE
ElseIf cboConsulta.Text = "FORNECEDOR" Then
  If cboConsDescricao.Text = "" Then Exit Sub
   sSQL = "SELECT fornecedor.codigo, fornecedor.razao, produtos_entrada.codigo AS var_codent,  produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML " & _
          "FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo " & _
          "WHERE (razao = '" & cboConsDescricao.Text & "') " & _
          "ORDER BY " & INDICE
          
ElseIf cboConsulta.Text = "PRODUTO" Then
  If cboConsDescricao.Text = "" Then Exit Sub
   sSQL = "SELECT DISTINCT produtos_entrada.notafiscal, produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML, produtos_entrada_itens.* " & _
          "FROM produtos_entrada_itens INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo INNER JOIN produtos ON produtos_entrada_itens.codigo_produto = produtos.codigo " & _
          "WHERE (produtos_entrada_itens.descricao = '" & cboConsDescricao.Text & "') " & _
          "ORDER BY " & INDICE

ElseIf cboConsulta.Text = "CÓD. BARRA" Then
  If cboConsDescricao.Text = "" Then Exit Sub
   sSQL = "SELECT produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML, produtos_entrada_itens.* " & _
          "FROM produtos_entrada_itens INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo INNER JOIN produtos ON produtos_entrada_itens.codigo_produto = produtos.codigo " & _
          "WHERE (cod_barra = '" & cboConsDescricao.Text & "') " & _
          "ORDER BY " & INDICE
          
ElseIf cboConsulta.Text = "REFERÊNCIA" Then
  If cboConsDescricao.Text = "" Then Exit Sub
   sSQL = "SELECT produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML, produtos_entrada_itens.* " & _
          "FROM produtos_entrada_itens INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo INNER JOIN produtos ON produtos_entrada_itens.codigo_produto = produtos.codigo " & _
          "WHERE (ref = '" & cboConsDescricao.Text & "') " & _
          "ORDER BY " & INDICE
          
ElseIf cboConsulta.Text = "FABRICANTE" Then
  If cboConsDescricao.Text = "" Then Exit Sub
   sSQL = "SELECT produtos_entrada.codigo AS var_codent,  produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML, produtos_entrada_itens.* " & _
          "FROM produtos_entrada_itens INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo INNER JOIN produtos ON produtos_entrada_itens.codigo_produto = produtos.codigo " & _
          "WHERE (fabricante = '" & cboConsDescricao.Text & "') " & _
          "ORDER BY " & INDICE

ElseIf cboConsulta.Text = "MENSAL" Then
   If Not ExistInList(cboConsDescricao) Then
      ShowMsg "Selecione o mês na lista.", vbExclamation
      Exit Sub
   End If
   
   If Not ExistInList(cboConsAno) Then
      ShowMsg "Selecione o ano na lista.", vbExclamation
      Exit Sub
   End If
   If cboConsDescricao.Text = "" Then Exit Sub
   sSQL = "SELECT fornecedor.codigo, fornecedor.razao, produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML  " & _
          "FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo " & _
          "WHERE (MONTH(data_entrada) = " & cboConsDescricao.ListIndex + 1 & ") AND (YEAR(data_entrada) = " & cboConsAno & ")  " & _
          "ORDER BY " & INDICE
ElseIf cboConsulta.Text = "PERÍODO" Then

   sSQL = "SELECT fornecedor.codigo, fornecedor.razao, produtos_entrada.codigo AS var_codent, produtos_entrada.data_entrada, produtos_entrada.notafiscal, produtos_entrada.data_EMISSAO, produtos_entrada.VALOR_FRETE, produtos_entrada.valor, (CASE WHEN XML = 1 THEN 'XML' ELSE 'MANUAL' END) AS vXML  " & _
          "FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo " & _
          "WHERE (DATA_EMISSAO >= CONVERT(DATETIME, '" & Format(mskInicial.Text, ocDATA) & "', 103)) AND (DATA_EMISSAO <= CONVERT(DATETIME, '" & Format(mskFinal.Text, ocDATA) & "', 103))  " & _
          "ORDER BY " & INDICE
End If

'Abre a consulta
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

'===FUNÇÃO DE CONTAR REGISTROS
lblQuant.Caption = Format(totalRegistros, "00")

If cboConsulta.Text = "TODOS" Or cboConsulta.Text = "NOTA FISCAL" Or cboConsulta.Text = "FORNECEDOR" Or cboConsulta.Text = "MENSAL" Or cboConsulta.Text = "PERÍODO" Then
    FormatarGrid r
Else
    FormatarGridConsProdutos r
End If

If r.State <> 0 Then r.Close
Set r = Nothing
   
printSQL = sSQL
End Sub
Private Sub cmdFechar_Click()
   If txtCodigo.Text <> "" And cmdSalvar.Enabled = True Then
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

If cboConsulta.Text = "TODOS" Or cboConsulta.Text = "NOTA FISCAL" Or cboConsulta.Text = "FORNECEDOR" Or cboConsulta.Text = "MENSAL" Or cboConsulta.Text = "PERÍODO" Then
   Set REL_Prod_Entrada_Nota.Relatorio.Recordset = r
   REL_Prod_Entrada_Nota.dfQuant.Caption = lblQuant.Caption
   REL_Prod_Entrada_Nota.dfBruto.Caption = lblValor.Caption
   
   If cboConsulta.Text = "MENSAL" Then
      REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Mês = " & cboConsDescricao.Text & "/" & cboConsAno.Text
   ElseIf cboConsulta.Text = "PERÍODO" Then
      REL_Prod_Entrada_Nota.dfTipo.Caption = "Período: " & mskInicial.Text & " até " & mskFinal.Text
   ElseIf cboConsulta.Text = "FORNECEDOR" Then
      REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Fornecedor = " & cboConsDescricao.Text & ""
   ElseIf cboConsulta.Text = "NOTA FISCAL" Then
      REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & cboConsDescricao.Text & ""
   Else
      REL_Prod_Entrada_Nota.dfTipo.Caption = "Tipo: Todas as notas"
   End If
   
   REL_Prod_Entrada_Nota.Relatorio.Ativar
   Unload REL_Prod_Entrada_Nota
Else
   Set REL_Prod_Entrada_PorProduto.Relatorio.Recordset = r
   REL_Prod_Entrada_PorProduto.dfQuant.Caption = lblQuant.Caption
   REL_Prod_Entrada_PorProduto.dfBruto.Caption = lblValor.Caption

   If cboConsulta.Text = "PRODUTO" Then
      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Produto = " & cboConsDescricao.Text & ""
   ElseIf cboConsulta.Text = "FABRICANTE" Then
      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Fabricante = " & cboConsDescricao.Text & ""
   ElseIf cboConsulta.Text = "REFERÊNCIA" Then
      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Referência " & cboConsDescricao.Text & ""
   ElseIf cboConsulta.Text = "CÓD. BARRA" Then
      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Cód. de Barra " & cboConsDescricao.Text & ""
   End If
   
   REL_Prod_Entrada_PorProduto.Relatorio.Ativar
   Unload REL_Prod_Entrada_PorProduto
End If

Me.Show 1
End Sub

Private Sub cmdNovo_Click()
frmNota.Enabled = True
frmTransporte.Enabled = True
frmItens.Enabled = True
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdImprimirEntrada.Enabled = False
lblTransportadora.Enabled = False
cboTransportadora.Enabled = False
lblFreteTotal.Enabled = False
txtFreteTotal.Enabled = False
Limpar_Objetos
Limpar_SubDados
Limpar_Valores
mskData.Text = Format(Date, "dd/mm/yy")
mskHora.Text = Format(Now, "hh:mm")
Mostrar_Historico
Auto_Numeracao
Mostrar_Itens
mskData.SetFocus
End Sub

Private Sub cmdRemover_Click()
If Grid_Cadastro.rows <= 1 Then Exit Sub

If Grid_Cadastro.TextMatrix(Grid_Cadastro.RowSel, 1) <> "" Then
    'dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo = " & Grid_Cadastro.TextMatrix(Grid_Cadastro.RowSel, 1) & ");"
    dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo = " & Grid_Cadastro.TextMatrix(Grid_Cadastro.Row, 1) & ") AND (codigo_entrada = " & txtCodigo.Text & ");"
    dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Grid_Cadastro.TextMatrix(Grid_Cadastro.RowSel, 6) & " WHERE (codigo = " & Grid_Cadastro.TextMatrix(Grid_Cadastro.RowSel, 4) & ");"
    dbData.Execute "DELETE FROM Produtos_Quant WHERE (COD_PRODUTO = " & Grid_Cadastro.TextMatrix(Grid_Cadastro.Row, 4) & ") AND (COD_ENTRADA = " & txtCodigo.Text & ");"
    dbData.Execute "DELETE FROM Produtos_Precos WHERE (COD_PRODUTO = " & Grid_Cadastro.TextMatrix(Grid_Cadastro.Row, 4) & ") AND (COD_ENTRADA = " & txtCodigo.Text & ");"
End If

Mostrar_Itens
End Sub

Private Sub cmdSalvar_Click()
If txtCodigo.Text = "" Or cboFornecedor.Text = "" Or txtNotaFiscal.Text = "" Or cboTipoFrete.Text = "" Then
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
Limpar_Valores
Mostrar_Itens
cmdExibir_Click

DesativarBotoes
End Sub

Private Function Atualizar_Dados() As Boolean
Dim sSQL As String
Dim varValorFrete As Currency

If txtFreteTotal.Text = "" Then varValorFrete = CCur(0) Else varValorFrete = CCur(txtFreteTotal.Text)

'Comando de atualização
sSQL = "UPDATE produtos_entrada SET " & _
   "data_entrada = CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
   "hora_entrada = '" & Format$(mskHora.Text, ocHORA) & "', " & _
   "COD_FORNECEDOR = " & TxtCodFornecedor.Text & ", " & _
   "COD_TRANSPORTADORA = " & txtCodTransportadora.Text & ", " & _
   "DATA_EMISSAO = CONVERT(DATETIME, '" & Format$(mskDataEmissao.Text, ocDATA) & "', 103), " & _
   "DATA_SAIDA = CONVERT(DATETIME, '" & Format$(mskDataSaida.Text, ocDATA) & "', 103), " & _
   "HORA_SAIDA = '" & Format$(mskHoraSaida.Text, ocHORA) & "', " & _
   "notafiscal = '" & txtNotaFiscal.Text & "', " & _
   "TIPO_FRETE = '" & cboTipoFrete.Text & "', " & _
   "valor_frete = " & FSQL(varValorFrete) & "  , " & _
   "valor = " & Replace(CCur(txtValor.Text), ",", ".")

'Condição para atualização
sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ");"

'Retorna o resultado da atualização
Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Function Inserir_Dados() As Boolean
Dim sSQL As String
Dim varValorFrete As Currency
If txtFreteTotal.Text = "" Then varValorFrete = CCur(0) Else varValorFrete = CCur(txtFreteTotal.Text)

sSQL = "INSERT INTO produtos_entrada (" & _
   "codigo, data_entrada, hora_entrada, cod_fornecedor, notafiscal, " & _
   "valor_frete, valor, DATA_EMISSAO, DATA_SAIDA, HORA_SAIDA, TIPO_FRETE, COD_TRANSPORTADORA, XML) VALUES ("

sSQL = sSQL & _
   txtCodigo & ", CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), '" & Format$(mskHora.Text, ocHORA) & "', " & TxtCodFornecedor.Text & ", " & txtNotaFiscal.Text & ",  " & FSQL(varValorFrete) & ", " & IIf((txtValor.Text = ""), "Null", FSQL(txtValor.Text)) & ", CONVERT(DATETIME, '" & Format$(mskDataEmissao.Text, ocDATA) & "', 103), CONVERT(DATETIME, '" & Format$(mskDataSaida.Text, ocDATA) & "', 103), '" & Format$(mskHoraSaida.Text, ocHORA) & "',  '" & cboTipoFrete.Text & "', " & IIf((txtCodTransportadora.Text = ""), "Null", txtCodTransportadora.Text) & ", 0 )"
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

Mostrar_Itens
LimparGrid_Consulta
End Sub

Private Sub mskDataEmissao_GotFocus()
SelectControl mskDataEmissao
End Sub


Private Sub mskDataEmissao_KeyPress(KeyAscii As Integer)
mskDataEmissao.Mask = "##/##/##"
End Sub


Private Sub mskDataEmissao_LostFocus()
If mskDataEmissao.Text = "" Or mskDataEmissao.Text = "__/__/__" Then
   mskDataEmissao.Mask = ""
   mskDataEmissao.Text = ""
Else
   If IsDate(mskDataEmissao.Text) Then
      Exit Sub
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskDataEmissao.SetFocus
   End If
End If

End Sub

Private Sub mskDataSaida_GotFocus()
SelectControl mskDataSaida
End Sub


Private Sub mskDataSaida_KeyPress(KeyAscii As Integer)
mskDataSaida.Mask = "##/##/##"
End Sub


Private Sub mskDataSaida_LostFocus()
If mskDataSaida.Text = "" Or mskDataSaida.Text = "__/__/__" Then
   mskDataSaida.Mask = ""
   mskDataSaida.Text = ""
Else
   If IsDate(mskDataSaida.Text) Then
      Exit Sub
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskDataSaida.SetFocus
   End If
End If
End Sub

Private Sub mskFinal_GotFocus()
SelectControl mskInicial
End Sub

Private Sub mskFinal_KeyPress(KeyAscii As Integer)
mskFinal.Mask = "##/##/##"
End Sub

Private Sub mskFinal_LostFocus()
If mskFinal.Text = "" Or mskFinal.Text = "__/__/__" Then
   mskFinal.Mask = ""
   mskFinal.Text = ""
   Exit Sub
Else
   If IsDate(mskFinal.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskFinal.SetFocus
      SelectControl mskFinal
   End If
End If
End Sub


Private Sub mskHora_KeyPress(KeyAscii As Integer)
mskHora.Mask = "##:##"
End Sub

Private Sub mskHora_LostFocus()
If mskHora.Text = "" Or mskHora.Text = "__:__" Then
   mskHora.Mask = ""
   mskHora.Text = ""
End If
End Sub

Private Sub mskHoraSaida_GotFocus()
SelectControl mskData
End Sub


Private Sub mskHoraSaida_KeyPress(KeyAscii As Integer)
mskHoraSaida.Mask = "##:##"
End Sub


Private Sub mskHoraSaida_LostFocus()
If mskHoraSaida.Text = "" Or mskHoraSaida.Text = "__:__" Then
   mskHoraSaida.Mask = ""
   mskHoraSaida.Text = ""
End If
End Sub

Private Sub mskInicial_GotFocus()
SelectControl mskInicial
End Sub

Private Sub mskInicial_KeyPress(KeyAscii As Integer)
mskInicial.Mask = "##/##/##"
End Sub

Private Sub mskInicial_LostFocus()
If mskInicial.Text = "" Or mskInicial.Text = "__/__/__" Then
   mskInicial.Mask = ""
   mskInicial.Text = ""
   Exit Sub
Else
   If IsDate(mskInicial.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskInicial.SetFocus
      SelectControl mskInicial
   End If
End If
End Sub


Private Sub txtAliqICMSProd_GotFocus()
SelectControl txtAliqICMSProd
End Sub

Private Sub txtAliqICMSProd_LostFocus()
If txtAliqICMSProd.Text = "" Then txtAliqICMSProd.Text = "0"
Moeda txtAliqICMSProd

Dim varValorCustoProd As Currency
Dim varAliqICMS As Double
Dim varValorICMS As Currency

If txtCusto.Text = "" Then Exit Sub
If txtAliqICMSProd.Text = "" Then Exit Sub

varValorCustoProd = txtCusto.Text
varAliqICMS = txtAliqICMSProd.Text
varValorICMS = ((varValorCustoProd * varAliqICMS) / 100)
txtValorICMSProd = Format(varValorICMS, ocMONEY)
End Sub


Private Sub txtAliqICMSSTProd_GotFocus()
SelectControl txtAliqICMSSTProd
End Sub

Private Sub txtAliqICMSSTProd_LostFocus()
If txtAliqICMSSTProd.Text = "" Then txtAliqICMSSTProd.Text = "0"
Moeda txtAliqICMSSTProd

Dim varValorCustoProd As Currency
Dim varAliqICMS As Double
Dim varValorICMS As Currency

If txtCusto.Text = "" Then Exit Sub
If txtAliqICMSSTProd.Text = "" Then Exit Sub

varValorCustoProd = txtCusto.Text
varAliqICMS = txtAliqICMSSTProd.Text
varValorICMS = ((varValorCustoProd * varAliqICMS) / 100)
txtValorICMSSTProd = Format(varValorICMS, ocMONEY)
End Sub


Private Sub txtAliqIPIProd_GotFocus()
SelectControl txtAliqIPIProd
End Sub

Private Sub txtAliqIPIProd_LostFocus()
If txtAliqIPIProd.Text = "" Then txtAliqIPIProd.Text = "0"
Moeda txtAliqIPIProd

Dim varValorCustoProd As Currency
Dim varAliqICMS As Double
Dim varValorICMS As Currency

If txtCusto.Text = "" Then Exit Sub
If txtAliqIPIProd.Text = "" Then Exit Sub

varValorCustoProd = txtCusto.Text
varAliqICMS = txtAliqIPIProd.Text
varValorICMS = ((varValorCustoProd * varAliqICMS) / 100)
txtValorIPIProd = Format(varValorICMS, ocMONEY)
End Sub


Private Sub txtCusto_GotFocus()
txtCusto.SelStart = 0
txtCusto.SelLength = Len(txtCusto)
End Sub

Private Sub txtCusto_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
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


Private Sub txtMargemAP_Validate(Cancel As Boolean)
txtMargemAP_LostFocus
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

Private Sub txtMargemAV_Validate(Cancel As Boolean)
txtMargemAV_LostFocus
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

Private Sub txtMargemVP_Validate(Cancel As Boolean)
txtMargemVP_LostFocus
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

txtMargemVV.Text = FormatNumber(varMargemVV, 2) & "%"
If txtMargemVP.Text = "" Then txtMargemVP.Text = txtMargemVV.Text
If txtMargemAV.Text = "" Then txtMargemAV.Text = txtMargemVV.Text
If txtMargemAP.Text = "" Then txtMargemAP.Text = txtMargemVV.Text
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

Private Sub Limpar_Valores()
txtMargemVV.Text = ""
txtMargemVP.Text = ""
txtMargemAV.Text = ""
txtMargemAP.Text = ""
txtValorVV.Text = ""
txtValorVP.Text = ""
txtValorAV.Text = ""
txtValorAP.Text = ""
txtCusto.Text = ""
End Sub

Private Sub txtCodBarra_GotFocus()
SelectControl txtCodBarra
End Sub

Private Sub txtCodProduto_Change()
'   If txtCodProduto.Text = "" Then
'      txtQuantAtual.Text = ""
'      cboDescricao.Text = ""
'      cboDescricao2.Text = ""
'      txtCodBarra.Text = ""
'      txtValorAtual.Text = ""
'   End If
   
   'Call Abrir_BancodeDados
   'SQL = "SELECT CODIGO, QUANT_ESTOQUE FROM PRODUTOS WHERE CODIGO = " & txtCodProduto.Text & ""
   'Set RS = BD.OpenRecordset(SQL)
   
   'If Not IsNull(RS!QUANT_ESTOQUE) Then txtQuantAtual.Text = RS!QUANT_ESTOQUE
End Sub

Private Sub txtFreteTotal_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub



Private Sub txtNotaFiscal_LostFocus()
   'Calcular_Frete
End Sub

Private Sub txtQuant_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub FormatarGridConsProdutos(rTabela As ADODB.Recordset, Optional ByVal Agrupar As Boolean = False)
Dim i As Integer, x As Integer

Dim aux As String, iRow As Long
Dim subtotalQtde As Double
Dim bNovoGrupo As Boolean

With Grid
   .Clear
   .Cols = 10
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 600
   .ColWidth(2) = 1100
   .ColWidth(3) = 1000
   .ColWidth(4) = 1000
   .ColWidth(5) = 4200
   .ColWidth(6) = 1200
   .ColWidth(7) = 1200
   .ColWidth(8) = 1200
   .ColWidth(9) = 1000
   
   .TextMatrix(0, 1) = "CÓD."
   .TextMatrix(0, 2) = "CADASTRO"
   .TextMatrix(0, 3) = "NO. NOTA"
   .TextMatrix(0, 4) = "EMISSÃO"
   .TextMatrix(0, 5) = "PRODUTO"
   .TextMatrix(0, 6) = "QUANT"
   .TextMatrix(0, 7) = "CUSTO"
   .TextMatrix(0, 8) = "VENDA VV"
   .TextMatrix(0, 9) = "CUSTO"
   
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
   
   .Redraw = False
   i = 1
            
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         'ALINHAMENTO
         .ColAlignment(2) = 1
         .TextMatrix(.rows - 1, 1) = rTabela("var_codent")
         .TextMatrix(.rows - 1, 2) = Format$(rTabela("data_entrada"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("notafiscal"))
         .TextMatrix(.rows - 1, 4) = Format$(rTabela("data_EMISSAO"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("descricao"))
         .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("QUANT"))
         .TextMatrix(.rows - 1, 7) = Format$(rTabela("CUSTO"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = Format$(rTabela("VALOR_VV"), ocMONEY)
         .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("vXML"))
         
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
   'MUDAR COR DE FONTE DA COLUNA
   For x = 1 To .rows - 1
      .Row = i
      .Col = 5
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
   
   .rows = .rows - 1
   .Redraw = True
End With

lblValor.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
lblQuant.Caption = SomaGrid(Grid, 6)
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset, Optional ByVal Agrupar As Boolean = False)
Dim i As Integer, x As Integer

Dim aux As String, iRow As Long
Dim subtotalQtde As Double
Dim bNovoGrupo As Boolean

With Grid
   .Clear
   .Cols = 9
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 1100
   .ColWidth(3) = 1000
   .ColWidth(4) = 1000
   .ColWidth(5) = 4200
   .ColWidth(6) = 1200
   .ColWidth(7) = 1200
   .ColWidth(8) = 1500
   
   .TextMatrix(0, 1) = "CÓD."
   .TextMatrix(0, 2) = "ENTRADA"
   .TextMatrix(0, 3) = "NO. NOTA"
   .TextMatrix(0, 4) = "EMISSÃO"
   .TextMatrix(0, 5) = "FORNECEDOR"
   .TextMatrix(0, 6) = "FRETE"
   .TextMatrix(0, 7) = "VALOR"
   .TextMatrix(0, 8) = "TIPO ENTRADA"
   
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
   
   .Redraw = False
   i = 1
            
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         'ALINHAMENTO
         .ColAlignment(2) = 1
         
         .TextMatrix(.rows - 1, 1) = rTabela("var_codent")
         .TextMatrix(.rows - 1, 2) = Format$(rTabela("DATA_ENTRADA"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("notafiscal"))
         .TextMatrix(.rows - 1, 4) = Format$(rTabela("data_EMISSAO"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("razao"))
         .TextMatrix(.rows - 1, 6) = Format$(rTabela("VALOR_FRETE"), ocMONEY)
         .TextMatrix(.rows - 1, 7) = Format$(rTabela("valor"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("vXML"))
         
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
   'MUDAR COR DE FONTE DA COLUNA
   For x = 1 To .rows - 1
      .Row = i
      .Col = 5
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
   
   .rows = .rows - 1
   .Redraw = True
End With

lblValor.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
End Sub

Private Sub FormatarGrid2(ByVal SaldoAnterior As Double, Movimento() As String)
   Dim i As Integer, x As Integer
   
   Dim aux As String, iRow As Long
   Dim subtotalQtde As Double
   Dim bNovoGrupo As Boolean
   
   With Grid
      .Clear
      .Cols = 6
      .rows = 2
      
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
      
      .Redraw = False
      
      'Adiciona o saldo anterior
      .TextMatrix(.rows - 1, 4) = "SALDO ANTERIOR"
      .TextMatrix(.rows - 1, 5) = Format$(SaldoAnterior, ocPESO)
      
      .Row = .rows - 1
      .Col = 5
      
      If CDbl(.TextMatrix(.Row, .Col)) > 0 Then
         .CellForeColor = RGB(0, 128, 0)
      ElseIf CDbl(.TextMatrix(.Row, .Col)) < 0 Then
         .CellForeColor = RGB(192, 0, 0)
      Else
         .CellForeColor = RGB(0, 0, 192)
      End If
      
      .CellFontBold = True
      .rows = .rows + 1
      
      For i = 1 To UBound(Movimento, 2)
         'ALINHAMENTO
         .ColAlignment(2) = 1
         
         '.TextMatrix(.Rows - 1, 1) = rTabela("var_codent")
         .TextMatrix(.rows - 1, 2) = Movimento(1, i)
         .TextMatrix(.rows - 1, 3) = Movimento(2, i)
         .TextMatrix(.rows - 1, 4) = Movimento(3, i)
         .TextMatrix(.rows - 1, 5) = Movimento(4, i)
         
         .Row = .rows - 1
         .Col = 5
         
         If CDbl(Movimento(4, i)) > 0 Then
            .CellForeColor = RGB(0, 128, 0)
         ElseIf CDbl(Movimento(4, i)) < 0 Then
            .CellForeColor = RGB(192, 0, 0)
         Else
            .CellForeColor = RGB(0, 0, 192)
         End If
         
         .CellFontBold = True
         .rows = .rows + 1
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      'For X = 1 To .Rows - 1
      '   .Row = i
      '   .Col = 5
      '   .CellForeColor = &HC0&
      '   .CellFontBold = True
      'Next
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   lblValor.Caption = Format(SomaGrid(Grid, 5), ocMONEY)
End Sub

Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
Dim i As Integer, j As Integer
Dim x As Integer

With Grid_Cadastro
   .Clear
   .Cols = 17
   .rows = 2
   
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 0
    .ColWidth(3) = 0
    .ColWidth(4) = 600
    .ColWidth(5) = 4000
    .ColWidth(6) = 700
    .ColWidth(7) = 850
    .ColWidth(8) = 650
    .ColWidth(9) = 850
    .ColWidth(10) = 650
    .ColWidth(11) = 850
    .ColWidth(12) = 650
    .ColWidth(13) = 850
    .ColWidth(14) = 650
    .ColWidth(15) = 850
    .ColWidth(16) = 0

    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "COD_ENTRADA"
    .TextMatrix(0, 3) = "COD_BARRA"
    .TextMatrix(0, 4) = "CÓD."
    .TextMatrix(0, 5) = "DESCRIÇÃO"
    .TextMatrix(0, 6) = "QTDE"
    .TextMatrix(0, 7) = "CUSTO"
    .TextMatrix(0, 8) = "% VV"
    .TextMatrix(0, 9) = "VALOR"
    .TextMatrix(0, 10) = "% VP "
    .TextMatrix(0, 11) = "VALOR"
    .TextMatrix(0, 12) = "% AV"
    .TextMatrix(0, 13) = "VALOR"
    .TextMatrix(0, 14) = "% AP"
    .TextMatrix(0, 15) = "VALOR"
    .TextMatrix(0, 16) = "SUBTOTAL"
    
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
            .TextMatrix(.rows - 1, 1) = rTabela("varCod")
            .TextMatrix(.rows - 1, 2) = rTabela("codigo_entrada")
            .TextMatrix(.rows - 1, 4) = Format$(rTabela("codigo_produto"), "0000")
         If tipoEmpresa = 4 Then
            .TextMatrix(.rows - 1, 5) = rTabela("descricao") & " /  " & rTabela("tamanho") & " / " & rTabela("var_ref")
         Else
            .TextMatrix(.rows - 1, 5) = rTabela("descricao")
         End If
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("quant"))
         .TextMatrix(.rows - 1, 7) = Format$(rTabela("custo"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = FormatNumber(rTabela("MARGEM_VV"), 2) & "%"
         .TextMatrix(.rows - 1, 9) = Format$(rTabela("VALOR_VV"), ocMONEY)
         .TextMatrix(.rows - 1, 10) = FormatNumber(rTabela("MARGEM_VP"), 2) & "%"
         .TextMatrix(.rows - 1, 11) = Format$(rTabela("VALOR_VP"), ocMONEY)
         .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("MARGEM_AV"), 2) & "%"
         .TextMatrix(.rows - 1, 13) = Format$(rTabela("VALOR_AV"), ocMONEY)
         .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela("MARGEM_AP"), 2) & "%"
         .TextMatrix(.rows - 1, 15) = Format$(rTabela("VALOR_AP"), ocMONEY)
         .TextMatrix(.rows - 1, 16) = Format$(rTabela("varTotalCustoItem"), ocMONEY)
         
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

Private Sub Form_Load()
SSTab1.Tab = 0
DesativarBotoes
lblTipoConsulta.Caption = "0"
txtCodUsuario.Text = Tela_Principal.txtCodFuncionario.Text
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")

'cboFiltroNota_GotFocus
'cboFiltroNota.ListIndex = 4
'cboConNotaMes.Text = Format(Date, "mmmm")
'cboConNotaAno.Text = Year(Date)
'cmdExibirConNotas_Click

Set moCombo = New cComboHelper
End Sub

Private Sub Mostrar_Historico()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cboFornecedor.Text = "" Then
      sSQL = "SELECT * FROM produtos_entrada WHERE 1 = 0"
   
   Else
      sSQL = "SELECT produtos_entrada.*, fornecedor.codigo, fornecedor.razao as varFornecedor " & _
             "FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo " & _
             "WHERE (cod_fornecedor = " & TxtCodFornecedor.Text & ") ORDER BY data_entrada;"
         '"FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _

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
   
   sSQL = "SELECT DISTINCT descricao, codigo, fabricante, tamanho FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If cboDescricao.Text <> "" Then var_cboTexto = cboDescricao.Text
   cboDescricao.Clear
   cboDescricao.Text = var_cboTexto
   
   Do While Not r.EOF
      If tipoEmpresa = 4 Then
          cboDescricao.AddItem ValidateNull(r("descricao")) & " /  " & ValidateNull(r("tamanho")) & " / " & ValidateNull(r("fabricante"))
          cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
      Else
         cboDescricao.AddItem ValidateNull(r("descricao"))
         cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
      End If
        
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
'CHECAR SE O PEDIDO ESTÁ FECHADO
If txtCodigo.Text = "" Then Exit Sub
If Grid_Cadastro.rows >= 1 And cmdNovo.Enabled = False Then cmdCancelar_Click
Set moCombo = Nothing
End Sub



Private Sub Grid_DblClick()
If Grid.rows <= 1 Then Exit Sub

If Grid.TextMatrix(Grid.Row, 8) = "XML" Then
    MsgBox "Não é permitida a visualização de uma nota fiscal com tipo de entrada igual a XML por aqui!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

cmdNovo.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
cmdImprimirEntrada.Enabled = True
frmNota.Enabled = True
frmTransporte.Enabled = True
frmItens.Enabled = True
SSTab1.Tab = 0
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub mskData_GotFocus()
SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
If mskData.Text = "" Or mskData.Text = "__/__/__" Then
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

Private Sub txtCodBarra_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodBarra_LostFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

If lblTipoConsulta.Caption = "0" Or lblTipoConsulta.Caption = "1" Then
    If txtCodBarra.Text = "" Then lblTipoConsulta.Caption = "0": txtCodProduto.Text = "": cboDescricao.Locked = False: cboDescricao.Text = "": Exit Sub
     'txtCodProduto.Text = "":
    sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, tamanho, REF, fabricante, quant_estoque FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "') AND (ativo = 1);"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then
       txtCodProduto.Text = r("var_codprod")
       
       If tipoEmpresa = 4 Then
           cboDescricao.Text = ValidateNull(r("var_desc")) & " /  " & ValidateNull(r("tamanho")) & " / " & ValidateNull(r("fabricante")) & " /  " & r("REF")
       Else
          cboDescricao.Text = ValidateNull(r("var_desc"))
       End If
       
       txtQuantAtual.Text = ValidateNull(r("quant_Estoque"))
       lblTipoConsulta.Caption = "1"
       cboDescricao.Locked = True
    Else
       ShowMsg "Produto Inexistente!", vbCritical
       txtCodBarra.Text = ""
       lblTipoConsulta.Caption = "0"
       cboDescricao.Locked = False
       cboDescricao.Text = ""
       txtCodProduto.Text = ""
       txtCodBarra.SetFocus
       Exit Sub
    End If
    
    MostrarValorVenda
    txtQuant.SetFocus
End If

On Local Error Resume Next
End Sub

Private Sub txtCodigo_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If cmdAlterar.Enabled = True Then
   If txtCodigo.Text = "" Then Exit Sub
   
   sSQL = "SELECT produtos_entrada.*, fornecedor.codigo, fornecedor.razao AS varFornecedor " & _
          "FROM produtos_entrada INNER JOIN fornecedor ON produtos_entrada.cod_fornecedor = fornecedor.codigo " & _
          "WHERE (produtos_entrada.codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   '          "INNER JOIN produtos_entrada ON produtos_entrada.cod_transportadora = transportadora.codigo  " & _
   If r.EOF Then Exit Sub
   '     "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
     "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _

   Limpar_Objetos
   Mostrar_Dados r
   Mostrar_Itens
   Mostrar_Historico
   mskData.SetFocus
End If
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then
'txtCodigo.Text = ValidateNull(rTabela("codigo"))
mskData.Text = Format$(rTabela("data_entrada"), "dd/mm/yy")
mskHora.Text = Format$(rTabela("hora_entrada"), ocHRMN)
TxtCodFornecedor.Text = ValidateNull(rTabela("COD_FORNECEDOR"))
txtNotaFiscal.Text = ValidateNull(rTabela("notafiscal"))
txtValor.Text = Format$(rTabela("valor"), ocMONEY)
txtFreteTotal.Text = Format$(rTabela("VALOR_frete"), ocMONEY)
mskDataEmissao.Text = Format$(rTabela("DATA_EMISSAO"), "dd/mm/yy")
mskDataSaida.Text = Format$(rTabela("DATA_SAIDA"), "dd/mm/yy")
mskHoraSaida.Text = Format$(rTabela("HORA_SAIDA"), ocHRMN)
cboTipoFrete.Text = ValidateNull(rTabela("TIPO_FRETE"))
txtCodTransportadora.Text = ValidateNull(rTabela("COD_TRANSPORTADORA"))
cboFornecedor.Text = ValidateNull(rTabela("varFornecedor"))
cboTransportadora.Text = ""
End If
End Sub

Private Sub txtFreteTotal_GotFocus()
SelectControl txtFreteTotal
End Sub

Private Sub txtFreteTotal_LostFocus()
Dim varFrete As Currency

If txtFreteTotal.Text = "" Then Exit Sub
varFrete = txtFreteTotal.Text

txtFreteTotal.Text = FormatNumber(varFrete, 2)
End Sub

Private Sub txtQuant_GotFocus()
   SelectControl txtQuant
End Sub

Private Sub CalcularPrecos()
Dim varValorCusto As Currency
If txtCusto.Text = "" Then Exit Sub
varValorCusto = txtCusto.Text

'CALCULAR PREÇO - VAREJO A VISTA
Dim varMargemVV As Currency
Dim varValorVV As Currency

If txtMargemVV.Text = "" Then Exit Sub

varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)

varValorVV = (varValorCusto * varMargemVV) / 100
varValorVV = varValorCusto + varValorVV
txtValorVV.Text = Format(varValorVV, ocMONEY)

'CALCULAR PREÇO - VAREJO A PRAZO
Dim varMargemVP As Currency
Dim varValorVP As Currency

If txtMargemVP.Text = "" Then Exit Sub

varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)

varValorVP = (varValorCusto * varMargemVP) / 100
varValorVP = varValorCusto + varValorVP
txtValorVP.Text = Format(varValorVP, ocMONEY)

'CALCULAR PREÇO - ATACADO A VISTA
Dim varMargemAV As Currency
Dim varValorAV As Currency

If txtMargemAV.Text = "" Then Exit Sub

varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)

varValorAV = (varValorCusto * varMargemAV) / 100
varValorAV = varValorCusto + varValorAV
txtValorAV.Text = Format(varValorAV, ocMONEY)

'CALCULAR PREÇO - ATACADO A PRAZO
Dim varMargemAP As Currency
Dim varValorAP As Currency

If txtMargemAP.Text = "" Then Exit Sub

varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)

varValorAP = (varValorCusto * varMargemAP) / 100
varValorAP = varValorCusto + varValorAP
txtValorAP.Text = Format(varValorAP, ocMONEY)
End Sub


Private Sub txtMargemAP_LostFocus()
Dim varMargemAP As Currency

If txtMargemAP.Text = "" Then txtMargemAP.Text = 0

If Right(txtMargemAP.Text, 1) = "%" Then
   varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)
Else
    varMargemAP = txtMargemAP.Text
End If

txtMargemAP.Text = FormatNumber(varMargemAP, 2) & "%"

CalcularPrecos
lblAviso.Visible = False
End Sub



Private Sub txtMargemAV_LostFocus()
Dim varMargemAV As Currency

If txtMargemAV.Text = "" Then txtMargemAV.Text = 0

If Right(txtMargemAV.Text, 1) = "%" Then
   varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)
Else
    varMargemAV = txtMargemAV.Text
End If

txtMargemAV.Text = FormatNumber(varMargemAV, 2) & "%"

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



Private Sub txtMargemVP_LostFocus()
Dim varMargemVP As Currency

If txtMargemVP.Text = "" Then txtMargemVP.Text = 0

If Right(txtMargemVP.Text, 1) = "%" Then
   varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)
Else
    varMargemVP = txtMargemVP.Text
End If
txtMargemVP.Text = FormatNumber(varMargemVP, 2) & "%"

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



Private Sub txtCusto_LostFocus()
Dim varLucro As Currency

If txtCusto.Text = "" Then Exit Sub
varLucro = txtCusto.Text

txtCusto.Text = FormatNumber(varLucro, 2)

CalcularPrecos
End Sub





Private Sub txtValor_GotFocus()
   SelectControl txtValor
End Sub

Private Sub txtValor_LostFocus()
   txtValor.Text = Format(txtValor.Text, "##,##0.00")
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
c = ((B - a) / a) * 100

txtMargemAP.Text = FormatNumber(c, 2) & "%"
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
c = ((B - a) / a) * 100

txtMargemAV.Text = FormatNumber(c, 2) & "%"
txtValorAV.Text = Format(txtValorAV.Text, ocMONEY)
End Sub


Private Sub txtValorICMSProd_GotFocus()
SelectControl txtValorICMSProd
End Sub


Private Sub txtValorICMSSTProd_GotFocus()
SelectControl txtAliqICMSProd
End Sub


Private Sub txtValorIPIProd_GotFocus()
SelectControl txtValorIPIProd
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
c = ((B - a) / a) * 100

txtMargemVP.Text = FormatNumber(c, 2) & "%"
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
c = ((B - a) / a) * 100

txtMargemVV.Text = FormatNumber(c, 2) & "%"
txtValorVV.Text = Format(txtValorVV.Text, ocMONEY)
End Sub


