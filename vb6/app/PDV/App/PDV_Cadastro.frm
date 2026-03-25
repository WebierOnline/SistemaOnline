VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PDV_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ESTONAR"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   Icon            =   "PDV_Cadastro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   10965
      TabIndex        =   17
      Top             =   60
      Width           =   10995
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7560
         TabIndex        =   78
         Top             =   540
         Width           =   1335
      End
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   9420
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "000000"
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   8880
         TabIndex        =   20
         Top             =   120
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "PDV_Cadastro.frx":23D2
         Top             =   0
         Width           =   1140
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ESTONAR VENDA"
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
         TabIndex        =   18
         Top             =   240
         Width           =   2670
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8835
      Left            =   60
      TabIndex        =   9
      Top             =   1020
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   15584
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
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
      TabCaption(0)   =   "PEDIDO"
      TabPicture(0)   =   "PDV_Cadastro.frx":8C18
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdFinalizarPedido"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdProdutos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdClientes"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAbrirVenda"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmProduto"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Picture2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.PictureBox Picture2 
         Height          =   6555
         Left            =   120
         ScaleHeight     =   6495
         ScaleWidth      =   10695
         TabIndex        =   15
         Top             =   1620
         Width           =   10755
         Begin VB.Frame frmVendaPrazo 
            BackColor       =   &H00C0C0FF&
            Caption         =   "VENDA À PRAZO"
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
            Left            =   1620
            TabIndex        =   48
            Top             =   360
            Visible         =   0   'False
            Width           =   7515
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0C0FF&
               Height          =   1575
               Left            =   120
               TabIndex        =   64
               Top             =   1620
               Width           =   7275
               Begin VB.ComboBox cboCliente 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   69
                  Top             =   420
                  Width           =   7035
               End
               Begin VB.TextBox txtCodCliente 
                  Appearance      =   0  'Flat
                  Height          =   225
                  Left            =   6420
                  TabIndex        =   68
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   735
               End
               Begin VB.TextBox txtValorParc 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2700
                  TabIndex        =   67
                  Text            =   "0"
                  Top             =   1080
                  Width           =   1215
               End
               Begin VB.TextBox txtEntrada 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  TabIndex        =   66
                  Top             =   1080
                  Width           =   1215
               End
               Begin VB.ComboBox cboPrazo 
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   65
                  Text            =   "30"
                  Top             =   1080
                  Width           =   1275
               End
               Begin MSMask.MaskEdBox mskVencimento 
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   70
                  Top             =   1080
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptChar      =   "_"
               End
               Begin ChamaleonBtn.chameleonButton cmdFinalizar 
                  Height          =   315
                  Left            =   5040
                  TabIndex        =   71
                  Top             =   1080
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "Finalizar"
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
                  MICON           =   "PDV_Cadastro.frx":8C34
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
                  Height          =   315
                  Left            =   6120
                  TabIndex        =   72
                  Top             =   1080
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
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
                  MICON           =   "PDV_Cadastro.frx":8C50
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cliente"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   77
                  Top             =   180
                  Width           =   480
               End
               Begin VB.Label lblValorParc 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor Rest."
                  Height          =   195
                  Left            =   2700
                  TabIndex        =   76
                  Top             =   840
                  Width           =   780
               End
               Begin VB.Label lblQuantParc 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prazo (dias)"
                  Height          =   195
                  Left            =   1380
                  TabIndex        =   75
                  Top             =   840
                  Width           =   825
               End
               Begin VB.Label lblInicio 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Vencimento"
                  Height          =   195
                  Left            =   3960
                  TabIndex        =   74
                  Top             =   840
                  Width           =   840
               End
               Begin VB.Label lblEntrada 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Entrada"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   73
                  Top             =   840
                  Width           =   555
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0FF&
               Height          =   735
               Left            =   120
               TabIndex        =   54
               Top             =   840
               Width           =   7275
               Begin VB.TextBox txtSubTotal 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   60
                  TabStop         =   0   'False
                  Text            =   "0,00"
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.TextBox txtTotalDesc 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   5700
                  Locked          =   -1  'True
                  TabIndex        =   59
                  TabStop         =   0   'False
                  Text            =   "0,00"
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.PictureBox Picture3 
                  BackColor       =   &H00C0C0FF&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   3180
                  ScaleHeight     =   210
                  ScaleWidth      =   1035
                  TabIndex        =   56
                  Top             =   300
                  Width           =   1035
                  Begin VB.OptionButton optDescPorc 
                     BackColor       =   &H00C0C0FF&
                     Caption         =   "%"
                     Height          =   210
                     Left            =   540
                     TabIndex        =   58
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   435
                  End
                  Begin VB.OptionButton optDescRS 
                     BackColor       =   &H00C0C0FF&
                     Caption         =   "R$"
                     Height          =   210
                     Left            =   0
                     TabIndex        =   57
                     TabStop         =   0   'False
                     Top             =   0
                     Value           =   -1  'True
                     Width           =   555
                  End
               End
               Begin VB.TextBox txtDesc 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   55
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Desc.:"
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
                  Left            =   2580
                  TabIndex        =   63
                  Top             =   300
                  Width           =   570
               End
               Begin VB.Label Label30 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total:"
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
                  Left            =   5160
                  TabIndex        =   62
                  Top             =   300
                  Width           =   510
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "SubTotal:"
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
                  Left            =   180
                  TabIndex        =   61
                  Top             =   300
                  Width           =   840
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00C0C0FF&
               Height          =   555
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   7275
               Begin VB.OptionButton optCheque 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Cheque"
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
                  Left            =   2640
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.OptionButton optCartao 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Cartão de Crédito"
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
                  Left            =   3720
                  TabIndex        =   52
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1995
               End
               Begin VB.OptionButton optPromissoria 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Promissória"
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
                  Left            =   1200
                  TabIndex        =   51
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.OptionButton optAvulso 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Avulso"
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
                  TabIndex        =   50
                  TabStop         =   0   'False
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   975
               End
            End
         End
         Begin VB.Frame frmVendaAvista 
            BackColor       =   &H00C0FFC0&
            Caption         =   "VENDA À VISTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   1620
            TabIndex        =   31
            Top             =   3780
            Visible         =   0   'False
            Width           =   7515
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0FFC0&
               Height          =   735
               Left            =   120
               TabIndex        =   36
               Top             =   900
               Width           =   7275
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H00C0FFC0&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   3180
                  ScaleHeight     =   210
                  ScaleWidth      =   1035
                  TabIndex        =   40
                  Top             =   300
                  Width           =   1035
                  Begin VB.OptionButton optDescRSAV 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "R$"
                     Height          =   210
                     Left            =   0
                     TabIndex        =   42
                     TabStop         =   0   'False
                     Top             =   0
                     Value           =   -1  'True
                     Width           =   555
                  End
                  Begin VB.OptionButton optDescPorcAV 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "%"
                     Height          =   210
                     Left            =   540
                     TabIndex        =   41
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   435
                  End
               End
               Begin VB.TextBox txtTotalDescAV 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   5700
                  Locked          =   -1  'True
                  TabIndex        =   39
                  TabStop         =   0   'False
                  Text            =   "0,00"
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.TextBox txtSubTotalAV 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   38
                  TabStop         =   0   'False
                  Text            =   "0,00"
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.TextBox txtDescAV 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   37
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "SubTotal:"
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
                  Left            =   180
                  TabIndex        =   45
                  Top             =   300
                  Width           =   840
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total:"
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
                  Left            =   5160
                  TabIndex        =   44
                  Top             =   300
                  Width           =   510
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Desc.:"
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
                  Index           =   1
                  Left            =   2580
                  TabIndex        =   43
                  Top             =   300
                  Width           =   570
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00C0FFC0&
               Height          =   555
               Left            =   120
               TabIndex        =   32
               Top             =   300
               Width           =   7275
               Begin VB.OptionButton optAVdinheiro 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Dinheiro"
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
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1095
               End
               Begin VB.OptionButton optAVcartao 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Cartão de Crédito"
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
                  Left            =   2340
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1995
               End
               Begin VB.OptionButton optAVcheque 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Cheque"
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
                  Left            =   1260
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin ChamaleonBtn.chameleonButton cmdAVfinalizar 
               Height          =   315
               Left            =   5280
               TabIndex        =   46
               Top             =   1740
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Finalizar"
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
               MICON           =   "PDV_Cadastro.frx":8C6C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdAVcancelar 
               Height          =   315
               Left            =   6360
               TabIndex        =   47
               Top             =   1740
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
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
               MICON           =   "PDV_Cadastro.frx":8C88
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
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   6135
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   10821
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label lbltotalGridProdutos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "R$ 0,00"
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
            Height          =   195
            Left            =   9960
            TabIndex        =   16
            Top             =   6240
            Width           =   690
         End
      End
      Begin VB.PictureBox frmProduto 
         Enabled         =   0   'False
         Height          =   1155
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   10695
         TabIndex        =   10
         Top             =   420
         Width           =   10755
         Begin VB.TextBox txtCodProduto 
            Appearance      =   0  'Flat
            Height          =   195
            Left            =   5280
            TabIndex        =   24
            Top             =   60
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   300
            Width           =   5715
         End
         Begin VB.TextBox txtQuant 
            Height          =   315
            Left            =   6960
            TabIndex        =   3
            Top             =   300
            Width           =   675
         End
         Begin VB.TextBox txtPreco 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5820
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtTotalPeca 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9300
            TabIndex        =   5
            Text            =   "0,00"
            Top             =   300
            Width           =   1335
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionar 
            Height          =   315
            Left            =   7320
            TabIndex        =   6
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Adicionar"
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
            MICON           =   "PDV_Cadastro.frx":8CA4
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
            Left            =   9000
            TabIndex        =   8
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Remover"
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
            MICON           =   "PDV_Cadastro.frx":8CC0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   675
            Left            =   7620
            ScaleHeight     =   675
            ScaleWidth      =   1695
            TabIndex        =   26
            Top             =   0
            Width           =   1695
            Begin VB.TextBox txtDesconto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   60
               TabIndex        =   4
               Text            =   "0,00"
               Top             =   300
               Width           =   1575
            End
            Begin VB.OptionButton optP 
               Caption         =   "%"
               Height          =   210
               Left            =   1140
               TabIndex        =   27
               Top             =   60
               Width           =   465
            End
            Begin VB.OptionButton optR 
               Caption         =   "R$"
               Height          =   210
               Left            =   600
               TabIndex        =   28
               Top             =   60
               Value           =   -1  'True
               Width           =   585
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desc.:"
               Height          =   195
               Left            =   60
               TabIndex        =   29
               Top             =   60
               Width           =   465
            End
         End
         Begin VB.Label lblAviso 
            AutoSize        =   -1  'True
            Caption         =   "Pressione a tecla  [ F2 ]  para obter os produtos."
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
            Left            =   60
            TabIndex        =   21
            Top             =   660
            Visible         =   0   'False
            Width           =   5820
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant."
            Height          =   195
            Left            =   6960
            TabIndex        =   13
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço"
            Height          =   195
            Left            =   5820
            TabIndex        =   12
            Top             =   60
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   195
            Left            =   9300
            TabIndex        =   11
            Top             =   60
            Width           =   360
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdAbrirVenda 
         Height          =   555
         Left            =   120
         TabIndex        =   0
         Top             =   8220
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "ABRIR VENDA"
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
         MICON           =   "PDV_Cadastro.frx":8CDC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdClientes 
         Height          =   555
         Left            =   7560
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Cadastro de Clientes"
         Top             =   8220
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Clientes"
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV_Cadastro.frx":8CF8
         PICN            =   "PDV_Cadastro.frx":8D14
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdProdutos 
         Height          =   555
         Left            =   9240
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Cadastro de Clientes"
         Top             =   8220
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   979
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV_Cadastro.frx":95D3
         PICN            =   "PDV_Cadastro.frx":95EF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdFinalizarPedido 
         Height          =   555
         Left            =   2100
         TabIndex        =   30
         Top             =   8220
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "FINALIZAR PEDIDO"
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
         MICON           =   "PDV_Cadastro.frx":9C63
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
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   25
      Top             =   9915
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15663
            Text            =   "Desenv.: OnLine.Info - Informática & Lan House - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: OnLine.Info - Informática & Lan House - Tel.: (89) 3544-2553"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "01:13"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
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
   Begin VB.Menu Menu_Pedidos 
      Caption         =   "Pedidos"
      Visible         =   0   'False
      Begin VB.Menu Menu_Novo 
         Caption         =   "Novo Pedido"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Menu_Finalizar 
         Caption         =   "Finalizar Pedido"
      End
      Begin VB.Menu Menu_Cancelar 
         Caption         =   "Cancelar Pedido"
      End
   End
   Begin VB.Menu Menu_Lista 
      Caption         =   "Lista"
      Visible         =   0   'False
      Begin VB.Menu Menu_Adicionar 
         Caption         =   "Adicionar à Lista"
      End
      Begin VB.Menu Menu_Remover 
         Caption         =   "Remover da Lista"
      End
   End
   Begin VB.Menu Menu_Ajuda 
      Caption         =   "Ajuda"
      Visible         =   0   'False
      Begin VB.Menu Menu_Teclas 
         Caption         =   "Teclas"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "PDV_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim RS As Recordset
Dim CAIXA_FECHADO As Boolean
Private moCombo As cComboHelper
Private Sub Abrir_Pedido()
    Call ABRIR_BD_SEM_DATA1
    SQL = "SELECT * FROM PEDIDOS WHERE COD_PEDIDO = " & txtCodPedido.Text & ""
    Set RS = BD.OpenRecordset(SQL)
    RS.Edit
        RS!STATUS_PEDIDO = False
    RS.Update
End Sub

Private Sub LimparGrid_Produtos()
Call ABRIR_BD_SEM_DATA1
SQL = "SELECT * FROM PEDIDOS_ITENS WHERE FALSE"
Set RS = BD.OpenRecordset(SQL)
FormatarGrid_Produtos
'BD.Close
End Sub
Private Sub MostrarDados_Pedido()
Call ABRIR_BD_SEM_DATA1
SQL = "SELECT * FROM PEDIDOS WHERE COD_PEDIDO = " & txtCodPedido.Text
Set RS = BD.OpenRecordset(SQL)

If Not IsNull(RS!TIPO_PAGAMENTO) Then
    If RS!TIPO_PAGAMENTO = "À Vista" Then
        frmVendaAvista.Visible = True
        frmVendaPrazo.Visible = False
        If RS!PAGAMENTO = "DINHEIRO" Then optAVdinheiro.Value = True
        If RS!PAGAMENTO = "CHEQUE" Then optAVcheque.Value = True
        If RS!PAGAMENTO = "CARTAO" Then optAVcartao.Value = True
        If Not IsNull(RS!SUBTOTAL) Then txtSubTotalAV.Text = Format(RS!SUBTOTAL, "##,##0.00")
        If RS!TIPO_DESC = "R" Then optDescRSAV.Value = True Else optDescPorcAV.Value = True
        If Not IsNull(RS!VALOR_DESC) Then txtDescAV.Text = Format(RS!VALOR_DESC, "##,##0.00")
        If Not IsNull(RS!Total) Then txtTotalDescAV.Text = Format(RS!Total, "##,##0.00")
    ElseIf RS!TIPO_PAGAMENTO = "À Prazo" Then
        frmVendaAvista.Visible = False
        frmVendaPrazo.Visible = True
        If RS!PAGAMENTO = "AVULSO" Then optAvulso.Value = True
        If RS!PAGAMENTO = "PROMISSORIA" Then optPromissoria.Value = True
        If RS!PAGAMENTO = "CHEQUE" Then optCheque.Value = True
        If RS!PAGAMENTO = "CARTAO" Then optCartao.Value = True
        If Not IsNull(RS!SUBTOTAL) Then txtSubTotal.Text = Format(RS!SUBTOTAL, "##,##0.00")
        If RS!TIPO_DESC = "R" Then optDescRS.Value = True Else optDescPorc.Value = True
        If Not IsNull(RS!VALOR_DESC) Then txtDesc.Text = Format(RS!VALOR_DESC, "##,##0.00")
        If Not IsNull(RS!Total) Then txtTotalDesc.Text = Format(RS!Total, "##,##0.00")
        If Not IsNull(RS!COD_CLIENTE) Then txtCodCliente.Text = RS!COD_CLIENTE
        If Not IsNull(RS!ENTRADA) Then txtEntrada.Text = Format(RS!ENTRADA, "##,##0.00")
        If Not IsNull(RS!PRAZO) Then cboPrazo.Text = RS!PRAZO
        If Not IsNull(RS!VALOR_PARC) Then txtValorParc.Text = Format(RS!VALOR_PARC, "##,##0.00")
        If Not IsNull(RS!VENCIMENTO) Then mskVencimento.Text = Format(RS!VENCIMENTO, "dd/mm/yy")

    End If
End If

If Not IsNull(RS!DATA_COMPRA) Then txtData.Text = Format(RS!DATA_COMPRA, "dd/mm/yy")


End Sub

Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Double
Dim i As Integer, Valor_Col As Double
For i = 0 To Grid.Rows - 1
  If IsNumeric(Grid.TextMatrix(i, Col)) Then
    Valor_Col = Valor_Col + CDbl(Grid.TextMatrix(i, Col))
  End If
Next i
SomaGrid = Valor_Col
End Function



Private Sub cmdAbrirVenda_Click()
'abre o caixa do pedido
Abrir_Caixa

'abre o pedido
Abrir_Pedido

'Apaga as parcelas do pedido
execSQL "DELETE FROM PARCELAS WHERE COD_PEDIDO = " & txtCodPedido.Text

cmdFinalizarPedido.Visible = True
cmdAbrirVenda.Visible = False

frmProduto.Enabled = True
txtDescricao.SetFocus
End Sub

Private Sub cmdClientes_Click()
Clientes_Cadastro.Show 1
End Sub
Private Sub Abrir_Caixa()
Call ABRIR_BD_SEM_DATA1
SQL = "SELECT * FROM CAIXA_DIA WHERE DATA = #" & Format(txtData.Text, "mm/dd/yy") & "# AND STATUS = TRUE"
Set RS = BD.OpenRecordset(SQL)

If RS.RecordCount <> 0 Then
    Call ABRIR_BD_SEM_DATA1
    SQL = "SELECT * FROM CAIXA_DIA WHERE DATA = #" & Format(txtData.Text, "mm/dd/yy") & "#"
    Set RS = BD.OpenRecordset(SQL)
    RS.Edit
        RS!Status = False
    RS.Update
End If
End Sub

Private Sub cmdFinalizarPedido_Click()
MostrarDados_Pedido
End Sub

Private Sub cmdProdutos_Click()
Produtos_Cadastro.Show 1
End Sub
Private Sub cmdRemover_Click()
On Error GoTo Erro

If Grid.TextMatrix(Grid.Row, 1) = "" Then GoSub Erro
If MsgBox("Deseja excluir o produto : " & Grid.TextMatrix(Grid.Row, 4) & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso do Sistema") = vbNo Then Exit Sub

execSQL "DELETE FROM PEDIDOS_ITENS WHERE CODIGO = " & Grid.TextMatrix(Grid.Row, 1) & " AND COD_PEDIDO = " & txtCodPedido.Text

PreencherGrid_Produtos

Exit Sub
Erro:
    MsgBox "Não existe nenhum produto para ser excluido!", vbExclamation, "Aviso do Sistema"
    Exit Sub
End Sub

Private Sub Checar_PEDIDOS()
'ABRIR_BD_com_Data Me.Data2
'Data2.RecordSource = "SELECT * FROM PEDIDOS WHERE COD_PEDIDO = " & txtCodPedido.Text & " AND STATUS_PEDIDO = true"
'Data2.Refresh

'If Data2.Recordset.RecordCount > 0 Then
'    MsgBox "Essa PEDIDOS encontra-se fechada para alterações!", vbInformation, "Aviso do Sistema"
'    Exit Sub
'End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
'If SSTab1.Tab = 1 Then cboCliente.SetFocus
End Sub

Private Sub txtCodCliente_Change()
If txtCodCliente.Text = "" Then Exit Sub

Dim SQL_Cliente As String
Dim RS_Cliente As Recordset
Call ABRIR_BD_SEM_DATA1
SQL_Cliente = "SELECT * FROM CLIENTE WHERE CODIGO = " & txtCodCliente.Text
Set RS_Cliente = BD.OpenRecordset(SQL_Cliente)

If Not IsNull(RS_Cliente!NOME) Then cboCliente.Text = RS_Cliente!NOME

End Sub

Private Sub txtCodPedido_Change()
If txtCodPedido.Text = "" Then Exit Sub
'MostrarDados_Pedido
Call ABRIR_BD_SEM_DATA1
SQL = "SELECT * FROM PEDIDOS WHERE COD_PEDIDO = " & txtCodPedido.Text
Set RS = BD.OpenRecordset(SQL)
If Not IsNull(RS!DATA_COMPRA) Then txtData.Text = Format(RS!DATA_COMPRA, "dd/mm/yy")


PreencherGrid_Produtos
End Sub
Private Sub txtDescricao_LostFocus()
lblAviso.Visible = False
End Sub
Private Sub Form_Load()
SSTab1.Tab = 0
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
'cmdNovo_Click
Set moCombo = New cComboHelper
End Sub
Private Sub FormatarGrid_Produtos()
With Grid
    
    .Clear
    .Cols = 10
    .Rows = 2
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 0
    .ColWidth(3) = 0
    .ColWidth(4) = 4500
    .ColWidth(5) = 1200
    .ColWidth(6) = 900
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1300


    
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "COD_PEDIDO"
    .TextMatrix(0, 3) = "COD_PRODUTO"
    .TextMatrix(0, 4) = "DESCRICAO"
    .TextMatrix(0, 5) = "PREÇO"
    .TextMatrix(0, 6) = "QUANT."
    .TextMatrix(0, 7) = "SUBTOTAL"
    .TextMatrix(0, 8) = "DESC."
    .TextMatrix(0, 9) = "TOTAL"


    
    'colocar os cabeçalho em negrito
    For x = 0 To .Cols - 1
    .Col = x
    .Row = 0
    .CellFontBold = True
    Next x
    
    'ALINHAMENTO
    '.ColAlignment(2) = 1
    
    'centralizar o titulo
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until RS.EOF
    
    'mudar a cor da coluna
    Dim i
    .Redraw = False
    For i = 1 To .Rows - 1
   .Row = i
   .Col = 5:   .CellBackColor = &HC0FFFF
   .Col = 9:   .CellBackColor = &HC0C0FF
    Next
    
    Grid.Redraw = False
    
  
    Grid.Redraw = True
    
    If Not IsNull(RS!CODIGO) Then .TextMatrix(.Rows - 1, 1) = RS!CODIGO
    If Not IsNull(RS!Cod_Pedido) Then .TextMatrix(.Rows - 1, 2) = RS!Cod_Pedido
    If Not IsNull(RS!COD_PRODUTO) Then .TextMatrix(.Rows - 1, 3) = RS!COD_PRODUTO
    If Not IsNull(RS!DESCRICAO) Then .TextMatrix(.Rows - 1, 4) = RS!DESCRICAO
    If Not IsNull(RS!Preco) Then .TextMatrix(.Rows - 1, 5) = Format(RS!Preco, "##,##0.00")
    If Not IsNull(RS!QUANTIDADE) Then .TextMatrix(.Rows - 1, 6) = RS!QUANTIDADE
    If Not IsNull(RS!SUBTOTAL) Then .TextMatrix(.Rows - 1, 7) = Format(RS!SUBTOTAL, "##,##0.00")
    If Not IsNull(RS!VALOR_DESC) Then .TextMatrix(.Rows - 1, 8) = Format(RS!VALOR_DESC, "##,##0.00")
    If Not IsNull(RS!LIQUIDO) Then .TextMatrix(.Rows - 1, 9) = Format(RS!LIQUIDO, "##,##0.00")

    RS.MoveNext
    .Rows = .Rows + 1
        
    Loop
    
    .Rows = .Rows - 1

End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
'CHECAR SE O PEDIDO ESTÁ FECHADA
If txtCodPedido.Text = "" Then Exit Sub

'If Grid.Rows >= 1 And cmdNovo.Enabled = False Then
'    If MsgBox("Existe um pedido em aberto. Deseja sair e cancelar o pedido?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso do Sistema") = vbNo Then Cancel = 1: Exit Sub
'    execSQL "DELETE FROM PEDIDOS_ITENS WHERE COD_PEDIDO = " & txtCodPedido.Text
'    execSQL "DELETE FROM PEDIDOS WHERE COD_PEDIDO = " & txtCodPedido.Text
'End If
Set moCombo = Nothing
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    'cmdAdicionar_Click
    End If
End Sub
Private Sub txtDescricao_GotFocus()
txtDescricao.SelStart = 0
txtDescricao.SelLength = Len(txtDescricao)
lblAviso.Visible = True
End Sub
Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    Vendas_Consulta.Show 1
End If
End Sub
Sub PreencherGrid_Produtos()
Call ABRIR_BD_SEM_DATA1
SQL = "SELECT DESCRICAO, QUANTIDADE, PRECO, (QUANTIDADE * PRECO) AS [SUBTOTAL], VALOR_DESC, IIf(TIPO_DESC = 'R', (QUANTIDADE * PRECO) - VALOR_DESC, TOTAL - ((QUANTIDADE * PRECO) * VALOR_DESC) / 100) as [LIQUIDO], CODIGO, COD_PRODUTO, COD_PEDIDO FROM PEDIDOS_ITENS WHERE COD_PEDIDO = " & txtCodPedido.Text
Set RS = BD.OpenRecordset(SQL)

FormatarGrid_Produtos

lbltotalGridProdutos.Caption = Format(SomaGrid(Grid, 9), "##,##0.00")
'txtSubTotal.Text = Format(SomaGrid(Grid, 9), "##,##0.00")
End Sub
Private Sub txtPreco_GotFocus()
txtPreco.SelStart = 0
txtPreco.SelLength = Len(txtPreco.Text)
End Sub
Private Sub txtPreco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    'cmdAdicionar_Click
    txtDescricao.SetFocus
    End If

KeyAscii = aNumeros(KeyAscii, True)

End Sub
Private Sub txtPreco_LostFocus()
txtPreco.Text = Format(txtPreco.Text, "##,##0.00")
End Sub
Private Sub txtQuant_Change()
If txtQuant.Text = "" Or txtPreco.Text = "" Then txtTotalPeca.Text = "0,00": Exit Sub

'txtDesconto_Change
''txtTotalPeca.Text = Format(txtQuant.Text * CDbl(txtPreco.Text), "##,##0.00")
End Sub
Private Sub txtQuant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdicionar.SetFocus
    'cmdAdicionar_Click
    txtDescricao.SetFocus
End If
    
KeyAscii = aNumeros(KeyAscii, True)
End Sub
Private Sub txtQuant_LostFocus()
If txtQuant.Text = "" Or txtQuant.Text = "0" Then
    If txtDescricao.Text <> "" Then
    txtQuant.Text = 1
    End If
End If
End Sub
