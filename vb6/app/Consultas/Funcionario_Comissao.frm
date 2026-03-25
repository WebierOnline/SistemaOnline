VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Funcionario_Comissao 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FUNCIONÁRIO - CONSULTA DE COMISSÃO"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   13215
   Icon            =   "Funcionario_Comissao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "COMISSÕES - Resultado Geral Mensal"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2940
      TabIndex        =   39
      Top             =   3000
      Width           =   10155
      Begin VB.Frame Frame6 
         Caption         =   "Serviços"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   6180
         TabIndex        =   55
         Top             =   240
         Width           =   1995
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   660
            TabIndex        =   59
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblComServicos 
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
            Left            =   660
            TabIndex        =   58
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label lblComServicosQtde 
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
            Left            =   60
            TabIndex        =   57
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   60
            TabIndex        =   56
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Recebidos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   4140
         TabIndex        =   50
         Top             =   240
         Width           =   1995
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   660
            TabIndex        =   54
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblComRecebido 
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
            Left            =   660
            TabIndex        =   53
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label lblComRecebidoQtde 
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
            Left            =   60
            TabIndex        =   52
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   60
            TabIndex        =   51
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Vendas á Prazo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   2100
         TabIndex        =   45
         Top             =   240
         Width           =   1995
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   660
            TabIndex        =   49
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblComAPrazo 
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
            Left            =   660
            TabIndex        =   48
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label lblComAPrazoQtde 
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
            Left            =   60
            TabIndex        =   47
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   60
            TabIndex        =   46
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Vendas á Vista"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   60
         TabIndex        =   40
         Top             =   240
         Width           =   1995
         Begin VB.Label lblComAvistaQtde 
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
            Left            =   60
            TabIndex        =   44
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   660
            TabIndex        =   43
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblComAvista 
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
            Left            =   660
            TabIndex        =   42
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   60
            TabIndex        =   41
            Top             =   240
            Width           =   330
         End
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
      Height          =   3255
      Left            =   60
      TabIndex        =   27
      Top             =   900
      Width           =   2835
      Begin VB.ComboBox cboTipoPgto 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   2220
         Width           =   2595
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   2595
      End
      Begin VB.ComboBox cboCriterioSec 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   1020
         Width           =   2595
      End
      Begin VB.ComboBox cboIndice 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   2820
         Width           =   2595
      End
      Begin VB.ComboBox cboFormaPgto 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1620
         Width           =   2595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pagamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1980
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Consulta"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Criterio"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Organizar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2580
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pgto"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1380
         Width           =   1035
      End
   End
   Begin VB.Frame Frame8 
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
      Height          =   2055
      Left            =   2940
      TabIndex        =   12
      Top             =   900
      Width           =   10215
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   4905
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   1500
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   1140
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1140
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtCodFunc 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   4320
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendario2 
         Height          =   315
         Left            =   2700
         TabIndex        =   17
         Tag             =   "Calendario"
         Top             =   1140
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
         MICON           =   "Funcionario_Comissao.frx":23D2
         PICN            =   "Funcionario_Comissao.frx":23EE
         PICH            =   "Funcionario_Comissao.frx":4741
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
         TabIndex        =   18
         Tag             =   "Calendario"
         Top             =   1140
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
         MICON           =   "Funcionario_Comissao.frx":6A94
         PICN            =   "Funcionario_Comissao.frx":6AB0
         PICH            =   "Funcionario_Comissao.frx":8E03
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1140
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
         TabIndex        =   20
         Top             =   1140
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton chameleonButton1 
         Height          =   495
         Left            =   3060
         TabIndex        =   36
         Top             =   960
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
         MICON           =   "Funcionario_Comissao.frx":B156
         PICN            =   "Funcionario_Comissao.frx":B172
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblVendedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor(a):"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lblAte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         Height          =   195
         Left            =   1440
         TabIndex        =   25
         Top             =   1200
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblFim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data final:"
         Height          =   195
         Left            =   1740
         TabIndex        =   24
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblInicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   900
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblAno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano:"
         Height          =   195
         Left            =   1500
         TabIndex        =   22
         Top             =   900
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblMes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   900
         Visible         =   0   'False
         Width           =   345
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
      Height          =   255
      Left            =   6900
      TabIndex        =   6
      Top             =   8820
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
      MICON           =   "Funcionario_Comissao.frx":BA4C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4740
      Picture         =   "Funcionario_Comissao.frx":BA68
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   2
      Top             =   5820
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   13065
      TabIndex        =   0
      Top             =   60
      Width           =   13095
      Begin VB.Image Image1 
         Height          =   720
         Left            =   540
         Picture         =   "Funcionario_Comissao.frx":CAA0
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FUNCIONÁRIO - CONSULTA DE COMISSÃO"
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
         TabIndex        =   1
         Top             =   180
         Width           =   6420
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   9135
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
            TextSave        =   "13:50"
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
      Height          =   4095
      Left            =   60
      TabIndex        =   7
      Top             =   4200
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7223
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirProdutos 
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   8340
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "EXIBIR PRODUTOS"
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
      MICON           =   "Funcionario_Comissao.frx":D57C
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
      Height          =   315
      Left            =   1860
      TabIndex        =   9
      Top             =   8340
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "EXIBIR PARCELAS"
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
      MICON           =   "Funcionario_Comissao.frx":D598
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
      Left            =   3720
      TabIndex        =   60
      Top             =   8340
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "IMPRIMIR"
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
      MICON           =   "Funcionario_Comissao.frx":D5B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCriarPDF 
      Height          =   315
      Left            =   5580
      TabIndex        =   61
      Top             =   8340
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "CRIAR PDF"
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
      MICON           =   "Funcionario_Comissao.frx":D5D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label11 
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
      Left            =   10680
      TabIndex        =   11
      Top             =   8700
      Width           =   510
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
      Left            =   11280
      TabIndex        =   10
      Top             =   8700
      Width           =   1815
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
      Left            =   11280
      TabIndex        =   4
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant. Parc.:"
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
      Left            =   10080
      TabIndex        =   3
      Top             =   8400
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   735
      Left            =   9960
      Top             =   8340
      Width           =   3195
   End
End
Attribute VB_Name = "Funcionario_Comissao"
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
Private Sub Limpar_Grid()
Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 10
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 1000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 700
      .ColWidth(5) = 4300
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1000
     
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "TIPO"
      .TextMatrix(0, 3) = "DATA"
      .TextMatrix(0, 4) = "PARC."
      .TextMatrix(0, 5) = "NOME DO CLIENTE"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "STATUS"
      .TextMatrix(0, 8) = "PGTO"
      .TextMatrix(0, 9) = "FORMA"
      

      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 7
      .ColAlignment(2) = 7
      .ColAlignment(3) = 7
      .ColAlignment(4) = 7
      .ColAlignment(5) = 1
      .ColAlignment(6) = 7
      .ColAlignment(7) = 7
      .ColAlignment(8) = 7
      .ColAlignment(9) = 7
      
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
   
   
    'lblSubtotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
    'lblSubtotalBruto.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
    'lblAcresc.Caption = Format(SomaGrid(Grid, 6), ocMONEY)
    'lblTotal.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
    'lblEntrada.Caption = Format(0, ocMONEY)
picAguarde.Visible = False

End Sub

Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
cboVendedor.Text = ""
'txtCodigo.Text = ""
'cboCliente.Text = ""
mskFim.Mask = ""
mskFim.Text = ""
mskInicio.Mask = ""
mskInicio.Text = ""
txtCodFunc.Text = ""
'txtCodCliente.Text = ""
End Sub


Private Sub PreencherCriterios()
cboCriterioSec.Clear
cboCriterioSec.AddItem "MENSAL"
End Sub

Private Sub PreencherTipoPgto()
End Sub

Private Sub PreencherIndice()
cboIndice.Clear
cboIndice.AddItem "PEDIDO"
cboIndice.AddItem "PGTO."
cboIndice.AddItem "FORMA PGTO"
cboIndice.AddItem "VALOR"
End Sub

Private Sub PreencherFormaPgto()
cboFormaPgto.Clear
cboFormaPgto.AddItem "À VISTA"
cboFormaPgto.AddItem "À PRAZO"
cboFormaPgto.AddItem "RECEBIDOS"
cboFormaPgto.AddItem "SERVIÇOS"
End Sub

Private Sub PreencherTipoConsulta()
cboTipo.Clear
cboTipo.AddItem "VENDA"
cboTipo.AddItem "SERVIÇOS"
End Sub

Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer, vUltimoNome As String

vUltimoNome = cboAno
cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAno.AddItem i
Next

cboAno.Text = vUltimoNome

moCombo.AttachTo cboAno
End Sub

Private Sub cboAno_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboCriterioSec_Change()
If cboCriterioSec.Text = "TODOS" Then
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblInicio.Visible = False
    lblFim.Visible = False
    lblAte.Visible = False
    mskInicio.Visible = False
    mskFim.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
End If

If cboCriterioSec.Text = "MENSAL" Then
    lblMes.Visible = True
    cboMes.Visible = True
    lblAno.Visible = True
    cboAno.Visible = True
    lblInicio.Visible = False
    lblFim.Visible = False
    lblAte.Visible = False
    mskInicio.Visible = False
    mskFim.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
End If

If cboCriterioSec.Text = "PERÍODO" Then
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblInicio.Visible = True
    lblFim.Visible = True
    lblAte.Visible = True
    mskInicio.Visible = True
    mskFim.Visible = True
    cmdCalendario1.Visible = True
    cmdCalendario2.Visible = True
End If
End Sub

Private Sub cboCriterioSec_Click()
   cboCriterioSec_Change
End Sub

Private Sub cboCriterioSec_GotFocus()
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
'cboAno.SetFocus
End Sub

Private Sub cboTipo_Change()
If cboTipo.Text = "VENDA" Then
   cmdExibirProdutos.Visible = True
   cmdExibirParcelas.Visible = True
   cmdExibirPedidos.Visible = False
ElseIf cboTipo.Text = "SERVIÇOS" Then
   cmdExibirProdutos.Visible = True
   cmdExibirParcelas.Visible = True
   cmdExibirPedidos.Visible = False
Else
   Exit Sub
End If
End Sub

Private Sub cboTipo_Click()
cboTipo_Change
End Sub

Private Sub cboTipo_GotFocus()
moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipoPgto_GotFocus()
cboTipoPgto.Clear
cboTipoPgto.AddItem "TODOS"
cboTipoPgto.AddItem "DINHEIRO"
cboTipoPgto.AddItem "PIX"
cboTipoPgto.AddItem "CARTÃO DÉBITO"
cboTipoPgto.AddItem "CARTÃO CRÉDITO"
cboTipoPgto.AddItem "TRANSFERÊNCIA"
cboTipoPgto.AddItem "DEPOSITO"
cboTipoPgto.AddItem "FINANCEIRA"
cboTipoPgto.AddItem "CHEQUE"
cboTipoPgto.AddItem "BOLETO"
cboTipoPgto.AddItem "PROMISSÓRIA"
moCombo.AttachTo cboTipoPgto
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
   'If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboVendedor_LostFocus()
   On Error GoTo TrataErro
   
   If cboVendedor.Text = "" Then txtCodFunc.Text = "": Exit Sub
   If cboVendedor.ListIndex = -1 Then txtCodFunc.Text = "": Exit Sub
   txtCodFunc = cboVendedor.ItemData(cboVendedor.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub chameleonButton1_Click()
If txtCodFunc.Text = "" Then Exit Sub
'INDICE===========================
Dim INDICE As String
If cboIndice.Text = "PGTO." Then
   INDICE = "parcelas.PAGAMENTO;"
ElseIf cboIndice.Text = "FORMA PGTO" Then
   INDICE = "parcelas.FORMA_PGTO;"
ElseIf cboIndice.Text = "VALOR" Then
   INDICE = "parcelas.VALOR_FINAL"
ElseIf cboIndice.Text = "PEDIDO" Then
   INDICE = "parcelas.COD_PEDIDO, parcelas.NUMERO"
Else
   INDICE = "parcelas.COD_PEDIDO, parcelas.NUMERO"
End If

'FORMA DE PAGAMENTO ===============
Dim vPago As String
If cboFormaPgto.Text = "À VISTA" Then
   vPago = " AND (parcelas.status = 1)"
ElseIf cboFormaPgto.Text = "À PRAZO" Then
   vPago = " AND (parcelas.status IN (1, 0))"
ElseIf cboFormaPgto.Text = "RECEBIDOS" Then
   vPago = " AND (parcelas.status = 1)"
Else
    vPago = " AND (parcelas.status IN (1, 0))"
End If

'FORMA DE PAGAMENTO ===============
Dim TipoPgto As String
If cboFormaPgto.Text = "À VISTA" Then
   TipoPgto = " AND (pedidos.TIPO_PAGAMENTO = 'À Vista')"
ElseIf cboFormaPgto.Text = "À PRAZO" Then
   TipoPgto = " AND (pedidos.TIPO_PAGAMENTO = 'À prazo')"
ElseIf cboFormaPgto.Text = "RECEBIDOS" Then
   TipoPgto = " AND (pedidos.TIPO_PAGAMENTO = 'À prazo')"
Else
    TipoPgto = " AND (pedidos.TIPO_PAGAMENTO IN ('À Vista', 'À prazo'))"
End If

'TIPO DE PAGAMENTO ===================
Dim vTipoPgtoParcelas As String
If cboTipoPgto.Text = "TODOS" Then
   vTipoPgtoParcelas = ""
ElseIf cboTipoPgto.Text = "DINHEIRO" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'DINHEIRO')"
ElseIf cboTipoPgto.Text = "CARTÃO DÉBITO" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CARTAO') and (parcelas.TIPO_CARTAO = 'D')"
ElseIf cboTipoPgto.Text = "CARTÃO CRÉDITO" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CARTAO') and (parcelas.TIPO_CARTAO = 'C')"
ElseIf cboTipoPgto.Text = "TRANSFERÊNCIA" Then
   vTipoPgtoParcelas = "  AND (parcelas.FORMA_PGTO = 'TRANSFERENCIA')"
ElseIf cboTipoPgto.Text = "DEPOSITO" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'DEPOSITO')"
ElseIf cboTipoPgto.Text = "FINANCEIRA" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'FINANCEIRA')"
ElseIf cboTipoPgto.Text = "PROMISSÓRIA" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'PROMISSORIA')"
ElseIf cboTipoPgto.Text = "CHEQUE" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'Cheque')"
ElseIf cboTipoPgto.Text = "BOLETO" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'BOLETO')"
ElseIf cboTipoPgto.Text = "PIX" Then
   vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'PIX')"
Else
    vTipoPgtoParcelas = ""
End If


Dim vTipoCriterio As String
If cboCriterioSec.Text = "TODOS" Then
    vTipoCriterio = ""
ElseIf cboCriterioSec.Text = "MENSAL" Then

    If cboMes.Text = "" Or cboAno.Text = "" Then Limpar_Grid: Exit Sub

    Dim vIndMes As Integer
    If cboMes.ListCount = 0 Then
        If cboMes.Text = "janeiro" Then
            vIndMes = cboMes.ListIndex + 2
        ElseIf cboMes.Text = "fevereiro" Then
            vIndMes = cboMes.ListIndex + 3
        ElseIf cboMes.Text = "março" Then
            vIndMes = cboMes.ListIndex + 4
        ElseIf cboMes.Text = "abril" Then
            vIndMes = cboMes.ListIndex + 5
        ElseIf cboMes.Text = "maio" Then
            vIndMes = cboMes.ListIndex + 6
        ElseIf cboMes.Text = "junho" Then
            vIndMes = cboMes.ListIndex + 7
        ElseIf cboMes.Text = "julho" Then
            vIndMes = cboMes.ListIndex + 8
        ElseIf cboMes.Text = "agosto" Then
            vIndMes = cboMes.ListIndex + 9
        ElseIf cboMes.Text = "setembro" Then
            vIndMes = cboMes.ListIndex + 10
        ElseIf cboMes.Text = "outubro" Then
            vIndMes = cboMes.ListIndex + 11
        ElseIf cboMes.Text = "novembro" Then
            vIndMes = cboMes.ListIndex + 12
        ElseIf cboMes.Text = "dezembro" Then
            vIndMes = cboMes.ListIndex + 13
        End If
    
        If cboFormaPgto.Text = "À VISTA" Then
            vTipoCriterio = " AND (Month(parcelas.PAGAMENTO) = " & vIndMes & ") And (Year(parcelas.PAGAMENTO) = " & cboAno & ")"
        ElseIf cboFormaPgto.Text = "À PRAZO" Then
            vTipoCriterio = " AND (Month(pedidos.DATA_COMPRA) = " & vIndMes & ") And (Year(pedidos.DATA_COMPRA) = " & cboAno & ")"
        ElseIf cboFormaPgto.Text = "RECEBIDOS" Then
            vTipoCriterio = " AND (Month(parcelas.PAGAMENTO) = " & vIndMes & ") And (Year(parcelas.PAGAMENTO) = " & cboAno & ")"
        End If
    
    Else
        If cboFormaPgto.Text = "À VISTA" Then
            vTipoCriterio = " AND (Month(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") And (Year(parcelas.PAGAMENTO) = " & cboAno & ")"
        ElseIf cboFormaPgto.Text = "À PRAZO" Then
            vTipoCriterio = " AND (Month(pedidos.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") And (Year(pedidos.DATA_COMPRA) = " & cboAno & ")"
        ElseIf cboFormaPgto.Text = "RECEBIDOS" Then
            vTipoCriterio = " AND (Month(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") And (Year(parcelas.PAGAMENTO) = " & cboAno & ")"
        End If
    End If

ElseIf cboCriterioSec.Text = "PERÍODO" Then
    If Not IsDate(mskInicio) Or Not IsDate(mskFim) Then Limpar_Grid: Exit Sub
    vTipoCriterio = " AND (parcelas.PAGAMENTO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas.PAGAMENTO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103))"
End If

'MONTAR O GRID
sSQL = "SELECT parcelas.COD_PEDIDO as var_codped, pedidos.DATA_COMPRA, parcelas.PAGAMENTO, parcelas.NUMERO, parcelas.VALOR_FINAL, parcelas.FORMA_PGTO as var_FormaPgto, (CASE WHEN parcelas.status = 1 THEN 'Pago' ELSE 'À Pagar' END) AS var_StatusPgto, pedidos.COD_FUNCIONARIO, cliente.Nome, pedidos.COD_CLIENTE, pedidos.TIPO_PAGAMENTO " & _
        "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
        "WHERE (pedidos.TIPO_PEDIDO = 'VENDA') AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") " & TipoPgto & " " & vPago & " AND (parcelas.STATUS = 1) " & vTipoCriterio & " "
Set r = dbData.OpenRecordset(sSQL, totalRegistros)
printSQL = sSQL '" & TipoPgto & "
'(pedidos.TIPO_PAGAMENTO = 'À Vista')
'Debug.Print sSQL
If Not r.EOF Then
    lblQtda.Caption = Format(totalRegistros, "00")
Else
    lblQtda.Caption = Format(0, "00")
End If

FormatarGrid r



'BUSCAR TOTAL DE AVISTA ===================================================================================
sSQL = "SELECT SUM(parcelas.VALOR_FINAL) AS varTotalAvista " & _
        "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
        "WHERE (pedidos.TIPO_PEDIDO = 'VENDA') AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") " & TipoPgto & " " & vPago & " " & vTipoPgtoParcelas & " " & vTipoCriterio & " "
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

Dim vValorTotalAvista As Currency
If Not r.EOF Then
    vValorTotalAvista = FormatNumber(ValidateNull(r("varTotalAvista")), 2)
Else
    vValorTotalAvista = FormatNumber(0, 2)
End If

'CONSULTAS COMISSÕES
sSQL = "SELECT Comissao_Avista1, Comissao_Avista2, Comissao_Avista3, Valor_Comissao1, Valor_Comissao2, Valor_Comissao3 " & _
       "FROM funcionario " & _
       "WHERE (CODIGO = " & txtCodFunc.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim vAlvoAvista As Currency
Dim vComissaoAvista As Currency

If Not r.EOF Then
    If vValorTotalAvista > r("Valor_Comissao1") Then
        If vValorTotalAvista < r("Valor_Comissao3") Then
            vComissaoAvista = FormatNumber(r("Comissao_Avista2"), 2)
        Else
            vComissaoAvista = FormatNumber(r("Comissao_Avista3"), 2)
        End If
    Else
        vComissaoAvista = FormatNumber(r("Comissao_Avista1"), 2)
    End If
Else
    vComissaoAvista = FormatNumber(0, 2)
End If

'COMISSÕES Á VISTA
sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL * " & Replace(CDbl(vComissaoAvista), ",", ".") & " / 100), 0) AS var_ComAvista, COUNT(parcelas.CODIGO) AS var_ContParcelas " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
                     "INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
                     "INNER JOIN funcionario ON pedidos.COD_FUNCIONARIO = funcionario.CODIGO " & _
       "WHERE (pedidos.TIPO_PEDIDO = 'VENDA') AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Vista') " & vPago & " " & vTipoPgtoParcelas & " AND (Month(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") And (Year(parcelas.PAGAMENTO) = " & cboAno & ") "
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If Not r.EOF Then
    lblComAvistaQtde.Caption = Format(r("var_ContParcelas"), "000")
    lblComAvista.Caption = FormatNumber(r("var_ComAvista"), 2)
Else
    lblComAvistaQtde.Caption = Format(0, "00")
    lblComAvista.Caption = FormatNumber(0, 2)
End If





'BUSCAR TOTAL DE RECEBIDO ===================================================================================
sSQL = "SELECT SUM(parcelas.VALOR_FINAL) AS varTotalRecebido " & _
        "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
        "WHERE (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Prazo') AND (parcelas.STATUS = 1) AND (Month(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") And (Year(parcelas.PAGAMENTO) = " & cboAno & ")"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)
'Debug.Print sSQL
Dim vValorTotalRecebido As Currency
If Not r.EOF Then
    vValorTotalRecebido = FormatNumber(ValidateNull(r("varTotalRecebido")), 2)
Else
    vValorTotalRecebido = FormatNumber(0, 2)
End If

'CONSULTAS COMISSÕES
sSQL = "SELECT Comissao_Recebido1, Comissao_Recebido2, Comissao_Recebido3, Valor_ComissaoRec1, Valor_ComissaoRec2, Valor_ComissaoRec3 " & _
       "FROM funcionario " & _
       "WHERE (CODIGO = " & txtCodFunc.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim vAlvoRecebido As Currency
Dim vComissaoRecebido As Currency

If Not r.EOF Then
    If vValorTotalRecebido > r("Valor_ComissaoRec1") Then
        If vValorTotalRecebido < r("Valor_ComissaoRec3") Then
            vComissaoRecebido = FormatNumber(r("Comissao_Recebido2"), 2)
        Else
            vComissaoRecebido = FormatNumber(r("Comissao_Recebido3"), 2)
        End If
    Else
        vComissaoRecebido = FormatNumber(r("Comissao_Recebido1"), 2)
    End If
Else
    vComissaoRecebido = FormatNumber(0, 2)
End If

'COMISSÃO à RECEBIDO
sSQL = "SELECT  ISNULL(SUM(parcelas.VALOR_FINAL * " & Replace(CDbl(vComissaoRecebido), ",", ".") & " / 100), 0) AS var_ComRecebido, count(parcelas.COD_PEDIDO) as var_ContParcelas " & _
        "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO INNER JOIN funcionario ON pedidos.COD_FUNCIONARIO = funcionario.CODIGO " & _
        "WHERE (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Prazo') AND (parcelas.STATUS = 1) AND (Month(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") And (Year(parcelas.PAGAMENTO) = " & cboAno & ")"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

'Debug.Print sSQL
If Not r.EOF Then
    lblComRecebidoQtde.Caption = Format(r("var_ContParcelas"), "000")
    lblComRecebido.Caption = FormatNumber(r("var_ComRecebido"), 2)
Else
    lblComRecebidoQtde.Caption = Format(0, "000")
    lblComRecebido.Caption = FormatNumber(0, 2)
End If

If lblComRecebido.Caption = "0,00" Then
    lblComRecebidoQtde.Caption = Format(0, "000")
End If





'COMISSÃO à PRAZO - EXIBIR TOTAIS
sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL * funcionario.Comissao_Prazo1 / 100), 0) AS var_ComAprazo, COUNT(parcelas.CODIGO) AS var_ContParcelas " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
                     "INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
                     "INNER JOIN funcionario ON pedidos.COD_FUNCIONARIO = funcionario.CODIGO " & _
       "WHERE (pedidos.TIPO_PEDIDO = 'VENDA') AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Prazo') " & vTipoPgtoParcelas & " AND (MONTH(pedidos.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos.DATA_COMPRA) = " & cboAno & ") "
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If Not r.EOF Then
    lblComAPrazoQtde.Caption = Format(r("var_ContParcelas"), "000")
    lblComAPrazo.Caption = FormatNumber(r("var_ComAprazo"), 2)
Else
    lblComAPrazoQtde.Caption = Format(0, "00")
    lblComAPrazo.Caption = FormatNumber(0, 2)
End If

'COMISSÃO DE SERVIÇOS
lblComServicosQtde.Caption = Format(0, "000")
lblComServicos.Caption = FormatNumber(0, 2)

If r.State <> 0 Then r.Close
Set r = Nothing
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


Private Sub cmdCriarPDF_Click()
Dim r As ADODB.Recordset

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_ImpNormal = "Impressora PDF"

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
'Debug.Print printSQL
   Set REL_Comissoes.Relatorio.Recordset = r
    
    REL_Comissoes.lblTitulo.Caption = "RELATÓRIO DE COMISSÕES"
   
    REL_Comissoes.dfQuantAvista.Caption = lblComAvistaQtde.Caption
    REL_Comissoes.dfQuantAPrazo.Caption = lblComAPrazoQtde.Caption
    REL_Comissoes.dfQuantRecebido.Caption = lblComRecebidoQtde.Caption
    REL_Comissoes.dfQuantServicos.Caption = lblComServicosQtde.Caption

    REL_Comissoes.dfTotalAVista.Caption = FormatNumber(lblComAvista.Caption, 2)
    REL_Comissoes.dfTotalAPrazo.Caption = FormatNumber(lblComAPrazo.Caption, 2)
    REL_Comissoes.dfTotalRecebido.Caption = FormatNumber(lblComRecebido.Caption, 2)
    REL_Comissoes.dfTotalServicos.Caption = FormatNumber(lblComServicos.Caption, 2)

   If cboFormaPgto.Text = "TODOS" Then
      REL_Comissoes.rfForma.Caption = "TODAS"
   ElseIf cboFormaPgto.Text = "À VISTA" Then
      REL_Comissoes.rfForma.Caption = "À VISTA"
   ElseIf cboFormaPgto.Text = "À PRAZO" Then
      REL_Comissoes.rfForma.Caption = "À PRAZO"
   Else
      REL_Comissoes.rfForma.Caption = "TODAS"
   End If

   
   If cboCriterioSec.Text = "MENSAL" Then
      REL_Comissoes.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
   End If
   REL_Comissoes.rfCons1.Caption = cboVendedor
   REL_Comissoes.rfCons3.Caption = FormatNumber(lblSubTotal.Caption, 2)
    
    REL_Comissoes.Relatorio.NomeImpressora = var_ImpNormal
    REL_Comissoes.Relatorio.Visualizar = False
    REL_Comissoes.Relatorio.Ativar
    Unload REL_Comissoes

Me.Show 1
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
'Debug.Print printSQL
   Set REL_Comissoes.Relatorio.Recordset = r
    
    REL_Comissoes.lblTitulo.Caption = "RELATÓRIO DE COMISSÕES"
   
    REL_Comissoes.dfQuantAvista.Caption = lblComAvistaQtde.Caption
    REL_Comissoes.dfQuantAPrazo.Caption = lblComAPrazoQtde.Caption
    REL_Comissoes.dfQuantRecebido.Caption = lblComRecebidoQtde.Caption
    REL_Comissoes.dfQuantServicos.Caption = lblComServicosQtde.Caption

    REL_Comissoes.dfTotalAVista.Caption = FormatNumber(lblComAvista.Caption, 2)
    REL_Comissoes.dfTotalAPrazo.Caption = FormatNumber(lblComAPrazo.Caption, 2)
    REL_Comissoes.dfTotalRecebido.Caption = FormatNumber(lblComRecebido.Caption, 2)
    REL_Comissoes.dfTotalServicos.Caption = FormatNumber(lblComServicos.Caption, 2)

   If cboFormaPgto.Text = "TODOS" Then
      REL_Comissoes.rfForma.Caption = "TODAS"
   ElseIf cboFormaPgto.Text = "À VISTA" Then
      REL_Comissoes.rfForma.Caption = "À VISTA"
   ElseIf cboFormaPgto.Text = "À PRAZO" Then
      REL_Comissoes.rfForma.Caption = "À PRAZO"
   Else
      REL_Comissoes.rfForma.Caption = "TODAS"
   End If

   
   If cboCriterioSec.Text = "MENSAL" Then
      REL_Comissoes.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
   End If
   REL_Comissoes.rfCons1.Caption = cboVendedor
   REL_Comissoes.rfCons3.Caption = FormatNumber(lblSubTotal.Caption, 2)
   
   
   REL_Comissoes.Relatorio.Ativar
   Unload REL_Comissoes



Me.Show 1
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




Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
   
'limpar o grid
PreencherTipoConsulta
cboTipo.ListIndex = 0

PreencherCriterios
cboCriterioSec.ListIndex = 0

PreencherIndice
cboIndice.ListIndex = 1

PreencherFormaPgto
cboFormaPgto.ListIndex = 0

cboTipoPgto.Text = "TODOS"

'cboMes.Text = Format(Date, "mmmm")
'cboAno.Text = Year(Date)

StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
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
      'cmdLocalizar.SetFocus
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

Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 10
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 1000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 700
      .ColWidth(5) = 4300
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1500
     
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "TIPO"
      .TextMatrix(0, 3) = "COMPRA"
      .TextMatrix(0, 4) = "PARC."
      .TextMatrix(0, 5) = "NOME DO CLIENTE"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "STATUS"
      .TextMatrix(0, 8) = "PGTO"
      .TextMatrix(0, 9) = "FORMA"
      

      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 7
      .ColAlignment(2) = 7
      .ColAlignment(3) = 7
      .ColAlignment(4) = 7
      .ColAlignment(5) = 1
      .ColAlignment(6) = 7
      .ColAlignment(7) = 7
      .ColAlignment(8) = 7
      .ColAlignment(9) = 7
      
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("TIPO_PAGAMENTO"))
            .TextMatrix(.rows - 1, 3) = Format(rTabela("DATA_COMPRA"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("NUMERO"))
            .TextMatrix(.rows - 1, 5) = Format(UCase(rTabela("NOME")), ocMONEY)
            .TextMatrix(.rows - 1, 6) = Format(rTabela("VALOR_FINAL"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_StatusPgto"))
            .TextMatrix(.rows - 1, 8) = Format(rTabela("Pagamento"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_FormaPgto"))
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
   
   
    lblSubTotal.Caption = Format(SomaGrid(Grid, 6), ocMONEY)
    'lblSubtotalBruto.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
    'lblAcresc.Caption = Format(SomaGrid(Grid, 6), ocMONEY)
    'lblTotal.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
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
