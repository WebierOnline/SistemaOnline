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
      TabIndex        =   37
      Top             =   2520
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
         TabIndex        =   53
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
            TabIndex        =   29
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
            TabIndex        =   56
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
            TabIndex        =   55
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
            TabIndex        =   54
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
         TabIndex        =   48
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
         TabIndex        =   43
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
         TabIndex        =   38
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
      Height          =   2775
      Left            =   60
      TabIndex        =   26
      Top             =   900
      Width           =   2835
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   2595
      End
      Begin VB.ComboBox cboIndice 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   2595
      End
      Begin VB.ComboBox cboFormaPgto 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   2595
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Consulta"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Organizar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pgto"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   840
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
      Height          =   1575
      Left            =   2940
      TabIndex        =   11
      Top             =   900
      Width           =   10215
      Begin VB.Frame frmDatas 
         Caption         =   "Datas"
         Height          =   495
         Left            =   3240
         TabIndex        =   57
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
         Begin VB.OptionButton optExecucao 
            Caption         =   "Execução"
            Height          =   195
            Left            =   1080
            TabIndex        =   59
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton optTermino 
            Caption         =   "Termino"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Width           =   4905
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   1500
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   1140
         Width           =   1155
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   1335
      End
      Begin VB.TextBox txtCodFunc 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   4320
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendario2 
         Height          =   315
         Left            =   2700
         TabIndex        =   16
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
         TabIndex        =   18
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
      Begin ChamaleonBtn.chameleonButton chameleonButton1 
         Height          =   315
         Left            =   5100
         TabIndex        =   34
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "Funcionario_Comissao.frx":B156
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
         Caption         =   "Colaborador(a):"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblAte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         Height          =   195
         Left            =   1440
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   900
         Width           =   330
      End
      Begin VB.Label lblMes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   900
         Width           =   345
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4800
      Picture         =   "Funcionario_Comissao.frx":B172
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   2
      Top             =   5340
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
         Picture         =   "Funcionario_Comissao.frx":C1AA
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
            TextSave        =   "10:17"
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
      Height          =   4575
      Left            =   60
      TabIndex        =   6
      Top             =   3720
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8070
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirProdutos 
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   8340
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "PRODUTOS/SERVIÇOS"
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
      MICON           =   "Funcionario_Comissao.frx":CC86
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
      Left            =   2040
      TabIndex        =   8
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
      MICON           =   "Funcionario_Comissao.frx":CCA2
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
      Left            =   3900
      TabIndex        =   35
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
      MICON           =   "Funcionario_Comissao.frx":CCBE
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
      Left            =   5760
      TabIndex        =   36
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
      MICON           =   "Funcionario_Comissao.frx":CCDA
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      Left            =   10560
      TabIndex        =   3
      Top             =   8400
      Width           =   645
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
Private vRelValorAvista As Currency
Private vRelValorAPrazo As Currency
Private vRelValorRecebido As Currency
Private vRelValorServicos As Currency
Private vRelPorcAvista As Currency
Private vRelPorcAPrazo As Currency
Private vRelPorcRecebido As Currency
Private vRelPorcServicos As Currency

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
      .Cols = 12
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 1000
      .ColWidth(2) = 800
      .ColWidth(3) = 800
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 700
      .ColWidth(7) = 4300
      .ColWidth(8) = 1000
      .ColWidth(9) = 1000
      .ColWidth(10) = 1000
      .ColWidth(11) = 1000
     
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "ORIGEM"
      .TextMatrix(0, 4) = "TIPO"
      .TextMatrix(0, 5) = "DATA"
      .TextMatrix(0, 6) = "PARC."
      .TextMatrix(0, 7) = "NOME DO CLIENTE"
      .TextMatrix(0, 8) = "VALOR"
      .TextMatrix(0, 9) = "STATUS"
      .TextMatrix(0, 10) = "PGTO"
      .TextMatrix(0, 11) = "FORMA"
      

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
      .ColAlignment(5) = 7
      .ColAlignment(6) = 7
      .ColAlignment(7) = 1
      .ColAlignment(8) = 7
      .ColAlignment(9) = 7
      .ColAlignment(10) = 7
      .ColAlignment(11) = 7
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA (STATUS: preto = Pago/Feito, vermelho escuro = A Pagar)
      For i = 1 To .rows - 1
         .Row = i
         .Col = 9
         If .TextMatrix(i, 9) = "Pago" Or .TextMatrix(i, 9) = "Feito" Then
            .CellForeColor = vbBlack
         Else
            .CellForeColor = RGB(139, 0, 0)
         End If
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
cboTipo.AddItem "TODOS"
cboTipo.AddItem "VENDA"
cboTipo.AddItem "OFICINA"
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



Private Sub cboFormaPgto_LostFocus()
PreencherTipoPgto
End Sub

Private Sub cboFormaPgto_Change()
If cboFormaPgto.Text = "SERVIÇOS" Then
    frmDatas.Visible = True
Else
    frmDatas.Visible = False
End If
End Sub

Private Sub cboFormaPgto_Click()
cboFormaPgto_Change
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
If cboTipo.Text = "TODOS" Then
   cmdExibirProdutos.Visible = True
   cmdExibirParcelas.Visible = True
ElseIf cboTipo.Text = "VENDA" Then
   cmdExibirProdutos.Visible = True
   cmdExibirParcelas.Visible = True
ElseIf cboTipo.Text = "OFICINA" Then
   cmdExibirProdutos.Visible = True
   cmdExibirParcelas.Visible = True
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



Private Sub cboVendedor_Click()
   cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_GotFocus()
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



Dim vTipoCriterio As String
Dim vCriterioPagamento As String
Dim vCriterioCompra As String
Dim vCriterioTermino As String
Dim vCriterioExecucao As String
Dim vCriterioServicos As String

'Sempre MENSAL (cboCriterioSec removido -- todas as consultas sao mensais)
If cboMes.Text = "" Or cboAno.Text = "" Then Limpar_Grid: Exit Sub

vCriterioPagamento = " AND (Month(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") And (Year(parcelas.PAGAMENTO) = " & cboAno & ")"
vCriterioCompra = " AND (Month(pedidos.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") And (Year(pedidos.DATA_COMPRA) = " & cboAno & ")"
vCriterioTermino = " AND (Month(OS.DATA_TERMINO) = " & cboMes.ListIndex + 1 & ") And (Year(OS.DATA_TERMINO) = " & cboAno & ")"
vCriterioExecucao = " AND (Month(sv.data) = " & cboMes.ListIndex + 1 & ") And (Year(sv.data) = " & cboAno & ")"

If optExecucao.Value = True Then
    vCriterioServicos = vCriterioExecucao
Else
    vCriterioServicos = vCriterioTermino
End If

If cboFormaPgto.Text = "À VISTA" Then
    vTipoCriterio = vCriterioPagamento
ElseIf cboFormaPgto.Text = "À PRAZO" Then
    vTipoCriterio = vCriterioCompra
ElseIf cboFormaPgto.Text = "RECEBIDOS" Then
    vTipoCriterio = vCriterioPagamento
Else
    vTipoCriterio = vCriterioCompra
End If

'MONTAR O GRID
'Ordem de colunas ALINHADA entre os dois SELECTs (necessario para o UNION ALL do modo TODOS nao misturar campos de posicoes diferentes)
Dim sSQLOficina As String
Dim sSQLVenda As String

sSQLOficina = "SELECT OS.COD_PEDIDO AS var_codped, OS.COD_OS AS var_CodOS, 'OS' AS var_Origem, pedidos.TIPO_PAGAMENTO, pedidos.DATA_COMPRA, parcelas.NUMERO, parcelas.VALOR_FINAL, parcelas.FORMA_PGTO as var_FormaPgto, (CASE WHEN parcelas.status = 1 THEN 'Pago' ELSE 'À Pagar' END) AS var_StatusPgto, pedidos.COD_FUNCIONARIO, cliente.Nome, OS.COD_CLIENTE, parcelas.PAGAMENTO " & _
    "FROM OS INNER JOIN pedidos ON OS.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN parcelas ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO LEFT JOIN cliente ON OS.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (pedidos.TIPO_PEDIDO = 'OFICINA') AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") " & TipoPgto & " " & vPago & " " & " AND (OS.STATUS_OS = 1) " & vTipoCriterio & " "

sSQLVenda = "SELECT parcelas.COD_PEDIDO as var_codped, NULL AS var_CodOS, 'PDV' AS var_Origem, pedidos.TIPO_PAGAMENTO, pedidos.DATA_COMPRA, parcelas.NUMERO, parcelas.VALOR_FINAL, parcelas.FORMA_PGTO as var_FormaPgto, (CASE WHEN parcelas.status = 1 THEN 'Pago' ELSE 'À Pagar' END) AS var_StatusPgto, pedidos.COD_FUNCIONARIO, cliente.Nome, pedidos.COD_CLIENTE, parcelas.PAGAMENTO " & _
    "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (pedidos.TIPO_PEDIDO = 'VENDA') AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") " & TipoPgto & " " & vPago & " " & " " & vTipoCriterio & " "

'SERVIÇOS: sempre trata como OFICINA, atribuindo por sv.cod_mecanico (quem executou o servico), ignorando cboTipo
Dim sSQLServicos As String
sSQLServicos = "SELECT OS.COD_PEDIDO AS var_codped, OS.COD_OS AS var_CodOS, 'OS' AS var_Origem, pedidos.TIPO_PAGAMENTO, OS.DATA_TERMINO AS DATA_COMPRA, '' AS NUMERO, sv.TOTAL AS VALOR_FINAL, '' AS var_FormaPgto, 'Feito' AS var_StatusPgto, sv.cod_mecanico AS COD_FUNCIONARIO, cliente.Nome, OS.COD_CLIENTE, sv.data AS PAGAMENTO " & _
    "FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os = OS.COD_OS INNER JOIN pedidos ON OS.COD_PEDIDO = pedidos.COD_PEDIDO LEFT JOIN cliente ON OS.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (pedidos.TIPO_PEDIDO = 'OFICINA') AND (pedidos.cancelado = 0) AND (sv.cod_mecanico = " & txtCodFunc.Text & ") " & vCriterioServicos & " "

If cboFormaPgto.Text = "SERVIÇOS" Then
    sSQL = sSQLServicos
ElseIf cboTipo.Text = "OFICINA" Then
    sSQL = sSQLOficina
ElseIf cboTipo.Text = "VENDA" Then
    sSQL = sSQLVenda
Else
    'TODOS
    sSQL = sSQLOficina & " UNION ALL " & sSQLVenda
End If
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
        "WHERE (pedidos.TIPO_PEDIDO IN ('VENDA', 'OFICINA')) AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") " & TipoPgto & " " & vPago & " " & " " & vTipoCriterio & " "
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

Dim vValorTotalAvista As Currency
If Not r.EOF Then
    vValorTotalAvista = FormatNumber(ValidateNull(r("varTotalAvista")), 2)
Else
    vValorTotalAvista = FormatNumber(0, 2)
End If
vRelValorAvista = vValorTotalAvista

'CONSULTAS COMISSÕES
sSQL = "SELECT Comissao_Avista1, Comissao_Avista2, Comissao_Avista3, Valor_ComissaoAV1, Valor_ComissaoAV2, Valor_ComissaoAV3 " & _
       "FROM funcionario " & _
       "WHERE (CODIGO = " & txtCodFunc.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim vAlvoAvista As Currency
Dim vComissaoAvista As Currency
Dim vMeta1Avista As Currency
Dim vPerc1Avista As Currency

If Not r.EOF Then
    vMeta1Avista = ValidateNull(r("Valor_ComissaoAV1"))
    vPerc1Avista = ValidateNull(r("Comissao_Avista1"))
    If vValorTotalAvista > r("Valor_ComissaoAV1") Then
        If vValorTotalAvista < r("Valor_ComissaoAV3") Then
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
vRelPorcAvista = vComissaoAvista

'Meta 1 zerada (valor e/ou %) = funcionario nao participa dessa comissao: zera tudo, nao busca parcelas
If vMeta1Avista = 0 Or vPerc1Avista = 0 Then
    lblComAvistaQtde.Caption = Format(0, "000")
    lblComAvista.Caption = FormatNumber(0, 2)
    vRelValorAvista = 0
    vRelPorcAvista = 0
Else
    'COMISSÕES À VISTA
    sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL * " & Replace(CDbl(vComissaoAvista), ",", ".") & " / 100), 0) AS var_ComAvista, COUNT(parcelas.CODIGO) AS var_ContParcelas " & _
           "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
                         "INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
                         "INNER JOIN funcionario ON pedidos.COD_FUNCIONARIO = funcionario.CODIGO " & _
           "WHERE (pedidos.TIPO_PEDIDO IN ('VENDA', 'OFICINA')) AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Vista') " & vPago & " " & " " & vCriterioPagamento & " "
    Set r = dbData.OpenRecordset(sSQL, totalRegistros)

    If Not r.EOF Then
        lblComAvistaQtde.Caption = Format(r("var_ContParcelas"), "000")
        lblComAvista.Caption = FormatNumber(r("var_ComAvista"), 2)
    Else
        lblComAvistaQtde.Caption = Format(0, "00")
        lblComAvista.Caption = FormatNumber(0, 2)
    End If
End If





'BUSCAR TOTAL DE RECEBIDO ===================================================================================
sSQL = "SELECT SUM(parcelas.VALOR_FINAL) AS varTotalRecebido " & _
        "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
        "WHERE (pedidos.TIPO_PEDIDO IN ('VENDA', 'OFICINA')) AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Prazo') AND (parcelas.STATUS = 1) " & " " & vCriterioPagamento & ""
Set r = dbData.OpenRecordset(sSQL, totalRegistros)
'Debug.Print sSQL
Dim vValorTotalRecebido As Currency
If Not r.EOF Then
    vValorTotalRecebido = FormatNumber(ValidateNull(r("varTotalRecebido")), 2)
Else
    vValorTotalRecebido = FormatNumber(0, 2)
End If
vRelValorRecebido = vValorTotalRecebido

'CONSULTAS COMISSÕES
sSQL = "SELECT Comissao_Recebido1, Comissao_Recebido2, Comissao_Recebido3, Valor_ComissaoRec1, Valor_ComissaoRec2, Valor_ComissaoRec3 " & _
       "FROM funcionario " & _
       "WHERE (CODIGO = " & txtCodFunc.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim vAlvoRecebido As Currency
Dim vComissaoRecebido As Currency
Dim vMeta1Recebido As Currency
Dim vPerc1Recebido As Currency

If Not r.EOF Then
    vMeta1Recebido = ValidateNull(r("Valor_ComissaoRec1"))
    vPerc1Recebido = ValidateNull(r("Comissao_Recebido1"))
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
vRelPorcRecebido = vComissaoRecebido

'Meta 1 zerada (valor e/ou %) = funcionario nao participa dessa comissao: zera tudo, nao busca parcelas
If vMeta1Recebido = 0 Or vPerc1Recebido = 0 Then
    lblComRecebidoQtde.Caption = Format(0, "000")
    lblComRecebido.Caption = FormatNumber(0, 2)
    vRelValorRecebido = 0
    vRelPorcRecebido = 0
Else
    'COMISSÃO à RECEBIDO
    sSQL = "SELECT  ISNULL(SUM(parcelas.VALOR_FINAL * " & Replace(CDbl(vComissaoRecebido), ",", ".") & " / 100), 0) AS var_ComRecebido, count(parcelas.COD_PEDIDO) as var_ContParcelas " & _
            "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO INNER JOIN funcionario ON pedidos.COD_FUNCIONARIO = funcionario.CODIGO " & _
            "WHERE (pedidos.TIPO_PEDIDO IN ('VENDA', 'OFICINA')) AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Prazo') AND (parcelas.STATUS = 1) " & " " & vCriterioPagamento & ""
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
End If





'COMISSÃO à PRAZO - EXIBIR TOTAIS
Dim vValorTotalAPrazo As Currency
sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varTotalAPrazo " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
                     "INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
                     "INNER JOIN funcionario ON pedidos.COD_FUNCIONARIO = funcionario.CODIGO " & _
       "WHERE (pedidos.TIPO_PEDIDO IN ('VENDA', 'OFICINA')) AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Prazo') " & " " & vCriterioCompra & " "
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If Not r.EOF Then
    vValorTotalAPrazo = FormatNumber(ValidateNull(r("varTotalAPrazo")), 2)
Else
    vValorTotalAPrazo = FormatNumber(0, 2)
End If
vRelValorAPrazo = vValorTotalAPrazo

sSQL = "SELECT Comissao_Prazo1, Comissao_Prazo2, Comissao_Prazo3, Valor_ComissaoAP1, Valor_ComissaoAP2, Valor_ComissaoAP3 " & _
       "FROM funcionario " & _
       "WHERE (CODIGO = " & txtCodFunc.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim vComissaoAPrazo As Currency
Dim vMeta1APrazo As Currency
Dim vPerc1APrazo As Currency

If Not r.EOF Then
    vMeta1APrazo = ValidateNull(r("Valor_ComissaoAP1"))
    vPerc1APrazo = ValidateNull(r("Comissao_Prazo1"))
    If vValorTotalAPrazo > r("Valor_ComissaoAP1") Then
        If vValorTotalAPrazo < r("Valor_ComissaoAP3") Then
            vComissaoAPrazo = FormatNumber(r("Comissao_Prazo2"), 2)
        Else
            vComissaoAPrazo = FormatNumber(r("Comissao_Prazo3"), 2)
        End If
    Else
        vComissaoAPrazo = FormatNumber(r("Comissao_Prazo1"), 2)
    End If
Else
    vComissaoAPrazo = FormatNumber(0, 2)
End If
vRelPorcAPrazo = vComissaoAPrazo

'Meta 1 zerada (valor e/ou %) = funcionario nao participa dessa comissao: zera tudo, nao busca parcelas
If vMeta1APrazo = 0 Or vPerc1APrazo = 0 Then
    lblComAPrazoQtde.Caption = Format(0, "000")
    lblComAPrazo.Caption = FormatNumber(0, 2)
    vRelValorAPrazo = 0
    vRelPorcAPrazo = 0
Else
    sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL * " & Replace(CDbl(vComissaoAPrazo), ",", ".") & " / 100), 0) AS var_ComAprazo, COUNT(parcelas.CODIGO) AS var_ContParcelas " & _
           "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
                         "INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
                         "INNER JOIN funcionario ON pedidos.COD_FUNCIONARIO = funcionario.CODIGO " & _
           "WHERE (pedidos.TIPO_PEDIDO IN ('VENDA', 'OFICINA')) AND (pedidos.cancelado = 0) AND (pedidos.COD_FUNCIONARIO = " & txtCodFunc.Text & ") AND (pedidos.TIPO_PAGAMENTO = 'À Prazo') " & " " & vCriterioCompra & " "
    Set r = dbData.OpenRecordset(sSQL, totalRegistros)

    If Not r.EOF Then
        lblComAPrazoQtde.Caption = Format(r("var_ContParcelas"), "000")
        lblComAPrazo.Caption = FormatNumber(r("var_ComAprazo"), 2)
    Else
        lblComAPrazoQtde.Caption = Format(0, "00")
        lblComAPrazo.Caption = FormatNumber(0, 2)
    End If
End If

'COMISSÃO DE SERVIÇOS
Dim vValorTotalServicos As Currency
sSQL = "SELECT ISNULL(SUM(sv.total), 0) AS varTotalServicos " & _
       "FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os = OS.COD_OS INNER JOIN pedidos ON OS.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.TIPO_PEDIDO = 'OFICINA') AND (pedidos.cancelado = 0) AND (sv.cod_mecanico = " & txtCodFunc.Text & ") " & vCriterioServicos & " "
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If Not r.EOF Then
    vValorTotalServicos = FormatNumber(ValidateNull(r("varTotalServicos")), 2)
Else
    vValorTotalServicos = FormatNumber(0, 2)
End If
vRelValorServicos = vValorTotalServicos

sSQL = "SELECT Comissao_Servico1, Comissao_Servico2, Comissao_Servico3, Valor_ComissaoServ1, Valor_ComissaoServ2, Valor_ComissaoServ3 " & _
       "FROM funcionario " & _
       "WHERE (CODIGO = " & txtCodFunc.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim vComissaoServicos As Currency
Dim vMeta1Servicos As Currency
Dim vPerc1Servicos As Currency

If Not r.EOF Then
    vMeta1Servicos = ValidateNull(r("Valor_ComissaoServ1"))
    vPerc1Servicos = ValidateNull(r("Comissao_Servico1"))
    If vValorTotalServicos > r("Valor_ComissaoServ1") Then
        If vValorTotalServicos < r("Valor_ComissaoServ3") Then
            vComissaoServicos = FormatNumber(r("Comissao_Servico2"), 2)
        Else
            vComissaoServicos = FormatNumber(r("Comissao_Servico3"), 2)
        End If
    Else
        vComissaoServicos = FormatNumber(r("Comissao_Servico1"), 2)
    End If
Else
    vComissaoServicos = FormatNumber(0, 2)
End If
vRelPorcServicos = vComissaoServicos

'Meta 1 zerada (valor e/ou %) = funcionario nao participa dessa comissao: zera tudo, nao busca servicos
If vMeta1Servicos = 0 Or vPerc1Servicos = 0 Then
    lblComServicosQtde.Caption = Format(0, "000")
    lblComServicos.Caption = FormatNumber(0, 2)
    vRelValorServicos = 0
    vRelPorcServicos = 0
Else
    sSQL = "SELECT ISNULL(SUM(sv.total * " & Replace(CDbl(vComissaoServicos), ",", ".") & " / 100), 0) AS var_ComServicos, COUNT(sv.codigo) AS var_ContServicos " & _
           "FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os = OS.COD_OS INNER JOIN pedidos ON OS.COD_PEDIDO = pedidos.COD_PEDIDO " & _
           "WHERE (pedidos.TIPO_PEDIDO = 'OFICINA') AND (pedidos.cancelado = 0) AND (sv.cod_mecanico = " & txtCodFunc.Text & ") " & vCriterioServicos & " "
    Set r = dbData.OpenRecordset(sSQL, totalRegistros)

    If Not r.EOF Then
        lblComServicosQtde.Caption = Format(r("var_ContServicos"), "000")
        lblComServicos.Caption = FormatNumber(r("var_ComServicos"), 2)
    Else
        lblComServicosQtde.Caption = Format(0, "000")
        lblComServicos.Caption = FormatNumber(0, 2)
    End If
End If

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


Private Sub MontarRelatorioComissoes(ByVal bGerarPDF As Boolean)
Dim r As ADODB.Recordset

If printSQL = "" Then
    ShowMsg "Clique em ""Exibir"" antes de gerar o relatório.", vbInformation
    Exit Sub
End If

If bGerarPDF Then
    Set oIni = New Ini
    oIni.Arquivo = appPathApp & "config.ini"
    var_ImpNormal = "Impressora PDF"

    Dim Prt As Printer
    Dim bImpressoraEncontrada As Boolean
    bImpressoraEncontrada = False
    For Each Prt In Printers
       If Prt.DeviceName = var_ImpNormal Then
          Set Printer = Prt
          bImpressoraEncontrada = True
          Exit For
       End If
    Next

    If Not bImpressoraEncontrada Then
        ShowMsg "Não foi encontrada nenhuma impressora chamada ""Impressora PDF"" instalada neste computador." & vbCrLf & vbCrLf & "Entre em contato com o suporte para instalar a Impressora PDF.", vbExclamation
        Exit Sub
    End If
Else
    Dim oIniNormal As Ini
    Dim var_Impressora As String
    Set oIniNormal = New Ini
    oIniNormal.Arquivo = appPathApp & "config.ini"
    var_Impressora = oIniNormal.LerTexto("DADOS_IMPRESSORA", "impressora")
    Set oIniNormal = Nothing
End If

Me.Hide

Set r = dbData.OpenRecordset(printSQL)
Set REL_Comissoes.Relatorio.Recordset = r

REL_Comissoes.lblTitulo.Caption = "RELATÓRIO DE COMISSÕES"

'Cabecalhos dinamicos: modo SERVICOS troca COMPRA/VALOR/PGTO por TERMINO/SERVICO/DATA e oculta FORMA (mesmo padrao do Grid/FormatarGrid)
Dim bModoServicosRel As Boolean
bModoServicosRel = (cboFormaPgto.Text = "SERVIÇOS")
REL_Comissoes.Label5.Caption = IIf(bModoServicosRel, "TÉRMINO", "COMPRA")
REL_Comissoes.Label1.Caption = IIf(bModoServicosRel, "SERVIÇO", "VALOR")
REL_Comissoes.Label9.Caption = IIf(bModoServicosRel, "DATA", "PGTO")
REL_Comissoes.Label10.Visible = Not bModoServicosRel
REL_Comissoes.ReportField7.Visible = Not bModoServicosRel

REL_Comissoes.dfQuantAvista.Caption = lblComAvistaQtde.Caption
REL_Comissoes.dfQuantAPrazo.Caption = lblComAPrazoQtde.Caption
REL_Comissoes.dfQuantRecebido.Caption = lblComRecebidoQtde.Caption
REL_Comissoes.dfQuantServicos.Caption = lblComServicosQtde.Caption

REL_Comissoes.dfTotalAVista.Caption = FormatNumber(lblComAvista.Caption, 2)
REL_Comissoes.dfTotalAPrazo.Caption = FormatNumber(lblComAPrazo.Caption, 2)
REL_Comissoes.dfTotalRecebido.Caption = FormatNumber(lblComRecebido.Caption, 2)
REL_Comissoes.dfTotalServicos.Caption = FormatNumber(lblComServicos.Caption, 2)

'Soma usada p/ definir a faixa de comissao + % da faixa escolhida (persistidos em modulo, ver vRelValor*/vRelPorc*)
REL_Comissoes.dfValorAvista.Caption = FormatNumber(vRelValorAvista, 2)
REL_Comissoes.dfValorAPrazo.Caption = FormatNumber(vRelValorAPrazo, 2)
REL_Comissoes.dfValorRecebido.Caption = FormatNumber(vRelValorRecebido, 2)
REL_Comissoes.dfValorServicos.Caption = FormatNumber(vRelValorServicos, 2)
REL_Comissoes.dfPorcAvista.Caption = FormatNumber(vRelPorcAvista, 2)
REL_Comissoes.dfPorcAPrazo.Caption = FormatNumber(vRelPorcAPrazo, 2)
REL_Comissoes.dfPorcRecebido.Caption = FormatNumber(vRelPorcRecebido, 2)
REL_Comissoes.dfPorcServicos.Caption = FormatNumber(vRelPorcServicos, 2)

'FORMA + TIPO (modo SERVIÇOS sempre trata como OFICINA, ignorando cboTipo)
If cboFormaPgto.Text = "SERVIÇOS" Then
   REL_Comissoes.rfForma.Caption = "SERVIÇOS (somente OFICINA)"
Else
   REL_Comissoes.rfForma.Caption = cboFormaPgto.Text & " (" & cboTipo.Text & ")"
End If

REL_Comissoes.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
If cboFormaPgto.Text = "SERVIÇOS" Then
    REL_Comissoes.rfCons2.Caption = REL_Comissoes.rfCons2.Caption & " (" & IIf(optExecucao.Value = True, "Execução do Serviço", "Término da OS") & ")"
End If
REL_Comissoes.rfCons1.Caption = cboVendedor
REL_Comissoes.rfCons3.Caption = FormatNumber(lblSubtotal.Caption, 2)

If bGerarPDF Then
    REL_Comissoes.Relatorio.NomeImpressora = var_ImpNormal
    REL_Comissoes.Relatorio.Visualizar = False
End If
REL_Comissoes.Relatorio.Ativar
Unload REL_Comissoes

Me.Show 1
End Sub

Private Sub cmdCriarPDF_Click()
MontarRelatorioComissoes True
End Sub

Private Sub cmdExibirParcelas_Click()
If Grid.Col = 0 Then Exit Sub
   If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
         Vendas_Consulta_Geral_Parcelas.loadInformacoes (Grid.TextMatrix(Grid.Row, 1))
         Vendas_Consulta_Geral_Parcelas.Show 1
   End If
End Sub

Private Sub cmdExibirProdutos_Click()
If Grid.Col = 0 Then Exit Sub
If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
   If Grid.Col = 1 Then
      If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), "OS"
      Parcelas_Consulta_Produtos.Show 1
   End If
End If
End Sub

Private Sub cmdImprimir_Click()
MontarRelatorioComissoes False
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
PreencherTipoConsulta
cboTipo.ListIndex = 0

PreencherIndice
cboIndice.ListIndex = 1

PreencherFormaPgto
cboFormaPgto.ListIndex = 0


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
      .Cols = 12
      .rows = 2
      
      Dim bModoServicos As Boolean
      bModoServicos = (cboFormaPgto.Text = "SERVIÇOS")
      
      .ColWidth(0) = 150
      .ColWidth(1) = 850
      .ColWidth(2) = 850
      .ColWidth(3) = 900
      .ColWidth(4) = 700
      .ColWidth(5) = 1000
      .ColWidth(6) = 300
      .ColWidth(7) = 4300
      .ColWidth(8) = 900
      .ColWidth(9) = 850
      .ColWidth(10) = 900
      .ColWidth(11) = IIf(bModoServicos, 0, 1100)
     
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "ORIGEM"
      .TextMatrix(0, 4) = "TIPO"
      .TextMatrix(0, 5) = IIf(bModoServicos, "TERMINO", "COMPRA")
      .TextMatrix(0, 6) = "Nº"
      .TextMatrix(0, 7) = "NOME DO CLIENTE"
      .TextMatrix(0, 8) = IIf(bModoServicos, "SERVIÇO", "VALOR")
      .TextMatrix(0, 9) = "STATUS"
      .TextMatrix(0, 10) = IIf(bModoServicos, "DATA", "PGTO")
      .TextMatrix(0, 11) = "FORMA"
      

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
      .ColAlignment(5) = 7
      .ColAlignment(6) = 7
      .ColAlignment(7) = 1
      .ColAlignment(8) = 7
      .ColAlignment(9) = 7
      .ColAlignment(10) = 7
      .ColAlignment(11) = 1
      
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            .TextMatrix(.rows - 1, 2) = IIf(IsNull(rTabela("var_CodOS")), "", Format(rTabela("var_CodOS"), "000000"))
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("var_Origem"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("TIPO_PAGAMENTO"))
            .TextMatrix(.rows - 1, 5) = Format(rTabela("DATA_COMPRA"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("NUMERO"))
            .TextMatrix(.rows - 1, 7) = Format(UCase(rTabela("NOME")), ocMONEY)
            .TextMatrix(.rows - 1, 8) = Format(rTabela("VALOR_FINAL"), ocMONEY)
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_StatusPgto"))
            If .TextMatrix(.rows - 1, 9) = "Pago" Or .TextMatrix(.rows - 1, 9) = "Feito" Then
               .TextMatrix(.rows - 1, 10) = Format(rTabela("Pagamento"), "dd/mm/yy")
               .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela("var_FormaPgto"))
            Else
               .TextMatrix(.rows - 1, 10) = ""
               .TextMatrix(.rows - 1, 11) = ""
            End If
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
      
      'MUDAR COR DE FONTE DA COLUNA (STATUS: preto = Pago/Feito, vermelho escuro = A Pagar)
      For i = 1 To .rows - 1
         .Row = i
         .Col = 9
         If .TextMatrix(i, 9) = "Pago" Or .TextMatrix(i, 9) = "Feito" Then
            .CellForeColor = vbBlack
         Else
            .CellForeColor = RGB(139, 0, 0)
         End If
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
   
    lblSubtotal.Caption = Format(SomaGrid(Grid, 8), ocMONEY)
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
