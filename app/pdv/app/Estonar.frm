VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Estonar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CONSULTA DE VENDAS"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   14805
   Icon            =   "Estonar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   60
      ScaleHeight     =   1065
      ScaleWidth      =   14625
      TabIndex        =   60
      Top             =   6960
      Width           =   14655
      Begin VB.Label lblTotaisFinanceiro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1440
         TabIndex        =   76
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   945
         TabIndex        =   75
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendas:"
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
         Left            =   345
         TabIndex        =   74
         Top             =   0
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblTotalGrid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   13380
         TabIndex        =   73
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label lblTotalVendas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   12885
         TabIndex        =   72
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendas:"
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
         Left            =   12285
         TabIndex        =   71
         Top             =   60
         Width           =   540
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Canceladas:"
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
         Left            =   12000
         TabIndex        =   70
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblQuantCanc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   12885
         TabIndex        =   69
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lblTotalCanc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   13380
         TabIndex        =   68
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Orçamento:"
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
         Left            =   12030
         TabIndex        =   67
         Top             =   540
         Width           =   795
      End
      Begin VB.Label lblQuantOrc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   12885
         TabIndex        =   66
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblTotalGridORC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   13380
         TabIndex        =   65
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label lblTotalGridConsignado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   13380
         TabIndex        =   64
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label lblQuantConsignado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   12885
         TabIndex        =   63
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Consignado:"
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
         Left            =   11940
         TabIndex        =   62
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.PictureBox frmSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   12660
      ScaleHeight     =   1245
      ScaleWidth      =   2025
      TabIndex        =   43
      Top             =   660
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   47
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtCodUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtNivelUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonBtn.chameleonButton cmdSenha 
         Height          =   315
         Left            =   1500
         TabIndex        =   48
         Top             =   840
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
         MICON           =   "Estonar.frx":23D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
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
         TabIndex        =   49
         Top             =   60
         Width           =   645
      End
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
      Height          =   1575
      Left            =   3480
      TabIndex        =   13
      Top             =   600
      Width           =   7275
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4020
         TabIndex        =   41
         Top             =   1140
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboTipoPgto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   39
         Top             =   1140
         Width           =   2175
      End
      Begin VB.ComboBox cboFormaPgto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   37
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtCodProdutoBarra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtCodBarra 
         Height          =   315
         Left            =   1260
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtCodProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cboProduto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   540
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboAno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2340
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox cboCliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   780
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   5715
      End
      Begin VB.TextBox txtCodCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox txtCodPedido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.TextBox txtCodPedidoCerto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5820
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optDig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Digitado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   15
         Top             =   180
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optEsc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolhendo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   14
         Top             =   180
         Visible         =   0   'False
         Width           =   1155
      End
      Begin ChamaleonBtn.chameleonButton cmdCal1 
         Height          =   315
         Left            =   2220
         TabIndex        =   27
         Tag             =   "Calendario"
         Top             =   360
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
         MICON           =   "Estonar.frx":23EE
         PICN            =   "Estonar.frx":240A
         PICH            =   "Estonar.frx":475D
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
         Height          =   285
         Left            =   600
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   615
         Left            =   5460
         TabIndex        =   59
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
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
         MICON           =   "Estonar.frx":6AB0
         PICN            =   "Estonar.frx":6ACC
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
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pagamento:"
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
         Left            =   180
         TabIndex        =   40
         Top             =   1140
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pagamento:"
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
         Left            =   180
         TabIndex        =   38
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label lblCodBarra 
         AutoSize        =   -1  'True
         Caption         =   "Cód. de Barra:"
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
         Left            =   180
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   180
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   180
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblMes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês:"
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
         Left            =   180
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblAno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano:"
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
         Left            =   1980
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   180
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblCodPedido 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cód. Pedido:"
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
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.Frame frmClassificacao 
      Caption         =   "Classificação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   60
      TabIndex        =   9
      Top             =   600
      Width           =   3375
      Begin VB.CheckBox chkIncompleto 
         Alignment       =   1  'Right Justify
         Caption         =   "Incompleto"
         Height          =   195
         Left            =   2160
         TabIndex        =   77
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cboTipoPedido 
         Height          =   315
         Left            =   1140
         TabIndex        =   42
         Top             =   1020
         Width           =   2175
      End
      Begin VB.ComboBox cboIndice 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboCriterios 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Organização:"
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
         Left            =   180
         TabIndex        =   12
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Pedido:"
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
         Left            =   180
         TabIndex        =   11
         Top             =   1020
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Critérios:"
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
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   14625
      TabIndex        =   6
      Top             =   60
      Width           =   14655
      Begin VB.Image Image1 
         Height          =   405
         Left            =   240
         Picture         =   "Estonar.frx":73A6
         Top             =   40
         Width           =   405
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE VENDAS"
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
         Left            =   900
         TabIndex        =   7
         Top             =   60
         Width           =   3600
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   8085
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17824
            MinWidth        =   1764
            Text            =   "ONLINE.INFO - INFORMÁTICA"
            TextSave        =   "ONLINE.INFO - INFORMÁTICA"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:48"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChamaleonBtn.chameleonButton cmdExcluirPedido 
      Height          =   315
      Left            =   3420
      TabIndex        =   5
      Tag             =   "2"
      Top             =   6600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "CANCELAR"
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
      MICON           =   "Estonar.frx":7B54
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdPedidoAbrir 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Tag             =   "1"
      Top             =   6600
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "REABRIR"
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
      MICON           =   "Estonar.frx":7B70
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdPedidoImprimir 
      Height          =   315
      Left            =   4860
      TabIndex        =   4
      Top             =   6600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "REIMPRIMIR"
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
      MICON           =   "Estonar.frx":7B8C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4335
      Left            =   60
      TabIndex        =   2
      Top             =   2220
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   7646
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   315
      Left            =   9360
      TabIndex        =   22
      Top             =   6600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "IMPRIMIR LISTA"
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
      MICON           =   "Estonar.frx":7BA8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdMostrarProdutos 
      Height          =   315
      Left            =   7380
      TabIndex        =   30
      Top             =   6600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "EXIBIR PRODUTOS"
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
      MICON           =   "Estonar.frx":7BC4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdReaberturas 
      Height          =   315
      Left            =   10800
      TabIndex        =   55
      Top             =   6600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "REABERTURAS"
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
      MICON           =   "Estonar.frx":7BE0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdModificar 
      Height          =   315
      Left            =   1980
      TabIndex        =   56
      Top             =   6600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "EDITAR"
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
      MICON           =   "Estonar.frx":7BFC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdPDF 
      Height          =   315
      Left            =   6300
      TabIndex        =   57
      Top             =   6600
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "CRIAR PDF"
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
      MICON           =   "Estonar.frx":7C18
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdModificarConsignado 
      Height          =   315
      Left            =   1980
      TabIndex        =   58
      Top             =   6600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "EDITAR"
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
      MICON           =   "Estonar.frx":7C34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdReimprimir 
      Height          =   255
      Left            =   10860
      TabIndex        =   61
      Top             =   660
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "&Abrir"
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
      MICON           =   "Estonar.frx":7C50
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imLogoCupom 
      Height          =   1125
      Left            =   11100
      Picture         =   "Estonar.frx":7C6C
      Top             =   720
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label lblCodUser2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   11700
      TabIndex        =   54
      Top             =   1800
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblUser2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   11700
      TabIndex        =   53
      Top             =   1980
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblUser1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   10920
      TabIndex        =   52
      Top             =   1980
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblCodUser1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód:"
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
      Left            =   11220
      TabIndex        =   51
      Top             =   1800
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "Estonar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim printSQL As String
Dim CAIXA_FECHADO As Boolean
Dim codPedido As String
Dim sSQL As String
Dim r As ADODB.Recordset
Public cCfg As ConfigItem
Public oIni As Ini
Public tipoEmpresa As String
Public vDeclararRecebedor As String
Dim vCodUsuario As Long
Public Ctrl As Control
Public varSegurancaAvancada As String
Dim vTipoParcelaImpressao As Integer
Dim vDescItensVenda As Currency
Dim rF As ADODB.Recordset
Dim varValorDescProc As Double
Dim A As Currency
Dim B As Currency

Private Sub Cancelar_NFCe(vCodPedido As String)
Dim sSQL As String, IdNFProd As Long

sSQL = "SELECT IdNFProd FROM TbNFCe WHERE Num_OS_VD_Origem  = " & vCodPedido
IdNFProd = SQLExecutaRetorno(sSQL, "IdNFProd", 0)
If IdNFProd > 0 Then
   sSQL = "SELECT NFCeChaveAcesso, NFCeProtocolo, NFCeCancelada, NFCeCanceladaProtocolo, NFCeCanceladaJustificativa FROM TbNFCe WHERE IdNFProd = " & IdNFProd
   NFeChaveAcesso = SQLExecutaRetorno(sSQL, "NFCeChaveAcesso", "")
   NFeNumeroProtocolo = SQLExecutaRetorno(sSQL, "NFCeProtocolo", 0)
   If Not Vazio(NFeChaveAcesso) And NFeNumeroProtocolo > 0 Then
      If CancelaNFCe(NFeChaveAcesso, NFeNumeroProtocolo, "DESISTENCIA DE COMPRA", True) Then
         sSQL = "UPDATE TbNFCe SET NFCeCancelada = 1, NFceCanceladaProtocolo = " & NFeNumeroProtocolo & ", NFCeCanceladaJustificativa = 'DESISTENCIA DE COMPRA' WHERE IDNFProd = " & IdNFProd
         dbData.Execute sSQL
      End If
   End If
End If
End Sub



Private Sub LiberarBotoesPermissoes()
If lblCodUser2.Caption = "" Then
    vCodUsuario = 0
Else
    vCodUsuario = lblCodUser2.Caption
End If

If vCodUsuario = False Then
    cmdPedidoAbrir.Enabled = False
    cmdExcluirPedido.Enabled = False
    cmdModificar.Enabled = False
    cmdModificarConsignado.Enabled = False
Else
    For Each Ctrl In Me.Controls
       If (TypeOf Ctrl Is chameleonButton) Then
           If Ctrl.Tag = "1" Then
               If LerPermissoesUsuario(vCodUsuario, 1) = True Then
                    Ctrl.Enabled = True
                Else
                    Ctrl.Enabled = False
               End If
           End If
           If Ctrl.Tag = "2" Then
               If LerPermissoesUsuario(vCodUsuario, 2) = True Then
                    Ctrl.Enabled = True
               Else
                    Ctrl.Enabled = False
               End If
           End If
       End If
    Next
End If
End Sub

Private Sub LimparGridPedidos()
'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, cliente.nome as var_Cliente, pedidos.DATA_COMPRA as var_Data,pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE 'NÃO' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE 'NÃO' END) AS Var_StatusCANCELADO, ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN 'SIM' ELSE 'NÃO' END) AS Var_NFCEInutilizada " & _
"FROM pedidos INNER JOIN pedidos_itens ON pedidos.COD_PEDIDO = pedidos_itens.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO LEFT OUTER JOIN TbNFCe ON TbNFCe.Num_OS_VD_Origem = pedidos.COD_PEDIDO  where 1 = 0" & _
"ORDER BY pedidos.cod_pedido"

Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Pedido r

lblTotalGrid.Caption = Format(0, "#,##0.00")
lblTotalVendas.Caption = Format(0, "000")
lblTotalGridORC.Caption = Format(0, "#,##0.00")
lblQuantOrc.Caption = Format(0, "000")
lblTotalCanc.Caption = Format(0, "#,##0.00")
lblQuantCanc.Caption = Format(0, "000")

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Public Sub PegarCodUsuario(vCod As Long)
vCodUsuario = vCod
End Sub

Private Sub PreencherCriterios()
Dim varTexto As String
varTexto = cboCriterios.Text
cboCriterios.Clear
cboCriterios.AddItem "NENHUM"
cboCriterios.AddItem "CÓD. PEDIDO"
cboCriterios.AddItem "CLIENTE"
cboCriterios.AddItem "DATA"
cboCriterios.AddItem "MENSAL"
'cboCriterios.AddItem "PRODUTO"
'cboCriterios.AddItem "CÓD. BARRA"
cboCriterios.Text = varTexto
moCombo.AttachTo cboCriterios
End Sub

Private Sub PreencherFormaPgto()
Dim varTexto As String
varTexto = cboFormaPgto.Text

cboFormaPgto.Clear
cboFormaPgto.AddItem "TODOS"
cboFormaPgto.AddItem "À VISTA"
cboFormaPgto.AddItem "À PRAZO"

cboFormaPgto.Text = varTexto
moCombo.AttachTo cboFormaPgto
End Sub

Private Sub PreencherTipoPedido()
Dim varTexto As String
varTexto = cboTipoPedido.Text
cboTipoPedido.Clear
'cboTipoPedido.AddItem "TODOS"
cboTipoPedido.AddItem "VENDA"
cboTipoPedido.AddItem "ORÇAMENTO"
cboTipoPedido.AddItem "CONSIGNADO"
cboTipoPedido.AddItem "ALUGUEL"
cboTipoPedido.AddItem "OFICINA"
cboTipoPedido.AddItem "CANCELADO"
cboTipoPedido.Text = varTexto
moCombo.AttachTo cboTipoPedido
End Sub

Private Sub PreencherIndice()
Dim varTexto As String
varTexto = cboIndice.Text
cboIndice.Clear
cboIndice.AddItem "CÓD. PEDIDO"
cboIndice.AddItem "CLIENTE"
cboIndice.AddItem "EMISSÃO"
cboIndice.AddItem "TIPO PGTO"
cboIndice.Text = varTexto
moCombo.AttachTo cboIndice
End Sub


Private Sub PreencherStatus()
Dim varTexto As String
varTexto = cboStatus.Text

cboStatus.Clear
cboStatus.AddItem "TODOS"
cboStatus.AddItem "FECHADO"
cboStatus.AddItem "ABERTO"
cboStatus.AddItem "PAUSADO"
cboStatus.AddItem "VAZIO"

cboStatus.Text = varTexto
moCombo.AttachTo cboStatus
End Sub

Private Sub PreencherTipoPgto()
Dim varTexto As String
varTexto = cboTipoPgto.Text

cboTipoPgto.Clear
cboTipoPgto.AddItem "TODOS"
cboTipoPgto.AddItem "DINHEIRO"
cboTipoPgto.AddItem "PIX"
cboTipoPgto.AddItem "PROMISSÓRIA"
cboTipoPgto.AddItem "CARTÃO"
cboTipoPgto.AddItem "CHEQUE"
cboTipoPgto.AddItem "BOLETO"
cboTipoPgto.AddItem "TRANSFERÊNCIA"

cboTipoPgto.Text = varTexto
moCombo.AttachTo cboTipoPgto
End Sub

Private Sub VerificarCaixa()
Dim sSQL As String
Dim r As ADODB.Recordset

If Grid.TextMatrix(Grid.Row, 19) = "0" Then
    CAIXA_FECHADO = True
Else
    sSQL = "SELECT * " & _
           "FROM caixa_dia " & _
           "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 18) & "') and (codcaixa = '" & Grid.TextMatrix(Grid.Row, 19) & "');"
    Set r = dbData.OpenRecordset(sSQL)
    
    CAIXA_FECHADO = r("status")
End If
End Sub

Private Sub Imprimir_Pedido()
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

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
        
codPedido = Grid.TextMatrix(Grid.Row, 2)

sSQL = "SELECT produtos.descricao as var_desc, produtos.fabricante as vFab, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
         "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
         "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") order by pedidos_itens.Codigo desc"
Set r = dbData.OpenRecordset(sSQL)

Dim vQuantItens As Integer

vQuantItens = r.RecordCount
varImpPDF = False
Me.Hide

If Grid.TextMatrix(Grid.Row, 5) = "ORÇAMENTO" Or Grid.TextMatrix(Grid.Row, 5) = "CONSIGNADO" Then
    If vQuantItens < 18 Then
        REL_Pedido_Orcamento.loadPedidos Grid.TextMatrix(Grid.Row, 2)
    Else
        If Grid.TextMatrix(Grid.Row, 11) = "0,00" Then
            'txtDescItens.Text = FormatNumber(0, 2)
            vDescItensVenda = FormatNumber(0, 2)
        Else
            'converter o desconto em dinheiro em porcentagem
            If Grid.TextMatrix(Grid.Row, 13) = "" Then Exit Sub
            If Grid.TextMatrix(Grid.Row, 10) = "" Then Exit Sub
            
            B = Grid.TextMatrix(Grid.Row, 13)
            A = Grid.TextMatrix(Grid.Row, 10)
            
            varValorDescProc = ((B - A) / A) * 100
            vDescItensVenda = Abs(FormatNumber(varValorDescProc, 2))
            vDescItensVenda = FormatNumber(vDescItensVenda, 2)
        End If
        
        'nome do funcionario
        
        Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 6) & ");")
        
        Set REL_Pedido_Completo.ReportMain1.Recordset = r
        REL_Pedido_Completo.txtDHead.Caption = "RELATÓRIO DO PEDIDO Nº " & Grid.TextMatrix(Grid.Row, 2)
        REL_Pedido_Completo.Mostrar_Parcelas Grid.TextMatrix(Grid.Row, 2)
        
        REL_Pedido_Completo.rfSubTotal.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 10), 2)
        REL_Pedido_Completo.txtDescontoRS.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 11), 2)
        REL_Pedido_Completo.rfTotal.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 13), 2)
        REL_Pedido_Completo.rfDesc.Caption = FormatNumber(vDescItensVenda, 2)
        
        REL_Pedido_Completo.rfCliente.Caption = Grid.TextMatrix(Grid.Row, 9)
        REL_Pedido_Completo.rfData.Caption = Grid.TextMatrix(Grid.Row, 4)
        REL_Pedido_Completo.rfForma.Caption = Grid.TextMatrix(Grid.Row, 7)
        REL_Pedido_Completo.rfFunc.Caption = rF("nome")
        REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_Pedido_Completo.ReportMain1.Ativar
        Unload REL_Pedido_Completo
    End If
Else
    If vQuantItens < 18 Then
        If Grid.TextMatrix(Grid.Row, 7) = "À Prazo" Then
            If vTipoParcelaImpressao = 1 Then
                REL_Pedido_Mod05.loadPedidos Grid.TextMatrix(Grid.Row, 2)
            Else
                REL_Pedido_APrazo.loadPedidos Grid.TextMatrix(Grid.Row, 2)
            End If
   '
        ElseIf Grid.TextMatrix(Grid.Row, 7) = "À Vista" Then
           REL_Pedido_Mod06.loadPedidos Grid.TextMatrix(Grid.Row, 2)
        ElseIf Grid.TextMatrix(Grid.Row, 7) = "Orçamento" Then
           REL_Pedido_Orcamento.loadPedidos Grid.TextMatrix(Grid.Row, 2)
        End If
    Else
        If Grid.TextMatrix(Grid.Row, 11) = "0,00" Then
            'txtDescItens.Text = FormatNumber(0, 2)
            vDescItensVenda = FormatNumber(0, 2)
        Else
            'converter o desconto em dinheiro em porcentagem
            If Grid.TextMatrix(Grid.Row, 13) = "" Then Exit Sub
            If Grid.TextMatrix(Grid.Row, 10) = "" Then Exit Sub
            
            B = Grid.TextMatrix(Grid.Row, 13)
            A = Grid.TextMatrix(Grid.Row, 10)
            
            varValorDescProc = ((B - A) / A) * 100
            vDescItensVenda = Abs(FormatNumber(varValorDescProc, 2))
            vDescItensVenda = FormatNumber(vDescItensVenda, 2)
        End If
        
        'nome do funcionario
        
        Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 6) & ");")
        
        Set REL_Pedido_Completo.ReportMain1.Recordset = r
        REL_Pedido_Completo.txtDHead.Caption = "RELATÓRIO DO PEDIDO Nº " & Grid.TextMatrix(Grid.Row, 2)
        REL_Pedido_Completo.Mostrar_Parcelas Grid.TextMatrix(Grid.Row, 2)
        
        REL_Pedido_Completo.rfSubTotal.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 10), 2)
        REL_Pedido_Completo.txtDescontoRS.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 11), 2)
        REL_Pedido_Completo.rfTotal.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 13), 2)
        REL_Pedido_Completo.rfDesc.Caption = FormatNumber(vDescItensVenda, 2)
        
        REL_Pedido_Completo.rfCliente.Caption = Grid.TextMatrix(Grid.Row, 9)
        REL_Pedido_Completo.rfData.Caption = Grid.TextMatrix(Grid.Row, 4)
        REL_Pedido_Completo.rfForma.Caption = Grid.TextMatrix(Grid.Row, 7)
        REL_Pedido_Completo.rfFunc.Caption = rF("nome")
        REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_Pedido_Completo.ReportMain1.Ativar
        Unload REL_Pedido_Completo
    End If
End If
Me.Show
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
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CboCliente_LostFocus()
cboCliente_Click
End Sub

Private Sub cboCriterios_Click()
cboCriterios_LostFocus
End Sub


Private Sub cboCriterios_GotFocus()
PreencherCriterios
SelectControl cboCriterios
End Sub


Private Sub cboCriterios_LostFocus()
If cboStatus.Text = "" Then Exit Sub

If cboCriterios.Text = "NENHUM" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = False
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
ElseIf cboCriterios.Text = "CÓD. PEDIDO" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = False
    lblCodPedido.Visible = True
    txtCodPedido.Visible = True
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = True
    optEsc.Visible = True
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cboStatus.ListIndex = 0
    txtCodPedido.SetFocus
ElseIf cboCriterios.Text = "CLIENTE" Then
    lblCliente.Visible = True
    cboCliente.Visible = True
    'lblDesc.Visible = False
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cboStatus.ListIndex = 0
    cboCliente.SetFocus
ElseIf cboCriterios.Text = "DATA" Then
   lblCliente.Visible = False
   cboCliente.Visible = False
   'lblDesc.Visible = False
   lblCodPedido.Visible = False
   txtCodPedido.Visible = False
   'lblCodBarra.Visible = True
   txtCodCliente.Text = ""
   optDig.Visible = False
   optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = True
    mskData.Visible = True
    cmdCal1.Visible = True
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
   cboStatus.ListIndex = 0
'   mskData.SetFocus
ElseIf cboCriterios.Text = "MENSAL" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = True
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = True
    cboMes.Visible = True
    lblAno.Visible = True
    cboAno.Visible = True
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cboStatus.ListIndex = 0
    cboMes.SetFocus
ElseIf cboCriterios.Text = "PRODUTO" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = False
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = True
    cboProduto.Visible = True
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cboProduto.SetFocus
ElseIf cboCriterios.Text = "CÓD. BARRA" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = False
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    'cboStatus.ListIndex = 0
    lblCodBarra.Visible = True
    txtCodBarra.Visible = True
    txtCodBarra.SetFocus
End If
End Sub



Private Sub cboformaPgto_Click()
cboFormaPgto_LostFocus
End Sub

Private Sub cboformaPgto_GotFocus()
PreencherFormaPgto
SelectControl cboFormaPgto
End Sub


Private Sub cboFormaPgto_LostFocus()
If cboFormaPgto.Text = "" Then Exit Sub
End Sub

Private Sub cboIndice_Click()
cboIndice_LostFocus
End Sub


Private Sub cboIndice_GotFocus()
PreencherIndice
SelectControl cboIndice
End Sub

Private Sub cboIndice_LostFocus()
If cboIndice.Text = "" Then Exit Sub
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


Private Sub cboProduto_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

cboProduto.Clear

sSQL = "SELECT * FROM produtos ORDER BY descricao;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboProduto.AddItem ValidateNull(r("descricao"))
   cboProduto.ItemData(cboProduto.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboProduto
End Sub


Private Sub cboProduto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboProduto_LostFocus()
On Error GoTo TrataErro

If cboProduto.Text = "" Then txtCodProduto.Text = "": Exit Sub
If cboProduto.ListIndex = -1 Then txtCodProduto.Text = "": Exit Sub
txtCodProduto = cboProduto.ItemData(cboProduto.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboStatus_Click()
cboStatus_LostFocus
End Sub


Private Sub cboStatus_GotFocus()
PreencherStatus
SelectControl cboStatus
End Sub


Private Sub cboStatus_LostFocus()
If cboStatus.Text = "" Then Exit Sub
'Mostrar_Pedido
End Sub

Private Sub cboTipoPedido_Change()
'If cboTipoPedido.Text = "CONSIGNADO" Then
'    chkMostarCancelados.Value = Unchecked
'Else
'    chkMostarCancelados.Value = Checked
'End If
End Sub

Private Sub cboTipoPedido_Click()
cboTipoPedido_Change
cboTipoPedido_LostFocus
End Sub

Private Sub cboTipoPedido_GotFocus()
PreencherTipoPedido
SelectControl cboTipoPedido
End Sub


Private Sub cboTipoPedido_LostFocus()
If cboTipoPedido.Text = "" Then Exit Sub
End Sub


Private Sub cboTipoPgto_Click()
cboTipoPgto_LostFocus
End Sub

Private Sub cboTipoPgto_GotFocus()
PreencherTipoPgto
SelectControl cboTipoPgto
End Sub


Private Sub cboTipoPgto_LostFocus()
If cboTipoPgto.Text = "" Then Exit Sub
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

Private Sub cmdExcluirPedido_Click()

If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 1) = "ALUGUEL" Then
    MsgBox "Não é possível cancelar um aluguel de equipamento!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 15) = "SIM" Then
    MsgBox "Não é possível cancelar um pedido já cancelado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 16) = "SIM" Then
    MsgBox "Não é possível cancelar um pedido que já emitiu NFCE!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

VerificarCaixa

If Grid.TextMatrix(Grid.Row, 1) <> "ORÇAMENTO" And Grid.TextMatrix(Grid.Row, 1) <> "CONSIGNADO" Then
    If CAIXA_FECHADO = True Then
        MsgBox "O " & Grid.TextMatrix(Grid.Row, 18) & " com o Cód.: " & Format(Grid.TextMatrix(Grid.Row, 19), "0000") & ", desse pedido encontra-se fechado!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
End If

If ShowMsg("Tem certeza que deseja cancelar o pedido " & Grid.TextMatrix(Grid.Row, 2) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "INSERT INTO Pedidos_Reabertura (COD_USUARIO, LOGIN, VLR_PEDIDO, DATA, HORA, CANCELADO, COD_PEDIDO) VALUES (" & lblCodUser2.Caption & ", '" & lblUser2.Caption & "', " & Replace(CCur(Grid.TextMatrix(Grid.Row, 13)), ",", ".") & ", CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4).Text, ocDATA) & "', 103), '" & Format(Now, ocHORA) & "', 1, " & Replace(CCur(Grid.TextMatrix(Grid.Row, 2)), ",", ".") & ");"
    'Retornar a quantidade de produtos ao estoque
    If cboStatus.Text <> "VAZIO" Then
        dbData.Execute "UPDATE produtos SET " & _
                       "quant_estoque = quant_estoque + pedidos_itens.quantidade " & _
                       "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                       "WHERE (pedidos_itens.cod_pedido = " & Grid.TextMatrix(Grid.Row, 2) & ")"
    End If
    
    'Apaga as parcelas do pedido
    dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & Grid.TextMatrix(Grid.Row, 2) & ");"
    
    'Colocar como cancelado os produtos do pedido
    dbData.Execute "UPDATE pedidos_itens SET cancelado = 1 WHERE (cod_pedido = " & Grid.TextMatrix(Grid.Row, 2) & ");"
    
    'Apaga a venda
    'dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & Grid.TextMatrix(Grid.Row, 2) & ");"
    dbData.Execute "UPDATE pedidos SET cancelado = 1 WHERE (cod_pedido = " & Grid.TextMatrix(Grid.Row, 2) & ");"

    
    Cancelar_NFCe " & Grid.TextMatrix(Grid.Row, 2) & "
End If

Mostrar_Pedido
End Sub

Private Sub cmdExibir_Click()
Mostrar_Pedido
cmdModificar.Enabled = False
cmdModificarConsignado.Enabled = False
'cmdPedidoAbrir.Caption = "Reabrir Venda"

If cboTipoPedido.Text = "VENDA" Then
    cmdPedidoAbrir.Caption = "REABRIR"
ElseIf cboTipoPedido.Text = "ORÇAMENTO" Then
    cmdPedidoAbrir.Caption = "CONVERTER":
End If
End Sub

Private Sub cmdImprimir_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhuma venda consultada!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide
Set r = dbData.OpenRecordset(printSQL)
Set REL_Estorno.Relatorio.Recordset = r
REL_Estorno.dfQuant.Caption = lblTotalVendas.Caption
REL_Estorno.dfBruto.Caption = Format(lblTotalGrid.Caption, "##,##0.00")

'REL_Parcelas_Agrupado.rfForma.Caption = cboForma.Text
'REL_Parcelas_Agrupado.rfTipo.Caption = cboTipo.Text

REL_Estorno.Relatorio.Ativar
Unload REL_Estorno

Me.Show
End Sub

Private Sub cmdModificar_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 15) = "SIM" Then
    MsgBox "Não é possível abrir um pedido cancelado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 1) = "ALUGUEL" Then
    MsgBox "Não é possível abrir um aluguel de equipamento!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 16) = "SIM" Then
    MsgBox "Não é possível abrir um pedido que já emitiu NFCE!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

VerificarCaixa

If Grid.TextMatrix(Grid.Row, 1) <> "ORÇAMENTO" And Grid.TextMatrix(Grid.Row, 1) <> "CONSIGNADO" Then
    If CAIXA_FECHADO = True Then
        MsgBox "O " & Grid.TextMatrix(Grid.Row, 18) & " com o Cód.: " & Format(Grid.TextMatrix(Grid.Row, 19), "0000") & ", desse pedido encontra-se fechado!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
End If

Dim varTipoPedido As String
varTipoPedido = Grid.TextMatrix(Grid.Row, 1)

'Dim codPedido As String
codPedido = Grid.TextMatrix(Grid.Row, 2)

If ShowMsg("Tem certeza que deseja editar o orçamento " & Grid.TextMatrix(Grid.Row, 2) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "INSERT INTO Pedidos_Reabertura (COD_USUARIO, LOGIN, VLR_PEDIDO, DATA, HORA, CANCELADO, COD_PEDIDO, STATUS_PEDIDO) VALUES (" & lblCodUser2.Caption & ", '" & lblUser2.Caption & "', " & Replace(CCur(Grid.TextMatrix(Grid.Row, 13)), ",", ".") & ", CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4).Text, ocDATA) & "', 103), '" & Format(Now, ocHORA) & "', 0, " & Replace(CCur(Grid.TextMatrix(Grid.Row, 2)), ",", ".") & ", 0);"
    PDV.frmAvancado.Visible = False
    PDV.frmSenha.Visible = False
    Unload Estonar
    vPedirPeso = False
    vTipoEdicao = "EDITAR"
    PDV.lblEstornar.Caption = "ESTORNO"
    PDV.lblTipoPedido.Caption = varTipoPedido
    PDV.txtCodPedido.Text = codPedido
End If
End Sub

Private Sub cmdModificarConsignado_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 15) = "SIM" Then
    MsgBox "Não é possível abrir um pedido cancelado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 1) = "ALUGUEL" Then
    MsgBox "Não é possível abrir um aluguel de equipamento!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 16) = "SIM" Then
    MsgBox "Não é possível abrir um pedido que já emitiu NFCE!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

VerificarCaixa

If Grid.TextMatrix(Grid.Row, 1) <> "ORÇAMENTO" And Grid.TextMatrix(Grid.Row, 1) <> "CONSIGNADO" Then
    If CAIXA_FECHADO = True Then
        MsgBox "O " & Grid.TextMatrix(Grid.Row, 18) & " com o Cód.: " & Format(Grid.TextMatrix(Grid.Row, 19), "0000") & ", desse pedido encontra-se fechado!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
End If

Dim varTipoPedido As String
varTipoPedido = Grid.TextMatrix(Grid.Row, 1)

'Dim codPedido As String
codPedido = Grid.TextMatrix(Grid.Row, 2)

If ShowMsg("Tem certeza que deseja editar o consignado " & Grid.TextMatrix(Grid.Row, 2) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "INSERT INTO Pedidos_Reabertura (COD_USUARIO, LOGIN, VLR_PEDIDO, DATA, HORA, CANCELADO, COD_PEDIDO, STATUS_PEDIDO) VALUES (" & lblCodUser2.Caption & ", '" & lblUser2.Caption & "', " & Replace(CCur(Grid.TextMatrix(Grid.Row, 13)), ",", ".") & ", CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4).Text, ocDATA) & "', 103), '" & Format(Now, ocHORA) & "', 0, " & Replace(CCur(Grid.TextMatrix(Grid.Row, 2)), ",", ".") & ", 0);"
    PDV.frmAvancado.Visible = False
    PDV.frmSenha.Visible = False
    Unload Estonar
    vPedirPeso = False
    vTipoEdicao = "EDITAR"
    PDV.lblEstornar.Caption = "ESTORNO"
    PDV.lblTipoPedido.Caption = varTipoPedido
    PDV.txtCodPedido.Text = codPedido
End If
End Sub


Private Sub cmdMostrarProdutos_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 2), Grid.TextMatrix(Grid.Row, 4)
Parcelas_Consulta_Produtos.Show 1
End Sub

Private Sub Imprimir_PedidoPDF()
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

codPedido = Grid.TextMatrix(Grid.Row, 2)

'sSQL = "SELECT produtos.descricao as var_desc, produtos.fabricante as vFab, quantidade, preco, pedidos_itens.subtotal as vSubtotal, produtos.codigo as vCodProd " & _
         "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
         "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") order by pedidos_itens.Codigo desc"
sSQL = "SELECT produtos.descricao as var_desc, produtos.fabricante as vFab, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
         "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
         "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") order by pedidos_itens.Codigo desc"
Set r = dbData.OpenRecordset(sSQL)
'Set r = dbData.OpenRecordset(sSQL)

Dim vQuantItens As Integer

vQuantItens = r.RecordCount

varImpPDF = True
Me.Hide

If Grid.TextMatrix(Grid.Row, 5) = "ORÇAMENTO" Or Grid.TextMatrix(Grid.Row, 5) = "CONSIGNADO" Then
    If vQuantItens < 18 Then
        REL_Pedido_Orcamento.loadPedidos Grid.TextMatrix(Grid.Row, 2)
    Else
        If Grid.TextMatrix(Grid.Row, 11) = "0,00" Then
            'txtDescItens.Text = FormatNumber(0, 2)
            vDescItensVenda = FormatNumber(0, 2)
        Else
            'converter o desconto em dinheiro em porcentagem
            If Grid.TextMatrix(Grid.Row, 13) = "" Then Exit Sub
            If Grid.TextMatrix(Grid.Row, 10) = "" Then Exit Sub
            
            B = Grid.TextMatrix(Grid.Row, 13)
            A = Grid.TextMatrix(Grid.Row, 10)
            
            varValorDescProc = ((B - A) / A) * 100
            vDescItensVenda = Abs(FormatNumber(varValorDescProc, 2))
            vDescItensVenda = FormatNumber(vDescItensVenda, 2)
        End If
        
        'nome do funcionario
        
        Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 6) & ");")
        
        Set REL_Pedido_Completo.ReportMain1.Recordset = r
        REL_Pedido_Completo.txtDHead.Caption = "RELATÓRIO DO ORÇAMENTO Nº " & Grid.TextMatrix(Grid.Row, 2)
        REL_Pedido_Completo.Mostrar_Parcelas Grid.TextMatrix(Grid.Row, 2)
        
        REL_Pedido_Completo.rfSubTotal.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 10), 2)
        REL_Pedido_Completo.txtDescontoRS.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 11), 2)
        REL_Pedido_Completo.rfTotal.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 13), 2)
        REL_Pedido_Completo.rfDesc.Caption = FormatNumber(vDescItensVenda, 2)
        
        REL_Pedido_Completo.rfCliente.Caption = Grid.TextMatrix(Grid.Row, 9)
        REL_Pedido_Completo.rfData.Caption = Grid.TextMatrix(Grid.Row, 4)
        REL_Pedido_Completo.rfForma.Caption = Grid.TextMatrix(Grid.Row, 7)
        REL_Pedido_Completo.rfFunc.Caption = rF("nome")
        REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_Pedido_Completo.ReportMain1.Ativar
        Unload REL_Pedido_Completo
        
        'Set REL_Pedido_Completo.ReportMain1.Recordset = r
        
        'REL_Pedido_Completo.txtDHead.Caption = "RELATORIO COMPLETO - ORÇAMENTO Nº " & txtCodPedido.Text
        'REL_Pedido_Completo.rfSubTotal.Caption = Format(Grid.TextMatrix(Grid.Row, 10), "#,##0.00")
        'REL_Pedido_Completo.rfDesc.Caption = Format(Grid.TextMatrix(Grid.Row, 11), "#,##0.00")
        'REL_Pedido_Completo.rfTotal.Caption = Format(Grid.TextMatrix(Grid.Row, 13), "#,##0.00")
        'REL_Pedido_Completo.rfCliente.Caption = Grid.TextMatrix(Grid.Row, 9)
        'REL_Pedido_Completo.rfData.Caption = Grid.TextMatrix(Grid.Row, 4)
        'REL_Pedido_Completo.rfForma.Caption = Grid.TextMatrix(Grid.Row, 7)
        'REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        'REL_Pedido_Completo.ReportMain1.Ativar
        'Unload REL_Pedido_Completo
    End If
Else
    If vQuantItens < 18 Then
        If Grid.TextMatrix(Grid.Row, 7) = "À Prazo" Then
           REL_Pedido_Mod05.loadPedidos Grid.TextMatrix(Grid.Row, 2)
        ElseIf Grid.TextMatrix(Grid.Row, 7) = "À Vista" Then
           REL_Pedido_Mod06.loadPedidos Grid.TextMatrix(Grid.Row, 2)
        ElseIf Grid.TextMatrix(Grid.Row, 7) = "Orçamento" Then
           REL_Pedido_Orcamento.loadPedidos Grid.TextMatrix(Grid.Row, 2)
        End If
    Else

        If Grid.TextMatrix(Grid.Row, 11) = "0,00" Then
            'txtDescItens.Text = FormatNumber(0, 2)
            vDescItensVenda = FormatNumber(0, 2)
        Else
            'converter o desconto em dinheiro em porcentagem
            If Grid.TextMatrix(Grid.Row, 13) = "" Then Exit Sub
            If Grid.TextMatrix(Grid.Row, 10) = "" Then Exit Sub
            
            B = Grid.TextMatrix(Grid.Row, 13)
            A = Grid.TextMatrix(Grid.Row, 10)
            
            varValorDescProc = ((B - A) / A) * 100
            vDescItensVenda = Abs(FormatNumber(varValorDescProc, 2))
            vDescItensVenda = FormatNumber(vDescItensVenda, 2)
        End If
        
        'nome do funcionario
        
        Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 6) & ");")
        
        Set REL_Pedido_Completo.ReportMain1.Recordset = r
        REL_Pedido_Completo.txtDHead.Caption = "RELATÓRIO DO PEDIDO Nº " & Grid.TextMatrix(Grid.Row, 2)
        REL_Pedido_Completo.Mostrar_Parcelas Grid.TextMatrix(Grid.Row, 2)
        
        REL_Pedido_Completo.rfSubTotal.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 10), 2)
        REL_Pedido_Completo.txtDescontoRS.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 11), 2)
        REL_Pedido_Completo.rfTotal.Caption = FormatNumber(Grid.TextMatrix(Grid.Row, 13), 2)
        REL_Pedido_Completo.rfDesc.Caption = FormatNumber(vDescItensVenda, 2)
        
        REL_Pedido_Completo.rfCliente.Caption = Grid.TextMatrix(Grid.Row, 9)
        REL_Pedido_Completo.rfData.Caption = Grid.TextMatrix(Grid.Row, 4)
        REL_Pedido_Completo.rfForma.Caption = Grid.TextMatrix(Grid.Row, 7)
        REL_Pedido_Completo.rfFunc.Caption = rF("nome")
        REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_Pedido_Completo.ReportMain1.Ativar
        Unload REL_Pedido_Completo

        'Set REL_Pedido_Completo.ReportMain1.Recordset = r
        
        'REL_Pedido_Completo.txtDHead.Caption = "RELATORIO COMPLETO - PEDIDO Nº " & txtCodPedido.Text
        'REL_Pedido_Completo.rfSubTotal.Caption = Format(Grid.TextMatrix(Grid.Row, 10), "#,##0.00")
        'REL_Pedido_Completo.rfDesc.Caption = Format(Grid.TextMatrix(Grid.Row, 11), "#,##0.00")
        'REL_Pedido_Completo.rfTotal.Caption = Format(Grid.TextMatrix(Grid.Row, 13), "#,##0.00")
        'REL_Pedido_Completo.rfCliente.Caption = Grid.TextMatrix(Grid.Row, 9)
        'REL_Pedido_Completo.rfData.Caption = Grid.TextMatrix(Grid.Row, 4)
        'REL_Pedido_Completo.rfForma.Caption = Grid.TextMatrix(Grid.Row, 7)
        'REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        'REL_Pedido_Completo.ReportMain1.Ativar
        'Unload REL_Pedido_Completo
    End If
End If
Me.Show
End Sub


Private Sub cmdPDF_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 1) = "ALUGUEL" Then
    MsgBox "Não é possível gerar um PDF de um aluguel de equipamento!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 15) = "SIM" Then
    MsgBox "Não é possivel gerar um PDF de um pedido já cancelado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

codPedido = Grid.TextMatrix(Grid.Row, 2)

If ShowMsg("Tem certeza que deseja gerar um PDF do pedido " & Grid.TextMatrix(Grid.Row, 2) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    Imprimir_PedidoPDF
End If
End Sub

Private Sub cmdPedidoAbrir_Click()
'If txtCodPedidoCerto.Text = "" Then Exit Sub

If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 15) = "SIM" Then
    MsgBox "Não é possível abrir um pedido cancelado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 1) = "ALUGUEL" Then
    MsgBox "Não é possível abrir um aluguel de equipamento!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 16) = "SIM" Then
    MsgBox "Não é possível abrir um pedido que já emitiu NFCE!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

VerificarCaixa

If Grid.TextMatrix(Grid.Row, 1) <> "ORÇAMENTO" And Grid.TextMatrix(Grid.Row, 1) <> "CONSIGNADO" Then
    If CAIXA_FECHADO = True Then
        MsgBox "O " & Grid.TextMatrix(Grid.Row, 18) & " com o Cód.: " & Format(Grid.TextMatrix(Grid.Row, 19), "0000") & ", desse pedido encontra-se fechado!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
End If

Dim varTipoPedido As String
varTipoPedido = Grid.TextMatrix(Grid.Row, 1)

'Dim codPedido As String
codPedido = Grid.TextMatrix(Grid.Row, 2)

If cboTipoPedido.Text = "ORÇAMENTO" Then
ElseIf cboTipoPedido.Text = "ORÇAMENTO" Then
Else

End If
If ShowMsg("Tem certeza que deseja reabrir o pedido " & Grid.TextMatrix(Grid.Row, 2) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "INSERT INTO Pedidos_Reabertura (COD_USUARIO, LOGIN, VLR_PEDIDO, DATA, HORA, CANCELADO, COD_PEDIDO, STATUS_PEDIDO) VALUES (" & lblCodUser2.Caption & ", '" & lblUser2.Caption & "', " & Replace(CCur(Grid.TextMatrix(Grid.Row, 13)), ",", ".") & ", CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4).Text, ocDATA) & "', 103), '" & Format(Now, ocHORA) & "', 0, " & Replace(CCur(Grid.TextMatrix(Grid.Row, 2)), ",", ".") & ", 0);"
    PDV.frmAvancado.Visible = False
    PDV.frmSenha.Visible = False
    Unload Estonar
    vPedirPeso = False
    vTipoEdicao = "REABRIR"
    PDV.lblEstornar.Caption = "ESTORNO"
    PDV.lblTipoPedido.Caption = varTipoPedido
    PDV.txtCodPedido.Text = codPedido
End If
End Sub

Private Sub cmdPedidoImprimir_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 1) = "ALUGUEL" Then
    MsgBox "Não é possível imprimir um aluguel de equipamento!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 15) = "SIM" Then
    MsgBox "Não é possivel imprimir um pedido já cancelado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

codPedido = Grid.TextMatrix(Grid.Row, 2)

If ShowMsg("Tem certeza que deseja reimprimir o pedido " & Grid.TextMatrix(Grid.Row, 2) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then

   Dim oCfg As ConfigItem
   Dim ii As Integer

If Grid.TextMatrix(Grid.Row, 7) = "À Prazo" Then
   Dim NumCopiasAP As Integer
   Dim varConfEntregaAP As Boolean
   Dim varNumCopiasAP As Integer
   Dim varConfImpressaoAP As Boolean
   Dim varImpressaoLiberadaAP As Boolean
   Dim varConfTipoImpressaoAP As Integer
   
   Set oCfg = sysConfig("COPIAS_AP")
   varNumCopiasAP = CInt(oCfg.Value)
   Set oCfg = Nothing
   
   Set oCfg = sysConfig("ENTREGA_AP")
   varConfEntregaAP = CBool(oCfg.Value)
   Set oCfg = Nothing

   Set oCfg = sysConfig("CONF_IMPRESSAO_AP")
   varConfImpressaoAP = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   Set oCfg = sysConfig("IMP_AP")
   varImpressaoLiberadaAP = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   Set oCfg = sysConfig("IMPRIMIR_AP")
   varConfTipoImpressaoAP = CInt(oCfg.Value)
   Set oCfg = Nothing
   
   If varNumCopiasAP <> 0 Then  'saber a quantidade de copias
      If varConfEntregaAP Then
         If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            NumCopiasAP = varNumCopiasAP + 1
         Else
            NumCopiasAP = varNumCopiasAP
         End If
      Else
         NumCopiasAP = varNumCopiasAP
      End If
   Else
      NumCopiasAP = 1
   End If
   
   If varImpressaoLiberadaAP = True Then       'Confirma se vai ter impressão
         If varConfTipoImpressaoAP = 1 Then
            For ii = 1 To NumCopiasAP
               Imprimir_Pedido
            Next
         ElseIf varConfTipoImpressaoAP = 2 Then
            For ii = 1 To NumCopiasAP
               Imprimir_CupomSerrilha
            Next
         ElseIf varConfTipoImpressaoAP = 3 Then
            For ii = 1 To NumCopiasAP
               Imprimir_CupomGuilhotina
            Next
         End If
   End If
ElseIf Grid.TextMatrix(Grid.Row, 7) = "À Vista" Then
   Dim NumCopiasAV As Integer
   Dim varConfEntregaAV As Boolean
   Dim varNumCopiasAV As Integer
   Dim varConfImpressaoAV As Boolean
   Dim varImpressaoLiberadaAV As Boolean
   Dim varConfTipoImpressaoAV As Integer
   
   Set oCfg = sysConfig("COPIAS_AV")
   varNumCopiasAV = CInt(oCfg.Value)
   Set oCfg = Nothing
   
   Set oCfg = sysConfig("ENTREGA_AV")
   varConfEntregaAV = CBool(oCfg.Value)
   Set oCfg = Nothing

   Set oCfg = sysConfig("CONF_IMPRESSAO_AV")
   varConfImpressaoAV = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   Set oCfg = sysConfig("IMP_AV")
   varImpressaoLiberadaAV = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   Set oCfg = sysConfig("IMPRIMIR_AV")
   varConfTipoImpressaoAV = CInt(oCfg.Value)
   Set oCfg = Nothing
   
   If varNumCopiasAV <> 0 Then  'saber a quantidade de copias
      If varConfEntregaAV Then
         If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            NumCopiasAV = varNumCopiasAV + 1
         Else
            NumCopiasAV = varNumCopiasAV
         End If
      Else
         NumCopiasAV = varNumCopiasAV
      End If
   Else
      NumCopiasAV = 1
   End If
   
   If varImpressaoLiberadaAV = True Then       'Confirma se vai ter impressão
         If varConfTipoImpressaoAV = 1 Then
            For ii = 1 To NumCopiasAV
               Imprimir_Pedido
            Next
         ElseIf varConfTipoImpressaoAV = 2 Then
            For ii = 1 To NumCopiasAV
               Imprimir_CupomSerrilha
            Next
         ElseIf varConfTipoImpressaoAV = 3 Then
            For ii = 1 To NumCopiasAV
               Imprimir_CupomGuilhotina
            Next
         End If
   End If
End If
End If
End Sub
Private Sub Imprimir_CupomGuilhotina()
'On Error GoTo Tratar_Erro
Dim sSQL As String
Dim r As ADODB.Recordset
Dim rP As ADODB.Recordset
Dim rPR As ADODB.Recordset
Dim rI As ADODB.Recordset
Dim rF As ADODB.Recordset
Dim rParc As ADODB.Recordset

Dim i As Integer
Dim f As Integer

If codPedido = "" Then Exit Sub

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

'consultar funcionario do pedido
Set rP = dbData.OpenRecordset("SELECT cod_funcionario, TIPO_PAGAMENTO, PAGAMENTO, DATA_COMPRA, SUBTOTAL, ValorAcrescReal, TOTAL FROM pedidos WHERE (cod_pedido = " & codPedido & ");")
Set rPR = dbData.OpenRecordset("SELECT * FROM Pedidos_Recebedor WHERE (cod_pedido = " & codPedido & ");")
Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rP("cod_funcionario") & ");")
Set rParc = dbData.OpenRecordset("SELECT COD_PEDIDO, NUMERO, PAGAMENTO, DATA, VALOR_FINAL, (CASE WHEN FORMA_PGTO = 'CARTAO' THEN (CASE WHEN TIPO_CARTAO = 'D' THEN 'CARTÃO DÉBITO' ELSE 'CARTÃO CRÉDITO' END) ELSE isnull(FORMA_PGTO, '') END) AS varFormaPgto FROM parcelas WHERE (cod_pedido = " & codPedido & ") order by NUMERO;")

'Recupera um número de arquivo disponível
f = FreeFile()

'pegar o nome da impressora no ini
Dim oIni As Ini
''Dim var_ImpTermica As String

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
Set oIni = Nothing

'logomarca impressa do cupom
Dim sLogo As String
Set oCfg = sysConfig("LOGO_CUPOM")
sLogo = oCfg.Value
Set oCfg = Nothing
If Dir$(sLogo) <> "" Then Set imLogoCupom.Picture = LoadPicture(sLogo)

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = var_ImpTermica Then
      Set Printer = Prt
      Exit For
   End If
Next

'Open "LPT1" For Output As #1
'Open "\\balcao04\TERMICA" For Output As #f
   

'Open "LPT1" For Output As #1
'Open "\\CAIXAO1\termica" For Output As #1

   With Printer
      .ScaleMode = vbPixels
      .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
      
      For i = 1 To 6
         Printer.Print " "
      Next
      
      .ScaleMode = vbCentimeters
      .FontName = "courier new"
      '.PrintQuality = vbPRPQHigh
      
      Fonte 8, False, False
      Printer.Print String(40, "-")
      Fonte 10, True, False
      Printer.Print Tab((35 - Len(r("fantasia"))) / 2); r("fantasia")   'Esse /2 é p/ centralizar
      Fonte 10, False, False
      'Printer.Print Tab((35 - Len(r("razao"))) / 2); r("razao")
      Printer.Print " "
      Fonte 8, False, False
      Printer.Print r("endereco") & ", " & r("cidade") & "-" & r("estado")
      Printer.Print "FONE: "; r("telefone")                                        '& " - (89) 9986-3739"
      Fonte 8, False, False
      Printer.Print "CNPJ:"; r("cnpj") & "  IE:" & r("ie")
      Printer.Print " "
      
      Fonte 10, True, False
      If Grid.TextMatrix(Grid.Row, 1) = "ORÇAMENTO" Then
         Printer.Print Tab(10); "O R Ç A M E N T O"
     ElseIf Grid.TextMatrix(Grid.Row, 1) = "CONSIGNADO" Then
        Printer.Print Tab(10); "C O N S I G N A D O"
     Else
         Printer.Print Tab(10); "CUPOM DE VENDA"
     End If
      
      Fonte 8, False, False
      Printer.Print Tab(2); Format(rP("DATA_COMPRA"), "dd/mm/yy"); " "; Format(Time, "hh:mm"); " "; "CÓD:"; Format(codPedido, "000000"); " "; rF("nome")

      Fonte 8, False, False
      Printer.Print Tab(2); "Tipo de Pgto:"; rP("TIPO_PAGAMENTO"); "  "; "Forma:"; rP("PAGAMENTO")

      Fonte 8, False, False
      Printer.Print String(40, "-")
      Printer.Print Tab(0); "DESCRIÇÃO";
      Printer.Print Tab(20); "PREÇO";
      Printer.Print Tab(26); "QTDE";
      Printer.Print Tab(35); "TOTAL"
      Printer.Print String(40, "-")
      
      sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.cod_pedido, pedidos_itens.preco, pedidos_itens.quantidade, (pedidos_itens.preco * pedidos_itens.quantidade) as total, produtos.descricao " & _
         "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
         "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") ORDER BY pedidos_itens.codigo DESC;"
      Set rI = dbData.OpenRecordset(sSQL)

      Do While Not rI.EOF
         '---------------imprime os dados da tabela----------------------------
         Printer.Print Tab(0); rI("descricao");
         Printer.Print Tab(19); Format$(Format$(rI("preco"), "0.00"), "@@@@@@@");
         Printer.Print Tab(26); Format$(Format$(rI("quantidade"), "0.000"), "@@@@@@@");
         Printer.Print Tab(33); Format$(Format$(rI("total"), "0.00"), "@@@@@@@")
         
         rI.MoveNext                 'vai para o proximo registro
      Loop
      
      Printer.Print String(40, "-")

         Fonte 8, False, False
         Printer.Print Tab(0); "*** PARCELAS ***";
         Printer.Print Tab(0); "No.";
         If Grid.TextMatrix(Grid.Row, 7) = "À PRAZO" Then
             Printer.Print Tab(5); "VENC.";
         Else
             Printer.Print Tab(5); "PGTO";
         End If
         Printer.Print Tab(17); "VALOR";
         Printer.Print Tab(25); "FORMA"
      
          Do While Not rParc.EOF
             Printer.Print Tab(0); rParc("NUMERO");
             If Grid.TextMatrix(Grid.Row, 7) = "À Prazo" Then
                 Printer.Print Tab(5); Format$(Format$(rParc("DATA"), "dd/mm/yy"), "@@@@@@@");
             Else
                 Printer.Print Tab(5); Format$(Format$(rParc("PAGAMENTO"), "dd/mm/yy"), "@@@@@@@");
             End If
             Printer.Print Tab(15); Format$(Format$(rParc("VALOR_FINAL"), "0.00"), "@@@@@@@");
             Printer.Print Tab(25); rParc("varFormaPgto")
             
             rParc.MoveNext                 'vai para o proximo registro
         Loop
      
         For i = 1 To 1
            Printer.Print " "
         Next
      
      If Grid.TextMatrix(Grid.Row, 7) = "À Prazo" Then
         'sub-total
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "SubTotal: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("SUBTOTAL"), "0.00"), "@@@@@@@@")
         
         'desconto
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "Desc.: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("ValorAcrescReal"), "0.00"), "@@@@@@@@")
         
         'total
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "Total: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("TOTAL"), "0.00"), "@@@@@@@@")
         
      ElseIf Grid.TextMatrix(Grid.Row, 7) = "À Vista" Then
      
         'sub-total
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "SubTotal: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("SUBTOTAL"), "0.00"), "@@@@@@@@")
         
         'desconto
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "Desc.: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("ValorAcrescReal"), "0.00"), "@@@@@@@@")
         
         'total
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "Total: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("TOTAL"), "0.00"), "@@@@@@@@")
         
         Printer.Print
         
         'Recebido
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "Receb.: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(0, "0.00"), "@@@@@@@@")
         
         'Troco
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "Troco: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(0, "0.00"), "@@@@@@@@")
      ElseIf Grid.TextMatrix(Grid.Row, 1) = "ORÇAMENTO" Then
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "SubTotal: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("SUBTOTAL"), "0.00"), "@@@@@@@@")
         
         'desconto
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "Desc.: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("ValorAcrescReal"), "0.00"), "@@@@@@@@")
         
         'total
         Fonte 8, False, False
         Printer.Print Tab(0); Tab(20); "Total: ";
         
         Fonte 10, True, False
         Printer.Print Tab(25); Format$(Format$(rP("TOTAL"), "0.00"), "@@@@@@@@")
         
         Printer.Print
         
         'Recebido
         'Fonte 8, False, False
         'Printer.Print Tab(0); Tab(20); "Receb.: ";
         
         'Fonte 10, True, False
         'Printer.Print Tab(25); Format$(Format$(txtRecebido.Text, "0.00"), "@@@@@@@@")
         
         'Troco
         'Fonte 8, False, False
         'Printer.Print Tab(0); Tab(20); "Troco: ";
         
         'Fonte 10, True, False
         'Printer.Print Tab(25); Format$(Format$(txtTroco.Text, "0.00"), "@@@@@@@@")
      End If
      
      Printer.Print
      
      Fonte 8, False, False
      Printer.Print Tab((40 - Len("ESTE CUPOM NÃO TEM VALOR FISCAL")) / 2); "ESTE CUPOM NÃO TEM VALOR FISCAL"
      Fonte 8, False, False
      Printer.Print Tab((40 - Len("Obrigado pela preferência")) / 2); "Obrigado pela preferência"
      
      For i = 1 To 4
            Printer.Print " "
      Next
      
     'If Grid.TextMatrix(Grid.Row, 1) <> "ORÇAMENTO" Then
     '    Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
     '    Printer.Print Tab((40 - Len(Grid.TextMatrix(Grid.Row, 9))) / 2); Grid.TextMatrix(Grid.Row, 9)
     '    'If Not rParc.EOF Then
     '    'Printer.Print Tab((40 - Len("VENCIMENTO:" & rParc("DATA"))) / 2); rParc("DATA")
     '    'End If
     'End If
     
     
     
        'If Grid.TextMatrix(Grid.Row, 1) <> "ORÇAMENTO" Then
            Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
            Printer.Print Tab((40 - Len(Grid.TextMatrix(Grid.Row, 9))) / 2); Grid.TextMatrix(Grid.Row, 9)
            If Grid.TextMatrix(Grid.Row, 1) = "CONSIGNADO" And Grid.TextMatrix(Grid.Row, 1) = "ORÇAMENTO" Then
                Printer.Print Tab((40 - Len("Data:" & rP("DATA_COMPRA"))) / 2); "Data:" & Format(rP("DATA_COMPRA"), "dd/mm/yy")
            Else
                'If Grid.TextMatrix(Grid.Row, 7) = "À PRAZO" Then
                 '   Printer.Print Tab((40 - Len("VENCIMENTO:" & Grid.TextMatrix(Grid.Row, 9))) / 2); "Pagar em:" & Format(Grid.TextMatrix(Grid.Row, 9), "dd/mm/yy")
            '    Else
            '        Printer.Print Tab((40 - Len("VENCIMENTO:" & Grid.TextMatrix(Grid.Row, 9))) / 2); "Pago em:" & Format(Grid.TextMatrix(Grid.Row, 9), "dd/mm/yy")
            '    End If
            'Else
            '    Printer.Print Tab((40 - Len("Recebido em:" & Grid.TextMatrix(Grid.Row, 9))) / 2); "Recebido em:" & Format(Grid.TextMatrix(Grid.Row, 9), "dd/mm/yy")
            'End If
            End If

        For i = 1 To 2
            Printer.Print " "
        Next
        
        'DADOS DO RECECEDOR
        If Grid.TextMatrix(Grid.Row, 7) <> "ORÇAMENTO" And Grid.TextMatrix(Grid.Row, 7) <> "CONSIGNADO" Then
            If vDeclararRecebedor = "SIM" Then
                If Not rPR.EOF Then
                    Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
                    Printer.Print Tab((40 - Len(rPR("Recebedor"))) / 2); rPR("Recebedor")
                    Printer.Print Tab((40 - Len("RECEBEDOR")) / 2); "RECEBEDOR"
                    Printer.Print Tab((40 - Len("Recebido em:" & rP("DATA_COMPRA"))) / 2); "Recebido em:" & Format(rP("DATA_COMPRA"), "dd/mm/yy")
                End If
            End If
        End If
     
     
     
      'For i = 1 To 10
      'Print #f, ""
      'Next
     
   Close #f
   .EndDoc
   'rsPedidos.Close
   'rsFunc.Close
   'RS.Close
   'BD.Close
End With

Tratar_Erro:
' Atribui a impressora inicial
'For Each Prt In Printers
'   If Prt.DeviceName = oldPrinter Then
'      Set Printer = Prt
'      Exit For
'   End If
'Next

If Not r Is Nothing Then If r.State <> 0 Then r.Close
If Not rP Is Nothing Then If rP.State <> 0 Then rP.Close
If Not rPR Is Nothing Then If rPR.State <> 0 Then rPR.Close
If Not rI Is Nothing Then If rI.State <> 0 Then rI.Close
If Not rF Is Nothing Then If rF.State <> 0 Then rF.Close
If Not rParc Is Nothing Then If rParc.State <> 0 Then rParc.Close

'If Err.Number = 52 Then
 '  ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
 '  Printer.KillDoc
 '  Exit Sub
'End If
End Sub
Private Sub Fonte(Tamanho As Byte, Negrito As Boolean, Italico As Boolean) 'Altera a fonte
   Printer.FontSize = Tamanho
   Printer.FontBold = Negrito
   Printer.FontItalic = Italico
End Sub
Private Sub Imprimir_CupomSerrilha()
'   'On Error GoTo TrataErro
'   Dim sSQL As String
'   Dim r As ADODB.Recordset
'   Dim rP As ADODB.Recordset
'   Dim rI As ADODB.Recordset
'   Dim rF As ADODB.Recordset
   
'   Dim i As Integer
'   Dim f As Integer
   
'   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
'   Set r = dbData.OpenRecordset(sSQL)
   
   'consultar funcionario do pedido
'   Set rP = dbData.OpenRecordset("SELECT * FROM pedidos WHERE (cod_pedido = " & codPedido & ");")
'   Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rP("cod_funcionario") & ");")
   
'   f = FreeFile()
   
'   'Open "LPT1" For Output As #1
'   Open "\\BALCAO01\termica" For Output As #f
'      Print #f, Chr$(27) & Chr(15)
'      Print #f, Spc(0); "----------------------------------------------------------------"
 '     Print #f, Tab((60 - Len(r("fantasia"))) / 2); r("fantasia")
'      Print #f, Tab((60 - Len(r("razao"))) / 2); r("razao")
'      Print #f, Tab((60 - Len(r("endereco") & ", " & r("cidade") & "-" & r("estado"))) / 2); r("endereco") & ", " & r("cidade") & "-" & r("estado")
'      Print #f, Tab((60 - Len(r("telefone"))) / 2); r("telefone")
'      Print #f, Tab((60 - Len(r("cnpj") & "  IE:" & r("ie"))) / 2); r("cnpj") & "  IE:" & r("ie")
'      Print #f, ""
'      Print #f, Spc(0); Format(Date, "dd/mm/yy"); Spc(3); Format(Time, "hh:mm"); Spc(4); "No. Cupom:"; Spc(1); Format(codPedido, "000000"); Spc(3); "Usuario:"; Spc(1); rF("nome")
'      Print #f, ""
'      Print #f, Spc(0); "                       C   U   P   O   M                     "
'      Print #f, Spc(0); "----------------------------------------------------------------"
'      Print #f, Tab(0); "DESCRICAO"; Tab(40); "PRECO"; Tab(48); "QUANT"; Tab(56); "TOTAL"
'      Print #f, Spc(0); "----------------------------------------------------------------"

'         sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.cod_pedido, pedidos_itens.preco, pedidos_itens.quantidade, (pedidos_itens.preco * pedidos_itens.quantidade) as total, produtos.descricao " & _
'            "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
'            "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") ORDER BY pedidos_itens.codigo DESC;"
'         Set rI = dbData.OpenRecordset(sSQL)
      
'      Do While Not rI.EOF
'         Print #f, Tab(0); rI("descricao"); Tab(38); Format$(Format$(r("preco"), "0.00"), "@@@@@@@"); Tab(46); Format$(Format$(rI("quantidade"), "0.000"), "@@@@@@@"); Tab(54); Format$(Format$(rI("total"), "0.00"), "@@@@@@@")
'         rI.MoveNext
'      Loop
      
'      Print #f, Spc(0); "----------------------------------------------------------------"
'      Print #f, Tab(45); "TOTAL: "; Tab(54); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
'      Print #f, ""
'      Print #f, Tab((60 - Len("ESTE CUPOM NAO TEM VALOR FISCAL")) / 2); "ESTE CUPOM NAO TEM VALOR FISCAL"
'      Print #f, Tab((60 - Len("Obrigado pela preferencia")) / 2); "Obrigado pela preferencia"
'      Print #f, ""
'      Print #f, ""
'      Print #f, ""
'      Print #f, ""
'      Print #f, ""
'      Print #f, ""
'      Print #f, ""
'      Print #f, ""
'   Close #f
   
'   If Not r Is Nothing Then If r.State <> 0 Then r.Close
'   If Not rP Is Nothing Then If rP.State <> 0 Then rP.Close
'   If Not rI Is Nothing Then If rI.State <> 0 Then rI.Close
'   If Not rF Is Nothing Then If rF.State <> 0 Then rF.Close
   

''TrataErro:
'   'MsgBox Err.Description, vbCritical, "Erro no Sistema, Impressora Inoperante"
End Sub

Private Sub cmdReaberturas_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Grid.TextMatrix(Grid.Row, 14) <> "SIM" Then
    MsgBox "Não existe um histórico de reabertura para esse pedido!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

codPedido = Grid.TextMatrix(Grid.Row, 2)

Estorno_ReabrirPedidos.loadInformacoes codPedido
Estorno_ReabrirPedidos.Show 1
End Sub

Private Sub cmdReimprimir_Click()
  
   Me.Hide
   PDV.lblEstornar.Caption = "REIMPRESSÃO"
   
   If cboCriterios.Text = "Cod. Pedido" Then
      If txtCodPedido.Text = "" Then Exit Sub
      PDV.txtCodPedido.Text = txtCodPedidoCerto.Text
   Else
      If txtCodPedido.Text = "NENHUM PEDIDO" Then
         ShowMsg "ESSE PEDIDO NÃO PODE SER REIMPRESSO", vbExclamation
         Exit Sub
      Else
         PDV.txtCodPedido.Text = txtCodPedidoCerto.Text
      End If
   End If
   
   PDV.Show
End Sub

Private Sub FormatarGrid_Pedido(rTabela As ADODB.Recordset)
Dim i As Integer
Dim j As Integer

   With Grid
      .Clear
      .Cols = 20
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 1050
      .ColWidth(2) = 630
      .ColWidth(3) = 0
      .ColWidth(4) = 750
      .ColWidth(5) = 0
      .ColWidth(6) = 270
      .ColWidth(7) = 750
      .ColWidth(8) = 1450
      .ColWidth(9) = 3000
      .ColWidth(10) = 900
      .ColWidth(11) = 900
      .ColWidth(12) = 900
      .ColWidth(13) = 900
      .ColWidth(14) = 800
      .ColWidth(15) = 800
      .ColWidth(16) = 800
      .ColWidth(17) = 0
      .ColWidth(18) = 800
      .ColWidth(19) = 800
      
      .TextMatrix(0, 1) = "TIPO"
      .TextMatrix(0, 2) = "PEDIDO"
      .TextMatrix(0, 3) = "STATUS"
      .TextMatrix(0, 4) = "EMISSÃO"
      .TextMatrix(0, 5) = "TIPO"
      .TextMatrix(0, 6) = "V"
      .TextMatrix(0, 7) = "FORMA"
      .TextMatrix(0, 8) = "TIPO"
      .TextMatrix(0, 9) = "CLIENTE"
      .TextMatrix(0, 10) = "SUBTOT."
      .TextMatrix(0, 11) = "DESC."
      .TextMatrix(0, 12) = "ACRESC."
      .TextMatrix(0, 13) = "VALOR"
      .TextMatrix(0, 14) = "Reaberto"
      .TextMatrix(0, 15) = "Cancel."
      .TextMatrix(0, 16) = "NFCe"
      .TextMatrix(0, 17) = "INUT"
      .TextMatrix(0, 18) = "CAIXA"
      .TextMatrix(0, 19) = "CÓD/CX"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      '.ColAlignment(2) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next i

      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("var_TipoPedido")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_CodPedido")
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("Var_StatusPEDIDO"))
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("var_Data"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("var_TipoPedido"))
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("varCod_Func"))
            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_TipoPagamento"))
            If chkIncompleto.Value = Unchecked Then
                .TextMatrix(.Rows - 1, 8) = ValidateNull(rTabela("var_Pagamento"))
            End If
        If cboStatus.Text <> "VAZIO" Then
            .TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("var_Cliente"))
        End If
            .TextMatrix(.Rows - 1, 10) = Format(rTabela("var_SUBTOTAL"), ocMONEY)
            .TextMatrix(.Rows - 1, 11) = Format(rTabela("var_DESC"), ocMONEY)
            .TextMatrix(.Rows - 1, 12) = Format(rTabela("var_ACRESC"), ocMONEY)
            .TextMatrix(.Rows - 1, 13) = Format(rTabela("var_Total"), ocMONEY)
            .TextMatrix(.Rows - 1, 14) = rTabela("Var_StatusREABERTO")
            .TextMatrix(.Rows - 1, 15) = rTabela("Var_StatusCANCELADO")
            .TextMatrix(.Rows - 1, 16) = ValidateNull(rTabela("Var_StatusNFCE"))
'            .TextMatrix(.Rows - 1, 17) = rTabela("Var_NFCEInutilizada")
            .TextMatrix(.Rows - 1, 18) = ValidateNull(rTabela("VarPEDCAIXA"))
            .TextMatrix(.Rows - 1, 19) = ValidateNull(rTabela("VarPEDCODCAIXA"))
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
   FlexCores &HFFFFFF, &HE0E0E0

      .Rows = .Rows - 1
   End With
End Sub
Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
Dim i As Integer, Valor As Currency

Valor = 0
For i = 0 To var_Grid.Rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CCur(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function
Private Sub Mostrar_Pedido()
Dim varCriterio As String
Dim varIndice As String
Dim varStatus As String

If cboStatus.Text = "" Then Exit Sub
lblTotalGrid.Caption = Format(0, ocMONEY)

'======================================================= CRITERIOS WHERE
If cboCriterios.Text = "NENHUM" Then
    varCriterio = " "
ElseIf cboCriterios.Text = "CÓD. PEDIDO" Then
    If txtCodPedidoCerto.Text = "" Then
        varCriterio = " and pedidos.cod_pedido = 0 "
    Else
        varCriterio = " and pedidos.cod_pedido = " & txtCodPedidoCerto.Text & " "
    End If
ElseIf cboCriterios.Text = "CÓD. BARRA" Then
    If txtCodProdutoBarra.Text = "0" Then
        varCriterio = " and (pedidos_itens.cod_produto < '0') "
    Else
        varCriterio = " and (pedidos_itens.cod_produto = " & txtCodProdutoBarra.Text & ") "
    End If
ElseIf cboCriterios.Text = "CLIENTE" Then
    If txtCodCliente.Text = "" Then
        varCriterio = " and cliente.codigo = 0"
    Else
        varCriterio = " and cliente.codigo = " & txtCodCliente.Text & ""
    End If
ElseIf cboCriterios.Text = "DATA" Then
    If IsDate(mskData) = True Then
        varCriterio = " and (pedidos.data_compra = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103))"
    Else
        varCriterio = " and pedidos.cod_pedido = 0"
    End If
ElseIf cboCriterios.Text = "MENSAL" Then
    If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
    varCriterio = " and (MONTH(DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(DATA_COMPRA) = " & cboAno & ")"
    'varCriterio = " and pedidos.cod_pedido = " & txtCodPedidoCerto.Text & ""
ElseIf cboCriterios.Text = "PRODUTO" Then
    If txtCodProduto.Text = "" Then Exit Sub
    varCriterio = " and (pedidos_itens.cod_produto = " & txtCodProduto.Text & ")"
End If

'====================================================== TIPO DE PEDIDO
Dim vTipoPedido As String
If cboTipoPedido.Text = "TODOS" Then
    vTipoPedido = " "
ElseIf cboTipoPedido.Text = "VENDA" Then
    vTipoPedido = " and pedidos.TIPO_PEDIDO = 'VENDA'"
ElseIf cboTipoPedido.Text = "ORÇAMENTO" Then
    vTipoPedido = " and pedidos.TIPO_PEDIDO = 'ORÇAMENTO'"
ElseIf cboTipoPedido.Text = "CONSIGNADO" Then
    vTipoPedido = " and pedidos.CANCELADO = 0 and pedidos.TIPO_PEDIDO = 'CONSIGNADO'"
ElseIf cboTipoPedido.Text = "ALUGUEL" Then
    vTipoPedido = " and pedidos.TIPO_PEDIDO = 'ALUGUEL'"
ElseIf cboTipoPedido.Text = "OFICINA" Then
    vTipoPedido = " and pedidos.TIPO_PEDIDO = 'OFICINA'"
ElseIf cboTipoPedido.Text = "CANCELADO" Then
    vTipoPedido = " and pedidos.CANCELADO = 1 and pedidos.TIPO_PEDIDO IN ('VENDA','CONSIGNADO') "
End If
'WHERE pedidos.TIPO_PEDIDO IN ('VENDA','CONSIGNADO');
'====================================================== FORMA DE PGTO
Dim varFormaPgto As String
If cboFormaPgto.Text = "TODOS" Then
    varFormaPgto = " "
ElseIf cboFormaPgto.Text = "À VISTA" Then
    varFormaPgto = " and pedidos.TIPO_PAGAMENTO = 'À Vista'"
ElseIf cboFormaPgto.Text = "À PRAZO" Then
    varFormaPgto = " and pedidos.TIPO_PAGAMENTO = 'À Prazo' and pedidos.TIPO_PEDIDO <> 'ORÇAMENTO'"
End If

'====================================================== TIPO DE PGTO
Dim varTipoPgto As String

If cboTipoPgto.Text = "TODOS" Then
    varTipoPgto = "  "
ElseIf cboTipoPgto.Text = "DINHEIRO" Then
    varTipoPgto = " and parcelas.FORMA_PGTO = 'DINHEIRO'"
ElseIf cboTipoPgto.Text = "PROMISSÓRIA" Then
    varTipoPgto = " and parcelas.FORMA_PGTO = 'PROMISSORIA'"
ElseIf cboTipoPgto.Text = "CARTÃO" Then
    varTipoPgto = " and parcelas.FORMA_PGTO = 'CARTAO'"
ElseIf cboTipoPgto.Text = "CHEQUE" Then
    varTipoPgto = " and parcelas.FORMA_PGTO = 'CHEQUE'"
ElseIf cboTipoPgto.Text = "BOLETO" Then
    varTipoPgto = " and parcelas.FORMA_PGTO = 'BOLETO'"
ElseIf cboTipoPgto.Text = "TRANSFERÊNCIA" Then
    varTipoPgto = " and parcelas.FORMA_PGTO = 'TRANSFERENCIA'"
ElseIf cboTipoPgto.Text = "PIX" Then
    varTipoPgto = " and parcelas.FORMA_PGTO = 'PIX'"
End If

'====================================================== INDICE
If cboIndice.Text = "CÓD. PEDIDO" Then
    varIndice = " pedidos.cod_pedido "
ElseIf cboIndice.Text = "CLIENTE" Then
    If cboStatus.Text = "VAZIO" Then
        varIndice = " pedidos.cod_pedido "
    Else
        varIndice = " cliente.codigo "
    End If
ElseIf cboIndice.Text = "EMISSÃO" Then
    varIndice = " var_Data "
ElseIf cboIndice.Text = "TIPO PGTO" Then
    varIndice = " pedidos.TIPO_PAGAMENTO "
Else
    varIndice = " pedidos.cod_pedido "
End If

'============================================= EXIBIR CANCELADOS
'Dim vExibirCancelados As String
'If chkMostarCancelados.Value = Unchecked Then
'    vExibirCancelados = " and pedidos.CANCELADO = 0"
'Else
'    vExibirCancelados = " and pedidos.CANCELADO = 1"
'End If

'============================================= CONSULTA
'If cboTipoPedido = "TODOS" Or cboTipoPedido = "VENDA" Or cboTipoPedido = "CONSIGNADO" Or cboTipoPedido = "ALUGUEL" Or cboTipoPedido = "OFICINA" Then


If cboStatus.Text = "TODOS" Then
    If chkIncompleto.Value = Unchecked Then
        varStatus = " pedidos.status_pedido <> 3 "
    Else
        varStatus = " pedidos.status_pedido = 0 "
    End If
    'SUBSTRING((SELECT ', ' + P.FORMA_PGTO FROM dbo.parcelas P WHERE P.COD_PEDIDO = C.COD_PEDIDO FOR XML PATH ('')), 2, 1000) var_Pagamento
     'sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, cliente.nome as var_Cliente, cliente.codigo, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, (CASE WHEN TbNFCe.NFCeEnviada = 1 THEN 'SIM' ELSE '' END) AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN 'SIM' ELSE '' END) AS Var_NFCEInutilizada, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa " & _
     "FROM  pedidos INNER JOIN Cliente ON pedidos.COD_CLIENTE = Cliente.CODIGO LEFT OUTER JOIN TbNFCe ON TbNFCe.Num_OS_VD_Origem = pedidos.COD_PEDIDO " & _
    "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & " " & vExibirCancelados & " AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
    'If cboTipoPedido.Text <> "ORÇAMENTO" Then
    If cboTipoPedido = "TODOS" Or cboTipoPedido = "VENDA" Or cboTipoPedido = "ALUGUEL" Or cboTipoPedido = "OFICINA" Then
        If chkIncompleto.Value = Unchecked Then
            If cboCriterios.Text <> "CLIENTE" Then
                'If cboTipoPedido.Text <> "CANCELADO" Then
                sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                    "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                    "SUBSTRING((SELECT ', ' + P.FORMA_PGTO FROM dbo.parcelas P WHERE P.COD_PEDIDO = pedidos.COD_PEDIDO FOR XML PATH ('')), 2, 1000) var_Pagamento,  " & _
                    "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE " & _
                    "FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO " & _
                    "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
            ElseIf cboCriterios.Text = "CLIENTE" Then
                sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                    "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                    "SUBSTRING((SELECT ', ' + P.FORMA_PGTO FROM dbo.parcelas P WHERE P.COD_PEDIDO = pedidos.COD_PEDIDO FOR XML PATH ('')), 2, 1000) var_Pagamento,  " & _
                    "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE " & _
                    "FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
                    "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
            End If
        Else
            If cboCriterios.Text <> "CLIENTE" Then
                'If cboTipoPedido.Text <> "CANCELADO" Then
                sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                    "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                    "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE " & _
                    "FROM pedidos " & _
                    "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
            ElseIf cboCriterios.Text = "CLIENTE" Then
                sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                    "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                    "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE " & _
                    "FROM pedidos " & _
                    "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
            End If
        End If
        
    ElseIf cboTipoPedido.Text = "CANCELADO" Then
        'If cboCriterios.Text <> "CLIENTE" Then
            sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                "SUBSTRING((SELECT ', ' + P.FORMA_PGTO FROM dbo.parcelas P WHERE P.COD_PEDIDO = pedidos.COD_PEDIDO FOR XML PATH ('')), 2, 1000) var_Pagamento,  " & _
                "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE " & _
                "FROM pedidos " & _
                "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
        'ElseIf cboCriterios.Text = "CLIENTE" Then
        '    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
        '        "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
        '        "SUBSTRING((SELECT ', ' + P.FORMA_PGTO FROM dbo.parcelas P WHERE P.COD_PEDIDO = pedidos.COD_PEDIDO FOR XML PATH ('')), 2, 1000) var_Pagamento,  " & _
        '        "(SELECT (CASE WHEN N .NFCeEnviada = 1 THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)) AS Var_StatusNFCE " & _
        '        "FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
        '        "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
        'End If
    ElseIf cboTipoPedido.Text = "CONSIGNADO" Then
    
    
        If cboCriterios.Text <> "CLIENTE" Then
            sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, '' var_Pagamento " & _
                "FROM pedidos " & _
                "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL' )"
        ElseIf cboCriterios.Text = "CLIENTE" Then
            sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, '' var_Pagamento " & _
                "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
                "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL' )"
        End If
    
    
        
    ElseIf cboTipoPedido.Text = "ORÇAMENTO" Then
        If cboCriterios.Text <> "CLIENTE" Then
            sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, '' var_Pagamento " & _
                "FROM pedidos " & _
                "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
        ElseIf cboCriterios.Text = "CLIENTE" Then
            sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _
                "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _
                "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, '' var_Pagamento " & _
                "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
                "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
        End If
    End If

ElseIf cboStatus.Text = "FECHADO" Then
    varStatus = " pedidos.status_pedido = 1 "
    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, cliente.nome as var_Cliente, cliente.codigo, pedidos.DATA_COMPRA as var_Data,pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, " & _
        "(CASE WHEN TbNFCe.Inutilizada = 1 THEN 'SIM' ELSE '' END) AS Var_NFCEInutilizada, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa " & _
        "FROM  pedidos INNER JOIN Cliente ON pedidos.COD_CLIENTE = Cliente.CODIGO LEFT OUTER JOIN TbNFCe ON TbNFCe.Num_OS_VD_Origem = pedidos.COD_PEDIDO " & _
        "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & " AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
ElseIf cboStatus.Text = "ABERTO" Then
    varStatus = " pedidos.status_pedido = 0 AND (not(COD_CLIENTE IS NULL)) "
    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, cliente.nome as var_Cliente, cliente.codigo, pedidos.DATA_COMPRA as var_Data,pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO , ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, " & _
       "(CASE WHEN TbNFCe.Inutilizada = 1 THEN 'SIM' ELSE '' END) AS Var_NFCEInutilizada, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa " & _
        "FROM  pedidos INNER JOIN Cliente ON pedidos.COD_CLIENTE = Cliente.CODIGO LEFT OUTER JOIN TbNFCe ON TbNFCe.Num_OS_VD_Origem = pedidos.COD_PEDIDO " & _
       "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & " AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
ElseIf cboStatus.Text = "PAUSADO" Then
    varStatus = " pedidos.status_pedido = -1 "
    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, cliente.nome as var_Cliente, cliente.codigo, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, " & _
        " (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, (CASE WHEN TbNFCe.NFCeEnviada = 1 THEN 'SIM' ELSE '' END) AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN 'SIM' ELSE '' END) AS Var_NFCEInutilizada, pedidos.caixa as varPedCaixa, " & _
        "pedidos.codcaixa as varPedCodCaixa, pedidos.TIPO_PEDIDO AS var_TIPOPedido " & _
        "FROM  pedidos INNER JOIN Cliente ON pedidos.COD_CLIENTE = Cliente.CODIGO LEFT OUTER JOIN TbNFCe ON TbNFCe.Num_OS_VD_Origem = pedidos.COD_PEDIDO " & _
       "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & " AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
ElseIf cboStatus.Text = "VAZIO" Then
    varStatus = " (STATUS_PEDIDO = 0) AND (COD_CLIENTE IS NULL)"
    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, cliente.codigo, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN 'SIM' ELSE '' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN 'SIM' ELSE '' END) AS Var_StatusCANCELADO, ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN 'SIM' ELSE '' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), '') AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN 'SIM' ELSE '' END) AS Var_NFCEInutilizada, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa " & _
        "FROM  pedidos INNER JOIN Cliente ON pedidos.COD_CLIENTE = Cliente.CODIGO LEFT OUTER JOIN TbNFCe ON TbNFCe.Num_OS_VD_Origem = pedidos.COD_PEDIDO " & _
       "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & " " & vTipoPedido & " AND (pedidos.TIPO_PEDIDO <> 'ALUGUEL')"
End If

sSQL = sSQL & "" & varCriterio & " ORDER BY " & varIndice

'Debug.Print sSQL

Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    cmdImprimir.Enabled = True
End If

FormatarGrid_Pedido r

Dim soma As Currency
Dim contar As Integer
Dim i As Integer

'Somar as vendas
soma = 0
contar = 0
With Grid
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 1) = "VENDA" Then
        If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 13))
        End If
      End If
   Next
End With

lblTotalGrid.Caption = Format(soma, "#,##0.00")
lblTotalVendas.Caption = Format(contar, "000")


'Somar as orçamento
soma = 0
contar = 0
With Grid
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 1) = "ORÇAMENTO" Then
        If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 13))
        End If
      End If
   Next
End With

lblTotalGridORC.Caption = Format(soma, "#,##0.00")
lblQuantOrc.Caption = Format(contar, "000")

'Somar as consignado
soma = 0
contar = 0
With Grid
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 1) = "CONSIGNADO" Then
        If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 13))
        End If
      End If
   Next
End With

lblTotalGridConsignado.Caption = Format(soma, "#,##0.00")
lblQuantConsignado.Caption = Format(contar, "000")

'canceladas
soma = 0
contar = 0
With Grid
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 15) = "SIM" Then
         contar = contar + 1
         soma = soma + CCur(.TextMatrix(i, 13))
      End If
   Next
End With

lblTotalCanc.Caption = Format(soma, "#,##0.00")
lblQuantCanc.Caption = Format(contar, "000")

'lblTotalGridORC.Caption = " Orçamentos: " & Format(soma, "#,##0.00")

'lblQuant.Caption = " Quant.: " & r.RecordCount
'lblTotalGrid.Caption = Format(SomaGrid(Grid, 13), ocMONEY)

printSQL = sSQL

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub



Private Sub Form_Activate()
If frmSenha.Visible = True Then
    cmdPedidoAbrir.Enabled = False
    cmdExcluirPedido.Enabled = False
End If
End Sub

Public Function LerPermissoesUsuario(vCodUser As Long, permissao As Long) As Boolean
sSQL = "SELECT Usuario_Acessos.Cod_Permissao FROM Usuario INNER JOIN Usuario_Acessos ON Usuario.Codigo = Usuario_Acessos.Cod_Usuario WHERE (Usuario_Acessos.Cod_Usuario = " & vCodUser & ") AND Usuario_Acessos.Cod_Permissao = " & permissao & ";"
Set r = dbData.OpenRecordset(sSQL)

If r.EOF And r.BOF Then
   LerPermissoesUsuario = False ' não achou a permissao
Else
   LerPermissoesUsuario = True 'aqui achou
End If
End Function
Private Sub Form_Load()
Set moCombo = New cComboHelper



CAIXA_FECHADO = True
txtCodPedidoCerto.Text = ""
txtCodCliente.Text = ""
txtCodPedido.Text = ""
cboCliente.Text = ""
optDig.Value = True
PreencherCriterios
cboCriterios.ListIndex = 3
PreencherIndice
cboIndice.ListIndex = 0
PreencherStatus
cboStatus.ListIndex = 1
PreencherFormaPgto
cboFormaPgto.ListIndex = 0
PreencherTipoPgto
cboTipoPgto.ListIndex = 0
PreencherTipoPedido
cboTipoPedido.ListIndex = 0

lblData.Visible = True
cmdCal1.Visible = True
mskData.Visible = True
cmdModificar.Visible = True
cmdModificarConsignado.Visible = False
mskData.Text = Format(Date, "dd/mm/yyyy")
StatusBar1.Panels(4).Text = Format(Date, "dd/mm/yyyy")
LimparGridPedidos

'tipo de empresa
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

'se precisa pedi senha nas opções do menu avançado
Set cCfg = sysConfig("SEGURANCAAVANCADA")
varSegurancaAvancada = cCfg.Value
Set cCfg = Nothing

'If varSegurancaAvancada = "SIM" Then frmSenha.Visible = True Else frmSenha.Visible = False

'If tipoEmpresa = 4 Then
'    Label11.Caption = "Consignado"
'Else
'    Label11.Caption = "Orçamento"
'End If

'se precisa pedi senha nas opções do menu avançado
Set cCfg = sysConfig("DECLARARRECEBEDOR")
vDeclararRecebedor = cCfg.Value
Set cCfg = Nothing


Set oCfg = sysConfig("TIPOIMPRESSAOPARCELAS")
vTipoParcelaImpressao = CInt(oCfg.Value)
Set oCfg = Nothing

'ver as permissões do cliente e liberar os botãos
'LiberarBotoesPermissoes

'fazer um loop que pega todas as permissões
'Dim y As Integer
'For y = 1 To 12 'seja lá quantas permissões vc tiver, aqui são 20
'    Dim vCodUsuario As Long
'    vCodUsuario = 1

'    Dim Ctrl As Control
'    For Each Ctrl In Me.Controls
'       If (TypeOf Ctrl Is chameleonButton) Then
'           If LerPermissoesUsuario(vCodUsuario, y) = True Then
'                Ctrl.Enabled = True
'           End If
'       End If
'    Next
'Next y
End Sub
Private Sub cmdSenha_Click()
If txtSenha.Text = "" Then ShowMsg "ACESSO NEGADO!" & vbCrLf & "Senha obrigatória", vbInformation: Exit Sub
If txtCodUsuario.Text = "" Then ShowMsg "ACESSO NEGADO!" & vbCrLf & "Usuário obrigatório", vbInformation: Exit Sub

sSQL = "SELECT codigo, password, nivel, login FROM Usuario WHERE (password = '" & txtSenha.Text & "') AND (codigo = " & txtCodUsuario.Text & ");"
Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then
        lblCodUser1.Visible = True
        lblCodUser2.Visible = True
        lblUser1.Visible = True
        lblUser2.Visible = True
        txtSenha.Text = ""
        frmSenha.Visible = False
        lblCodUser2.Caption = Format(r("codigo"))
        lblUser2.Caption = r("login")
   Else
        ShowMsg "ACESSO NEGADO!" & vbCrLf & "Você não tem nivel de acesso a esse recurso", vbInformation
        txtSenha.Text = ""
        'frmSenha.Visible = False
        lblCodUser2.Caption = ""
        lblUser2.Caption = ""
        vCodUsuario = 0
  End If
LiberarBotoesPermissoes
cmdModificar.Enabled = False
cmdModificarConsignado.Enabled = False
End Sub


Private Sub Grid_Click()
Dim vQuantLinhas As Integer

vQuantLinhas = Grid.Rows - 1

If vQuantLinhas >= 1 Then
    If Grid.TextMatrix(Grid.Row, 1) = "ORÇAMENTO" Then cmdModificar.Enabled = True: cmdModificar.Visible = True: cmdModificarConsignado.Visible = False: cmdModificarConsignado.Enabled = False: cmdPedidoAbrir.Caption = "CONVERTER": cmdModificar.Caption = "EDITAR" Else cmdModificar.Enabled = False: cmdPedidoAbrir.Caption = "REABRIR"
    
    If Grid.TextMatrix(Grid.Row, 1) <> "ORÇAMENTO" Then
        If Grid.TextMatrix(Grid.Row, 1) = "CONSIGNADO" Then cmdModificarConsignado.Enabled = True: cmdModificarConsignado.Visible = True: cmdModificar.Visible = False: cmdModificar.Enabled = False: cmdPedidoAbrir.Caption = "CONVERTER": cmdModificarConsignado.Caption = "EDITAR" Else cmdModificarConsignado.Enabled = False: cmdPedidoAbrir.Caption = "REABRIR"
    End If
    
    If Grid.TextMatrix(Grid.Row, 14) = "SIM" Then cmdReaberturas.Enabled = True Else cmdReaberturas.Enabled = False
    cmdPedidoImprimir.Enabled = True
    cmdPDF.Enabled = True
    cmdMostrarProdutos.Enabled = True
    
    'permissões
    LiberarBotoesPermissoes
End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdSenha_Click
End Sub
Private Sub cboUsuario_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
cboUsuario.Clear
sSQL = "SELECT codigo, login FROM usuario WHERE (visivel = 1) ORDER BY login;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboUsuario.AddItem r("login")
   cboUsuario.ItemData(cboUsuario.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboUsuario
End Sub
Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cboUsuario_LostFocus()
On Error GoTo TrataErro

If cboUsuario.Text = "" Then txtCodUsuario.Text = "": txtNivelUsuario.Text = "": Exit Sub
If cboUsuario.ListIndex = -1 Then txtCodUsuario.Text = "": txtNivelUsuario.Text = "": Exit Sub
txtCodUsuario = cboUsuario.ItemData(cboUsuario.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub txtCodUsuario_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodUsuario.Text = "" Then Exit Sub

sSQL = "SELECT codigo, nivel FROM usuario WHERE (codigo = " & txtCodUsuario.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    vCodUsuario = r("codigo")
Else
    vCodUsuario = 0
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Sub FlexCores(lCorPar As Long, lCorImpar As Long)
   'ZEBRAR O FLEXGRID
   Dim iLinha As Integer
   Dim lCor As OLE_COLOR
   
   Grid.FillStyle = flexFillRepeat
   
   For iLinha = 1 To Grid.Rows - 1
      With Grid
         .Row = iLinha
         
         If EImpar(iLinha) Then 'Se a linha for impar:
            lCor = lCorImpar
         Else
            lCor = lCorPar
         End If
         
         .Col = 1                'Seleciona a partir da primeira coluna
         .ColSel = .Cols - 1     'Seleciona até a última coluna
         .CellBackColor = lCor   'Aplica a cor
      End With
   Next
   
   Grid.FillStyle = flexFillSingle
End Sub

Function EImpar(ByVal iNum As Long) As Boolean
   EImpar = (iNum Mod 2)
End Function
Private Sub Form_Unload(Cancel As Integer)
HabilitaObjetosVenda False
Set moCombo = Nothing
End Sub


Private Sub mskData_GotFocus()
SelectControl mskData
End Sub


Private Sub mskData_KeyPress(KeyAscii As Integer)
mskData.Mask = "##/##/##"
End Sub


Private Sub mskData_LostFocus()
If Not IsDate(mskData.Text) Then
   mskData.Mask = ""
   mskData.Text = ""
End If
End Sub


Private Sub optDig_Click()
txtCodPedido.Clear
If txtCodPedido.Text = "" Then txtCodPedidoCerto.Text = "": Exit Sub
End Sub

Private Sub optEsc_Click()
txtCodPedido.Clear
End Sub

Private Sub txtCodBarra_Validate(Cancel As Boolean)
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodBarra.Text = "" Then txtCodProdutoBarra.Text = "0": Exit Sub

sSQL = "SELECT COD_BARRA, CODIGO FROM produtos WHERE (produtos.cod_barra = '" & txtCodBarra.Text & "') AND (produtos.ativo = 1);"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    txtCodProdutoBarra.Text = ValidateNull(r("CODIGO"))
Else
    'MsgBox "PRODUTO NÃO CADASTRADO!", vbCritical, "Alerta"
    txtCodProdutoBarra.Text = "0"
    'Exit Sub
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub


Private Sub txtCodPedido_Change()
If optDig.Value = True Then
    If txtCodPedido.Text <> "" Then
        txtCodPedidoCerto.Text = txtCodPedido.Text
    Else
        txtCodPedidoCerto.Text = ""
    End If
Else
    If txtCodPedido.Text <> "" Then
    txtCodPedidoCerto.Text = Mid(txtCodPedido.Text, 1, InStr(1, txtCodPedido.Text, "->", vbTextCompare) - 1)
    End If
End If
End Sub

Private Sub txtCodPedido_Click()
If optDig.Value = True Then
    If txtCodPedido.Text <> "" Then
    txtCodPedidoCerto.Text = txtCodPedido.Text
    End If
Else
    If txtCodPedido.Text <> "" Then
    txtCodPedidoCerto.Text = Mid(txtCodPedido.Text, 1, InStr(1, txtCodPedido.Text, "->", vbTextCompare) - 1)
    End If
End If
txtCodPedido_LostFocus
End Sub

Private Sub txtCodPedido_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   txtCodPedido.Clear
   
   sSQL = "SELECT top 50 * FROM pedidos ORDER BY cod_pedido DESC;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.BOF Then
      txtCodPedido.AddItem "NENHUM PEDIDO"
   Else
      Do While Not r.EOF

        If optDig.Value = True Then
            txtCodPedido.Clear
        Else
            txtCodPedido.AddItem Format(r("cod_pedido"), "000000") & " -> " & Format(r("data_compra"), "dd/mm/yyyy") & " -> " & Format(r("total"), ocMONEY)
        End If
         
         r.MoveNext
      Loop
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

If optEsc.Value = True Then
   txtCodPedido.ListIndex = 0
End If
End Sub


Private Sub txtCodPedido_LostFocus()
If txtCodPedido.Text = "" Then Exit Sub
'Mostrar_Pedido
End Sub

Private Sub txtCodPedido_Validate(Cancel As Boolean)
If txtCodPedido.Text = "" Then
    If optEsc.Value = True Then
    txtCodPedidoCerto.Text = Mid(txtCodPedido.Text, 1, InStr(1, txtCodPedido.Text, "->", vbTextCompare) - 1)
    End If
End If
End Sub



