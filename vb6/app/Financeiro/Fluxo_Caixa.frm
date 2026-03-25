VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Fluxo_Caixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FLUXO DE CAIXA"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "Fluxo_Caixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodFuncionario 
      Height          =   285
      Left            =   6060
      TabIndex        =   38
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmSenha 
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
      Left            =   60
      TabIndex        =   33
      Top             =   8400
      Visible         =   0   'False
      Width           =   2355
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin ChamaleonBtn.chameleonButton cmdSenha 
         Height          =   315
         Left            =   1500
         TabIndex        =   35
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
         MICON           =   "Fluxo_Caixa.frx":23D2
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
         Height          =   315
         Left            =   1920
         TabIndex        =   37
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "X"
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
         MICON           =   "Fluxo_Caixa.frx":23EE
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   60
      ScaleHeight     =   975
      ScaleWidth      =   11685
      TabIndex        =   0
      Top             =   1080
      Width           =   11715
      Begin VB.Frame frmCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Escolha os critérios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   4935
         Begin VB.ComboBox cboOrganizar 
            Height          =   315
            Left            =   1620
            TabIndex        =   23
            Top             =   480
            Width           =   1515
         End
         Begin VB.ComboBox cboConsulta 
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   480
            Width           =   1515
         End
         Begin VB.ComboBox cboCaixa 
            Height          =   315
            Left            =   3180
            TabIndex        =   17
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Organizar"
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
            TabIndex        =   24
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consultar por"
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
            TabIndex        =   22
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caixa"
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
            Left            =   3180
            TabIndex        =   18
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame frmObjetos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   5160
         TabIndex        =   11
         Top             =   60
         Width           =   3135
         Begin ChamaleonBtn.chameleonButton cmdConsData2 
            Height          =   315
            Left            =   2700
            TabIndex        =   39
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
            MICON           =   "Fluxo_Caixa.frx":240A
            PICN            =   "Fluxo_Caixa.frx":2426
            PICH            =   "Fluxo_Caixa.frx":4779
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   1635
         End
         Begin ChamaleonBtn.chameleonButton cmdConsData 
            Height          =   315
            Left            =   1020
            TabIndex        =   25
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
            MICON           =   "Fluxo_Caixa.frx":6ACC
            PICN            =   "Fluxo_Caixa.frx":6AE8
            PICH            =   "Fluxo_Caixa.frx":8E3B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskConsData 
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskConsData2 
            Height          =   315
            Left            =   1800
            TabIndex        =   27
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label lblAte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Até"
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
            Left            =   1380
            TabIndex        =   40
            Top             =   540
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano"
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
            Left            =   1800
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês"
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
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   615
         Left            =   8400
         TabIndex        =   19
         Top             =   180
         Width           =   1575
         _ExtentX        =   2778
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
         MICON           =   "Fluxo_Caixa.frx":B18E
         PICN            =   "Fluxo_Caixa.frx":B1AA
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
         Height          =   615
         Left            =   10020
         TabIndex        =   20
         Top             =   180
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
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
         MICON           =   "Fluxo_Caixa.frx":BA84
         PICN            =   "Fluxo_Caixa.frx":BAA0
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   9240
      ScaleHeight     =   1005
      ScaleWidth      =   2505
      TabIndex        =   1
      Top             =   7980
      Width           =   2535
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Totais"
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
         Left            =   360
         TabIndex        =   7
         Top             =   660
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saídas"
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
         Left            =   300
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entradas"
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
         TabIndex        =   5
         Top             =   60
         Width           =   765
      End
      Begin VB.Label lblTotalSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   1020
         TabIndex        =   4
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblTotalEntrada 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   1020
         TabIndex        =   3
         Top             =   60
         Width           =   1365
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   1020
         TabIndex        =   2
         Top             =   660
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   11685
      TabIndex        =   8
      Top             =   60
      Width           =   11715
      Begin VB.Image Image1 
         Height          =   825
         Left            =   60
         Picture         =   "Fluxo_Caixa.frx":BDBA
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FLUXO DE CAIXA"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   300
         Width           =   2640
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   9210
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16563
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "20:36"
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
      Height          =   5715
      Left            =   60
      TabIndex        =   28
      Top             =   2160
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   10081
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdAbrirCaixa 
      Height          =   375
      Left            =   2940
      TabIndex        =   29
      Top             =   7980
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Reabrir"
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
      MICON           =   "Fluxo_Caixa.frx":124FB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExcluirCaixa 
      Height          =   375
      Left            =   6420
      TabIndex        =   30
      Top             =   7980
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Excluir"
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
      MICON           =   "Fluxo_Caixa.frx":12517
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdMostarCaixa 
      Height          =   375
      Left            =   60
      TabIndex        =   31
      Top             =   7980
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Mostrar"
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
      MICON           =   "Fluxo_Caixa.frx":12533
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimirCaixa 
      Height          =   375
      Left            =   1500
      TabIndex        =   32
      Top             =   7980
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Imprimir"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Fluxo_Caixa.frx":1254F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFecharCaixa 
      Height          =   375
      Left            =   4380
      TabIndex        =   36
      Top             =   7980
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Fechar Temporariamente"
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
      MICON           =   "Fluxo_Caixa.frx":1256B
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
Attribute VB_Name = "Fluxo_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Private printSQL As String
Dim varSaberBotaoCaixa As String
Dim sSQL As String
Dim r As ADODB.Recordset
Private Sub AbrirCaixa()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "UPDATE caixa_dia SET " & _
   "status = 0, REABERTO = 1 " & _
   " WHERE (codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ") and (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "');"
dbData.Execute sSQL
End Sub

Private Sub FecharCaixa()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "UPDATE caixa_dia SET " & _
   "status = 1 " & _
   " WHERE (codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ") and (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "');"
dbData.Execute sSQL
End Sub


Private Sub PreencherAno()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAno.AddItem i
Next
End Sub

Private Sub PreencherCaixa()
cboCaixa.AddItem "CAIXA01"
cboCaixa.AddItem "CAIXA02"
cboCaixa.AddItem "CAIXA03"
cboCaixa.AddItem "CAIXA04"
cboCaixa.AddItem "TODOS"
End Sub

Private Sub PreencherConsulta()
cboConsulta.AddItem "TODOS"
cboConsulta.AddItem "DATA"
cboConsulta.AddItem "MENSAL"
End Sub

Private Sub PreencherMes()
Dim vMes As Integer
cboMes.Clear

For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next
End Sub

Private Sub PreencherOrganizar()
cboOrganizar.AddItem "CÓD. CAIXA"
cboOrganizar.AddItem "DATA"
cboOrganizar.AddItem "CAIXA"
End Sub

Private Sub VerificarCaixa()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT status, CAIXA, CODCAIXA " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    MsgBox "Já Existe um caixa aberto. Não é possivel ter 2 caixas abertos simultaneamente!", vbInformation, "Aviso do Sistema"
Else
    AbrirCaixa
End If
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

Private Sub cboConsulta_GotFocus()
cboConsulta.Clear
cboConsulta.AddItem "TODOS"
cboConsulta.AddItem "DATA"
cboConsulta.AddItem "MENSAL"
cboConsulta.AddItem "PERÍODO"
moCombo.AttachTo cboConsulta
End Sub


Private Sub cboConsulta_LostFocus()
If cboConsulta.Text = "TODOS" Then
    lblMes.Visible = False
    lblAno.Visible = False
    lblAte.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
    cmdConsData.Visible = False
    mskConsData.Visible = False
    cmdConsData2.Visible = False
    mskConsData2.Visible = False
ElseIf cboConsulta.Text = "DATA" Then
    lblMes.Visible = True
    lblMes.Caption = "Data"
    lblAno.Visible = False
    lblAte.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
    cmdConsData.Visible = True
    mskConsData.Visible = True
    cmdConsData2.Visible = False
    mskConsData2.Visible = False
ElseIf cboConsulta.Text = "MENSAL" Then
    lblMes.Visible = True
    lblMes.Caption = "Mês"
    lblAno.Caption = "Ano"
    lblAno.Visible = True
    lblAte.Visible = False
    cboMes.Visible = True
    cboAno.Visible = True
    cmdConsData.Visible = False
    mskConsData.Visible = False
    cmdConsData2.Visible = False
    mskConsData2.Visible = False
ElseIf cboConsulta.Text = "PERÍODO" Then
    lblMes.Visible = True
    lblMes.Caption = "Data Inicial"
    lblAno.Visible = True
    lblAno.Caption = "Data Final"
    lblAte.Visible = True
    cboMes.Visible = False
    cboAno.Visible = False
    cmdConsData.Visible = True
    mskConsData.Visible = True
    cmdConsData2.Visible = True
    mskConsData2.Visible = True
End If
End Sub


Private Sub cboMes_GotFocus()
Dim vMes As Integer

cboMes.Clear

For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMes
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub cboOrganizar_GotFocus()
cboOrganizar.Clear
cboOrganizar.AddItem "CÓD. CAIXA"
cboOrganizar.AddItem "DATA"
cboOrganizar.AddItem "CAIXA"
moCombo.AttachTo cboOrganizar
End Sub


Private Sub cmdAbrirCaixa_Click()
If Grid.TextMatrix(Grid.Row, 13) = "ABERTO" Then MsgBox "O caixa já se encontra aberto!", vbInformation, "Aviso do Sistema": Exit Sub
txtSenha.Text = ""
frmSenha.Visible = True
txtSenha.SetFocus
varSaberBotaoCaixa = "ABRIR"
End Sub

Private Sub cmdConsData2_Click()
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

mskConsData2 = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdExcluirCaixa_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If Grid.TextMatrix(Grid.Row, 3) = "00000" Then MsgBox "Não é permitido excluir caixas anterior as atualizações", vbExclamation, "Aviso do sistema": Exit Sub

'parcelas
sSQL = "SELECT * from PARCELAS WHERE CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & " "
Set r = dbData.OpenRecordset(sSQL)

Dim varQuantParc As Boolean
If r.RecordCount > 0 Then varQuantParc = True Else varQuantParc = False

If r.State <> 0 Then r.Close
Set r = Nothing

'parcelas_haver
sSQL = "SELECT * from parcelas_haver WHERE CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & " "
Set r = dbData.OpenRecordset(sSQL)

Dim varQuantParcHaver As Boolean
If r.RecordCount > 0 Then varQuantParcHaver = True Else varQuantParcHaver = False

If r.State <> 0 Then r.Close
Set r = Nothing

'Caixa_sangria
sSQL = "SELECT * from caixa_saida WHERE CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & " "
Set r = dbData.OpenRecordset(sSQL)

Dim varQuantCaixaSaida As Boolean
If r.RecordCount > 0 Then varQuantCaixaSaida = True Else varQuantCaixaSaida = False

If r.State <> 0 Then r.Close
Set r = Nothing

'Caixa_sangria
sSQL = "SELECT * from caixa_entrada WHERE CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & " "
Set r = dbData.OpenRecordset(sSQL)

Dim varQuantCaixaEntrada As Boolean
If r.RecordCount > 0 Then varQuantCaixaEntrada = True Else varQuantCaixaEntrada = False

If r.State <> 0 Then r.Close
Set r = Nothing

'apagar o caixa

If ShowMsg("Deseja excluir o Caixa '" & Grid.TextMatrix(Grid.Row, 2) & "' com o Cód. Caixa: " & Grid.TextMatrix(Grid.Row, 3) & " ?", vbInformation + vbYesNo) = vbNo Then Exit Sub

If varQuantParc = False And varQuantParcHaver = False And varQuantCaixaSaida = False And varQuantCaixaEntrada = False And Grid.TextMatrix(Grid.Row, 13) = "FECHADO" Then
    sSQL = "delete from caixa_dia WHERE CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & " "
    dbData.Execute (sSQL)
Else
    MsgBox "Não é permitido excluir caixa com itens no totalizados", vbInformation, "Aviso do Sistema"
End If

cmdExibir_Click
End Sub

Private Sub cmdExibir_Click()
Dim INDICE As String
Dim varQualCaixa As String

If cboConsulta.Text = "" Then Exit Sub

'Verifica qual o índice a ser utlizado
If cboOrganizar.Text = "CÓD. CAIXA" Then
   INDICE = " CODCAIXA;"
ElseIf cboOrganizar.Text = "DATA" Then
   INDICE = " DATA_ABERTURA;"
ElseIf cboOrganizar.Text = "CAIXA" Then
   INDICE = " CAIXA;"
Else
    INDICE = " DATA_ABERTURA;"
End If

'Verifica qual o índice a ser utlizado
If cboCaixa.Text = "CAIXA01" Then
   varQualCaixa = " where caixa = 'CAIXA01' and "
ElseIf cboCaixa.Text = "CAIXA02" Then
   varQualCaixa = " where caixa = 'CAIXA02' and "
ElseIf cboCaixa.Text = "CAIXA03" Then
   varQualCaixa = " where caixa = 'CAIXA03' and "
ElseIf cboCaixa.Text = "TODOS" Then
   varQualCaixa = " where "
Else
    varQualCaixa = " where caixa = 'CAIXA01' and "
End If

If cboConsulta.Text = "TODOS" Then
   sSQL = "SELECT *, CASE status WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END AS varStatus, CASE reaberto WHEN 0 THEN 'NÃO' ELSE 'SIM' END AS varReaberto FROM caixa_dia " & varQualCaixa & " CODIGO = CODIGO ORDER BY " & INDICE
ElseIf cboConsulta.Text = "MENSAL" Then
   If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
   sSQL = "SELECT *, CASE status WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END AS varStatus, CASE reaberto WHEN 0 THEN 'NÃO' ELSE 'SIM' END AS varReaberto FROM caixa_dia " & varQualCaixa & " (MONTH(data_abertura) = " & cboMes.ListIndex + 1 & ") AND (YEAR(data_abertura) = " & cboAno & ") ORDER BY " & INDICE
ElseIf cboConsulta.Text = "DATA" Then
   If mskConsData.Text = "" Then Exit Sub
   sSQL = "SELECT *, CASE status WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END AS varStatus, CASE reaberto WHEN 0 THEN 'NÃO' ELSE 'SIM' END AS varReaberto FROM caixa_dia " & varQualCaixa & " (data_abertura = CONVERT(DATETIME, '" & Format(mskConsData, ocDATA) & "', 103)) ORDER BY " & INDICE
ElseIf cboConsulta.Text = "PERÍODO" Then
   If Not IsDate(mskConsData.Text) = True Or Not IsDate(mskConsData2.Text) = True Then Exit Sub
   sSQL = "SELECT *, CASE status WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END AS varStatus, CASE reaberto WHEN 0 THEN 'NÃO' ELSE 'SIM' END AS varReaberto FROM caixa_dia " & varQualCaixa & " ( data_abertura  >= CONVERT(DATETIME, '" & Format(mskConsData.Text, ocDATA) & "', 103)) AND ( data_abertura  <= CONVERT(DATETIME, '" & Format(mskConsData2.Text, ocDATA) & "', 103)) ORDER BY " & INDICE
End If
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer

With Grid
   .Clear
   .Cols = 15
   .rows = 2
   
   .ColWidth(0) = 0
   
   .ColWidth(1) = 0
   .ColWidth(2) = 850
   .ColWidth(3) = 850
   .ColWidth(4) = 850
   .ColWidth(5) = 700
   .ColWidth(6) = 500
   .ColWidth(7) = 1100
   .ColWidth(8) = 1000
   .ColWidth(9) = 1100
   .ColWidth(10) = 850
   .ColWidth(11) = 700
   .ColWidth(12) = 500
   .ColWidth(13) = 1000
   .ColWidth(14) = 1100
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "CAIXA"
   .TextMatrix(0, 3) = "CÓD.CX"
   .TextMatrix(0, 4) = "ABERTURA"
   .TextMatrix(0, 5) = "HORA"
   .TextMatrix(0, 6) = "FUNC."
   .TextMatrix(0, 7) = "ENTRADAS"
   .TextMatrix(0, 8) = "SAÍDAS"
   .TextMatrix(0, 9) = "SALDO"
   .TextMatrix(0, 10) = "FECHADO"
   .TextMatrix(0, 11) = "DATA"
   .TextMatrix(0, 12) = "FUNC."
   .TextMatrix(0, 13) = "STATUS"
   .TextMatrix(0, 14) = "REABERTO"

   .Redraw = False
   i = 1
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.rows - 1, 2) = rTabela("CAIXA")
         .TextMatrix(.rows - 1, 3) = Format(ValidateNull(rTabela("CODCAIXA")), "00000")
         .TextMatrix(.rows - 1, 4) = Format(rTabela("DATA_ABERTURA"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 5) = Format(rTabela("HORA_ABERTURA"), "hh:mm")
         .TextMatrix(.rows - 1, 6) = rTabela("COD_FUNC_ABERTURA")
         .TextMatrix(.rows - 1, 7) = Format(rTabela("entrada"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = Format(rTabela("saida"), ocMONEY)
         .TextMatrix(.rows - 1, 9) = Format(rTabela("saldo"), ocMONEY)
         .TextMatrix(.rows - 1, 10) = Format(rTabela("DATA_FECHAMENTO"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 11) = Format(rTabela("HORA_FECHAMENTO"), "hh:mm")
         .TextMatrix(.rows - 1, 12) = rTabela("COD_FUNC_FECHAMENTO")
         .TextMatrix(.rows - 1, 13) = rTabela("varStatus")
         .TextMatrix(.rows - 1, 14) = rTabela("varReaberto")
         
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If

   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 9
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
   
   .rows = .rows - 1
   .Redraw = True
End With

lblTotalEntrada.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
lblTotalSaida.Caption = Format(SomaGrid(Grid, 8), ocMONEY)
lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
End Sub

Private Sub cmdFechar_Click()
txtSenha.Text = ""
frmSenha.Visible = False
varSaberBotaoCaixa = ""
End Sub

Private Sub cmdFecharCaixa_Click()
If Grid.TextMatrix(Grid.Row, 13) = "FECHADO" Then MsgBox "O caixa já se encontra fechado!", vbInformation, "Aviso do Sistema": Exit Sub
txtSenha.Text = ""
frmSenha.Visible = True
txtSenha.SetFocus
varSaberBotaoCaixa = "FECHAR"
End Sub

Private Sub cmdImprimir_Click()
Dim r As ADODB.Recordset

'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

If cboConsulta.Text = "MENSAL" Then
   If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
End If

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Fluxo_Caixa.Relatorio.Recordset = r

REL_Fluxo_Caixa.dfENTRADAS.Caption = Format(lblTotalEntrada.Caption, ocMONEY)
REL_Fluxo_Caixa.dfSAIDAS.Caption = Format(lblTotalSaida.Caption, ocMONEY)
REL_Fluxo_Caixa.dfTOTAIS.Caption = Format(lblTotal.Caption, ocMONEY)

If cboConsulta.Text = "MENSAL" Then
   REL_Fluxo_Caixa.dfMES.Caption = "Mês = " & cboMes.Text & "/" & cboAno.Text
ElseIf cboConsulta.Text = "DATA" Then
   REL_Fluxo_Caixa.dfMES.Caption = "Data = " & mskConsData.Text
ElseIf cboConsulta.Text = "PERÍODO" Then
   REL_Fluxo_Caixa.dfMES.Caption = "Período = " & mskConsData.Text & " até " & mskConsData2.Text
Else
   REL_Fluxo_Caixa.dfMES.Caption = "TODOS"
End If

cboConsulta.AddItem "TODOS"
cboConsulta.AddItem "DATA"
cboConsulta.AddItem "MENSAL"
cboConsulta.AddItem "PERÍODO"

REL_Fluxo_Caixa.Relatorio.NomeImpressora = var_Impressora
REL_Fluxo_Caixa.Relatorio.Ativar
Unload REL_Fluxo_Caixa

Me.Show 1
End Sub

Private Sub cmdImprimirCaixa_Click()
ImprimirCaixa
End Sub

Private Sub ImprimirCaixa()
'=========== DEFINIR A IMPRESSORA
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

'Me.Hide
        
Dim sSQL As String
Dim r As ADODB.Recordset

Dim SETOR_CAIXA As String
Dim var_Setor As String
Dim varTipoCartao2 As String
'       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") "

'================ PREENCHER O GRID
        Dim Maquina_Parcela As String
        Maquina_Parcela = "AND (parcelas.caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') "
        
        Dim Maquina_Haver As String
        Maquina_Haver = "AND (parcelas_haver.caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') "
        
        Dim Maquina_Suprimento As String
        Maquina_Suprimento = "AND (caixa_entrada.caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') "
        
        Dim Maquina_Sangria As String
        Maquina_Sangria = "AND (caixa_saida.caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') "
        
        SETOR_CAIXA = "AND (pedidos.tipo_pedido = 'VENDA') "
        
        var_Setor = "AND (setor <> 'BOSTA') "
        
        sSQL = "SELECT " & _
           "parcelas.tipo as varTipoLanc, " & _
           "parcelas.hora AS varHora, " & _
           "parcelas.codigo AS varCodigo, " & _
           "pedidos.cod_pedido AS varCodPedido, " & _
           "cliente.nome AS varCliente, " & _
           "parcelas.forma_pgto AS varFormaPgto, " & _
           "parcelas.valor_final AS varValorLanc, " & _
           "0 AS varValorSaida, " & _
           "(parcelas.valor_final  - 0) AS campo04, " & _
           "(CASE WHEN parcelas.tipo_cartao = 'D' THEN 'DÉBITO' WHEN parcelas.tipo_cartao = 'C'  THEN 'CRÉDITO' Else '' End) AS varTipoCartao, " & _
           "parcelas.pagamento AS data, " & _
           "pedidos.cod_pedido AS pedido, " & _
           "pedidos.cod_cliente AS cliente, " & _
           "'' AS setor, " & _
           "parcelas.caixa " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
           "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ") " & Maquina_Parcela & _
           "UNION ALL "
        
        'parcelas_haveres
        sSQL = sSQL & "SELECT " & _
           "'HAVER' as varTipoLanc, " & _
           "parcelas_haver.hora AS varHora, " & _
           "0 AS varCodigo, " & _
           "parcelas.cod_pedido AS varCodPedido, " & _
           "(SELECT nome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE (pedidos.cod_pedido = parcelas.cod_pedido)) AS varCliente, " & _
           "parcelas_haver.forma_pgto AS varFormaPgto, " & _
           "parcelas_haver.valor_haver AS varValorLanc, " & _
           "0 AS varValorSaida, " & _
           "parcelas_haver.valor_haver AS campo04, " & _
           "(CASE WHEN parcelas_haver.tipo_cartao = 'D' THEN 'DÉBITO' WHEN parcelas_haver.tipo_cartao = 'C'  THEN 'CRÉDITO' Else '' End) AS varTipoCartao, " & _
           "parcelas_haver.haver AS data, " & _
           "parcelas_haver.codigo AS campotc, " & _
           "0 AS cliente, " & _
           "'' AS  setor, " & _
           "'' AS maquina " & _
           "FROM parcelas_haver INNER JOIN parcelas ON parcelas_haver.cod_parcela = parcelas.codigo " & _
           "WHERE (parcelas_haver.codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ") " & Maquina_Haver & _
           "UNION ALL "
        
        'suprimentos
        sSQL = sSQL & "SELECT " & _
           "'SUPRIMENTO' as varTipoLanc, " & _
           "hora AS varHora, " & _
           "0 AS varCodigo, " & _
           "0 AS varCodPedido, " & _
           "descricao AS varCliente, " & _
           "caixa_entrada.forma_pgto AS varFormaPgto, " & _
           "valor AS varValorLanc, " & _
           "0 AS varValorSaida, " & _
           "valor AS campo04, " & _
           "'' AS varTipoCartao, " & _
           "data, " & _
           "0 AS pedido, " & _
           "0 AS cliente, " & _
           "setor, " & _
           " '' AS maquina " & _
           "FROM caixa_entrada WHERE (caixa_entrada.codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ")  " & Maquina_Suprimento & var_Setor & _
           "UNION ALL "
           
        'sangria
        sSQL = sSQL & "SELECT " & _
           "'SANGRIA' as varTipoLanc, " & _
           "caixa_saida.hora AS varHora, " & _
           "0 AS varCodigo, " & _
           "0 AS varCodPedido, " & _
           "caixa_saida.descricao AS varCliente, " & _
           "'DINHEIRO' AS varFormaPgto, " & _
           "0 AS varValorLanc, " & _
           "caixa_saida.valor AS varValorSaida, " & _
           "(0 - caixa_saida.valor) AS campo04, " & _
           "'' AS varTipoCartao, " & _
           "data, " & _
           "0, " & _
           "0, " & _
           "setor, " & _
           "'' AS maquina " & _
           "FROM caixa_saida WHERE (caixa_saida.codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ")  " & Maquina_Sangria & var_Setor & _
           " ORDER BY 2"
        
        Set r = dbData.OpenRecordset(sSQL)

        If Grid.TextMatrix(Grid.Row, 7) = "0,00" Then   'fiz esse if para imprimir caixa sem saldo
            If r.State <> 0 Then r.Close
            Set r = Nothing
        End If


''==================== PEGAR OS DADOS DO FECHAMENTO
Dim sSQLusuario As String
Dim r_usuario As ADODB.Recordset

sSQLusuario = "SELECT DATA_ABERTURA, HORA_ABERTURA, COD_FUNC_ABERTURA, DATA_FECHAMENTO, HORA_FECHAMENTO, COD_FUNC_FECHAMENTO, (CASE WHEN status = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS VarStatus, " & _
        "(SELECT Usuario.Login FROM Usuario INNER JOIN caixa_dia ON Usuario.Codigo = caixa_dia.COD_FUNC_ABERTURA wHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ")) AS Nome_Abertura, " & _
        "(SELECT Usuario_2.Login FROM Usuario AS Usuario_2 INNER JOIN caixa_dia AS caixa_dia_2 ON Usuario_2.Codigo = caixa_dia_2.COD_FUNC_FECHAMENTO WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ")) AS Nome_Fechamento " & _
       "FROM caixa_dia AS caixa_dia_1 " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ");"
Set r_usuario = dbData.OpenRecordset(sSQLusuario)

'SETAR O RELATORIO
Set REL_Caixa_Fech_Imp.ReportMain1.Recordset = r

'==================== CABEÇALHO
    If Not r_usuario.EOF Then
    REL_Caixa_Fech_Imp.txtDHead.Caption = "FECHAMENTO DE CAIXA - ABERTURA: " & Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
    End If

'===========================CALCULO DOS TOTAIS

'VENDAS DINHEIRO============
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaVendas " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'VENDA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalDinheiroVenda As Currency
If Not r.EOF Then
    varTotalDinheiroVenda = Format(ValidateNull(r("varSomaVendas")), "#,##0.00")
Else
    varTotalDinheiroVenda = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp.rfDinheiro.Caption = Format(varTotalDinheiroVenda, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'VENDA')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfDinheiroQuant.Caption = Format(r.RecordCount, "000") & " "


'PARCELAS DINHEIRO============
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaParcelas " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'PARCELA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalDinheiroParcela As Currency
If Not r.EOF Then
    varTotalDinheiroParcela = Format(ValidateNull(r("varSomaParcelas")), "#,##0.00")
Else
    varTotalDinheiroParcela = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp.rfParcelas.Caption = Format(varTotalDinheiroParcela, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'PARCELA')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfParcelasQuant.Caption = Format(r.RecordCount, "000") & " "

'PARCELAS HAVER============
sSQL = "SELECT SUM(VALOR_HAVER) AS varSomaHaveres " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'DINHEIRO')"
Set r = dbData.OpenRecordset(sSQL)

Dim varValorHaveres As Currency
If Not r.EOF Then
    varValorHaveres = Format(ValidateNull(r("varSomaHaveres")), "#,##0.00")
Else
    varValorHaveres = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp.rfHaveres.Caption = Format(varValorHaveres, "#,##0.00") & " "

sSQL = "SELECT CODIGO " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'DINHEIRO')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfHaveresQuant.Caption = Format(r.RecordCount, "000") & " "

'CARTÃO============
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaCartao1 " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'CARTAO')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalCartao As Currency
varTotalCartao = Format(ValidateNull(r("varSomaCartao1")))

sSQL = "SELECT SUM(VALOR_HAVER) AS varSomaCartao2 " & _
       "FROM parcelas_HAVER " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'CARTAO') "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalCartao2 As Currency
varTotalCartao2 = Format(ValidateNull(r("varSomaCartao2")))

varTotalCartao = varTotalCartao + varTotalCartao2

'If Not r.EOF Then
'    varTotalCartao = ValidateNull(r("varTotalCartao"))
'Else
'    varTotalCartao = Format(0, "#,##0.00")
'End If

REL_Caixa_Fech_Imp.rfCartao.Caption = Format(varTotalCartao, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'CARTAO')"
Set r = dbData.OpenRecordset(sSQL)

Dim ContaCartao1 As Integer
Dim ContaCartao2 As Integer
ContaCartao1 = r.RecordCount

sSQL = "SELECT codigo " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'CARTAO')"
Set r = dbData.OpenRecordset(sSQL)

ContaCartao2 = ContaCartao1 + r.RecordCount

REL_Caixa_Fech_Imp.rfCartaoQuant.Caption = Format(ContaCartao2, "000") & " "

'CHEQUE============

sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaCheque " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'CHEQUE')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalCheque As Currency

If Not r.EOF Then
    varTotalCheque = Format(ValidateNull(r("varSomaCheque")), "#,##0.00")
Else
    varTotalCheque = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp.rfCheque.Caption = Format(varTotalCheque, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'CHEQUE')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfChequeQuant.Caption = Format(r.RecordCount, "000") & " "

'DEPOSITO/TRANSFERENCIA/BOLETO/FINANCEIRA============
'boleto
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaBoleto " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'BOLETO')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalBoleto As Currency
If Not r.EOF Then
    varTotalBoleto = ValidateNull(r("varSomaBoleto"))
Else
    varTotalBoleto = 0
End If

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'BOLETO')"
Set r = dbData.OpenRecordset(sSQL)

Dim contaBoleto As Integer
contaBoleto = Format(r.RecordCount, "000")

'transferencia
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaTransferencia " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'TRANSFERENCIA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalTransferencia As Currency
If Not r.EOF Then
    varTotalTransferencia = ValidateNull(r("varSomaTransferencia"))
Else
    varTotalTransferencia = 0
End If

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'TRANSFERENCIA')"
Set r = dbData.OpenRecordset(sSQL)

Dim contaTransferencia As Integer
contaTransferencia = Format(r.RecordCount, "000")

'deposito
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaDeposito " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'DEPOSITO')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalDeposito As Currency
If Not r.EOF Then
    varTotalDeposito = ValidateNull(r("varSomaDeposito"))
Else
    varTotalDeposito = 0
End If

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'DEPOSITO')"
Set r = dbData.OpenRecordset(sSQL)

Dim contaDeposito As Integer
contaDeposito = Format(r.RecordCount, "000")

'financeira
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaFinanceira " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'FINANCEIRA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalFinanceira As Currency
If Not r.EOF Then
    varTotalFinanceira = ValidateNull(r("varSomaFinanceira"))
Else
    varTotalFinanceira = 0
End If

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (FORMA_PGTO = 'FINANCEIRA')"
Set r = dbData.OpenRecordset(sSQL)

'soma
Dim varTotalBTD As Currency
varTotalBTD = varTotalBoleto + varTotalTransferencia + varTotalDeposito + varTotalFinanceira

REL_Caixa_Fech_Imp.rfOutros.Caption = Format(varTotalBTD, "#,##0.00") & " "

Dim contaFinanceira As Integer
contaFinanceira = Format(r.RecordCount, "000")

Dim ContaOutros As Integer
ContaOutros = contaFinanceira + contaDeposito + contaTransferencia + contaBoleto

REL_Caixa_Fech_Imp.rfOutrosQuant.Caption = Format(ContaOutros, "000") & " "

'SANGRIA============
sSQL = "SELECT SUM(VALOR) AS varSomaSangria " & _
       "FROM caixa_saida " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") and (FONTE = 'CAIXA ATUAL') "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalSangria As Currency

If Not r.EOF Then
    varTotalSangria = Format(ValidateNull(r("varSomaSangria")), "#,##0.00")
Else
    varTotalSangria = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp.rfSaida.Caption = Format(varTotalSangria, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM caixa_saida " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") and (FONTE = 'CAIXA ATUAL') "
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfSaidaQuant.Caption = Format(r.RecordCount, "000") & " "

'SUPRIMENTO============
sSQL = "SELECT SUM(VALOR) AS varSomaSuprimento " & _
       "FROM caixa_entrada " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalSuprimento As Currency

If Not r.EOF Then
    varTotalSuprimento = Format(ValidateNull(r("varSomaSuprimento")), "#,##0.00")
Else
    varTotalSuprimento = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp.rfSuprimentos.Caption = Format(varTotalSuprimento, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM caixa_entrada " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") "
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfSuprimentosQuant.Caption = Format(r.RecordCount, "000") & " "

'TROCO============
sSQL = "SELECT SUM(VALOR) AS varSomaTROCO " & _
       "FROM caixa_troco " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") "
Set r = dbData.OpenRecordset(sSQL)

'If Not r.EOF Then
'    REL_Caixa_Fech_Imp.rfTroco.Caption = Format(ValidateNull(r("varSomaTROCO")), "#,##0.00") & " "
'Else
'    REL_Caixa_Fech_Imp.rfTroco.Caption = Format(0, "#,##0.00") & " "
'End If

'VENDA A PRAZO ================
sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaPrazo " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ") AND pedidos.caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalPrazo As Currency

If Not r.EOF Then
    varTotalPrazo = ValidateNull(r("varSomaPrazo"))
Else
    varTotalPrazo = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp.rfPrazo.Caption = Format(varTotalPrazo, "#,##0.00") & " "

sSQL = "SELECT parcelas.cod_pedido " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ") AND pedidos.caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfPrazoQuant.Caption = Format(r.RecordCount, "000") & " "

'CALCULAR TOTAIS================
Dim varTotaisEntrada As Currency
Dim varTotaisSaida As Currency
varTotaisEntrada = varTotalDinheiroVenda + varTotalDinheiroParcela + varValorHaveres + varTotalCheque + varTotalSuprimento
varTotaisSaida = varTotaisEntrada - varTotalSangria

REL_Caixa_Fech_Imp.rfSaldoFisico.Caption = Format(varTotaisSaida, "#,##0.00") & " "

Dim varTotaisGeral As Currency
varTotaisGeral = varTotaisSaida + varTotalBTD + varTotalCartao

REL_Caixa_Fech_Imp.rfSaldoGeral.Caption = Format(varTotaisGeral, "#,##0.00") & " "

Dim varTotaisFAT As Currency
varTotaisFAT = varTotaisGeral + varTotalPrazo
REL_Caixa_Fech_Imp.rfFaturamento.Caption = Format(varTotaisFAT, "#,##0.00") & " "

'===========================RODAPÉ
If Not r_usuario.EOF Then
    REL_Caixa_Fech_Imp.rfCodUsuarioA.Caption = Format(r_usuario("COD_FUNC_ABERTURA"), "00")
    REL_Caixa_Fech_Imp.rfNomeUsuarioA.Caption = ValidateNull(r_usuario("Nome_Abertura"))
    REL_Caixa_Fech_Imp.rfDataA.Caption = Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
    REL_Caixa_Fech_Imp.rfHoraA.Caption = Format(ValidateNull(r_usuario("HORA_ABERTURA")), "hh:mm")
    
    REL_Caixa_Fech_Imp.rfNomeUsuarioF.Caption = ValidateNull(r_usuario("Nome_Fechamento"))
    If IsNull(r_usuario("DATA_FECHAMENTO")) Then
        REL_Caixa_Fech_Imp.rfDataF.Caption = ""
        REL_Caixa_Fech_Imp.rfCodUsuarioF.Caption = ""
        REL_Caixa_Fech_Imp.rfHoraF.Caption = ""
    Else
        REL_Caixa_Fech_Imp.rfCodUsuarioF.Caption = Format(ValidateNull(r_usuario("COD_FUNC_FECHAMENTO")), "00")
        REL_Caixa_Fech_Imp.rfDataF.Caption = Format(ValidateNull(r_usuario("DATA_FECHAMENTO")), "dd/mm/yyyy")
        REL_Caixa_Fech_Imp.rfHoraF.Caption = Format(ValidateNull(r_usuario("HORA_FECHAMENTO")), "hh:mm")
    End If

    REL_Caixa_Fech_Imp.rfSituacao.Caption = ValidateNull(r_usuario("VARSTATUS"))
End If

REL_Caixa_Fech_Imp.rfCaixa.Caption = Grid.TextMatrix(Grid.Row, 2)
REL_Caixa_Fech_Imp.rfCodCaixa.Caption = Format(Grid.TextMatrix(Grid.Row, 3), "0000")


'=========================CALCULO DO FATURAMENTO

'VENDAS
sSQL = "SELECT ISNULL(SUM(VALOR_FINAL),0) AS varSomaVendasFAT " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (TIPO = 'VENDA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalVendaFAT As Currency
varTotalVendaFAT = ValidateNull(r("varSomaVendasFAT"))
REL_Caixa_Fech_Imp.rfT1.Caption = Format(varTotalVendaFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (TIPO = 'VENDA')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfF1.Caption = Format(r.RecordCount, "000") & " "

'PARCELAS
sSQL = "SELECT ISNULL(SUM(VALOR_FINAL),0) AS varSomaParcelasFAT " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (TIPO = 'PARCELA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalParcelasFAT As Currency
varTotalParcelasFAT = ValidateNull(r("varSomaParcelasFAT"))
REL_Caixa_Fech_Imp.rfT2.Caption = Format(varTotalParcelasFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") AND (TIPO = 'PARCELA')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfF2.Caption = Format(r.RecordCount, "000") & " "

'HAVER
sSQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) AS varSomaHaveresFAT " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ")"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalHaveresFAT As Currency
varTotalHaveresFAT = ValidateNull(r("varSomaHaveresFAT"))
REL_Caixa_Fech_Imp.rfT3.Caption = Format(varTotalHaveresFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ")"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfF3.Caption = Format(r.RecordCount, "000") & " "

'SUPRIMENTO
sSQL = "SELECT ISNULL(SUM(VALOR),0) AS varSomaSuprimentoFAT " & _
       "FROM caixa_entrada " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalSuprimentoFAT As Currency
varTotalSuprimentoFAT = ValidateNull(r("varSomaSuprimentoFAT"))
REL_Caixa_Fech_Imp.rfT4.Caption = Format(varTotalSuprimentoFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM caixa_entrada " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") "
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfF4.Caption = Format(r.RecordCount, "000") & " "

'PRAZO
sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaPrazoFAT " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ") AND pedidos.caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalPrazoFAT As Currency
varTotalPrazoFAT = ValidateNull(r("varSomaPrazoFAT"))
REL_Caixa_Fech_Imp.rfT5.Caption = Format(varTotalPrazoFAT, "#,##0.00") & " "

sSQL = "SELECT parcelas.codigo " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & Grid.TextMatrix(Grid.Row, 3) & ") AND pedidos.caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfF5.Caption = Format(r.RecordCount, "000") & " "

'SANGRIA
sSQL = "SELECT ISNULL(SUM(VALOR),0) AS varSomaSangriaFAT " & _
       "FROM caixa_saida " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") and (FONTE = 'CAIXA ATUAL') "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalSangriaFAT As Currency
varTotalSangriaFAT = ValidateNull(r("varSomaSangriaFAT"))
REL_Caixa_Fech_Imp.rfT6.Caption = Format(varTotalSangriaFAT, "#,##0.00") & " "

sSQL = "SELECT CODIGO " & _
       "FROM caixa_saida " & _
       "WHERE (caixa = '" & Grid.TextMatrix(Grid.Row, 2) & "') and (CODCAIXA = " & Grid.TextMatrix(Grid.Row, 3) & ") and (FONTE = 'CAIXA ATUAL') "
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp.rfF6.Caption = Format(r.RecordCount, "000") & " "

'CALCULAR TOTAIS================
Dim varTotaisEntradaFAT As Currency
Dim varTotaisSaidaFAT As Currency

varTotaisEntradaFAT = varTotalVendaFAT + varTotalParcelasFAT + varTotalHaveresFAT + varTotalSuprimentoFAT + varTotalPrazoFAT
varTotaisSaidaFAT = varTotaisEntradaFAT - varTotalSangriaFAT

REL_Caixa_Fech_Imp.rfFTotal.Caption = Format(varTotaisSaidaFAT, "#,##0.00") & " "


'REL_Caixa_Fech_Imp.Relatorio.NomeImpressora = var_Impressora
REL_Caixa_Fech_Imp.ReportMain1.Ativar
Unload REL_Caixa_Fech_Imp

'Me.Show 1
End Sub
Private Sub cmdMostarCaixa_Click()
varFluxoCaixa = True
varFluxoNomeCaixa = Grid.TextMatrix(Grid.Row, 2)
varFluxoCodCaixa = Grid.TextMatrix(Grid.Row, 3)
varFluxoCaixaSituacao = Grid.TextMatrix(Grid.Row, 13)
varFluxoCaixaData = Format(Grid.TextMatrix(Grid.Row, 4), "dd/mm/yy")
varCodCaixa = Grid.TextMatrix(Grid.Row, 3)
Caixa_Controle_semOS.cmdAbrirCaixa.Visible = True
Caixa_Controle_semOS.cmdAbrirCaixa.Enabled = False
Caixa_Controle_semOS.cmdFecharCaixa.Enabled = False
Caixa_Controle_semOS.cmdTroco.Enabled = False
Caixa_Controle_semOS.cmdImprimir.Enabled = True
Caixa_Controle_semOS.cmdMostrar_Click
Caixa_Controle_semOS.Show
End Sub


Private Sub cmdSenha_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * FROM usuario WHERE (password = '" & txtSenha.Text & "') AND (nivel = 1);" 'desabilitei pelo jacobina usar
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   txtSenha.Text = ""
   frmSenha.Visible = False
   If varSaberBotaoCaixa = "FECHAR" Then
        FecharCaixa
   ElseIf varSaberBotaoCaixa = "ABRIR" Then
        VerificarCaixa
   End If
   varSaberBotaoCaixa = ""
Else
   ShowMsg "ACESSO NEGADO!" & vbCrLf & "Você não tem nivel de acesso a esse recurso", vbInformation
   txtSenha.Text = ""
   frmSenha.Visible = False
   varSaberBotaoCaixa = ""
End If
cmdExibir_Click
End Sub

Private Sub Form_Activate()
'cmdExibir_Click
End Sub

Private Sub cmdConsData_Click()
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

mskConsData = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cboCaixa_GotFocus()
cboCaixa.Clear
cboCaixa.AddItem "CAIXA01"
cboCaixa.AddItem "CAIXA02"
cboCaixa.AddItem "CAIXA03"
cboCaixa.AddItem "CAIXA04"
cboCaixa.AddItem "TODOS"
moCombo.AttachTo cboCaixa
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")

PreencherConsulta
cboConsulta.ListIndex = 2

PreencherOrganizar
cboOrganizar.ListIndex = 1

PreencherCaixa
cboCaixa.ListIndex = 0

cboConsulta_LostFocus

PreencherMes
cboMes.ListIndex = Month(Date) - 1

'PreencherAno
cboAno.Text = Year(Date)

cmdExibir_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub


Private Sub mskConsData_GotFocus()
SelectControl mskConsData
End Sub

Private Sub mskConsData_KeyPress(KeyAscii As Integer)
mskConsData.Mask = "##/##/##"
End Sub


Private Sub mskConsData2_GotFocus()
SelectControl mskConsData2
End Sub


Private Sub mskConsData2_KeyPress(KeyAscii As Integer)
mskConsData2.Mask = "##/##/##"
End Sub


Private Sub txtCodFuncionario_Change()
If txtCodFuncionario.Text = "" Then Exit Sub

sSQL = "SELECT Usuario_permissoes.Codigo, Usuario_permissoes.permissao " & _
       "FROM Usuario_permissoes INNER JOIN Usuario_Acessos ON Usuario_permissoes.Codigo = Usuario_Acessos.Cod_Permissao " & _
       "WHERE (Usuario_permissoes.permissao = 'FLUXO DE CAIXA') AND (Usuario_Acessos.Cod_Usuario = " & txtCodFuncionario.Text & ")"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    cmdAbrirCaixa.Enabled = True
    cmdFecharCaixa.Enabled = True
    cmdExcluirCaixa.Enabled = True
Else
    cmdAbrirCaixa.Enabled = False
    cmdFecharCaixa.Enabled = False
    cmdExcluirCaixa.Enabled = False
End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSenha_Click
End Sub


