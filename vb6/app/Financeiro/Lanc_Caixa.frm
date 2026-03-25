VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Lanc_Caixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LANÇAMENTOS DE CAIXAS"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   ForeColor       =   &H00000000&
   Icon            =   "Lanc_Caixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   9345
      TabIndex        =   7
      Top             =   60
      Width           =   9375
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LANÇAMENTOS"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   180
         Width           =   2400
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   1260
         Picture         =   "Lanc_Caixa.frx":23D2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   60
      ScaleHeight     =   1425
      ScaleWidth      =   9345
      TabIndex        =   6
      Top             =   8160
      Width           =   9375
      Begin VB.Frame Frame4 
         Caption         =   "Último Caixa Fechado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6540
         TabIndex        =   15
         Top             =   60
         Width           =   2715
         Begin VB.TextBox txtConsData 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "0,00"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtUltimaCaixa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "0,00"
            Top             =   460
            Width           =   1575
         End
         Begin VB.TextBox txtUltimoCodCaixa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0,00"
            Top             =   710
            Width           =   1575
         End
         Begin VB.TextBox txtConsSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   255
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0,00"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caixa:"
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
            Left            =   420
            TabIndex        =   27
            Top             =   480
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód.Caixa:"
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
            TabIndex        =   26
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo:"
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
            Left            =   420
            TabIndex        =   18
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
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
            Left            =   480
            TabIndex        =   17
            Top             =   240
            Width           =   480
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdDetalhar 
         Height          =   315
         Left            =   60
         TabIndex        =   33
         Top             =   60
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Mostrar Retiradas"
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
         MICON           =   "Lanc_Caixa.frx":8B13
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMostrarCaixa 
         Height          =   315
         Left            =   1980
         TabIndex        =   34
         Top             =   60
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Mostrar Caixa"
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
         MICON           =   "Lanc_Caixa.frx":8B2F
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
      Height          =   6195
      Left            =   60
      ScaleHeight     =   6165
      ScaleWidth      =   9345
      TabIndex        =   5
      Top             =   1920
      Width           =   9375
      Begin VB.Frame frmDetalhamento 
         Caption         =   "DETALHAMENTO"
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
         Height          =   3435
         Left            =   540
         TabIndex        =   30
         Top             =   1260
         Visible         =   0   'False
         Width           =   8055
         Begin MSFlexGridLib.MSFlexGrid Grid_Det 
            Height          =   3075
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5424
            _Version        =   393216
            BackColor       =   16777215
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdFecharDet 
            Height          =   195
            Left            =   7740
            TabIndex        =   32
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   344
            BTYPE           =   3
            TX              =   "X"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            FCOL            =   192
            FCOLO           =   192
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lanc_Caixa.frx":8B4B
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
         Height          =   5955
         Left            =   60
         TabIndex        =   21
         Top             =   120
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   10504
         _Version        =   393216
         BackColor       =   16777215
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   60
      ScaleHeight     =   1065
      ScaleWidth      =   9345
      TabIndex        =   4
      Top             =   780
      Width           =   9375
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   975
         Left            =   60
         TabIndex        =   12
         Top             =   0
         Width           =   1875
         Begin VB.ComboBox cboConsulta 
            Height          =   315
            Left            =   60
            TabIndex        =   35
            Top             =   480
            Width           =   1755
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Consulta:"
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo"
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
         Left            =   1980
         TabIndex        =   9
         Top             =   0
         Width           =   6375
         Begin ChamaleonBtn.chameleonButton cmdConsData2 
            Height          =   315
            Left            =   2520
            TabIndex        =   29
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
            MICON           =   "Lanc_Caixa.frx":8B67
            PICN            =   "Lanc_Caixa.frx":8B83
            PICH            =   "Lanc_Caixa.frx":AED6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdConsData 
            Height          =   315
            Left            =   1140
            TabIndex        =   28
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
            MICON           =   "Lanc_Caixa.frx":D229
            PICN            =   "Lanc_Caixa.frx":D245
            PICH            =   "Lanc_Caixa.frx":F598
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
            Left            =   1620
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Visible         =   0   'False
            Width           =   1515
         End
         Begin MSMask.MaskEdBox mskTermino 
            Height          =   315
            Left            =   1500
            TabIndex        =   3
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdExibir 
            Height          =   615
            Left            =   3300
            TabIndex        =   19
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
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
            MICON           =   "Lanc_Caixa.frx":118EB
            PICN            =   "Lanc_Caixa.frx":11907
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
            Left            =   4740
            TabIndex        =   20
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
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
            MICON           =   "Lanc_Caixa.frx":121E1
            PICN            =   "Lanc_Caixa.frx":121FD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblTermino 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Termino"
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
            Left            =   1500
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblInicio 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Inicio"
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
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   480
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
            Left            =   1620
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   360
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   22
      Top             =   9660
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12488
            Text            =   "Online.Info - Informática  - Tel.: (89) 9 8817-7036"
            TextSave        =   "Online.Info - Informática  - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "20:00"
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
Attribute VB_Name = "Lanc_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Private moCombo As cComboHelper
Private printSQL As String
Dim i As Integer

Private Sub FormatarGrid_Retirada(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Det
      .Clear
      .Cols = 6
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 950
      .ColWidth(2) = 1500
      .ColWidth(3) = 0
      .ColWidth(4) = 3700
      .ColWidth(5) = 1250
      
      .TextMatrix(0, 1) = "PGTO"
      .TextMatrix(0, 2) = "ORIGEM"
      .TextMatrix(0, 3) = "COD_DESCRICAO"
      .TextMatrix(0, 4) = "DESCRICAO"
      .TextMatrix(0, 5) = "VALOR"
      
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
            'mudar a cor da coluna
            'For i = 1 To .Rows - 1
            '   .Row = i
            '   .Col = 6
            '   .CellBackColor = &HC0FFFF
            'Next
            
            'ALINHAMENTO
            .ColAlignment(1) = 3
            
            .TextMatrix(.rows - 1, 1) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 2) = rTabela("tipo")
            .TextMatrix(.rows - 1, 3) = rTabela("cod_descricao")
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("descricao"))
            .TextMatrix(.rows - 1, 5) = Format(rTabela("valor"), ocMONEY)
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .Redraw = True
      
      .rows = .rows - 1
      .BackColor = &HC0FFFF
      
      'SOMAR REGUSTROS
      'lblTotal.Caption = Format(SomaGrid(Grid, 5), "##,##0.00")
   End With
End Sub

Private Sub PreencherConsulta()
cboConsulta.AddItem "TODOS"
cboConsulta.AddItem "MENSAL"
cboConsulta.AddItem "PERÍODO"
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
Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer
Dim m_Saldo As Currency
Dim Saldo_Anterior As Currency

With Grid
   .Clear
   .Cols = 9
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 950
   .ColWidth(3) = 1400
   .ColWidth(4) = 1400
   .ColWidth(5) = 1400
   .ColWidth(6) = 1500
   .ColWidth(7) = 1100
   .ColWidth(8) = 1200
   
   .TextMatrix(0, 1) = "CÓD"
   .TextMatrix(0, 2) = "DATA"
   .TextMatrix(0, 3) = "SALDO ANT."
   .TextMatrix(0, 4) = "ENTRADA"
   .TextMatrix(0, 5) = "RETIRADA"
   .TextMatrix(0, 6) = "SALDO ATUAL"
   .TextMatrix(0, 7) = "CAIXA"
   .TextMatrix(0, 8) = "CÓD. CAIXA"
   
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
   
   .BackColor = &HFFFFFF
   .Redraw = False
     
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         'mudar a cor da coluna
         For i = 1 To .rows - 1
            .Row = i
            .Col = 6
            .CellBackColor = &HC0FFFF
         Next
         
         'ALINHAMENTO
         .ColAlignment(1) = 3
         
         'If .Rows <= 2 Then
         '   Saldo_Anterior = Format(0, ocMONEY)
         'ElseIf .Rows > 2 Then
         '   Saldo_Anterior = Format(.TextMatrix(.Rows - 2, 6), ocMONEY)
         'End If
         
         .TextMatrix(.rows - 1, 1) = Format(rTabela("CODIGO"), "0000")
         .TextMatrix(.rows - 1, 2) = Format(rTabela("DATA"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = Format(rTabela("SALDO_ANTERIOR"), ocMONEY)
         .TextMatrix(.rows - 1, 4) = Format(rTabela("ENTRADA"), ocMONEY)
         .TextMatrix(.rows - 1, 5) = Format(rTabela("RETIRADA"), ocMONEY)
         .TextMatrix(.rows - 1, 6) = Format(rTabela("SALDO_ATUAL"), ocMONEY)
         .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("CAIXA"))
         .TextMatrix(.rows - 1, 8) = Format(rTabela("CODCAIXA"), "00000")
         
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .rows = .rows - 1
   .Redraw = True
End With
End Sub

Private Sub Mostrar_Ult_Caixa()
'   Dim sSQL As String
'   Dim r As ADODB.Recordset
   
'   sSQL = "SELECT TOP 1 codigo, data_abertura, saldo FROM caixa_dia ORDER BY data_abertura DESC;"
'   Set r = dbData.OpenRecordset(sSQL)
   
'   If Not r.BOF Then
'      txtConsSaldo.Text = Format(r("saldo"), ocMONEY)
'      txtConsData.Text = Format(r("data_abertura"), "dd/mm/yy")
'   End If
   
'   If r.State <> 0 Then r.Close
'   Set r = Nothing
End Sub

Private Sub PreencherGrid_Lancamentos()
   Dim INDICE As String       'INDICE PARA ORGANIZAR OS DADOS
   Dim Tipo_Data As String
   
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   'If optOrdVencimento.Value = True Then
   '   INDICE = "vencimento"
   'ElseIf optOrdFavorecido.Value = True Then
   '   INDICE = "favorecido"
   'ElseIf optOrdReferente.Value = True Then
   '   INDICE = "descricao"
   'ElseIf optOrdForma.Value = True Then
   '   INDICE = "forma"
   'Else
   '   optOrdVencimento.Value = True
      INDICE = "vencimento"
   'End If

   If cboConsulta.Text = "TODOS" Then
      'lblCONmes.Enabled = False
      cboMes.Enabled = False
      cboAno.Enabled = False
      'cmdCONmes.Enabled = False
      'lblCONint1.Enabled = False
      mskInicio.Enabled = False
      'lblCONint2.Enabled = False
      mskTermino.Enabled = False
      'cmdCONintervalo.Enabled = False
      'lblCONnome.Enabled = False
      'cboNome.Enabled = False
      'cmdCONnome.Enabled = False
      'optVencimento.Enabled = False
      'optPagamento.Enabled = False
      'chkPgtoMes.Enabled = False
      'chkPgtoOutros.Enabled = False
      
      'MOSTRAR OS DADOS
      'If cboCONForma.Text = "TODOS" Then
      '   sSQL = "SELECT *, (valor_antigo - ISNULL(haveres, 0)) AS total, ISNULL(haveres, 0) AS hav " & _
      '      "FROM a_pagar WHERE (status = '" & cboCONStatus.Text & "') AND (setor = '" & cboCONsetor.Text & "') " & _
      '      "ORDER BY " & INDICE
      'Else
      '   sSQL = "SELECT *, (valor_antigo - ISNULL(haveres, 0)) AS total, ISNULL(haveres, 0) AS hav " & _
      '      "FROM a_pagar WHERE (status = '" & cboCONStatus.Text & "') AND (setor = '" & cboCONsetor.Text & "') " & _
      '      "AND (forma = '" & cboCONForma.Text & "') ORDER BY " & INDICE
      'End If
      
   ElseIf cboConsulta.Text = "MENSAL" Then
      If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
    
      'INDICE SOMENTE PARA CONSULTA DE DATAS
      'If optVencimento.Value = True Then
      '   INDICE = "vencimento"
      'ElseIf optPagamento.Value = True Then
      '   INDICE = "data_pgto"
      'End If
      
      'DATA DE VENCIMENTO OU PAGAMENTO
      'If optVencimento.Value = True Then
      '   Tipo_Data = "(MONTH(vencimento) = " & cboMes.ListIndex + 1 & ") AND (YEAR(vencimento) = " & cboAno & ")"
      'ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 0 And chkPgtoMes.Value = 0 Then
      '   Tipo_Data = "(MONTH(data_pgto) = " & cboMes.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ")"
      'ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 1 And chkPgtoMes.Value = 0 Then
      '   Tipo_Data = "(MONTH(data_pgto) = " & cboMes.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ") AND (MONTH(vencimento) <> " & cboMes.ListIndex + 1 & ")"
      'ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 0 And chkPgtoMes.Value = 1 Then
      '   Tipo_Data = "(MONTH(data_pgto) = " & cboMes.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ") AND (MONTH(vencimento) = " & cboMes.ListIndex + 1 & " AND (YEAR(vencimento) = " & cboAno & ")"
      'Else
      '   Tipo_Data = "(MONTH(vencimento) = " & cboMes.ListIndex + 1 & ") AND (YEAR(vencimento) = " & cboAno & ")"
      'End If
      
      'MOSTRAR OS DADOS
      'If cboCONForma.Text = "TODOS" Then
      '   sSQL = "SELECT *, (valor_antigo - ISNULL(haveres, 0)) AS total, ISNULL(haveres, 0) AS hav " & _
      '      "FROM a_pagar WHERE " & Tipo_Data & " AND (status = '" & cboCONStatus.Text & "') AND (setor = '" & cboCONsetor.Text & "') " & _
      '      "ORDER BY " & INDICE
      'Else
      '   sSQL = "SELECT *, (valor_antigo - ISNULL(haveres, 0)) AS total, ISNULL(haveres, 0) AS hav " & _
      '      "FROM a_pagar WHERE " & Tipo_Data & " AND (status = '" & cboCONStatus.Text & "') AND (setor = '" & cboCONsetor.Text & "') " & _
      '      "AND (forma = '" & cboCONForma.Text & "') ORDER BY " & INDICE
      'End If
      
   ElseIf cboConsulta.Text = "PERÍODO" Then
      If mskInicio.Text = "" Or mskTermino.Text = "" Then Exit Sub
      If Not IsDate(mskInicio.Text) Or Not IsDate(mskTermino.Text) Then Exit Sub
      
      If mskInicio.Text = "" Or mskTermino.Text = "" Then
         ShowMsg "Digite a DATA INICIAL e DATA FINAL!", vbExclamation
         mskInicio.SetFocus
         Exit Sub
      End If
      
      'MOSTRAR OS DADOS
      'If cboCONForma.Text = "TODOS" Then
      '   sSQL = "SELECT *, (valor_antigo - ISNULL(haveres, 0)) AS total, ISNULL(haveres, 0) AS hav " & _
      '      "FROM a_pagar WHERE (status = '" & cboCONStatus.Text & "') AND (setor = '" & cboCONsetor.Text & "') " & _
      '      "AND (vencimento >= '" & Format(Mask1, ocDATA_EUA) & "') AND (vencimento <= '" & Format(Mask2, ocDATA_EUA) & "') " & _
      '      "ORDER BY " & INDICE
      'Else
      '   sSQL = "SELECT *, (valor_antigo - ISNULL(haveres, 0)) AS total, ISNULL(haveres, 0) AS hav " & _
      '      "FROM a_pagar WHERE (status = '" & cboCONStatus.Text & "') AND (setor = '" & cboCONsetor.Text & "') " & _
      '      "AND (forma = '" & cboCONForma.Text & "') AND (vencimento >= '" & Format(Mask1, ocDATA_EUA) & "') " & _
      '      "AND (vencimento <= '" & Format(Mask2, ocDATA_EUA) & "') ORDER BY " & INDICE
      'End If
           
   'ElseIf optNome.Value = True Then
   '   If cboNome.Text = "" Then Exit Sub
   '
   '   'MOSTRAR OS DADOS
   '   If cboCONForma.Text = "TODOS" Then
   '      sSQL = "SELECT *, (valor_antigo - ISNULL(haveres, 0)) AS total, ISNULL(haveres, 0) AS hav " & _
   '         "FROM a_pagar WHERE (status = '" & cboCONStatus.Text & "') AND (setor = '" & cboCONsetor.Text & "') " & _
   '         "AND (favorecido = '" & cboNome.Text & "') ORDER BY " & INDICE
   '   Else
   '      sSQL = "SELECT *, (valor_antigo - ISNULL(haveres, 0)) AS total, ISNULL(haveres, 0) AS hav " & _
   '         "FROM a_pagar WHERE (status = '" & cboCONStatus.Text & "') AND (setor = '" & cboCONsetor.Text & "') " & _
   '         "AND (forma = '" & cboCONForma.Text & "') AND (favorecido = '" & cboNome.Text & "') ORDER BY " & INDICE
   '   End If
   End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   
   'CONTAR REGISTROS - CONTAS
   'txtCONquant.Text = Format(totalRegistros, "00")
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
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
   
   'For i = ANO To FirstYear Step -1
   '   cboAno.AddItem i
   'Next
   '
   'iAno = iAno + 1
   'For i = iAno To LastYear
   '   cboAno.AddItem i
   'Next
   
   moCombo.AttachTo cboAno
End Sub

Private Sub cboConsulta_Click()
cboConsulta_LostFocus
End Sub

Private Sub cboConsulta_GotFocus()
cboConsulta.Clear
PreencherConsulta
moCombo.AttachTo cboConsulta
End Sub

Private Sub cboConsulta_LostFocus()
If cboConsulta.Text = "TODOS" Then
    lblMes.Visible = False
    lblAno.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
    lblInicio.Visible = False
    lblTermino.Visible = False
    mskInicio.Visible = False
    mskTermino.Visible = False
    cmdConsData.Visible = False
    cmdConsData2.Visible = False
    cmdExibir_Click
ElseIf cboConsulta.Text = "MENSAL" Then
    lblMes.Visible = True
    lblAno.Visible = True
    cboMes.Visible = True
    cboAno.Visible = True
    lblInicio.Visible = False
    lblTermino.Visible = False
    mskInicio.Visible = False
    mskTermino.Visible = False
    cmdConsData.Visible = False
    cmdConsData2.Visible = False
    If cboMes.Visible = True Then cboMes.SetFocus
ElseIf cboConsulta.Text = "PERÍODO" Then
    lblMes.Visible = False
    lblAno.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
    lblInicio.Visible = True
    lblTermino.Visible = True
    mskInicio.Visible = True
    mskTermino.Visible = True
    cmdConsData.Visible = True
    cmdConsData2.Visible = True
    mskInicio.SetFocus
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

mskTermino = Format(varData, "dd/mm/yy")   'Exibe a data no campo
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

mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
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

mskTermino = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdDetalhar_Click()
i = Grid.Row
frmDetalhamento.Visible = True

If Grid.TextMatrix(i, 1) = "" Then Exit Sub

sSQL = "SELECT * FROM caixa_saldo_retirada WHERE (COD_SALDO = " & Grid.TextMatrix(i, 1) & ") ORDER BY codigo;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Retirada r

If r.State <> 0 Then r.Close
Set r = Nothing
'End If
End Sub

Private Sub cmdFecharDet_Click()
frmDetalhamento.Visible = False
End Sub


Private Sub cmdImprimir_Click()
   'colocar o nome da maquina na barra de status
   Dim var_Impressora As String
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   If cboConsulta.Text = "TODOS" Then
      sSQL = "SELECT data, saldo, (SELECT ISNULL(SUM(valor), 0) FROM caixa_saldo_retirada WHERE (data = t0.data)) AS soma_retirada " & _
         "FROM caixa_dia AS t0 ORDER BY data;"
      
   ElseIf cboConsulta.Text = "MENSAL" Then
      If cboMes.Text = "" Then Exit Sub
      sSQL = "SELECT data, saldo, (SELECT ISNULL(SUM(valor), 0) FROM caixa_saldo_retirada WHERE (data = t0.data)) AS soma_retirada " & _
         "FROM caixa_dia AS t0 WHERE (MONTH(data) = " & cboMes.Text & ") AND (YEAR(data) = " & cboAno & ") ORDER BY data;"
      
   ElseIf cboConsulta.Text = "PERÍODO" Then
      If mskInicio.Text = "" Or mskTermino.Text = "" Then Exit Sub
      
      sSQL = "SELECT data, saldo, (SELECT ISNULL(SUM(valor), 0) FROM caixa_saldo_retirada WHERE (data = t0.data)) AS soma_retirada " & _
         "FROM caixa_dia AS t0 WHERE (data >= CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103)) AND (data <= CONVERT(DATETIME, '" & Format(mskTermino, ocDATA) & "', 103)) ORDER BY data;"
      
   End If
   
   Me.Hide
   
   Set r = dbData.OpenRecordset(printSQL)
   
   Set REL_Lanc_Caixa.Relatorio.Recordset = r
   'REL_Lanc_Caixa.dfData.Caption = Format(txtData.Text, "dd/mm/yy")
   'REL_Lanc_Caixa.dfSaldo.Caption = txtSaldoAtual.Text
   'REL_Lanc_Caixa.dfRelData.Caption = Format(txtConsData.Text, "dd/mm/yy")
   'REL_Lanc_Caixa.dfRelSaldo.Caption = txtConsSaldo.Text
   
   '.TextMatrix(.Rows - 1, 2) = IIf(IsNull(RS!Data) = True, "", Format(RS!Data, "dd/mm/yy"))
   '.TextMatrix(.Rows - 1, 3) = Format(Saldo_Anterior, "##,##0.00")
   '.TextMatrix(.Rows - 1, 4) = Format(IIf(IsNull(RS!SALDO) = True, 0, RS!SALDO), "##,##0.00")
   '.TextMatrix(.Rows - 1, 5) = Format(IIf(IsNull(RS!Soma_retirada) = True, 0, RS!Soma_retirada), "##,##0.00")
   '.TextMatrix(.Rows - 1, 6) = Format(CDbl(.TextMatrix(.Rows - 1, 3)) + CDbl(.TextMatrix(.Rows - 1, 4)) - CDbl(.TextMatrix(.Rows - 1, 5)), "##,##0.00")
   
   REL_Lanc_Caixa.ReportField1.Campo = "r.data"
   REL_Lanc_Caixa.ReportField5.Campo = "rs.Saldo_Anterior"
   REL_Lanc_Caixa.ReportField2.Campo = "SALDO"
   REL_Lanc_Caixa.ReportField3.Campo = "Soma_retirada"
   REL_Lanc_Caixa.ReportField4.Campo = "SALDO_ANTERIOR"

   REL_Lanc_Caixa.ReportField1.Formato = "dd/mm/yy"
   'REL_Lanc_Caixa.Relatorio.NomeImpressora = var_Impressora
   REL_Lanc_Caixa.Relatorio.Ativar
   Unload REL_Lanc_Caixa
   
   Me.Show 1
End Sub



Private Sub cmdExibir_Click()
If cboConsulta.Text = "TODOS" Then
    sSQL = "SELECT * FROM caixa_saldo ORDER BY codigo desc;"
         
ElseIf cboConsulta.Text = "MENSAL" Then
   If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
   
   sSQL = "SELECT * " & _
      "FROM caixa_saldo WHERE (MONTH(DATA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(DATA) = " & cboAno & ") ORDER BY codigo desc;"

ElseIf cboConsulta.Text = "PERÍODO" Then
   If mskInicio.Text = "" Or mskTermino.Text = "" Then Exit Sub
   
   sSQL = "SELECT * " & _
      "FROM caixa_saldo WHERE (DATA >= CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103)) AND (DATA <= CONVERT(DATETIME, '" & Format(mskTermino, ocDATA) & "', 103)) ORDER BY codigo desc;"

End If

Set r = dbData.OpenRecordset(sSQL)
'Debug.Print sSQL

FormatarGrid r

'FormatarGrid_Lancamentos r
'If r.State <> 0 Then r.Close
'Set r = Nothing

'printSQL = sSQL

'======== CODIGO NOVO


'If cboConsulta.Text = "TODOS" Then
   
         
'ElseIf cboConsulta.Text = "MENSAL" Then
'   If cboMES.Text = "" Then Exit Sub
   
'   sSQL = "SELECT data_abertura, saldo, (SELECT ISNULL(SUM(valor), 0) FROM caixa_saldo_retirada WHERE (data = t0.data_abertura)) AS soma_retirada " & _
'      "FROM caixa_dia AS t0 WHERE (MONTH(data_abertura) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_abertura) = " & cboAno & ") ORDER BY data_abertura;"

'ElseIf cboConsulta.Text = "PERÍODO" Then
'   If mskInicio.Text = "" Or mskTermino.Text = "" Then Exit Sub
   
'   sSQL = "SELECT data_abertura, saldo, (SELECT ISNULL(SUM(valor), 0) FROM caixa_saldo_retirada WHERE (data = t0.data_abertura)) AS soma_retirada " & _
'      "FROM caixa_dia AS t0 WHERE (data_abertura >= CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103)) AND (data_abertura <= CONVERT(DATETIME, '" & Format(mskTermino, ocDATA) & "', 103)) ORDER BY data_abertura;"

'End If

'Set r = dbData.OpenRecordset(sSQL)

'FormatarGrid r


'mostrar os totais
sSQL = "SELECT TOP 1 codigo, DATA, SALDO_ATUAL, CODCAIXA, CAIXA FROM caixa_saldo ORDER BY codigo desc;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    txtConsData.Text = r("DATA")
    txtUltimaCaixa.Text = ValidateNull(r("CAIXA"))
    txtUltimoCodCaixa.Text = ValidateNull(r("CODCAIXA"))
    txtConsSaldo.Text = Format(r("SALDO_ATUAL"), ocMONEY)
End If

If r.State <> 0 Then r.Close
Set r = Nothing

End Sub

Private Sub cmdMostrarCaixa_Click()
varFluxoCaixa = True
varFluxoNomeCaixa = Grid.TextMatrix(Grid.Row, 7)
varFluxoCodCaixa = Grid.TextMatrix(Grid.Row, 8)
varFluxoCaixaSituacao = "FECHADO"
varFluxoCaixaData = Format(Grid.TextMatrix(Grid.Row, 2), "dd/mm/yy")
varCodCaixa = Grid.TextMatrix(Grid.Row, 8)
Caixa_Controle_semOS.cmdAbrirCaixa.Visible = True
Caixa_Controle_semOS.cmdAbrirCaixa.Enabled = False
Caixa_Controle_semOS.cmdFecharCaixa.Enabled = False
Caixa_Controle_semOS.cmdTroco.Enabled = False
Caixa_Controle_semOS.cmdImprimir.Enabled = True
Caixa_Controle_semOS.cmdMostrar_Click
Caixa_Controle_semOS.Show
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")

PreencherConsulta
cboConsulta.ListIndex = 1

'cmdExibir_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
   mskInicio.Mask = "##/##/##"
End Sub
Private Sub mskTermino_KeyPress(KeyAscii As Integer)
   mskTermino.Mask = "##/##/##"
End Sub

Private Sub mskTermino_LostFocus()
   cmdExibir.SetFocus
End Sub

