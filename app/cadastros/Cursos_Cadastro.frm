VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Cursos_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CRONOGRAMA"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "Cursos_Cadastro.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   7425
      TabIndex        =   88
      Top             =   60
      Width           =   7455
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   8940
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CRONOGRAMA"
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
         Left            =   1380
         TabIndex        =   90
         Top             =   240
         Width           =   2310
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   300
         Picture         =   "Cursos_Cadastro.frx":23D2
         Top             =   -60
         Width           =   960
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   60
      TabIndex        =   54
      Top             =   1020
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabMaxWidth     =   2593
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CURSOS"
      TabPicture(0)   =   "Cursos_Cadastro.frx":4D61
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdExcluirCurso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAlterarCurso"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNovoCurso"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelarCurso"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSalvarCurso"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmPrincipalCurso"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmSecundario"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "PACOTES"
      TabPicture(1)   =   "Cursos_Cadastro.frx":4D7D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCancelarPac"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdSalvarPac"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdExcluirPac"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAlterarPac"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdNovoPac"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "frmPrincipalPacote"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "frmSecundarioPacote"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "SALAS"
      TabPicture(2)   =   "Cursos_Cadastro.frx":4D99
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdExcluirSala"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdAlterarSala"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdNovoSala"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdCancelarSala"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdSalvarSala"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "frmPrincipalSala"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame5"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "TEMPORADAS"
      TabPicture(3)   =   "Cursos_Cadastro.frx":4DB5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdExcluirTemporada"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdAlterarTemporada"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdNovoTemporada"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdCancelarTemporada"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdSalvarTemporada"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame4"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "frmPrincipalTemporada"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "HORÁRIOS"
      TabPicture(4)   =   "Cursos_Cadastro.frx":4DD1
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "cmdExcluirHorario"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdAlteraHorario"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdNovoHorario"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmdCancelarHorario"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdSalvarHorario"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "frmCadHorario"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Frame8"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).ControlCount=   7
      Begin VB.Frame Frame8 
         Height          =   3975
         Left            =   120
         TabIndex        =   108
         Top             =   3540
         Width           =   7215
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            ScaleHeight     =   345
            ScaleWidth      =   6945
            TabIndex        =   110
            Top             =   3480
            Width           =   6975
            Begin VB.OptionButton optSala 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Sala"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   4920
               TabIndex        =   116
               Top             =   60
               Width           =   855
            End
            Begin VB.OptionButton optPacote 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Pacote"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   960
               TabIndex        =   114
               Top             =   60
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optTemporada 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Temporada"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1860
               TabIndex        =   113
               Top             =   60
               Width           =   1215
            End
            Begin VB.OptionButton optDias 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Dias"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3120
               TabIndex        =   112
               Top             =   60
               Width           =   735
            End
            Begin VB.OptionButton optHorario 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Horário"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3900
               TabIndex        =   111
               Top             =   60
               Width           =   855
            End
            Begin VB.Label Label21 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Ordem:"
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
               Height          =   195
               Left            =   120
               TabIndex        =   115
               Top             =   60
               Width           =   615
            End
         End
         Begin MSFlexGridLib.MSFlexGrid GridHorario 
            Height          =   3255
            Left            =   120
            TabIndex        =   109
            Top             =   180
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5741
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame frmCadHorario 
         Enabled         =   0   'False
         Height          =   3135
         Left            =   120
         TabIndex        =   91
         Top             =   420
         Width           =   7215
         Begin VB.TextBox txtDias 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4860
            TabIndex        =   107
            Top             =   1560
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.TextBox txtCodHorario 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   0
            TabIndex        =   106
            Top             =   -60
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cboHORPacote 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   43
            Top             =   420
            Width           =   1995
         End
         Begin VB.ComboBox cboHORSala 
            Height          =   315
            Left            =   5640
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   420
            Width           =   1455
         End
         Begin VB.ComboBox cboHORTemporada 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   44
            Top             =   420
            Width           =   3435
         End
         Begin VB.Frame Frame7 
            Caption         =   "Dias"
            Height          =   675
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   6975
            Begin VB.CheckBox chkDias 
               Caption         =   "Seg."
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
               Index           =   0
               Left            =   120
               TabIndex        =   101
               Tag             =   "2ª"
               Top             =   300
               Width           =   735
            End
            Begin VB.CheckBox chkDias 
               Caption         =   "Ter."
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
               Left            =   960
               TabIndex        =   100
               Tag             =   "3ª"
               Top             =   300
               Width           =   735
            End
            Begin VB.CheckBox chkDias 
               Caption         =   "Qua."
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
               Index           =   2
               Left            =   1800
               TabIndex        =   99
               Tag             =   "4ª"
               Top             =   300
               Width           =   735
            End
            Begin VB.CheckBox chkDias 
               Caption         =   "Qui."
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
               Index           =   3
               Left            =   2700
               TabIndex        =   98
               Tag             =   "5ª"
               Top             =   300
               Width           =   735
            End
            Begin VB.CheckBox chkDias 
               Caption         =   "Sex."
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
               Index           =   4
               Left            =   3600
               TabIndex        =   97
               Tag             =   "6ª"
               Top             =   300
               Width           =   735
            End
            Begin VB.CheckBox chkDias 
               Caption         =   "Sáb."
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
               Index           =   5
               Left            =   4440
               TabIndex        =   96
               Tag             =   "Sáb"
               Top             =   300
               Width           =   735
            End
            Begin VB.CheckBox chkDias 
               Caption         =   "Dom."
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
               Index           =   6
               Left            =   5220
               TabIndex        =   95
               Tag             =   "Dom"
               Top             =   300
               Width           =   795
            End
         End
         Begin VB.TextBox txtCodHORPacote 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1500
            TabIndex        =   94
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtCodHORTemporada 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            TabIndex        =   93
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtCodHORSala 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6420
            TabIndex        =   92
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSMask.MaskEdBox mskHorario 
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   1800
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pacote"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   180
            Width           =   510
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sala"
            Height          =   195
            Left            =   5640
            TabIndex        =   104
            Top             =   180
            Width           =   315
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Temporada"
            Height          =   195
            Left            =   2160
            TabIndex        =   103
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horário"
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   1560
            Width           =   510
         End
      End
      Begin VB.Frame frmPrincipalTemporada 
         Enabled         =   0   'False
         Height          =   915
         Left            =   -74880
         TabIndex        =   79
         Top             =   360
         Width           =   7215
         Begin VB.ComboBox cboEtapa 
            Height          =   315
            Left            =   1800
            TabIndex        =   35
            Top             =   420
            Width           =   1635
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   3480
            TabIndex        =   36
            Top             =   420
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   420
            Width           =   1635
         End
         Begin VB.TextBox txtCodTemporada 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6480
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSMask.MaskEdBox mskTermino 
            Height          =   315
            Left            =   5220
            TabIndex        =   37
            Top             =   420
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdCalendario1 
            Height          =   315
            Left            =   4800
            TabIndex        =   117
            Top             =   420
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
            MICON           =   "Cursos_Cadastro.frx":4DED
            PICN            =   "Cursos_Cadastro.frx":4E09
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
            Left            =   6540
            TabIndex        =   118
            Top             =   420
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
            MICON           =   "Cursos_Cadastro.frx":71EB
            PICN            =   "Cursos_Cadastro.frx":7207
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
            BackStyle       =   0  'Transparent
            Caption         =   "Etapa"
            Height          =   195
            Left            =   1800
            TabIndex        =   84
            Top             =   180
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Termino"
            Height          =   195
            Left            =   5220
            TabIndex        =   83
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Left            =   3540
            TabIndex        =   82
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   180
            Width           =   285
         End
      End
      Begin VB.Frame Frame4 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   77
         Top             =   1320
         Width           =   7215
         Begin MSFlexGridLib.MSFlexGrid GridTemporada 
            Height          =   5835
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   10292
            _Version        =   393216
            SelectionMode   =   1
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   76
         Top             =   1320
         Width           =   7215
         Begin MSFlexGridLib.MSFlexGrid GridSala 
            Height          =   5835
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   10292
            _Version        =   393216
            SelectionMode   =   1
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "PACOTES CADASTRADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         TabIndex        =   75
         Top             =   6000
         Width           =   7215
         Begin MSFlexGridLib.MSFlexGrid GridPacote 
            Height          =   1815
            Left            =   60
            TabIndex        =   24
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   3201
            _Version        =   393216
            SelectionMode   =   1
         End
      End
      Begin VB.Frame frmSecundarioPacote 
         Enabled         =   0   'False
         Height          =   4095
         Left            =   -74880
         TabIndex        =   72
         Top             =   1260
         Width           =   7215
         Begin VB.Frame Frame3 
            Caption         =   "NESSE PACOTE:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3795
            Left            =   4740
            TabIndex        =   74
            Top             =   180
            Width           =   2415
            Begin MSFlexGridLib.MSFlexGrid GridCursoADDPacote 
               Height          =   3435
               Left            =   60
               TabIndex        =   19
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   6059
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "TODOS OS CURSOS:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3795
            Left            =   60
            TabIndex        =   73
            Top             =   180
            Width           =   2415
            Begin MSFlexGridLib.MSFlexGrid GridCursoPacote 
               Height          =   3495
               Left            =   60
               TabIndex        =   17
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   6165
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionaCurso 
            Height          =   555
            Left            =   2700
            TabIndex        =   18
            Top             =   1260
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   979
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
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Cursos_Cadastro.frx":95E9
            PICN            =   "Cursos_Cadastro.frx":9605
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdRemoverCursos 
            Height          =   555
            Left            =   2700
            TabIndex        =   20
            Top             =   2400
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   979
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
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Cursos_Cadastro.frx":9B69
            PICN            =   "Cursos_Cadastro.frx":9B85
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
      Begin VB.Frame frmPrincipalPacote 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   -74880
         TabIndex        =   66
         Top             =   360
         Width           =   7215
         Begin VB.TextBox txtParc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6000
            TabIndex        =   16
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtQuant 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5160
            TabIndex        =   15
            Top             =   480
            Width           =   795
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            TabIndex        =   14
            Top             =   480
            Width           =   1275
         End
         Begin VB.TextBox txtCodPac 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtPacote 
            Height          =   315
            Left            =   120
            MaxLength       =   19
            TabIndex        =   11
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox cboTipoDur 
            Height          =   315
            Left            =   1980
            TabIndex        =   12
            Top             =   480
            Width           =   1035
         End
         Begin VB.TextBox txtDuracao 
            Height          =   315
            Left            =   3060
            MaxLength       =   2
            TabIndex        =   13
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parcela:"
            Height          =   195
            Left            =   6000
            TabIndex        =   87
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant."
            Height          =   195
            Left            =   5160
            TabIndex        =   86
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   3840
            TabIndex        =   85
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Dur."
            Height          =   195
            Left            =   1980
            TabIndex        =   70
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pacote:"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duração:"
            Height          =   195
            Left            =   3060
            TabIndex        =   68
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame frmPrincipalSala 
         Height          =   915
         Left            =   -74880
         TabIndex        =   62
         Top             =   360
         Width           =   7215
         Begin VB.TextBox txtSala 
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   420
            Width           =   1515
         End
         Begin VB.TextBox txtCodSala 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2580
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtVagas 
            Height          =   315
            Left            =   1680
            TabIndex        =   27
            Top             =   420
            Width           =   975
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sala:"
            Height          =   315
            Left            =   120
            TabIndex        =   65
            Top             =   180
            Width           =   360
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vagas"
            Height          =   195
            Left            =   1680
            TabIndex        =   64
            Top             =   180
            Width           =   450
         End
      End
      Begin VB.Frame frmSecundario 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   61
         Top             =   1260
         Width           =   7215
         Begin MSFlexGridLib.MSFlexGrid GridCursos 
            Height          =   6015
            Left            =   60
            TabIndex        =   7
            Top             =   180
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   10610
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin VB.Frame frmPrincipalCurso 
         Enabled         =   0   'False
         Height          =   915
         Left            =   -74880
         TabIndex        =   55
         Top             =   360
         Width           =   7215
         Begin VB.TextBox txtCurso 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   420
            Width           =   2835
         End
         Begin VB.ComboBox cboClassificacao 
            Height          =   315
            Left            =   3000
            TabIndex        =   2
            Top             =   420
            Width           =   1875
         End
         Begin VB.TextBox txtCodigoCurso 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6600
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtCarga 
            Height          =   315
            Left            =   4920
            TabIndex        =   3
            Top             =   420
            Width           =   975
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5940
            TabIndex        =   4
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Curso:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Classificação:"
            Height          =   195
            Left            =   3000
            TabIndex        =   59
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carga Hor."
            Height          =   195
            Left            =   4920
            TabIndex        =   58
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Left            =   5940
            TabIndex        =   57
            Top             =   180
            Width           =   360
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarCurso 
         Height          =   555
         Left            =   -74880
         TabIndex        =   5
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":A105
         PICN            =   "Cursos_Cadastro.frx":A121
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarCurso 
         Height          =   555
         Left            =   -73140
         TabIndex        =   6
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":109EB
         PICN            =   "Cursos_Cadastro.frx":10A07
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovoCurso 
         Height          =   555
         Left            =   -69360
         TabIndex        =   0
         Top             =   7620
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":174AB
         PICN            =   "Cursos_Cadastro.frx":174C7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterarCurso 
         Height          =   555
         Left            =   -74880
         TabIndex        =   8
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":181A1
         PICN            =   "Cursos_Cadastro.frx":181BD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluirCurso 
         Height          =   555
         Left            =   -73140
         TabIndex        =   9
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":18A97
         PICN            =   "Cursos_Cadastro.frx":18AB3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovoPac 
         Height          =   555
         Left            =   -69420
         TabIndex        =   10
         Top             =   5400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":18DCD
         PICN            =   "Cursos_Cadastro.frx":18DE9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterarPac 
         Height          =   555
         Left            =   -74880
         TabIndex        =   52
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":19AC3
         PICN            =   "Cursos_Cadastro.frx":19ADF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluirPac 
         Height          =   555
         Left            =   -73140
         TabIndex        =   23
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":1A3B9
         PICN            =   "Cursos_Cadastro.frx":1A3D5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarPac 
         Height          =   555
         Left            =   -74880
         TabIndex        =   21
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":1A6EF
         PICN            =   "Cursos_Cadastro.frx":1A70B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarPac 
         Height          =   555
         Left            =   -73140
         TabIndex        =   22
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":20FD5
         PICN            =   "Cursos_Cadastro.frx":20FF1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarSala 
         Height          =   555
         Left            =   -74880
         TabIndex        =   28
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":27A95
         PICN            =   "Cursos_Cadastro.frx":27AB1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarSala 
         Height          =   555
         Left            =   -73140
         TabIndex        =   29
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":2E37B
         PICN            =   "Cursos_Cadastro.frx":2E397
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovoSala 
         Height          =   555
         Left            =   -69360
         TabIndex        =   25
         Top             =   7620
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":34E3B
         PICN            =   "Cursos_Cadastro.frx":34E57
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterarSala 
         Height          =   555
         Left            =   -74880
         TabIndex        =   31
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":35B31
         PICN            =   "Cursos_Cadastro.frx":35B4D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluirSala 
         Height          =   555
         Left            =   -73140
         TabIndex        =   32
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":36427
         PICN            =   "Cursos_Cadastro.frx":36443
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarTemporada 
         Height          =   555
         Left            =   -74880
         TabIndex        =   38
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":3675D
         PICN            =   "Cursos_Cadastro.frx":36779
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarTemporada 
         Height          =   555
         Left            =   -73140
         TabIndex        =   39
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":3D043
         PICN            =   "Cursos_Cadastro.frx":3D05F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovoTemporada 
         Height          =   555
         Left            =   -69360
         TabIndex        =   33
         Top             =   7620
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":43B03
         PICN            =   "Cursos_Cadastro.frx":43B1F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterarTemporada 
         Height          =   555
         Left            =   -74880
         TabIndex        =   40
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":447F9
         PICN            =   "Cursos_Cadastro.frx":44815
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluirTemporada 
         Height          =   555
         Left            =   -73140
         TabIndex        =   41
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":450EF
         PICN            =   "Cursos_Cadastro.frx":4510B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarHorario 
         Height          =   555
         Left            =   120
         TabIndex        =   48
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":45425
         PICN            =   "Cursos_Cadastro.frx":45441
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarHorario 
         Height          =   555
         Left            =   1860
         TabIndex        =   49
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cursos_Cadastro.frx":4BD0B
         PICN            =   "Cursos_Cadastro.frx":4BD27
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovoHorario 
         Height          =   555
         Left            =   5640
         TabIndex        =   42
         Top             =   7620
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":527CB
         PICN            =   "Cursos_Cadastro.frx":527E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlteraHorario 
         Height          =   555
         Left            =   120
         TabIndex        =   50
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":534C1
         PICN            =   "Cursos_Cadastro.frx":534DD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluirHorario 
         Height          =   555
         Left            =   1860
         TabIndex        =   51
         Top             =   7620
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         MICON           =   "Cursos_Cadastro.frx":53DB7
         PICN            =   "Cursos_Cadastro.frx":53DD3
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
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   555
      Left            =   5820
      TabIndex        =   53
      Top             =   9360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Cursos_Cadastro.frx":540ED
      PICN            =   "Cursos_Cadastro.frx":54109
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   71
      Top             =   10035
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9419
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "13:06"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
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
End
Attribute VB_Name = "Cursos_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Private Sub Calcular_Parcelas()
If txtTotal.Text = "" Or txtQuant.Text = "" Then
    Exit Sub
Else
    Dim QtdaParc As Integer
    Dim Total As Currency
    Dim RESULTADO As Currency
    
    Total = txtTotal
    QtdaParc = txtQuant
    RESULTADO = Total / QtdaParc
    txtParc.Text = Format(RESULTADO, "##,##0.00")
End If

txtTotal.SelStart = 0
txtTotal.SelLength = Len(txtTotal)
End Sub

Private Sub Limpar_Objetos_Modulos()
'txtCodigoModulo.Text = ""
'cboModulo.Text = ""
'txtMatricula.Text = ""
'txtTotal.Text = ""
'txtQuantParc.Text = ""
'txtParc.Text = ""
End Sub




Private Sub Limpar_Objetos_Pacotes()
txtCodPac.Text = ""
txtPacote.Text = ""
txtDuracao.Text = ""
cboTipoDur.Text = ""
txtTotal.Text = ""
txtParc.Text = ""
txtQuant.Text = ""
End Sub




Private Sub LimparGrid_CursoADDPacote()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From PACOTES_ITENS WHERE 1=0"
   Set r = dbData.OpenRecordset(sSQL)
   Debug.Print sSQL
   'Mostra os dados no grid
   FormatarGrid_CursoADDPacote r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Limpar_Objetos_Horario()
If cmdAlteraHorario.Visible = False Then txtCodHorario.Text = ""
txtCodHORPacote.Text = ""
txtCodHORTemporada.Text = ""
txtCodHORSala.Text = ""
cboHORPacote.Text = ""
cboHORTemporada.Text = ""
cboHORSala.Text = ""
txtDias.Text = ""
mskHorario.Mask = ""
mskHorario.Text = ""

Dim i As Integer
    For i = 0 To 6
        chkDias(i).Value = Unchecked
    Next i
End Sub

Private Sub MontarGrid_Temporada()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From TEMPORADAS ORDER BY ANO, ETAPA"
   Set r = dbData.OpenRecordset(sSQL)
   
   'Mostra os dados no grid
   FormatarGrid_Temporada r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub MontarGrid_Sala()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From SALAS ORDER BY SALA"
   Set r = dbData.OpenRecordset(sSQL)
   
   'Mostra os dados no grid
   FormatarGrid_Sala r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub MontarGrid_Pacote()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From PACOTES ORDER BY PACOTE"
   Set r = dbData.OpenRecordset(sSQL)
   
   'Mostra os dados no grid
   FormatarGrid_Pacote r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub MontarGrid_CursoPacote()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From CURSOS ORDER BY CURSO"
   Set r = dbData.OpenRecordset(sSQL)
   
   'Mostra os dados no grid
   FormatarGrid_CursoPacote r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub MontarGrid_CursoADDPacote()
If txtCodPac.Text = "" Then Exit Sub
    
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From PACOTES_ITENS WHERE COD_PACOTE = " & txtCodPac.Text & " ORDER BY CURSO"
   Set r = dbData.OpenRecordset(sSQL)

   FormatarGrid_CursoADDPacote r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub MontarGrid_Horario()
  'Indice
    Dim INDICE As String
    If optPacote.Value = True Then
        INDICE = "PACOTES.CODIGO"
    ElseIf optTemporada.Value = True Then
        INDICE = "TEMPORADAS.CODIGO"
    ElseIf optDias.Value = True Then
        INDICE = "HORARIO.DIAS"
    ElseIf optHorario.Value = True Then
        INDICE = "HORARIO.HORARIO"
    ElseIf optSala.Value = True Then
        INDICE = "SALAS.CODIGO"
    Else
        optPacote.Value = True
        INDICE = "PACOTES.CODIGO"
    End If

   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT HORARIO.CODIGO AS var_CODIGO, HORARIO.COD_PACOTE AS varCodPac, HORARIO.COD_TEMPORADA as varCodTem, HORARIO.COD_SALA as varCodSal, HORARIO.DIAS as varDias, HORARIO.HORARIO as varHor, PACOTES.CODIGO, PACOTES.PACOTE AS VARPAC, TEMPORADAS.CODIGO, TEMPORADAS.ANO AS VARANO, TEMPORADAS.ETAPA AS VARETAPA, SALAS.CODIGO, SALAS.SALA AS VARSAL FROM SALAS INNER JOIN (TEMPORADAS INNER JOIN (PACOTES INNER JOIN HORARIO ON PACOTES.CODIGO = HORARIO.COD_PACOTE) ON (TEMPORADAS.CODIGO = HORARIO.COD_TEMPORADA)) ON SALAS.CODIGO = HORARIO.COD_SALA ORDER BY " & INDICE
   Set r = dbData.OpenRecordset(sSQL)
   
   'Mostra os dados no grid
   FormatarGrid_Horario r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub MontarGrid_Cursos()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM cursos ORDER BY CURSO;"
   Set r = dbData.OpenRecordset(sSQL)
   
   'Mostra os dados no grid
   FormatarGrid_Cursos r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub MostrarPacote()
If txtCodHORPacote.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From PACOTES where CODIGO = " & txtCodHORPacote.Text & ""
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.EOF Then
       If Not IsNull(r!PACOTE) Then cboHORPacote.Text = r!PACOTE
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub MostrarTemporada()
If txtCodHORTemporada.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From TEMPORADAS where CODIGO = " & txtCodHORTemporada.Text & ""
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.EOF Then
       If Not IsNull(r!ANO) Then cboHORTemporada.Text = r!ANO & "/" & r!ETAPA & "  -> " & Format(r!INICIO, "dd/mm/yy") & " -> " & Format(r!TERMINO, "dd/mm/yy")
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub MostrarSala()
If txtCodHORSala.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From SALAS where CODIGO = " & txtCodHORSala.Text & ""
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.EOF Then
       If Not IsNull(r!SALA) Then cboHORSala.Text = r!SALA
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub MostrarHorario()
If txtCodHorario.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "Select * From HORARIO where CODIGO = " & txtCodHorario.Text & ""
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.EOF Then
       If Not IsNull(r!COD_PACOTE) Then txtCodHORPacote.Text = r!COD_PACOTE
       If Not IsNull(r!COD_TEMPORADA) Then txtCodHORTemporada.Text = r!COD_TEMPORADA
       If Not IsNull(r!COD_SALA) Then txtCodHORSala.Text = r!COD_SALA
       If Not IsNull(r!Dias) Then txtDias.Text = r!Dias
       If Not IsNull(r!HORARIO) Then mskHorario.Text = Format(r!HORARIO, "hh:mm")
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub cboAno_GotFocus()
Dim ANO, FirstYear, LastYear As Integer
Dim x As Long

If cboAno.ListCount = 0 Then

    ANO = Year(Date)
    FirstYear = ANO - 2
    LastYear = ANO + 2
     
    For x = ANO To FirstYear Step -1
        cboAno.AddItem x
    Next x
     
    ANO = ANO + 1
    For x = ANO To LastYear
        cboAno.AddItem x
    Next x
End If
End Sub


Private Sub cboClassificacao_GotFocus()
If cboClassificacao.ListCount = 0 Then
    cboClassificacao.AddItem "INFORMÁTICA"
    cboClassificacao.AddItem "LIVRE"
End If
    moCombo.AttachTo cboClassificacao
End Sub





Private Sub cboEtapa_GotFocus()
If cboEtapa.ListCount = 0 Then
    cboEtapa.AddItem "1"
    cboEtapa.AddItem "2"
    cboEtapa.AddItem "3"
    cboEtapa.AddItem "4"
    cboEtapa.AddItem "5"
    cboEtapa.AddItem "6"
    cboEtapa.AddItem "7"
    cboEtapa.AddItem "8"
End If
    moCombo.AttachTo cboEtapa
End Sub


Private Sub cboHORPacote_GotFocus()
If cboHORPacote.ListCount = 0 Then
   Dim sSQL As String
   Dim r As ADODB.Recordset
   'cboDesc.Clear
   
   sSQL = "SELECT DISTINCT PACOTE, CODIGO, Total FROM PACOTES ORDER BY PACOTE"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboHORPacote.AddItem r!PACOTE & " -> " & Format(r!Total, "##,##0.00")
      cboHORPacote.ItemData(cboHORPacote.NewIndex) = r!Codigo
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
    moCombo.AttachTo cboHORPacote
End Sub


Private Sub cboHORPacote_LostFocus()
On Error GoTo TrataErro

    If cboHORPacote.Text = "" Then txtCodHORPacote.Text = "": Exit Sub
    If cboHORPacote.ListIndex = -1 Then txtCodHORPacote.Text = "": Exit Sub

        txtCodHORPacote = cboHORPacote.ItemData(cboHORPacote.ListIndex)
  Exit Sub

TrataErro:
  If Err.Number = 381 Then
     Exit Sub
  End If
End Sub


Private Sub cboHORSala_GotFocus()
'If cboHORSala.ListCount = 0 Then

   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim varTexto As String
   Dim varCodText As String
   varTexto = cboHORSala
   varCodText = txtCodHORSala
   cboHORSala.Clear
   
   sSQL = "SELECT DISTINCT SALA, CODIGO FROM SALAS ORDER BY SALA"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboHORSala.AddItem r!SALA
      cboHORSala.ItemData(cboHORSala.NewIndex) = r!Codigo
      r.MoveNext
   Loop
   
   cboHORSala = varTexto
   txtCodHORSala = varCodText
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
'End If
    moCombo.AttachTo cboHORSala
End Sub


Private Sub cboHORSala_LostFocus()
On Error GoTo TrataErro

    If cboHORSala.Text = "" Then txtCodHORSala.Text = "": Exit Sub
    If cmdAlteraHorario.Visible = False Then If cboHORSala.ListIndex = -1 Then txtCodHORSala.Text = "": Exit Sub

        txtCodHORSala = cboHORSala.ItemData(cboHORSala.ListIndex)
  Exit Sub

TrataErro:
  If Err.Number = 381 Then
     Exit Sub
  End If
End Sub


Private Sub cboHORTemporada_GotFocus()
If cboHORTemporada.ListCount = 0 Then
   Dim sSQL As String
   Dim r As ADODB.Recordset
   'cboDesc.Clear
   
   sSQL = "SELECT DISTINCT ANO, CODIGO, ETAPA, INICIO, TERMINO FROM TEMPORADAS ORDER BY ANO"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboHORTemporada.AddItem r!ANO & "/" & r!ETAPA & "  -> " & Format(r!INICIO, "dd/mm/yy") & " -> " & Format(r!TERMINO, "dd/mm/yy")
      cboHORTemporada.ItemData(cboHORTemporada.NewIndex) = r!Codigo
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
    moCombo.AttachTo cboHORTemporada
End Sub


Private Sub cboHORTemporada_LostFocus()
On Error GoTo TrataErro

    If cboHORTemporada.Text = "" Then txtCodHORTemporada.Text = "": Exit Sub
    If cmdAlteraHorario.Visible = False Then If cboHORTemporada.ListIndex = -1 Then txtCodHORTemporada.Text = "": Exit Sub

        txtCodHORTemporada = cboHORTemporada.ItemData(cboHORTemporada.ListIndex)
  Exit Sub

TrataErro:
  If Err.Number = 381 Then
     Exit Sub
  End If
End Sub


Private Sub cboTipoDur_GotFocus()
If cboTipoDur.ListCount = 0 Then
    cboTipoDur.AddItem "HORA(S)"
    cboTipoDur.AddItem "MES(ES)"
End If
    moCombo.AttachTo cboTipoDur
End Sub



Private Sub chkDias_Click(index As Integer)
txtDias.Text = Dias
End Sub

Function Dias() As String
    Dim i As Integer
    Dias = ""
    For i = 0 To 6
        If chkDias(i).Value Then
            If Dias = "" Then
                Dias = chkDias(i).Tag
            Else
                Dias = Dias & ", " & chkDias(i).Tag
                
            End If
        End If
    Next i
    If Len(Dias) > 5 Then
        If InStr(1, Right(Dias, 5), ",") > 0 Then
        Dias = Left(Dias, Len(Dias) - 5) & Replace(Right(Dias, 5), ",", " e")
        End If
    End If
End Function

Private Sub cmdAdicionaCurso_Click()
If txtCodPac.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim x As Long
   
   sSQL = "SELECT MAX(CODIGO) AS ULTIMO FROM PACOTES_ITENS"
   Set r = dbData.OpenRecordset(sSQL)
   
   x = IIf(IsNull(r!ULTIMO) = True, 1, r!ULTIMO + 1)
   
  dbData.Execute ("INSERT INTO PACOTES_ITENS VALUES(" & x & "," & txtCodPac.Text & "," & GridCursoPacote.TextMatrix(GridCursoPacote.Row, 1) & ",'" & GridCursoPacote.TextMatrix(GridCursoPacote.Row, 2) & "')")
                  
   MontarGrid_CursoADDPacote

   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdAlteraHorario_Click()
If txtCodHorario.Text = "" Then Exit Sub

   If Not Atualizar_Dados_Horarios Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos_Horario
   cmdNovoHorario.Enabled = True
   cmdSalvarHorario.Visible = False
   cmdCancelarHorario.Visible = False
   cmdAlteraHorario.Visible = False
   cmdExcluirHorario.Visible = False
   frmCadHorario.Enabled = False
   MontarGrid_Horario
End Sub

Private Sub cmdAlterarCurso_Click()
If txtCodigoCurso.Text = "" Then Exit Sub

   If Not Atualizar_Dados_Cursos Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos_Cursos
   Form_Load
End Sub

Private Function Atualizar_Dados_Horarios() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE horario SET COD_PACOTE = " & txtCodHORPacote.Text & ", COD_TEMPORADA = " & txtCodHORTemporada.Text & ", COD_SALA = " & txtCodHORSala.Text & ", DIAS = '" & txtDias.Text & "', HORARIO = '" & Format$(mskHorario, ocHRMN) & "'  WHERE (Codigo = " & txtCodHorario.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Horarios = dbData.Execute(sSQL)
End Function
Private Function Atualizar_Dados_Cursos() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE cursos SET CURSO = '" & txtCurso.Text & "', CLASSIFICACAO = '" & cboClassificacao.Text & "', CARGA = " & txtCarga.Text & ", Valor = " & Replace(CCur(txtValor.Text), ",", ".") & "  WHERE (Codigo = " & txtCodigoCurso.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Cursos = dbData.Execute(sSQL)
End Function
Private Sub cmdAlterarPac_Click()
If txtCodPac.Text = "" Then Exit Sub

   If Not Atualizar_Dados_Pacotes Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos_Pacotes
   cmdNovoPac.Enabled = True
   cmdSalvarPac.Visible = False
   cmdCancelarPac.Visible = False
   cmdExcluirPac.Visible = False
   cmdAlterarPac.Visible = False
   frmPrincipalPacote.Enabled = False
   frmSecundarioPacote.Enabled = False
   LimparGrid_CursoADDPacote
   MontarGrid_Pacote
End Sub

Private Sub cmdAlterarSala_Click()
If txtCodSala.Text = "" Then Exit Sub

   If Not Atualizar_Dados_Salas Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
    Limpar_Objetos_Sala
    cmdNovoSala.Enabled = True
    cmdSalvarSala.Visible = False
    cmdCancelarSala.Visible = False
    cmdAlterarSala.Visible = False
    cmdExcluirSala.Visible = False
    frmPrincipalSala.Enabled = False
    MontarGrid_Sala
End Sub


Private Sub cmdAlterarTemporada_Click()
If txtCodTemporada.Text = "" Then Exit Sub

   If Not Atualizar_Dados_Temporada Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
    Limpar_Objetos_Temporada
    cmdNovoTemporada.Enabled = True
    cmdSalvarTemporada.Visible = False
    cmdCancelarTemporada.Visible = False
    cmdAlterarTemporada.Visible = False
    cmdExcluirTemporada.Visible = False
    frmPrincipalTemporada.Enabled = False
    MontarGrid_Temporada
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
  
   mskTermino = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCancelarCurso_Click()
    Limpar_Objetos_Cursos
    cmdNovoCurso.Enabled = True
    cmdSalvarCurso.Visible = False
    cmdCancelarCurso.Visible = False
    frmPrincipalCurso.Enabled = False
    MontarGrid_Cursos
End Sub

Private Sub Limpar_Objetos_Cursos()
If cmdAlterarCurso.Visible = False Then txtCodigoCurso.Text = ""
txtCurso.Text = ""
cboClassificacao.Text = ""
txtCarga.Text = ""
txtValor.Text = ""
End Sub




Private Sub cmdCancelarHorario_Click()
   Limpar_Objetos_Horario
   cmdNovoHorario.Enabled = True
   cmdSalvarHorario.Visible = False
   cmdCancelarHorario.Visible = False
   cmdAlteraHorario.Visible = False
   cmdExcluirHorario.Visible = False
   frmCadHorario.Enabled = False
   MontarGrid_Horario
End Sub

Private Sub cmdCancelarPac_Click()
If MsgBox("Cancelando o pacote todos os cursos adicionado até agora serão removidos!" & vbCrLf & "Deseja cancelar esse pacote?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso do Sistema") = vbNo Then Exit Sub
dbData.Execute ("DELETE FROM PACOTES_ITENS WHERE COD_PACOTE = " & txtCodPac.Text)
dbData.Execute ("DELETE FROM PACOTES WHERE CODIGO = " & txtCodPac.Text)

   Limpar_Objetos_Pacotes
   cmdNovoPac.Enabled = True
   cmdSalvarPac.Visible = False
   cmdCancelarPac.Visible = False
   cmdExcluirPac.Visible = False
   cmdAlterarPac.Visible = False
   frmPrincipalPacote.Enabled = False
   frmSecundarioPacote.Enabled = False
   LimparGrid_CursoADDPacote
   MontarGrid_Pacote
End Sub

Private Sub cmdCancelarSala_Click()
    Limpar_Objetos_Sala
    cmdNovoSala.Enabled = True
    cmdSalvarSala.Visible = False
    cmdCancelarSala.Visible = False
    cmdAlterarSala.Visible = False
    cmdExcluirSala.Visible = False
    frmPrincipalSala.Enabled = False
    MontarGrid_Sala
End Sub




Private Sub cmdCancelarTemporada_Click()
    Limpar_Objetos_Temporada
    cmdNovoTemporada.Enabled = True
    cmdSalvarTemporada.Visible = False
    cmdCancelarTemporada.Visible = False
    cmdAlterarTemporada.Visible = False
    cmdExcluirTemporada.Visible = False
    frmPrincipalTemporada.Enabled = False
    MontarGrid_Temporada
End Sub


Private Sub cmdExcluirCurso_Click()
If txtCodigoCurso.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim bRet As Boolean
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir esse curso?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM cursos WHERE (codigo = " & txtCodigoCurso.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   Limpar_Objetos_Cursos
   Form_Load
End Sub

Private Sub cmdExcluirHorario_Click()
If txtCodHorario.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim bRet As Boolean
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir esse horário?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE from HORARIO where (CODIGO = " & txtCodHorario.Text & ")"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   Limpar_Objetos_Horario
   cmdNovoHorario.Enabled = True
   cmdSalvarHorario.Visible = False
   cmdCancelarHorario.Visible = False
   cmdAlteraHorario.Visible = False
   cmdExcluirHorario.Visible = False
   frmCadHorario.Enabled = False
   MontarGrid_Horario
End Sub

Private Sub cmdExcluirPac_Click()
If MsgBox("Excluindo o pacote todos os cursos adicionado até agora serão removidos!" & vbCrLf & "Deseja excluir esse pacote?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso do Sistema") = vbNo Then Exit Sub
dbData.Execute ("DELETE FROM PACOTES_ITENS WHERE COD_PACOTE = " & txtCodPac.Text)
dbData.Execute ("DELETE FROM PACOTES WHERE CODIGO = " & txtCodPac.Text)

   Limpar_Objetos_Pacotes
   cmdNovoPac.Enabled = True
   cmdSalvarPac.Visible = False
   cmdCancelarPac.Visible = False
   cmdExcluirPac.Visible = False
   cmdAlterarPac.Visible = False
   frmPrincipalPacote.Enabled = False
   frmSecundarioPacote.Enabled = False
   LimparGrid_CursoADDPacote
   MontarGrid_Pacote
End Sub

Private Sub cmdExcluirSala_Click()
If txtCodSala.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim bRet As Boolean
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir essa sala?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE from SALAS where (CODIGO = " & txtCodSala.Text & ")"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
    Limpar_Objetos_Sala
    cmdNovoSala.Enabled = True
    cmdSalvarSala.Visible = False
    cmdCancelarSala.Visible = False
    cmdAlterarSala.Visible = False
    cmdExcluirSala.Visible = False
    frmPrincipalSala.Enabled = False
    MontarGrid_Sala
End Sub

Private Sub cmdExcluirTemporada_Click()
If txtCodTemporada.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim bRet As Boolean
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir esse temporada?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE from TEMPORADAS where (CODIGO = " & txtCodTemporada.Text & ")"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
    Limpar_Objetos_Temporada
    cmdNovoTemporada.Enabled = True
    cmdSalvarTemporada.Visible = False
    cmdCancelarTemporada.Visible = False
    cmdAlterarTemporada.Visible = False
    cmdExcluirTemporada.Visible = False
    frmPrincipalTemporada.Enabled = False
    MontarGrid_Temporada
End Sub


Private Sub cmdFechar_Click()
Unload Me
End Sub




Private Sub cmdNovoCurso_Click()
Limpar_Objetos_Cursos
frmPrincipalCurso.Enabled = True
txtCurso.SetFocus
cmdSalvarCurso.Visible = True
cmdCancelarCurso.Visible = True
cmdNovoCurso.Enabled = False
cmdAlterarCurso.Visible = False
cmdExcluirCurso.Visible = False
Autonumeracao_Cursos
End Sub




Private Sub cmdNovoHorario_Click()
frmCadHorario.Enabled = True
Limpar_Objetos_Horario
cmdSalvarHorario.Visible = True
cmdCancelarHorario.Visible = True
cmdAlteraHorario.Visible = False
cmdExcluirHorario.Visible = False
cmdNovoHorario.Enabled = False
Autonumeracao_Horario
cboHORPacote.SetFocus
End Sub

Private Sub cmdNovoPac_Click()
Limpar_Objetos_Pacotes
frmPrincipalPacote.Enabled = True
frmSecundarioPacote.Enabled = True
cmdSalvarPac.Visible = True
cmdCancelarPac.Visible = True
cmdNovoPac.Enabled = False
cmdAlterarPac.Visible = False
cmdExcluirPac.Visible = False
Autonumeracao_Pacotes
txtPacote.SetFocus
End Sub

Private Sub cmdNovoSala_Click()
Limpar_Objetos_Sala
frmPrincipalSala.Enabled = True
cmdSalvarSala.Visible = True
cmdCancelarSala.Visible = True
cmdNovoSala.Enabled = False
cmdAlterarSala.Visible = False
cmdExcluirSala.Visible = False
Autonumeracao_Sala
txtSala.SetFocus
End Sub

Private Sub cmdNovoTemporada_Click()
Limpar_Objetos_Temporada
frmPrincipalTemporada.Enabled = True
cmdSalvarTemporada.Visible = True
cmdCancelarTemporada.Visible = True
cmdAlterarTemporada.Visible = False
cmdExcluirTemporada.Visible = False
cmdNovoTemporada.Enabled = False
Autonumeracao_Temporada
cboAno.SetFocus
End Sub


Private Sub cmdRemoverCursos_Click()
On Error GoTo erro

If GridCursoADDPacote.TextMatrix(GridCursoADDPacote.Row, 1) = "" Then GoSub erro
If MsgBox("Deseja remover o curso: " & GridCursoADDPacote.TextMatrix(GridCursoADDPacote.Row, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso do Sistema") = vbNo Then Exit Sub

dbData.Execute ("DELETE FROM PACOTES_ITENS WHERE CODIGO = " & GridCursoADDPacote.TextMatrix(GridCursoADDPacote.Row, 1) & " AND COD_PACOTE = " & txtCodPac.Text)

MontarGrid_CursoADDPacote

Exit Sub
erro:
    MsgBox "Não existe nenhum curso selecionado para ser removido!", vbExclamation, "Aviso do Sistema"
    Exit Sub
End Sub

Private Sub cmdSalvarCurso_Click()
If txtCodigoCurso.Text = "" Or txtCurso.Text = "" Then Exit Sub
    
   If Not Inserir_Dados_Cursos Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos_Cursos
   Form_Load
   Exit Sub
End Sub



Private Function Inserir_Dados_Cursos() As Boolean
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO cursos (Codigo, CURSO, CLASSIFICACAO, CARGA, Valor) VALUES (" & _
      txtCodigoCurso.Text & ", '" & txtCurso.Text & "', '" & cboClassificacao.Text & "', " & txtCarga.Text & ", " & Replace(CCur(txtValor.Text), ",", ".") & ");"
      
   'Retorna o resultado da inclusão
   Inserir_Dados_Cursos = dbData.Execute(sSQL)
End Function


Private Function Inserir_Dados_Horarios() As Boolean
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO horario (Codigo, COD_PACOTE, COD_TEMPORADA, COD_SALA, DIAS, HORARIO) VALUES (" & _
      txtCodHorario.Text & ", " & txtCodHORPacote.Text & ", " & txtCodHORTemporada.Text & ", " & txtCodHORSala.Text & ", '" & txtDias.Text & "', '" & Format$(mskHorario, ocHRMN) & "' );"

   'Retorna o resultado da inclusão
   Inserir_Dados_Horarios = dbData.Execute(sSQL)
End Function
Private Function Inserir_Dados_Pacotes() As Boolean
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO pacotes (Codigo, PACOTE, TIPO_DURACAO, DURACAO, Total, PARC, QUANT) VALUES (" & _
      txtCodPac.Text & ", '" & txtPacote.Text & "', '" & cboTipoDur.Text & "', " & txtDuracao.Text & ", " & Replace(CCur(txtTotal.Text), ",", ".") & ", " & Replace(CCur(txtParc.Text), ",", ".") & ", " & txtQuant.Text & ");"

   'Retorna o resultado da inclusão
   Inserir_Dados_Pacotes = dbData.Execute(sSQL)
End Function
Private Function Inserir_Dados_Salas() As Boolean
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO salas (Codigo, SALA, VAGAS) VALUES (" & _
      txtCodSala.Text & ", '" & txtSala.Text & "', " & txtVagas.Text & ");"

   'Retorna o resultado da inclusão
   Inserir_Dados_Salas = dbData.Execute(sSQL)
End Function
Private Function Inserir_Dados_Temporadas() As Boolean
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO temporadas (CODIGO, ANO, ETAPA, INICIO, TERMINO) VALUES (" & _
      txtCodTemporada.Text & ", '" & cboAno.Text & "', " & cboEtapa.Text & ", CONVERT(DATETIME, '" & Format$(mskInicio.Text, ocDATA) & "', 103), CONVERT(DATETIME, '" & Format$(mskTermino.Text, ocDATA) & "', 103));"

   'Retorna o resultado da inclusão
   Inserir_Dados_Temporadas = dbData.Execute(sSQL)
End Function
Private Sub cmdSalvarHorario_Click()
If cboHORPacote.Text = "" Or txtDias.Text = "" Then Exit Sub
    
   If Not Inserir_Dados_Horarios Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
    
   Limpar_Objetos_Horario
   cmdNovoHorario.Enabled = True
   cmdSalvarHorario.Visible = False
   cmdCancelarHorario.Visible = False
   cmdAlteraHorario.Visible = False
   cmdExcluirHorario.Visible = False
   frmCadHorario.Enabled = False
   MontarGrid_Horario
End Sub

Private Sub cmdSalvarPac_Click()
If txtCodPac.Text = "" Or txtParc.Text = "" Then Exit Sub
    
   If Not Inserir_Dados_Pacotes Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos_Pacotes
   cmdNovoPac.Enabled = True
   cmdSalvarPac.Visible = False
   cmdCancelarPac.Visible = False
   cmdExcluirPac.Visible = False
   cmdAlterarPac.Visible = False
   frmPrincipalPacote.Enabled = False
   frmSecundarioPacote.Enabled = False
   LimparGrid_CursoADDPacote
   MontarGrid_Pacote
End Sub

Private Sub cmdSalvarSala_Click()
If txtCodSala.Text = "" Or txtSala.Text = "" Then Exit Sub
    
   If Not Inserir_Dados_Salas Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
    
    Limpar_Objetos_Sala
    cmdNovoSala.Enabled = True
    cmdSalvarSala.Visible = False
    cmdCancelarSala.Visible = False
    cmdAlterarSala.Visible = False
    cmdExcluirSala.Visible = False
    frmPrincipalSala.Enabled = False
    MontarGrid_Sala
End Sub

Private Sub cmdSalvarTemporada_Click()
If txtCodTemporada.Text = "" Or cboAno.Text = "" Then Exit Sub
    
   If Not Inserir_Dados_Temporadas Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
    
    Limpar_Objetos_Temporada
    cmdNovoTemporada.Enabled = True
    cmdSalvarTemporada.Visible = False
    cmdCancelarTemporada.Visible = False
    cmdAlterarTemporada.Visible = False
    cmdExcluirTemporada.Visible = False
    frmPrincipalTemporada.Enabled = False
    MontarGrid_Temporada
End Sub

Private Sub GridCursos_DblClick()
Limpar_Objetos_Cursos
frmPrincipalCurso.Enabled = True
txtCurso.SetFocus
cmdSalvarCurso.Visible = False
cmdCancelarCurso.Visible = False
cmdNovoCurso.Enabled = True
cmdAlterarCurso.Visible = True
cmdExcluirCurso.Visible = True
txtCodigoCurso.Text = ""
txtCodigoCurso.Text = (GridCursos.TextMatrix(GridCursos.Row, 1))
txtCurso.Text = (GridCursos.TextMatrix(GridCursos.Row, 2))
cboClassificacao.Text = (GridCursos.TextMatrix(GridCursos.Row, 3))
txtCarga.Text = (GridCursos.TextMatrix(GridCursos.Row, 4))
txtValor.Text = (GridCursos.TextMatrix(GridCursos.Row, 5))
End Sub










Private Sub Form_Load()
Set moCombo = New cComboHelper
cmdNovoCurso.Enabled = True
cmdSalvarCurso.Visible = False
cmdCancelarCurso.Visible = False
cmdAlterarCurso.Visible = False
cmdExcluirCurso.Visible = False


StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
SSTab1.Tab = 0
MontarGrid_Cursos
MontarGrid_CursoPacote
MontarGrid_CursoADDPacote
MontarGrid_Pacote
MontarGrid_Sala
MontarGrid_Temporada
LimparGrid_CursoADDPacote
MontarGrid_Horario
End Sub


Private Sub FormatarGrid_Temporada(r As ADODB.Recordset)
With GridTemporada
    
    .Clear
    .Cols = 6
    .rows = 2
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 1500
    .ColWidth(3) = 1000
    .ColWidth(4) = 1100
    .ColWidth(5) = 1100

 
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "ANO"
    .TextMatrix(0, 3) = "ETAPA"
    .TextMatrix(0, 4) = "INICIO"
    .TextMatrix(0, 5) = "TERMINO"

    
    'colocar os cabeçalho em negrito
    Dim x As Integer
    For x = 0 To .Cols - 1
    .Col = x
    .Row = 0
    .CellFontBold = True
    Next x
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until r.EOF
    
    .Redraw = False
    
    If Not IsNull(r!Codigo) Then .TextMatrix(.rows - 1, 1) = r!Codigo
    If Not IsNull(r!ANO) Then .TextMatrix(.rows - 1, 2) = r!ANO
    If Not IsNull(r!ETAPA) Then .TextMatrix(.rows - 1, 3) = r!ETAPA
    If Not IsNull(r!INICIO) Then .TextMatrix(.rows - 1, 4) = Format(r!INICIO, "dd/mm/yy")
    If Not IsNull(r!TERMINO) Then .TextMatrix(.rows - 1, 5) = Format(r!TERMINO, "dd/mm/yy")
    r.MoveNext
    .rows = .rows + 1
        
    Loop
    
    .rows = .rows - 1
    .Redraw = True
End With
End Sub
Private Sub FormatarGrid_Sala(r As ADODB.Recordset)
With GridSala
    
    .Clear
    .Cols = 4
    .rows = 2
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 1500
    .ColWidth(3) = 1000

 
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "SALA"
    .TextMatrix(0, 3) = "VAGAS"

    
    'colocar os cabeçalho em negrito
    Dim x As Integer
    For x = 0 To .Cols - 1
    .Col = x
    .Row = 0
    .CellFontBold = True
    Next x
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until r.EOF
    
    
    .Redraw = False
    
    If Not IsNull(r!Codigo) Then .TextMatrix(.rows - 1, 1) = r!Codigo
    If Not IsNull(r!SALA) Then .TextMatrix(.rows - 1, 2) = r!SALA
    If Not IsNull(r!VAGAS) Then .TextMatrix(.rows - 1, 3) = r!VAGAS
    r.MoveNext
    .rows = .rows + 1
        
    Loop
    
    .rows = .rows - 1
    .Redraw = True
    
End With
End Sub


Private Sub FormatarGrid_Pacote(r As ADODB.Recordset)
With GridPacote
    
    .Clear
    .Cols = 8
    .rows = 2
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 1700
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
 
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "PACOTE"
    .TextMatrix(0, 3) = "DURAÇÃO"
    .TextMatrix(0, 4) = "TIPO"
    .TextMatrix(0, 5) = "TOTAL"
    .TextMatrix(0, 6) = "QUANT."
    .TextMatrix(0, 7) = "PARCELA"

    
    'colocar os cabeçalho em negrito
    Dim x As Integer
    For x = 0 To .Cols - 1
    .Col = x
    .Row = 0
    .CellFontBold = True
    Next x
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until r.EOF
    
    'mudar a cor da coluna
    'Dim i As Integer
    'For i = 1 To .Rows - 1
    '.Row = i
    '.Col = 6:   .CellBackColor = &HC0FFFF
    'Next

    
    .Redraw = False
    
    'ALINHAMENTO
    '.ColAlignment(2) = 1
    
    
    If Not IsNull(r!Codigo) Then .TextMatrix(.rows - 1, 1) = r!Codigo
    If Not IsNull(r!PACOTE) Then .TextMatrix(.rows - 1, 2) = r!PACOTE
    If Not IsNull(r!DURACAO) Then .TextMatrix(.rows - 1, 3) = r!DURACAO
    If Not IsNull(r!TIPO_DURACAO) Then .TextMatrix(.rows - 1, 4) = r!TIPO_DURACAO
    If Not IsNull(r!Total) Then .TextMatrix(.rows - 1, 5) = Format(r!Total, "##,##0.00")
    If Not IsNull(r!QUANT) Then .TextMatrix(.rows - 1, 6) = r!QUANT
    If Not IsNull(r!PARC) Then .TextMatrix(.rows - 1, 7) = Format(r!PARC, "##,##0.00")
    r.MoveNext
    .rows = .rows + 1
        
    Loop
    
    .rows = .rows - 1
   .Redraw = True

End With
End Sub


Private Sub FormatarGrid_CursoPacote(r As ADODB.Recordset)
With GridCursoPacote
    
    .Clear
    .Cols = 3
    .rows = 2
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 2000
 
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "CURSO"

    
    'colocar os cabeçalho em negrito
    Dim x As Integer
    For x = 0 To .Cols - 1
    .Col = x
    .Row = 0
    .CellFontBold = True
    Next x
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until r.EOF
    
    .Redraw = False
    
    
    
    If Not IsNull(r!Codigo) Then .TextMatrix(.rows - 1, 1) = r!Codigo
    If Not IsNull(r!CURSO) Then .TextMatrix(.rows - 1, 2) = r!CURSO
    r.MoveNext
    .rows = .rows + 1
        
    Loop
    
    .rows = .rows - 1
    .Redraw = True
    
End With
End Sub


Private Sub FormatarGrid_CursoADDPacote(r As ADODB.Recordset)
With GridCursoADDPacote
    .Clear
    .Cols = 3
    .rows = 2
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 2000
    
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "CURSO"
    
    'colocar os cabeçalho em negrito
    Dim x As Integer
    For x = 0 To .Cols - 1
    .Col = x
    .Row = 0
    .CellFontBold = True
    Next x
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    .Redraw = False
    
    Do Until r.EOF
    
    If Not IsNull(r!COD_CURSO) Then .TextMatrix(.rows - 1, 1) = r!COD_CURSO
    If Not IsNull(r!CURSO) Then .TextMatrix(.rows - 1, 2) = r!CURSO
    r.MoveNext
    .rows = .rows + 1
        
    Loop
    
    .rows = .rows - 1
    .Redraw = True
    
End With
End Sub


Private Sub FormatarGrid_Cursos(r As ADODB.Recordset)
With GridCursos
    .Clear
    .Cols = 6
    .rows = 2
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 2400
    .ColWidth(3) = 2300
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "CURSO"
    .TextMatrix(0, 3) = "CLASSIFICAÇÃO"
    .TextMatrix(0, 4) = "CARGA"
    .TextMatrix(0, 5) = "VALOR"

    'colocar os cabeçalho em negrito
    Dim x As Integer
    For x = 0 To .Cols - 1
    .Col = x
    .Row = 0
    .CellFontBold = True
    Next x
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until r.EOF
    
    .Redraw = False
    
    If Not IsNull(r!Codigo) Then .TextMatrix(.rows - 1, 1) = r!Codigo
    If Not IsNull(r!CURSO) Then .TextMatrix(.rows - 1, 2) = r!CURSO
    If Not IsNull(r!CLASSIFICACAO) Then .TextMatrix(.rows - 1, 3) = r!CLASSIFICACAO
    If Not IsNull(r!CARGA) Then .TextMatrix(.rows - 1, 4) = r!CARGA
    If Not IsNull(r!Valor) Then .TextMatrix(.rows - 1, 5) = Format(r!Valor, "##,##0.00")
    r.MoveNext
    .rows = .rows + 1
        
    Loop
    
    .rows = .rows - 1
    .Redraw = True
End With
End Sub
Private Function FormatarGrid_Horario(r As ADODB.Recordset)
With GridHorario
    .Clear
    .Cols = 10
    .rows = 2
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 0
    .ColWidth(3) = 2100
    .ColWidth(4) = 0
    .ColWidth(5) = 1500
    .ColWidth(6) = 0
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    
    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "COD_PAC"
    .TextMatrix(0, 3) = "PACOTE"
    .TextMatrix(0, 4) = "COD_TEM"
    .TextMatrix(0, 5) = "TEMPORADA"
    .TextMatrix(0, 6) = "COD_SALA"
    .TextMatrix(0, 7) = "SALA"
    .TextMatrix(0, 8) = "DIAS"
    .TextMatrix(0, 9) = "HORARIO"

    'colocar os cabeçalho em negrito
    Dim x As Integer
    For x = 0 To .Cols - 1
    .Col = x
    .Row = 0
    .CellFontBold = True
    Next x
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    Do Until r.EOF
    
    .Redraw = False
    If Not IsNull(r!VAR_CODIGO) Then .TextMatrix(.rows - 1, 1) = r!VAR_CODIGO
    If Not IsNull(r!varcodpac) Then .TextMatrix(.rows - 1, 2) = r!varcodpac
    If Not IsNull(r!varpac) Then .TextMatrix(.rows - 1, 3) = r!varpac
    If Not IsNull(r!varcodtem) Then .TextMatrix(.rows - 1, 4) = r!varcodtem
    If Not IsNull(r!varANO) Then .TextMatrix(.rows - 1, 5) = r!varANO & "/" & r!VARETAPA
    If Not IsNull(r!varcodsal) Then .TextMatrix(.rows - 1, 6) = r!varcodsal
    If Not IsNull(r!varSAL) Then .TextMatrix(.rows - 1, 7) = r!varSAL
    If Not IsNull(r!vardias) Then .TextMatrix(.rows - 1, 8) = r!vardias
    If Not IsNull(r!varhor) Then .TextMatrix(.rows - 1, 9) = Format(r!varhor, "hh:mm")
    r.MoveNext
    .rows = .rows + 1
        
    Loop
    
    .rows = .rows - 1
    .Redraw = True
End With
End Function

Private Sub Autonumeracao_Temporada()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS COD_TEMPORADA FROM TEMPORADAS;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then txtCodTemporada.Text = r("COD_TEMPORADA") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Autonumeracao_Sala()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS COD_SALA FROM SALAS;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then txtCodSala.Text = r("COD_SALA") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Autonumeracao_Horario()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS COD_HORARIO FROM HORARIO;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then txtCodHorario.Text = r("COD_HORARIO") + 1
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub Autonumeracao_Pacotes()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS COD_PACOTE FROM PACOTES;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then txtCodPac.Text = r("COD_PACOTE") + 1
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Private Sub Autonumeracao_Cursos()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS COD_CURSO FROM CURSOS;"
   
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigoCurso.Text = r("COD_CURSO") + 1
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

'Private Sub Atualizar_Dados_Horario()
'   Dim sSQL As String
   
   'Comando de atualização
'   sSQL = "UPDATE cursos SET COD_PACOTE = " & txtCodHORPacote.Text & ", COD_TEMPORADA = " & txtCodHORTemporada.Text & ", COD_SALA = " & txtCodHORSala.Text & ", Dias = '" & txtDias.Text & "', HORARIO = '" & Format$(mskHorario, ocHRMN) & "'  WHERE (Codigo = " & txtCodHorario.Text & ");"
   
   'Retorna o resultado da atualização
'   Atualizar_Dados_Horario = dbData.Execute(sSQL)
'End Sub







Private Function Atualizar_Dados_Salas() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE SALAs SET SALA = '" & txtSala.Text & "', VAGAS = " & txtVagas.Text & "  WHERE (Codigo = " & txtCodSala.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Salas = dbData.Execute(sSQL)
End Function
Private Sub Limpar_Objetos_Temporada()
txtCodTemporada.Text = ""
cboAno.Text = ""
cboEtapa.Text = ""
mskInicio.Mask = ""
mskInicio.Text = ""
mskTermino.Mask = ""
mskTermino.Text = ""
End Sub
Private Sub Limpar_Objetos_Sala()
txtCodSala.Text = ""
txtSala.Text = ""
txtVagas.Text = ""
End Sub
Private Function Atualizar_Dados_Temporada() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE temporadas SET ANO = '" & cboAno.Text & "', ETAPA = '" & cboEtapa.Text & "', INICIO = CONVERT(DATETIME, '" & Format$(mskInicio.Text, ocDATA) & "', 103), TERMINO = CONVERT(DATETIME, '" & Format$(mskTermino.Text, ocDATA) & "', 103) WHERE (Codigo = " & txtCodTemporada.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Temporada = dbData.Execute(sSQL)
End Function



Private Function Atualizar_Dados_Pacotes() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE pacotes SET PACOTE = '" & txtPacote.Text & "', TIPO_DURACAO = '" & cboTipoDur.Text & "', DURACAO = " & txtDuracao.Text & ", Total = " & Replace(CCur(txtTotal.Text), ",", ".") & ", PARC = " & Replace(CCur(txtParc.Text), ",", ".") & ", QUANT = " & txtQuant.Text & "  WHERE (Codigo = " & txtCodPac.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Pacotes = dbData.Execute(sSQL)
End Function
Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

Private Sub GridHorario_DblClick()
frmCadHorario.Enabled = True
cmdSalvarHorario.Visible = False
cmdCancelarHorario.Visible = False
cmdAlteraHorario.Visible = True
cmdExcluirHorario.Visible = True
cmdNovoHorario.Enabled = False
txtCodHorario.Text = ""
txtCodHorario.Text = (GridHorario.TextMatrix(GridHorario.Row, 1))
End Sub


Private Sub GridPacote_DblClick()
frmPrincipalPacote.Enabled = True
frmSecundarioPacote.Enabled = True
txtPacote.SetFocus
cmdSalvarPac.Visible = False
cmdCancelarPac.Visible = False
cmdNovoPac.Enabled = True
cmdAlterarPac.Visible = True
cmdExcluirPac.Visible = True
txtCodPac.Text = ""
txtCodPac.Text = (GridPacote.TextMatrix(GridPacote.Row, 1))
txtPacote.Text = (GridPacote.TextMatrix(GridPacote.Row, 2))
txtDuracao.Text = (GridPacote.TextMatrix(GridPacote.Row, 3))
cboTipoDur.Text = (GridPacote.TextMatrix(GridPacote.Row, 4))
txtTotal.Text = Format((GridPacote.TextMatrix(GridPacote.Row, 5)), "##,##0.00")
txtQuant.Text = (GridPacote.TextMatrix(GridPacote.Row, 6))
txtParc.Text = Format((GridPacote.TextMatrix(GridPacote.Row, 7)), "##,##0.00")
End Sub


Private Sub GridSala_DblClick()
frmPrincipalSala.Enabled = True
txtSala.SetFocus
cmdSalvarSala.Visible = False
cmdCancelarSala.Visible = False
cmdNovoSala.Enabled = True
cmdAlterarSala.Visible = True
cmdExcluirSala.Visible = True
txtCodSala.Text = ""
txtCodSala.Text = (GridSala.TextMatrix(GridSala.Row, 1))
txtSala.Text = (GridSala.TextMatrix(GridSala.Row, 2))
txtVagas.Text = (GridSala.TextMatrix(GridSala.Row, 3))
End Sub


Private Sub GridTemporada_DblClick()
frmPrincipalTemporada.Enabled = True
cboAno.SetFocus
cmdSalvarTemporada.Visible = False
cmdCancelarTemporada.Visible = False
cmdNovoTemporada.Enabled = True
cmdAlterarTemporada.Visible = True
cmdExcluirTemporada.Visible = True
txtCodTemporada.Text = ""
txtCodTemporada.Text = (GridTemporada.TextMatrix(GridTemporada.Row, 1))
cboAno.Text = (GridTemporada.TextMatrix(GridTemporada.Row, 2))
cboEtapa.Text = (GridTemporada.TextMatrix(GridTemporada.Row, 3))
mskInicio.Text = Format((GridTemporada.TextMatrix(GridTemporada.Row, 4)), "dd/mm/yy")
mskTermino.Text = Format((GridTemporada.TextMatrix(GridTemporada.Row, 5)), "dd/mm/yy")
End Sub


Private Sub mskHorario_KeyPress(KeyAscii As Integer)
mskHorario.Mask = "##:##"
End Sub


Private Sub mskHorario_LostFocus()
If mskHorario.Text = "" Or mskHorario.Text = "__:__" Then
    mskHorario.Mask = ""
    mskHorario.Text = ""
    Exit Sub
End If
End Sub


Private Sub mskInicio_KeyPress(KeyAscii As Integer)
mskInicio.Mask = "##/##/##"
End Sub


Private Sub mskInicio_LostFocus()
If mskInicio.Text = "" Or mskInicio.Text = "__/__/__" Then
    mskInicio.Mask = ""
    mskInicio.Text = ""
    Exit Sub
Else
    If IsDate(mskInicio.Text) Then
        Exit Sub
    Else
        MsgBox "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation, "Aviso do Sistema"
        mskInicio.SetFocus
        mskInicio.SelStart = 0
        mskInicio.SelLength = Len(mskInicio)
    End If
End If
End Sub


Private Sub mskTermino_KeyPress(KeyAscii As Integer)
mskTermino.Mask = "##/##/##"
End Sub


Private Sub mskTermino_LostFocus()
If mskTermino.Text = "" Or mskTermino.Text = "__/__/__" Then
    mskTermino.Mask = ""
    mskTermino.Text = ""
    Exit Sub
Else
    If IsDate(mskTermino.Text) Then
        Exit Sub
    Else
        MsgBox "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation, "Aviso do Sistema"
        mskTermino.SetFocus
        mskTermino.SelStart = 0
        mskTermino.SelLength = Len(mskTermino)
    End If
End If
End Sub


Private Sub optDias_Click()
MontarGrid_Horario
End Sub

Private Sub optHorario_Click()
MontarGrid_Horario
End Sub


Private Sub optPacote_Click()
MontarGrid_Horario
End Sub

Private Sub optSala_Click()
MontarGrid_Horario
End Sub

Private Sub optTemporada_Click()
MontarGrid_Horario
End Sub


Private Sub txtCodHorario_Change()
If cmdAlteraHorario.Visible = True Then MostrarHorario
End Sub

Private Sub txtCodHORPacote_Change()
If cmdAlteraHorario.Visible = True Then MostrarPacote
End Sub


Private Sub txtCodHORSala_Change()
If cmdAlteraHorario.Visible = True Then MostrarSala
End Sub

Private Sub txtCodHORTemporada_Change()
If cmdAlteraHorario.Visible = True Then MostrarTemporada
End Sub


Private Sub txtCodPac_Change()
MontarGrid_CursoADDPacote
End Sub

Private Sub txtCurso_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtPacote_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtParc_GotFocus()
txtParc.SelStart = 0
txtParc.SelLength = Len(txtParc)
End Sub

Private Sub txtParc_LostFocus()
If txtParc.Text = "" Then txtParc.Text = Format(0, "##,##0.00") Else txtParc.Text = Format(txtParc, "##,##0.00")
Calcular_Parcelas
End Sub

Private Sub txtQuant_LostFocus()
Calcular_Parcelas
End Sub


Private Sub txtSala_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtTotal_LostFocus()
If txtTotal.Text = "" Then txtTotal.Text = Format(0, "##,##0.00") Else txtTotal.Text = Format(txtTotal, "##,##0.00")
Calcular_Parcelas
End Sub


Private Sub txtValor_LostFocus()
If txtValor.Text = "" Then Exit Sub
txtValor.Text = FormatNumber(txtValor.Text, 2)
End Sub


