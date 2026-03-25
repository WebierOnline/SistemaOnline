VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Consultorio_Agendamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AGENDAMENTO DE CONSULTA"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   Icon            =   "Consultorio_Agendamento.frx":0000
   LinkTopic       =   "Form73"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   60
      TabIndex        =   21
      Top             =   1080
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   14420
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "AGENDAMENTO"
      TabPicture(0)   =   "Consultorio_Agendamento.frx":23D2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmReserva"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Grid"
      Tab(0).Control(3)=   "cmdSair"
      Tab(0).Control(4)=   "cmdSalvar"
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(6)=   "cmdAlterar"
      Tab(0).Control(7)=   "cmdExcluir"
      Tab(0).Control(8)=   "cmdCliente"
      Tab(0).Control(9)=   "cmdImprimir"
      Tab(0).Control(10)=   "cmdNovo"
      Tab(0).Control(11)=   "lblQuant"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "HISTÓRICO"
      TabPicture(1)   =   "Consultorio_Agendamento.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "CONSULTA"
      TabPicture(2)   =   "Consultorio_Agendamento.frx":240A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblQuantConsulta"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Grid_Consulta"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame1 
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
         Height          =   975
         Left            =   120
         TabIndex        =   45
         Top             =   7080
         Width           =   11235
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   6420
            Sorted          =   -1  'True
            TabIndex        =   56
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboMES 
            Height          =   315
            ItemData        =   "Consultorio_Agendamento.frx":2426
            Left            =   4260
            List            =   "Consultorio_Agendamento.frx":2428
            TabIndex        =   55
            Top             =   480
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox txtCodCriterio 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8640
            TabIndex        =   54
            Top             =   180
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cboCriterio 
            Height          =   315
            Left            =   4260
            TabIndex        =   48
            Top             =   480
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.ComboBox cboOrdem 
            Height          =   315
            Left            =   2340
            TabIndex        =   47
            Top             =   480
            Width           =   1875
         End
         Begin VB.ComboBox cboTipoConsulta 
            Height          =   315
            Left            =   180
            TabIndex        =   46
            Top             =   480
            Width           =   2115
         End
         Begin MSMask.MaskEdBox mskDataCriterio 
            Height          =   315
            Left            =   4260
            TabIndex        =   49
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizar 
            Height          =   555
            Left            =   9420
            TabIndex        =   53
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
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
            MICON           =   "Consultorio_Agendamento.frx":242A
            PICN            =   "Consultorio_Agendamento.frx":2446
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCalendario 
            Height          =   315
            Left            =   5760
            TabIndex        =   60
            Top             =   480
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
            MICON           =   "Consultorio_Agendamento.frx":2D20
            PICN            =   "Consultorio_Agendamento.frx":2D3C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            Caption         =   "Ano"
            Height          =   195
            Left            =   6420
            TabIndex        =   59
            Top             =   240
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            Caption         =   "Mês"
            Height          =   195
            Left            =   4260
            TabIndex        =   58
            Top             =   240
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   4260
            TabIndex        =   57
            Top             =   240
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblCriterio 
            AutoSize        =   -1  'True
            Caption         =   "Critério"
            Height          =   195
            Left            =   4260
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   2340
            TabIndex        =   51
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Consulta:"
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   240
            Width           =   1245
         End
      End
      Begin VB.Frame frmReserva 
         Caption         =   "AGENDAMENTO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -74880
         TabIndex        =   28
         Top             =   420
         Width           =   11235
         Begin VB.ComboBox cboRecepcionista 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1140
            Width           =   2415
         End
         Begin VB.TextBox txtCodRecepcionista 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   40
            Top             =   840
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   4560
            TabIndex        =   10
            Top             =   1140
            Width           =   1815
         End
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   4815
         End
         Begin VB.ComboBox cboProfissional 
            Height          =   315
            Left            =   4980
            TabIndex        =   2
            Top             =   480
            Width           =   2955
         End
         Begin VB.TextBox txtCodCliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4380
            TabIndex        =   30
            Top             =   180
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cboSala 
            Height          =   315
            Left            =   7980
            TabIndex        =   3
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtCodProfissional 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6420
            TabIndex        =   29
            Top             =   180
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   315
            Left            =   8880
            TabIndex        =   4
            Top             =   480
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdCalendario2 
            Height          =   315
            Left            =   10020
            TabIndex        =   5
            Top             =   480
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
            MICON           =   "Consultorio_Agendamento.frx":511E
            PICN            =   "Consultorio_Agendamento.frx":513A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskHora 
            Height          =   315
            Left            =   10380
            TabIndex        =   6
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataCadastro 
            Height          =   315
            Left            =   2580
            TabIndex        =   8
            Top             =   1140
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskHoraCadastro 
            Height          =   315
            Left            =   3780
            TabIndex        =   9
            Top             =   1140
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hora"
            Height          =   195
            Left            =   3780
            TabIndex        =   43
            Top             =   900
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   2580
            TabIndex        =   42
            Top             =   900
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Recepcionista"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   900
            Width           =   1020
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   4560
            TabIndex        =   39
            Top             =   900
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hora"
            Height          =   195
            Left            =   10380
            TabIndex        =   38
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   8880
            TabIndex        =   33
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissional"
            Height          =   195
            Left            =   4980
            TabIndex        =   32
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Sala"
            Height          =   195
            Left            =   7980
            TabIndex        =   31
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   23
         Top             =   7440
         Width           =   3315
         Begin VB.CommandButton Command12 
            Caption         =   "ok"
            Height          =   255
            Left            =   1620
            TabIndex        =   27
            Top             =   300
            Width           =   375
         End
         Begin VB.CommandButton cmdVoltar 
            Caption         =   "<<"
            Height          =   255
            Left            =   2040
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   300
            Width           =   555
         End
         Begin VB.CommandButton cmdAvancar 
            Caption         =   ">>"
            Height          =   255
            Left            =   2640
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   300
            Width           =   495
         End
         Begin MSMask.MaskEdBox mskDataConsulta 
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   22
         Top             =   2820
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   7435
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   555
         Left            =   -65340
         TabIndex        =   15
         Top             =   7500
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
         MICON           =   "Consultorio_Agendamento.frx":751C
         PICN            =   "Consultorio_Agendamento.frx":7538
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
         Height          =   555
         Left            =   -74880
         TabIndex        =   11
         Top             =   2100
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
         MICON           =   "Consultorio_Agendamento.frx":7852
         PICN            =   "Consultorio_Agendamento.frx":786E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   555
         Left            =   -73080
         TabIndex        =   12
         Top             =   2100
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
         MICON           =   "Consultorio_Agendamento.frx":E138
         PICN            =   "Consultorio_Agendamento.frx":E154
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   555
         Left            =   -74880
         TabIndex        =   13
         Top             =   2100
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
         MICON           =   "Consultorio_Agendamento.frx":14BF8
         PICN            =   "Consultorio_Agendamento.frx":14C14
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
         Height          =   555
         Left            =   -73080
         TabIndex        =   14
         Top             =   2100
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
         MICON           =   "Consultorio_Agendamento.frx":154EE
         PICN            =   "Consultorio_Agendamento.frx":1550A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCliente 
         Height          =   555
         Left            =   -65340
         TabIndex        =   35
         Top             =   2100
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "&Cliente"
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
         MICON           =   "Consultorio_Agendamento.frx":15824
         PICN            =   "Consultorio_Agendamento.frx":15840
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
         Left            =   -71520
         TabIndex        =   36
         Top             =   7440
         Width           =   1575
         _ExtentX        =   2778
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
         MICON           =   "Consultorio_Agendamento.frx":15B5A
         PICN            =   "Consultorio_Agendamento.frx":15B76
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
         Height          =   555
         Left            =   -67080
         TabIndex        =   0
         Top             =   2100
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
         MICON           =   "Consultorio_Agendamento.frx":16304
         PICN            =   "Consultorio_Agendamento.frx":16320
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta 
         Height          =   6315
         Left            =   120
         TabIndex        =   44
         Top             =   420
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   11139
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label lblQuantConsulta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   11100
         TabIndex        =   61
         Top             =   6780
         Width           =   225
      End
      Begin VB.Label lblQuant 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -63900
         TabIndex        =   37
         Top             =   7080
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   11445
      TabIndex        =   17
      Top             =   60
      Width           =   11475
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9660
         TabIndex        =   20
         Top             =   300
         Visible         =   0   'False
         Width           =   1635
      End
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "Consultorio_Agendamento.frx":16FFA
         Top             =   0
         Width           =   960
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AGENDAMENTO"
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
         TabIndex        =   19
         Top             =   240
         Width           =   2460
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   16
      Top             =   9345
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16140
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "23:21"
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
Attribute VB_Name = "Consultorio_Agendamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private printSQL As String
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Private Function Atualizar_Dados() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE consultorio_agendamento SET " & _
      "cod_cliente = '" & txtCodCliente.Text & "', " & _
      "tipo = '" & cboTipo.Text & "', " & _
      "cod_profissional = '" & txtCodProfissional.Text & "', " & _
      "cod_recepcionista = '" & txtCodRecepcionista.Text & "', " & _
      "data = CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
      "hora = " & IIf(mskHora.Text = "", "Null", "'" & Format$(mskHora.Text, ocHORA) & "'") & ", " & _
      "data_cadastro = CONVERT(DATETIME, '" & Format$(mskDataCadastro.Text, ocDATA) & "', 103), " & _
      "hora_cadastro = " & IIf(mskHoraCadastro.Text = "", "Null", "'" & Format$(mskHoraCadastro.Text, ocHORA) & "'") & ", " & _
      "sala = '" & cboSala.Text & "' "
   
   'Condição para atualização
   sSQL = sSQL & "WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Limpar_Objetos()
txtCodigo.Text = ""
cboCliente.Text = ""
txtCodCliente.Text = ""
cboTipo.Text = ""
cboProfissional.Text = ""
txtCodProfissional.Text = ""
cboRecepcionista.Text = ""
txtCodRecepcionista.Text = ""
cboSala.Text = ""
mskData.Mask = ""
mskData.Text = ""
mskHora.Mask = ""
mskHora.Text = ""
mskDataCadastro.Mask = ""
mskDataCadastro.Text = ""
mskHoraCadastro.Mask = ""
mskHoraCadastro.Text = ""
cmdNovo.Visible = True
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = False
cmdExcluir.Visible = False
frmReserva.Enabled = False
End Sub

Private Sub Mostrar_Grid()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long

If Not IsDate(mskDataConsulta) Then Exit Sub
sSQL = "SELECT cliente.*, cliente.nome, consultorio_agendamento.codigo as varCod, consultorio_agendamento.*, funcionario.codigo, funcionario.nome as profissional " & _
"FROM ((consultorio_agendamento INNER JOIN cliente ON cliente.codigo = consultorio_agendamento.cod_cliente) " & _
"INNER JOIN funcionario ON consultorio_agendamento.cod_profissional = funcionario.codigo) " & _
"WHERE (data = CONVERT(DATETIME, '" & Format(mskDataConsulta.Text, ocDATA) & "', 103)) " & _
"ORDER BY consultorio_agendamento.data, consultorio_agendamento.hora;"
   

Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid r
If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL

'MOSTRAR A QUANTIDADE REGISTROS
lblQuant.Caption = Format(totalRegistros, "00")
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then
   cboTipo.Text = rTabela("tipo")
   mskData.Text = Format(rTabela("data"), "dd/mm/yy")
   mskHora.Text = Format(rTabela("hora"), "hh:mm")
   mskDataCadastro.Text = Format(rTabela("data_cadastro"), "dd/mm/yy")
   mskHoraCadastro.Text = Format(rTabela("hora_cadastro"), "hh:mm")
   cboSala.Text = rTabela("sala")
   txtCodCliente.Text = rTabela("cod_cliente")
   txtCodProfissional.Text = rTabela("cod_profissional")
   txtCodRecepcionista.Text = rTabela("cod_recepcionista")
End If
End Sub

Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = LastYear To FirstYear Step -1
   cboAno.AddItem i
Next
End Sub


Private Sub CboCliente_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboCliente.Text
codAtual = txtCodCliente.Text
cboCliente.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboCliente.AddItem r("nome")
   cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboCliente.Text = itemAtual
txtCodCliente.Text = codAtual
moCombo.AttachTo cboCliente
End Sub


Private Sub CboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Maiuscula(KeyAscii)
End Sub


Private Sub CboCliente_LostFocus()
  On Error GoTo TrataErro
   'If chkCodPedido.Value = Unchecked Then
   If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   'If chkCodPedido.Value = Unchecked Then
   If cmdAlterar.Visible = False Then If cboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
   
   txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
 '  If chkCodPedido.Value = Unchecked Then Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboCriterio_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim itemAtual As String
   Dim codAtual As String
   
   
If cboTipoConsulta.Text = "CLIENTE" Then
   itemAtual = cboCriterio.Text
   codAtual = txtCodCriterio.Text
   cboCriterio.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCriterio.AddItem r("nome")
      cboCriterio.ItemData(cboCriterio.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboCriterio.Text = itemAtual
   txtCodCriterio.Text = codAtual
   moCombo.AttachTo cboCriterio
ElseIf cboTipoConsulta.Text = "PROFISSIONAL" Then
   itemAtual = cboCriterio.Text
   codAtual = txtCodCriterio.Text
   cboCriterio.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM funcionario WHERE cargo = 'dentista' ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCriterio.AddItem r("nome")
      cboCriterio.ItemData(cboCriterio.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboCriterio.Text = itemAtual
   txtCodCriterio.Text = codAtual
   moCombo.AttachTo cboCriterio
ElseIf cboTipoConsulta.Text = "TIPO" Then
   itemAtual = cboCriterio.Text
   cboCriterio.Clear
   
   sSQL = "SELECT tipo FROM consultorio_agendamento GROUP BY tipo;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCriterio.AddItem ValidateNull(r("tipo"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboCriterio.Text = itemAtual
   moCombo.AttachTo cboCriterio
Else
   Exit Sub
End If
End Sub


Private Sub cboCriterio_LostFocus()
  On Error GoTo TrataErro
   If cboCriterio.Text = "" Then txtCodCriterio.Text = "": Exit Sub
   If cmdAlterar.Visible = False Then If cboCriterio.ListIndex = -1 Then txtCodCriterio.Text = "": Exit Sub
   
   txtCodCriterio = cboCriterio.ItemData(cboCriterio.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboMes_GotFocus()
Dim vMes As Integer

cboMes.Clear

For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMes
End Sub


Private Sub cboOrdem_GotFocus()
cboOrdem.Clear
cboOrdem.AddItem "DATA"
cboOrdem.AddItem "CLIENTE"
cboOrdem.AddItem "SALA"
cboOrdem.AddItem "PROFISSIONAL"
cboOrdem.AddItem "TIPO"
moCombo.AttachTo cboOrdem
End Sub


Private Sub cboProfissional_LostFocus()
  On Error GoTo TrataErro
   If cboProfissional.Text = "" Then txtCodProfissional.Text = "": Exit Sub
   If cmdAlterar.Visible = False Then If cboProfissional.ListIndex = -1 Then txtCodProfissional.Text = "": Exit Sub
   
   txtCodProfissional = cboProfissional.ItemData(cboProfissional.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboRecepcionista_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboRecepcionista.Text
codAtual = txtCodRecepcionista.Text
cboRecepcionista.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM funcionario WHERE cargo = 'recepcionista' ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboRecepcionista.AddItem r("nome")
   cboRecepcionista.ItemData(cboRecepcionista.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboRecepcionista.Text = itemAtual
txtCodRecepcionista.Text = codAtual
moCombo.AttachTo cboRecepcionista
End Sub


Private Sub cboRecepcionista_LostFocus()
  On Error GoTo TrataErro
   If cboRecepcionista.Text = "" Then txtCodRecepcionista.Text = "": Exit Sub
   If cmdAlterar.Visible = False Then If cboRecepcionista.ListIndex = -1 Then txtCodRecepcionista.Text = "": Exit Sub
   
   txtCodRecepcionista = cboRecepcionista.ItemData(cboRecepcionista.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboSala_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

Dim itemAtual As String
itemAtual = cboSala.Text
cboSala.Clear

sSQL = "SELECT sala FROM consultorio_agendamento GROUP BY sala;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboSala.AddItem ValidateNull(r("sala"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboSala.Text = itemAtual
moCombo.AttachTo cboSala
End Sub


Private Sub cboTipo_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

Dim itemAtual As String
itemAtual = cboTipo.Text
cboTipo.Clear

sSQL = "SELECT tipo FROM consultorio_agendamento GROUP BY tipo;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboTipo.AddItem ValidateNull(r("tipo"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboTipo.Text = itemAtual
moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
KeyAscii = Maiuscula(KeyAscii)
End Sub


Private Sub cboProfissional_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboProfissional.Text
codAtual = txtCodProfissional.Text
cboProfissional.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM funcionario WHERE cargo = 'dentista' ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboProfissional.AddItem r("nome")
   cboProfissional.ItemData(cboProfissional.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboProfissional.Text = itemAtual
txtCodProfissional.Text = codAtual
moCombo.AttachTo cboProfissional
End Sub


Private Sub cboProfissional_KeyPress(KeyAscii As Integer)
KeyAscii = Maiuscula(KeyAscii)
End Sub


Private Sub cboTipoConsulta_Click()
cboTipoConsulta_LostFocus
End Sub

Private Sub cboTipoConsulta_GotFocus()
cboTipoConsulta.Clear
cboTipoConsulta.AddItem "CLIENTE"
cboTipoConsulta.AddItem "PROFISSIONAL"
cboTipoConsulta.AddItem "MENSAL"
cboTipoConsulta.AddItem "DATA"
cboTipoConsulta.AddItem "TIPO"
moCombo.AttachTo cboTipoConsulta
End Sub


Private Sub cboTipoConsulta_LostFocus()
If cboTipoConsulta.Text = "CLIENTE" Then
   cboCriterio.Visible = True
   lblCriterio.Visible = True
   mskDataCriterio.Visible = False
   lblData.Visible = False
   cmdCalendario.Visible = False
   cboMes.Visible = False
   lblMes.Visible = False
   cboAno.Visible = False
   lblAno.Visible = False
ElseIf cboTipoConsulta.Text = "PROFISSIONAL" Then
   cboCriterio.Visible = True
   lblCriterio.Visible = True
   mskDataCriterio.Visible = False
   lblData.Visible = False
   cmdCalendario.Visible = False
   cboMes.Visible = False
   lblMes.Visible = False
   cboAno.Visible = False
   lblAno.Visible = False
ElseIf cboTipoConsulta.Text = "MENSAL" Then
   cboCriterio.Visible = False
   lblCriterio.Visible = False
   mskDataCriterio.Visible = False
   lblData.Visible = False
   cmdCalendario.Visible = True
   cboMes.Visible = True
   lblMes.Visible = True
   cboAno.Visible = True
   lblAno.Visible = True
ElseIf cboTipoConsulta.Text = "DATA" Then
   cboCriterio.Visible = False
   lblCriterio.Visible = False
   mskDataCriterio.Visible = True
   lblData.Visible = True
   cmdCalendario.Visible = False
   cboMes.Visible = False
   lblMes.Visible = False
   cboAno.Visible = False
   lblAno.Visible = False
ElseIf cboTipoConsulta.Text = "TIPO" Then
   cboCriterio.Visible = True
   lblCriterio.Visible = True
   mskDataCriterio.Visible = False
   lblData.Visible = False
   cmdCalendario.Visible = False
   cboMes.Visible = False
   lblMes.Visible = False
   cboAno.Visible = False
   lblAno.Visible = False
End If
End Sub


Private Sub cmdAlterar_Click()
   If txtCodigo.Text = "" Then Exit Sub
   
   'Faz a atualização de forma direta e verifica se houve algum erro
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos
   Mostrar_Grid
End Sub

Private Sub cmdAvancar_Click()
Dim DataNova As Date
DataNova = Format(DateAdd("d", 1, mskDataConsulta), "dd/mm/yy")
mskDataConsulta.Text = Format(DataNova, "dd/mm/yy")
Mostrar_Grid
End Sub

Private Sub cmdCalendario_Click()
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
  
   mskDataCriterio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
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
  
   mskData = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCancelar_Click()
Limpar_Objetos
cmdNovo.Visible = True
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = False
cmdExcluir.Visible = False
frmReserva.Enabled = False
End Sub

Private Sub cmdCliente_Click()
Clientes_Cadastro.Show
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cboTipo.Text = "" Or txtCodCliente.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   Dim bRet As Boolean
   
   If txtCodigo.Text = "" Then Exit Sub
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir esse agendamento?", vbInformation + vbYesNo) = vbNo Then Exit Sub

   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM consultorio_agendamento WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   Limpar_Objetos
   Mostrar_Grid
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

'Set REL_consultorio_agendamento.Relatorio.Recordset = r

'REL_consultorio_agendamento.rfQuant.Caption = lblQuant.Caption
'REL_consultorio_agendamento.rfData.Caption = Format(mskDataConsulta.Text, "dd/mm/yy")

'REL_consultorio_agendamento.Relatorio.NomeImpressora = var_Impressora
'REL_consultorio_agendamento.Relatorio.Ativar
'Unload REL_consultorio_agendamento

Me.Show 1
End Sub

Private Sub cmdLocalizar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Dim INDICE As String

If cboTipoConsulta.Text = "DATA" Then
   INDICE = "DATA;"
ElseIf cboTipoConsulta.Text = "CLIENTE" Then
   INDICE = "COD_CLIENTE;"
ElseIf cboTipoConsulta.Text = "SALA" Then
   INDICE = "SALA;"
ElseIf cboTipoConsulta.Text = "PROFISSIONAL" Then
   INDICE = "COD_PROFISSIONAL;"
ElseIf cboTipoConsulta.Text = "TIPO" Then
   INDICE = "consultorio_agendamento.TIPO;"
Else
   INDICE = "DATA"
End If


If cboTipoConsulta.Text = "CLIENTE" Then
   If cboCriterio.Text = "" Then Exit Sub
   sSQL = "SELECT cliente.*, cliente.nome, consultorio_agendamento.codigo as varCod, consultorio_agendamento.*, funcionario.codigo, funcionario.nome as profissional " & _
   "FROM ((consultorio_agendamento INNER JOIN cliente ON cliente.codigo = consultorio_agendamento.cod_cliente) " & _
   "INNER JOIN funcionario ON consultorio_agendamento.cod_profissional = funcionario.codigo) " & _
   "WHERE cod_cliente = " & txtCodCriterio & " " & _
   "ORDER BY " & INDICE
ElseIf cboTipoConsulta.Text = "PROFISSIONAL" Then
   If cboCriterio.Text = "" Then Exit Sub
   sSQL = "SELECT cliente.*, cliente.nome, consultorio_agendamento.codigo as varCod, consultorio_agendamento.*, funcionario.codigo, funcionario.nome as profissional " & _
   "FROM ((consultorio_agendamento INNER JOIN cliente ON cliente.codigo = consultorio_agendamento.cod_cliente) " & _
   "INNER JOIN funcionario ON consultorio_agendamento.cod_profissional = funcionario.codigo) " & _
   "WHERE cod_profissional = " & txtCodCriterio & " " & _
   "ORDER BY " & INDICE
ElseIf cboTipoConsulta.Text = "MENSAL" Then
   If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
   sSQL = "SELECT cliente.*, cliente.nome, consultorio_agendamento.codigo as varCod, consultorio_agendamento.*, funcionario.codigo, funcionario.nome as profissional " & _
   "FROM ((consultorio_agendamento INNER JOIN cliente ON cliente.codigo = consultorio_agendamento.cod_cliente) " & _
   "INNER JOIN funcionario ON consultorio_agendamento.cod_profissional = funcionario.codigo) " & _
   "WHERE (MONTH(data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ") " & _
   "ORDER BY " & INDICE
ElseIf cboTipoConsulta.Text = "DATA" Then
   If Not IsDate(mskDataCriterio) Then Exit Sub
   sSQL = "SELECT cliente.*, cliente.nome, consultorio_agendamento.codigo as varCod, consultorio_agendamento.*, funcionario.codigo, funcionario.nome as profissional " & _
   "FROM ((consultorio_agendamento INNER JOIN cliente ON cliente.codigo = consultorio_agendamento.cod_cliente) " & _
   "INNER JOIN funcionario ON consultorio_agendamento.cod_profissional = funcionario.codigo) " & _
   "WHERE (data = CONVERT(DATETIME, '" & Format(mskDataCriterio.Text, ocDATA) & "', 103)) " & _
   "ORDER BY " & INDICE
ElseIf cboTipoConsulta.Text = "TIPO" Then
   If cboCriterio.Text = "" Then Exit Sub
   sSQL = "SELECT cliente.*, cliente.nome, consultorio_agendamento.codigo as varCod, consultorio_agendamento.*, funcionario.codigo, funcionario.nome as profissional " & _
   "FROM ((consultorio_agendamento INNER JOIN cliente ON cliente.codigo = consultorio_agendamento.cod_cliente) " & _
   "INNER JOIN funcionario ON consultorio_agendamento.cod_profissional = funcionario.codigo) " & _
   "WHERE consultorio_agendamento.tipo = '" & cboCriterio & "' " & _
   "ORDER BY " & INDICE
Else
   Exit Sub
End If

Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGridConsulta r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL

'MOSTRAR A QUANTIDADE REGISTROS
lblQuantConsulta.Caption = Format(totalRegistros, "00")
End Sub

Private Sub cmdNovo_Click()
Limpar_Objetos
frmReserva.Enabled = True
mskDataCadastro = Format(Date, "dd/mm/yy")
mskHoraCadastro = Format(Now, "hh:mm")
cmdNovo.Visible = False
cmdSalvar.Visible = True
cmdCancelar.Visible = True
cmdAlterar.Visible = False
cmdExcluir.Visible = False
cboCliente.SetFocus
End Sub

Private Sub cmdSair_Click()
  Unload Me
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lNovoCod As Long
   
   If txtCodCliente.Text = "" Or cboTipo.Text = "" Or mskData.Text = "" Then
      ShowMsg "Formulário incompleto!", vbInformation
      cboCliente.SetFocus
      Exit Sub
   End If
   
   'ADICIONAR REGISTRO
   lNovoCod = AutoNumeracao
   
   'Faz a inserção de forma direta e verifica se houve algum erro
   If Not Inserir_Dados(lNovoCod) Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos
   Mostrar_Grid
   Exit Sub
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub


Private Sub cmdVoltar_Click()
Dim DataNova As Date
DataNova = Format(DateAdd("d", -1, mskDataConsulta), "dd/mm/yy")
mskDataConsulta.Text = Format(DataNova, "dd/mm/yy")
Mostrar_Grid
End Sub

Private Sub Command12_Click()
Mostrar_Grid
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
frmReserva.Enabled = False
cboCriterio.Visible = True
lblCriterio.Visible = True
mskDataConsulta.Text = Format(DateAdd("d", Val(1), Date), "dd/mm/yy")
Mostrar_Grid
Set moCombo = New cComboHelper
End Sub
Private Sub FormatarGridConsulta(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Consulta
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 3800
      .ColWidth(3) = 2000
      .ColWidth(4) = 900
      .ColWidth(5) = 900
      .ColWidth(6) = 1700

      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "CLIENTE"
      .TextMatrix(0, 3) = "PROFISSIONAL"
      .TextMatrix(0, 4) = "HORA"
      .TextMatrix(0, 5) = "SALA"
      .TextMatrix(0, 6) = "TIPO"
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("varCod"))
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("nome"))
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("profissional"))
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("hora"), "hh:mm")
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("sala"), "00")
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("tipo"))
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   'lblValor.Caption = Format(SomaGrid(GridSuprimentos, 6), ocMONEY)
End Sub
Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 3800
      .ColWidth(3) = 2000
      .ColWidth(4) = 900
      .ColWidth(5) = 900
      .ColWidth(6) = 1700

      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "CLIENTE"
      .TextMatrix(0, 3) = "PROFISSIONAL"
      .TextMatrix(0, 4) = "HORA"
      .TextMatrix(0, 5) = "SALA"
      .TextMatrix(0, 6) = "TIPO"
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("varCod"))
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("nome"))
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("profissional"))
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("hora"), "hh:mm")
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("sala"), "00")
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("tipo"))
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   'lblValor.Caption = Format(SomaGrid(GridSuprimentos, 6), ocMONEY)
End Sub

Private Function Inserir_Dados(ByVal Codigo As Long) As Boolean
   'A inclusão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO consultorio_agendamento (codigo, cod_cliente, cod_profissional, cod_recepcionista, sala, data, hora, data_cadastro, hora_cadastro, tipo) VALUES (" & _
      Codigo & ", " & txtCodCliente.Text & ", " & txtCodProfissional.Text & ", " & txtCodRecepcionista.Text & ", '" & cboSala.Text & "', CONVERT(DATETIME, '" & _
      Format$(mskData.Text, ocDATA) & "', 103), " & IIf(mskHora.Text = "", "Null", "'" & Format$(mskHora.Text, ocHORA) & "'") & ", CONVERT(DATETIME, '" & _
      Format$(mskDataCadastro.Text, ocDATA) & "', 103), " & IIf(mskHoraCadastro.Text = "", "Null", "'" & Format$(mskHoraCadastro.Text, ocHORA) & "'") & ", '" & cboTipo.Text & "');"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function AutoNumeracao() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_reserva FROM consultorio_agendamento;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then lRet = r("cod_reserva") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao = lRet
End Function

Function Maiuscula(KeyAscii As Integer)
   If KeyAscii > 96 And KeyAscii < 123 Then
      KeyAscii = KeyAscii - 32
   End If
   Maiuscula = KeyAscii
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
   'TBcompromissos.Close
End Sub

Private Sub Grid_DblClick()
frmReserva.Enabled = True
cmdNovo.Visible = True
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = True
cmdExcluir.Visible = True
'Limpar_Objetos
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub


Private Sub mskData_KeyPress(KeyAscii As Integer)
mskData.Mask = "##/##/##"
End Sub


Private Sub mskDataCadastro_KeyPress(KeyAscii As Integer)
mskDataCadastro.Mask = "##/##/##"
End Sub


Private Sub mskDataConsulta_Validate(Cancel As Boolean)
If mskDataConsulta.Text = "__/__/__" Then
   mskDataConsulta.SetFocus
   Exit Sub
End If

If Not IsDate(mskDataConsulta) Then
   ShowMsg "DATA INVÁLIDA" & vbCrLf & "Digite a data novamente!", vbInformation
   mskDataConsulta.SetFocus
   mskDataConsulta.SelStart = 0
   mskDataConsulta.SelLength = Len(mskDataConsulta)
Exit Sub
End If
End Sub


Private Sub mskDataCriterio_KeyPress(KeyAscii As Integer)
mskDataCriterio.Mask = "##/##/##"
End Sub

Private Sub mskDataCriterio_Validate(Cancel As Boolean)
If mskDataCriterio.Text = "__/__/__" Then
   mskDataCriterio.SetFocus
   Exit Sub
End If

If Not IsDate(mskDataCriterio) Then
   ShowMsg "DATA INVÁLIDA" & vbCrLf & "Digite a data novamente!", vbInformation
   mskDataCriterio.SetFocus
   mskDataCriterio.SelStart = 0
   mskDataCriterio.SelLength = Len(mskDataCriterio)
   Exit Sub
End If
End Sub


Private Sub mskHora_KeyPress(KeyAscii As Integer)
mskHora.Mask = "##:##"
End Sub


Private Sub mskHoraCadastro_KeyPress(KeyAscii As Integer)
mskHoraCadastro.Mask = "##:##"
End Sub


Private Sub TxtCodCliente_Change()
If txtCodCliente.Text = "" Then Exit Sub
Dim r As ADODB.Recordset
If cmdAlterar.Visible = True Then
      sSQL = "SELECT codigo, nome FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If Not r.BOF Then cboCliente.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub

Private Sub txtCodigo_Change()
   If txtCodigo.Text = "" Then Exit Sub
   
   If cmdAlterar.Visible = True Then
      sSQL = "SELECT * FROM consultorio_agendamento WHERE (codigo = " & txtCodigo.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If Not r.BOF Then Mostrar_Dados r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub


Private Sub txtCodProfissional_Change()
If txtCodProfissional.Text = "" Then Exit Sub
Dim r As ADODB.Recordset
If cmdAlterar.Visible = True Then
      sSQL = "SELECT codigo, nome FROM funcionario WHERE (codigo = " & txtCodProfissional.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If Not r.BOF Then cboProfissional.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub


Private Sub txtCodRecepcionista_Change()
If txtCodRecepcionista.Text = "" Then Exit Sub
Dim r As ADODB.Recordset
If cmdAlterar.Visible = True Then
      sSQL = "SELECT codigo, nome FROM funcionario WHERE (codigo = " & txtCodRecepcionista.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If Not r.BOF Then cboRecepcionista.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub


