VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Caixa_Saida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SANGRIA DO CAIXA"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "Caixa_Saida.frx":0000
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   9480
      Top             =   960
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   10545
      TabIndex        =   34
      Top             =   0
      Width           =   10575
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Image Image2 
         Height          =   840
         Left            =   1740
         Picture         =   "Caixa_Saida.frx":23D2
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1020
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SANGRIA DO CAIXA"
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
         Left            =   2925
         TabIndex        =   35
         Top             =   300
         Width           =   3030
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5835
      Left            =   60
      TabIndex        =   23
      Top             =   1080
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   10292
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
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
      TabPicture(0)   =   "Caixa_Saida.frx":739D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExcluir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSalvar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSair"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Caixa_Saida.frx":73B9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblQuant"
      Tab(1).Control(1)=   "lblValor"
      Tab(1).Control(2)=   "frmConsulta"
      Tab(1).ControlCount=   3
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   60
         ScaleHeight     =   5145
         ScaleWidth      =   8565
         TabIndex        =   27
         Top             =   480
         Width           =   8595
         Begin VB.Frame frmCadastro 
            Enabled         =   0   'False
            Height          =   2655
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   8415
            Begin VB.TextBox txtCodFunc 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1860
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   840
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.ComboBox cboFuncionario 
               Height          =   315
               Left            =   120
               TabIndex        =   3
               Top             =   1140
               Width           =   2295
            End
            Begin VB.TextBox txtCodConta 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5400
               TabIndex        =   40
               Top             =   0
               Visible         =   0   'False
               Width           =   915
            End
            Begin ChamaleonBtn.chameleonButton cmdCal1 
               Height          =   315
               Left            =   6660
               TabIndex        =   39
               Tag             =   "Calendario"
               Top             =   1140
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
               MICON           =   "Caixa_Saida.frx":73D5
               PICN            =   "Caixa_Saida.frx":73F1
               PICH            =   "Caixa_Saida.frx":9744
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboFonte 
               Height          =   315
               ItemData        =   "Caixa_Saida.frx":BA97
               Left            =   4380
               List            =   "Caixa_Saida.frx":BA99
               TabIndex        =   6
               Top             =   1140
               Width           =   1395
            End
            Begin VB.ComboBox cboSubDesc 
               Height          =   315
               Left            =   120
               TabIndex        =   1
               Top             =   480
               Width           =   2535
            End
            Begin VB.ComboBox cboDesc 
               Height          =   315
               Left            =   2700
               TabIndex        =   2
               Top             =   480
               Width           =   5595
            End
            Begin VB.TextBox txtValor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   7020
               TabIndex        =   8
               Top             =   1140
               Width           =   1275
            End
            Begin VB.ComboBox cboSetor 
               Height          =   315
               ItemData        =   "Caixa_Saida.frx":BA9B
               Left            =   2460
               List            =   "Caixa_Saida.frx":BA9D
               TabIndex        =   5
               Top             =   1140
               Width           =   1875
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   315
               Left            =   5760
               TabIndex        =   7
               Top             =   1140
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Funcionário"
               Height          =   195
               Left            =   120
               TabIndex        =   48
               Top             =   900
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fonte"
               Height          =   195
               Left            =   4380
               TabIndex        =   36
               Top             =   900
               Width           =   405
            End
            Begin VB.Label Label7 
               Caption         =   "Sub-Descrição"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               Height          =   195
               Left            =   2700
               TabIndex        =   32
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data"
               Height          =   195
               Left            =   5760
               TabIndex        =   31
               Top             =   900
               Width           =   345
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   6960
               TabIndex        =   30
               Top             =   900
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Setor"
               Height          =   195
               Left            =   2460
               TabIndex        =   29
               Top             =   900
               Width           =   375
            End
         End
      End
      Begin VB.Frame frmConsulta 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5115
         Left            =   -74880
         TabIndex        =   24
         Top             =   360
         Width           =   10335
         Begin VB.ComboBox cboConsSetor 
            Height          =   315
            ItemData        =   "Caixa_Saida.frx":BA9F
            Left            =   4140
            List            =   "Caixa_Saida.frx":BAA1
            TabIndex        =   50
            Top             =   3120
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Frame Frame2 
            Caption         =   "Critérios"
            Height          =   915
            Left            =   8040
            TabIndex        =   45
            Top             =   180
            Width           =   2175
            Begin VB.ComboBox cboMES 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               ItemData        =   "Caixa_Saida.frx":BAA3
               Left            =   60
               List            =   "Caixa_Saida.frx":BAA5
               TabIndex        =   18
               Top             =   480
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.ComboBox cboAno 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1320
               Sorted          =   -1  'True
               TabIndex        =   19
               Top             =   480
               Visible         =   0   'False
               Width           =   795
            End
            Begin ChamaleonBtn.chameleonButton cmdConsData 
               Height          =   315
               Left            =   1020
               TabIndex        =   46
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
               MICON           =   "Caixa_Saida.frx":BAA7
               PICN            =   "Caixa_Saida.frx":BAC3
               PICH            =   "Caixa_Saida.frx":DE16
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
               Left            =   60
               TabIndex        =   17
               Top             =   480
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12648447
               PromptChar      =   "_"
            End
            Begin VB.Label lblCONmes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E&scolha o mês/ano:"
               Height          =   195
               Left            =   60
               TabIndex        =   47
               Top             =   240
               Visible         =   0   'False
               Width           =   1425
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Consulta"
            Height          =   915
            Left            =   60
            TabIndex        =   41
            Top             =   180
            Width           =   7935
            Begin VB.ComboBox cboConsFontes 
               Height          =   315
               ItemData        =   "Caixa_Saida.frx":10169
               Left            =   120
               List            =   "Caixa_Saida.frx":1016B
               TabIndex        =   52
               Top             =   480
               Width           =   1395
            End
            Begin VB.ComboBox cboCriterio 
               Height          =   315
               ItemData        =   "Caixa_Saida.frx":1016D
               Left            =   1560
               List            =   "Caixa_Saida.frx":1016F
               TabIndex        =   14
               Top             =   480
               Width           =   1695
            End
            Begin VB.ComboBox cboOrigem 
               Height          =   315
               ItemData        =   "Caixa_Saida.frx":10171
               Left            =   3300
               List            =   "Caixa_Saida.frx":10173
               TabIndex        =   15
               Top             =   480
               Width           =   1215
            End
            Begin VB.ComboBox cboIndice 
               Height          =   315
               ItemData        =   "Caixa_Saida.frx":10175
               Left            =   4560
               List            =   "Caixa_Saida.frx":10177
               TabIndex        =   16
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fontes:"
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
               TabIndex        =   53
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Left            =   1560
               TabIndex        =   44
               Top             =   240
               Width           =   660
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Origem:"
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
               Left            =   3300
               TabIndex        =   43
               Top             =   240
               Width           =   660
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ordem"
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
               Left            =   4560
               TabIndex        =   42
               Top             =   240
               Width           =   555
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibir 
            Height          =   315
            Left            =   7140
            TabIndex        =   20
            Top             =   1140
            Width           =   1515
            _ExtentX        =   2672
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
            MICON           =   "Caixa_Saida.frx":10179
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
            Left            =   8700
            TabIndex        =   21
            Top             =   1140
            Width           =   1515
            _ExtentX        =   2672
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
            MICON           =   "Caixa_Saida.frx":10195
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSFlexGridLib.MSFlexGrid GridSaidas 
            Height          =   3495
            Left            =   60
            TabIndex        =   37
            Top             =   1500
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   6165
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Setor"
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
            Left            =   4140
            TabIndex        =   51
            Top             =   2880
            Width           =   465
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   8700
         TabIndex        =   13
         Top             =   3780
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Caixa_Saida.frx":101B1
         PICN            =   "Caixa_Saida.frx":101CD
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
         Left            =   8700
         TabIndex        =   9
         Top             =   1140
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Salvar"
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
         MICON           =   "Caixa_Saida.frx":104E7
         PICN            =   "Caixa_Saida.frx":10503
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
         Left            =   8700
         TabIndex        =   11
         Top             =   2460
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Alterar"
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
         MICON           =   "Caixa_Saida.frx":12295
         PICN            =   "Caixa_Saida.frx":122B1
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
         Height          =   615
         Left            =   8700
         TabIndex        =   10
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Cancelar"
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
         MICON           =   "Caixa_Saida.frx":12B8B
         PICN            =   "Caixa_Saida.frx":12BA7
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
         Left            =   8700
         TabIndex        =   12
         Top             =   3120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Excluir"
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
         MICON           =   "Caixa_Saida.frx":19481
         PICN            =   "Caixa_Saida.frx":1949D
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
         Left            =   8700
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Novo"
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
         MICON           =   "Caixa_Saida.frx":197B7
         PICN            =   "Caixa_Saida.frx":197D3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblValor 
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -65340
         TabIndex        =   26
         Top             =   5520
         Width           =   690
      End
      Begin VB.Label lblQuant 
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74880
         TabIndex        =   25
         Top             =   5520
         Width           =   225
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   38
      Top             =   7005
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10795
            Text            =   "Desenv.: Online.Info Sistemas"
            TextSave        =   "Desenv.: Online.Info Sistemas"
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
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
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
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "00:00"
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
      Left            =   7860
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Caixa_Saida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Dim varCodCaixa As Long
Dim printSQL As String
Dim var_Cod_Retirada As Long
Dim ULTRAPASSOU_VALOR As Boolean
Dim CAIXA_FECHADO As Boolean
Dim varNovoCodSaldoRet As Integer
Dim var_Caixa As String
Dim varMaquina As String
Dim IMPRIMIR As Boolean
Dim var_ImpTermica As String
Dim var_ImpNormal As String
Dim varTipoRecPgto As String
Dim varTipoRecHaver As String
Private moCombo As cComboHelper


Private Function Atualizar_Dados() As Boolean
Dim sSQL As String

sSQL = "UPDATE caixa_saida SET " & _
   "descricao = '" & cboDesc.Text & "', " & _
   "setor = '" & cboSetor.Text & "', " & _
   "subdescricao = '" & cboSubDesc.Text & "', " & _
   "fonte = '" & cboFonte.Text & "', " & _
   "caixa = '" & StatusBar1.Panels(2).Text & "', " & _
   "codcaixa = " & varCodCaixa & ", " & _
   "data = CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
   "valor = " & Replace(CCur(txtValor.Text), ",", ".") & ", " & _
   "hora = '" & Format$(lblHora, ocHRMN) & "', COD_FUNCIONARIO = " & txtCodFunc & " "

sSQL = sSQL & "WHERE (codigo = " & txtCodigo.Text & ");"

Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Function Inserir_Dados(ByVal Codigo As Long) As Boolean
Dim sSQL As String

'Comando de inclusão
sSQL = "INSERT INTO caixa_saida (codigo, descricao, setor, subdescricao, fonte, data, valor, hora, caixa, codcaixa, COD_CONTA, MAQUINA, COD_FUNCIONARIO) VALUES (" & Codigo & ", '" & cboDesc.Text & "', '" & cboSetor.Text & "', '" & cboSubDesc.Text & "', '" & cboFonte.Text & "', CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & Replace(CCur(txtValor.Text), ",", ".") & ", '" & StatusBar1.Panels(4).Text & "', '" & StatusBar1.Panels(2).Text & "', " & varCodCaixa & ", 0, '" & varMaquina & "', " & txtCodFunc & ");"

'Retorna o resultado da inclusão
Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function AutoNumeracao() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 1
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_saida FROM caixa_saida;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod_saida") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao = lRet
End Function

Private Sub Campos_Brancos()
txtCodigo.Text = ""
mskData.Mask = ""
mskData.Text = ""
txtValor.Text = ""
cboDesc.Text = ""
cboDesc.Clear
cboSetor.Text = ""
cboSetor.Clear
cboSubDesc.Text = ""
txtCodConta.Text = ""
cboSubDesc.Clear
cboFonte.Text = "CAIXA ATUAL"
End Sub

Private Sub Criar_Novo()
Campos_Brancos
End Sub

Private Function Verificar_Caixa() As Integer
Dim sSQL As String
Dim r As ADODB.Recordset
Dim cxaStatus As Integer

cxaStatus = -1   'Não foi aberto
If cmdAlterar.Enabled = True Then
   sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskData.FormattedText, ocDATA) & "', 103)) AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
Else
   sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(5), ocDATA) & "', 103)) AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
End If
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then cxaStatus = CInt(ValidateNull(r("status")))   '0 = aberto, 1 = fechado
If r.State <> 0 Then r.Close
Set r = Nothing
Verificar_Caixa = cxaStatus
End Function

Private Sub Verificar_Valor_Saida()
Dim sSQL As String
Dim r As ADODB.Recordset

Dim Ent_Parcelas As Currency
Dim Ent_Haveres As Currency
Dim Ent_Entradas As Currency
Dim Soma_Entradas As Currency
Dim TotalSaidas As Currency
Dim Valor_Saida As Currency

Ent_Parcelas = 0
Ent_Entradas = 0

'parcelas
sSQL = "SELECT ISNULL(SUM(valor_final), 0) AS var_total FROM parcelas WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = '" & varCodCaixa & "') AND (FORMA_PGTO IN ('DINHEIRO', 'CHEQUE'));"
'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then Ent_Parcelas = r("var_total")

 'haveres
sSQL = "SELECT ISNULL(SUM(VALOR_HAVER), 0) AS varTotalHaver FROM parcelas_haver WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = '" & varCodCaixa & "') AND (FORMA_PGTO IN ('DINHEIRO', 'CHEQUE'));"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then Ent_Haveres = r("varTotalHaver")

 'suprimentos
sSQL = "SELECT ISNULL(SUM(valor), 0) AS varTotalSuprimentos FROM caixa_entrada WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = '" & varCodCaixa & "');"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then Ent_Entradas = r("varTotalSuprimentos")

'saidas
sSQL = "SELECT ISNULL(SUM(valor), 0) AS varTotalSaidas FROM caixa_saida WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = '" & varCodCaixa & "') and (FONTE = 'CAIXA ATUAL') ;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then TotalSaidas = r("varTotalSaidas")

If r.State <> 0 Then r.Close
Set r = Nothing

Soma_Entradas = Ent_Parcelas + Ent_Entradas + Ent_Haveres

Valor_Saida = txtValor.Text
Valor_Saida = Valor_Saida + TotalSaidas

ULTRAPASSOU_VALOR = False

If Valor_Saida > Soma_Entradas Then
   ShowMsg "O valor da sangria ultrapassou o seu saldo do caixa atual!", vbInformation
   ULTRAPASSOU_VALOR = True
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

Private Sub cboAno_LostFocus()
   If cboAno.Text = "" Then Exit Sub Else cmdExibir.SetFocus
End Sub

Private Sub cboConsFontes_Click()
cmdExibir_Click
End Sub

Private Sub cboConsFontes_GotFocus()
cboConsFontes.Clear
cboConsFontes.AddItem "TODOS"
cboConsFontes.AddItem "CAIXA ATUAL"
cboConsFontes.AddItem "SALDOS"
If cboConsFontes.Text = "" Then cboConsFontes.ListIndex = 0
moCombo.AttachTo cboConsFontes
End Sub


Private Sub cboConsSetor_GotFocus()
cboConsSetor.Clear

sSQL = "SELECT DISTINCT setor FROM setor ORDER BY setor;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboConsSetor.AddItem r("setor")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing


moCombo.AttachTo cboConsSetor
End Sub


Private Sub cboCriterio_Change()
cboCriterio_LostFocus
End Sub

Private Sub cboCriterio_Click()
cboCriterio_LostFocus
End Sub


Private Sub cboCriterio_GotFocus()
cboCriterio.Clear
cboCriterio.AddItem "TODOS"
cboCriterio.AddItem "CAIXA ATUAL"
cboCriterio.AddItem "MENSAL"
cboCriterio.AddItem "DATA"
'If cboCriterio.Text = "" Then cboCriterio.ListIndex = 0
moCombo.AttachTo cboCriterio
End Sub


Private Sub cboCriterio_LostFocus()
If cboCriterio.Text = "TODOS" Then
    mskConsData.Visible = False
    cmdConsData.Visible = False
    cboMES.Visible = False
    cboAno.Visible = False
    lblCONmes.Visible = False
ElseIf cboCriterio.Text = "CAIXA ATUAL" Then
    mskConsData.Visible = False
    cmdConsData.Visible = False
    cboMES.Visible = False
    cboAno.Visible = False
    lblCONmes.Visible = False
ElseIf cboCriterio.Text = "MENSAL" Then
    mskConsData.Visible = False
    cmdConsData.Visible = False
    cboMES.Visible = True
    cboAno.Visible = True
    lblCONmes.Visible = True
    lblCONmes.Caption = "Escolha o mês/ano:"
    cboMES.SetFocus
ElseIf cboCriterio.Text = "DATA" Then
    mskConsData.Visible = True
    cmdConsData.Visible = True
    cboMES.Visible = False
    cboAno.Visible = False
    lblCONmes.Visible = True
    lblCONmes.Caption = "Data:"
    If mskConsData.Visible = True Then mskConsData.SetFocus
End If
End Sub


Private Sub cboFonte_GotFocus()
cboFonte.Clear
cboFonte.AddItem "CAIXA ATUAL"
cboFonte.AddItem "SALDOS"
If cboFonte.Text = "" Then cboFonte.ListIndex = 0
moCombo.AttachTo cboFonte
End Sub

Private Sub cboFuncionario_GotFocus()
Dim varNomeAntes As String
Dim varCodAntes As String

varNomeAntes = cboFuncionario.Text
varCodAntes = txtCodFunc.Text

cboFuncionario.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM funcionario ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFuncionario.AddItem r("nome")
   cboFuncionario.ItemData(cboFuncionario.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboFuncionario.Text = varNomeAntes
txtCodFunc.Text = varCodAntes

SelectControl cboFuncionario
moCombo.AttachTo cboFuncionario
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cboIndice_GotFocus()
cboIndice.Clear
cboIndice.AddItem "DATA"
cboIndice.AddItem "DESCRIÇÃO"
If cboIndice.Text = "" Then cboIndice.ListIndex = 0
moCombo.AttachTo cboIndice
End Sub


Private Sub cboMes_GotFocus()
Dim vMes As Integer

cboMES.Clear

For vMes = 1 To 12
   cboMES.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMES
End Sub

Private Sub cboMes_LostFocus()
   If cboMES.Text = "" Then Exit Sub Else cboAno.SetFocus
End Sub

Private Sub cboOrigem_Click()
cmdExibir_Click
End Sub

Private Sub cboOrigem_GotFocus()
cboOrigem.Clear
cboOrigem.AddItem "TODOS"
cboOrigem.AddItem "SANGRIA"
cboOrigem.AddItem "CONTA"
If cboOrigem.Text = "" Then cboOrigem.ListIndex = 0
moCombo.AttachTo cboOrigem
End Sub


Private Sub cboSETOR_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAlterar_Click()
If txtCodigo.Text = "" Then Exit Sub
If cboFonte.Text = "SALDOS" Then MsgBox "Não é permitir alterar sangrias de saldo.", vbInformation, "Aviso do Sistema": Exit Sub
If txtCodFunc.Text = "" Then txtCodFunc.Text = "0"

'verificar situação do caixa
ConsultarCaixaAtual

If varCodCaixa = 0 Then
    MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim sSQL As String
Dim r As ADODB.Recordset
 
'verificar se é uma conta à pagar
If txtCodConta.Text <> "0" Then
    sSQL = "SELECT codigo, cod_conta FROM caixa_saida WHERE (cod_conta = " & txtCodConta.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then
       ShowMsg "Essa saída somente poderá ser alterada nas CONTAS À PAGAR!", vbExclamation
       Exit Sub
    End If
End If
 
'Faz a atualização de forma direta e verifica se houve algum erro
If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If
vCodFunc = 0
txtCodFunc.Text = ""
cboFuncionario.Text = ""
Campos_Brancos
Form_Load
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

Private Sub cmdCancelar_Click()
vCodFunc = 0
txtCodFunc.Text = ""
cboFuncionario.Text = ""
Campos_Brancos
Form_Load
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

Private Sub cmdExcluir_Click()
'Dim sSQL As String
Dim bRet As Boolean

If txtCodigo.Text = "" Then Exit Sub

'verificar se o caixa está aberto
ConsultarCaixaAtual

If varCodCaixa = 0 Then
    MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If ShowMsg("Excluir essa sangria?", vbInformation + vbYesNo) = vbNo Then Exit Sub

If cboFonte.Text = "CAIXA ATUAL" Then
    'verificar se é uma conta à pagar
    sSQL = "SELECT codigo, cod_conta FROM caixa_saida WHERE (cod_conta = " & txtCodConta.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then
        If r("cod_conta") <> 0 Then
            ShowMsg "Essa saída somente poderá ser excluída nas CONTAS À PAGAR!", vbExclamation
            Exit Sub
        End If
    End If
    
    'excluir a sangria
    sSQL = "DELETE FROM caixa_saida WHERE (codigo = " & txtCodigo.Text & ");"
    bRet = dbData.Execute(sSQL)
    
    If Not bRet Then
       ShowMsg "Não foi possível excluir o registro.", vbCritical
       Exit Sub
    End If
ElseIf cboFonte.Text = "SALDOS" Then
    'descobrir o codigo do saldo que  foi retirado o dinheiro
    sSQL = "SELECT COD_SALDO FROM caixa_saldo_retirada where TIPO = 'SANGRIA' AND COD_DESCRICAO = " & txtCodigo.Text & ";"
    Set r = dbData.OpenRecordset(sSQL)
    
    Dim varCodSaldo As Integer
    
    If Not r.BOF Then
        varCodSaldo = r("COD_SALDO")
    Else
        varCodSaldo = 0
    End If
    
    'descobrir o valor do saldo para voltar a quantia retirada
    sSQL = "SELECT ISNULL(RETIRADA, 0) as Ret, ISNULL(SALDO_ATUAL, 0) AS Sald FROM caixa_saldo where CODIGO = " & varCodSaldo & ";"
    Set r = dbData.OpenRecordset(sSQL)

    Dim varValorRetAtual As Currency
    Dim varValorRetNovas As Currency
    Dim varValorSaldoAtual As Currency
    Dim varValorSaldoNovo As Currency
        
    If Not r.BOF Then
        varValorRetAtual = r("Ret")
        varValorSaldoAtual = r("Sald")
    Else
        varValorRetAtual = 0
        varValorSaldoAtual = 0
    End If
    
    
    varValorRetNovas = varValorRetAtual - txtValor.Text
    varValorSaldoNovo = varValorSaldoAtual + txtValor.Text
    
    'atualizar o valor da retirada e saldo (acrescentar)
    dbData.Execute "UPDATE caixa_saldo SET RETIRADA = " & Replace(CCur(varValorRetNovas), ",", ".") & ", SALDO_ATUAL = " & Replace(CCur(varValorSaldoNovo), ",", ".") & " WHERE CODIGO = " & varCodSaldo & ";"
    
    'apagar a retirada de saldo
    dbData.Execute "DELETE caixa_saldo_retirada WHERE COD_SALDO = " & varCodSaldo & " and COD_DESCRICAO = " & txtCodigo.Text & ";"
    
    'apagar a sangria
    dbData.Execute "DELETE caixa_saida WHERE FONTE = 'SALDOS' AND CODIGO = " & txtCodigo.Text & ";"
End If
vCodFunc = 0
txtCodFunc.Text = ""
cboFuncionario.Text = ""
Campos_Brancos
Form_Load
End Sub

Private Function AutoNumeracao_Saldo_Retirada() As Long
sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS cod FROM caixa_saldo_retirada;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then varNovoCodSaldoRet = r("cod") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Function

Private Sub cmdExibir_Click()
If cboConsFontes.Text = "" Then cboConsFontes.Text = "CAIXA ATUAL"

Dim varCriterio As String
If cboCriterio.Text = "TODOS" Then
    varCriterio = " "
ElseIf cboCriterio.Text = "CAIXA ATUAL" Then
    varCriterio = "WHERE (codcaixa = " & varCodCaixa & ") and (CAIXA = '" & StatusBar1.Panels(2).Text & "') "
ElseIf cboCriterio.Text = "MENSAL" Then
   If cboMES.Text = "" Or cboMES.ListIndex = -1 Then Exit Sub
   If cboAno.Text = "" Or cboAno.ListIndex = -1 Then Exit Sub
    varCriterio = "WHERE (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ") "
ElseIf cboCriterio.Text = "DATA" Then
   If Not IsDate(mskConsData) Then Exit Sub
    varCriterio = "WHERE (data = CONVERT(DATETIME, '" & Format(mskConsData.Text, ocDATA) & "', 103)) and (CAIXA = '" & StatusBar1.Panels(2).Text & "') "
Else
    varCriterio = " "
End If

Dim varFonte As String
If cboCriterio.Text = "TODOS" Then
    If cboConsFontes.Text = "TODOS" Then
        varFonte = " "
    ElseIf cboConsFontes.Text = "CAIXA ATUAL" Then
        varFonte = "WHERE (FONTE = 'CAIXA ATUAL') "
    ElseIf cboConsFontes.Text = "SALDOS" Then
        varFonte = "WHERE (FONTE = 'SALDOS') "
    Else
        varFonte = " "
    End If
Else
    If cboConsFontes.Text = "TODOS" Then
        varFonte = " "
    ElseIf cboConsFontes.Text = "CAIXA ATUAL" Then
        varFonte = "AND (FONTE = 'CAIXA ATUAL') "
    ElseIf cboConsFontes.Text = "SALDOS" Then
        varFonte = "AND (FONTE = 'SALDOS') "
    Else
        varFonte = " "
    End If
End If

Dim varOrigem As String
If cboCriterio.Text = "TODOS" And cboConsFontes.Text = "TODOS" Then
    If cboOrigem.Text = "TODOS" Then
        varOrigem = " "
    ElseIf cboOrigem.Text = "SANGRIA" Then
        varOrigem = "WHERE (COD_CONTA = 0) "
    ElseIf cboOrigem.Text = "CONTA" Then
        varOrigem = "WHERE (COD_CONTA <> 0) "
    Else
        varOrigem = " "
    End If
Else
    If cboOrigem.Text = "TODOS" Then
        varOrigem = " "
    ElseIf cboOrigem.Text = "SANGRIA" Then
        varOrigem = "AND (COD_CONTA = 0) "
    ElseIf cboOrigem.Text = "CONTA" Then
        varOrigem = "AND (COD_CONTA <> 0) "
    Else
        varOrigem = " "
    End If
End If

Dim vConsSetor As String
If cboConsSetor.Text = "TODOS" Then
    vConsSetor = " "
ElseIf cboConsSetor.Text <> "TODOS" Then
    vConsSetor = " and (SETOR = '" & cboConsSetor.Text & "') "
End If

Dim vIndice As String
If cboIndice.Text = "DATA" Then
    vIndice = " order by DATA"
ElseIf cboIndice.Text = "DESCRIÇÃO" Then
    vIndice = " order by DESCRICAO"
End If

Dim totalRegistros As Long

sSQL = "SELECT * FROM caixa_saida  " & varCriterio & varFonte & varOrigem & vConsSetor & vIndice & ""

Set r = dbData.OpenRecordset(sSQL, totalRegistros)
FormatarGrid r

Debug.Print sSQL

printSQL = sSQL

If r.State <> 0 Then r.Close
Set r = Nothing

'MOSTRAR A QUANTIDADE REGISTROS
lblQuant.Caption = Format(totalRegistros, "00")
End Sub

Private Sub cmdImprimir_Click()
cmdExibir_Click
Dim r As ADODB.Recordset

'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Caixa_Sangria.Relatorio.Recordset = r
REL_Caixa_Sangria.lblTitulo.Caption = "RELATÓRIO DE CAIXA - SANGRIAS"

REL_Caixa_Sangria.rfCriterio.Caption = cboCriterio.Text
REL_Caixa_Sangria.rfFontes.Caption = cboConsFontes.Text
REL_Caixa_Sangria.rfTitOrigem.Caption = cboOrigem.Text

If cboCriterio.Text = "TODOS" Then
    REL_Caixa_Sangria.rfCriterioDesc.Caption = ""
ElseIf cboCriterio.Text = "CAIXA ATUAL" Then
    REL_Caixa_Sangria.rfCriterioDesc.Caption = var_Caixa & " - " & Format(varCodCaixa, "0000")
ElseIf cboCriterio.Text = "MENSAL" Then
    REL_Caixa_Sangria.rfCriterioDesc.Caption = cboMES.Text & " / " & cboAno.Text
ElseIf cboCriterio.Text = "DATA" Then
    REL_Caixa_Sangria.rfCriterioDesc.Caption = mskConsData.Text
End If

REL_Caixa_Sangria.rfQuant.Caption = lblQuant.Caption
REL_Caixa_Sangria.rfSubTotal.Caption = Format(lblValor.Caption, ocMONEY)

'REL_Caixa_Sangria.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
'REL_Caixa_Sangria.rfCodCaixa.Caption = varFluxoCodCaixa
'REL_Caixa_Sangria.rfNomeCaixa.Caption = varFluxoNomeCaixa

REL_Caixa_Sangria.Relatorio.NomeImpressora = var_Impressora
REL_Caixa_Sangria.Relatorio.Ativar
Unload REL_Caixa_Sangria

Me.Show 1
End Sub

Private Sub cmdNovo_Click()
frmCadastro.Enabled = True
Criar_Novo
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub ConsultarCaixaAtual()
'Dim sSQL As String
'Dim r As ADODB.Recordset
sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    varCodCaixa = ValidateNull(r("codcaixa"))
Else
    varCodCaixa = 0
End If
End Sub

Private Function Inserir_Dados_Saldo_Retirada(ByVal CodCaixa As Long, ByVal CodLancto As Long) As Boolean

End Function

Private Sub cboDesc_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cboSubDesc.Text = "VALE" Then
      cboDesc.Clear
      
      sSQL = "SELECT DISTINCT nome, codigo FROM funcionario ORDER BY nome;"
      Set r = dbData.OpenRecordset(sSQL)
      
      Do While Not r.EOF
         cboDesc.AddItem r("nome")
         cboDesc.ItemData(cboDesc.NewIndex) = r("codigo")
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
   Else
      cboDesc.Clear
      
      sSQL = "SELECT DISTINCT descricao FROM caixa_saida ORDER BY descricao;"
      Set r = dbData.OpenRecordset(sSQL)
      
      Do While Not r.EOF
         cboDesc.AddItem r("descricao")
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
   
   End If

   moCombo.AttachTo cboDesc
End Sub

Private Sub cboDesc_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSetor_GotFocus()
cboSetor.Clear

If cboSubDesc.Text = "VALE" And cboDesc.Text <> "" Then
   sSQL = "SELECT DISTINCT setor FROM funcionario WHERE (nome = '" & cboDesc.Text & "') ORDER BY setor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboSetor.AddItem r("setor")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   cboSetor.ListIndex = 0

Else
   sSQL = "SELECT DISTINCT setor FROM setor ORDER BY setor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboSetor.AddItem r("setor")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
End If

If cboSetor.ListCount <> 0 Then cboSetor.ListIndex = 0
moCombo.AttachTo cboSetor
End Sub

Private Sub cboSubDesc_GotFocus()
   cboSubDesc.Clear
   cboSubDesc.AddItem "ALIMENTAÇÃO"
   cboSubDesc.AddItem "COMPRA"
   cboSubDesc.AddItem "SERVIÇO"
   cboSubDesc.AddItem "LOCAÇÃO"
   cboSubDesc.AddItem "FARMACIA"
   cboSubDesc.AddItem "PGTO DE CONTA"
   cboSubDesc.AddItem "HAVER EM CONTA"
   cboSubDesc.AddItem "DIVERSOS"
   cboSubDesc.AddItem "PESSOAL"
   moCombo.AttachTo cboSubDesc
End Sub

Private Sub cboSubDesc_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdSalvar_Click()
If txtValor.Text = "" Or cboSubDesc.Text = "" Or cboDesc.Text = "" Or cboFonte.Text = "" Then
   ShowMsg "Formulário incompleto!", vbInformation
   cboSubDesc.SetFocus
   Exit Sub
End If

If txtCodFunc.Text = "" Then txtCodFunc.Text = "0"

Dim lNovoCod As Long
   
If cboFonte.Text = "CAIXA ATUAL" Then
    'verificar se o caixa está aberto
    If varCodCaixa = 0 Then
        MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
   
    'verificar o saldo do caixa
    Verificar_Valor_Saida
    If ULTRAPASSOU_VALOR = True Then Exit Sub
    
    'criar uma sangria
    'Dim lNovoCod As Long
    lNovoCod = AutoNumeracao
    
    'Faz a inserção de forma direta e verifica se houve algum erro
    If Not Inserir_Dados(lNovoCod) Then
       ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
       Exit Sub
    End If
      
ElseIf cboFonte.Text = "SALDOS" Then
    'pegar o valor do ultimo do saldos
    sSQL = "SELECT top 1 SALDO_ATUAL FROM caixa_saldo order by codigo desc;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Dim varValorUltimoSaldo As Currency
    If Not r.BOF Then varValorUltimoSaldo = r("SALDO_ATUAL")
    
    Dim varValorSaida As Currency
    varValorSaida = txtValor.Text
    
    If varValorSaida > varValorUltimoSaldo Then MsgBox "O valor da sangria é maior que seu saldo atual", vbInformation, "Aviso do Sistema": Exit Sub
    
    'criar uma sangria
    lNovoCod = AutoNumeracao
    
    'inserir dados na tabela caixa_saida
    If Not Inserir_Dados(lNovoCod) Then
       ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
       Exit Sub
    End If
    
    'pegar o ultimo codigo do saldos
    sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM caixa_saldo;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Dim varCodSaldo As Integer
    If Not r.BOF Then varCodSaldo = r("cod")

    'criar um retirada do saldo
    AutoNumeracao_Saldo_Retirada

    dbData.Execute "INSERT INTO caixa_saldo_retirada (codigo, cod_saldo, tipo, data, cod_descricao, valor, descricao) VALUES (" & _
       varNovoCodSaldoRet & ", " & varCodSaldo & ", 'SANGRIA', CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
       lNovoCod & ", " & Replace(CCur(txtValor.Text), ",", ".") & ", '" & cboSubDesc.Text & " / " & cboDesc.Text & "');"
    
    sSQL = "UPDATE caixa_saldo SET " & _
       "RETIRADA = RETIRADA + " & Replace(CCur(txtValor), ",", ".") & ", " & _
       "SALDO_ATUAL = SALDO_ANTERIOR + ENTRADA - (RETIRADA + " & Replace(CCur(txtValor), ",", ".") & ")" & _
       " WHERE (codigo = " & varCodSaldo & ") ;"
     'Debug.Print sSQL
    dbData.Execute sSQL


End If

If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    If varTipoRecPgto = "CUPOM" Then
        Imprimir_ReciboCupom
    Else
        Imprimir_ReciboFolha
    End If
End If
vCodFunc = 0
txtCodFunc.Text = ""
cboFuncionario.Text = ""
Campos_Brancos
Form_Load
End Sub
Private Sub Imprimir_ReciboFolha()
Dim rUsuario As ADODB.Recordset
Dim rEmpresa As ADODB.Recordset
Dim vCidadeUF As String

'tabela empresa
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set rEmpresa = dbData.OpenRecordset(sSQL)
vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")

If txtCodFunc.Text = "" Then txtCodFunc.Text = "1"
Set rUsuario = dbData.OpenRecordset("SELECT codigo, login FROM Usuario WHERE  (codigo = " & txtCodFunc.Text & ");")

Me.Hide

With REL_ReciboSangria
    .txtSubDescricao.Caption = UCase(cboSubDesc.Text)
    .txtDescricao.Caption = UCase(cboDesc.Text)
    .txtFonte.Caption = UCase(cboFonte.Text)
    .txtUsuario.Caption = UCase(rUsuario("login"))
    .txtFormaPgto.Caption = UCase("DINHEIRO")
    .txtValor.Caption = UCase(NumeroExtenso(txtValor.Text, True))
    .txthead.Caption = "R$ " & Format(txtValor.Text, "##,##0.00")
    '.txtProveniente.Caption = "Pagamento da " & txtNumParcela.Text & "ª parcela do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
    .txtData.Caption = "" & vCidadeUF & ", " & Day(mskData) & " de " & MonthName(Month(mskData)) & " de " & Year(mskData)
    
    .Relatorio.NumeroRegistros = 1
    .Relatorio.NomeImpressora = var_ImpNormal
    .Relatorio.Ativar
End With

Unload REL_ReciboSangria
Me.Show
End Sub


Private Sub Imprimir_ReciboCupom()
'On Error GoTo Tratar_Erro
Dim sSQL As String
Dim rEmpresa As ADODB.Recordset

Dim i As Integer
Dim f As Integer

'tabela empresa
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set rEmpresa = dbData.OpenRecordset(sSQL)

'Recupera um número de arquivo disponível
f = FreeFile()
   
  'pegar o nome da impressora no ini
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
   Set oIni = Nothing
   
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

   With Printer
      .ScaleMode = vbPixels
      '.PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
      
      For i = 1 To 6
         Printer.Print " "
      Next
      
      .ScaleMode = vbCentimeters
      .FontName = "courier new"
      '.PrintQuality = vbPRPQHigh
      

      Fonte 10, True, False
      Printer.Print Tab((35 - Len(rEmpresa("fantasia"))) / 2); rEmpresa("fantasia")   'Esse /2 é p/ centralizar
      Fonte 10, False, False
      Printer.Print Tab((35 - Len(rEmpresa("razao"))) / 2); rEmpresa("razao")
      Fonte 8, False, False
      Printer.Print rEmpresa("endereco") & ", " & rEmpresa("cidade") & "-" & rEmpresa("estado")
      Printer.Print "FONE: "; rEmpresa("telefone")                                        '& " - (89) 9986-3739"
      Fonte 8, False, False
      Printer.Print "CNPJ:"; rEmpresa("cnpj") & "  IE:" & rEmpresa("ie")
      Fonte 8, False, False
      Printer.Print String(40, "-")
      
       For i = 1 To 2
         Printer.Print " "
      Next
      
      Fonte 10, True, False
      Printer.Print Tab((40 - Len("R E C I B O  D E  S A Í D A")) / 2); "R E C I B O  D E  S A Í D A"
      
      For i = 1 To 2
         Printer.Print " "
      Next
  
    
      Fonte 8, False, False
      Printer.Print Tab(2); "Subdescrição: "
      Fonte 8, True, False
      Printer.Print Tab(2); cboSubDesc.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Fonte 8, False, False
      Printer.Print Tab(2); "Descrição: "
      Fonte 8, True, False
      Printer.Print Tab(2); cboDesc.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next

      Fonte 8, False, False
      Printer.Print Tab(2); "Setor: "
      Fonte 8, True, False
      Printer.Print Tab(2); cboSetor.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next

      Fonte 8, False, False
      Printer.Print Tab(2); "Fonte: "
      Fonte 8, True, False
      Printer.Print Tab(2); cboFonte.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Fonte 8, False, False
      Printer.Print Tab(2); "Data: "
      Fonte 8, True, False
      Printer.Print Tab(2); mskData.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next

      Fonte 8, False, False
      Printer.Print Tab(2); "Valor: "
      Fonte 8, True, False
      Printer.Print Tab(2); txtValor.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
     
      For i = 1 To 3
            Printer.Print " "
      Next
      
      Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
      Printer.Print Tab((40 - Len("Assinatura")) / 2); "Assinatura"
      

     
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

If Not rEmpresa Is Nothing Then If rEmpresa.State <> 0 Then rEmpresa.Close

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


Private Sub Form_Load()
SSTab1.Tab = 0
cmdNovo.Enabled = True
frmCadastro.Enabled = False
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False

cboCriterio.AddItem "TODOS"
cboCriterio.AddItem "CAIXA ATUAL"
cboCriterio.AddItem "MENSAL"
cboCriterio.AddItem "DATA"
If cboCriterio.Text = "" Then cboCriterio.ListIndex = 3

cboFonte.AddItem "CAIXA ATUAL"
cboFonte.AddItem "SALDOS"
If cboFonte.Text = "" Then cboFonte.ListIndex = 0

 mskConsData.Text = Format(Date, "dd/mm/yy")

'colocar o nome da maquina na barra de status
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
varMaquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
'Set oIni = Nothing

StatusBar1.Panels(2).Text = var_Caixa

'verificar se o caixa está aberto
ConsultarCaixaAtual

StatusBar1.Panels(3).Text = varCodCaixa
StatusBar1.Panels(4).Text = Format(Now, "hh:mm")
StatusBar1.Panels(5).Text = Format(Date, "dd/mm/yy")

'cboCriterio.Text = "DATA"
cboConsFontes.Text = "TODOS"
cboOrigem.Text = "TODOS"
cboIndice.Text = "DATA"
cboConsSetor.Text = "TODOS"

'nome da maquina
var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
Set oIni = Nothing  'fecha o ini

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

'tipo de recibo de pagamento
Set oCfg = sysConfig("TIPORECPGTO")
varTipoRecPgto = oCfg.Value
Set oCfg = Nothing
Set oIni = Nothing

cmdExibir_Click
Set moCombo = New cComboHelper
End Sub

Private Sub Mostrar_Saida(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then
   cboDesc.Text = rTabela("descricao")
   txtCodFunc.Text = rTabela("COD_FUNCIONARIO")
   cboSetor.Text = rTabela("setor")
   cboSubDesc.Text = rTabela("subdescricao")
   txtCodConta.Text = ValidateNull(rTabela("cod_conta"))
   cboFonte.Text = ValidateNull(rTabela("fonte"))
   mskData.Text = Format(rTabela("data"), "dd/mm/yy")
   txtValor.Text = Format(rTabela("valor"), ocMONEY)
   If cboFonte.Text = "SALDOS" Then cmdAlterar.Enabled = False
End If
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

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer

With GridSaidas
    .Clear
    .Cols = 11
    .rows = 2
        
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 800
    .ColWidth(3) = 650
    .ColWidth(4) = 650
    .ColWidth(5) = 1250
    .ColWidth(6) = 3000
    .ColWidth(7) = 700
    .ColWidth(8) = 700
    .ColWidth(9) = 1150
    .ColWidth(10) = 1000
   
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "DATA"
   .TextMatrix(0, 3) = "HORA"
   .TextMatrix(0, 4) = "FUNC"
   .TextMatrix(0, 5) = "SUBDESC"
   .TextMatrix(0, 6) = "DESC"
   .TextMatrix(0, 7) = "CAIXA"
   .TextMatrix(0, 8) = "CÓD."
   .TextMatrix(0, 9) = "FONTE"
   .TextMatrix(0, 10) = "VALOR"
   .Redraw = False
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.rows - 1, 2) = Format(rTabela("data"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = Format(rTabela("hora"), ocHRMN)
         .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("COD_FUNCIONARIO"))
         .TextMatrix(.rows - 1, 5) = rTabela("subdescricao")
         .TextMatrix(.rows - 1, 6) = rTabela("descricao")
         .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("CAIXA"))
         .TextMatrix(.rows - 1, 8) = Format(rTabela("CODCAIXA"), "000000")
         .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("FONTE"))
         .TextMatrix(.rows - 1, 10) = Format(rTabela("valor"), ocMONEY)
         
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
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

lblValor.Caption = Format(SomaGrid(GridSaidas, 10), ocMONEY)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If vChamouCaixa = "PDV" Then
    Me.Hide
    'PDV.Show  'desativei somente para geerar o online comerce
Else
    Me.Hide
    'PDV.Show 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'HabilitaObjetosVenda False
Set moCombo = Nothing
End Sub

Private Sub GridSaidas_DblClick()
'verificar ser o caixa do haver selecionado está em aberto
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (codcaixa = " & varCodCaixa & ") and (caixa = '" & GridSaidas.TextMatrix(GridSaidas.Row, 2) & "') and caixa_dia.status = 1;"
Set r = dbData.OpenRecordset(sSQL)

If r.RecordCount > 0 Then
    MsgBox "O caixa onde essa sangria foi adicionado encontra-se fechado!", vbInformation, "Aviso do Sistema"
    r.Close
    Set r = Nothing
    Exit Sub
End If

'INICIO DA ROTINA
frmCadastro.Enabled = True
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
txtCodigo.Text = ""
txtCodigo.Text = (GridSaidas.TextMatrix(GridSaidas.Row, 1))
End Sub

Private Sub mskConsData_GotFocus()
   SelectControl mskConsData
End Sub

Private Sub mskConsData_KeyPress(KeyAscii As Integer)
mskConsData.Mask = "##/##/##"
End Sub

Private Sub mskConsData_LostFocus()
   If mskConsData.Text = "" Or mskConsData.Text = "__/__/__" Then
      mskConsData.Mask = ""
      mskConsData.Text = ""
   Else
      If IsDate(mskConsData.Text) Then
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskConsData.SetFocus
         SelectControl mskConsData
      End If
   End If
End Sub

Private Sub mskData_GotFocus()
   If cmdAlterar.Enabled = False Then mskData.Text = Format(Date, "dd/mm/yy")
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
         SelectControl mskData
      End If
   End If
End Sub

Private Sub cboFuncionario_LostFocus()
On Error GoTo TrataErro

If cboFuncionario.Text = "" Then txtCodFunc.Text = "": Exit Sub

'If cmdAlterar.Enabled = False Then
    'If cboFuncionario.ListIndex = -1 Then
    '    txtCodFunc.Text = ""
    '    cboFuncionario.Text = ""
    '    Exit Sub
    'Else
        txtCodFunc = cboFuncionario.ItemData(cboFuncionario.ListIndex)
   'End If
'End If

Exit Sub

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub Timer1_Timer()
   lblHora.Caption = Format(Time, "hh:mm")
End Sub

Private Sub txtCodFunc_Change()
If txtCodFunc.Text = "" Then Exit Sub

'If cmdAlterar.Enabled = True Then
   sSQL = "SELECT codigo, nome FROM funcionario WHERE (codigo = " & txtCodFunc.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cboFuncionario.Text = r("nome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
'End If
End Sub
Private Sub txtCodigo_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

If cmdAlterar.Enabled = True Then
   sSQL = "SELECT * FROM caixa_saida WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then Mostrar_Saida r
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If

SSTab1.Tab = 0
End Sub

Private Sub txtValor_GotFocus()
SelectControl txtValor
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtValor_LostFocus()
If txtValor.Text = "" Then
   txtValor.Text = Format(0, ocMONEY)
Else
   txtValor.Text = Format(txtValor, ocMONEY)
End If
End Sub
