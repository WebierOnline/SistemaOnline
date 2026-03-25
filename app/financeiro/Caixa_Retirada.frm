VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Caixa_Retirada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RETIRADAS DO CAIXA"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "Caixa_Retirada.frx":0000
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10005
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
      ScaleWidth      =   9825
      TabIndex        =   22
      Top             =   0
      Width           =   9855
      Begin VB.Image Image2 
         Height          =   795
         Left            =   1560
         Picture         =   "Caixa_Retirada.frx":23D2
         Stretch         =   -1  'True
         Top             =   60
         Width           =   795
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RETIRADAS DO CAIXA"
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
         Left            =   2460
         TabIndex        =   23
         Top             =   300
         Width           =   3420
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5835
      Left            =   60
      TabIndex        =   11
      Top             =   1080
      Width           =   9885
      _ExtentX        =   17436
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
      TabPicture(0)   =   "Caixa_Retirada.frx":641E
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
      TabPicture(1)   =   "Caixa_Retirada.frx":643A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmConsulta"
      Tab(1).Control(1)=   "lblValor"
      Tab(1).Control(2)=   "lblQuant"
      Tab(1).ControlCount=   3
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   60
         ScaleHeight     =   5145
         ScaleWidth      =   7725
         TabIndex        =   15
         Top             =   480
         Width           =   7755
         Begin VB.Frame frmCadastro 
            Enabled         =   0   'False
            Height          =   1635
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   7575
            Begin VB.TextBox txtCodFunc 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6480
               TabIndex        =   27
               Top             =   240
               Visible         =   0   'False
               Width           =   915
            End
            Begin ChamaleonBtn.chameleonButton cmdCal1 
               Height          =   315
               Left            =   1020
               TabIndex        =   26
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
               MICON           =   "Caixa_Retirada.frx":6456
               PICN            =   "Caixa_Retirada.frx":6472
               PICH            =   "Caixa_Retirada.frx":87C5
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox txtCodigo 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6660
               TabIndex        =   21
               Top             =   -60
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.ComboBox cboSubDesc 
               Height          =   315
               Left            =   120
               TabIndex        =   1
               Top             =   480
               Width           =   2295
            End
            Begin VB.ComboBox cboFuncionario 
               Height          =   315
               Left            =   2460
               TabIndex        =   2
               Top             =   480
               Width           =   4995
            End
            Begin VB.TextBox txtValor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2100
               TabIndex        =   5
               Top             =   1140
               Width           =   1455
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   315
               Left            =   120
               TabIndex        =   3
               Top             =   1140
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskHora 
               Height          =   315
               Left            =   1320
               TabIndex        =   4
               Top             =   1140
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hora"
               Height          =   195
               Left            =   1320
               TabIndex        =   30
               Top             =   900
               Width           =   345
            End
            Begin VB.Label Label7 
               Caption         =   "Descrição"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Funcionário"
               Height          =   195
               Left            =   2460
               TabIndex        =   19
               Top             =   240
               Width           =   825
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   900
               Width           =   345
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   2040
               TabIndex        =   17
               Top             =   900
               Width           =   360
            End
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
            Left            =   3840
            TabIndex        =   31
            Top             =   2280
            Visible         =   0   'False
            Width           =   495
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
         TabIndex        =   12
         Top             =   360
         Width           =   9675
         Begin VB.Frame Frame2 
            Caption         =   "Critérios"
            Height          =   915
            Left            =   4740
            TabIndex        =   39
            Top             =   180
            Width           =   3555
            Begin VB.ComboBox cboMES 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               ItemData        =   "Caixa_Retirada.frx":AB18
               Left            =   120
               List            =   "Caixa_Retirada.frx":AB1A
               TabIndex        =   42
               Top             =   480
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.ComboBox cboAno 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1860
               Sorted          =   -1  'True
               TabIndex        =   41
               Top             =   480
               Visible         =   0   'False
               Width           =   1335
            End
            Begin ChamaleonBtn.chameleonButton cmdConsData 
               Height          =   315
               Left            =   1140
               TabIndex        =   40
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
               MICON           =   "Caixa_Retirada.frx":AB1C
               PICN            =   "Caixa_Retirada.frx":AB38
               PICH            =   "Caixa_Retirada.frx":CE8B
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
               TabIndex        =   44
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
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Visible         =   0   'False
               Width           =   1425
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Consulta"
            Height          =   915
            Left            =   60
            TabIndex        =   32
            Top             =   180
            Width           =   4635
            Begin VB.ComboBox cboCriterio 
               Height          =   315
               ItemData        =   "Caixa_Retirada.frx":F1DE
               Left            =   120
               List            =   "Caixa_Retirada.frx":F1E0
               TabIndex        =   35
               Top             =   480
               Width           =   1575
            End
            Begin VB.ComboBox cboIndice 
               Height          =   315
               ItemData        =   "Caixa_Retirada.frx":F1E2
               Left            =   3300
               List            =   "Caixa_Retirada.frx":F1E4
               TabIndex        =   34
               Top             =   480
               Width           =   1215
            End
            Begin VB.ComboBox cboCaixa 
               Height          =   315
               Left            =   1740
               TabIndex        =   33
               Top             =   480
               Width           =   1515
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
               Left            =   120
               TabIndex        =   38
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
               Left            =   3300
               TabIndex        =   37
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label5 
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
               Left            =   1740
               TabIndex        =   36
               Top             =   240
               Width           =   480
            End
         End
         Begin MSFlexGridLib.MSFlexGrid GridSaidas 
            Height          =   3855
            Left            =   60
            TabIndex        =   24
            Top             =   1140
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   6800
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdExibir 
            Height          =   435
            Left            =   8340
            TabIndex        =   28
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   767
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
            MICON           =   "Caixa_Retirada.frx":F1E6
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
            Height          =   375
            Left            =   8340
            TabIndex        =   29
            Top             =   720
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
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
            MICON           =   "Caixa_Retirada.frx":F202
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
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   7920
         TabIndex        =   10
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
         MICON           =   "Caixa_Retirada.frx":F21E
         PICN            =   "Caixa_Retirada.frx":F23A
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
         Left            =   7920
         TabIndex        =   6
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
         MICON           =   "Caixa_Retirada.frx":F554
         PICN            =   "Caixa_Retirada.frx":F570
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
         Left            =   7920
         TabIndex        =   8
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
         MICON           =   "Caixa_Retirada.frx":11302
         PICN            =   "Caixa_Retirada.frx":1131E
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
         Left            =   7920
         TabIndex        =   7
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
         MICON           =   "Caixa_Retirada.frx":11BF8
         PICN            =   "Caixa_Retirada.frx":11C14
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
         Left            =   7920
         TabIndex        =   9
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
         MICON           =   "Caixa_Retirada.frx":184EE
         PICN            =   "Caixa_Retirada.frx":1850A
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
         Left            =   7920
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
         MICON           =   "Caixa_Retirada.frx":18824
         PICN            =   "Caixa_Retirada.frx":18840
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
         Left            =   -65940
         TabIndex        =   14
         Top             =   5550
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
         TabIndex        =   13
         Top             =   5520
         Width           =   225
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   25
      Top             =   7005
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9551
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
End
Attribute VB_Name = "Caixa_Retirada"
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
Private moCombo As cComboHelper
Dim IMPRIMIR As Boolean
Dim var_ImpTermica As String
Dim var_ImpNormal As String
Dim varTipoRecPgto As String
Dim varTipoRecHaver As String




Private Function Atualizar_Dados() As Boolean
Dim sSQL As String
If txtCodFunc.Text = "" Then txtCodFunc.Text = 0

sSQL = "UPDATE caixa_retirada SET " & _
   "descricao = '" & cboSubDesc.Text & "', " & _
   "cod_funcionario = '" & txtCodFunc.Text & "', " & _
   "caixa = '" & StatusBar1.Panels(2).Text & "', " & _
   "codcaixa = " & varCodCaixa & ", " & _
   "data = CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
   "valor = " & Replace(CCur(txtValor.Text), ",", ".") & ", " & _
   "hora = '" & Format$(mskHora.Text, ocHRMN) & "' "

sSQL = sSQL & "WHERE (codigo = " & txtCodigo.Text & ");"

Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Function Inserir_Dados(ByVal Codigo As Long) As Boolean
Dim sSQL As String
If txtCodFunc.Text = "" Then txtCodFunc.Text = 0

'Comando de inclusão
sSQL = "INSERT INTO caixa_retirada (codigo, descricao, data, valor, hora, caixa, codcaixa, COD_FUNCIONARIO) VALUES (" & Codigo & ", '" & cboSubDesc.Text & "', CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & Replace(CCur(txtValor.Text), ",", ".") & ", '" & StatusBar1.Panels(4).Text & "', '" & StatusBar1.Panels(2).Text & "', " & varCodCaixa & ", " & txtCodFunc & ");"

'Retorna o resultado da inclusão
Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function AutoNumeracao() As Long
Dim lRet As Long

lRet = 1
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_retirada FROM caixa_retirada;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lRet = r("cod_retirada") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

AutoNumeracao = lRet
End Function

Private Sub Campos_Brancos()
txtCodigo.Text = ""
mskData.Mask = ""
mskData.Text = ""
mskHora.Mask = ""
mskHora.Text = ""
txtValor.Text = ""
cboSubDesc.Text = ""
cboSubDesc.Clear
lblHora.Caption = ""
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



Private Sub cboCaixa_GotFocus()
cboCaixa.Clear
cboCaixa.AddItem "CAIXA01"
cboCaixa.AddItem "CAIXA02"
cboCaixa.AddItem "CAIXA03"
cboCaixa.AddItem "CAIXA04"
cboCaixa.AddItem "TODOS"
moCombo.AttachTo cboCaixa
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
    cboMES.SetFocus
ElseIf cboCriterio.Text = "DATA" Then
    mskConsData.Visible = True
    cmdConsData.Visible = True
    cboMES.Visible = False
    cboAno.Visible = False
    lblCONmes.Visible = False
    'mskConsData.SetFocus
End If
End Sub


Private Sub cboFuncionario_LostFocus()
On Error GoTo TrataErro

If cboFuncionario.Text = "" Then txtCodFunc.Text = "": Exit Sub

If cmdAlterar.Enabled = False Then
   If cboFuncionario.ListIndex = -1 Then
      'txtCodFunc.Text = ""
      'Exit Sub
   End If
End If

txtCodFunc = cboFuncionario.ItemData(cboFuncionario.ListIndex)
Exit Sub

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboIndice_GotFocus()
cboIndice.Clear
cboIndice.AddItem "DATA"
cboIndice.AddItem "HORA"
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

Private Sub cmdAlterar_Click()
If txtCodigo.Text = "" Then Exit Sub

'verificar situação do caixa
ConsultarCaixaAtual

If varCodCaixa = 0 Then
    MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

'Faz a atualização de forma direta e verifica se houve algum erro
If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

cmdExibir_Click
vCodFunc = 0
cboFuncionario.Text = ""
cboFuncionario.Clear
txtCodFunc.Text = ""
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
cboFuncionario.Text = ""
cboFuncionario.Clear
txtCodFunc.Text = ""
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
Dim bRet As Boolean
If txtCodigo.Text = "" Then Exit Sub

'verificar se o caixa está aberto
ConsultarCaixaAtual

If varCodCaixa = 0 Then
    MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If ShowMsg("Excluir essa retirada?", vbInformation + vbYesNo) = vbNo Then Exit Sub

'excluir a sangria
sSQL = "DELETE FROM caixa_retirada WHERE (codigo = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

cmdExibir_Click
vCodFunc = 0
cboFuncionario.Text = ""
cboFuncionario.Clear
txtCodFunc.Text = ""
Campos_Brancos
Form_Load
End Sub


Private Sub cmdExibir_Click()
Dim varCriterio As String
If cboCriterio.Text = "TODOS" Then
    varCriterio = " where codigo <> 0 "
ElseIf cboCriterio.Text = "MENSAL" Then
   If cboMES.Text = "" Or cboMES.ListIndex = -1 Then Exit Sub
   If cboAno.Text = "" Or cboAno.ListIndex = -1 Then Exit Sub
    varCriterio = "WHERE (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ") "
ElseIf cboCriterio.Text = "DATA" Then
   If Not IsDate(mskConsData) Then Exit Sub
    varCriterio = "WHERE (data = CONVERT(DATETIME, '" & Format(mskConsData.Text, ocDATA) & "', 103))  "
Else
    varCriterio = " "
End If

Dim varCAIXA As String
If cboCaixa.Text = "TODOS" Then
    varCAIXA = " "
Else
    varCAIXA = " and (CAIXA = '" & cboCaixa.Text & "') "
End If

Dim vIndice As String
If cboIndice.Text = "DATA" Then
    vIndice = " order by DATA"
ElseIf cboIndice.Text = "HORA" Then
    vIndice = " order by HORA"
End If

Dim totalRegistros As Long

sSQL = "SELECT * FROM caixa_retirada  " & varCriterio & varCAIXA & vIndice & ""

Set r = dbData.OpenRecordset(sSQL, totalRegistros)
FormatarGrid r

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

Set REL_Caixa_Retirada_Cons.Relatorio.Recordset = r
REL_Caixa_Retirada_Cons.lblTitulo.Caption = "RELATÓRIO DE CAIXA - RETIRADAS"

REL_Caixa_Retirada_Cons.rfCriterio.Caption = cboCriterio.Text
'REL_Caixa_Retirada_Cons.rfFontes.Caption = cboConsFontes.Text
'REL_Caixa_Retirada_Cons.rfTitOrigem.Caption = cboOrigem.Text

If cboCriterio.Text = "TODOS" Then
    REL_Caixa_Retirada_Cons.rfCriterioDesc.Caption = ""
ElseIf cboCriterio.Text = "MENSAL" Then
    REL_Caixa_Retirada_Cons.rfCriterioDesc.Caption = cboMES.Text & " / " & cboAno.Text
ElseIf cboCriterio.Text = "DATA" Then
    REL_Caixa_Retirada_Cons.rfCriterioDesc.Caption = mskConsData.Text
End If

REL_Caixa_Retirada_Cons.rfQuant.Caption = lblQuant.Caption
REL_Caixa_Retirada_Cons.rfSubTotal.Caption = Format(lblValor.Caption, ocMONEY)

'REL_Caixa_Retirada_cons.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
'REL_Caixa_Retirada_cons.rfCodCaixa.Caption = varFluxoCodCaixa
'REL_Caixa_Retirada_cons.rfNomeCaixa.Caption = varFluxoNomeCaixa

REL_Caixa_Retirada_Cons.Relatorio.NomeImpressora = var_Impressora
REL_Caixa_Retirada_Cons.Relatorio.Ativar
Unload REL_Caixa_Retirada_Cons

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
mskHora.Text = Format(Now, "hh:mm")
mskData.Text = Format(Date, "dd/mm/yy")
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

cboFuncionario.SelStart = 0
cboFuncionario.SelLength = Len(cboFuncionario)

moCombo.AttachTo cboFuncionario
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSubDesc_GotFocus()
Dim varNomeAntes As String
varNomeAntes = cboSubDesc.Text

cboSubDesc.Clear
cboSubDesc.AddItem "RETIRADA"

cboSubDesc.Text = varNomeAntes
SelectControl cboSubDesc
moCombo.AttachTo cboSubDesc
End Sub

Private Sub cboSubDesc_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdSalvar_Click()
If txtValor.Text = "" Or cboSubDesc.Text = "" Or cboFuncionario.Text = "" Then
   ShowMsg "Formulário incompleto!", vbInformation
   cboSubDesc.SetFocus
   Exit Sub
End If

Dim lNovoCod As Long
   
 'verificar se o caixa está aberto
 If varCodCaixa = 0 Then
     MsgBox "O caixa ainda não foi aberto", vbInformation, "Aviso do Sistema"
     Exit Sub
 End If

 'criar uma retirada
 lNovoCod = AutoNumeracao
 
 'Faz a inserção de forma direta e verifica se houve algum erro
 If Not Inserir_Dados(lNovoCod) Then
    ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
    Exit Sub
 End If
      
If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    If varTipoRecPgto = "CUPOM" Then
        Imprimir_ReciboCupom
    Else
        Imprimir_ReciboFolha
    End If
End If
vCodFunc = 0
cboFuncionario.Text = ""
cboFuncionario.Clear
txtCodFunc.Text = ""
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

With REL_ReciboRetirada
    .txtDescricao.Caption = UCase(cboSubDesc.Text)
    .txtUsuario.Caption = UCase(cboFuncionario.Text)
    .txtFormaPgto.Caption = UCase("DINHEIRO")
    .txtValor.Caption = UCase(NumeroExtenso(txtValor.Text, True))
    .txthead.Caption = "R$ " & Format(txtValor.Text, "##,##0.00")
    '.txtProveniente.Caption = "Pagamento da " & txtNumParcela.Text & "ª parcela do PEDIDO Nº " & Format(txtCodPedido.Text, "000000")
    .txtData.Caption = "" & vCidadeUF & ", " & Day(mskData) & " de " & MonthName(Month(mskData)) & " de " & Year(mskData)
    
    .Relatorio.NumeroRegistros = 1
    .Relatorio.NomeImpressora = var_ImpNormal
    .Relatorio.Ativar
End With

Unload REL_ReciboRetirada
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
      Printer.Print Tab(2); cboFuncionario.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next

      'Fonte 8, False, False
      'Printer.Print Tab(2); "Setor: "
      'Fonte 8, True, False
      'Printer.Print Tab(2); cboSetor.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next

      'Fonte 8, False, False
      'Printer.Print Tab(2); "Fonte: "
      'Fonte 8, True, False
      'Printer.Print Tab(2); cboFonte.Text
      
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

'colocar o nome da maquina na barra de status
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")


StatusBar1.Panels(2).Text = var_Caixa

'verificar se o caixa está aberto
ConsultarCaixaAtual

StatusBar1.Panels(3).Text = varCodCaixa
StatusBar1.Panels(4).Text = Format(Now, "hh:mm")
StatusBar1.Panels(5).Text = Format(Date, "dd/mm/yy")

cboCriterio.Text = "DATA"
cboCaixa.Text = "CAIXA01"
cboIndice.Text = "HORA"
mskConsData.Text = Format(Date, "dd/mm/yy")


'colocar o nome da maquina na barra de status
'Dim var_Caixa As String
'Dim oIni As Ini

'Set oIni = New Ini
'oIni.Arquivo = appPathApp & "config.ini"    'abre o ini
'var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")


'StatusBar1.Panels(2).Text = var_Caixa
'StatusBar1.Panels(4).Text = Format(Date, "dd/mm/yy")

'abrindo arquivo .ini
'Set oIni = New Ini
'oIni.Arquivo = appPathApp & "config.ini"

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

'cmdExibir_Click
Set moCombo = New cComboHelper
End Sub

Private Sub Mostrar_Saida(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then
   cboSubDesc.Text = rTabela("descricao")
   txtCodFunc.Text = rTabela("cod_funcionario")
   mskData.Text = Format(rTabela("data"), "dd/mm/yy")
   txtValor.Text = Format(rTabela("valor"), ocMONEY)
   mskHora.Text = Format(rTabela("hora"), "hh:mm")
   StatusBar1.Panels(2).Text = rTabela("CAIXA")
   StatusBar1.Panels(3).Text = rTabela("CODCAIXA")
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
    .Cols = 9
    .rows = 2
        
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 800
    .ColWidth(3) = 650
    .ColWidth(4) = 650
    .ColWidth(5) = 3000
    .ColWidth(6) = 1500
    .ColWidth(7) = 900
    .ColWidth(8) = 900
   
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "DATA"
   .TextMatrix(0, 3) = "HORA"
   .TextMatrix(0, 4) = "FUNC"
   .TextMatrix(0, 5) = "DESCRIÇÃO"
   .TextMatrix(0, 6) = "VALOR"
   .TextMatrix(0, 7) = "CAIXA"
   .TextMatrix(0, 8) = "CÓD."
   .Redraw = False
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.rows - 1, 2) = Format(rTabela("data"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = Format(rTabela("hora"), ocHRMN)
         .TextMatrix(.rows - 1, 4) = rTabela("COD_FUNCIONARIO")
         .TextMatrix(.rows - 1, 5) = rTabela("DESCRICAO")
         .TextMatrix(.rows - 1, 6) = Format(rTabela("valor"), ocMONEY)
         .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("CAIXA"))
         .TextMatrix(.rows - 1, 8) = Format(rTabela("CODCAIXA"), "000000")
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

lblValor.Caption = Format(SomaGrid(GridSaidas, 6), ocMONEY)
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
    MsgBox "O caixa onde essa retirada foi adicionado encontra-se fechado!", vbInformation, "Aviso do Sistema"
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

Private Sub MaskEdBox1_Change()

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

Private Sub mskHora_GotFocus()
SelectControl mskHora
If cmdAlterar.Enabled = False Then
    mskHora.Text = Format(Now, "hh:mm")
End If
End Sub

Private Sub mskHora_KeyPress(KeyAscii As Integer)
mskHora.Mask = "##:##"
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
   sSQL = "SELECT * FROM caixa_retirada WHERE (codigo = " & txtCodigo.Text & ");"
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
