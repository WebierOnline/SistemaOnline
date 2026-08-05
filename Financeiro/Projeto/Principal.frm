VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Configuracao_Geral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONFIGURAÇÕES"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   8865
      TabIndex        =   15
      Top             =   60
      Width           =   8895
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIGURAÇÕES GERAIS"
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
         TabIndex        =   16
         Top             =   300
         Width           =   4005
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   300
         Picture         =   "Principal.frx":23D2
         Top             =   0
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   14
      Top             =   9930
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "16:46"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8955
      Left            =   60
      TabIndex        =   1
      Top             =   1080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   15796
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabMaxWidth     =   2999
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
      TabPicture(0)   =   "Principal.frx":28A9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label60"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label59"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label58"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label57"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label56"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label54"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label53"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDesmarcarTodos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdDesmarcar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdMarcar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdLocalizar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdNovo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdPrepara2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdPrepara"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdMostrarSenha"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAdicionar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "mskCPF"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Grid"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chameleonButton1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCodDesbloqueioTemp"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCodDesbloqueio"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtFantasia"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboAno"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cboMes"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtRazao"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "0"
      TabPicture(1)   =   "Principal.frx":28C5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "0"
      TabPicture(2)   =   "Principal.frx":28E1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdSalvarBalanca"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "0"
      TabPicture(3)   =   "Principal.frx":28FD
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "0"
      TabPicture(4)   =   "Principal.frx":2919
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.TextBox txtRazao 
         Height          =   315
         Left            =   3000
         TabIndex        =   21
         Top             =   5460
         Width           =   3735
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   6420
         Width           =   1515
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   6420
         Width           =   1155
      End
      Begin VB.TextBox txtFantasia 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   5460
         Width           =   2835
      End
      Begin VB.TextBox txtCodDesbloqueio 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   6420
         Width           =   975
      End
      Begin VB.TextBox txtCodDesbloqueioTemp 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   6420
         Width           =   975
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarBalanca 
         Height          =   615
         Left            =   -68400
         TabIndex        =   18
         Top             =   7980
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Salvar"
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
         MICON           =   "Principal.frx":2935
         PICN            =   "Principal.frx":2951
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton chameleonButton1 
         Height          =   315
         Left            =   8400
         TabIndex        =   20
         Top             =   5460
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "C"
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
         MICON           =   "Principal.frx":46E3
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
         Height          =   4635
         Left            =   120
         TabIndex        =   22
         Top             =   420
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8176
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox mskCPF 
         Height          =   315
         Left            =   6780
         TabIndex        =   2
         Top             =   5460
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdAdicionar 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   5820
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Salvar"
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
         MICON           =   "Principal.frx":46FF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMostrarSenha 
         Height          =   315
         Left            =   2880
         TabIndex        =   11
         Top             =   6420
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
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
         MICON           =   "Principal.frx":471B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdPrepara 
         Height          =   315
         Left            =   5040
         TabIndex        =   23
         Top             =   6420
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "C"
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
         MICON           =   "Principal.frx":4737
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdPrepara2 
         Height          =   315
         Left            =   6420
         TabIndex        =   24
         Top             =   6420
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "C"
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
         MICON           =   "Principal.frx":4753
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
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   5820
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Novo"
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
         MICON           =   "Principal.frx":476F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   315
         Left            =   2220
         TabIndex        =   5
         Top             =   5820
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Localizar"
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
         MICON           =   "Principal.frx":478B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMarcar 
         Height          =   315
         Left            =   5220
         TabIndex        =   6
         Top             =   5820
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Marcar"
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
         MICON           =   "Principal.frx":47A7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdDesmarcar 
         Height          =   315
         Left            =   6240
         TabIndex        =   7
         Top             =   5820
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Desmarcar"
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
         MICON           =   "Principal.frx":47C3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdDesmarcarTodos 
         Height          =   315
         Left            =   7320
         TabIndex        =   8
         Top             =   5820
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Desmarcar Todos"
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
         MICON           =   "Principal.frx":47DF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razão"
         Height          =   195
         Left            =   3060
         TabIndex        =   31
         Top             =   5220
         Width           =   465
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   6780
         TabIndex        =   30
         Top             =   5220
         Width           =   405
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ano"
         Height          =   195
         Left            =   1680
         TabIndex        =   29
         Top             =   6180
         Width           =   285
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mês"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   6180
         Width           =   300
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fantasia"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   5220
         Width           =   600
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Certo"
         Height          =   195
         Left            =   4020
         TabIndex        =   26
         Top             =   6180
         Width           =   375
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Temporário"
         Height          =   195
         Left            =   5460
         TabIndex        =   25
         Top             =   6180
         Width           =   795
      End
   End
   Begin VB.Label Label51 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "À Vista:"
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   540
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Confirmar fechamendo da venda:"
      Height          =   195
      Left            =   540
      TabIndex        =   17
      Top             =   5040
      Width           =   2355
   End
End
Attribute VB_Name = "Configuracao_Geral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private moCombo As cComboHelper
Private Caminho As String
'Dim oCfg As ConfigItem
Dim sSQL As String
Dim r As ADODB.Recordset
Dim i As Integer
Private Sub Form_Load()
'Set moCombo = New cComboHelper
End Sub


Private Sub mskCPF_KeyPress(KeyAscii As Integer)
mskCPF.Mask = "##.###.###/####-##"
End Sub
Private Sub chameleonButton1_Click()
Clipboard.Clear
Clipboard.SetText mskCPF.TEXT
End Sub
Private Sub cmdAdicionar_Click()
If txtFantasia.TEXT = "" Or txtRazao.TEXT = "" Or mskCPF.TEXT = "" Then Exit Sub

sSQL = "SELECT CNPJ FROM  empresas_desbloueio WHERE CNPJ = '" & mskCPF.TEXT & "';"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    MsgBox "Empresa já cadastrada!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

MostrarEmpresa
LimparEmpresa

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub cmdNovo_Click()
LimparEmpresa
MostrarEmpresa
End Sub
Private Sub LimparEmpresa()
txtFantasia.TEXT = ""
txtRazao.TEXT = ""
mskCPF.Mask = ""
mskCPF.TEXT = ""
End Sub

Private Sub MostrarEmpresa()
sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado FROM  empresas_desbloueio ORDER BY FANTASIA;"
Set r = dbData.OpenRecordset(sSQL)
FormatarGrid r
End Sub
Private Sub cmdLocalizar_Click()
sSQL = "SELECT * FROM  empresas_desbloueio where (FANTASIA LIKE '%" & txtFantasia.TEXT & "%')"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim x As Integer

With Grid
   .Clear
   .Cols = 6
   .Rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 2500
   .ColWidth(2) = 4000
   .ColWidth(3) = 1800
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "FANTASIA"
   .TextMatrix(0, 2) = "RAZÃO."
   .TextMatrix(0, 3) = "CNPJ"
   .TextMatrix(0, 4) = "CODIGO"
   
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.Rows - 1, 1) = rTabela("FANTASIA")
         .TextMatrix(.Rows - 1, 2) = rTabela("RAZAO")
         .TextMatrix(.Rows - 1, 3) = rTabela("CNPJ")
         .TextMatrix(.Rows - 1, 4) = rTabela("CODIGO")
         .TextMatrix(.Rows - 1, 5) = rTabela("vMarcado")
         rTabela.MoveNext
         
         .Rows = .Rows + 1
         i = i + 1
      Loop
   End If
   
   
   For i = 1 To .Rows - 1
       For j = 0 To .Cols - 1
          .Col = j
          .Row = i
    
          If .TextMatrix(i, 5) = "NÃO" Then
             .CellForeColor = vbBlack
          ElseIf .TextMatrix(i, 5) = "SIM" Then
             .CellForeColor = vbRed
          Else
             .CellForeColor = vbBlack
          End If
          
       Next
    Next
   
   .Rows = .Rows - 1
   .Redraw = True
End With
End Sub
Private Sub cmdMarcar_Click()
i = Grid.Row
dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 1 WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
MostrarEmpresa
End Sub
Private Sub cmdDesmarcar_Click()
i = Grid.Row
dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 0 WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
MostrarEmpresa
End Sub
Private Sub cmdDesmarcarTodos_Click()
dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 0;"
MostrarEmpresa
End Sub
Private Sub cboMes_GotFocus()
Dim vMes As Integer

cboMes.Clear

For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMes
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
Private Sub cmdMostrarSenha_Click()
If cboMes.TEXT = "" Then Exit Sub
If cboAno.TEXT = "" Then Exit Sub

Dim vCnpj As Integer
Dim vQuantRazao As Integer

vCnpj = SomarDigitos(mskCPF.TEXT)
vQuantRazao = Len(txtRazao.TEXT)

Dim vNumeroMes As Integer
If cboMes.TEXT = "Janeiro" Then
    vNumeroMes = 1
ElseIf cboMes.TEXT = "Fevereiro" Then
    vNumeroMes = 2
ElseIf cboMes.TEXT = "Março" Then
    vNumeroMes = 3
ElseIf cboMes.TEXT = "Abril" Then
    vNumeroMes = 4
ElseIf cboMes.TEXT = "Maio" Then
    vNumeroMes = 5
ElseIf cboMes.TEXT = "Junho" Then
    vNumeroMes = 6
ElseIf cboMes.TEXT = "Julho" Then
    vNumeroMes = 7
ElseIf cboMes.TEXT = "Agosto" Then
    vNumeroMes = 8
ElseIf cboMes.TEXT = "Setembro" Then
    vNumeroMes = 9
ElseIf cboMes.TEXT = "Outubro" Then
    vNumeroMes = 10
ElseIf cboMes.TEXT = "Novembro" Then
    vNumeroMes = 11
ElseIf cboMes.TEXT = "Dezembro" Then
    vNumeroMes = 12
End If

'começa a criação
Dim vDataInicio As Date
Dim vDia As Integer
Dim vMes As Integer
Dim vMesInt As String
Dim vAno As Integer
Dim vMesRef As String

vDia = 30
vMes = vNumeroMes
vAno = cboAno

Dim vDataBloqueio As String

    vDataInicio = vDia & " / " & vMes & " / " & vAno
    vMesInt = Format(vDataInicio, "mmmm")
    
    Dim vCodDesbloqueio As String
    Dim vCodDesbTemp As String
    
    'Desbloqueio
    If vNumeroMes Mod 2 = 0 Then
        'MsgBox "Par!"
        vCodDesbloqueio = Left(vCnpj, 1) & "" & Left(vQuantRazao, 1) & "" & Len(vMesInt) & "" & vNumeroMes & "" & UCase(Mid(vMesInt, 3, 1))
    Else
        'MsgBox "Ímpar!"
        vCodDesbloqueio = Mid(vCnpj, 2, 1) & "" & Mid(vQuantRazao, 2, 1) & "" & Len(vMesInt) - 1 & "" & vNumeroMes & "" & UCase(Mid(vMesInt, 2, 1))
    End If

    'Desbloqueio temporario
    If vNumeroMes Mod 2 = 0 Then
        'MsgBox "Par!"
        vCodDesbTemp = Left(vCodDesbloqueio, 1) & "" & Left(vCodDesbloqueio, 1) & "" & vNumeroMes + 1 & "" & UCase(Mid(vMesInt, 4, 1))
    Else
        'MsgBox "Ímpar!"
        vCodDesbTemp = Mid(vCodDesbloqueio, 2, 1) & "" & Mid(vCodDesbloqueio, 2, 1) & "" & Len(vMesInt) - 1 & "" & vNumeroMes + 1 & "" & UCase(Mid(vMesInt, 4, 1))
    End If
    
txtCodDesbloqueio.TEXT = vCodDesbloqueio
txtCodDesbloqueioTemp.TEXT = vCodDesbTemp
End Sub
Private Sub cmdPrepara_Click()
Clipboard.Clear
Clipboard.SetText "[MENSAGEM AUTOMÁTICA]: Seu código de desbloqueio é: " & txtCodDesbloqueio.TEXT & "   - Obs: O último caractere é uma letra. "
End Sub
Private Sub cmdPrepara2_Click()
Clipboard.Clear
Clipboard.SetText "Seu código de desbloqueio temporário é: " & txtCodDesbloqueioTemp.TEXT
End Sub
